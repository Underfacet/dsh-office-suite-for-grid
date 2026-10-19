/**
 * Office 排版（docs/P2-设计.md §2.3 / §3.3）
 *
 * 单文件：formatOp(ctx, config, args, exec)
 *   args = {path, spec:{rules?, styleSet?, template?, placeholders?, scope?}, template?, engine?}
 *   兼容工具形态：{path, rules?, styleSet?, template?, placeholders?, engine?, scope?}
 * 批量：batchFormatOp(ctx, config, args, exec, store?) → {taskId, filesTotal}（tasks.js，kind='office.batch'）
 *
 * rules / styleSet / template 由 Host 解析：
 * - styleSet：内置 suite\templates\styles\<id>.json，用户 runtime\templates\styles\<id>.json（同名用户优先）；
 * - template：suite\templates\docs\<name>.docx，用户 runtime\templates\docs\<name>.docx；
 * - template 优先于 rules（模板提供版式骨架，仅填充占位符）。
 * 修改前备份由 backup.js 负责（失败即中止）。
 */
import { readdir, readFile, stat } from 'node:fs/promises';
import { basename, extname, isAbsolute, join, resolve } from 'node:path';
import {
  OpsError,
  PACKAGE_DIR,
  RUNTIME_DIR,
  fail,
  isAbortError,
  ok,
  resolveInputPath,
  throwIfAborted,
  toErrorPayload,
} from '../ops.js';
import { assertInsideRoot, resolveBoundaryRoot } from '../boundary.js';
import { backupFile } from './backup.js';
import { familyOfExtension, decideOfficeEngine, executeOfficeEngine } from './office-edit.js';
import { tasks as defaultTasks } from '../tasks.js';

const TARGETS = new Set(['heading', 'body', 'page', 'header', 'footer']);
const NUMBERING_MODES = new Set(['chinese', 'arabic', 'decimalDot', 'none']);
const ALIGNS = new Set(['center', 'left', 'right', 'justify']);
const MAX_RULES = 100;
/** 批量任务单次最多文件数（防御异常目录）。 */
const MAX_BATCH_FILES = 200;

// ───────────────────────── 排版规则校验（§3.3） ─────────────────────────

function isPlainObject(value) {
  return !!value && typeof value === 'object' && !Array.isArray(value);
}

function ruleError(index, message) {
  return new OpsError('E2006', index === null ? message : `rules[${index}] ${message}`);
}

function finiteNumber(value, field, index, { min = -Infinity, max = Infinity } = {}) {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < min || value > max) {
    throw ruleError(index, `${field} 必须是 ${min === -Infinity ? '' : `${min}–`}${max === Infinity ? '有限数字' : max} 的数字，收到: ${JSON.stringify(value)}`);
  }
  return value;
}

function normalizeFont(font, index) {
  if (!isPlainObject(font)) throw ruleError(index, 'font 必须是对象');
  const out = {};
  if (font.name !== undefined) {
    if (typeof font.name !== 'string' || !font.name.trim()) throw ruleError(index, 'font.name 必须是非空字符串');
    out.name = font.name.trim();
  }
  if (font.sizePt !== undefined) out.sizePt = finiteNumber(font.sizePt, 'font.sizePt', index, { min: 1, max: 200 });
  if (font.bold !== undefined) {
    if (typeof font.bold !== 'boolean') throw ruleError(index, 'font.bold 必须是布尔值');
    out.bold = font.bold;
  }
  if (!Object.keys(out).length) throw ruleError(index, 'font 至少需要 name/sizePt/bold 之一');
  return out;
}

function normalizeParagraphFormat(paragraph, index) {
  if (!isPlainObject(paragraph)) throw ruleError(index, 'paragraph 必须是对象');
  const out = {};
  if (paragraph.lineSpacing !== undefined) {
    out.lineSpacing = finiteNumber(paragraph.lineSpacing, 'paragraph.lineSpacing', index, { min: 0.5, max: 10 });
  }
  if (paragraph.lineSpacingPt !== undefined) {
    out.lineSpacingPt = finiteNumber(paragraph.lineSpacingPt, 'paragraph.lineSpacingPt', index, { min: 1, max: 200 });
  }
  if (out.lineSpacing !== undefined && out.lineSpacingPt !== undefined) {
    throw ruleError(index, 'paragraph.lineSpacing 与 lineSpacingPt 只能二选一（倍数 / 固定磅值）');
  }
  if (paragraph.spaceBeforePt !== undefined) {
    out.spaceBeforePt = finiteNumber(paragraph.spaceBeforePt, 'paragraph.spaceBeforePt', index, { min: 0, max: 500 });
  }
  if (paragraph.spaceAfterPt !== undefined) {
    out.spaceAfterPt = finiteNumber(paragraph.spaceAfterPt, 'paragraph.spaceAfterPt', index, { min: 0, max: 500 });
  }
  if (paragraph.align !== undefined) {
    if (!ALIGNS.has(paragraph.align)) {
      throw ruleError(index, `paragraph.align 必须为 ${[...ALIGNS].join('|')}，收到: ${JSON.stringify(paragraph.align)}`);
    }
    out.align = paragraph.align;
  }
  if (!Object.keys(out).length) throw ruleError(index, 'paragraph 至少需要一个排版字段');
  return out;
}

function normalizePage(page, index) {
  if (!isPlainObject(page)) throw ruleError(index, 'page 必须是对象');
  const margins = page.marginsMm;
  if (!isPlainObject(margins)) throw ruleError(index, 'page.marginsMm 必须是对象');
  const out = { marginsMm: {} };
  for (const side of ['top', 'bottom', 'left', 'right']) {
    if (margins[side] === undefined) continue;
    out.marginsMm[side] = finiteNumber(margins[side], `page.marginsMm.${side}`, index, { min: 1, max: 500 });
  }
  if (!Object.keys(out.marginsMm).length) throw ruleError(index, 'page.marginsMm 至少需要一项（top/bottom/left/right）');
  return out;
}

function normalizeHeaderFooter(value, index, label) {
  if (!isPlainObject(value)) throw ruleError(index, `${label} 必须是对象`);
  if (typeof value.text !== 'string') throw ruleError(index, `${label}.text 必须是字符串`);
  const out = { text: value.text };
  if (value.font !== undefined) out.font = normalizeFont(value.font, index);
  return out;
}

function normalizeRule(rule, index) {
  if (!isPlainObject(rule)) throw ruleError(index, '规则必须是对象');
  if (!TARGETS.has(rule.target)) {
    throw ruleError(index, `target 必须为 ${[...TARGETS].join('|')}，收到: ${JSON.stringify(rule.target)}`);
  }
  const out = { target: rule.target };
  if (rule.target === 'heading') {
    if (!Number.isInteger(rule.level) || rule.level < 1 || rule.level > 3) {
      throw ruleError(index, `heading 规则 level 必须为 1–3 的整数，收到: ${JSON.stringify(rule.level)}`);
    }
    out.level = rule.level;
  }
  if (rule.font !== undefined) out.font = normalizeFont(rule.font, index);
  if (rule.paragraph !== undefined) out.paragraph = normalizeParagraphFormat(rule.paragraph, index);
  if (rule.numbering !== undefined) {
    if (!NUMBERING_MODES.has(rule.numbering)) {
      throw ruleError(index, `numbering 必须为 ${[...NUMBERING_MODES].join('|')}，收到: ${JSON.stringify(rule.numbering)}`);
    }
    out.numbering = rule.numbering;
  }
  if (rule.page !== undefined) out.page = normalizePage(rule.page, index);
  if (rule.header !== undefined) out.header = normalizeHeaderFooter(rule.header, index, 'header');
  if (rule.footer !== undefined) out.footer = normalizeHeaderFooter(rule.footer, index, 'footer');

  if (rule.target === 'page' && !out.page) throw ruleError(index, 'page 规则必须提供 page.marginsMm');
  if (rule.target === 'header' && !out.header) throw ruleError(index, 'header 规则必须提供 header.text');
  if (rule.target === 'footer' && !out.footer) throw ruleError(index, 'footer 规则必须提供 footer.text');
  if (rule.target === 'heading' || rule.target === 'body') {
    if (!out.font && !out.paragraph && out.numbering === undefined) {
      throw ruleError(index, `${rule.target} 规则至少需要 font/paragraph/numbering 之一`);
    }
  }
  return out;
}

/** 校验 scope：'all' | {headings:[1..3]} | {paragraphs:[idx]}。 */
export function validateScope(scope) {
  if (scope === undefined || scope === null || scope === 'all') return { type: 'all' };
  if (!isPlainObject(scope)) {
    throw new OpsError('E2006', "scope 必须为 'all' 或 {headings:[1..3]} / {paragraphs:[索引]}");
  }
  if (scope.headings !== undefined) {
    if (!Array.isArray(scope.headings) || !scope.headings.length) throw new OpsError('E2006', 'scope.headings 必须是非空数组');
    const levels = [...new Set(scope.headings)];
    for (const level of levels) {
      if (!Number.isInteger(level) || level < 1 || level > 3) {
        throw new OpsError('E2006', `scope.headings 只能包含 1–3 的整数，收到: ${JSON.stringify(level)}`);
      }
    }
    return { type: 'headings', headings: levels.sort((a, b) => a - b) };
  }
  if (scope.paragraphs !== undefined) {
    if (!Array.isArray(scope.paragraphs) || !scope.paragraphs.length) throw new OpsError('E2006', 'scope.paragraphs 必须是非空数组');
    const paragraphs = [...new Set(scope.paragraphs)];
    for (const index of paragraphs) {
      if (!Number.isInteger(index) || index < 1) {
        throw new OpsError('E2006', `scope.paragraphs 只能包含正整数，收到: ${JSON.stringify(index)}`);
      }
    }
    return { type: 'paragraphs', paragraphs: paragraphs.sort((a, b) => a - b) };
  }
  throw new OpsError('E2006', 'scope 需为 {headings:[...]} 或 {paragraphs:[...]}');
}

/**
 * 校验排版规格（rules / styleSet / template / placeholders / scope）。
 * @returns {object} 归一化后的 spec（rules 已逐条校验；styleSet 仅校验类型）
 * @throws {OpsError} E2006 FORMAT_RULE_INVALID
 */
export function validateFormatSpec(spec) {
  const raw = Array.isArray(spec) ? { rules: spec } : spec;
  if (!isPlainObject(raw)) {
    throw new OpsError('E2006', 'spec 必须是对象（或 rules 数组）');
  }
  const out = {};
  if (raw.rules !== undefined) {
    if (!Array.isArray(raw.rules)) throw new OpsError('E2006', 'spec.rules 必须是数组');
    if (raw.rules.length > MAX_RULES) {
      throw new OpsError('E2006', `spec.rules 数量超过上限（${MAX_RULES}）`, `收到 ${raw.rules.length} 条`);
    }
    out.rules = raw.rules.map((rule, index) => normalizeRule(rule, index));
  } else {
    out.rules = [];
  }

  if (raw.styleSet !== undefined) {
    if (typeof raw.styleSet === 'string') {
      if (!raw.styleSet.trim()) throw new OpsError('E2006', 'styleSet 不能为空');
      out.styleSet = raw.styleSet.trim();
    } else if (isPlainObject(raw.styleSet)) {
      if (!Array.isArray(raw.styleSet.rules)) throw new OpsError('E2006', 'styleSet 对象必须包含 rules 数组');
      out.styleSet = { ...raw.styleSet, rules: raw.styleSet.rules.map((rule, index) => normalizeRule(rule, index)) };
    } else {
      throw new OpsError('E2006', 'styleSet 必须是样式集名称或 {rules:[...]} 对象');
    }
  }

  if (raw.template !== undefined) {
    if (typeof raw.template !== 'string' || !raw.template.trim()) throw new OpsError('E2006', 'template 必须是非空字符串（模板名或路径）');
    out.template = raw.template.trim();
  }

  if (raw.placeholders !== undefined) {
    if (!isPlainObject(raw.placeholders)) throw new OpsError('E2006', 'placeholders 必须是对象');
    const placeholders = {};
    for (const [key, value] of Object.entries(raw.placeholders)) {
      if (!key) continue;
      if (value === undefined || value === null) continue;
      if (!['string', 'number', 'boolean'].includes(typeof value)) {
        throw new OpsError('E2006', `placeholders.${key} 必须是字符串/数字/布尔值`);
      }
      placeholders[key] = String(value);
    }
    out.placeholders = placeholders;
  }

  out.scope = validateScope(raw.scope);
  return out;
}

// ───────────────────────── 样式集 / 模板解析 ─────────────────────────

function templatesRoot(config = {}) {
  const custom = typeof config.officeTemplatesDir === 'string' ? config.officeTemplatesDir.trim() : '';
  if (custom) return resolve(custom);
  return join(RUNTIME_DIR, 'templates');
}

function styleDirs(config) {
  return [...new Set([join(templatesRoot(config), 'styles'), join(PACKAGE_DIR, 'templates', 'styles')])];
}

function templateDirs(config) {
  return [...new Set([join(templatesRoot(config), 'docs'), join(PACKAGE_DIR, 'templates', 'docs')])];
}

function safeName(name) {
  const value = String(name ?? '').trim();
  if (!value || /[\\/:*?"<>|]/.test(value) || value.includes('..')) return null;
  return value;
}

async function exists(path) {
  try {
    await stat(path);
    return true;
  } catch {
    return false;
  }
}

/** 加载样式集（用户目录优先于内置）。 */
export async function resolveStyleSet(styleSet, config = {}) {
  if (isPlainObject(styleSet)) {
    if (!Array.isArray(styleSet.rules)) throw new OpsError('E2006', 'styleSet 对象必须包含 rules 数组');
    return {
      id: typeof styleSet.id === 'string' ? styleSet.id : 'inline',
      builtin: false,
      file: null,
      rules: styleSet.rules.map((rule, index) => normalizeRule(rule, index)),
    };
  }
  const name = safeName(styleSet);
  if (!name) throw new OpsError('E2006', `styleSet 名称非法: ${JSON.stringify(styleSet)}`);
  for (const dir of styleDirs(config)) {
    const file = join(dir, `${name}.json`);
    if (!(await exists(file))) continue;
    let parsed;
    try {
      parsed = JSON.parse((await readFile(file, 'utf8')).replace(/^\uFEFF/, ''));
    } catch (err) {
      throw new OpsError('E2006', `样式集 JSON 解析失败: ${file}`, String(err?.message ?? err));
    }
    const rules = Array.isArray(parsed) ? parsed : parsed?.rules;
    if (!Array.isArray(rules)) throw new OpsError('E2006', `样式集缺少 rules 数组: ${file}`);
    return {
      id: name,
      builtin: dir.startsWith(join(PACKAGE_DIR, 'templates')),
      file,
      rules: rules.map((rule, index) => normalizeRule(rule, index)),
    };
  }
  throw new OpsError(
    'E2007',
    `样式集不存在: ${name}`,
    `内置目录：${join(PACKAGE_DIR, 'templates', 'styles')}；用户目录：${join(templatesRoot(config), 'styles')}`,
  );
}

/** 解析模板（名称或路径；名称在用户/内置 docs 目录中查找 .docx/.doc）。 */
export async function resolveTemplate(template, ctx, config = {}, exec = {}) {
  const raw = String(template ?? '').trim();
  if (!raw) throw new OpsError('E2006', 'template 不能为空');
  const looksLikePath = isAbsolute(raw) || /[\\/]/.test(raw) || /\.(docx?|doc)$/i.test(raw);
  if (looksLikePath) {
    const abs = assertInsideRoot(await resolveInputPath(ctx, exec, raw), {
      write: false,
      root: resolveBoundaryRoot(config),
      allowExternalPaths: config.allowExternalPaths === true,
    });
    if (!(await exists(abs))) throw new OpsError('E2007', `模板文件不存在: ${abs}`);
    return { path: abs, name: basename(abs, extname(abs)), builtin: false };
  }
  const name = safeName(raw);
  if (!name) throw new OpsError('E2006', `template 名称非法: ${JSON.stringify(template)}`);
  for (const dir of templateDirs(config)) {
    for (const ext of ['.docx', '.doc']) {
      const file = join(dir, `${name}${ext}`);
      if (await exists(file)) {
        return { path: file, name, builtin: dir.startsWith(join(PACKAGE_DIR, 'templates')) };
      }
    }
  }
  throw new OpsError(
    'E2007',
    `模板不存在: ${name}`,
    `内置目录：${join(PACKAGE_DIR, 'templates', 'docs')}；用户目录：${join(templatesRoot(config), 'docs')}`,
  );
}

// ───────────────────────── 单文件排版 ─────────────────────────

function errorWithPartial(parsed, backup, fallback) {
  const source = parsed?.error ?? {};
  const result = parsed?.result ?? {};
  return {
    ok: false,
    error: {
      code: String(source.code ?? 'E1051'),
      message: String(source.message ?? '排版引擎执行失败'),
      ...(source.hint ? { hint: String(source.hint) } : {}),
      ...(Array.isArray(result.applied) ? { applied: result.applied } : {}),
      ...(result.saved !== undefined ? { saved: result.saved === true } : {}),
      ...(result.outputPath ? { outputPath: result.outputPath } : {}),
      ...(backup ? { backup } : {}),
      ...(fallback ? { engineFallback: 'com → python' } : {}),
    },
  };
}

async function prepareBackup(abs, config) {
  const result = await backupFile(abs, { copies: config?.backupCopies, root: resolveBoundaryRoot(config) });
  if (!result.ok) {
    throw new OpsError(
      'E2008',
      `修改前备份失败，已中止排版: ${result.error.message}`,
      result.error.hint ?? '请检查 runtime\\backups 是否可写（磁盘空间/权限）',
    );
  }
  return { path: result.result.backupPath, id: result.result.id, time: result.result.time };
}

function specFromArgs(args) {
  if (isPlainObject(args?.spec)) return args.spec;
  const spec = {};
  for (const key of ['rules', 'styleSet', 'template', 'placeholders', 'scope']) {
    if (args?.[key] !== undefined) spec[key] = args[key];
  }
  return spec;
}

function summarizeApplied(result) {
  if (Array.isArray(result?.applied)) return result.applied;
  return [];
}

/**
 * office_format / suite.format：单文件排版。
 * @param {object} ctx DSH 插件上下文
 * @param {object} config 有效配置
 * @param {object} args {path, spec?, rules?, styleSet?, template?, placeholders?, engine?, scope?}
 * @param {object} exec 工具/命令执行上下文
 */
export async function formatOp(ctx, config = {}, args = {}, exec = {}) {
  try {
    const rawPath = typeof args?.path === 'string' ? args.path.trim() : '';
    if (!rawPath) throw new OpsError('E1041', '缺少 path');
    const root = resolveBoundaryRoot(config);
    const abs = assertInsideRoot(await resolveInputPath(ctx, exec, rawPath), { write: true, root });
    const ext = extname(abs).toLowerCase();
    if (familyOfExtension(ext) !== 'word') {
      throw new OpsError('E2006', `排版仅支持 Word 文档（.doc/.docx），收到: ${ext || '(无扩展名)'}`, 'Excel 排版不在 P2 范围');
    }
    const info = await stat(abs).catch(() => null);
    if (info?.isDirectory()) {
      throw new OpsError('E2006', `path 是目录，不能做单文件排版: ${abs}`, '批量排版请使用 suite.format.batch / batchFormatOp');
    }

    const rawSpec = specFromArgs(args);
    const spec = validateFormatSpec(rawSpec);

    // styleSet 规则在前，显式 rules 在后（后者覆盖前者的同目标设置）
    const rules = [];
    let styleSetInfo = null;
    if (spec.styleSet !== undefined) {
      styleSetInfo = await resolveStyleSet(spec.styleSet, config);
      rules.push(...styleSetInfo.rules);
    }
    rules.push(...spec.rules);

    let templateInfo = null;
    if (spec.template !== undefined) {
      templateInfo = await resolveTemplate(spec.template, ctx, config, exec);
      if (templateInfo.path.toLowerCase() === abs.toLowerCase()) {
        throw new OpsError('E2006', `模板与目标路径相同，拒绝覆盖模板: ${abs}`, '请把 template 指向骨架文件，path 指向输出文件');
      }
    }
    if (!rules.length && !templateInfo) {
      throw new OpsError('E2006', '未提供任何排版规则（rules / styleSet / template 至少其一）', '示例：{"rules":[{"target":"body","font":{"name":"宋体","sizePt":12}}]}');
    }
    if (!info && !templateInfo) {
      throw new OpsError('E1023', `文件不存在: ${abs}`, '套用模板时目标可以不存在（将从模板新建）；纯规则排版需目标已存在');
    }

    // 目标存在才备份；模板新建时无原文件可备份
    throwIfAborted(exec?.signal);
    const backup = info ? await prepareBackup(abs, config) : null;
    const decided = await decideOfficeEngine(ctx, config, exec, args?.engine, 'word', ext);
    throwIfAborted(exec?.signal);

    const payload = {
      mode: 'format',
      path: abs,
      ext,
      app: 'word',
      spec: {
        rules: templateInfo ? [] : rules,
        scope: spec.scope,
        styleSetId: styleSetInfo?.id ?? null,
      },
      templatePath: templateInfo?.path ?? null,
      placeholders: spec.placeholders ?? {},
      outputPath: abs,
    };
    const parsed = await executeOfficeEngine(ctx, config, exec, decided.engine, decided.python, payload);
    if (!parsed.ok) {
      const failure = errorWithPartial(parsed, backup, decided.fallback);
      if (backup) {
        failure.error.hint = [failure.error.hint, `修改前备份保留在：${backup.path}`].filter(Boolean).join('；');
      }
      return failure;
    }

    const result = parsed.result ?? {};
    return ok({
      outputPath: result.outputPath ?? abs,
      engine: decided.engine,
      ...(decided.fallback ? { engineFallback: 'com → python' } : {}),
      backup,
      styleSet: styleSetInfo?.id ?? null,
      template: templateInfo?.name ?? null,
      applied: summarizeApplied(result),
      stats: result.stats ?? {},
      ...(Array.isArray(result.warnings) && result.warnings.length ? { warnings: result.warnings } : {}),
    });
  } catch (err) {
    if (isAbortError(err)) throw err;
    const payload = toErrorPayload(err);
    return fail(payload.code, payload.message, payload.hint);
  }
}

// ───────────────────────── 批量排版 ─────────────────────────

/** glob → 正则（仅支持 * 与 ?，忽略大小写）。 */
function globToRegExp(glob) {
  const escaped = String(glob).replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
  return new RegExp(`^${escaped}$`, 'i');
}

/** 扫描批量排版输入目录（不递归；排除 ~$ 与 .bak；可按 glob 过滤）。 */
export async function scanFormatFiles(dir, glob) {
  let dirents;
  try {
    dirents = await readdir(dir, { withFileTypes: true });
  } catch (err) {
    if (err?.code === 'ENOENT') throw new OpsError('E1023', `目录不存在: ${dir}`);
    if (err?.code === 'EACCES' || err?.code === 'EPERM') throw new OpsError('E1022', `无权限读取目录: ${dir}`);
    throw new OpsError('E1099', `读取目录失败: ${String(err?.message ?? err)}`);
  }
  const matcher = glob ? globToRegExp(glob) : null;
  const files = [];
  for (const dirent of dirents) {
    if (!dirent.isFile()) continue;
    if (/^~\$/.test(dirent.name) || /\.(bak|tmp)$/i.test(dirent.name)) continue;
    const ext = extname(dirent.name).toLowerCase();
    if (ext !== '.doc' && ext !== '.docx') continue;
    if (matcher && !matcher.test(dirent.name)) continue;
    files.push(resolve(dir, dirent.name));
  }
  files.sort((a, b) => a.localeCompare(b, 'zh-Hans-CN', { numeric: true }));
  return files.slice(0, MAX_BATCH_FILES);
}

function resolveBatchParallel(value, engine) {
  if (value === 'auto') return engine === 'python' ? 4 : 1;
  const n = Number(value);
  if (Number.isInteger(n) && n >= 1 && n <= 4) return n;
  return 1;
}

async function runBatchFormat({ ctx, config, task, files, baseArgs, batchExec, store, parallel }) {
  const log = (level, message) => {
    try { store.appendLog(task.id, level, message); } catch { /* 任务可能已被清理 */ }
  };
  try {
    store.setState(task.id, 'running');
    const total = files.length;
    let done = 0;
    const written = [];
    const failures = [];

    const work = async (file) => {
      throwIfAborted(task.controller.signal);
      store.tick(task.id, { currentFile: basename(file) });
      const result = await formatOp(ctx, config, { ...baseArgs, path: file }, batchExec);
      done += 1;
      if (result.ok) {
        written.push({
          file,
          outputPath: result.result.outputPath,
          engine: result.result.engine,
          backup: result.result.backup?.path ?? null,
        });
        log('info', `[${done}/${total}] 完成 ${basename(file)}（${result.result.engine}）`);
      } else {
        failures.push({ file, code: result.error.code, message: result.error.message });
        log('error', `[${done}/${total}] 失败 ${basename(file)} ${result.error.code}: ${result.error.message}`);
      }
      store.tick(task.id, { filesDone: done });
    };

    if (parallel <= 1) {
      for (const file of files) await work(file);
    } else {
      let cursor = 0;
      const worker = async () => {
        for (;;) {
          const index = cursor;
          cursor += 1;
          if (index >= files.length) return;
          await work(files[index]);
        }
      };
      await Promise.all(Array.from({ length: Math.min(parallel, files.length) }, worker));
    }

    if (failures.length && !written.length) {
      store.fail(task.id, {
        code: 'E2009',
        message: `全部 ${total} 个文件排版失败`,
        hint: failures[0] ? `${failures[0].code}: ${failures[0].message}` : undefined,
      });
      return;
    }
    store.finish(task.id, {
      filesTotal: total,
      filesDone: total,
      written: written.length,
      failed: failures.length,
      failures: failures.slice(0, 100),
      outputs: written.slice(0, 100),
    });
    log('info', `批量排版完成：成功 ${written.length}，失败 ${failures.length}`);
  } catch (err) {
    if (isAbortError(err) || task.controller.signal.aborted) {
      const current = store.get(task.id);
      if (current && !['cancelled', 'done', 'failed'].includes(current.state)) {
        try { store.cancel(task.id); } catch { /* 已终态 */ }
      }
      log('warn', '批量排版已取消');
      return;
    }
    const payload = toErrorPayload(err);
    try { store.fail(task.id, payload); } catch { /* 已终态 */ }
    log('error', `批量排版失败 ${payload.code}: ${payload.message}`);
  }
}

/**
 * suite.format.batch：批量排版（异步任务，kind='office.batch'）。
 * @param {object} ctx
 * @param {object} config
 * @param {object} args {dir?, paths?, glob?, spec?/rules?/styleSet?/template?/placeholders?, engine?, parallel?}
 * @param {object} exec
 * @param {import('../tasks.js').TaskStore} [store]
 * @returns {Promise<{ok:true,result:{taskId,filesTotal,files,state}}|{ok:false,error:object}>}
 */
export async function batchFormatOp(ctx, config = {}, args = {}, exec = {}, store = defaultTasks) {
  try {
    const root = resolveBoundaryRoot(config);
    let files = [];
    let inputLabel = '';

    if (Array.isArray(args?.paths) && args.paths.length) {
      for (const raw of args.paths) {
        if (typeof raw !== 'string' || !raw.trim()) throw new OpsError('E1041', 'paths 中存在空路径');
        const abs = assertInsideRoot(await resolveInputPath(ctx, exec, raw), { write: true, root });
        const ext = extname(abs).toLowerCase();
        if (ext !== '.doc' && ext !== '.docx') {
          throw new OpsError('E2006', `批量排版仅支持 .doc/.docx: ${abs}`);
        }
        if (!(await exists(abs))) throw new OpsError('E1023', `文件不存在: ${abs}`);
        files.push(abs);
      }
      inputLabel = `${files.length} 个指定文件`;
    } else {
      const rawDir = typeof args?.dir === 'string' ? args.dir.trim() : (typeof args?.path === 'string' ? args.path.trim() : '');
      if (!rawDir) throw new OpsError('E1041', '缺少 dir（批量排版目录）或 paths 数组');
      const dirAbs = assertInsideRoot(await resolveInputPath(ctx, exec, rawDir), { write: true, root });
      const info = await stat(dirAbs).catch(() => null);
      if (!info) throw new OpsError('E1023', `目录不存在: ${dirAbs}`);
      if (!info.isDirectory()) {
        throw new OpsError('E2006', `批量排版的 dir 必须是目录: ${dirAbs}`, '单文件请使用 formatOp / suite.format');
      }
      files = await scanFormatFiles(dirAbs, typeof args?.glob === 'string' ? args.glob : undefined);
      inputLabel = dirAbs;
    }

    if (!files.length) {
      throw new OpsError('E1023', '没有可排版的 Word 文件（.doc/.docx；已排除 ~$ 与 .bak）', `输入：${inputLabel}`);
    }

    // 任务创建前先校验规格与资源，避免异步任务一开始就失败
    const spec = validateFormatSpec(specFromArgs(args));
    const rules = [];
    if (spec.styleSet !== undefined) {
      const styleSetInfo = await resolveStyleSet(spec.styleSet, config);
      rules.push(...styleSetInfo.rules);
    }
    rules.push(...spec.rules);
    if (spec.template !== undefined) await resolveTemplate(spec.template, ctx, config, exec);
    if (!rules.length && spec.template === undefined) {
      throw new OpsError('E2006', '未提供任何排版规则（rules / styleSet / template 至少其一）');
    }

    const engineHint = typeof args?.engine === 'string' ? args.engine : (config.officeEngine ?? 'auto');
    const parallel = resolveBatchParallel(args?.parallel ?? config.batchParallel ?? 1, engineHint);
    const baseArgs = { ...specFromArgs(args) };
    if (typeof args?.engine === 'string') baseArgs.engine = args.engine;

    const task = store.create({
      kind: 'office.batch',
      mode: 'format',
      engine: engineHint === 'auto' ? null : engineHint,
      input: inputLabel,
      output: null,
    });
    store.appendLog(task.id, 'info', `批量排版任务创建：${files.length} 个文件（parallel=${parallel}）`);
    store.appendLog(task.id, 'info', `规则：${rules.length} 条${spec.template ? `；模板：${spec.template}` : ''}`);
    store.tick(task.id, { filesTotal: files.length, filesDone: 0 });

    const batchExec = { agent: exec?.agent, signal: task.controller.signal };
    setTimeout(() => {
      void runBatchFormat({ ctx, config, task, files, baseArgs, batchExec, store, parallel });
    }, 0);

    return ok({
      taskId: task.id,
      filesTotal: files.length,
      files: files.map((file) => basename(file)),
      state: 'queued',
    });
  } catch (err) {
    if (isAbortError(err)) throw err;
    const payload = toErrorPayload(err);
    return fail(payload.code, payload.message, payload.hint);
  }
}

/** suite.office.status（只过滤 kind='office.batch'）。 */
export function batchFormatStatus(store, args = {}) {
  const taskId = typeof args?.taskId === 'string' ? args.taskId.trim() : '';
  if (!taskId) {
    const all = store?.tasks instanceof Map ? [...store.tasks.values()] : [];
    const tasks = all
      .filter((task) => task.kind === 'office.batch')
      .sort((a, b) => (a.createdAt < b.createdAt ? 1 : a.createdAt > b.createdAt ? -1 : 0))
      .map((task) => ({ kind: task.kind, ...store.summary(task.id) }));
    return ok({ tasks });
  }
  const task = store.get(taskId);
  if (!task || task.kind !== 'office.batch') {
    return fail('TASK_NOT_FOUND', `批量排版任务不存在或已过期: ${taskId}`);
  }
  return ok(store.snapshot(taskId));
}

/** suite.office.cancel。 */
export function batchFormatCancel(store, args = {}) {
  const taskId = typeof args?.taskId === 'string' ? args.taskId.trim() : '';
  if (!taskId) return fail('E1041', '缺少 taskId');
  const task = store.get(taskId);
  if (!task || task.kind !== 'office.batch') {
    return fail('TASK_NOT_FOUND', `批量排版任务不存在或已过期: ${taskId}`);
  }
  try {
    return ok(store.cancel(taskId));
  } catch (err) {
    const payload = toErrorPayload(err);
    return fail(payload.code, payload.message, payload.hint);
  }
}
