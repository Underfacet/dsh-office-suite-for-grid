/**
 * Office 结构化编辑（docs/P2-设计.md §2.3 / §3.1 / §3.2）
 *
 * 调用链：boundary 校验（写）→ 指令全量校验（E2005）→ 修改前备份（E2008 中止）
 *        → 引擎选择（COM 优先 → venv Python）→ 执行 → 结果规整。
 *
 * 引擎输入：office-com.ps1 / office-python.py 统一 -Json '<payload>'，
 *          payload = { mode:'edit', path, ext, app, ops, outputPath, create }。
 * 引擎输出：{ok:true,result:{outputPath,applied,stats}} | {ok:false,error,result?}（op 级失败带 applied/saved）。
 * 降级策略：只在探测/选择阶段降级；COM 执行失败不自动切 python 重试（避免二次修改）。
 *
 * 导出：
 *   editOp(ctx, config, args, exec)                 args: {path, ops, engine?}
 *   validateEditOps(path, ext, ops)
 *   chooseOfficeEngine(probe, config, engineOverride, options)
 *   decideOfficeEngine(ctx, config, exec, engineOverride, app, ext)
 *   executeOfficeEngine(ctx, config, exec, engine, python, payload)
 *   familyOfExtension(ext) / appForFamily(family) / comUsableFromProbe(probe, app)
 */
import { stat } from 'node:fs/promises';
import { extname, join } from 'node:path';
import {
  OpsError,
  PACKAGE_DIR,
  abortError,
  fail,
  isAbortError,
  ok,
  parseJson,
  resolveInputPath,
  throwIfAborted,
  toErrorPayload,
} from '../ops.js';
import { assertInsideRoot, resolveBoundaryRoot } from '../boundary.js';
import { runExe, runPs1 } from '../engines.js';
import { probeEngines } from './cad.js';
import { backupFile } from './backup.js';
import { ensureVenv } from './venv.js';

/** Word 文件扩展名。 */
export const WORD_EXTENSIONS = Object.freeze(['.doc', '.docx']);
/** Excel 文件扩展名。 */
export const EXCEL_EXTENSIONS = Object.freeze(['.xls', '.xlsx']);

const WORD_ACTIONS = new Set(['create', 'paragraph', 'heading', 'table', 'saveAs']);
const EXCEL_ACTIONS = new Set(['create', 'cell', 'row', 'column', 'sheet', 'saveAs']);
const PARAGRAPH_OPS = new Set(['setText', 'setStyle', 'insert', 'delete', 'append']);
const HEADING_OPS = new Set(['setText', 'setStyle', 'insert', 'delete']);
const TABLE_OPS = new Set(['setCell', 'addRow', 'deleteRow', 'addColumn', 'deleteColumn']);
const SHEET_OPS = new Set(['rename', 'add', 'delete']);
/** 单次编辑最多指令数（防御异常输入）。 */
const MAX_OPS = 500;
/** 引擎默认超时（毫秒）。 */
const DEFAULT_OFFICE_TIMEOUT_MS = 180_000;

// ───────────────────────── 扩展名 / 引擎工具 ─────────────────────────

/** 扩展名 → 'word' | 'excel' | null。 */
export function familyOfExtension(ext) {
  const value = String(ext ?? '').toLowerCase();
  if (WORD_EXTENSIONS.includes(value)) return 'word';
  if (EXCEL_EXTENSIONS.includes(value)) return 'excel';
  return null;
}

/** 家族 → probe.officeCom 字段名。 */
export function appForFamily(family) {
  return family === 'excel' ? 'excel' : 'word';
}

function appLabel(app) {
  return app === 'excel' ? 'Excel' : 'Word';
}

/**
 * probe 结果 → COM 是否可用：
 * - officeCom.<app>.comUsable 明确 true/false 时以实测为准；
 * - comUsable 为 null（未 -Refresh 强刷）时回退到 present（安装键）；
 * - 兼容 v0.1 probe（仅顶层 ProgID.present）。
 */
export function comUsableFromProbe(probe, app) {
  const node = probe?.officeCom?.[app];
  if (node && typeof node === 'object') {
    if (node.comUsable === true) return true;
    if (node.comUsable === false) return false;
    return node.present === true;
  }
  const legacy = probe?.[app === 'excel' ? 'Excel.Application' : 'Word.Application'];
  return legacy?.present === true;
}

function normalizeEngine(value) {
  if (value === undefined || value === null || value === '') return null;
  const text = String(value).trim();
  if (text === 'auto' || text === 'com' || text === 'python') return text;
  throw new OpsError('PARAM_INVALID', `engine 必须为 'auto' | 'com' | 'python'，收到: ${JSON.stringify(value)}`);
}

/**
 * 选择 Office 编辑引擎（纯函数，不探测、不建 venv）。
 * @param {object} probe probe.ps1 结果
 * @param {object} config 有效配置（officeEngine）
 * @param {string} [engineOverride] 工具参数 engine（优先级高于配置）
 * @param {{app?:'word'|'excel', venvHealthy?:boolean|null}} [options]
 * @returns {{engine:'com'|'python', configured:string, app:string, comAvailable:boolean, pythonHealthy:boolean|null, needsVenv:boolean, fallback:boolean}}
 * @throws {OpsError} PARAM_INVALID / E2001
 */
export function chooseOfficeEngine(probe, config = {}, engineOverride, options = {}) {
  const app = options.app === 'excel' ? 'excel' : 'word';
  const override = normalizeEngine(engineOverride);
  const configured = override ?? normalizeEngine(config?.officeEngine) ?? 'auto';
  const comAvailable = comUsableFromProbe(probe, app);
  const venvHealthy = options.venvHealthy === true ? true : (options.venvHealthy === false ? false : null);

  if (configured === 'com') {
    if (!comAvailable) {
      throw new OpsError(
        'E2001',
        `${appLabel(app)} COM 不可用（未安装或无法实例化）`,
        `可改用 engine='python'，或在设置中把 officeEngine 改为 'auto'/'python'`,
      );
    }
    return { engine: 'com', configured, app, comAvailable, pythonHealthy: venvHealthy, needsVenv: false, fallback: false };
  }

  if (configured === 'python') {
    if (venvHealthy === false) {
      throw new OpsError('E2001', 'Python venv 不可用', "可改用 engine='com'（需本机装有 Office）或在诊断页重建 venv");
    }
    return {
      engine: 'python',
      configured,
      app,
      comAvailable,
      pythonHealthy: venvHealthy,
      needsVenv: venvHealthy !== true,
      fallback: false,
    };
  }

  // auto：COM 优先 → Python venv
  if (comAvailable) {
    return { engine: 'com', configured, app, comAvailable, pythonHealthy: venvHealthy, needsVenv: false, fallback: false };
  }
  if (venvHealthy === false) {
    throw new OpsError(
      'E2001',
      `${appLabel(app)} COM 不可用，且 Python venv 不健康`,
      '请安装 Microsoft Office，或运行诊断页/venv ensure 重建 Python 依赖',
    );
  }
  return {
    engine: 'python',
    configured,
    app,
    comAvailable,
    pythonHealthy: venvHealthy,
    needsVenv: venvHealthy !== true,
    fallback: true,
  };
}

/**
 * 探测 + 选择 + 必要时 ensure venv（会创建 runtime\venv）。
 * @returns {Promise<{engine:'com'|'python', python:string|null, fallback:boolean, probe:object, venv:object|null}>}
 * @throws {OpsError} E2001 / E2002 / E2003 / PARAM_INVALID
 */
export async function decideOfficeEngine(ctx, config = {}, exec = {}, engineOverride, app = 'word', ext = '') {
  const override = normalizeEngine(engineOverride);
  const probeResult = await probeEngines(ctx, config, { signal: exec?.signal });
  const probe = probeResult.ok ? probeResult.result : {};
  if (!probeResult.ok) {
    console.warn(`[office-suite] 引擎探测失败，按无 COM 处理：${probeResult.error?.message ?? ''}`);
  }

  const decision = chooseOfficeEngine(probe, config, override, { app, venvHealthy: null });
  let python = null;
  let venv = null;

  if (decision.engine === 'python') {
    const legacy = String(ext ?? '').toLowerCase();
    if (legacy === '.doc' || legacy === '.xls') {
      throw new OpsError(
        'E2001',
        `旧格式 ${legacy} 无法使用 Python 引擎（python-docx/openpyxl 不支持二进制旧格式）`,
        "请安装 Microsoft Office 后使用 engine='com'，或先把文件另存为 .docx/.xlsx",
      );
    }
    const ensured = await ensureVenv(ctx, config, exec);
    if (!ensured.ok) {
      throw new OpsError(ensured.error.code ?? 'E2002', ensured.error.message ?? 'venv 创建失败', ensured.error.hint);
    }
    python = ensured.result.python;
    venv = ensured.result;
  }

  return { engine: decision.engine, python, fallback: decision.fallback, probe, venv };
}

/** 运行 Office 引擎（COM ps1 / venv Python），返回引擎 JSON（含 ok:false 情形）。 */
export async function executeOfficeEngine(ctx, config, exec, engine, python, payload) {
  const timeoutMs = Number.isFinite(config?.officeTimeoutMs) ? config.officeTimeoutMs : DEFAULT_OFFICE_TIMEOUT_MS;

  let execution;
  if (engine === 'python') {
    if (!python) throw new OpsError('E2001', 'Python 引擎路径缺失');
    const script = join(PACKAGE_DIR, 'engines', 'office-python.py');
    execution = await runExe(ctx, python, [script, '-Json', JSON.stringify(payload)], {
      signal: exec?.signal,
      timeoutMs,
      graceMs: 120_000,
    });
  } else {
    execution = await runPs1(ctx, 'office-com.ps1', payload, {
      signal: exec?.signal,
      timeoutMs,
      graceMs: 120_000,
    });
  }

  if (exec?.signal?.aborted) throw abortError('任务已取消');
  const parsed = parseJson(execution.stdout);
  const stderr = String(execution.stderr ?? '').trim();
  if (!parsed) {
    const reason = execution.signal
      ? `引擎被终止（${execution.signal ?? '超时'}）`
      : `引擎未返回可解析的 JSON（exit ${execution.exitCode ?? 'null'}）`;
    throw new OpsError(
      'E1051',
      `${engine} 引擎执行失败：${reason}`,
      (stderr || String(execution.stdout ?? '')).slice(0, 600) || undefined,
    );
  }
  if (typeof parsed.ok !== 'boolean') {
    throw new OpsError('E1051', `${engine} 引擎返回结构非法（缺少 ok 字段）`, stderr.slice(0, 600) || undefined);
  }
  return parsed;
}

function errorWithPartial(parsed, backup, fallback) {
  const source = parsed?.error ?? {};
  const result = parsed?.result ?? {};
  return {
    ok: false,
    error: {
      code: String(source.code ?? 'E1051'),
      message: String(source.message ?? 'Office 引擎执行失败'),
      ...(source.hint ? { hint: String(source.hint) } : {}),
      ...(Array.isArray(result.applied) ? { applied: result.applied } : {}),
      ...(result.saved !== undefined ? { saved: result.saved === true } : {}),
      ...(result.outputPath ? { outputPath: result.outputPath } : {}),
      ...(backup ? { backup } : {}),
      ...(fallback ? { engineFallback: 'com → python' } : {}),
    },
  };
}

// ───────────────────────── 指令校验（§3.2） ─────────────────────────

function isPlainObject(value) {
  return !!value && typeof value === 'object' && !Array.isArray(value);
}

function opError(index, message, hint) {
  return new OpsError('E2005', `ops[${index}] ${message}`, hint);
}

function requirePositiveInt(value, field, index, { optional = false } = {}) {
  if (optional && (value === undefined || value === null || value === '')) return;
  if (!Number.isInteger(value) || value < 1) {
    throw opError(index, `${field} 必须为正整数（1 起），收到: ${JSON.stringify(value)}`);
  }
}

function requireString(value, field, index, { allowEmpty = false } = {}) {
  if (typeof value !== 'string') {
    throw opError(index, `${field} 必须是字符串，收到: ${JSON.stringify(value)}`);
  }
  if (!allowEmpty && !value.trim()) {
    throw opError(index, `${field} 不能为空`);
  }
}

function requireNonEmptyString(value, field, index) {
  requireString(value, field, index, { allowEmpty: false });
}

/** 列字母（A–XFD）→ 列号；非法返回 null。 */
function columnNumber(letters) {
  const text = String(letters ?? '').toUpperCase();
  if (!/^[A-Z]{1,3}$/.test(text)) return null;
  let value = 0;
  for (const char of text) value = value * 26 + (char.charCodeAt(0) - 64);
  return value >= 1 && value <= 16384 ? value : null;
}

function requireColumn(value, field, index) {
  if (Number.isInteger(value) && value >= 1 && value <= 16384) return value;
  if (typeof value === 'string') {
    const upper = value.trim().toUpperCase();
    if (columnNumber(upper)) return upper;
  }
  throw opError(index, `${field} 必须是 1–16384 的列号或 A–XFD 列字母，收到: ${JSON.stringify(value)}`);
}

function validateWordOp(op, index) {
  switch (op.action) {
    case 'create': {
      if (index !== 0) throw opError(index, 'create 只能作为 ops[0]');
      if (op.kind !== 'docx') {
        throw opError(index, `Word 文档 create.kind 必须为 'docx'，收到: ${JSON.stringify(op.kind)}`);
      }
      return { ...op };
    }
    case 'paragraph': {
      if (!PARAGRAPH_OPS.has(op.op)) {
        throw opError(index, `paragraph.op 必须为 ${[...PARAGRAPH_OPS].join('|')}，收到: ${JSON.stringify(op.op)}`);
      }
      if (op.op === 'append') {
        requireString(op.text, 'text', index, { allowEmpty: true });
        return { ...op };
      }
      requirePositiveInt(op.at, 'at', index);
      if (op.op === 'setText' || op.op === 'insert') requireString(op.text, 'text', index, { allowEmpty: true });
      if (op.op === 'setStyle') requireNonEmptyString(op.style, 'style', index);
      return { ...op };
    }
    case 'heading': {
      if (!HEADING_OPS.has(op.op)) {
        throw opError(index, `heading.op 必须为 ${[...HEADING_OPS].join('|')}，收到: ${JSON.stringify(op.op)}`);
      }
      if (!Number.isInteger(op.level) || op.level < 1 || op.level > 3) {
        throw opError(index, `heading.level 必须为 1–3 的整数，收到: ${JSON.stringify(op.level)}`);
      }
      requirePositiveInt(op.at, 'at', index);
      if (op.op === 'setText' || op.op === 'insert') requireString(op.text, 'text', index, { allowEmpty: true });
      if (op.op === 'setStyle' && op.style !== undefined) requireNonEmptyString(op.style, 'style', index);
      return { ...op };
    }
    case 'table': {
      if (!TABLE_OPS.has(op.op)) {
        throw opError(index, `table.op 必须为 ${[...TABLE_OPS].join('|')}，收到: ${JSON.stringify(op.op)}`);
      }
      requirePositiveInt(op.table, 'table', index);
      switch (op.op) {
        case 'setCell':
          requirePositiveInt(op.row, 'row', index);
          requirePositiveInt(op.col, 'col', index);
          requireString(op.text, 'text', index, { allowEmpty: true });
          break;
        case 'deleteRow':
          requirePositiveInt(op.row, 'row', index);
          break;
        case 'addRow': {
          if (op.at !== undefined && op.at !== null && op.at !== 'end') requirePositiveInt(op.at, "at（'end' 或行号）", index);
          break;
        }
        case 'deleteColumn':
          requirePositiveInt(op.col, 'col', index);
          break;
        case 'addColumn': {
          if (op.col !== undefined && op.col !== null) requirePositiveInt(op.col, 'col', index);
          if (op.at !== undefined && op.at !== null && op.at !== 'end') requirePositiveInt(op.at, "at（'end' 或列号）", index);
          break;
        }
        default:
          break;
      }
      return { ...op };
    }
    case 'saveAs': {
      requireNonEmptyString(op.path, 'path', index);
      if (familyOfExtension(extname(op.path)) !== 'word') {
        throw opError(index, `saveAs 目标必须是 .doc/.docx，收到: ${op.path}`);
      }
      return { ...op };
    }
    default:
      throw opError(index, `Word 文档不支持指令: ${JSON.stringify(op.action)}`, `可用指令：${[...WORD_ACTIONS].join(', ')}`);
  }
}

function validateExcelOp(op, index) {
  switch (op.action) {
    case 'create': {
      if (index !== 0) throw opError(index, 'create 只能作为 ops[0]');
      if (op.kind !== 'xlsx') {
        throw opError(index, `Excel 文档 create.kind 必须为 'xlsx'，收到: ${JSON.stringify(op.kind)}`);
      }
      return { ...op };
    }
    case 'cell': {
      if (op.sheet !== undefined) requireNonEmptyString(op.sheet, 'sheet', index);
      if (typeof op.at !== 'string' || !/^[A-Za-z]{1,3}[1-9]\d{0,6}$/.test(op.at.trim())) {
        throw opError(index, `cell.at 必须是 A1 形式的单元格地址，收到: ${JSON.stringify(op.at)}`);
      }
      const at = op.at.trim().toUpperCase();
      const letters = at.match(/^[A-Z]+/)?.[0] ?? '';
      if (!columnNumber(letters)) throw opError(index, `cell.at 列超出 A–XFD: ${at}`);
      const hasValue = Object.prototype.hasOwnProperty.call(op, 'value') && op.value !== undefined;
      const hasFormula = op.formula !== undefined && op.formula !== null && op.formula !== '';
      if (hasValue === hasFormula) {
        throw opError(index, 'cell 必须且只能提供 value 或 formula 之一');
      }
      if (hasFormula) {
        if (typeof op.formula !== 'string' || !op.formula.trim().startsWith('=')) {
          throw opError(index, `cell.formula 必须以 '=' 开头，收到: ${JSON.stringify(op.formula)}`);
        }
      } else if (!['string', 'number', 'boolean'].includes(typeof op.value)) {
        throw opError(index, `cell.value 必须是字符串/数字/布尔值，收到: ${JSON.stringify(op.value)}`);
      }
      return { ...op, at };
    }
    case 'row': {
      if (!['insert', 'delete'].includes(op.op)) {
        throw opError(index, `row.op 必须为 'insert' | 'delete'，收到: ${JSON.stringify(op.op)}`);
      }
      if (op.sheet !== undefined) requireNonEmptyString(op.sheet, 'sheet', index);
      requirePositiveInt(op.at, 'at', index);
      return { ...op };
    }
    case 'column': {
      if (!['insert', 'delete'].includes(op.op)) {
        throw opError(index, `column.op 必须为 'insert' | 'delete'，收到: ${JSON.stringify(op.op)}`);
      }
      if (op.sheet !== undefined) requireNonEmptyString(op.sheet, 'sheet', index);
      const at = requireColumn(op.at, 'at', index);
      return { ...op, at };
    }
    case 'sheet': {
      if (!SHEET_OPS.has(op.op)) {
        throw opError(index, `sheet.op 必须为 ${[...SHEET_OPS].join('|')}，收到: ${JSON.stringify(op.op)}`);
      }
      requireNonEmptyString(op.name, 'name', index);
      if (op.op === 'rename') requireNonEmptyString(op.newName, 'newName', index);
      return { ...op };
    }
    case 'saveAs': {
      requireNonEmptyString(op.path, 'path', index);
      if (familyOfExtension(extname(op.path)) !== 'excel') {
        throw opError(index, `saveAs 目标必须是 .xls/.xlsx，收到: ${op.path}`);
      }
      return { ...op };
    }
    default:
      throw opError(index, `Excel 文档不支持指令: ${JSON.stringify(op.action)}`, `可用指令：${[...EXCEL_ACTIONS].join(', ')}`);
  }
}

/**
 * 全量校验编辑指令（按扩展名限定指令集；不预读文件做索引越界校验）。
 * @param {string} path 目标文件（用于错误提示）
 * @param {string} ext 小写扩展名
 * @param {object[]} ops 指令数组
 * @returns {object[]} 归一化后的指令（浅拷贝，Excel 单元格/列字母转大写）
 * @throws {OpsError} E2005 EDIT_OP_INVALID
 */
export function validateEditOps(path, ext, ops) {
  const family = familyOfExtension(ext);
  if (!family) {
    throw new OpsError('E1042', `不支持的文件类型: ${ext || '(无扩展名)'}`, '仅支持 .doc/.docx/.xls/.xlsx');
  }
  if (!Array.isArray(ops) || ops.length === 0) {
    throw new OpsError('E2005', 'ops 必须是非空数组', '示例：{"ops":[{"action":"paragraph","op":"setText","at":1,"text":"新文本"}]}');
  }
  if (ops.length > MAX_OPS) {
    throw new OpsError('E2005', `ops 数量超过上限（${MAX_OPS}）`, `收到 ${ops.length} 条；请拆分为多次编辑`);
  }
  const normalized = [];
  const seenCreate = [];
  for (let index = 0; index < ops.length; index++) {
    const op = ops[index];
    if (!isPlainObject(op)) throw opError(index, '指令必须是对象', `文件：${path}`);
    if (typeof op.action !== 'string' || !op.action.trim()) {
      throw opError(index, '缺少 action', `文件：${path}`);
    }
    if (op.action === 'create') seenCreate.push(index);
    normalized.push(family === 'word' ? validateWordOp(op, index) : validateExcelOp(op, index));
  }
  if (seenCreate.length > 1) {
    throw new OpsError('E2005', 'create 指令最多出现一次', `收到 ${seenCreate.length} 次（ops 索引 ${seenCreate.join(', ')}）`);
  }
  return normalized;
}

// ───────────────────────── 路径 / 备份 ─────────────────────────

async function isFile(path) {
  try {
    return (await stat(path)).isFile();
  } catch {
    return false;
  }
}

async function prepareBackup(abs, config) {
  const result = await backupFile(abs, {
    copies: config?.backupCopies,
    root: resolveBoundaryRoot(config),
  });
  if (!result.ok) {
    throw new OpsError(
      'E2008',
      `修改前备份失败，已中止编辑: ${result.error.message}`,
      result.error.hint ?? '请检查 runtime\\backups 是否可写（磁盘空间/权限）',
    );
  }
  return { path: result.result.backupPath, id: result.result.id, time: result.result.time };
}

/** 解析并校验 saveAs 目标（写边界 + 同族扩展名）。 */
async function resolveSaveAsTargets(ctx, config, exec, family, ops, inputAbs) {
  let outputPath = inputAbs;
  for (const op of ops) {
    if (op.action !== 'saveAs') continue;
    const target = assertInsideRoot(await resolveInputPath(ctx, exec, op.path), {
      write: true,
      root: resolveBoundaryRoot(config),
    });
    const targetExt = extname(target).toLowerCase();
    if (familyOfExtension(targetExt) !== family) {
      throw new OpsError('E2005', `saveAs 目标扩展名与源文件不一致: ${target}`);
    }
    op.path = target;
    outputPath = target;
  }
  return outputPath;
}

// ───────────────────────── 主操作 ─────────────────────────

/**
 * office_edit / suite.edit：结构化编辑 Word / Excel。
 * @param {object} ctx DSH 插件上下文
 * @param {object} config 有效配置
 * @param {{path:string, ops:object[], engine?:'auto'|'com'|'python'}} args
 * @param {object} exec 工具/命令执行上下文（agent/signal）
 * @returns {Promise<{ok:true,result:object}|{ok:false,error:object}>}
 */
export async function editOp(ctx, config = {}, args = {}, exec = {}) {
  try {
    const rawPath = typeof args?.path === 'string' ? args.path.trim() : '';
    if (!rawPath) throw new OpsError('E1041', '缺少 path');
    const root = resolveBoundaryRoot(config);
    const abs = assertInsideRoot(await resolveInputPath(ctx, exec, rawPath), { write: true, root });
    const ext = extname(abs).toLowerCase();
    const family = familyOfExtension(ext);
    if (!family) {
      throw new OpsError('E1042', `不支持的文件类型: ${ext || '(无扩展名)'}`, '仅支持 .doc/.docx（Word）与 .xls/.xlsx（Excel）');
    }

    const ops = validateEditOps(abs, ext, args?.ops);
    const outputPath = await resolveSaveAsTargets(ctx, config, exec, family, ops, abs);

    const existsNow = await isFile(abs);
    const hasCreate = ops.some((op) => op.action === 'create');
    if (hasCreate && existsNow) {
      throw new OpsError('E2005', `create 仅用于不存在的文件，目标已存在: ${abs}`, '如需覆盖已有文档请去掉 create 指令');
    }
    if (!existsNow && !hasCreate) {
      throw new OpsError(
        'E1023',
        `文件不存在: ${abs}`,
        "如需新建文档，请在 ops[0] 加入 {action:'create', kind:'docx'} 或 {action:'create', kind:'xlsx'}",
      );
    }

    // 修改前备份（原文件存在时；备份失败即中止）
    throwIfAborted(exec?.signal);
    const backup = existsNow ? await prepareBackup(abs, config) : null;

    const app = appForFamily(family);
    const decided = await decideOfficeEngine(ctx, config, exec, args?.engine, app, ext);
    throwIfAborted(exec?.signal);

    const payload = {
      mode: 'edit',
      path: abs,
      ext,
      app,
      ops,
      outputPath,
      create: hasCreate,
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
      outputPath: result.outputPath ?? outputPath,
      engine: decided.engine,
      ...(decided.fallback ? { engineFallback: 'com → python' } : {}),
      backup,
      applied: Array.isArray(result.applied)
        ? result.applied
        : ops.map((op, index) => ({ op: op.action, index, ok: true })),
      stats: result.stats ?? {},
      ...(Array.isArray(result.warnings) && result.warnings.length ? { warnings: result.warnings } : {}),
      ...(result.saved !== undefined ? { saved: result.saved === true } : {}),
    });
  } catch (err) {
    if (isAbortError(err)) throw err;
    const payload = toErrorPayload(err);
    return fail(payload.code, payload.message, payload.hint);
  }
}
