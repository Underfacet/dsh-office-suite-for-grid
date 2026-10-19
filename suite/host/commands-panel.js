/**
 * 面板远程命令集合（docs/P1-设计.md §2 的 10 个 + docs/P2-设计.md §2.2 增量）
 *
 * 约定：命令名小写、无前导斜杠；参数为单个 JSON 文本（rawInput）；
 * ops 结果统一序列化为 text：{ok:true,result} | {ok:false,error:{code,message,hint?}}；
 * 仅 rawInput 不是单个 JSON 对象时返回 {kind:'error', text}（PARAM_INVALID）。
 *
 * 与对话工具共用同一操作层（ops/*）；所有路径先过 boundary（§5.3）。
 * P2 新增：suite.edit / suite.format / suite.format.batch / suite.office.status /
 *         suite.office.cancel / suite.office.backup.list / suite.office.backup.restore /
 *         suite.templates；suite.draw 扩展电网三图；suite.preview 按 §3.7 契约修正。
 */
import { existsSync, mkdirSync } from 'node:fs';
import { readdir, readFile, stat } from 'node:fs/promises';
import { basename, dirname, extname, join, resolve } from 'node:path';
import { runPs1 } from './engines.js';
import { assertInsideRoot, boundaryInfo, resolveBoundaryRoot } from './boundary.js';
import {
  OpsError,
  PACKAGE_DIR,
  RUNTIME_DIR,
  UI_CONFIG_PATH,
  WRITABLE_CONFIG_KEYS,
  effectiveConfig,
  fail,
  loadUiConfig,
  ok,
  parseJson,
  resolveInputPath,
  saveUiConfig,
  sessionCwd,
  toErrorPayload,
  validateConfigEntry,
} from './ops.js';
import { listDirectory } from './ops/files.js';
import { cadCancel, cadStatus, convertCad, probeEngines } from './ops/cad.js';
import { editOp } from './ops/office-edit.js';
import { batchFormatCancel, batchFormatOp, batchFormatStatus, formatOp } from './ops/office-format.js';
import { backupOp } from './ops/backup.js';
import { expandGridLinks, loadSymbols, validateGridModel } from './ops/grid-model.js';
import { TaskStore, tasks as defaultTasks } from './tasks.js';

/** suite.draw / visio_draw 支持的图类型（P1 流程图 + P2 电网三图）。 */
export const DRAW_KINDS = Object.freeze(['flowchart', 'singleLine', 'distribution', 'networkTopology']);

/** suite.preview 支持的扩展名（officeToPdf OfficeExtension 的子集，§3.7）。 */
const PREVIEW_EXTENSIONS = Object.freeze(['.doc', '.docx', '.xls', '.xlsx']);

/** 用户模板根目录（与 ops/office-format.js 的 templatesRoot 口径一致）。 */
function templatesRoot(config = {}) {
  const custom = typeof config.officeTemplatesDir === 'string' ? config.officeTemplatesDir.trim() : '';
  if (custom) return resolve(custom);
  return join(RUNTIME_DIR, 'templates');
}

/** 列出目录内满足条件的文件（目录不存在返回空数组）。 */
async function listFilesIn(dir, predicate) {
  let entries;
  try {
    entries = await readdir(dir, { withFileTypes: true });
  } catch {
    return [];
  }
  return entries
    .filter((entry) => entry.isFile() && predicate(entry.name))
    .map((entry) => join(dir, entry.name));
}

/**
 * suite.templates：合并列出样式集（styles/*.json）与模板（docs/*.docx）。
 * 优先级（同名先出现者胜）：config.officeTemplatesDir → runtime\templates → suite\templates。
 */
async function listTemplateCatalog(config = {}) {
  const userDir = templatesRoot(config);
  const builtinStyles = join(PACKAGE_DIR, 'templates', 'styles');
  const builtinDocs = join(PACKAGE_DIR, 'templates', 'docs');
  const styleDirs = [...new Set([join(userDir, 'styles'), join(RUNTIME_DIR, 'templates', 'styles'), builtinStyles])];
  const docDirs = [...new Set([join(userDir, 'docs'), join(RUNTIME_DIR, 'templates', 'docs'), builtinDocs])];

  const styles = [];
  const seenStyles = new Set();
  for (const dir of styleDirs) {
    for (const file of await listFilesIn(dir, (name) => name.toLowerCase().endsWith('.json'))) {
      const id = basename(file, extname(file));
      if (seenStyles.has(id)) continue;
      seenStyles.add(id);
      const entry = { id, name: id, file, builtin: dir === builtinStyles, rules: 0 };
      try {
        const parsed = JSON.parse((await readFile(file, 'utf8')).replace(/^\uFEFF/, ''));
        const rules = Array.isArray(parsed) ? parsed : parsed?.rules;
        if (!Array.isArray(rules)) {
          entry.invalid = true;
        } else {
          entry.rules = rules.length;
        }
        if (typeof parsed?.name === 'string' && parsed.name.trim()) entry.name = parsed.name.trim();
        if (typeof parsed?.description === 'string' && parsed.description.trim()) entry.description = parsed.description.trim();
      } catch {
        entry.invalid = true;
      }
      styles.push(entry);
    }
  }

  const templates = [];
  const seenTemplates = new Set();
  for (const dir of docDirs) {
    for (const file of await listFilesIn(dir, (name) => /\.docx?$/i.test(name))) {
      const id = basename(file, extname(file));
      if (seenTemplates.has(id)) continue;
      seenTemplates.add(id);
      templates.push({ id, name: id, file, builtin: dir === builtinDocs, placeholders: [] });
    }
  }

  const hints = [];
  if (!styles.length) {
    hints.push(`未找到样式集 JSON（可放置于 runtime\\templates\\styles\\<id>.json 或随包 suite\\templates\\styles）：${styleDirs.join('；')}`);
  }
  if (!templates.length) {
    hints.push(`未找到模板 docx（可放置于 runtime\\templates\\docs\\<name>.docx 或随包 suite\\templates\\docs）：${docDirs.join('；')}`);
  }
  return ok({
    styles,
    templates,
    userDir,
    ...(hints.length ? { hint: hints.join(' ') } : {}),
  });
}

/** 解析 rawInput：空 → {}；非单个 JSON 对象 → undefined（PARAM_INVALID）。 */
export function parseCommandInput(rawInput) {
  const text = String(rawInput ?? '').trim();
  if (!text) return {};
  try {
    const value = JSON.parse(text);
    return value && typeof value === 'object' && !Array.isArray(value) ? value : undefined;
  } catch {
    return undefined;
  }
}

/** 统一成功返回（含 ops 失败，按 §2 约定统一为 kind:'success' + JSON）。 */
export function commandSuccess(payload) {
  return { kind: 'success', text: JSON.stringify(payload) };
}

/** 参数非法（JSON 解析失败）。 */
export function commandParamInvalid(message = '参数必须是单个 JSON 对象文本') {
  return {
    kind: 'error',
    text: JSON.stringify({ ok: false, error: { code: 'PARAM_INVALID', message } }),
  };
}

/**
 * 注册面板命令（返回 disposer；index.js apply 内收集）。
 * @param {object} ctx DSH 插件上下文（需 commands / workingDirectory / subprocess）
 * @param {{config?:object, tasks?:import('./tasks.js').TaskStore, officeTasks?:import('./tasks.js').TaskStore}} [options]
 * @returns {() => void}
 */
export function registerPanelCommands(ctx, options = {}) {
  const pluginConfig = options.config ?? {};
  const store = options.tasks ?? defaultTasks;
  /** 批量排版任务注册表（P2 独立实例；kind='office.batch'，与 CAD 任务互不干扰）。 */
  const officeTasks = options.officeTasks ?? new TaskStore();
  const disposers = [];

  const register = (spec) => {
    try {
      const dispose = ctx.commands.register(spec);
      if (typeof dispose === 'function') disposers.push(dispose);
    } catch (err) {
      console.error(`[office-suite] 命令注册失败（${spec.name}）:`, err);
    }
  };

  const configNow = () => effectiveConfig(pluginConfig);
  const boundaryOptions = (config, write) => ({
    write,
    allowExternalPaths: config.allowExternalPaths === true,
    root: resolveBoundaryRoot(config),
    warn: (message) => console.warn(`[office-suite] ${message}`),
  });

  /** 包装 handler：解析参数 + 统一异常 → ops 错误 JSON。 */
  const wrap = (fn) => async (invocation) => {
    try {
      const args = parseCommandInput(invocation?.rawInput);
      if (args === undefined) return commandParamInvalid();
      return await fn(args, invocation ?? {});
    } catch (err) {
      const payload = toErrorPayload(err);
      return commandSuccess(fail(payload.code, payload.message, payload.hint));
    }
  };

  // ── 1) suite.env：引擎探测（60s 缓存 + refresh 强刷） ──
  register({
    name: 'suite.env',
    description: '检测本机办公/绘图/CAD 引擎环境（Word/Excel/Visio/AutoCAD/中望/浩辰/ODA/Python），60 秒缓存，refresh:true 强刷',
    handler: wrap(async (args, invocation) => {
      const result = await probeEngines(ctx, configNow(), {
        signal: invocation.signal,
        refresh: args.refresh === true,
      });
      return commandSuccess(result);
    }),
  });

  // ── 2) suite.read：Word / Excel 读取（office-read.ps1） ──
  register({
    name: 'suite.read',
    description: '读取 Word/Excel（doc/docx/xls/xlsx）：大纲、表格数、正文预览、工作表清单与样例',
    handler: wrap(async (args, invocation) => {
      const config = configNow();
      const rawPath = typeof args.path === 'string' ? args.path.trim() : '';
      if (!rawPath) return commandSuccess(fail('E1041', '缺少 path'));
      const abs = assertInsideRoot(await resolveInputPath(ctx, invocation, rawPath), boundaryOptions(config, false));
      if (!existsSync(abs)) return commandSuccess(fail('E1023', `文件不存在: ${abs}`));
      const ext = extname(abs).toLowerCase();
      if (!['.doc', '.docx', '.xls', '.xlsx'].includes(ext)) {
        return commandSuccess(fail('E1042', `暂不支持的文件类型: ${ext || '(无扩展名)'}`, '仅支持 doc/docx/xls/xlsx'));
      }
      const mode = args.mode === 'full' ? 'full' : 'outline';
      const result = await runPs1(ctx, 'office-read.ps1', { path: abs, mode }, {
        signal: invocation.signal,
        timeoutMs: 180_000,
      });
      const parsed = parseJson(result.stdout);
      if (!parsed) {
        return commandSuccess(fail('E1051', result.stderr || result.stdout || '读取失败'));
      }
      return commandSuccess(parsed);
    }),
  });

  // ── 3) suite.draw：Visio 绘图（flowchart + P2 电网三图） ──
  register({
    name: 'suite.draw',
    description: "调用本机 Visio 绘图并保存为 .vsdx：kind='flowchart'（nodes/edges）；kind='singleLine'|'distribution'|'networkTopology' 需 model（电网图数据模型，见 docs/P2-设计.md §4.1）",
    handler: wrap(async (args, invocation) => {
      const config = configNow();
      if (!DRAW_KINDS.includes(args.kind)) {
        return commandSuccess(fail(
          'E1041',
          `kind 必须为 ${DRAW_KINDS.join(' | ')}，收到: ${JSON.stringify(args.kind)}`,
          'flowchart 使用 nodes/edges；电网三图使用 model',
        ));
      }
      const cwd = await sessionCwd(ctx, invocation);

      // P1 行为不变：流程图
      if (args.kind === 'flowchart') {
        const nodes = Array.isArray(args.nodes) ? args.nodes : null;
        const edges = Array.isArray(args.edges) ? args.edges : null;
        if (!nodes || !edges) return commandSuccess(fail('E1041', 'nodes 与 edges 必须是数组'));
        for (const [index, node] of nodes.entries()) {
          if (!node || typeof node.id !== 'string' || !node.id.trim() || typeof node.label !== 'string') {
            return commandSuccess(fail('E1041', `nodes[${index}] 需要 id/label 字符串`));
          }
          if (node.kind != null && !['start', 'process', 'decision', 'end'].includes(node.kind)) {
            return commandSuccess(fail('E1041', `nodes[${index}].kind 非法: ${JSON.stringify(node.kind)}`));
          }
        }
        for (const [index, edge] of edges.entries()) {
          if (!edge || typeof edge.from !== 'string' || typeof edge.to !== 'string') {
            return commandSuccess(fail('E1041', `edges[${index}] 需要 from/to 字符串`));
          }
        }
        const outputPath = args.outputPath
          ? assertInsideRoot(await resolveInputPath(ctx, invocation, String(args.outputPath)), boundaryOptions(config, true))
          : assertInsideRoot(join(cwd, 'output', 'diagrams', `flowchart-${Date.now()}.vsdx`), boundaryOptions(config, true));
        mkdirSync(dirname(outputPath), { recursive: true });
        const result = await runPs1(ctx, 'visio-flowchart.ps1', {
          kind: args.kind,
          title: typeof args.title === 'string' ? args.title : '',
          nodes,
          edges,
          outputPath,
        }, {
          signal: invocation.signal,
          timeoutMs: 300_000,
          cwd,
        });
        const parsed = parseJson(result.stdout);
        if (!parsed) return commandSuccess(fail('E1051', result.stderr || result.stdout || '绘图失败'));
        if (!parsed.ok) {
          const message = String(parsed.error?.message ?? '绘图失败');
          const code = /visio|com|progid/i.test(message) ? 'E1052' : (parsed.error?.code ?? 'E1051');
          return commandSuccess(fail(code, message, parsed.error?.hint));
        }
        return commandSuccess(ok({
          outputPath: parsed.result?.outputPath ?? outputPath,
          stats: { nodes: nodes.length, edges: edges.length },
          format: parsed.result?.format ?? 'vsdx',
        }));
      }

      // P2 电网三图：符号库 → 模型/接线校验 → 展开连线 → visio-grid.ps1
      if (args.model == null || typeof args.model !== 'object' || Array.isArray(args.model)) {
        return commandSuccess(fail('E2021', '电网图需要 model 对象', '模型结构见 docs/P2-设计.md §4.1'));
      }
      const loaded = loadSymbols(args.kind);
      if (!loaded.ok) return commandSuccess(loaded);
      const validated = validateGridModel(args.kind, args.model, { symbols: loaded.result.symbols });
      if (!validated.ok) return commandSuccess(validated);
      const expanded = expandGridLinks(validated.result.model);
      const outputPath = args.outputPath
        ? assertInsideRoot(await resolveInputPath(ctx, invocation, String(args.outputPath)), boundaryOptions(config, true))
        : assertInsideRoot(join(cwd, 'output', 'diagrams', `${args.kind}-${Date.now()}.vsdx`), boundaryOptions(config, true));
      mkdirSync(dirname(outputPath), { recursive: true });
      const result = await runPs1(ctx, 'visio-grid.ps1', {
        kind: args.kind,
        title: typeof args.title === 'string' && args.title ? args.title : (validated.result.model.title ?? ''),
        model: validated.result.model,
        symbols: loaded.result.symbols,
        outputPath,
        validations: validated.result.validations,
      }, {
        signal: invocation.signal,
        timeoutMs: 300_000,
        cwd,
      });
      const parsed = parseJson(result.stdout);
      if (!parsed) return commandSuccess(fail('E1051', result.stderr || result.stdout || '绘图失败'));
      if (!parsed.ok) {
        const message = String(parsed.error?.message ?? '绘图失败');
        const code = parsed.error?.code ?? (/visio|com|progid/i.test(message) ? 'E1052' : 'E1051');
        return commandSuccess(fail(code, message, parsed.error?.hint));
      }
      const drawn = parsed.result ?? {};
      return commandSuccess(ok({
        outputPath: drawn.outputPath ?? outputPath,
        format: drawn.format ?? 'vsdx',
        kind: drawn.kind ?? args.kind,
        paper: drawn.paper ?? null,
        stats: drawn.stats ?? { symbols: 0, shapes: 0, links: expanded.length, buses: 0, bays: 0, groups: 0 },
        validations: Array.isArray(drawn.validations) ? drawn.validations : validated.result.validations,
      }));
    }),
  });

  // ── 4) suite.files：目录列举（工作区 / 边界内任意目录） ──
  register({
    name: 'suite.files',
    description: '列举目录内容（name/type/size），支持 extensions 过滤；路径受 2026 边界校验',
    handler: wrap(async (args, invocation) => {
      const config = configNow();
      const result = await listDirectory(ctx, {
        path: typeof args.path === 'string' ? args.path : undefined,
        extensions: args.extensions,
        limit: args.limit,
      }, invocation, config);
      return commandSuccess(result);
    }),
  });

  // ── 5) suite.cad.convert：校验 + 注册任务 + 异步启动 ──
  register({
    name: 'suite.cad.convert',
    description: 'CAD→PDF 转换（merge 文件夹合并 / split 单文件多图框分割）；立即返回 taskId，进度用 suite.cad.status 轮询',
    handler: wrap(async (args, invocation) => {
      const result = await convertCad(ctx, args, invocation, configNow(), store);
      return commandSuccess(result);
    }),
  });

  // ── 6) suite.cad.status：任务状态（面板 800ms 轮询端点） ──
  register({
    name: 'suite.cad.status',
    description: '查询 CAD 转换任务状态/进度/日志；无 taskId 时返回任务列表',
    handler: wrap(async (args) => commandSuccess(cadStatus(store, args))),
  });

  // ── 7) suite.cad.cancel：取消任务（AbortController） ──
  register({
    name: 'suite.cad.cancel',
    description: '取消排队/运行中的 CAD 转换任务',
    handler: wrap(async (args) => commandSuccess(cadCancel(store, args))),
  });

  // ── 8) suite.config.get：有效配置（默认 ← 插件 Config ← UI 覆盖层） ──
  register({
    name: 'suite.config.get',
    description: '读取有效运行配置（含 runtime/ui-config.json 覆盖层）、可写键与边界信息',
    handler: wrap(async () => {
      const config = configNow();
      return commandSuccess(ok({
        effective: config,
        overrides: loadUiConfig(),
        writableKeys: [...WRITABLE_CONFIG_KEYS],
        path: UI_CONFIG_PATH,
        boundary: boundaryInfo(config),
      }));
    }),
  });

  // ── 9) suite.config.set：白名单键写 runtime\ui-config.json ──
  register({
    name: 'suite.config.set',
    description: `修改运行配置（白名单键：${WRITABLE_CONFIG_KEYS.join('/')}），写入 runtime/ui-config.json`,
    handler: wrap(async (args) => {
      const key = typeof args.key === 'string' ? args.key.trim() : '';
      const validation = validateConfigEntry(key, args.value);
      if (!validation.ok) return commandSuccess(fail(validation.code, validation.message));
      const config = configNow();
      if (key === 'outputRoot' && validation.value) {
        assertInsideRoot(validation.value, boundaryOptions(config, true));
      }
      if (key === 'odaPath' && validation.value && !existsSync(validation.value)) {
        return commandSuccess(fail('CONFIG_INVALID_VALUE', `ODAFileConverter 路径不存在: ${validation.value}`));
      }
      const overrides = { ...loadUiConfig(), [key]: validation.value };
      saveUiConfig(overrides);
      const next = configNow();
      return commandSuccess(ok({
        saved: { key, value: validation.value },
        effective: next,
        overrides: loadUiConfig(),
        writableKeys: [...WRITABLE_CONFIG_KEYS],
        path: UI_CONFIG_PATH,
      }));
    }),
  });

  // ── 10) suite.preview：officeToPdf 渲染预览（按 P2 §3.7 已核实契约修正） ──
  register({
    name: 'suite.preview',
    description: '通过 Host officeToPdf.convert 把 doc/docx/xls/xlsx 渲染为 PDF（base64）；服务缺失返回 PREVIEW_UNSUPPORTED',
    handler: wrap(async (args, invocation) => {
      const config = configNow();
      // 先判空 Host 服务（access.optional：ctx.get 不存在时返回 undefined）
      let service;
      try {
        service = typeof ctx.get === 'function' ? ctx.get('officeToPdf') : undefined;
      } catch {
        service = undefined;
      }
      if (!service || typeof service.convert !== 'function') {
        return commandSuccess(fail(
          'PREVIEW_UNSUPPORTED',
          'Host officeToPdf 服务不可用',
          '预览需要宿主提供 officeToPdf.convert；可改用系统应用打开文件，或在 Client 走 remote.officeToPdf 直连（§3.7 方案 A）',
        ));
      }

      const rawPath = typeof args.path === 'string' ? args.path.trim() : '';
      if (!rawPath) return commandSuccess(fail('E1041', '缺少 path'));
      const abs = assertInsideRoot(await resolveInputPath(ctx, invocation, rawPath), boundaryOptions(config, false));
      const ext = extname(abs).toLowerCase();
      if (!PREVIEW_EXTENSIONS.includes(ext)) {
        return commandSuccess(fail(
          'PREVIEW_UNSUPPORTED',
          `该类型不支持预览: ${ext || '(无扩展名)'}`,
          `仅支持 ${PREVIEW_EXTENSIONS.join('/')}（officeToPdf 的 OfficeExtension 不含 ppt/vsdx/dwg/pdf）`,
        ));
      }
      const info = await stat(abs).catch(() => null);
      if (!info || !info.isFile()) return commandSuccess(fail('E1023', `文件不存在: ${abs}`));

      const version = String(info.mtimeMs ?? 0);
      const priority = config.previewPriority === 'background' ? 'background' : 'foreground';
      try {
        const request = {
          extension: ext.slice(1), // 'doc' | 'docx' | 'xls' | 'xlsx'
          priority,
          source: {
            key: 'suite-preview',
            version,
            bytes: info.size,
            read: async (signal, maxBytes) => {
              const limit = Number(maxBytes);
              if (Number.isFinite(limit) && limit > 0 && info.size > limit) {
                throw new OpsError('E2011', `文件超过预览大小上限（${info.size} > ${limit} 字节）`);
              }
              const bytes = await readFile(abs, { signal });
              return { bytes, version };
            },
          },
        };
        const output = await service.convert(request, invocation.signal);
        if (!output?.pdf) {
          return commandSuccess(fail('E2011', 'officeToPdf.convert 返回结果缺少 pdf 字节'));
        }
        return commandSuccess(ok({
          pdfBase64: Buffer.from(output.pdf).toString('base64'),
          missingFonts: Array.isArray(output.missingFonts) ? output.missingFonts : [],
          generation: output.generation ?? null,
          cacheKey: output.cacheKey ?? null,
        }));
      } catch (err) {
        return commandSuccess(fail(
          'E2011',
          `预览渲染失败: ${String(err?.message ?? err)}`,
          '请确认文件未被占用且未超过渲染大小上限；也可改用系统应用打开',
        ));
      }
    }),
  });

  // ── 11) suite.edit：Office 结构化编辑（Word / Excel） ──
  register({
    name: 'suite.edit',
    description: '结构化编辑 Word/Excel：{path, ops, engine?}；修改前自动备份，结果含 applied/engine/backup',
    handler: wrap(async (args, invocation) => commandSuccess(await editOp(ctx, configNow(), args, invocation))),
  });

  // ── 12) suite.format：Word 单文件排版（rules / styleSet / template） ──
  register({
    name: 'suite.format',
    description: 'Word 单文件排版：{path, rules?/styleSet?/template?/placeholders?/scope?, engine?}；修改前自动备份',
    handler: wrap(async (args, invocation) => commandSuccess(await formatOp(ctx, configNow(), args, invocation))),
  });

  // ── 13) suite.format.batch：批量排版（异步任务，立即返回 taskId） ──
  register({
    name: 'suite.format.batch',
    description: '批量 Word 排版：{dir|paths, glob?, rules?/styleSet?/template?/placeholders?, parallel?}；立即返回 taskId，进度用 suite.office.status 轮询',
    handler: wrap(async (args, invocation) => commandSuccess(await batchFormatOp(ctx, configNow(), args, invocation, officeTasks))),
  });

  // ── 14) suite.office.status：批量排版任务状态（800ms 轮询端点） ──
  register({
    name: 'suite.office.status',
    description: "查询批量排版任务状态/进度/日志；无 taskId 时返回批量任务摘要列表（kind='office.batch'）",
    handler: wrap(async (args) => commandSuccess(batchFormatStatus(officeTasks, args))),
  });

  // ── 15) suite.office.cancel：取消批量排版任务 ──
  register({
    name: 'suite.office.cancel',
    description: '取消排队/运行中的批量排版任务（{taskId}）',
    handler: wrap(async (args) => commandSuccess(batchFormatCancel(officeTasks, args))),
  });

  // ── 16) suite.office.backup.list：列出文档自动备份 ──
  register({
    name: 'suite.office.backup.list',
    description: '列出文档的自动备份记录（{path}）；无备份返回 BACKUP_NONE',
    handler: wrap(async (args, invocation) => commandSuccess(await backupOp(ctx, configNow(), { ...args, action: 'list' }, invocation))),
  });

  // ── 17) suite.office.backup.restore：恢复备份（恢复前先备份当前版本） ──
  register({
    name: 'suite.office.backup.restore',
    description: '恢复文档备份（{path, id}）：恢复前自动备份当前版本，返回 restoredTo/preRestoreBackup',
    handler: wrap(async (args, invocation) => commandSuccess(await backupOp(ctx, configNow(), { ...args, action: 'restore' }, invocation))),
  });

  // ── 18) suite.templates：样式集与模板目录枚举 ──
  register({
    name: 'suite.templates',
    description: '列出可用样式集（suite/runtime templates/styles/*.json 合并，同名用户优先）与 docx 模板；目录缺失时返回提示',
    handler: wrap(async () => commandSuccess(await listTemplateCatalog(configNow()))),
  });

  return () => {
    for (const dispose of [...disposers].reverse()) {
      try { dispose(); } catch { /* 卸载时忽略个别清理失败 */ }
    }
  };
}
