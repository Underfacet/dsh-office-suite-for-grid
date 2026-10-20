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
 * P3 新增（docs/P3-设计.md §2.1/§3.1/§5.3）：suite.cad.catalog（图纸目录表）、
 *         suite.export（导出 PDF）；suite.cad.convert 增 inputs 批量队列；suite.draw 扩
 *         P3 五类图（protection/substationPlan/geoNetwork/powerFlow/terminalStrip）；
 *         suite.preview 改由 ops/office-pdf.js 的 previewPdfOp 实现；suite.cad.status
 *         补齐任务 kind 与 cad.queue 父任务摘要。
 *
 * 路由解耦（面板 typert 网关 descriptor 不一致的规避）：
 * 所有命令的实现在 `runSuiteCommand(ctx, config, name, args, signal)` 中，命令注册
 * （`ctx.commands.register`）与 panel-rpc.js 的自注册 HTTP 路由共用它；命令注册通道的
 * 行为与返回格式保持不变。另提供两条只读路由命令供面板文件浏览：
 *   workspace.list {path} / workspace.stat {path}（复用 ops/files.js 与 boundary，写=否）。
 */
import { existsSync, mkdirSync } from 'node:fs';
import { readdir, readFile, stat } from 'node:fs/promises';
import { basename, dirname, extname, isAbsolute, join, resolve } from 'node:path';
import { runPs1 } from './engines.js';
import { assertInsideRoot, boundaryInfo, resolveBoundaryRoot } from './boundary.js';
import {
  PACKAGE_DIR,
  PROJECT_DIR,
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
import { batchCadOp, QUEUE_KIND } from './ops/cad-batch.js';
import { buildCatalogOp } from './ops/cad-catalog.js';
import { editOp } from './ops/office-edit.js';
import { batchFormatCancel, batchFormatOp, batchFormatStatus, formatOp } from './ops/office-format.js';
import { previewPdfOp } from './ops/office-pdf.js';
import { backupOp } from './ops/backup.js';
import {
  GRID_KINDS_PENDING,
  P2_GRID_KINDS,
  P3_GRID_KINDS,
  expandGridLinks,
  loadSymbols,
  validateGridModel,
} from './ops/grid-model.js';
import { TaskStore, tasks as defaultTasks } from './tasks.js';

/**
 * suite.draw / visio_draw 支持的图类型：P1 流程图 + P2 电网三图 + P3 五类（共 9 类）。
 * cableRoute 属 P4（GRID_KINDS_PENDING），命令层返回 E2024 而不是「未知 kind」。
 */
export const DRAW_KINDS = Object.freeze(['flowchart', ...P2_GRID_KINDS, ...P3_GRID_KINDS]);

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
 * 任务存储按 ctx 隔离：命令注册与 HTTP 路由（panel-rpc.js）必须共享同一实例，
 * 否则 suite.cad.status/cancel、suite.office.status/cancel 看不到启动任务。
 * `options.tasks/officeTasks` 仍可用于测试注入（与旧 registerPanelCommands 选项一致）。
 */
const TASK_STORES = new WeakMap();

/** 取（必要时创建）某 ctx 的任务存储。 */
function storesFor(ctx, options = {}) {
  const key = ctx && (typeof ctx === 'object' || typeof ctx === 'function') ? ctx : null;
  let stores = key ? TASK_STORES.get(key) : undefined;
  if (!stores) {
    stores = {
      cad: options.tasks ?? defaultTasks,
      office: options.officeTasks ?? new TaskStore(),
    };
    if (key) TASK_STORES.set(key, stores);
  }
  return stores;
}

/** boundary 选项（与旧 registerPanelCommands 内闭包同口径）。 */
function boundaryOptionsFor(config, write) {
  return {
    write,
    allowExternalPaths: config.allowExternalPaths === true,
    root: resolveBoundaryRoot(config),
    warn: (message) => console.warn(`[office-suite] ${message}`),
  };
}

/**
 * 面板文件浏览的路径解析：以传入 path 为绝对路径处理；相对路径按项目根解析；
 * 缺省 = 项目根（仓库根，含 docs/tests/runtime/output）。
 */
function resolveWorkspacePath(rawPath) {
  const text = typeof rawPath === 'string' ? rawPath.trim() : '';
  if (!text) return PROJECT_DIR;
  return isAbsolute(text) ? resolve(text) : resolve(PROJECT_DIR, text);
}

/** workspace.list：只读目录列举（复用 ops/files.js 的 readdir 实现与排序）。 */
async function workspaceListOp(ctx, config, args, exec) {
  const abs = resolveWorkspacePath(args?.path);
  return listDirectory(ctx, { path: abs, extensions: args?.extensions, limit: args?.limit }, exec, config);
}

/** workspace.stat：只读文件元信息（bytes/version/absolutePath 供面板元信息区展示）。 */
async function workspaceStatOp(config, args) {
  const rawPath = typeof args?.path === 'string' ? args.path.trim() : '';
  if (!rawPath) return fail('E1041', '缺少 path');
  try {
    const abs = assertInsideRoot(resolveWorkspacePath(rawPath), boundaryOptionsFor(config, false));
    let info;
    try {
      info = await stat(abs);
    } catch (error) {
      if (error?.code === 'ENOENT') return fail('E1023', `文件不存在: ${abs}`);
      if (error?.code === 'EACCES' || error?.code === 'EPERM') return fail('E1022', `无权限读取: ${abs}`);
      return fail('E1099', `读取文件元信息失败: ${String(error?.message ?? error)}`);
    }
    return ok({
      absolutePath: abs,
      path: abs,
      bytes: info.size,
      size: info.size,
      version: String(info.mtimeMs ?? 0),
      mtimeMs: info.mtimeMs,
      isFile: info.isFile(),
      isDirectory: info.isDirectory(),
    });
  } catch (error) {
    const payload = toErrorPayload(error);
    return fail(payload.code, payload.message, payload.hint);
  }
}

/**
 * P3 §3.2 兼容层：ops/cad.js 的 cadStatus 只输出任务摘要（summary 无 kind 字段），
 * 而批量队列父任务 kind='cad.queue' 与子任务同在一个 TaskStore 中。此处按 taskId
 * 回查 kind 并给列表补 `queues` 摘要（itemsDone/itemsTotal、完成后的 failures），
 * 不修改 cad.js（保持既有 cadStatus 契约与 P1/P2 行为）。
 */
function augmentCadStatus(store, result) {
  if (!result?.ok || !result.result) return result;
  const value = result.result;
  const kindOf = (taskId) => (typeof taskId === 'string' ? store.get(taskId)?.kind ?? null : null);
  if (Array.isArray(value.tasks)) {
    const tasks = value.tasks.map((summary) => {
      const kind = kindOf(summary?.taskId);
      if (!kind) return summary;
      const entry = { ...summary, kind };
      if (kind === QUEUE_KIND) {
        const output = summary.output && typeof summary.output === 'object' ? summary.output : null;
        entry.queue = {
          itemsDone: summary.progress?.itemsDone ?? 0,
          itemsTotal: summary.progress?.itemsTotal ?? 0,
          succeeded: output?.succeeded ?? null,
          failed: output?.failed ?? null,
          failures: Array.isArray(output?.failures) ? output.failures : [],
        };
      }
      return entry;
    });
    return ok({
      ...value,
      tasks,
      queues: tasks.filter((item) => item.kind === QUEUE_KIND),
    });
  }
  if (typeof value.taskId === 'string') {
    return ok({ ...value, kind: kindOf(value.taskId) });
  }
  return result;
}

/**
 * 执行一条面板命令（命令注册通道与 panel-rpc HTTP 路由共用）。
 *
 * @param {object} ctx DSH 插件上下文（commands / workingDirectory / subprocess / officeToPdf 等）
 * @param {object} config 插件 Config（同 registerPanelCommands 的 options.config）
 * @param {string} name 命令名（suite.* 或 workspace.list / workspace.stat）
 * @param {object} args 已解析参数对象
 * @param {AbortSignal} [signal] 取消信号（命令注册通道来自 invocation.signal；HTTP 路由来自请求生命周期）
 * @param {object} [invocation] 可选：命令注册通道的 invocation（保留 agent 会话上下文；HTTP 路由不传）
 * @returns {Promise<{kind:'success'|'error',text:string}>} 命令层信封（与既有格式一致）
 */
export async function runSuiteCommand(ctx, config, name, args, signal, invocation) {
  const pluginConfig = config ?? {};
  // 命令注册通道保留完整 invocation（相对路径按会话工作区解析）；HTTP 路由只有 signal。
  const exec = invocation && typeof invocation === 'object'
    ? { ...invocation, signal: signal ?? invocation.signal }
    : { signal };
  const stores = storesFor(ctx);
  const configNow = () => effectiveConfig(pluginConfig);
  const safeArgs = args && typeof args === 'object' && !Array.isArray(args) ? args : {};

  try {
    switch (name) {
      // ── 1) suite.env：引擎探测（60s 缓存 + refresh 强刷） ──
      case 'suite.env': {
        const result = await probeEngines(ctx, configNow(), {
          signal: exec.signal,
          refresh: safeArgs.refresh === true,
        });
        return commandSuccess(result);
      }

      // ── 2) suite.read：Word / Excel 读取（office-read.ps1） ──
      case 'suite.read': {
        const config = configNow();
        const rawPath = typeof safeArgs.path === 'string' ? safeArgs.path.trim() : '';
        if (!rawPath) return commandSuccess(fail('E1041', '缺少 path'));
        const abs = assertInsideRoot(await resolveInputPath(ctx, exec, rawPath), boundaryOptionsFor(config, false));
        if (!existsSync(abs)) return commandSuccess(fail('E1023', `文件不存在: ${abs}`));
        const ext = extname(abs).toLowerCase();
        if (!['.doc', '.docx', '.xls', '.xlsx'].includes(ext)) {
          return commandSuccess(fail('E1042', `暂不支持的文件类型: ${ext || '(无扩展名)'}`, '仅支持 doc/docx/xls/xlsx'));
        }
        const mode = safeArgs.mode === 'full' ? 'full' : 'outline';
        const result = await runPs1(ctx, 'office-read.ps1', { path: abs, mode }, {
          signal: exec.signal,
          timeoutMs: 180_000,
        });
        const parsed = parseJson(result.stdout);
        if (!parsed) {
          return commandSuccess(fail('E1051', result.stderr || result.stdout || '读取失败'));
        }
        return commandSuccess(parsed);
      }

      // ── 3) suite.draw：Visio 绘图（flowchart + P2 电网三图 + P3 五类） ──
      case 'suite.draw': {
        const config = configNow();
        if (GRID_KINDS_PENDING.includes(safeArgs.kind)) {
          // P3 §4.4⑥：schema 定稿但本期未实现（E2024，不静默）
          return commandSuccess(loadSymbols(safeArgs.kind));
        }
        if (!DRAW_KINDS.includes(safeArgs.kind)) {
          return commandSuccess(fail(
            'E1041',
            `kind 必须为 ${DRAW_KINDS.join(' | ')}，收到: ${JSON.stringify(safeArgs.kind)}`,
            'flowchart 使用 nodes/edges；电网图使用 model（P2 §4.1 / P3 §4.4）',
          ));
        }
        const cwd = await sessionCwd(ctx, exec);

        // P1 行为不变：流程图
        if (safeArgs.kind === 'flowchart') {
          const nodes = Array.isArray(safeArgs.nodes) ? safeArgs.nodes : null;
          const edges = Array.isArray(safeArgs.edges) ? safeArgs.edges : null;
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
          const outputPath = safeArgs.outputPath
            ? assertInsideRoot(await resolveInputPath(ctx, exec, String(safeArgs.outputPath)), boundaryOptionsFor(config, true))
            : assertInsideRoot(join(cwd, 'output', 'diagrams', `flowchart-${Date.now()}.vsdx`), boundaryOptionsFor(config, true));
          mkdirSync(dirname(outputPath), { recursive: true });
          const result = await runPs1(ctx, 'visio-flowchart.ps1', {
            kind: safeArgs.kind,
            title: typeof safeArgs.title === 'string' ? safeArgs.title : '',
            nodes,
            edges,
            outputPath,
          }, {
            signal: exec.signal,
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

        // P2 电网三图（visio-grid.ps1）/ P3 五类（visio-layouts.ps1）：payload 同构
        if (safeArgs.model == null || typeof safeArgs.model !== 'object' || Array.isArray(safeArgs.model)) {
          return commandSuccess(fail('E2021', '电网图需要 model 对象', '模型结构见 docs/P2-设计.md §4.1 与 docs/P3-设计.md §4.4'));
        }
        const engineScript = P3_GRID_KINDS.includes(safeArgs.kind) ? 'visio-layouts.ps1' : 'visio-grid.ps1';
        const loaded = loadSymbols(safeArgs.kind);
        if (!loaded.ok) return commandSuccess(loaded);
        const validated = validateGridModel(safeArgs.kind, safeArgs.model, { symbols: loaded.result.symbols });
        if (!validated.ok) return commandSuccess(validated);
        const expanded = expandGridLinks(validated.result.model);
        const outputPath = safeArgs.outputPath
          ? assertInsideRoot(await resolveInputPath(ctx, exec, String(safeArgs.outputPath)), boundaryOptionsFor(config, true))
          : assertInsideRoot(join(cwd, 'output', 'diagrams', `${safeArgs.kind}-${Date.now()}.vsdx`), boundaryOptionsFor(config, true));
        mkdirSync(dirname(outputPath), { recursive: true });
        const result = await runPs1(ctx, engineScript, {
          kind: safeArgs.kind,
          title: typeof safeArgs.title === 'string' && safeArgs.title ? safeArgs.title : (validated.result.model.title ?? ''),
          model: validated.result.model,
          symbols: loaded.result.symbols,
          outputPath,
          validations: validated.result.validations,
        }, {
          signal: exec.signal,
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
          kind: drawn.kind ?? safeArgs.kind,
          engine: drawn.engine ?? engineScript,
          paper: drawn.paper ?? null,
          ...(drawn.pages != null ? { pages: drawn.pages } : {}),
          stats: drawn.stats ?? { symbols: 0, shapes: 0, links: expanded.length, buses: 0, bays: 0, groups: 0 },
          validations: Array.isArray(drawn.validations) ? drawn.validations : validated.result.validations,
          ...(Array.isArray(drawn.warnings) && drawn.warnings.length ? { warnings: drawn.warnings } : {}),
        }));
      }

      // ── 4) suite.files：目录列举（工作区 / 边界内任意目录） ──
      case 'suite.files': {
        const config = configNow();
        const result = await listDirectory(ctx, {
          path: typeof safeArgs.path === 'string' ? safeArgs.path : undefined,
          extensions: safeArgs.extensions,
          limit: safeArgs.limit,
        }, exec, config);
        return commandSuccess(result);
      }

      // ── 5) suite.cad.convert：校验 + 注册任务 + 异步启动（P3：inputs 走批量队列） ──
      case 'suite.cad.convert': {
        const config = configNow();
        if (Array.isArray(safeArgs.inputs) && safeArgs.inputs.length) {
          // P3 §3.1：inputs 每项一个子任务，父任务聚合（与 input 互斥，由 ops 校验）
          return commandSuccess(await batchCadOp(ctx, config, safeArgs, exec, stores.cad));
        }
        const result = await convertCad(ctx, safeArgs, exec, config, stores.cad);
        return commandSuccess(result);
      }

      // ── 6) suite.cad.status：任务状态（面板 800ms 轮询端点） ──
      case 'suite.cad.status':
        return commandSuccess(augmentCadStatus(stores.cad, cadStatus(stores.cad, safeArgs)));

      // ── 7) suite.cad.cancel：取消任务（AbortController） ──
      case 'suite.cad.cancel':
        return commandSuccess(cadCancel(stores.cad, safeArgs));

      // ── 8) suite.config.get：有效配置（默认 ← 插件 Config ← UI 覆盖层） ──
      case 'suite.config.get': {
        const config = configNow();
        return commandSuccess(ok({
          effective: config,
          overrides: loadUiConfig(),
          writableKeys: [...WRITABLE_CONFIG_KEYS],
          path: UI_CONFIG_PATH,
          boundary: boundaryInfo(config),
        }));
      }

      // ── 9) suite.config.set：白名单键写 runtime\ui-config.json ──
      case 'suite.config.set': {
        const key = typeof safeArgs.key === 'string' ? safeArgs.key.trim() : '';
        const validation = validateConfigEntry(key, safeArgs.value);
        if (!validation.ok) return commandSuccess(fail(validation.code, validation.message));
        const config = configNow();
        if (key === 'outputRoot' && validation.value) {
          assertInsideRoot(validation.value, boundaryOptionsFor(config, true));
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
      }

      // ── 10) suite.preview：officeToPdf 渲染预览（P3 §5.1：previewPdfOp） ──
      case 'suite.preview':
        return commandSuccess(await previewPdfOp(ctx, configNow(), { path: safeArgs.path }, exec));

      // ── 11) suite.edit：Office 结构化编辑（Word / Excel） ──
      case 'suite.edit':
        return commandSuccess(await editOp(ctx, configNow(), safeArgs, exec));

      // ── 12) suite.format：Word 单文件排版（rules / styleSet / template） ──
      case 'suite.format':
        return commandSuccess(await formatOp(ctx, configNow(), safeArgs, exec));

      // ── 13) suite.format.batch：批量排版（异步任务，立即返回 taskId） ──
      case 'suite.format.batch':
        return commandSuccess(await batchFormatOp(ctx, configNow(), safeArgs, exec, stores.office));

      // ── 14) suite.office.status：批量排版任务状态（800ms 轮询端点） ──
      case 'suite.office.status':
        return commandSuccess(batchFormatStatus(stores.office, safeArgs));

      // ── 15) suite.office.cancel：取消批量排版任务 ──
      case 'suite.office.cancel':
        return commandSuccess(batchFormatCancel(stores.office, safeArgs));

      // ── 16) suite.office.backup.list：列出文档自动备份 ──
      case 'suite.office.backup.list':
        return commandSuccess(await backupOp(ctx, configNow(), { ...safeArgs, action: 'list' }, exec));

      // ── 17) suite.office.backup.restore：恢复备份（恢复前先备份当前版本） ──
      case 'suite.office.backup.restore':
        return commandSuccess(await backupOp(ctx, configNow(), { ...safeArgs, action: 'restore' }, exec));

      // ── 18) suite.templates：样式集与模板目录枚举 ──
      case 'suite.templates':
        return commandSuccess(await listTemplateCatalog(configNow()));

      // ── 19) suite.cad.catalog：图纸目录表（P3 §2.1；同步执行，不走任务注册表） ──
      case 'suite.cad.catalog':
        return commandSuccess(await buildCatalogOp(ctx, configNow(), safeArgs, exec));

      // ── 20) suite.export：导出 PDF 文件落地（P3 §5.3；与 office_export 工具同一 ops） ──
      case 'suite.export': {
        const output = typeof safeArgs.output === 'string' && safeArgs.output.trim() ? safeArgs.output.trim() : true;
        const result = await previewPdfOp(
          ctx,
          configNow(),
          { path: safeArgs.path, saveAs: output, ...(safeArgs.engine ? { engine: safeArgs.engine } : {}) },
          exec,
        );
        return commandSuccess(result);
      }

      // ── 21) workspace.list：面板工作区目录列举（只读；相对路径按项目根解析） ──
      case 'workspace.list':
        return commandSuccess(await workspaceListOp(ctx, configNow(), safeArgs, exec));

      // ── 22) workspace.stat：面板工作区文件元信息（只读；boundary 写=否） ──
      case 'workspace.stat':
        return commandSuccess(await workspaceStatOp(configNow(), safeArgs));

      default:
        return commandSuccess(fail('COMMAND_NOT_FOUND', `未知命令: ${String(name)}`));
    }
  } catch (error) {
    const payload = toErrorPayload(error);
    return commandSuccess(fail(payload.code, payload.message, payload.hint));
  }
}

/**
 * 注册面板命令（返回 disposer；index.js apply 内收集）。
 * handler 只负责解析 rawInput，实现统一走 runSuiteCommand（与 panel-rpc.js 共用）。
 * @param {object} ctx DSH 插件上下文（需 commands / workingDirectory / subprocess）
 * @param {{config?:object, tasks?:import('./tasks.js').TaskStore, officeTasks?:import('./tasks.js').TaskStore}} [options]
 * @returns {() => void}
 */
export function registerPanelCommands(ctx, options = {}) {
  const pluginConfig = options.config ?? {};
  storesFor(ctx, options);
  const disposers = [];

  const register = (spec) => {
    try {
      const dispose = ctx.commands.register(spec);
      if (typeof dispose === 'function') disposers.push(dispose);
    } catch (err) {
      console.error(`[office-suite] 命令注册失败（${spec.name}）:`, err);
    }
  };

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

  /** 命令注册通道 → runSuiteCommand（invocation 透传，保留会话工作目录上下文）。 */
  const invoke = (name) => wrap(async (args, invocation) => runSuiteCommand(
    ctx,
    pluginConfig,
    name,
    args,
    invocation?.signal,
    invocation,
  ));

  // ── 1) suite.env：引擎探测（60s 缓存 + refresh 强刷） ──
  register({
    name: 'suite.env',
    description: '检测本机办公/绘图/CAD 引擎环境（Word/Excel/Visio/AutoCAD/中望/浩辰/ODA/Python），60 秒缓存，refresh:true 强刷',
    handler: invoke('suite.env'),
  });

  // ── 2) suite.read：Word / Excel 读取（office-read.ps1） ──
  register({
    name: 'suite.read',
    description: '读取 Word/Excel（doc/docx/xls/xlsx）：大纲、表格数、正文预览、工作表清单与样例',
    handler: invoke('suite.read'),
  });

  // ── 3) suite.draw：Visio 绘图（flowchart + P2 电网三图 + P3 五类） ──
  register({
    name: 'suite.draw',
    description: "调用本机 Visio 绘图并保存为 .vsdx：kind='flowchart'（nodes/edges）；kind='singleLine'|'distribution'|'networkTopology'（docs/P2-设计.md §4.1）与 'protection'|'substationPlan'|'geoNetwork'|'powerFlow'|'terminalStrip'（docs/P3-设计.md §4.4）需 model；cableRoute 为 P4（E2024）",
    handler: invoke('suite.draw'),
  });

  // ── 4) suite.files：目录列举（工作区 / 边界内任意目录） ──
  register({
    name: 'suite.files',
    description: '列举目录内容（name/type/size），支持 extensions 过滤；路径受 2026 边界校验',
    handler: invoke('suite.files'),
  });

  // ── 5) suite.cad.convert：校验 + 注册任务 + 异步启动（P3：inputs 走批量队列） ──
  register({
    name: 'suite.cad.convert',
    description: 'CAD→PDF 转换（merge 文件夹合并 / split 单文件多图框分割）：单任务用 input；批量队列用 inputs（1–100 项，父任务 kind=cad.queue）。立即返回 taskId，进度用 suite.cad.status 轮询',
    handler: invoke('suite.cad.convert'),
  });

  // ── 6) suite.cad.status：任务状态（面板 800ms 轮询端点） ──
  register({
    name: 'suite.cad.status',
    description: "查询 CAD 任务状态/进度/日志（含 batch 队列父任务 kind='cad.queue'）；无 taskId 时返回任务列表（含 kind 与队列摘要）",
    handler: invoke('suite.cad.status'),
  });

  // ── 7) suite.cad.cancel：取消任务（AbortController） ──
  register({
    name: 'suite.cad.cancel',
    description: '取消排队/运行中的 CAD 转换任务',
    handler: invoke('suite.cad.cancel'),
  });

  // ── 8) suite.config.get：有效配置（默认 ← 插件 Config ← UI 覆盖层） ──
  register({
    name: 'suite.config.get',
    description: '读取有效运行配置（含 runtime/ui-config.json 覆盖层）、可写键与边界信息',
    handler: invoke('suite.config.get'),
  });

  // ── 9) suite.config.set：白名单键写 runtime\ui-config.json ──
  register({
    name: 'suite.config.set',
    description: `修改运行配置（白名单键：${WRITABLE_CONFIG_KEYS.join('/')}），写入 runtime/ui-config.json`,
    handler: invoke('suite.config.set'),
  });

  // ── 10) suite.preview：officeToPdf 渲染预览（P3 §5.1：改由 ops/office-pdf.js 的 previewPdfOp 实现） ──
  register({
    name: 'suite.preview',
    description: '通过 Host officeToPdf.convert 把 doc/docx/xls/xlsx/ppt/pptx 渲染为 PDF（base64，供面板 iframe）；服务缺失/类型不支持返回 E2012',
    handler: invoke('suite.preview'),
  });

  // ── 11) suite.edit：Office 结构化编辑（Word / Excel） ──
  register({
    name: 'suite.edit',
    description: '结构化编辑 Word/Excel：{path, ops, engine?}；修改前自动备份，结果含 applied/engine/backup',
    handler: invoke('suite.edit'),
  });

  // ── 12) suite.format：Word 单文件排版（rules / styleSet / template） ──
  register({
    name: 'suite.format',
    description: 'Word 单文件排版：{path, rules?/styleSet?/template?/placeholders?/scope?, engine?}；修改前自动备份',
    handler: invoke('suite.format'),
  });

  // ── 13) suite.format.batch：批量排版（异步任务，立即返回 taskId） ──
  register({
    name: 'suite.format.batch',
    description: '批量 Word 排版：{dir|paths, glob?, rules?/styleSet?/template?/placeholders?, parallel?}；立即返回 taskId，进度用 suite.office.status 轮询',
    handler: invoke('suite.format.batch'),
  });

  // ── 14) suite.office.status：批量排版任务状态（800ms 轮询端点） ──
  register({
    name: 'suite.office.status',
    description: "查询批量排版任务状态/进度/日志；无 taskId 时返回批量任务摘要列表（kind='office.batch'）",
    handler: invoke('suite.office.status'),
  });

  // ── 15) suite.office.cancel：取消批量排版任务 ──
  register({
    name: 'suite.office.cancel',
    description: '取消排队/运行中的批量排版任务（{taskId}）',
    handler: invoke('suite.office.cancel'),
  });

  // ── 16) suite.office.backup.list：列出文档自动备份 ──
  register({
    name: 'suite.office.backup.list',
    description: '列出文档的自动备份记录（{path}）；无备份返回 BACKUP_NONE',
    handler: invoke('suite.office.backup.list'),
  });

  // ── 17) suite.office.backup.restore：恢复备份（恢复前先备份当前版本） ──
  register({
    name: 'suite.office.backup.restore',
    description: '恢复文档备份（{path, id}）：恢复前自动备份当前版本，返回 restoredTo/preRestoreBackup',
    handler: invoke('suite.office.backup.restore'),
  });

  // ── 18) suite.templates：样式集与模板目录枚举 ──
  register({
    name: 'suite.templates',
    description: '列出可用样式集（suite/runtime templates/styles/*.json 合并，同名用户优先）与 docx 模板；目录缺失时返回提示',
    handler: invoke('suite.templates'),
  });

  // ── 19) suite.cad.catalog：图纸目录表（P3 §2.1；同步执行，不走任务注册表） ──
  register({
    name: 'suite.cad.catalog',
    description: '生成图纸目录表 xlsx（图号/图名/图幅/比例/张数）：{input(.dwg/.pdf/目录), output?, sort?, engine?}；同步返回 outputPath/rows/fields，超限报 E1041、空表报 E1061',
    handler: invoke('suite.cad.catalog'),
  });

  // ── 20) suite.export：导出 PDF 文件落地（P3 §5.3；与 office_export 工具同一 ops） ──
  register({
    name: 'suite.export',
    description: '把 doc/docx/xls/xlsx/ppt/pptx 导出为 PDF 文件：{path, output?(目录或 .pdf 路径), engine?}；缺省 output\\pdf 下同名（重名加 -<时间戳>），服务缺失报 E2012、写盘失败报 E2013',
    handler: invoke('suite.export'),
  });

  return () => {
    for (const dispose of [...disposers].reverse()) {
      try { dispose(); } catch { /* 卸载时忽略个别清理失败 */ }
    }
  };
}
