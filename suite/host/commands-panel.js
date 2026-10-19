/**
 * 面板远程命令集合（docs/P1-设计.md §2，共 10 个命令）
 *
 * 约定：命令名小写、无前导斜杠；参数为单个 JSON 文本（rawInput）；
 * ops 结果统一序列化为 text：{ok:true,result} | {ok:false,error:{code,message,hint?}}；
 * 仅 rawInput 不是单个 JSON 对象时返回 {kind:'error', text}（PARAM_INVALID）。
 *
 * 与对话工具共用同一操作层（ops/*）；所有路径先过 boundary（§5.3）。
 */
import { existsSync, mkdirSync } from 'node:fs';
import { dirname, extname, join } from 'node:path';
import { runPs1 } from './engines.js';
import { assertInsideRoot, boundaryInfo, resolveBoundaryRoot } from './boundary.js';
import {
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
import { tasks as defaultTasks } from './tasks.js';

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
 * @param {{config?:object, tasks?:import('./tasks.js').TaskStore}} [options]
 * @returns {() => void}
 */
export function registerPanelCommands(ctx, options = {}) {
  const pluginConfig = options.config ?? {};
  const store = options.tasks ?? defaultTasks;
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

  // ── 3) suite.draw：Visio 绘图（P1 仅 flowchart） ──
  register({
    name: 'suite.draw',
    description: '调用本机 Visio 绘制流程图并保存为 .vsdx（P1 仅 kind=flowchart）',
    handler: wrap(async (args, invocation) => {
      const config = configNow();
      if (args.kind !== 'flowchart') {
        return commandSuccess(fail(
          'NOT_IMPLEMENTED_P2',
          `P1 仅支持 kind='flowchart'，收到: ${JSON.stringify(args.kind)}`,
          '主接线图/配网图等图类型将在 P2 提供',
        ));
      }
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
      const cwd = await sessionCwd(ctx, invocation);
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

  // ── 10) suite.preview：officeToPdf 渲染预览（P1 预留，失败不抛崩） ──
  register({
    name: 'suite.preview',
    description: '（预留）尝试通过 Host officeToPdf 服务把 doc/docx/xls/xlsx 渲染为 PDF（base64）',
    handler: wrap(async (args, invocation) => {
      const config = configNow();
      const rawPath = typeof args.path === 'string' ? args.path.trim() : '';
      if (!rawPath) return commandSuccess(fail('E1041', '缺少 path'));
      const abs = assertInsideRoot(await resolveInputPath(ctx, invocation, rawPath), boundaryOptions(config, false));
      const ext = extname(abs).toLowerCase();
      if (!['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx'].includes(ext)) {
        return commandSuccess(fail('PREVIEW_UNSUPPORTED', `该类型不支持预览: ${ext || '(无扩展名)'}`));
      }
      let service;
      try {
        service = typeof ctx.get === 'function' ? ctx.get('officeToPdf') : undefined;
      } catch {
        service = undefined;
      }
      if (!service || typeof service.convert !== 'function') {
        return commandSuccess(fail(
          'PREVIEW_UNSUPPORTED',
          'Host officeToPdf 服务不可用（预览将在 P2 接通）',
          '§0.6 #1：远程命名空间挂载待实测；当前仅保留接入点',
        ));
      }
      try {
        const request = {
          path: abs,
          absolutePath: abs,
          extension: ext.slice(1),
          format: 'pdf',
        };
        const output = await service.convert(request, invocation.signal);
        if (!output || !output.pdf) {
          return commandSuccess(fail('PREVIEW_UNSUPPORTED', 'officeToPdf.convert 返回结果缺少 pdf 字节'));
        }
        return commandSuccess(ok({
          pdfBase64: Buffer.from(output.pdf).toString('base64'),
          missingFonts: Array.isArray(output.missingFonts) ? output.missingFonts : [],
          generation: output.generation ?? null,
        }));
      } catch (err) {
        return commandSuccess(fail(
          'PREVIEW_UNSUPPORTED',
          `预览失败: ${String(err?.message ?? err)}`,
          'officeToPdf 契约形状待实测（§0.6 #1）',
        ));
      }
    }),
  });

  return () => {
    for (const dispose of [...disposers].reverse()) {
      try { dispose(); } catch { /* 卸载时忽略个别清理失败 */ }
    }
  };
}
