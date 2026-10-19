/**
 * CAD → PDF 引擎编排（docs/P1-设计.md §3.1 / §3.2 / §3.3 / §3.6）
 *
 * 引擎选择（探测缓存 60s，refresh 强刷）：
 *   ① CAD COM（AutoCAD / 中望 / 浩辰）—— 经 runPs1 调 engines\cad-convert.ps1；
 *   ② ODAFileConverter —— 整文件转换后交 PDF 级引擎；
 *   ③ PDF 级引擎（纯 Node，pdfjs-dist + pdf-lib）—— 输入为已有 PDF（或 ODA 产出）。
 * 无 CAD/ODA 且输入 .dwg → E1001 CAD_NO_ENGINE；降级在任务日志写 engineFallback。
 *
 * 场景 A（split）：单文件多图框 → 每图一页 PDF；
 * 场景 B（merge）：文件夹多图 → 按图号排序合并为一个 PDF。
 */
import { copyFile, mkdir, readdir, rm, stat } from 'node:fs/promises';
import { basename, dirname, extname, isAbsolute, join, resolve } from 'node:path';
import { runPs1, runExe } from '../engines.js';
import { assertInsideRoot, resolveBoundaryRoot } from '../boundary.js';
import {
  OpsError,
  TMP_DIR,
  fail,
  isAbortError,
  ok,
  parseJson,
  resolveInputPath,
  resolveParallel,
  sessionCwd,
  throwIfAborted,
  toErrorPayload,
} from '../ops.js';
import { sortFiles } from './drawingNo.js';
import { splitPdfFile } from './pdf-split.js';
import { mergePdfs } from './pdf-merge.js';

const PROBE_TTL_MS = 60_000;
const PROBE_TIMEOUT_MS = 60_000;
const DEFAULT_CAD_TIMEOUT_MS = 300_000;
const MERGE_EXTENSIONS = new Set(['.pdf', '.dwg']);
const EXCLUDED_FILE = /^~\$|\.(bak|dwl|dwl2|tmp)$/i;

// ───────────────────────── 引擎探测（缓存 60s） ─────────────────────────

let probeCache = { at: 0, data: null, pending: null };

/** 清空探测缓存（suite.env refresh / 配置变更时调用）。 */
export function invalidateProbeCache() {
  probeCache = { at: 0, data: null, pending: null };
}

/**
 * 运行 probe.ps1 并缓存结果（60s）。refresh=true 强刷。
 * @returns {Promise<{ok:true,result:object}|{ok:false,error:object}>}
 */
export async function probeEngines(ctx, config = {}, options = {}) {
  const refresh = options.refresh === true;
  if (refresh) invalidateProbeCache();
  const now = Date.now();
  if (!refresh && probeCache.data && now - probeCache.at < PROBE_TTL_MS) {
    return ok(probeCache.data);
  }
  if (!refresh && probeCache.pending) {
    try {
      return ok(await probeCache.pending);
    } catch (err) {
      return { ok: false, error: toErrorPayload(err, 'E1051') };
    }
  }
  const pending = (async () => {
    const result = await runPs1(ctx, 'probe.ps1', {}, {
      signal: options.signal,
      timeoutMs: Number.isFinite(options.timeoutMs) ? options.timeoutMs : PROBE_TIMEOUT_MS,
    });
    const parsed = parseJson(result.stdout);
    if (!parsed) {
      throw new OpsError('E1051', result.stderr || result.stdout || '环境探测失败', 'probe.ps1 未返回可解析的 JSON');
    }
    return parsed;
  })();
  probeCache.pending = pending;
  try {
    const data = await pending;
    probeCache = { at: Date.now(), data, pending: null };
    return ok(data);
  } catch (err) {
    probeCache.pending = null;
    return { ok: false, error: toErrorPayload(err, 'E1051') };
  }
}

/** 从探测结果提取引擎可用性。 */
export function analyzeEngines(probe) {
  const platforms = [];
  const cadPlatforms = Array.isArray(probe?.cad?.platforms) ? probe.cad.platforms : [];
  for (const platform of cadPlatforms) {
    if (!platform?.present) continue;
    platforms.push({
      id: String(platform.id ?? ''),
      name: String(platform.name ?? platform.progId ?? ''),
      progId: String(platform.progId ?? ''),
      curVer: platform.curVer ?? null,
      installPaths: Array.isArray(platform.installPaths) ? platform.installPaths : [],
    });
  }
  // 兼容 v0.1 probe 输出（仅顶层 ProgID 字段）
  if (!platforms.length) {
    for (const progId of ['AutoCAD.Application', 'ZWCAD.Application', 'GstarCAD.Application']) {
      const info = probe?.[progId];
      if (info?.present) {
        const id = progId.split('.')[0].toLowerCase();
        platforms.push({ id, name: progId.split('.')[0], progId, curVer: info.curVer ?? null, installPaths: [] });
      }
    }
  }

  const odaPaths = [];
  const pushPath = (value) => {
    if (typeof value === 'string' && value.trim() && !odaPaths.includes(value)) odaPaths.push(value);
  };
  pushPath(probe?.oda?.converter);
  for (const path of Array.isArray(probe?.oda?.paths) ? probe.oda.paths : []) pushPath(path);
  pushPath(probe?.odaConverter?.converter);
  for (const path of Array.isArray(probe?.odaConverter?.paths) ? probe.odaConverter.paths : []) pushPath(path);

  return {
    cadCom: { available: platforms.length > 0, platforms },
    oda: {
      found: probe?.oda?.found === true || probe?.odaConverter?.found === true || odaPaths.length > 0,
      converter: odaPaths[0] ?? null,
      paths: odaPaths,
    },
    pdf: { available: true },
  };
}

/** 选择主引擎：pdf 输入 → pdf；否则 com → oda → none。 */
export function chooseEngine({ isPdf, engines }) {
  if (isPdf) return 'pdf';
  if (engines?.cadCom?.available) return 'com';
  if (engines?.oda?.found) return 'oda';
  return 'none';
}

/** 降级链（日志用）：com → oda → pdf（按可用性过滤，pdf 恒可用）。 */
export function engineChain(engines, isPdf = false) {
  if (isPdf) return ['pdf'];
  const chain = [];
  if (engines?.cadCom?.available) chain.push('com');
  if (engines?.oda?.found) chain.push('oda');
  chain.push('pdf');
  return chain;
}

// ───────────────────────── 参数与路径准备 ─────────────────────────

function timestamp() {
  return new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14);
}

function randomSuffix() {
  return Math.random().toString(36).slice(2, 8);
}

function validateParallel(value) {
  if (value === undefined || value === null || value === '') return undefined;
  if (value === 'auto') return 'auto';
  if (Number.isInteger(value) && value >= 1 && value <= 8) return value;
  throw new OpsError('E1041', "parallel 必须为 'auto' 或 1–8 的整数", `收到: ${JSON.stringify(value)}`);
}

function validatePaper(value) {
  if (value === undefined || value === null || value === '') return undefined;
  if (value === 'A4-landscape' || value === 'keep') return value;
  throw new OpsError('E1041', "paper 必须为 'A4-landscape' | 'keep'", `收到: ${JSON.stringify(value)}`);
}

function validateSortMode(value) {
  if (value === undefined || value === null || value === '') return undefined;
  if (value === 'drawingNo' || value === 'filename') return value;
  throw new OpsError('E1041', "sort 必须为 'drawingNo' | 'filename'", `收到: ${JSON.stringify(value)}`);
}

async function exists(path) {
  try {
    await stat(path);
    return true;
  } catch {
    return false;
  }
}

/** 扫描合并输入目录（不递归；排除 ~$ / .bak / .dwl / .dwl2 / 输出文件）。 */
export async function scanMergeInput(dir, outputPath = null) {
  let dirents;
  try {
    dirents = await readdir(dir, { withFileTypes: true });
  } catch (err) {
    if (err?.code === 'ENOENT') throw new OpsError('E1023', `输入目录不存在: ${dir}`);
    if (err?.code === 'EACCES' || err?.code === 'EPERM') throw new OpsError('E1022', `无权限读取目录: ${dir}`);
    throw new OpsError('E1099', `读取输入目录失败: ${String(err?.message ?? err)}`);
  }
  const outputAbs = outputPath ? resolve(outputPath) : null;
  const files = [];
  for (const dirent of dirents) {
    if (!dirent.isFile()) continue;
    if (EXCLUDED_FILE.test(dirent.name)) continue;
    const ext = extname(dirent.name).toLowerCase();
    if (!MERGE_EXTENSIONS.has(ext)) continue;
    const full = resolve(dir, dirent.name);
    if (outputAbs && full === outputAbs) continue;
    files.push(full);
  }
  return sortFiles(files, 'filename');
}

/** 默认输出根：config.outputRoot 或 会话工作区 output\pdf。 */
async function defaultOutputRoot(ctx, exec, config) {
  if (typeof config?.outputRoot === 'string' && config.outputRoot.trim()) {
    const configured = config.outputRoot.trim();
    return isAbsolute(configured) ? resolve(configured) : resolve(await sessionCwd(ctx, exec), configured);
  }
  return join(await sessionCwd(ctx, exec), 'output', 'pdf');
}

/**
 * 校验参数 + 解析路径 + 选择引擎（不创建任务）。内部使用，抛 OpsError。
 */
export async function prepareCadSpec(ctx, args, exec, config = {}) {
  const mode = args?.mode;
  if (mode !== 'merge' && mode !== 'split') {
    throw new OpsError('E1041', "mode 必须为 'merge' | 'split'", `收到: ${JSON.stringify(mode)}`);
  }
  const inputArg = typeof args?.input === 'string' ? args.input.trim() : '';
  if (!inputArg) throw new OpsError('E1041', '缺少 input（文件或目录路径）');
  const paper = validatePaper(args?.paper) ?? (config.pdfPaper === 'keep' ? 'keep' : 'A4-landscape');
  const sortMode = validateSortMode(args?.sort) ?? (config.sortMode === 'filename' ? 'filename' : 'drawingNo');
  const parallel = validateParallel(args?.parallel) ?? config.parallel ?? 'auto';

  const root = resolveBoundaryRoot(config);
  const allowExternalPaths = config.allowExternalPaths === true;
  const inputAbs = assertInsideRoot(await resolveInputPath(ctx, exec, inputArg), {
    write: false,
    allowExternalPaths,
    root,
    warn: (message) => console.warn(`[office-suite] ${message}`),
  });
  if (!(await exists(inputAbs))) {
    throw new OpsError('E1023', `输入路径不存在: ${inputAbs}`);
  }
  const inputStat = await stat(inputAbs);

  let outputBase = await defaultOutputRoot(ctx, exec, config);
  outputBase = resolve(outputBase);
  let outputAbs;
  const outputArg = typeof args?.output === 'string' ? args.output.trim() : '';
  if (outputArg) {
    const candidate = assertInsideRoot(await resolveInputPath(ctx, exec, outputArg), { write: true, root });
    const candidateIsDir = await (async () => {
      try { return (await stat(candidate)).isDirectory(); } catch { return false; }
    })();
    outputAbs = candidateIsDir
      ? join(candidate, `${mode === 'merge' ? 'merged' : 'split'}-${timestamp()}.pdf`)
      : candidate;
  } else {
    outputAbs = join(outputBase, `${mode === 'merge' ? 'merged' : 'split'}-${timestamp()}.pdf`);
  }
  outputAbs = assertInsideRoot(outputAbs, { write: true, root });
  if (extname(outputAbs).toLowerCase() !== '.pdf') {
    throw new OpsError('E1041', `输出必须是 .pdf 文件: ${outputAbs}`);
  }

  let files = [];
  let isPdf = false;
  if (mode === 'split') {
    if (!inputStat.isFile()) {
      throw new OpsError('E1041', '分割模式的 input 必须是单个文件（.dwg 或 .pdf）', `收到目录: ${inputAbs}`);
    }
    const ext = extname(inputAbs).toLowerCase();
    if (ext !== '.dwg' && ext !== '.pdf') {
      throw new OpsError('E1041', '分割模式仅支持 .dwg / .pdf 输入', `收到: ${ext || '(无扩展名)'}`);
    }
    files = [inputAbs];
    isPdf = ext === '.pdf';
  } else {
    if (!inputStat.isDirectory()) {
      throw new OpsError('E1041', '合并模式的 input 必须是文件夹', `收到文件: ${inputAbs}`);
    }
    const allFiles = await scanMergeInput(inputAbs, outputAbs);
    if (!allFiles.length) {
      throw new OpsError('E1023', `输入目录没有可转换的 .dwg / .pdf 文件: ${inputAbs}`);
    }
    const dwgFiles = allFiles.filter((file) => extname(file).toLowerCase() === '.dwg');
    const pdfFiles = allFiles.filter((file) => extname(file).toLowerCase() === '.pdf');
    if (dwgFiles.length) {
      files = dwgFiles;
      isPdf = false;
      if (pdfFiles.length) {
        console.warn(`[office-suite] 合并目录同时含 .dwg 与 .pdf，按 CAD 引擎只处理 ${dwgFiles.length} 个 .dwg（忽略 ${pdfFiles.length} 个 .pdf）`);
      }
    } else {
      files = pdfFiles;
      isPdf = true;
    }
  }

  // PDF 输入直接走 PDF 级引擎，不做 CAD/ODA 探测（避免探测失败阻塞纯 PDF 场景）
  let engines;
  if (isPdf) {
    engines = {
      cadCom: { available: false, platforms: [] },
      oda: { found: false, converter: null, paths: [] },
      pdf: { available: true },
    };
  } else {
    const probe = await probeEngines(ctx, config);
    if (!probe.ok) {
      throw new OpsError(probe.error.code ?? 'E1051', probe.error.message ?? '环境探测失败', probe.error.hint);
    }
    engines = analyzeEngines(probe.result);
  }

  const engine = chooseEngine({ isPdf, engines });
  if (engine === 'none') {
    throw new OpsError(
      'E1001',
      '本机未检测到 AutoCAD/中望/浩辰 COM，也未检测到 ODAFileConverter，且输入不是 PDF',
      '可安装 CAD 或 ODAFileConverter，或先用其它工具把图纸导出为 PDF 后使用 PDF 级引擎',
    );
  }
  if (engine === 'oda' && !engines.oda.found) {
    throw new OpsError('E1002', '未找到 ODAFileConverter', '请在设置中填写 odaPath（ODAFileConverter.exe 绝对路径）');
  }

  const tmpDir = join(TMP_DIR, `cad-${timestamp()}-${randomSuffix()}`);
  assertInsideRoot(tmpDir, { write: true, root });

  return {
    mode,
    inputAbs,
    outputAbs,
    paper,
    sortMode,
    parallel,
    engine,
    engines,
    isPdf,
    files,
    filesTotal: mode === 'merge' ? files.length : 1,
    chain: engineChain(engines, isPdf),
    tmpDir,
  };
}

// ───────────────────────── 任务启动 ─────────────────────────

/**
 * suite.cad.convert 操作：校验 + 注册任务 + 异步启动（立即返回 taskId）。
 * @returns {Promise<{ok:true,result:{taskId:string,...}}|{ok:false,error:object}>}
 */
export async function convertCad(ctx, args, exec, config = {}, store) {
  try {
    const spec = await prepareCadSpec(ctx, args, exec, config);
    const task = store.create({
      kind: 'cad.convert',
      mode: spec.mode,
      engine: spec.engine,
      input: spec.inputAbs,
      output: spec.outputAbs,
    });
    store.appendLog(task.id, 'info', `任务创建：${spec.mode} ${spec.inputAbs} → ${spec.outputAbs}`);
    store.appendLog(task.id, 'info', `引擎链路：${spec.chain.join(' → ')}（选定 ${spec.engine}；parallel=${spec.parallel}）`);
    if (spec.engine !== spec.chain[0]) {
      store.appendLog(task.id, 'warn', `engineFallback: ${spec.chain.join(' → ')}（主链路不可用，降级到 ${spec.engine}）`);
    }
    if (spec.mode === 'merge') store.tick(task.id, { filesTotal: spec.filesTotal });
    setTimeout(() => {
      void runCadTask({ ctx, task, spec, config, store });
    }, 0);
    return ok({
      taskId: task.id,
      mode: spec.mode,
      engine: spec.engine,
      chain: spec.chain,
      input: spec.inputAbs,
      output: spec.outputAbs,
      filesTotal: spec.filesTotal,
    });
  } catch (err) {
    const payload = toErrorPayload(err);
    return fail(payload.code, payload.message, payload.hint);
  }
}

/** 任务主流程：按场景分发；所有异常映射到任务状态。 */
export async function runCadTask({ ctx, task, spec, config = {}, store }) {
  const { signal } = task.controller;
  const log = (level, msg) => {
    try { store.appendLog(task.id, level, msg); } catch { /* 任务可能已被清理 */ }
  };
  try {
    store.setState(task.id, 'running');
    if (spec.mode === 'split') {
      await runSplitScenario({ ctx, task, spec, config, store, log });
    } else {
      await runMergeScenario({ ctx, task, spec, config, store, log });
    }
    store.finish(task.id, spec.outputAbs);
    log('info', `任务完成：${spec.outputAbs}`);
  } catch (err) {
    if (isAbortError(err) || signal.aborted) {
      const current = store.get(task.id);
      if (current && current.state !== 'cancelled' && current.state !== 'done' && current.state !== 'failed') {
        try { store.cancel(task.id); } catch { /* 已终态 */ }
      }
      log('warn', '任务已取消');
      return;
    }
    const payload = toErrorPayload(err);
    store.fail(task.id, payload);
    log('error', `任务失败 ${payload.code}: ${payload.message}`);
  } finally {
    if (config.keepTempFiles) {
      log('info', `保留临时目录：${spec.tmpDir}`);
    } else {
      try { await rm(spec.tmpDir, { recursive: true, force: true }); } catch { /* 清理失败忽略 */ }
    }
  }
}

// ───────────────────────── 场景 A：分割 ─────────────────────────

async function runSplitScenario({ ctx, task, spec, config, store, log }) {
  const { signal } = task.controller;
  store.setState(task.id, 'converting');
  let pdfPath = spec.inputAbs;

  if (spec.engine === 'com') {
    const frameDir = join(spec.tmpDir, 'frames');
    await mkdir(frameDir, { recursive: true });
    store.tick(task.id, { currentFile: basename(spec.inputAbs) });
    const payload = {
      mode: 'split',
      input: spec.inputAbs,
      outputDir: frameDir,
      paper: spec.paper,
      sortMode: spec.sortMode,
      frameBlockNames: config.frameBlockNames,
      frameSizes: config.frameSizes,
      frameToleranceMm: config.frameToleranceMm,
      titleBlockZone: config.titleBlockZone,
      drawingNoPatterns: config.drawingNoPatterns,
      cadTimeoutMs: config.cadTimeoutMs,
      pc3: config.pc3 ?? '',
    };
    log('info', `COM 引擎：cad-convert.ps1（场景 A，逐图框窗口打印）`);
    const result = await runPs1(ctx, 'cad-convert.ps1', payload, {
      signal,
      timeoutMs: Number.isFinite(config.cadTimeoutMs) ? config.cadTimeoutMs : DEFAULT_CAD_TIMEOUT_MS,
    });
    const parsed = parseJson(result.stdout);
    if (!parsed) {
      throw new OpsError('E1051', result.stderr || result.stdout || 'cad-convert.ps1 未返回可解析的 JSON');
    }
    if (!parsed.ok) {
      throw new OpsError(parsed.error?.code ?? 'E1004', parsed.error?.message ?? 'COM 打印失败', parsed.error?.hint);
    }
    const frames = Array.isArray(parsed.result?.frames) ? parsed.result.frames : [];
    if (!frames.length) throw new OpsError('E1012', 'COM 引擎未识别到任何图框');
    store.tick(task.id, { framesTotal: frames.length });
    log('info', `COM 引擎产出 ${frames.length} 个图框 PDF，开始合并`);
    store.setState(task.id, 'merging');
    const merged = await mergePdfs({
      files: frames.map((frame, index) => ({
        path: frame.path,
        fileName: basename(frame.path) || `frame-${index + 1}.pdf`,
        drawingNo: frame.drawingNo ?? null,
      })),
      outputPath: spec.outputAbs,
      paper: 'keep',
      sortMode: 'given',
      config,
      signal,
      onProgress: (patch) => store.tick(task.id, patch),
      log,
    });
    store.tick(task.id, { framesDone: frames.length, framesTotal: frames.length });
    return merged;
  }

  if (spec.engine === 'oda') {
    const staging = join(spec.tmpDir, 'oda-in');
    const outDir = join(spec.tmpDir, 'oda-out');
    await mkdir(staging, { recursive: true });
    await copyFile(spec.inputAbs, join(staging, basename(spec.inputAbs)));
    await convertWithOda(ctx, { inDir: staging, outDir, config, signal, log });
    const expected = join(outDir, `${basename(spec.inputAbs, extname(spec.inputAbs))}.pdf`);
    if (!(await exists(expected))) {
      const produced = await readdir(outDir).catch(() => []);
      const pdf = produced.find((name) => name.toLowerCase().endsWith('.pdf'));
      if (!pdf) {
        throw new OpsError('E1051', 'ODAFileConverter 未产出 PDF', `输出目录: ${outDir}`);
      }
      pdfPath = join(outDir, pdf);
    } else {
      pdfPath = expected;
    }
    log('info', `ODA 引擎产出：${pdfPath}`);
  }

  store.setState(task.id, 'splitting');
  const result = await splitPdfFile({
    inputPath: pdfPath,
    outputPath: spec.outputAbs,
    paper: spec.paper,
    sortMode: spec.sortMode,
    config,
    signal,
    onProgress: (patch) => store.tick(task.id, patch),
    log,
  });
  const frames = result?.result?.frames?.length ?? 0;
  store.tick(task.id, { framesDone: frames, framesTotal: frames });
  return result;
}

// ───────────────────────── 场景 B：合并 ─────────────────────────

async function runMergeScenario({ ctx, task, spec, config, store, log }) {
  const { signal } = task.controller;
  store.setState(task.id, 'converting');
  let pdfEntries = [];

  if (spec.engine === 'com') {
    const outDir = join(spec.tmpDir, 'com-pdf');
    await mkdir(outDir, { recursive: true });
    store.tick(task.id, { filesTotal: spec.files.length, filesDone: 0 });
    for (let index = 0; index < spec.files.length; index++) {
      throwIfAborted(signal);
      const file = spec.files[index];
      const outPath = join(outDir, `${String(index + 1).padStart(3, '0')}-${basename(file, extname(file))}.pdf`);
      store.tick(task.id, { currentFile: basename(file) });
      log('info', `COM 打印 ${index + 1}/${spec.files.length}: ${basename(file)}`);
      const payload = {
        mode: 'merge',
        input: file,
        output: outPath,
        paper: spec.paper,
        cadTimeoutMs: config.cadTimeoutMs,
        pc3: config.pc3 ?? '',
      };
      const result = await runPs1(ctx, 'cad-convert.ps1', payload, {
        signal,
        timeoutMs: Number.isFinite(config.cadTimeoutMs) ? config.cadTimeoutMs : DEFAULT_CAD_TIMEOUT_MS,
      });
      const parsed = parseJson(result.stdout);
      if (!parsed) {
        throw new OpsError('E1051', result.stderr || result.stdout || 'cad-convert.ps1 未返回可解析的 JSON');
      }
      if (!parsed.ok) {
        throw new OpsError(parsed.error?.code ?? 'E1004', parsed.error?.message ?? `打印失败: ${basename(file)}`, parsed.error?.hint);
      }
      pdfEntries.push({ path: outPath, fileName: basename(file), drawingNo: null });
      store.tick(task.id, { filesDone: index + 1 });
    }
  } else if (spec.engine === 'oda') {
    const outDir = join(spec.tmpDir, 'oda-out');
    const staging = join(spec.tmpDir, 'oda-in');
    await mkdir(staging, { recursive: true });
    store.tick(task.id, { filesTotal: spec.files.length, filesDone: 0, currentFile: basename(spec.files[0] ?? '') });
    for (const file of spec.files) {
      throwIfAborted(signal);
      await copyFile(file, join(staging, basename(file)));
    }
    await convertWithOda(ctx, { inDir: staging, outDir, config, signal, log });
    const produced = await readdir(outDir).catch(() => []);
    const byName = new Map(produced.map((name) => [name.toLowerCase(), name]));
    for (const file of spec.files) {
      const expected = `${basename(file, extname(file)).toLowerCase()}.pdf`;
      const actual = byName.get(expected);
      if (!actual) {
        log('warn', `ODA 未产出对应 PDF，跳过：${basename(file)}`);
        continue;
      }
      pdfEntries.push({ path: join(outDir, actual), fileName: basename(file), drawingNo: null });
    }
    if (!pdfEntries.length) {
      throw new OpsError('E1051', 'ODAFileConverter 未产出任何 PDF', `输出目录: ${outDir}`);
    }
    store.tick(task.id, { filesDone: spec.files.length, filesTotal: spec.files.length });
  } else {
    pdfEntries = spec.files.map((file) => ({ path: file, fileName: basename(file), drawingNo: null }));
    store.tick(task.id, { filesTotal: pdfEntries.length, filesDone: pdfEntries.length });
  }

  if (pdfEntries.length < spec.files.length) {
    log('warn', `${spec.files.length - pdfEntries.length} 个输入文件未产出 PDF（已跳过）`);
  }
  log('info', `引擎降级检查：${spec.chain.join(' → ')}（实际 ${spec.engine}）`);
  store.setState(task.id, 'merging');
  const merged = await mergePdfs({
    files: pdfEntries,
    outputPath: spec.outputAbs,
    paper: spec.paper,
    sortMode: spec.sortMode,
    config,
    signal,
    concurrency: resolveParallel(config),
    onProgress: (patch) => store.tick(task.id, patch),
    log,
  });
  return merged;
}

// ───────────────────────── ODA 引擎 ─────────────────────────

/** 解析 ODAFileConverter 路径：config.odaPath 优先，其次探测结果。 */
export async function resolveOdaConverter(ctx, config = {}, engines = null) {
  const configured = typeof config.odaPath === 'string' ? config.odaPath.trim() : '';
  if (configured) {
    if (!(await exists(configured))) {
      throw new OpsError('E1002', `配置的 ODAFileConverter 不存在: ${configured}`, '请在设置中更新 odaPath');
    }
    return configured;
  }
  let converter = engines?.oda?.converter ?? null;
  if (!converter) {
    const probe = await probeEngines(ctx, config);
    if (probe.ok) converter = analyzeEngines(probe.result).oda.converter;
  }
  if (!converter) {
    throw new OpsError('E1002', '未找到 ODAFileConverter', '请安装 ODA File Converter 或在设置中填写 odaPath');
  }
  return converter;
}

/**
 * 调用 ODAFileConverter（整目录转换）。
 * 命令形态（ODA 官方）：
 *   ODAFileConverter <inDir> <outDir> <version> <type> <recurse> <audit> [filter]
 * 注：设计文档要求 DWG→PDF；ODAFileConverter 的 PDF 输出能力依本机安装版本而定，
 *     失败时按 E1051 上报（待实测，见交付报告遗留项）。
 */
async function convertWithOda(ctx, { inDir, outDir, config, signal, log }) {
  await mkdir(outDir, { recursive: true });
  const converter = await resolveOdaConverter(ctx, config);
  const version = typeof config.odaVersion === 'string' && config.odaVersion ? config.odaVersion : 'ACAD2018';
  const recurse = config.odaRecurse === false ? '0' : '1';
  const argv = [inDir, outDir, version, 'PDF', recurse, '0', '*.DWG'];
  log('info', `ODA 转换：${converter} ${argv.join(' ')}`);
  const result = await runExe(ctx, converter, argv, {
    signal,
    timeoutMs: Number.isFinite(config.cadTimeoutMs) ? config.cadTimeoutMs : DEFAULT_CAD_TIMEOUT_MS,
    cwd: dirname(converter),
  });
  if (result.exitCode !== 0) {
    throw new OpsError(
      'E1051',
      `ODAFileConverter 失败（exit ${result.exitCode ?? result.signal}）`,
      (result.stderr || result.stdout || '').slice(0, 500) || '请确认 ODA 版本支持 PDF 输出',
    );
  }
  return outDir;
}

// ───────────────────────── 状态 / 取消 ─────────────────────────

/** suite.cad.status。 */
export function cadStatus(store, args = {}) {
  const taskId = typeof args.taskId === 'string' ? args.taskId.trim() : '';
  if (!taskId) {
    return ok({ tasks: store.list() });
  }
  const snapshot = store.snapshot(taskId);
  if (!snapshot) {
    return fail('TASK_NOT_FOUND', `任务不存在或已过期: ${taskId}`);
  }
  return ok(snapshot);
}

/** suite.cad.cancel。 */
export function cadCancel(store, args = {}) {
  const taskId = typeof args.taskId === 'string' ? args.taskId.trim() : '';
  if (!taskId) return fail('E1041', '缺少 taskId');
  try {
    return ok(store.cancel(taskId));
  } catch (err) {
    const payload = toErrorPayload(err);
    return fail(payload.code, payload.message, payload.hint);
  }
}
