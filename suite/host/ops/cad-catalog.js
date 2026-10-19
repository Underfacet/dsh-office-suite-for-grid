/**
 * 图纸目录表（docs/P3-设计.md §2）
 *
 * 命令契约：suite.cad.catalog / buildCatalogOp(ctx, config, args, exec)
 *   参数：{ input: .dwg/.pdf 单文件或目录, engine?: 'auto'|'com'|'pdf', output?: .xlsx, paper?, sort? }
 *   返回：{ outputPath, engine, sources:[{file,frames,withDrawingNo}], rows, sheets, fields, warnings }
 *
 * 数据获取（引擎只提供文本，抽取规则统一在 Node，避免 COM/PDF 两处实现漂移）：
 *   - .pdf / ODA 产出：pdf-split.js 的 detectFramesOnPage + getPageTextItems → 每框相对坐标；
 *   - .dwg + COM：cad-convert.ps1 mode=catalog（P3 §1.3-2，**待真机联调**，本文件只预留调用契约）；
 *   - .dwg + ODA：ODAFileConverter 转整图 PDF → 同上；
 *   - .dwg 且无任何引擎：E1001（同 cad.convert 语义）。
 *
 * 写表：engines/catalog-write.py（venv python，openpyxl）；venv 不可用 → 纯 Node CSV 降级（UTF-8 BOM）。
 * 错误码：E1021/E1023/E1041（路径参数）、E1001（无引擎）、E1012（无图框）、
 *         E1061 CATALOG_EMPTY、E1062 CATALOG_WRITE_FAILED、E1051。
 *
 * 导出：
 *   extractCatalogFields(kind, source)        统一字段抽取（kind='pdf'|'com'）
 *   extractCatalogTitle(items, options)       图名（纯函数）
 *   extractCatalogScale(items, options)       比例（纯函数）
 *   readCatalogFromPdf(inputPath, config, o)  PDF 级逐框抽取（测试与 buildCatalogOp 共用）
 *   sortCatalogRows(rows, sortMode) / buildCatalogSheets(rows, sources) / rowsToCsv(rows)
 *   buildCatalogOp(ctx, config, args, exec)
 */
import { copyFile, mkdir, readdir, readFile, rm, stat, writeFile } from 'node:fs/promises';
import { basename, dirname, extname, isAbsolute, join, resolve } from 'node:path';
import { runExe, runPs1 } from '../engines.js';
import { assertInsideRoot, resolveBoundaryRoot } from '../boundary.js';
import {
  OpsError,
  PACKAGE_DIR,
  TMP_DIR,
  fail,
  isAbortError,
  ok,
  parseJson,
  resolveInputPath,
  sessionCwd,
  throwIfAborted,
  toErrorPayload,
} from '../ops.js';
import { analyzeEngines, probeEngines, resolveOdaConverter } from './cad.js';
import {
  MM_PER_PT,
  compareNatural,
  extractDrawingNo,
  matchFrameSize,
  mergeTextLines,
  normalizeZone,
  toFrameRelativeItems,
  zOrderFrames,
} from './drawingNo.js';
import { detectFramesOnPage, getPageTextItems, openPdfjs, pageSize } from './pdf-split.js';
import { venvStatus } from './venv.js';

// ───────────────────────── 常量与配置 ─────────────────────────

/** 目录表表头（§2.4；顺序即列序）。 */
export const CATALOG_HEADERS = Object.freeze(['序号', '图号', '图名', '图幅', '比例', '张数', '来源文件']);

/** 单次目录表最多来源文件数与图框数（§2.1 规模上限）。 */
export const CATALOG_MAX_FILES = 100;
export const DEFAULT_CATALOG_MAX_FRAMES = 500;

/** 图名区默认相对坐标（§2.3，原点左下）。 */
export const DEFAULT_CATALOG_TITLE_ZONE = Object.freeze([0.3, 0.25, 1.0, 0.55]);

/** 比例提取默认正则（§2.3）：先「比例: x:x」，再全框裸写比号（左值 1–3 位，避免匹配图号连字符）。 */
export const DEFAULT_CATALOG_SCALE_PATTERNS = Object.freeze([
  '比例\\s*[:：]?\\s*(\\d{1,3}\\s*[:：]\\s*\\d{1,4})',
  '(?<![\\dA-Za-z])(\\d{1,3}\\s*[:：]\\s*\\d{1,4})(?![\\dA-Za-z])',
]);

/** P3 新增配置键的模块默认值（DEFAULT_CONFIG 由父智能体在 ops.js 增补，本模块不越权修改）。 */
export const CATALOG_DEFAULTS = Object.freeze({
  catalogTitleNameZone: [...DEFAULT_CATALOG_TITLE_ZONE],
  catalogScalePatterns: [...DEFAULT_CATALOG_SCALE_PATTERNS],
  catalogMaxFrames: DEFAULT_CATALOG_MAX_FRAMES,
});

/** 写入引擎超时（毫秒）。 */
const CATALOG_WRITE_TIMEOUT_MS = 180_000;
/** 目录扫描排除项（与 cad.js 同口径）。 */
const EXCLUDED_FILE = /^~\$|\.(bak|dwl|dwl2|tmp)$/i;
const CATALOG_EXTENSIONS = new Set(['.pdf', '.dwg']);
const CJK_RE = /[\u4e00-\u9fa5]/g;
/** 图名候选行的排除标签（图签字段/日期/比例行不是图名）。 */
const TITLE_NOISE_LABEL_RE = /(图\s*号|编\s*号|图\s*别|日\s*期|比\s*例|比\s*列|设计|制图|校\s*对|审\s*核|批\s*准|会\s*签)/;
const TITLE_LABEL_PREFIX_RE = /^(?:图\s*名|工程名称|图纸名称|项目名称)\s*[:：]?\s*/;
const SCALE_DATE_RE = /(日\s*期|时\s*间|年\s*月|共\s*\d+\s*页)/;

/** 规范化目录表配置（缺省键回落模块默认值；不修改入参）。 */
export function catalogEffectiveConfig(config = {}) {
  const source = config && typeof config === 'object' ? config : {};
  const hasTitleZone = Array.isArray(source.catalogTitleNameZone) && source.catalogTitleNameZone.length >= 4;
  const hasScalePatterns = Array.isArray(source.catalogScalePatterns) && source.catalogScalePatterns.length > 0;
  const maxFrames = Number(source.catalogMaxFrames);
  return {
    ...CATALOG_DEFAULTS,
    ...source,
    catalogTitleNameZone: hasTitleZone ? source.catalogTitleNameZone : CATALOG_DEFAULTS.catalogTitleNameZone,
    catalogScalePatterns: hasScalePatterns ? source.catalogScalePatterns : CATALOG_DEFAULTS.catalogScalePatterns,
    catalogMaxFrames: Number.isFinite(maxFrames) && maxFrames > 0 ? Math.floor(maxFrames) : CATALOG_DEFAULTS.catalogMaxFrames,
  };
}

// ───────────────────────── 纯函数：图名 / 比例 ─────────────────────────

function cjkCount(text) {
  const matches = String(text ?? '').match(CJK_RE);
  return matches ? matches.length : 0;
}

function compileScalePatterns(patterns) {
  const list = Array.isArray(patterns) && patterns.length ? patterns : DEFAULT_CATALOG_SCALE_PATTERNS;
  const compiled = [];
  for (const item of list) {
    if (item instanceof RegExp) {
      compiled.push(item);
      continue;
    }
    try {
      compiled.push(new RegExp(String(item), 'i'));
    } catch { /* 非法正则忽略（配置容错） */ }
  }
  return compiled;
}

/** 比例值归一化：去空白、全角冒号 → 半角。 */
export function normalizeScaleText(raw) {
  if (raw == null) return null;
  const text = String(raw).replace(/\s+/g, '').replace(/：/g, ':');
  return /^\d{1,3}:\d{1,4}$/.test(text) ? text : null;
}

/**
 * 图名抽取（§2.3）：图名区（默认 [0.3,0.25,1.0,0.55]）内取**中文最多的文本行**（≥2 汉字），
 * 排除纯数字/比例/日期/图签标签行；去掉「图名：」等前缀。
 * @param {Array<{text:string,x:number,y:number,height?:number}>} items 相对图框坐标（0..1，原点左下）
 * @param {{zone?:number[],lineToleranceRatio?:number}} [options]
 * @returns {string|null}
 */
export function extractCatalogTitle(items, options = {}) {
  const zone = normalizeZone(options.zone ?? DEFAULT_CATALOG_TITLE_ZONE);
  const [x0, y0, x1, y1] = zone;
  const selected = [];
  for (const item of Array.isArray(items) ? items : []) {
    if (!item) continue;
    const text = String(item.text ?? '');
    if (!text.trim()) continue;
    const x = Number(item.x);
    const y = Number(item.y);
    if (!Number.isFinite(x) || !Number.isFinite(y)) continue;
    if (x < x0 || x > x1 || y < y0 || y > y1) continue;
    selected.push({ text, x, y, height: Number(item.height) || 0 });
  }
  if (!selected.length) return null;
  const lines = mergeTextLines(selected, { lineToleranceRatio: options.lineToleranceRatio });
  let best = null;
  for (const line of lines) {
    let text = String(line.text ?? '');
    text = text.replace(TITLE_LABEL_PREFIX_RE, '').replace(/\s+/g, ' ').trim();
    if (cjkCount(text) < 2) continue;
    if (TITLE_NOISE_LABEL_RE.test(text)) continue;
    if (SCALE_DATE_RE.test(text)) continue;
    if (compileScalePatterns(options.scalePatterns).some((pattern) => pattern.test(text))) continue;
    const score = cjkCount(text);
    const noise = text.replace(CJK_RE, '').trim().length;
    if (!best
      || score > best.score
      || (score === best.score && noise < best.noise)
      || (score === best.score && noise === best.noise && text.length > best.text.length)) {
      best = { text, score, noise };
    }
  }
  return best ? best.text : null;
}

/**
 * 比例抽取（§2.3）：全框文本按「比例」前缀优先，再裸写比号（排除日期/时间行）。
 * @param {Array<{text:string,x:number,y:number,height?:number}>} items 相对图框坐标（0..1）
 * @param {{patterns?:Array<string|RegExp>}} [options]
 * @returns {string|null} 形如 '1:100'
 */
export function extractCatalogScale(items, options = {}) {
  const patterns = compileScalePatterns(options.patterns);
  const lines = mergeTextLines(items);
  for (const pattern of patterns) {
    for (const line of lines) {
      const text = String(line.text ?? '');
      if (!text.trim()) continue;
      if (SCALE_DATE_RE.test(text)) continue;
      const match = text.match(pattern);
      if (!match) continue;
      const value = normalizeScaleText(match[1] ?? match[0]);
      if (value) return value;
    }
  }
  return null;
}

// ───────────────────────── 统一字段抽取（PDF / COM） ─────────────────────────

/** 图幅展示标签：A4（横）/ A3（竖）。 */
export function formatFrameSizeLabel(frameSize, orientation) {
  const name = frameSize == null || frameSize === '' ? '' : String(frameSize);
  if (!name) return '';
  return `${name}（${orientation === 'portrait' ? '竖' : '横'}）`;
}

function frameSizeFromFrame(frame) {
  if (!frame) return null;
  if (frame.widthMm > 0 && frame.heightMm > 0) {
    return matchFrameSize(frame.widthMm, frame.heightMm);
  }
  if (Number(frame.width) > 0 && Number(frame.height) > 0) {
    return matchFrameSize(Number(frame.width) * MM_PER_PT, Number(frame.height) * MM_PER_PT);
  }
  return null;
}

function normalizeRelativeItems(items) {
  const relative = [];
  for (const item of Array.isArray(items) ? items : []) {
    if (!item) continue;
    const text = String(item.text ?? '');
    if (!text.trim()) continue;
    const x = Number(item.x);
    const y = Number(item.y);
    if (!Number.isFinite(x) || !Number.isFinite(y)) continue;
    relative.push({ text, x, y, height: Number(item.height) || 0 });
  }
  return relative;
}

/**
 * 统一目录字段抽取（§2.3）：图号 / 图名 / 图幅 / 比例（张数由调用方按来源文件统计）。
 *
 * @param {'pdf'|'com'} kind 数据来源（'pdf' = PDF 级，'com' = CAD COM 级）
 * @param {object} source
 *   pdf: { frame, pageItems, frameSize?, orientation?, fileName?, config? }
 *   com: { items（相对图框坐标 0..1）, frameSize?, orientation?, drawingNo?, fileName?, config? }
 * @returns {{drawingNo:string|null,title:string|null,titleSource:'text'|'filename'|null,
 *            frameSize:string|null,orientation:'landscape'|'portrait'|null,scale:string|null,
 *            texts:number,kind:string,enginePending?:boolean}}
 */
export function extractCatalogFields(kind, source = {}) {
  if (kind !== 'pdf' && kind !== 'com') {
    throw new OpsError('E1041', "extractCatalogFields kind 必须为 'pdf' | 'com'", `收到: ${JSON.stringify(kind)}`);
  }
  const config = catalogEffectiveConfig(source.config ?? {});
  const frame = source.frame ?? null;

  let relative;
  let frameSize = source.frameSize ?? null;
  let orientation = source.orientation ?? null;
  if (kind === 'pdf') {
    if (!frame) throw new OpsError('E1041', "extractCatalogFields('pdf', ...) 缺少 frame");
    relative = normalizeRelativeItems(toFrameRelativeItems(source.pageItems ?? source.items ?? [], frame));
    if (!frameSize || !orientation) {
      const matched = frameSizeFromFrame(frame);
      if (matched) {
        frameSize = frameSize ?? matched.name;
        orientation = orientation ?? matched.orientation;
      }
    }
  } else {
    // COM 级（cad-convert.ps1 mode=catalog）：引擎直接给相对图框坐标文本项；本接口 P3 预留，待真机联调。
    relative = normalizeRelativeItems(source.items);
  }

  const drawingNo = source.drawingNo != null && String(source.drawingNo).trim()
    ? String(source.drawingNo).trim()
    : extractDrawingNo(relative, {
      titleBlockZone: source.titleBlockZone ?? config.titleBlockZone,
      drawingNoPatterns: source.drawingNoPatterns ?? config.drawingNoPatterns,
    });

  const titleText = extractCatalogTitle(relative, {
    zone: source.titleZone ?? config.catalogTitleNameZone,
    scalePatterns: config.catalogScalePatterns,
  });
  const fallbackTitle = source.fileName
    ? basename(String(source.fileName), extname(String(source.fileName)))
    : null;
  const title = titleText ?? fallbackTitle;
  const scale = extractCatalogScale(relative, { patterns: config.catalogScalePatterns });

  return {
    drawingNo: drawingNo ?? null,
    title: title ?? null,
    titleSource: titleText ? 'text' : (fallbackTitle ? 'filename' : null),
    frameSize: frameSize ?? null,
    orientation: orientation ?? null,
    scale,
    texts: relative.length,
    kind,
    ...(kind === 'com' ? { enginePending: true } : {}),
  };
}

// ───────────────────────── PDF 级读取 ─────────────────────────

/**
 * 逐页检测图框 → 逐框抽取目录字段（PDF 级；buildCatalogOp 与单测共用）。
 * @param {string} inputPath PDF 绝对路径
 * @param {object} config 有效配置
 * @param {{signal?:AbortSignal, log?:(level:string,msg:string)=>void, onProgress?:(patch:object)=>void}} [options]
 * @returns {Promise<{rows:object[],framesTotal:number,pageCount:number,warnings:string[]}>}
 */
export async function readCatalogFromPdf(inputPath, config = {}, options = {}) {
  const settings = catalogEffectiveConfig(config);
  let bytes;
  try {
    bytes = await readFile(inputPath);
  } catch (err) {
    throw new OpsError('E1023', `无法读取输入 PDF: ${inputPath}`, String(err?.message ?? err));
  }
  throwIfAborted(options.signal);
  const doc = await openPdfjs(bytes);
  const rows = [];
  const warnings = [];
  let pageCount = 0;
  let framesTotal = 0;
  try {
    pageCount = doc.numPages;
    for (let pageNumber = 1; pageNumber <= pageCount; pageNumber++) {
      throwIfAborted(options.signal);
      const page = await doc.getPage(pageNumber);
      try {
        const pageFrames = await detectFramesOnPage(page, {
          sizes: settings.frameSizes,
          toleranceMm: settings.frameToleranceMm,
        });
        if (!pageFrames.length) {
          warnings.push(`第 ${pageNumber} 页：未识别到标准图框`);
          options.log?.('info', `目录表：第 ${pageNumber} 页未识别到标准图框`);
          continue;
        }
        const pageItemList = await getPageTextItems(page);
        const { width, height } = pageSize(page);
        // 页内图框按 Z 形平面序（与 split 引擎一致），保证行序稳定
        const ordered = zOrderFrames(pageFrames);
        for (let frameIndex = 0; frameIndex < ordered.length; frameIndex++) {
          throwIfAborted(options.signal);
          const frame = ordered[frameIndex];
          const fields = extractCatalogFields('pdf', {
            frame,
            pageItems: pageItemList,
            frameSize: frame.frameSize,
            orientation: frame.orientation,
            fileName: inputPath,
            config: settings,
          });
          rows.push({
            ...fields,
            pageIndex: pageNumber - 1,
            frameIndex,
            source: inputPath,
            pageWidth: width,
            pageHeight: height,
            frame: {
              x: Math.round(frame.x * 100) / 100,
              y: Math.round(frame.y * 100) / 100,
              width: Math.round(frame.width * 100) / 100,
              height: Math.round(frame.height * 100) / 100,
            },
          });
          framesTotal += 1;
        }
        options.onProgress?.({ framesTotal });
      } finally {
        try { page.cleanup(); } catch { /* ignore */ }
      }
    }
  } finally {
    try { await doc.destroy(); } catch { /* ignore */ }
  }
  if (!framesTotal) {
    throw new OpsError(
      'E1012',
      `未识别到任何标准图框: ${inputPath}`,
      '目录表需要矢量 PDF（含 A0–A4 标准图框矩形）；扫描件/图片 PDF 不支持',
    );
  }
  return { rows, framesTotal, pageCount, warnings };
}

// ───────────────────────── COM / ODA 读取 ─────────────────────────

/**
 * COM 级读取（cad-convert.ps1 mode=catalog）—— **P3 预留接口，待真机联调**。
 * 契约（P3 §1.3-2）：result.frames = [{ drawingNo?, frameSize?, orientation?, texts:[{text,x,y}] }]
 *                 texts 为相对图框坐标（0..1，原点左下）的全框文本项。
 * @returns {Promise<{rows:object[],framesTotal:number,warnings:string[]}>}
 * @throws {OpsError} E1051/E1004/E1012/E1041（引擎拒绝 mode=catalog 时由调用方决定是否降级）
 */
export async function readCatalogFromCom(ctx, config, inputPath, exec = {}, options = {}) {
  const settings = catalogEffectiveConfig(config);
  const payload = {
    mode: 'catalog', // 待真机联调：当前随包 cad-convert.ps1 仅支持 split/merge（P3 §1.3-2 增补项）
    input: inputPath,
    listBlocks: true,
    frameBlockNames: settings.frameBlockNames,
    frameSizes: settings.frameSizes,
    frameToleranceMm: settings.frameToleranceMm,
    titleBlockZone: settings.titleBlockZone,
    drawingNoPatterns: settings.drawingNoPatterns,
    cadTimeoutMs: settings.cadTimeoutMs,
    pc3: settings.pc3 ?? '',
  };
  const result = await runPs1(ctx, 'cad-convert.ps1', payload, {
    signal: exec?.signal,
    timeoutMs: Number.isFinite(settings.cadTimeoutMs) ? settings.cadTimeoutMs : 300_000,
  });
  const parsed = parseJson(result.stdout);
  if (!parsed) {
    throw new OpsError('E1051', result.stderr || result.stdout || 'cad-convert.ps1 未返回可解析的 JSON');
  }
  if (!parsed.ok) {
    const code = parsed.error?.code ?? 'E1004';
    const message = parsed.error?.message ?? 'COM 目录表读取失败（mode=catalog 待真机联调）';
    const catalogHint = code === 'E1041' && /mode/i.test(message)
      ? '随包 cad-convert.ps1 暂只支持 split/merge；P3 §1.3-2 的 mode=catalog 待真机联调，可先用 CAD/ODA 导出整图 PDF 再生成目录表'
      : parsed.error?.hint;
    throw new OpsError(code, message, catalogHint);
  }
  const frames = Array.isArray(parsed.result?.frames) ? parsed.result.frames : [];
  if (!frames.length) throw new OpsError('E1012', 'COM 引擎未识别到任何图框', '真机联调时用 listBlocks 输出核对图框块名与图幅容差');
  const rows = [];
  const warnings = [];
  frames.forEach((frame, frameIndex) => {
    const fields = extractCatalogFields('com', {
      items: Array.isArray(frame?.texts) ? frame.texts : [],
      frameSize: frame?.frameSize ?? null,
      orientation: frame?.orientation ?? null,
      drawingNo: frame?.drawingNo ?? null,
      fileName: inputPath,
      config: settings,
    });
    rows.push({ ...fields, pageIndex: frame?.pageIndex ?? 0, frameIndex, source: inputPath });
  });
  warnings.push('目录表 COM 级路径（mode=catalog）为 P3 预留接口，尚未真机联调（T-02/D5）');
  options.log?.('warn', '目录表 COM 级路径（cad-convert.ps1 mode=catalog）待真机联调');
  return { rows, framesTotal: rows.length, warnings };
}

/** ODA：整文件转 PDF（与 cad.js 编排同口径；cad.js 的 convertWithOda 未导出，故在此复刻调用）。 */
async function convertDwgWithOda(ctx, config, inputPath, tmpDir, exec = {}) {
  const converter = await resolveOdaConverter(ctx, config);
  const staging = join(tmpDir, 'oda-in');
  const outDir = join(tmpDir, 'oda-out');
  await mkdir(staging, { recursive: true });
  await mkdir(outDir, { recursive: true });
  await copyFile(inputPath, join(staging, basename(inputPath)));
  const version = typeof config.odaVersion === 'string' && config.odaVersion ? config.odaVersion : 'ACAD2018';
  const recurse = config.odaRecurse === false ? '0' : '1';
  const argv = [staging, outDir, version, 'PDF', recurse, '0', '*.DWG'];
  const result = await runExe(ctx, converter, argv, {
    signal: exec?.signal,
    timeoutMs: Number.isFinite(config.cadTimeoutMs) ? config.cadTimeoutMs : 300_000,
    cwd: dirname(converter),
  });
  if (result.exitCode !== 0) {
    throw new OpsError(
      'E1051',
      `ODAFileConverter 失败（exit ${result.exitCode ?? result.signal}）`,
      (result.stderr || result.stdout || '').slice(0, 500) || '请确认 ODA 版本支持 PDF 输出',
    );
  }
  const produced = await readdir(outDir).catch(() => []);
  const pdf = produced.find((name) => name.toLowerCase().endsWith('.pdf'));
  if (!pdf) throw new OpsError('E1051', 'ODAFileConverter 未产出 PDF', `输出目录: ${outDir}`);
  return join(outDir, pdf);
}

// ───────────────────────── 行排序 / 张数 / 工作表 ─────────────────────────

function hasText(value) {
  return value != null && String(value).trim() !== '';
}

function compareCatalogRows(a, b) {
  const bySource = compareNatural(a.source, b.source);
  const byPage = (Number(a.pageIndex) || 0) - (Number(b.pageIndex) || 0);
  const byFrame = (Number(a.frameIndex) || 0) - (Number(b.frameIndex) || 0);
  return bySource || byPage || byFrame;
}

/**
 * 行排序（§2.3）：'drawingNo' 图号优先（null 排后，再按来源文件+页序）；'filename' 按来源文件+页序。
 * @param {object[]} rows
 * @param {'drawingNo'|'filename'} [sortMode]
 */
export function sortCatalogRows(rows, sortMode = 'drawingNo') {
  const list = Array.isArray(rows) ? rows : [];
  const indexed = list.map((row, index) => ({ row, index }));
  indexed.sort((a, b) => {
    if (sortMode === 'filename') return compareCatalogRows(a.row, b.row) || a.index - b.index;
    const an = hasText(a.row.drawingNo) ? String(a.row.drawingNo) : null;
    const bn = hasText(b.row.drawingNo) ? String(b.row.drawingNo) : null;
    if (an && !bn) return -1;
    if (!an && bn) return 1;
    if (an && bn) {
      const byNo = compareNatural(an, bn);
      if (byNo) return byNo;
    }
    return compareCatalogRows(a.row, b.row) || a.index - b.index;
  });
  return indexed.map((entry) => entry.row);
}

/**
 * 张数统计（§2.3）：同一来源文件内按图号计数；图号缺失时按「图名+图幅」组合计数。
 * @param {object[]} rows
 * @returns {object[]} 原地写入 count 并返回同一数组
 */
export function annotateCatalogCounts(rows) {
  const list = Array.isArray(rows) ? rows : [];
  const counters = new Map();
  const keyOf = (row) => (hasText(row.drawingNo)
    ? `no:${String(row.drawingNo)}`
    : `t:${String(row.title ?? '')}|${String(row.frameSize ?? '')}|${String(row.orientation ?? '')}`);
  for (const row of list) {
    const key = `${row.source}\u0000${keyOf(row)}`;
    counters.set(key, (counters.get(key) ?? 0) + 1);
  }
  for (const row of list) {
    row.count = counters.get(`${row.source}\u0000${keyOf(row)}`) ?? 1;
  }
  return list;
}

/** 行 → xlsx 数据行（字段名与 engines/catalog-write.py 的 HEADER_FIELDS 对齐）。 */
export function toCatalogSheetRow(row) {
  return {
    seq: row.seq,
    drawingNo: hasText(row.drawingNo) ? String(row.drawingNo) : null,
    title: row.title != null ? String(row.title) : null,
    frameSize: row.frameSize ?? null,
    frameSizeLabel: formatFrameSizeLabel(row.frameSize, row.orientation),
    orientation: row.orientation ?? null,
    scale: row.scale ?? null,
    count: Number.isFinite(Number(row.count)) ? Number(row.count) : 1,
    source: basename(String(row.source ?? '')),
  };
}

function sheetNameFor(file, used) {
  const raw = basename(String(file ?? ''), extname(String(file ?? ''))).replace(/[\[\]:*?/\\]/g, '_');
  const base = (raw || 'sheet').slice(0, 31);
  if (!used.has(base)) return base;
  for (let suffix = 2; ; suffix += 1) {
    const tail = `(${suffix})`;
    const name = base.length + tail.length > 31 ? base.slice(0, 31 - tail.length) + tail : base + tail;
    if (!used.has(name)) return name;
  }
}

/**
 * 工作表组织（§2.4）：来源文件数 >1 时「汇总」+ 每来源一个 sheet；单来源只有「汇总」。
 * @param {object[]} rows 已排序带 seq 的行
 * @param {Array<{file:string}>} sources
 * @returns {Array<{name:string,rows:object[]}>}
 */
export function buildCatalogSheets(rows, sources) {
  const list = Array.isArray(rows) ? rows : [];
  const summary = { name: '汇总', rows: list.map(toCatalogSheetRow) };
  const sourceList = Array.isArray(sources) ? sources : [];
  if (sourceList.length <= 1) return [summary];
  const used = new Set(['汇总']);
  const sheets = [summary];
  for (const source of sourceList) {
    const name = sheetNameFor(source.file, used);
    used.add(name);
    sheets.push({
      name,
      rows: list.filter((row) => row.source === source.file).map(toCatalogSheetRow),
    });
  }
  return sheets;
}

function csvCell(value) {
  const text = value == null ? '' : String(value);
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

/**
 * 纯 Node CSV 降级（venv 不可用时；UTF-8 BOM + CRLF，仅「汇总」表）。
 * @param {object[]} rows 数据行
 * @returns {string}
 */
export function rowsToCsv(rows) {
  const lines = [CATALOG_HEADERS.join(',')];
  for (const row of Array.isArray(rows) ? rows : []) {
    const sheetRow = toCatalogSheetRow(row);
    lines.push([
      sheetRow.seq,
      sheetRow.drawingNo,
      sheetRow.title,
      sheetRow.frameSizeLabel,
      sheetRow.scale,
      sheetRow.count,
      sheetRow.source,
    ].map(csvCell).join(','));
  }
  return `\uFEFF${lines.join('\r\n')}\r\n`;
}

// ───────────────────────── 参数 / 路径工具 ─────────────────────────

function timestamp() {
  return new Date().toISOString().replace(/[-:T]/g, '').slice(0, 14);
}

function randomSuffix() {
  return Math.random().toString(36).slice(2, 8);
}

function validateCatalogEngine(value) {
  if (value === undefined || value === null || value === '') return 'auto';
  if (value === 'auto' || value === 'com' || value === 'pdf') return value;
  throw new OpsError('E1041', "engine 必须为 'auto' | 'com' | 'pdf'", `收到: ${JSON.stringify(value)}`);
}

function validateCatalogSort(value) {
  if (value === undefined || value === null || value === '') return 'drawingNo';
  if (value === 'drawingNo' || value === 'filename') return value;
  throw new OpsError('E1041', "sort 必须为 'drawingNo' | 'filename'", `收到: ${JSON.stringify(value)}`);
}

function validateCatalogPaper(value) {
  if (value === undefined || value === null || value === '') return undefined;
  if (value === 'A4-landscape' || value === 'keep') return value;
  throw new OpsError('E1041', "paper 必须为 'A4-landscape' | 'keep'", `收到: ${JSON.stringify(value)}`);
}

async function exists(path) {
  try {
    await stat(path);
    return true;
  } catch {
    return false;
  }
}

/** 默认输出根：config.outputRoot 或 会话工作区 output\pdf（语义同 cad.convert）。 */
async function defaultCatalogOutputRoot(ctx, exec, config) {
  if (typeof config?.outputRoot === 'string' && config.outputRoot.trim()) {
    const configured = config.outputRoot.trim();
    return isAbsolute(configured) ? resolve(configured) : resolve(await sessionCwd(ctx, exec), configured);
  }
  return join(await sessionCwd(ctx, exec), 'output', 'pdf');
}

/** 扫描目录内的 .dwg/.pdf（不递归；排除 ~$/.bak/.dwl/.dwl2/.tmp）。 */
async function scanCatalogInput(dir) {
  let dirents;
  try {
    dirents = await readdir(dir, { withFileTypes: true });
  } catch (err) {
    if (err?.code === 'ENOENT') throw new OpsError('E1023', `输入目录不存在: ${dir}`);
    if (err?.code === 'EACCES' || err?.code === 'EPERM') throw new OpsError('E1022', `无权限读取目录: ${dir}`);
    throw new OpsError('E1099', `读取输入目录失败: ${String(err?.message ?? err)}`);
  }
  const files = [];
  for (const dirent of dirents) {
    if (!dirent.isFile()) continue;
    if (EXCLUDED_FILE.test(dirent.name)) continue;
    if (!CATALOG_EXTENSIONS.has(extname(dirent.name).toLowerCase())) continue;
    files.push(resolve(dir, dirent.name));
  }
  return files.sort((a, b) => compareNatural(basename(a), basename(b)));
}

// ───────────────────────── 写表 ─────────────────────────

function catalogFailure(err, fallbackCode = 'E1062') {
  if (err instanceof OpsError) return err;
  return new OpsError(fallbackCode, `目录表写出失败: ${String(err?.message ?? err)}`);
}

/**
 * 写 xlsx：venv python + engines/catalog-write.py；venv 不可用 → Node CSV 降级（§2.4）。
 * 说明：目录表是**同步轻量操作**，这里只读 venv 状态（venvStatus），不触发建 venv/装包；
 *       venv 缺失或版本标记不匹配时直接 CSV 降级，避免命令通道长时间阻塞。
 * @returns {Promise<{engine:'python'|'csv',outputPath:string,sheets:string[],warnings:string[]}>}
 */
export async function writeCatalogWorkbook(ctx, config, exec, targetPath, rows, sources) {
  const sheets = buildCatalogSheets(rows, sources);
  const payload = {
    path: targetPath,
    title: '图纸目录表',
    headers: [...CATALOG_HEADERS],
    sheets,
  };
  const status = venvStatus();
  if (!status.healthy || !status.python) {
    const csvPath = targetPath.replace(/\.xlsx$/i, '.csv');
    await mkdir(dirname(csvPath), { recursive: true });
    await writeFile(csvPath, rowsToCsv(rows), 'utf8');
    return {
      engine: 'csv',
      outputPath: csvPath,
      sheets: ['汇总'],
      warnings: [
        `Python venv 不可用（present=${status.present}，healthy=${status.healthy}），已降级 CSV：无表头样式、仅「汇总」表`,
        '如需 xlsx，请在诊断页重建 venv（suite 依赖 openpyxl）后重试',
      ],
    };
  }
  const python = status.python;
  const script = join(PACKAGE_DIR, 'engines', 'catalog-write.py');
  const tmpDir = join(TMP_DIR, `catalog-${timestamp()}-${randomSuffix()}`);
  await mkdir(tmpDir, { recursive: true });
  const payloadPath = join(tmpDir, 'catalog-payload.json');
  // 走 -JsonFile：500 行中文表头/图名可能超过 Windows 32K 命令行上限；脚本同时支持 stdin 与 -Json
  await writeFile(payloadPath, JSON.stringify(payload), 'utf8');
  try {
    const execution = await runExe(ctx, python, [script, '-JsonFile', payloadPath], {
      signal: exec?.signal,
      timeoutMs: CATALOG_WRITE_TIMEOUT_MS,
      graceMs: 60_000,
    });
    const parsed = parseJson(execution.stdout);
    if (!parsed) {
      throw new OpsError(
        'E1062',
        'catalog-write.py 未返回可解析的 JSON',
        (execution.stderr || execution.stdout || '').slice(0, 600) || undefined,
      );
    }
    if (!parsed.ok) {
      throw new OpsError(
        parsed.error?.code ?? 'E1062',
        parsed.error?.message ?? '目录表写出失败',
        parsed.error?.hint,
      );
    }
    const result = parsed.result ?? {};
    const outputPath = result.outputPath ?? targetPath;
    if (!(await exists(outputPath))) {
      throw new OpsError('E1062', `目录表写出后未找到目标文件: ${outputPath}`);
    }
    return {
      engine: result.engine === 'csv' ? 'csv' : 'python',
      outputPath,
      sheets: Array.isArray(result.sheets) ? result.sheets : sheets.map((sheet) => sheet.name),
      warnings: Array.isArray(result.warnings) ? result.warnings : [],
    };
  } catch (err) {
    throw catalogFailure(err);
  } finally {
    try { await rm(tmpDir, { recursive: true, force: true }); } catch { /* 清理失败忽略 */ }
  }
}

// ───────────────────────── 字段命中统计 ─────────────────────────

function hitStat(rows, field) {
  const total = rows.length;
  let hit = 0;
  for (const row of rows) {
    if (field === 'title') {
      if (row.titleSource === 'text') hit += 1;
    } else if (hasText(row[field])) {
      hit += 1;
    }
  }
  return { hit, total, rate: total ? Math.round((hit / total) * 10000) / 10000 : 0 };
}

function catalogFieldsStats(rows) {
  return {
    drawingNo: hitStat(rows, 'drawingNo'),
    title: hitStat(rows, 'title'),
    frameSize: hitStat(rows, 'frameSize'),
    scale: hitStat(rows, 'scale'),
    count: hitStat(rows, 'count'),
  };
}

// ───────────────────────── 主操作 ─────────────────────────

/**
 * suite.cad.catalog：图纸目录表（同步执行，不走任务注册表；规模上限见 §2.1）。
 * @param {object} ctx DSH 插件上下文
 * @param {object} config 有效配置
 * @param {{input:string,engine?:'auto'|'com'|'pdf',output?:string,paper?:'A4-landscape'|'keep',sort?:'drawingNo'|'filename'}} args
 * @param {object} exec 工具/命令执行上下文（agent/signal）
 * @returns {Promise<{ok:true,result:object}|{ok:false,error:object}>}
 */
export async function buildCatalogOp(ctx, config = {}, args = {}, exec = {}) {
  const startedAt = Date.now();
  const settings = catalogEffectiveConfig(config);
  const warnings = [];
  const tmpDir = join(TMP_DIR, `catalog-${timestamp()}-${randomSuffix()}`);
  try {
    const rawInput = typeof args?.input === 'string' ? args.input.trim() : '';
    if (!rawInput) throw new OpsError('E1041', '缺少 input（.dwg/.pdf 文件或目录路径）');
    const engineArg = validateCatalogEngine(args?.engine);
    const sortMode = validateCatalogSort(args?.sort ?? config.sortMode);
    const paper = validateCatalogPaper(args?.paper);

    const root = resolveBoundaryRoot(config);
    const allowExternalPaths = config.allowExternalPaths === true;
    const inputAbs = assertInsideRoot(await resolveInputPath(ctx, exec, rawInput), {
      write: false,
      allowExternalPaths,
      root,
      warn: (message) => console.warn(`[office-suite] ${message}`),
    });
    if (!(await exists(inputAbs))) throw new OpsError('E1023', `输入路径不存在: ${inputAbs}`);
    const inputStat = await stat(inputAbs);

    let files = [];
    if (inputStat.isDirectory()) {
      files = await scanCatalogInput(inputAbs);
      if (!files.length) {
        throw new OpsError('E1023', `输入目录没有可生成目录表的 .dwg / .pdf 文件: ${inputAbs}`);
      }
    } else {
      const ext = extname(inputAbs).toLowerCase();
      if (!CATALOG_EXTENSIONS.has(ext)) {
        throw new OpsError('E1041', '目录表 input 仅支持 .dwg / .pdf 文件或目录', `收到: ${ext || '(无扩展名)'}`);
      }
      files = [inputAbs];
    }
    if (files.length > CATALOG_MAX_FILES) {
      throw new OpsError(
        'E1041',
        `目录表单次最多 ${CATALOG_MAX_FILES} 个来源文件，收到 ${files.length} 个`,
        '请拆分为多次生成（目录表是同步操作，防止长阻塞命令通道）',
      );
    }

    // 输出路径：指定则校验 .xlsx（或已存在目录）；缺省 outputRoot\catalog-<时间戳>.xlsx
    let outputAbs;
    const outputArg = typeof args?.output === 'string' ? args.output.trim() : '';
    if (outputArg) {
      const candidate = assertInsideRoot(await resolveInputPath(ctx, exec, outputArg), { write: true, root });
      const candidateIsDir = await (async () => {
        try { return (await stat(candidate)).isDirectory(); } catch { return false; }
      })();
      if (candidateIsDir) {
        outputAbs = join(candidate, `catalog-${timestamp()}.xlsx`);
      } else if (extname(candidate).toLowerCase() === '.xlsx') {
        outputAbs = candidate;
      } else {
        throw new OpsError('E1041', `输出必须是 .xlsx 文件: ${candidate}`, 'CSV 降级时由写出引擎自动改用同名 .csv');
      }
    } else {
      outputAbs = join(await defaultCatalogOutputRoot(ctx, exec, settings), `catalog-${timestamp()}.xlsx`);
    }
    outputAbs = assertInsideRoot(outputAbs, { write: true, root });
    if (extname(outputAbs).toLowerCase() !== '.xlsx') {
      throw new OpsError('E1041', `输出必须是 .xlsx 文件: ${outputAbs}`);
    }

    const needsDwg = files.some((file) => extname(file).toLowerCase() === '.dwg');
    let engines = null;
    if (needsDwg) {
      const probe = await probeEngines(ctx, settings);
      if (!probe.ok) {
        throw new OpsError(probe.error.code ?? 'E1051', probe.error.message ?? '环境探测失败', probe.error.hint);
      }
      engines = analyzeEngines(probe.result);
      if (engineArg === 'pdf') {
        throw new OpsError(
          'E1001',
          'PDF 级引擎无法直接读取 .dwg（需要先转换）',
          "可改用 engine='auto'/'com'，或用 CAD/ODA 把图纸导出为 PDF 后再生成目录表",
        );
      }
      if (engineArg === 'com' && !engines.cadCom.available && !engines.oda.found) {
        throw new OpsError(
          'E1001',
          '本机未检测到 AutoCAD/中望/浩辰 COM，也未检测到 ODAFileConverter',
          '可安装 CAD 或 ODAFileConverter，或先把图纸导出为整图 PDF 后用 PDF 级引擎',
        );
      }
      if (!engines.cadCom.available && !engines.oda.found) {
        throw new OpsError(
          'E1001',
          '本机未检测到 AutoCAD/中望/浩辰 COM，也未检测到 ODAFileConverter，无法读取 .dwg',
          '可安装 CAD 或 ODAFileConverter，或先用其它工具把图纸导出为 PDF 后使用 PDF 级引擎',
        );
      }
      if (engineArg === 'com' && !engines.cadCom.available && engines.oda.found) {
        warnings.push("engine='com' 但本机无 CAD COM，已降级 ODA → PDF 级目录表");
      }
      if (engines.cadCom.available) {
        warnings.push('COM 级目录表（cad-convert.ps1 mode=catalog）为 P3 预留接口，尚未真机联调（T-02/D5）');
      }
    } else if (engineArg === 'com') {
      warnings.push("输入均为 PDF，目录表固定走 PDF 级引擎（engine='com' 已忽略）");
    }

    await mkdir(tmpDir, { recursive: true });
    const sources = [];
    const allRows = [];
    let skippedNoFrame = null;
    for (let index = 0; index < files.length; index++) {
      throwIfAborted(exec?.signal);
      const file = files[index];
      const ext = extname(file).toLowerCase();
      const log = (level, message) => {
        if (level === 'warn') warnings.push(message);
        console.warn(`[office-suite] catalog ${message}`);
      };
      let read;
      let readEngine;
      try {
        if (ext === '.pdf') {
          read = await readCatalogFromPdf(file, settings, { signal: exec?.signal, log });
          readEngine = 'pdf';
        } else {
          // .dwg：com → oda 降级链（engineArg 决定首选）
          const preferCom = engines.cadCom.available && (engineArg === 'com' || engineArg === 'auto');
          let comError = null;
          if (preferCom) {
            try {
              read = await readCatalogFromCom(ctx, settings, file, exec, { log });
              readEngine = 'com';
            } catch (err) {
              if (!(err instanceof OpsError)) throw err;
              comError = err;
              if (!engines.oda.found) throw err;
              warnings.push(`COM 目录表读取失败（${err.code}），已降级 ODA → PDF 级：${file}`);
            }
          }
          if (!read) {
            const pdfPath = await convertDwgWithOda(ctx, settings, file, join(tmpDir, `oda-${index + 1}`), exec);
            read = await readCatalogFromPdf(pdfPath, settings, { signal: exec?.signal, log });
            readEngine = 'oda';
            if (comError) warnings.push(`COM 失败原因：${comError.code} ${comError.message}`);
          }
        }
      } catch (err) {
        if (isAbortError(err)) throw err;
        if (!(err instanceof OpsError)) throw err;
        // 单个文件无图框/解析失败：跳过并继续（目录表批量场景），全部失败再上报
        if (err.code === 'E1012' || err.code === 'E1011') {
          warnings.push(`跳过 ${basename(file)}（${err.code}）：${err.message}`);
          skippedNoFrame = skippedNoFrame ?? err;
          continue;
        }
        throw err;
      }
      if (read.warnings?.length) warnings.push(...read.warnings.map((item) => `${basename(file)}：${item}`));
      for (const row of read.rows) {
        allRows.push({ ...row, source: file, readEngine });
      }
      sources.push({
        file,
        frames: read.framesTotal,
        withDrawingNo: read.rows.filter((row) => hasText(row.drawingNo)).length,
        readEngine,
        pages: read.pageCount ?? null,
      });
    }

    if (!allRows.length) {
      if (skippedNoFrame) throw skippedNoFrame;
      throw new OpsError('E1061', '未提取到任何目录字段', '图中未识别到图框文本；请确认输入为矢量 PDF 或含图框块的 DWG');
    }
    if (allRows.length > settings.catalogMaxFrames) {
      throw new OpsError(
        'E1041',
        `目录表图框合计 ${allRows.length} 个，超过上限 ${settings.catalogMaxFrames}`,
        '请拆分为多个文件/目录分批生成（catalogMaxFrames 可调）',
      );
    }
    const extracted = allRows.filter((row) => hasText(row.drawingNo) || row.titleSource === 'text' || hasText(row.scale));
    if (!extracted.length) {
      throw new OpsError(
        'E1061',
        `识别到 ${allRows.length} 个图框，但未提取到任何图号/图名/比例`,
        '请核对图签区/图名区配置（titleBlockZone / catalogTitleNameZone）与图号正则（drawingNoPatterns）',
      );
    }

    const ordered = sortCatalogRows(annotateCatalogCounts(allRows), sortMode);
    ordered.forEach((row, index) => { row.seq = index + 1; });

    const withNo = ordered.filter((row) => hasText(row.drawingNo)).length;
    if (withNo < ordered.length) {
      warnings.push(`E1013 同义告警：${ordered.length - withNo}/${ordered.length} 个图框未提取到图号（排序回退来源文件/页序）`);
    }
    const titleFallback = ordered.filter((row) => row.titleSource === 'filename').length;
    if (titleFallback) {
      warnings.push(`${titleFallback}/${ordered.length} 个图框图名未命中图名区，已回退文件名`);
    }
    const scaleMissed = ordered.filter((row) => !hasText(row.scale)).length;
    if (scaleMissed) warnings.push(`${scaleMissed}/${ordered.length} 个图框未提取到比例`);

    const write = await writeCatalogWorkbook(ctx, settings, exec, outputAbs, ordered, sources);
    if (write.warnings.length) warnings.push(...write.warnings);

    return ok({
      outputPath: write.outputPath,
      engine: write.engine,
      writeEngine: write.engine,
      readEngine: [...new Set(ordered.map((row) => row.readEngine))].join('+') || null,
      ...(paper ? { paper } : {}),
      sort: sortMode,
      sources,
      rows: ordered.length,
      sheets: write.sheets,
      fields: catalogFieldsStats(ordered),
      warnings,
      elapsedMs: Date.now() - startedAt,
    });
  } catch (err) {
    if (isAbortError(err)) throw err;
    const payload = toErrorPayload(err, 'E1062');
    return fail(payload.code, payload.message, payload.hint);
  } finally {
    if (!settings.keepTempFiles) {
      try { await rm(tmpDir, { recursive: true, force: true }); } catch { /* 清理失败忽略 */ }
    }
  }
}
