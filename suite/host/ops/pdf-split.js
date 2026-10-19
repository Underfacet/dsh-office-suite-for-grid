/**
 * PDF 级引擎：多图框 PDF → 每图一页（docs/P1-设计.md §3.5）
 *
 * 解析：pdfjs-dist（legacy Node 构建）扫描矢量矩形（re / 4 段正交闭合路径）+ getTextContent 坐标还原；
 * 裁剪/缩放/合并：pdf-lib（embedPage boundingBox 裁剪 + 白底 + 居中缩放）。
 * 防御：单页字节上限 64MB、每 50 页检查点 save、pdfjs 对象 cleanup/destroy、加密 PDF 报 E1011。
 */
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { dirname, join, sep } from 'node:path';
import { createRequire } from 'node:module';
import { PDFDocument, rgb } from 'pdf-lib';
import { OpsError, isAbortError, throwIfAborted } from '../ops.js';
import {
  A4_LANDSCAPE_PT,
  MM_PER_PT,
  matchFrameSize,
  extractDrawingNo,
  toFrameRelativeItems,
  sortFrames,
  normalizeFrameSizes,
  normalizeZone,
} from './drawingNo.js';

/** 单页字节上限（超过则跳过该页并告警，R-08）。 */
export const MAX_PAGE_BYTES = 64 * 1024 * 1024;
/** 图框最小边长（pt，噪声过滤）。 */
const MIN_FRAME_SIDE_PT = 20;
/** 每 N 页做一次 save 检查点（§3.5）。 */
const CHECKPOINT_EVERY = 50;
/** 矩形匹配容差（pt）。 */
const RECT_CORNER_TOL_PT = 0.05;

const IDENTITY_MATRIX = Object.freeze([1, 0, 0, 1, 0, 0]);

let pdfjsPromise;
let standardFontDataUrlCache;

/** 动态加载 pdfjs-dist legacy 构建（R-04：锁定 legacy/build/pdf.mjs）。 */
export async function loadPdfjs() {
  if (!pdfjsPromise) pdfjsPromise = import('pdfjs-dist/legacy/build/pdf.mjs');
  return pdfjsPromise;
}

/** 标准字体目录（避免 pdfjs 打印 standardFontDataUrl 警告；失败则省略）。 */
function standardFontDataUrl() {
  if (standardFontDataUrlCache !== undefined) return standardFontDataUrlCache;
  try {
    const require = createRequire(import.meta.url);
    const pkg = require.resolve('pdfjs-dist/package.json');
    standardFontDataUrlCache = `${join(dirname(pkg), 'standard_fonts')}${sep}`;
  } catch {
    standardFontDataUrlCache = undefined;
  }
  return standardFontDataUrlCache;
}

function mapPdfOpenError(err) {
  const name = String(err?.name ?? '');
  const message = String(err?.message ?? err ?? '未知错误');
  if (name === 'PasswordException' || /password|encrypt/i.test(message)) {
    return new OpsError('E1011', `加密 PDF 不支持: ${message}`, '请先去除 PDF 密码后重试（P1 不支持加密 PDF）');
  }
  return new OpsError('E1011', `PDF 解析失败: ${message}`);
}

function mapPdfLibError(err) {
  const message = String(err?.message ?? err ?? '未知错误');
  if (/encrypted/i.test(message)) {
    return new OpsError('E1011', `加密 PDF 不支持: ${message}`, '请先去除 PDF 密码后重试（P1 不支持加密 PDF）');
  }
  return new OpsError('E1011', `PDF 解析失败: ${message}`);
}

/** pdf-lib 打开 PDF（加密 → E1011）。 */
export async function loadPdfLib(bytes) {
  try {
    return await PDFDocument.load(new Uint8Array(bytes), { updateMetadata: false });
  } catch (err) {
    throw mapPdfLibError(err);
  }
}

/** pdfjs 打开 PDF（加密 → E1011）。 */
export async function openPdfjs(bytes) {
  const pdfjs = await loadPdfjs();
  const fontUrl = standardFontDataUrl();
  const loadingTask = pdfjs.getDocument({
    data: new Uint8Array(bytes),
    disableFontFace: true,
    isEvalSupported: false,
    useSystemFonts: false,
    verbosity: 0,
    ...(fontUrl ? { standardFontDataUrl: fontUrl } : {}),
  });
  try {
    return await loadingTask.promise;
  } catch (err) {
    try { await loadingTask.destroy(); } catch { /* 已销毁 */ }
    throw mapPdfOpenError(err);
  }
}

/** 页面尺寸（user space，pt）与 UserUnit。 */
export function pageSize(page) {
  const view = Array.isArray(page.view) ? page.view : [0, 0, 612, 792];
  const width = Math.abs(view[2] - view[0]);
  const height = Math.abs(view[3] - view[1]);
  return {
    width: width || 612,
    height: height || 792,
    userUnit: Number(page.userUnit) > 0 ? Number(page.userUnit) : 1,
  };
}

// ───────────────────────── 矢量矩形检测 ─────────────────────────

/** 仿射矩阵乘法：mul(m1, m2) 表示先作用 m2、再作用 m1。 */
function multiply(m1, m2) {
  return [
    m1[0] * m2[0] + m1[2] * m2[1],
    m1[1] * m2[0] + m1[3] * m2[1],
    m1[0] * m2[2] + m1[2] * m2[3],
    m1[1] * m2[2] + m1[3] * m2[3],
    m1[0] * m2[4] + m1[2] * m2[5] + m1[4],
    m1[1] * m2[4] + m1[3] * m2[5] + m1[5],
  ];
}

function applyMatrix(matrix, x, y) {
  return [matrix[0] * x + matrix[2] * y + matrix[4], matrix[1] * x + matrix[3] * y + matrix[5]];
}

/**
 * 把一条（可能闭合的）折线点列判别为轴对齐矩形。
 * 要求：去重后恰为 4 个角点，且每个点都落在四点包围盒的角上。
 */
function polygonToRect(points) {
  if (!Array.isArray(points) || points.length < 4) return null;
  const cleaned = [];
  for (const point of points) {
    if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) return null;
    const last = cleaned[cleaned.length - 1];
    if (last && Math.abs(last.x - point.x) < 1e-6 && Math.abs(last.y - point.y) < 1e-6) continue;
    cleaned.push(point);
  }
  if (cleaned.length > 4) {
    const first = cleaned[0];
    const last = cleaned[cleaned.length - 1];
    if (Math.abs(first.x - last.x) < 1e-6 && Math.abs(first.y - last.y) < 1e-6) cleaned.pop();
  }
  if (cleaned.length !== 4) return null;
  let minX = Infinity; let minY = Infinity; let maxX = -Infinity; let maxY = -Infinity;
  for (const point of cleaned) {
    if (point.x < minX) minX = point.x;
    if (point.y < minY) minY = point.y;
    if (point.x > maxX) maxX = point.x;
    if (point.y > maxY) maxY = point.y;
  }
  const width = maxX - minX;
  const height = maxY - minY;
  if (!(width > 0 && height > 0)) return null;
  const tolerance = Math.max(RECT_CORNER_TOL_PT, Math.max(width, height) * 0.002);
  const cornerKeys = new Set();
  for (const point of cleaned) {
    const atMinX = Math.abs(point.x - minX) <= tolerance;
    const atMaxX = Math.abs(point.x - maxX) <= tolerance;
    const atMinY = Math.abs(point.y - minY) <= tolerance;
    const atMaxY = Math.abs(point.y - maxY) <= tolerance;
    if (!((atMinX || atMaxX) && (atMinY || atMaxY))) return null;
    cornerKeys.add(`${atMaxX ? 1 : 0}${atMaxY ? 1 : 0}`);
  }
  if (cornerKeys.size !== 4) return null;
  return { x: minX, y: minY, width, height };
}

/** 解析 constructPath 的 [fnArray, coords] 为若干矩形（已应用 CTM）。 */
function rectsFromConstructPath(fns, coords, ctm, ops, output) {
  if (!Array.isArray(fns) || !Array.isArray(coords)) return;
  let index = 0;
  let current = null;
  const subpaths = [];
  const begin = () => {
    current = [];
    subpaths.push(current);
  };
  const push = (x, y) => {
    if (!current) begin();
    const [dx, dy] = applyMatrix(ctm, x, y);
    current.push({ x: dx, y: dy });
  };
  for (const fn of fns) {
    if (fn === ops.moveTo) {
      const x = coords[index++];
      const y = coords[index++];
      begin();
      push(x, y);
    } else if (fn === ops.lineTo) {
      const x = coords[index++];
      const y = coords[index++];
      if (!current) begin();
      push(x, y);
    } else if (fn === ops.curveTo) {
      // 折线化：仅取曲线终点（图框通常为直线段；圆角曲线不参与矩形判定）
      index += 6;
      const x = coords[index - 2];
      const y = coords[index - 1];
      if (Number.isFinite(x) && Number.isFinite(y)) {
        if (!current) begin();
        push(x, y);
      }
    } else if (fn === ops.curveTo2 || fn === ops.curveTo3) {
      index += 4;
      const x = coords[index - 2];
      const y = coords[index - 1];
      if (Number.isFinite(x) && Number.isFinite(y)) {
        if (!current) begin();
        push(x, y);
      }
    } else if (fn === ops.rectangle) {
      const x = coords[index++];
      const y = coords[index++];
      const width = coords[index++];
      const height = coords[index++];
      if ([x, y, width, height].every(Number.isFinite)) {
        begin();
        push(x, y);
        push(x + width, y);
        push(x + width, y + height);
        push(x, y + height);
        const first = current[0];
        current.push({ ...first });
      }
    } else if (fn === ops.closePath) {
      if (current?.length) {
        const first = current[0];
        const last = current[current.length - 1];
        if (Math.abs(first.x - last.x) > 1e-6 || Math.abs(first.y - last.y) > 1e-6) current.push({ ...first });
      }
    } else {
      // 未知路径子操作：保守放弃本路径
      return;
    }
  }
  for (const subpath of subpaths) {
    const rect = polygonToRect(subpath);
    if (rect) output.push(rect);
  }
}

/**
 * 从 operator list 检测矢量矩形（已按 CTM 变换到页面 user space）。
 * @param {number[]} fnArray
 * @param {any[]} argsArray
 * @param {object} ops pdfjs OPS 常量表
 * @returns {Array<{x:number,y:number,width:number,height:number}>}
 */
export function detectRectsFromOperatorList(fnArray, argsArray, ops) {
  const rects = [];
  if (!Array.isArray(fnArray) || !Array.isArray(argsArray) || !ops) return rects;
  let ctm = IDENTITY_MATRIX;
  const stack = [];
  for (let i = 0; i < fnArray.length; i++) {
    const fn = fnArray[i];
    const args = argsArray[i];
    if (fn === ops.save) {
      stack.push(ctm);
    } else if (fn === ops.restore) {
      ctm = stack.pop() ?? IDENTITY_MATRIX;
    } else if (fn === ops.transform) {
      if (Array.isArray(args) && args.length >= 6 && args.slice(0, 6).every((v) => Number.isFinite(Number(v)))) {
        ctm = multiply(ctm, args.slice(0, 6).map(Number));
      }
    } else if (fn === ops.paintFormXObjectBegin) {
      stack.push(ctm);
      if (Array.isArray(args) && Array.isArray(args[0]) && args[0].length >= 6) {
        ctm = multiply(ctm, args[0].slice(0, 6).map(Number));
      }
    } else if (fn === ops.paintFormXObjectEnd) {
      ctm = stack.pop() ?? IDENTITY_MATRIX;
    } else if (fn === ops.constructPath) {
      if (Array.isArray(args) && Array.isArray(args[0])) {
        rectsFromConstructPath(args[0], Array.from(args[1] ?? []), ctm, ops, rects);
      }
    } else if (fn === ops.rectangle) {
      if (Array.isArray(args) && args.length >= 4) {
        rectsFromConstructPath([ops.rectangle], Array.from(args), ctm, ops, rects);
      }
    }
  }
  return rects;
}

/** 面积交并比。 */
function iou(a, b) {
  const left = Math.max(a.x, b.x);
  const right = Math.min(a.x + a.width, b.x + b.width);
  const bottom = Math.max(a.y, b.y);
  const top = Math.min(a.y + a.height, b.y + b.height);
  if (right <= left || top <= bottom) return 0;
  const inter = (right - left) * (top - bottom);
  const union = a.width * a.height + b.width * b.height - inter;
  return union > 0 ? inter / union : 0;
}

function contains(outer, inner) {
  return inner.x >= outer.x - 0.5
    && inner.y >= outer.y - 0.5
    && inner.x + inner.width <= outer.x + outer.width + 0.5
    && inner.y + inner.height <= outer.y + outer.height + 0.5;
}

/**
 * 去重聚类：
 * - IoU > 0.9 视为同一图框（重复绘制）；
 * - 整页边框识别：某候选包含 ≥2 个图框候选、且被包含候选面积合计 ≥ 自身 50%，
 *   判为图纸外框（如 A2 页内排 4 个 A4 图框时的 A2 边框），丢弃外框、保留内部图框；
 * - 其余包含关系（小框在大框内且面积 <50%）判为框内嵌套线，丢弃小框（§3.4）。
 */
export function dedupeFrames(frames) {
  const area = (frame) => frame.width * frame.height;
  const sorted = [...frames].sort((a, b) => area(b) - area(a));
  const isPageBorder = (frame) => {
    const nested = sorted.filter((other) => other !== frame
      && contains(frame, other)
      && area(other) < area(frame) * 0.5);
    if (nested.length < 2) return false;
    const covered = nested.reduce((sum, other) => sum + area(other), 0);
    return covered >= area(frame) * 0.5;
  };
  const kept = [];
  for (const frame of sorted) {
    if (isPageBorder(frame)) continue;
    let drop = false;
    for (const candidate of kept) {
      if (iou(frame, candidate) > 0.9) { drop = true; break; }
      if (contains(candidate, frame) && area(frame) < area(candidate) * 0.5) { drop = true; break; }
    }
    if (!drop) kept.push(frame);
  }
  return kept;
}

/** 检测一页中的图框（尺寸表匹配）。 */
export async function detectFramesOnPage(page, options = {}) {
  const pdfjs = await loadPdfjs();
  const opList = await page.getOperatorList();
  const rawRects = detectRectsFromOperatorList(opList.fnArray, opList.argsArray, pdfjs.OPS);
  const { width: pageWidth, height: pageHeight, userUnit } = pageSize(page);
  const sizes = options.sizes ? normalizeFrameSizes(options.sizes) : undefined;
  const frames = [];
  for (const rect of rawRects) {
    if (rect.width < MIN_FRAME_SIDE_PT || rect.height < MIN_FRAME_SIDE_PT) continue;
    const widthMm = rect.width * MM_PER_PT * userUnit;
    const heightMm = rect.height * MM_PER_PT * userUnit;
    const match = matchFrameSize(widthMm, heightMm, {
      sizes,
      toleranceMm: options.toleranceMm,
      aspectTolerance: options.aspectTolerance,
    });
    if (!match) continue;
    frames.push({
      ...rect,
      pageWidth,
      pageHeight,
      widthMm,
      heightMm,
      frameSize: match.name,
      orientation: match.orientation,
      confidence: 'size',
    });
  }
  return dedupeFrames(frames);
}

/** 页面文本项（user space 坐标，y 向上）。 */
export async function getPageTextItems(page) {
  const content = await page.getTextContent();
  const items = [];
  for (const item of content.items ?? []) {
    if (!item || typeof item.str !== 'string' || !item.str) continue;
    const transform = Array.isArray(item.transform) ? item.transform : null;
    if (!transform || transform.length < 6) continue;
    const height = Math.abs(transform[3]) || Math.hypot(transform[2], transform[3]) || Number(item.height) || 0;
    items.push({
      text: item.str,
      x: Number(transform[4]) || 0,
      y: Number(transform[5]) || 0,
      height,
    });
  }
  return items;
}

/**
 * 从已有 PDF 字节提取首页图号（合并场景的单页图框按整页图签区处理）。
 * @param {Uint8Array} bytes
 * @param {object} [options] { pageIndex, titleBlockZone, drawingNoPatterns }
 * @returns {Promise<string|null>}
 */
export async function extractDrawingNoFromPdfBytes(bytes, options = {}) {
  const pageIndex = Number.isInteger(options.pageIndex) ? options.pageIndex : 0;
  const doc = await openPdfjs(bytes);
  try {
    const page = await doc.getPage(pageIndex + 1);
    try {
      const items = await getPageTextItems(page);
      const { width, height } = pageSize(page);
      if (!(width > 0 && height > 0)) return null;
      const relative = items.map((item) => ({
        text: item.text,
        x: item.x / width,
        y: item.y / height,
        height: item.height / height,
      }));
      return extractDrawingNo(relative, {
        titleBlockZone: options.titleBlockZone,
        drawingNoPatterns: options.drawingNoPatterns,
      });
    } finally {
      try { page.cleanup(); } catch { /* ignore */ }
    }
  } finally {
    try { await doc.destroy(); } catch { /* ignore */ }
  }
}

/**
 * 分割主流程：多图框 PDF → 按图框裁剪 → 每图一页 → 排序合并。
 * @param {object} options
 * @param {string} options.inputPath
 * @param {string} options.outputPath
 * @param {'A4-landscape'|'keep'} [options.paper]
 * @param {'drawingNo'|'filename'} [options.sortMode]
 * @param {object} [options.config] 运行配置（frameSizes/frameToleranceMm/titleBlockZone/drawingNoPatterns）
 * @param {AbortSignal} [options.signal]
 * @param {(patch:object)=>void} [options.onProgress]
 * @param {(level:string,msg:string)=>void} [options.log]
 */
export async function splitPdfFile(options) {
  const startedAt = Date.now();
  const {
    inputPath,
    outputPath,
    paper = 'A4-landscape',
    sortMode = 'drawingNo',
    config = {},
    signal,
    onProgress,
    log,
  } = options;
  if (!inputPath) throw new OpsError('E1023', '缺少输入 PDF 路径');
  if (!outputPath) throw new OpsError('E1014', '缺少输出 PDF 路径');

  let bytes;
  try {
    bytes = await readFile(inputPath);
  } catch (err) {
    throw new OpsError('E1023', `无法读取输入 PDF: ${inputPath}`, String(err?.message ?? err));
  }
  throwIfAborted(signal);

  const doc = await openPdfjs(bytes);
  const frames = [];
  let skippedPages = 0;
  try {
    const pageCount = doc.numPages;
    const pageByteEstimate = Math.ceil(bytes.length / Math.max(1, pageCount));
    for (let pageNumber = 1; pageNumber <= pageCount; pageNumber++) {
      throwIfAborted(signal);
      const page = await doc.getPage(pageNumber);
      try {
        if (pageByteEstimate > MAX_PAGE_BYTES) {
          skippedPages += 1;
          log?.('warn', `第 ${pageNumber} 页平均字节 ${pageByteEstimate} 超过 64MB 上限，跳过（R-08）`);
          continue;
        }
        const pageFrames = await detectFramesOnPage(page, {
          sizes: config.frameSizes,
          toleranceMm: config.frameToleranceMm,
        });
        if (!pageFrames.length) {
          log?.('info', `第 ${pageNumber} 页：未识别到标准图框`);
          continue;
        }
        const textItems = await getPageTextItems(page);
        const { width, height } = pageSize(page);
        for (const frame of pageFrames) {
          throwIfAborted(signal);
          const relative = toFrameRelativeItems(textItems, frame);
          const drawingNo = extractDrawingNo(relative, {
            titleBlockZone: normalizeZone(config.titleBlockZone),
            drawingNoPatterns: config.drawingNoPatterns,
          });
          frames.push({
            ...frame,
            pageIndex: pageNumber - 1,
            pageWidth: width,
            pageHeight: height,
            drawingNo,
            confidence: drawingNo ? 'size+title' : 'size',
          });
        }
        log?.('info', `第 ${pageNumber} 页：识别图框 ${pageFrames.length} 个`);
        onProgress?.({ framesTotal: frames.length, pageIndex: pageNumber, pageCount });
      } finally {
        try { page.cleanup(); } catch { /* ignore */ }
      }
    }
  } finally {
    try { await doc.destroy(); } catch { /* ignore */ }
  }

  throwIfAborted(signal);
  if (!frames.length) {
    throw new OpsError('E1012', '未识别到任何图框', '确认输入为矢量 PDF（含 A0–A4 标准图框矩形）；扫描件/图片 PDF 不支持');
  }
  const withNo = frames.filter((frame) => frame.drawingNo).length;
  if (withNo < frames.length) {
    log?.('warn', `E1013：${frames.length - withNo}/${frames.length} 个图框未提取到图号，排序回退文件名/平面序`);
  }

  const ordered = sortFrames(frames, {
    sortMode,
    columnTolerance: options.columnTolerance,
  });
  onProgress?.({ framesTotal: ordered.length });

  const srcDoc = await loadPdfLib(bytes);
  const outDoc = await PDFDocument.create();
  const useA4 = paper !== 'keep';
  const targetWidth = useA4 ? A4_LANDSCAPE_PT.width : 0;
  const targetHeight = useA4 ? A4_LANDSCAPE_PT.height : 0;
  let checkpoints = 0;
  const results = [];
  try {
    for (let index = 0; index < ordered.length; index++) {
      throwIfAborted(signal);
      const frame = ordered[index];
      const srcPage = srcDoc.getPage(frame.pageIndex);
      let embedded;
      try {
        embedded = await outDoc.embedPage(srcPage, {
          left: frame.x,
          bottom: frame.y,
          right: frame.x + frame.width,
          top: frame.y + frame.height,
        });
      } catch (err) {
        throw new OpsError('E1011', `裁剪图框失败（第 ${index + 1} 个）: ${String(err?.message ?? err)}`);
      }
      const pageWidth = useA4 ? targetWidth : frame.width;
      const pageHeight = useA4 ? targetHeight : frame.height;
      const outPage = outDoc.addPage([pageWidth, pageHeight]);
      outPage.drawRectangle({ x: 0, y: 0, width: pageWidth, height: pageHeight, color: rgb(1, 1, 1) });
      const scale = Math.min(pageWidth / frame.width, pageHeight / frame.height);
      outPage.drawPage(embedded, {
        x: (pageWidth - frame.width * scale) / 2,
        y: (pageHeight - frame.height * scale) / 2,
        xScale: scale,
        yScale: scale,
      });
      results.push({
        seq: index + 1,
        pageIndex: frame.pageIndex,
        drawingNo: frame.drawingNo ?? null,
        frameSize: frame.frameSize,
        widthMm: Math.round(frame.widthMm * 100) / 100,
        heightMm: Math.round(frame.heightMm * 100) / 100,
        crop: {
          x: Math.round(frame.x * 100) / 100,
          y: Math.round(frame.y * 100) / 100,
          width: Math.round(frame.width * 100) / 100,
          height: Math.round(frame.height * 100) / 100,
        },
      });
      onProgress?.({ framesDone: index + 1, framesTotal: ordered.length, currentFile: frame.drawingNo ?? `frame-${index + 1}` });
      if ((index + 1) % CHECKPOINT_EVERY === 0 && index + 1 < ordered.length) {
        await outDoc.save();
        checkpoints += 1;
        log?.('info', `第 ${index + 1} 页检查点：已 save（${checkpoints}）`);
      }
    }

    // 合并/写出阶段（状态由调用方维护；此处仅落盘）
    await mkdir(dirname(outputPath), { recursive: true });
    const outBytes = await outDoc.save();
    try {
      await writeFile(outputPath, outBytes);
    } catch (err) {
      throw new OpsError('E1014', `写出 PDF 失败: ${outputPath}`, String(err?.message ?? err));
    }
    return {
      ok: true,
      result: {
        outputPath,
        pages: results.length,
        frames: results,
        paper,
        sortMode,
        skippedPages,
        checkpoints,
        elapsedMs: Date.now() - startedAt,
      },
    };
  } catch (err) {
    if (isAbortError(err)) throw err;
    if (err instanceof OpsError) throw err;
    throw new OpsError('E1011', `PDF 分割失败: ${String(err?.message ?? err)}`);
  }
}
