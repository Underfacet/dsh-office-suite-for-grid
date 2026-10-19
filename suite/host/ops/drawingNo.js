/**
 * 图幅 / 图框 / 图号 / 排序算法（docs/P1-设计.md §3.4）
 *
 * 本模块是纯函数（无 fs / 无 ctx / 无第三方依赖），COM 引擎（cad-convert.ps1 的参数）、
 * PDF 级引擎（pdf-split.js / pdf-merge.js）与单元测试共用同一套规则：
 * - FRAME_SIZES：标准图幅尺寸表（mm，长边×短边），config.frameSizes 可增补覆盖；
 * - matchFrameSize：尺寸 + 宽高比双重匹配（容差 ±2mm、宽高比 0.5%）；
 * - matchBlockName：图框块名通配匹配（`*图框*` / `A0` …）；
 * - extractDrawingNo：图签区相对坐标 + 行合并 + 两条默认正则；
 * - compareNatural / sortFiles / sortFrames / zOrderFrames：自然排序与 Z 形平面序。
 */

/** 1pt = 25.4/72 mm；1mm = 72/25.4 pt。 */
export const MM_PER_PT = 25.4 / 72;
export const PT_PER_MM = 72 / 25.4;

/** A4 横向（PDF 输出纸型）尺寸（pt）。 */
export const A4_LANDSCAPE_PT = Object.freeze({ width: 297 * PT_PER_MM, height: 210 * PT_PER_MM });

/**
 * 标准图幅尺寸表（mm，长边×短边）。§3.4 表格共 16 项。
 * name 使用 ASCII 便于日志/文件名，label 保留排版形式。
 */
const RAW_FRAME_SIZES = [
  { name: 'A4', width: 297, height: 210 },
  { name: 'A3', width: 420, height: 297 },
  { name: 'A2', width: 594, height: 420 },
  { name: 'A1', width: 841, height: 594 },
  { name: 'A0', width: 1189, height: 841 },
  { name: 'A4x3', width: 891, height: 210, label: 'A4×3' },
  { name: 'A4x4', width: 1189, height: 210, label: 'A4×4' },
  { name: 'A3x3', width: 891, height: 420, label: 'A3×3' },
  { name: 'A3x4', width: 1189, height: 420, label: 'A3×4' },
  { name: 'A2x3', width: 1261, height: 594, label: 'A2×3' },
  { name: 'A2x4', width: 1682, height: 594, label: 'A2×4' },
  { name: 'A2x5', width: 2102, height: 594, label: 'A2×5' },
  { name: 'A1x3', width: 1783, height: 841, label: 'A1×3' },
  { name: 'A1x4', width: 2378, height: 841, label: 'A1×4' },
  { name: 'A0x2', width: 1682, height: 1189, label: 'A0×2' },
  { name: 'A0x3', width: 2523, height: 1189, label: 'A0×3' },
];

export const FRAME_SIZES = Object.freeze(RAW_FRAME_SIZES.map((size) => Object.freeze({
  ...size,
  long: Math.max(size.width, size.height),
  short: Math.min(size.width, size.height),
})));

/** 图框块名默认通配清单（COM 路径用）。 */
export const DEFAULT_FRAME_BLOCK_NAMES = Object.freeze(['*图框*', 'A0', 'A1', 'A2', 'A3', 'A4', '*TK*']);

/** 图签区默认相对坐标 [x0, y0, x1, y1]（相对图框，原点左下，y 向上）。 */
export const DEFAULT_TITLE_BLOCK_ZONE = Object.freeze([0.6, 0, 1.0, 0.25]);

/** 图号提取默认正则（源字符串，可在 config.drawingNoPatterns 覆盖）。 */
export const DEFAULT_DRAWING_NO_PATTERNS = Object.freeze([
  '[/／](?:图号|编号|图别)[:：]?\\s*([A-Za-z0-9\\u4e00-\\u9fa5_./-]+)',
  '\\b([A-Z]{0,4}-?\\d{2,5}(?:-\\d{1,3})?(?:[A-Z]?\\d{0,3})?)\\b',
]);

/** 默认图幅匹配容差（mm）与宽高比容差。 */
export const DEFAULT_FRAME_TOLERANCE_MM = 2;
export const DEFAULT_ASPECT_TOLERANCE = 0.005;

/**
 * 归一化图幅尺寸表：接受 [{name,width,height}]、{name:[w,h]}、{name:{width,height}}、[[name,w,h]]。
 * 非法条目忽略；空输入 → 默认表。
 */
export function normalizeFrameSizes(input) {
  const list = [];
  const push = (name, width, height) => {
    const w = Number(width);
    const h = Number(height);
    if (!name || !Number.isFinite(w) || !Number.isFinite(h) || w <= 0 || h <= 0) return;
    list.push({ name: String(name), width: w, height: h, long: Math.max(w, h), short: Math.min(w, h) });
  };
  if (Array.isArray(input)) {
    input.forEach((entry, index) => {
      if (Array.isArray(entry)) {
        if (entry.length >= 3) push(entry[0], entry[1], entry[2]);
        else if (entry.length === 2) push(`custom${index + 1}`, entry[0], entry[1]);
        return;
      }
      if (entry && typeof entry === 'object') push(entry.name ?? `custom${index + 1}`, entry.width, entry.height);
    });
  } else if (input && typeof input === 'object') {
    for (const [name, value] of Object.entries(input)) {
      if (Array.isArray(value)) push(name, value[0], value[1]);
      else if (value && typeof value === 'object') push(name, value.width, value.height);
    }
  }
  return list.length ? list : FRAME_SIZES.map((size) => ({ ...size }));
}

/**
 * 图幅匹配：矩形宽高（mm）→ 命中的标准幅面。
 * 规则：长/短边分别容差 ±toleranceMm，且宽高比相对误差 ≤ aspectTolerance。
 * @returns {null | {name:string,width:number,height:number,long:number,short:number,orientation:'landscape'|'portrait',error:number,aspectError:number}}
 */
export function matchFrameSize(widthMm, heightMm, options = {}) {
  const w = Number(widthMm);
  const h = Number(heightMm);
  if (!Number.isFinite(w) || !Number.isFinite(h) || w <= 0 || h <= 0) return null;
  const sizes = options.sizes ? normalizeFrameSizes(options.sizes) : FRAME_SIZES;
  const tolerance = Number.isFinite(options.toleranceMm) ? Number(options.toleranceMm) : DEFAULT_FRAME_TOLERANCE_MM;
  const aspectTolerance = Number.isFinite(options.aspectTolerance) ? Number(options.aspectTolerance) : DEFAULT_ASPECT_TOLERANCE;
  const long = Math.max(w, h);
  const short = Math.min(w, h);
  const actualAspect = long / short;
  let best = null;
  for (const size of sizes) {
    const sLong = size.long ?? Math.max(size.width, size.height);
    const sShort = size.short ?? Math.min(size.width, size.height);
    if (!(sLong > 0 && sShort > 0)) continue;
    if (Math.abs(long - sLong) > tolerance || Math.abs(short - sShort) > tolerance) continue;
    const expectedAspect = sLong / sShort;
    const aspectError = Math.abs(actualAspect - expectedAspect) / expectedAspect;
    if (aspectError > aspectTolerance) continue;
    const error = Math.abs(long - sLong) + Math.abs(short - sShort);
    if (!best || error < best.error) {
      best = {
        name: size.name,
        label: size.label ?? size.name,
        width: size.width,
        height: size.height,
        long: sLong,
        short: sShort,
        orientation: w >= h ? 'landscape' : 'portrait',
        error,
        aspectError,
      };
    }
  }
  return best;
}

/** 通配符 → 正则（`*` 任意长度，`?` 单字符；大小写不敏感）。 */
export function globToRegExp(pattern) {
  const source = String(pattern ?? '');
  const escaped = source.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
  return new RegExp(`^${escaped}$`, 'i');
}

/** 块名是否命中任一图框通配模式。 */
export function matchBlockName(name, patterns = DEFAULT_FRAME_BLOCK_NAMES) {
  if (name == null) return false;
  const target = String(name);
  const list = Array.isArray(patterns) && patterns.length ? patterns : DEFAULT_FRAME_BLOCK_NAMES;
  for (const pattern of list) {
    try {
      if (globToRegExp(pattern).test(target)) return true;
    } catch { /* 非法通配模式忽略 */ }
  }
  return false;
}

/** 编译图号正则（字符串 → RegExp，忽略非法项；空输入用默认）。 */
export function compileDrawingNoPatterns(patterns) {
  const list = Array.isArray(patterns) && patterns.length ? patterns : DEFAULT_DRAWING_NO_PATTERNS;
  const compiled = [];
  for (const item of list) {
    if (item instanceof RegExp) {
      compiled.push(item);
      continue;
    }
    try {
      compiled.push(new RegExp(String(item), 'i'));
    } catch { /* 非法正则忽略 */ }
  }
  return compiled;
}

/** 清洗图号：去首尾空白/冒号/标点，去内部空白。 */
export function cleanDrawingNo(raw) {
  if (raw == null) return null;
  let text = String(raw).trim();
  text = text.replace(/^[\s:：\-–—_]+/, '');
  text = text.replace(/[\s:：;；,，。.、]+$/, '');
  text = text.replace(/\s+/g, '');
  return text.length ? text : null;
}

/**
 * 文本行合并：y 差 < lineToleranceRatio × 行高视为同一行；行内按 x 升序拼接。
 * @param {Array<{text:string,x:number,y:number,height?:number}>} items 相对坐标（0..1）文本项
 * @returns {Array<{text:string,y:number,height:number,items:Array}>}
 */
export function mergeTextLines(items, options = {}) {
  const ratio = Number.isFinite(options.lineToleranceRatio) ? Number(options.lineToleranceRatio) : 0.6;
  const input = Array.isArray(items) ? items.filter((item) => item && String(item.text ?? '').trim()) : [];
  const sorted = [...input].sort((a, b) => (Number(b.y) || 0) - (Number(a.y) || 0));
  const lines = [];
  for (const item of sorted) {
    const y = Number(item.y) || 0;
    const height = Math.max(Number(item.height) || 0, 1e-9);
    let line = null;
    for (const candidate of lines) {
      const tolerance = ratio * Math.max(candidate.height, height);
      if (Math.abs(candidate.y - y) <= tolerance) {
        line = candidate;
        break;
      }
    }
    if (!line) {
      line = { y, height, items: [] };
      lines.push(line);
    } else {
      line.height = Math.max(line.height, height);
    }
    line.items.push(item);
  }
  for (const line of lines) {
    line.items.sort((a, b) => (Number(a.x) || 0) - (Number(b.x) || 0));
    line.text = line.items.map((item) => String(item.text)).join('');
  }
  // 行按 y 降序（上 → 下），保证结果稳定
  lines.sort((a, b) => b.y - a.y);
  return lines;
}

/** 归一化图签区（容忍 x1<x0 / y1<y0 的写法）。 */
export function normalizeZone(zone = DEFAULT_TITLE_BLOCK_ZONE) {
  if (!Array.isArray(zone) || zone.length < 4) return [...DEFAULT_TITLE_BLOCK_ZONE];
  const [a, b, c, d] = zone.map((v) => Number(v));
  if (![a, b, c, d].every(Number.isFinite)) return [...DEFAULT_TITLE_BLOCK_ZONE];
  return [Math.min(a, c), Math.min(b, d), Math.max(a, c), Math.max(b, d)];
}

/**
 * 图号提取（§3.4 第 3–4 步）。
 * @param {Array<{text:string,x:number,y:number,height?:number}>} items 相对图框坐标（0..1，原点左下）
 * @param {object} options { titleBlockZone, drawingNoPatterns }
 * @returns {string|null}
 */
export function extractDrawingNo(items, options = {}) {
  const zone = normalizeZone(options.titleBlockZone);
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
  // 图签区行序：自下而上（图号通常位于图签最底行，比例/日期等在其上方）。
  const lines = mergeTextLines(selected, options).sort((a, b) => a.y - b.y);
  const patterns = compileDrawingNoPatterns(options.drawingNoPatterns);
  // 模式优先：先按 config.drawingNoPatterns 顺序，再按行（图签底部 → 顶部）扫描；取首个命中。
  for (const pattern of patterns) {
    for (const line of lines) {
      const match = line.text.match(pattern);
      if (match) {
        const value = cleanDrawingNo(match[1] ?? match[0]);
        if (value) return value;
      }
    }
  }
  return null;
}

/**
 * 提取文本项在某个图框内的相对坐标（供 extractDrawingNo 使用）。
 * @param {Array<{text,x,y,height}>} items 页面坐标（pt 或 mm，y 向上）
 * @param {{x:number,y:number,width:number,height:number}} frame
 */
export function toFrameRelativeItems(items, frame) {
  const fw = Number(frame.width);
  const fh = Number(frame.height);
  if (!(fw > 0 && fh > 0)) return [];
  return (Array.isArray(items) ? items : []).map((item) => ({
    text: String(item.text ?? ''),
    x: (Number(item.x) - Number(frame.x)) / fw,
    y: (Number(item.y) - Number(frame.y)) / fh,
    height: (Number(item.height) || 0) / fh,
  }));
}

/** 自然排序比较（localeCompare numeric:true）。 */
export function compareNatural(a, b) {
  return String(a ?? '').localeCompare(String(b ?? ''), undefined, { numeric: true, sensitivity: 'base' });
}

/** 文件记录统一取字段。 */
export function fileNameOf(file) {
  if (typeof file === 'string') return file;
  return String(file?.fileName ?? file?.name ?? file?.path ?? '');
}

export function drawingNoOf(file) {
  if (typeof file === 'string') return null;
  const value = file?.drawingNo;
  return value == null || value === '' ? null : String(value);
}

/** 图号优先排序：有图号者在前（按图号自然序），无图号者按文件名自然序追加在后。 */
export function sortFilesByDrawingNo(files) {
  const list = Array.isArray(files) ? [...files] : [];
  return list.sort((a, b) => {
    const an = drawingNoOf(a);
    const bn = drawingNoOf(b);
    if (an && !bn) return -1;
    if (!an && bn) return 1;
    if (an && bn) {
      const byNo = compareNatural(an, bn);
      if (byNo) return byNo;
    }
    return compareNatural(fileNameOf(a), fileNameOf(b));
  });
}

/** 文件名自然排序。 */
export function sortFilesByFilename(files) {
  const list = Array.isArray(files) ? [...files] : [];
  return list.sort((a, b) => compareNatural(fileNameOf(a), fileNameOf(b)));
}

/** 按排序模式排序（'drawingNo' | 'filename'）。 */
export function sortFiles(files, sortMode = 'drawingNo') {
  return sortMode === 'filename' ? sortFilesByFilename(files) : sortFilesByDrawingNo(files);
}

/**
 * Z 形平面序：按 x 分列（容差内同列），列内 y 降序，列间 x 升序。
 * @param {Array<{x:number,y:number,width:number}>} frames
 */
export function zOrderFrames(frames, options = {}) {
  const list = Array.isArray(frames) ? frames.filter(Boolean) : [];
  if (!list.length) return [];
  const widths = list.map((frame) => Number(frame.width) || 0).filter((value) => value > 0).sort((a, b) => a - b);
  const medianWidth = widths.length ? widths[Math.floor(widths.length / 2)] : 0;
  const tolerance = Number.isFinite(options.columnTolerance)
    ? Number(options.columnTolerance)
    : Math.max(1e-6, medianWidth * 0.5);
  const byX = [...list].sort((a, b) => (Number(a.x) || 0) - (Number(b.x) || 0));
  const columns = [];
  for (const frame of byX) {
    const x = Number(frame.x) || 0;
    let column = columns.find((candidate) => Math.abs(candidate.x - x) <= tolerance);
    if (!column) {
      column = { x, frames: [] };
      columns.push(column);
    }
    column.frames.push(frame);
  }
  columns.sort((a, b) => a.x - b.x);
  const ordered = [];
  columns.forEach((column, index) => {
    column.frames.sort((a, b) => (Number(b.y) || 0) - (Number(a.y) || 0));
    for (const frame of column.frames) ordered.push({ ...frame, columnIndex: index });
  });
  return ordered;
}

/**
 * 图框排序：先 Z 形平面序，再按 sortMode 稳定排序（§3.4 排序）。
 * 图号缺失时回退文件名，仍相同则保留平面序（Array.sort 稳定）。
 */
export function sortFrames(frames, options = {}) {
  const ordered = zOrderFrames(frames, options);
  const indexed = ordered.map((frame, index) => ({ frame, index }));
  indexed.sort((a, b) => {
    if (options.sortMode === 'filename') {
      return compareNatural(fileNameOf(a.frame), fileNameOf(b.frame)) || a.index - b.index;
    }
    const an = drawingNoOf(a.frame);
    const bn = drawingNoOf(b.frame);
    if (an && !bn) return -1;
    if (!an && bn) return 1;
    if (an && bn) {
      const byNo = compareNatural(an, bn);
      if (byNo) return byNo;
    }
    return compareNatural(fileNameOf(a.frame), fileNameOf(b.frame)) || a.index - b.index;
  });
  return indexed.map((entry) => entry.frame);
}
