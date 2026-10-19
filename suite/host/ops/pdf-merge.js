/**
 * PDF 合并 + 纸型适配 + 排序（docs/P1-设计.md §3.2 / §3.5）
 *
 * - 纸型适配 adaptPaper：统一 A4 横向（缩放 fit + 居中 + 白底）；paper:'keep' 原样保留；
 * - 排序：调用 drawingNo.js 的 sortFiles（图号优先，回退文件名自然序）；
 *   图号可由调用方提供，或从 PDF 首页图签区文本提取（pdf-split.js）。
 * - 合并：pdf-lib 逐页嵌入/复制，输出到 outputPath；进度每文件一次。
 */
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { basename, dirname } from 'node:path';
import { PDFDocument, rgb } from 'pdf-lib';
import { OpsError, isAbortError, throwIfAborted } from '../ops.js';
import { A4_LANDSCAPE_PT, sortFiles } from './drawingNo.js';
import { extractDrawingNoFromPdfBytes, loadPdfLib } from './pdf-split.js';

/** 归一化合并输入：字符串 → { path, fileName }。 */
export function normalizeMergeEntries(files) {
  const list = Array.isArray(files) ? files : [];
  return list.map((entry, index) => {
    if (typeof entry === 'string') {
      return { path: entry, fileName: basename(entry), drawingNo: null, order: index };
    }
    const path = String(entry?.path ?? '');
    return {
      path,
      fileName: String(entry?.fileName ?? entry?.name ?? (path ? basename(path) : '')),
      drawingNo: entry?.drawingNo == null || entry?.drawingNo === '' ? null : String(entry.drawingNo),
      order: index,
    };
  }).filter((entry) => entry.path);
}

/** 有界并发 map（保持结果顺序）。 */
async function mapLimit(items, limit, fn) {
  const results = new Array(items.length);
  if (!items.length) return results;
  const width = Math.max(1, Math.min(limit, items.length));
  let cursor = 0;
  const workers = Array.from({ length: width }, async () => {
    while (true) {
      const index = cursor++;
      if (index >= items.length) return;
      results[index] = await fn(items[index], index);
    }
  });
  await Promise.all(workers);
  return results;
}

/**
 * 合并 PDF 文件为单个 PDF。
 * @param {object} options
 * @param {Array<string|{path:string,fileName?:string,drawingNo?:string}>} options.files
 * @param {string} options.outputPath
 * @param {'A4-landscape'|'keep'} [options.paper]
 * @param {'drawingNo'|'filename'|'given'} [options.sortMode]
 * @param {object} [options.config]
 * @param {AbortSignal} [options.signal]
 * @param {(patch:object)=>void} [options.onProgress]
 * @param {(level:string,msg:string)=>void} [options.log]
 */
export async function mergePdfs(options) {
  const startedAt = Date.now();
  const {
    files,
    outputPath,
    paper = 'A4-landscape',
    sortMode = 'drawingNo',
    config = {},
    signal,
    onProgress,
    log,
  } = options ?? {};
  if (!outputPath) throw new OpsError('E1014', '缺少输出 PDF 路径');
  const entries = normalizeMergeEntries(files);
  if (!entries.length) {
    throw new OpsError('E1023', '没有可合并的 PDF 文件', '请检查输入目录中是否存在 .pdf 文件');
  }
  const effectiveSort = sortMode === 'given' ? 'given' : sortMode === 'filename' ? 'filename' : 'drawingNo';

  // 第一遍：读取 + 图号提取（图号优先排序需要；失败时告警并回退文件名）
  if (effectiveSort === 'drawingNo') {
    const concurrency = Number.isInteger(options?.concurrency) && options.concurrency > 0
      ? Math.min(options.concurrency, 8)
      : 1;
    const pending = entries.filter((entry) => !entry.drawingNo);
    const extract = async (entry) => {
      throwIfAborted(signal);
      try {
        const bytes = await readFile(entry.path);
        entry.drawingNo = await extractDrawingNoFromPdfBytes(bytes, {
          titleBlockZone: config.titleBlockZone,
          drawingNoPatterns: config.drawingNoPatterns,
        });
      } catch (err) {
        if (isAbortError(err)) throw err;
        entry.drawingNo = null;
        log?.('warn', `图号提取失败（回退文件名排序）: ${entry.fileName} — ${String(err?.message ?? err)}`);
      }
    };
    if (concurrency > 1) {
      await mapLimit(pending, concurrency, extract);
    } else {
      for (const entry of pending) await extract(entry);
    }
  }

  const ordered = effectiveSort === 'given' ? entries : sortFiles(entries, effectiveSort);
  const missingNo = ordered.filter((entry) => !entry.drawingNo).length;
  if (effectiveSort === 'drawingNo' && missingNo) {
    log?.('warn', `E1013：${missingNo}/${ordered.length} 个文件未提取到图号，排序回退文件名`);
  }

  // 第二遍：逐文件合并
  const outDoc = await PDFDocument.create();
  const useA4 = paper !== 'keep';
  let pageCount = 0;
  const mergedFiles = [];
  try {
    for (let index = 0; index < ordered.length; index++) {
      throwIfAborted(signal);
      const entry = ordered[index];
      let bytes;
      try {
        bytes = await readFile(entry.path);
      } catch (err) {
        throw new OpsError('E1023', `无法读取 PDF: ${entry.path}`, String(err?.message ?? err));
      }
      const srcDoc = await loadPdfLib(bytes);
      if (useA4) {
        const sourcePages = srcDoc.getPages();
        for (const srcPage of sourcePages) {
          throwIfAborted(signal);
          const { width, height } = srcPage.getSize();
          if (!(width > 0 && height > 0)) continue;
          const embedded = await outDoc.embedPage(srcPage);
          const targetWidth = A4_LANDSCAPE_PT.width;
          const targetHeight = A4_LANDSCAPE_PT.height;
          const outPage = outDoc.addPage([targetWidth, targetHeight]);
          outPage.drawRectangle({ x: 0, y: 0, width: targetWidth, height: targetHeight, color: rgb(1, 1, 1) });
          const scale = Math.min(targetWidth / width, targetHeight / height);
          outPage.drawPage(embedded, {
            x: (targetWidth - width * scale) / 2,
            y: (targetHeight - height * scale) / 2,
            xScale: scale,
            yScale: scale,
          });
          pageCount += 1;
        }
      } else {
        const copied = await outDoc.copyPages(srcDoc, srcDoc.getPageIndices());
        for (const page of copied) outDoc.addPage(page);
        pageCount += copied.length;
      }
      mergedFiles.push({ fileName: entry.fileName, drawingNo: entry.drawingNo ?? null, path: entry.path });
      onProgress?.({
        filesDone: index + 1,
        filesTotal: ordered.length,
        currentFile: entry.fileName,
      });
      log?.('info', `合并 ${index + 1}/${ordered.length}: ${entry.fileName}${entry.drawingNo ? `（图号 ${entry.drawingNo}）` : ''}`);
    }

    await mkdir(dirname(outputPath), { recursive: true });
    const outBytes = await outDoc.save();
    try {
      await writeFile(outputPath, outBytes);
    } catch (err) {
      throw new OpsError('E1014', `写出合并 PDF 失败: ${outputPath}`, String(err?.message ?? err));
    }
    return {
      ok: true,
      result: {
        outputPath,
        pages: pageCount,
        files: mergedFiles,
        paper,
        sortMode: effectiveSort,
        elapsedMs: Date.now() - startedAt,
      },
    };
  } catch (err) {
    if (isAbortError(err) || err instanceof OpsError) throw err;
    throw new OpsError('E1014', `合并 PDF 失败: ${String(err?.message ?? err)}`);
  }
}
