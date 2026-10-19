/**
 * D1 夹具：合成「3 个单页 A4 纵向 PDF」的合并集（文件名与图号顺序故意不同）。
 *
 *   图纸-01.pdf → D0201-03
 *   图纸-02.pdf → D0201-01
 *   图纸-03.pdf → D0201-02
 *
 * 因此：sort=filename → 01,02,03；sort=drawingNo → 02,03,01（按图号 01,02,03）。
 *
 * 用法：node tests/fixtures/make-merge-set.js [输出目录]
 *   默认写出 runtime/tmp/fixtures/merge-set/。
 */
import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

// 依赖安装在 suite/node_modules（npm install 在 suite/ 执行）；测试夹具从包位置解析
const requireFromSuite = createRequire(new URL('../../suite/package.json', import.meta.url));
const { PDFDocument, StandardFonts, rgb } = requireFromSuite('pdf-lib');

const HERE = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = resolve(HERE, '..', '..');

export const PT_PER_MM = 72 / 25.4;
export const A4_PORTRAIT = Object.freeze({ width: 210 * PT_PER_MM, height: 297 * PT_PER_MM });

/** 合并集定义：文件名顺序与图号顺序相反。 */
export const MERGE_SET = Object.freeze([
  { fileName: '图纸-01.pdf', drawingNo: 'D0201-03', label: 'PLAN-A' },
  { fileName: '图纸-02.pdf', drawingNo: 'D0201-01', label: 'PLAN-B' },
  { fileName: '图纸-03.pdf', drawingNo: 'D0201-02', label: 'PLAN-C' },
]);

async function buildSinglePagePdf(entry, options = {}) {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const width = options.width ?? A4_PORTRAIT.width;
  const height = options.height ?? A4_PORTRAIT.height;
  const page = doc.addPage([width, height]);
  page.drawRectangle({ x: 0, y: 0, width, height, borderColor: rgb(0, 0, 0), borderWidth: 1 });
  // 图签区（相对整页右下）：x ∈ [0.6,1]、y ∈ [0,0.25]
  page.drawText(entry.drawingNo, {
    x: width * 0.66,
    y: height * 0.06,
    size: 12,
    font,
    color: rgb(0, 0, 0),
  });
  page.drawText(entry.label, {
    x: width * 0.66,
    y: height * 0.11,
    size: 9,
    font,
    color: rgb(0, 0, 0),
  });
  return doc.save();
}

/** 生成并写出合并集，返回 [{ path, fileName, drawingNo }]。 */
export async function writeMergeSet(outDir) {
  const targetDir = outDir ?? join(REPO_ROOT, 'runtime', 'tmp', 'fixtures', 'merge-set');
  await mkdir(dirname(join(targetDir, 'x.pdf')), { recursive: true });
  const files = [];
  for (const entry of MERGE_SET) {
    const bytes = await buildSinglePagePdf(entry);
    const path = join(targetDir, entry.fileName);
    await writeFile(path, bytes);
    files.push({ path, fileName: entry.fileName, drawingNo: entry.drawingNo });
  }
  return files;
}

const invokedDirectly = process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url;
if (invokedDirectly) {
  const dir = process.argv[2] ? resolve(process.argv[2]) : undefined;
  const files = await writeMergeSet(dir);
  console.log(JSON.stringify({ ok: true, files }, null, 2));
}
