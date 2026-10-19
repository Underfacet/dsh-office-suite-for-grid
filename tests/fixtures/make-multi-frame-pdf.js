/**
 * D1 夹具：合成「A2 页面内排 4 个 A4 图框 + 图签区图号文本」的多图框 PDF。
 *
 * 版式（A2 横向 = 297×210 mm 的 2×2，正好铺满）：
 *   ┌──────────────┬──────────────┐
 *   │ 02 (左上)    │ 04 (右上)    │
 *   ├──────────────┼──────────────┤
 *   │ 03 (左下)    │ 01 (右下)    │
 *   └──────────────┴──────────────┘
 * Z 形平面序 = 02 → 03 → 04 → 01；按图号排序 = 01 → 02 → 03 → 04（顺序发生变化，便于断言）。
 *
 * 用法：
 *   node tests/fixtures/make-multi-frame-pdf.js [输出路径]
 *   默认写出 runtime/tmp/fixtures/multi-frame-a2.pdf（runtime/ 已被 git 忽略）。
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
export const A4_WIDTH_PT = 297 * PT_PER_MM;   // 841.8898
export const A4_HEIGHT_PT = 210 * PT_PER_MM;  // 595.2756
export const A2_WIDTH_PT = A4_WIDTH_PT * 2;   // 1683.7795
export const A2_HEIGHT_PT = A4_HEIGHT_PT * 2; // 1190.5512

/** 4 个图框的版式（col/row 从 0 开始；row=0 为下排）。 */
export const FRAME_LAYOUT = Object.freeze([
  { drawingNo: 'D0101-03', col: 0, row: 0 },
  { drawingNo: 'D0101-01', col: 1, row: 0 },
  { drawingNo: 'D0101-02', col: 0, row: 1 },
  { drawingNo: 'D0101-04', col: 1, row: 1 },
]);

/**
 * 生成多图框 PDF。
 * @param {{pageWidth?:number,pageHeight?:number,fontSize?:number}} [options]
 * @returns {Promise<{bytes:Uint8Array,frames:Array<{drawingNo:string,col:number,row:number,x:number,y:number,width:number,height:number}>,pageWidth:number,pageHeight:number}>}
 */
export async function buildMultiFramePdf(options = {}) {
  const pageWidth = options.pageWidth ?? A2_WIDTH_PT;
  const pageHeight = options.pageHeight ?? A2_HEIGHT_PT;
  const fontSize = options.fontSize ?? 12;
  const offset = options.offset ?? 0;
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const page = doc.addPage([pageWidth, pageHeight]);
  if (options.pageBorder) {
    // 图纸外框（A2 幅面），用于验证「整页边框 + 内部多图框」的去重逻辑
    page.drawRectangle({
      x: 1,
      y: 1,
      width: pageWidth - 2,
      height: pageHeight - 2,
      borderColor: rgb(0, 0, 0),
      borderWidth: 1,
    });
  }
  const frames = [];
  for (const layout of FRAME_LAYOUT) {
    const x = offset + layout.col * A4_WIDTH_PT;
    const y = offset + layout.row * A4_HEIGHT_PT;
    page.drawRectangle({
      x,
      y,
      width: A4_WIDTH_PT,
      height: A4_HEIGHT_PT,
      borderColor: rgb(0, 0, 0),
      borderWidth: 1,
    });
    // 图签区（图框右下）内绘制图号文本（标准字体仅支持 ASCII，故用裸编号，命中第二条默认正则）
    page.drawText(layout.drawingNo, {
      x: x + A4_WIDTH_PT * 0.66,
      y: y + A4_HEIGHT_PT * 0.06,
      size: fontSize,
      font,
      color: rgb(0, 0, 0),
    });
    page.drawText('SCALE 1:100', {
      x: x + A4_WIDTH_PT * 0.66,
      y: y + A4_HEIGHT_PT * 0.11,
      size: 8,
      font,
      color: rgb(0, 0, 0),
    });
    frames.push({
      drawingNo: layout.drawingNo,
      col: layout.col,
      row: layout.row,
      x,
      y,
      width: A4_WIDTH_PT,
      height: A4_HEIGHT_PT,
    });
  }
  const bytes = await doc.save();
  return { bytes, frames, pageWidth, pageHeight };
}

/** 生成并写出多图框 PDF，返回 { path, ...buildMultiFramePdf }。 */
export async function writeMultiFramePdf(outPath, options = {}) {
  const target = outPath ?? join(REPO_ROOT, 'runtime', 'tmp', 'fixtures', 'multi-frame-a2.pdf');
  const built = await buildMultiFramePdf(options);
  await mkdir(dirname(target), { recursive: true });
  await writeFile(target, built.bytes);
  return { path: target, ...built };
}

const invokedDirectly = process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url;
if (invokedDirectly) {
  const target = process.argv[2] ? resolve(process.argv[2]) : undefined;
  const fixture = await writeMultiFramePdf(target);
  console.log(JSON.stringify({
    ok: true,
    path: fixture.path,
    pageWidth: fixture.pageWidth,
    pageHeight: fixture.pageHeight,
    frames: fixture.frames.map((frame) => frame.drawingNo),
  }, null, 2));
}
