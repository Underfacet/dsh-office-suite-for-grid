/**
 * D2：PDF 级分割引擎测试（node:test）
 * 夹具：tests/fixtures/make-multi-frame-pdf.js（A2 页内 4 个 A4 图框 + 图签图号）
 *
 * 断言：帧数 / 裁剪 / 图号 / 排序 / 纸型（§5.2 D2）。
 */
import test, { after, before } from 'node:test';
import assert from 'node:assert/strict';
import { mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

// 依赖安装在 suite/node_modules（npm install 在 suite/ 执行）
const requireFromSuite = createRequire(new URL('../suite/package.json', import.meta.url));
const { PDFDocument, StandardFonts, rgb } = requireFromSuite('pdf-lib');
import {
  A2_HEIGHT_PT,
  A2_WIDTH_PT,
  A4_HEIGHT_PT,
  A4_WIDTH_PT,
  PT_PER_MM,
  writeMultiFramePdf,
} from './fixtures/make-multi-frame-pdf.js';
import {
  detectFramesOnPage,
  getPageTextItems,
  openPdfjs,
  splitPdfFile,
} from '../suite/host/ops/pdf-split.js';
import {
  compareNatural,
  extractDrawingNo,
  matchBlockName,
  matchFrameSize,
  toFrameRelativeItems,
  zOrderFrames,
} from '../suite/host/ops/drawingNo.js';
import { assertInsideRoot, DEFAULT_BOUNDARY_ROOT, resolveBoundaryRoot } from '../suite/host/boundary.js';

const HERE = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = resolve(HERE, '..');
const TMP = join(REPO_ROOT, 'runtime', 'tmp', 'tests', `split-${process.pid}`);

let fixture;

before(async () => {
  await mkdir(TMP, { recursive: true });
  fixture = await writeMultiFramePdf(join(TMP, 'multi-frame-a2.pdf'));
});

after(async () => {
  await rm(TMP, { recursive: true, force: true });
});

/** 逐页读取文本项（含坐标），用于裁剪断言。 */
async function readPageTextItems(bytes) {
  const doc = await openPdfjs(bytes);
  const pages = [];
  try {
    for (let index = 1; index <= doc.numPages; index++) {
      const page = await doc.getPage(index);
      try {
        const content = await page.getTextContent();
        const view = page.view;
        pages.push({
          width: Math.abs(view[2] - view[0]),
          height: Math.abs(view[3] - view[1]),
          items: (content.items ?? [])
            .filter((item) => typeof item.str === 'string' && item.str)
            .map((item) => ({ text: item.str, x: item.transform[4], y: item.transform[5] })),
        });
      } finally {
        page.cleanup();
      }
    }
  } finally {
    await doc.destroy();
  }
  return pages;
}

/** 只保留落在页面可视区域内的文本（Form XObject BBox 之外的邻居内容不计）。 */
function inBoundsText(page) {
  return page.items
    .filter((item) => item.x >= -1 && item.x <= page.width + 1 && item.y >= -1 && item.y <= page.height + 1)
    .map((item) => item.text)
    .join('');
}

async function buildSingleA3FramePdf(outPath) {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const page = doc.addPage([594 * PT_PER_MM, 420 * PT_PER_MM]); // A2 横向
  const frameX = 50;
  const frameY = 50;
  const frameW = 420 * PT_PER_MM; // A3 横向
  const frameH = 297 * PT_PER_MM;
  page.drawRectangle({ x: frameX, y: frameY, width: frameW, height: frameH, borderColor: rgb(0, 0, 0), borderWidth: 1 });
  page.drawText('D0301-01', {
    x: frameX + frameW * 0.66,
    y: frameY + frameH * 0.06,
    size: 12,
    font,
    color: rgb(0, 0, 0),
  });
  const bytes = await doc.save();
  await mkdir(dirname(outPath), { recursive: true });
  await writeFile(outPath, bytes);
  return outPath;
}

test('夹具：检测 4 个 A4 图框并逐框提取图号', async () => {
  const bytes = await readFile(fixture.path);
  const doc = await openPdfjs(bytes);
  try {
    const page = await doc.getPage(1);
    try {
      const frames = await detectFramesOnPage(page);
      assert.equal(frames.length, 4, '应识别 4 个 A4 图框');
      assert.ok(frames.every((frame) => frame.frameSize === 'A4'));
      const items = await getPageTextItems(page);
      const drawingNos = frames
        .map((frame) => extractDrawingNo(toFrameRelativeItems(items, frame), {}))
        .sort();
      assert.deepEqual(drawingNos, ['D0101-01', 'D0101-02', 'D0101-03', 'D0101-04']);
    } finally {
      page.cleanup();
    }
  } finally {
    await doc.destroy();
  }
});

test('整页 A2 外框 + 4 个 A4 图框：丢弃外框、保留 4 个图框', async () => {
  // 页面比 A2 大一圈，图框整体内缩：外框完整包住 4 个 A4 图框（真实 CAD PDF 常见版式）
  const bordered = await writeMultiFramePdf(join(TMP, 'multi-frame-a2-border.pdf'), {
    pageBorder: true,
    offset: 30,
    pageWidth: A2_WIDTH_PT + 60,
    pageHeight: A2_HEIGHT_PT + 60,
  });
  const bytes = await readFile(bordered.path);
  const doc = await openPdfjs(bytes);
  try {
    const page = await doc.getPage(1);
    try {
      const frames = await detectFramesOnPage(page);
      assert.equal(frames.length, 4, '整页 A2 边框应被识别为图纸外框并丢弃');
      assert.ok(frames.every((frame) => frame.frameSize === 'A4'));
    } finally {
      page.cleanup();
    }
  } finally {
    await doc.destroy();
  }
});

test('分割：4 页 / A4 横向 / 图号排序 / 裁剪只含本图框', async () => {
  const outputPath = join(TMP, 'split-a4.pdf');
  const result = await splitPdfFile({
    inputPath: fixture.path,
    outputPath,
    paper: 'A4-landscape',
    sortMode: 'drawingNo',
    config: {},
  });
  assert.equal(result.ok, true);
  assert.equal(result.result.pages, 4);
  assert.deepEqual(
    result.result.frames.map((frame) => frame.drawingNo),
    ['D0101-01', 'D0101-02', 'D0101-03', 'D0101-04'],
    '按图号自然序输出',
  );

  const outputBytes = await readFile(outputPath);
  const outputDoc = await PDFDocument.load(outputBytes);
  assert.equal(outputDoc.getPageCount(), 4);
  for (const page of outputDoc.getPages()) {
    const { width, height } = page.getSize();
    assert.ok(Math.abs(width - A4_WIDTH_PT) < 0.5, `页宽应为 A4 横向: ${width}`);
    assert.ok(Math.abs(height - A4_HEIGHT_PT) < 0.5, `页高应为 A4 横向: ${height}`);
  }

  const pages = await readPageTextItems(outputBytes);
  const expected = ['D0101-01', 'D0101-02', 'D0101-03', 'D0101-04'];
  pages.forEach((page, index) => {
    const text = inBoundsText(page);
    assert.match(text, new RegExp(expected[index]), `第 ${index + 1} 页应含图号 ${expected[index]}`);
    for (const other of expected) {
      if (other === expected[index]) continue;
      assert.ok(!text.includes(other), `第 ${index + 1} 页不应含其它图框图号 ${other}`);
    }
  });
});

test("分割：paper='keep' 保留图框原幅面（A3）", async () => {
  const inputPath = await buildSingleA3FramePdf(join(TMP, 'a3-frame.pdf'));
  const outputPath = join(TMP, 'split-keep-a3.pdf');
  const result = await splitPdfFile({
    inputPath,
    outputPath,
    paper: 'keep',
    sortMode: 'drawingNo',
    config: {},
  });
  assert.equal(result.result.pages, 1);
  const outputDoc = await PDFDocument.load(await readFile(outputPath));
  const [page] = outputDoc.getPages();
  const { width, height } = page.getSize();
  assert.ok(Math.abs(width - 420 * PT_PER_MM) < 0.5, `keep 页宽应为 A3 横向: ${width}`);
  assert.ok(Math.abs(height - 297 * PT_PER_MM) < 0.5, `keep 页高应为 A3 横向: ${height}`);
});

test('分割：无图框 PDF 报 E1012', async () => {
  const doc = await PDFDocument.create();
  doc.addPage([200, 200]).drawRectangle({ x: 10, y: 10, width: 50, height: 30 });
  const inputPath = join(TMP, 'no-frame.pdf');
  await writeFile(inputPath, await doc.save());
  await assert.rejects(
    () => splitPdfFile({ inputPath, outputPath: join(TMP, 'no-frame-out.pdf'), paper: 'A4-landscape', sortMode: 'drawingNo', config: {} }),
    (err) => err.code === 'E1012',
  );
});

test('图号提取：图签区 / 行合并 / 两条默认正则', () => {
  const items = [
    { text: '图号：', x: 0.65, y: 0.08, height: 0.02 },
    { text: 'D0101-01', x: 0.73, y: 0.08, height: 0.02 },
    { text: '比例 1:100', x: 0.65, y: 0.14, height: 0.02 },
    // 图签区外（x < 0.6）不应参与
    { text: 'D9999-99', x: 0.2, y: 0.1, height: 0.02 },
  ];
  assert.equal(extractDrawingNo(items, {}), 'D0101-01');
  // 模式优先级：带“/图号”的行优先（第一条正则）
  const slashItems = [
    { text: '图号：D0101-01', x: 0.65, y: 0.08, height: 0.02 },
    { text: '/图号：D0101-02', x: 0.65, y: 0.16, height: 0.02 },
  ];
  assert.equal(extractDrawingNo(slashItems, {}), 'D0101-02');
  assert.equal(extractDrawingNo([{ text: 'A4', x: 0.7, y: 0.1, height: 0.02 }], {}), null);
});

test('图幅匹配：容差 ±2mm / 宽高比 0.5%', () => {
  assert.equal(matchFrameSize(297, 210)?.name, 'A4');
  assert.equal(matchFrameSize(297.5, 209.6)?.name, 'A4');
  assert.equal(matchFrameSize(841, 594)?.name, 'A1');
  assert.equal(matchFrameSize(1189, 841)?.name, 'A0');
  assert.equal(matchFrameSize(320, 210), null, '宽高比偏差过大应不匹配');
  assert.equal(matchFrameSize(0, 0), null);
});

test('平面序与自然排序', () => {
  const frames = [
    { id: 'bottom-right', drawingNo: 'D0101-01', x: 100, y: 0, width: 50, height: 30 },
    { id: 'top-left', drawingNo: 'D0101-02', x: 0, y: 40, width: 50, height: 30 },
    { id: 'bottom-left', drawingNo: 'D0101-03', x: 0, y: 0, width: 50, height: 30 },
    { id: 'top-right', drawingNo: 'D0101-04', x: 100, y: 40, width: 50, height: 30 },
  ];
  assert.deepEqual(zOrderFrames(frames).map((frame) => frame.id), ['top-left', 'bottom-left', 'top-right', 'bottom-right']);
  assert.ok(compareNatural('图纸-10', '图纸-2') > 0);
  assert.ok(compareNatural('D0101-01', 'D0101-2') < 0);
  assert.equal(matchBlockName('A2-图框-01', ['*图框*']), true);
  assert.equal(matchBlockName('图框', ['A0', 'A1']), false);
});

test('边界：默认根=仓库根上溯两级；越界读写 E1021；allowExternalPaths 只放宽读取', () => {
  const root = resolveBoundaryRoot({});
  assert.equal(root, DEFAULT_BOUNDARY_ROOT);
  // 默认根与机器无关：= tests 目录上溯三级（tests → 仓库根 → 上级 → 上上级）
  const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
  assert.equal(resolve(root), resolve(repoRoot, '..', '..'));
  assert.throws(
    () => assertInsideRoot('C:\\Windows\\notepad.exe', { write: false, root }),
    (err) => err.code === 'E1021',
  );
  assert.equal(
    assertInsideRoot('C:\\Windows\\notepad.exe', { write: false, root, allowExternalPaths: true }),
    'C:\\Windows\\notepad.exe',
  );
  assert.throws(
    () => assertInsideRoot('C:\\Windows\\out.pdf', { write: true, root, allowExternalPaths: true }),
    (err) => err.code === 'E1021',
  );
  assert.equal(assertInsideRoot(join(root, 'x.pdf'), { write: true, root }), join(root, 'x.pdf'));
  // 自定义 projectRoot 覆盖
  assert.equal(resolveBoundaryRoot({ projectRoot: root }), root);
});
