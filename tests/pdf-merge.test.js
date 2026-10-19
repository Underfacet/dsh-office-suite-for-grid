/**
 * D2：PDF 合并引擎测试（node:test）
 * 夹具：tests/fixtures/make-merge-set.js（3 个单页 A4 纵向 PDF，文件名序与图号序相反）
 *
 * 断言：页序 / 纸型 / 图号排序 / 文件名排序（§5.2 D2）。
 */
import test, { after, before } from 'node:test';
import assert from 'node:assert/strict';
import { mkdir, readFile, rm } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

// 依赖安装在 suite/node_modules（npm install 在 suite/ 执行）
const requireFromSuite = createRequire(new URL('../suite/package.json', import.meta.url));
const { PDFDocument } = requireFromSuite('pdf-lib');
import { A4_PORTRAIT, writeMergeSet } from './fixtures/make-merge-set.js';
import { A4_HEIGHT_PT, A4_WIDTH_PT } from './fixtures/make-multi-frame-pdf.js';
import { mergePdfs } from '../suite/host/ops/pdf-merge.js';
import { openPdfjs } from '../suite/host/ops/pdf-split.js';

const HERE = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = resolve(HERE, '..');
const TMP = join(REPO_ROOT, 'runtime', 'tmp', 'tests', `merge-${process.pid}`);

let mergeSet;

before(async () => {
  await mkdir(TMP, { recursive: true });
  mergeSet = await writeMergeSet(join(TMP, 'merge-set'));
});

after(async () => {
  await rm(TMP, { recursive: true, force: true });
});

async function readPageTexts(bytes) {
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
          texts: (content.items ?? [])
            .filter((item) => typeof item.str === 'string' && item.str)
            .filter((item) => item.transform[4] >= -1 && item.transform[4] <= Math.abs(view[2] - view[0]) + 1
              && item.transform[5] >= -1 && item.transform[5] <= Math.abs(view[3] - view[1]) + 1)
            .map((item) => item.str)
            .join(''),
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

test('合并：图号排序 + A4 横向纸型适配', async () => {
  const outputPath = join(TMP, 'merged-drawingno.pdf');
  const result = await mergePdfs({
    files: mergeSet,
    outputPath,
    paper: 'A4-landscape',
    sortMode: 'drawingNo',
    config: {},
  });
  assert.equal(result.ok, true);
  assert.equal(result.result.pages, 3);
  assert.deepEqual(
    result.result.files.map((file) => file.drawingNo),
    ['D0201-01', 'D0201-02', 'D0201-03'],
    '按图号自然序',
  );
  assert.deepEqual(
    result.result.files.map((file) => file.fileName),
    ['图纸-02.pdf', '图纸-03.pdf', '图纸-01.pdf'],
    '文件名顺序应与图号顺序不同（证明按图号排序生效）',
  );

  const outputBytes = await readFile(outputPath);
  const outputDoc = await PDFDocument.load(outputBytes);
  assert.equal(outputDoc.getPageCount(), 3);
  for (const page of outputDoc.getPages()) {
    const { width, height } = page.getSize();
    assert.ok(Math.abs(width - A4_WIDTH_PT) < 0.5, `页宽应适配 A4 横向: ${width}`);
    assert.ok(Math.abs(height - A4_HEIGHT_PT) < 0.5, `页高应适配 A4 横向: ${height}`);
  }

  const pages = await readPageTexts(outputBytes);
  const expected = ['D0201-01', 'D0201-02', 'D0201-03'];
  pages.forEach((page, index) => {
    assert.match(page.texts, new RegExp(expected[index]), `第 ${index + 1} 页应含图号 ${expected[index]}`);
  });
});

test("合并：paper='keep' 保留原纸型 + 文件名自然排序", async () => {
  const outputPath = join(TMP, 'merged-keep.pdf');
  const result = await mergePdfs({
    files: mergeSet,
    outputPath,
    paper: 'keep',
    sortMode: 'filename',
    config: {},
  });
  assert.equal(result.result.pages, 3);
  assert.deepEqual(
    result.result.files.map((file) => file.fileName),
    ['图纸-01.pdf', '图纸-02.pdf', '图纸-03.pdf'],
  );

  const outputDoc = await PDFDocument.load(await readFile(outputPath));
  for (const page of outputDoc.getPages()) {
    const { width, height } = page.getSize();
    assert.ok(Math.abs(width - A4_PORTRAIT.width) < 0.5, `keep 页宽应保持 A4 纵向: ${width}`);
    assert.ok(Math.abs(height - A4_PORTRAIT.height) < 0.5, `keep 页高应保持 A4 纵向: ${height}`);
  }
});

test('合并：sortMode=given 保留调用方顺序（COM 逐框输出场景）', async () => {
  const files = [...mergeSet].reverse().map((file) => file.path);
  const outputPath = join(TMP, 'merged-given.pdf');
  const result = await mergePdfs({ files, outputPath, paper: 'keep', sortMode: 'given', config: {} });
  assert.equal(result.result.pages, 3);
  assert.deepEqual(
    result.result.files.map((file) => file.fileName),
    ['图纸-03.pdf', '图纸-02.pdf', '图纸-01.pdf'],
    'given 模式保持调用方传入顺序',
  );
});

test('合并：空文件列表报 E1023', async () => {
  await assert.rejects(
    () => mergePdfs({ files: [], outputPath: join(TMP, 'empty.pdf'), paper: 'A4-landscape', sortMode: 'drawingNo', config: {} }),
    (err) => err.code === 'E1023',
  );
});
