/**
 * P4：市场索引构建 / 官方格式 1 校验（docs/P4-发布与分发.md §1.4）
 *
 * 断言来源：
 *   - 官方 `@deepseek-ai/dsh-plugin-market-index` 的 parseMarketIndex（2026-10-20 源码级核对）；
 *   - docs/P4-发布与分发.md §0.1 的字段规则表。
 *
 * 本文件不触网、不写 suite/、不触碰 tests-local；临时产物只落在 git 忽略的 runtime\tmp\tests\。
 */
import test, { after, before } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync } from 'node:fs';
import { mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';
import {
  buildIndex,
  compareSemver,
  isValidDshRange,
  MarketIndexError,
  PACKAGE_NAME,
  SEMVER,
  validateIndex,
} from '../tools/build-index.mjs';

const HERE = dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = resolve(HERE, '..');
const TMP = join(REPO_ROOT, 'runtime', 'tmp', 'tests', `market-index-${process.pid}`);
const EXPECTED_INDEX_URL = 'https://underfacet.github.io/dsh-office-suite-for-grid/index.json';
const MAX_TIMESTAMP = 1e11;

let pkg;
let entry;
let built;

before(async () => {
  await mkdir(TMP, { recursive: true });
  pkg = JSON.parse(await readFile(join(REPO_ROOT, 'suite', 'package.json'), 'utf8'));
  entry = JSON.parse(await readFile(join(REPO_ROOT, 'market-entry.json'), 'utf8'));
  built = await buildIndex({ out: join(TMP, 'index.json') });
});

after(async () => {
  await rm(TMP, { recursive: true, force: true });
});

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

/** 修改合法产物的一处，断言校验器在指定 JSON path 抛 MarketIndexError。 */
function expectInvalid(mutator, path) {
  const candidate = clone(built.index);
  mutator(candidate);
  assert.throws(
    () => validateIndex(candidate, EXPECTED_INDEX_URL),
    (error) => {
      assert.ok(error instanceof MarketIndexError, `应抛 MarketIndexError，实际：${error}`);
      assert.equal(error.path, path);
      return true;
    },
  );
}

test('build-index：产物与磁盘一致，indexUrl 为 Pages 直连规范 URL', async () => {
  assert.equal(built.indexUrl, EXPECTED_INDEX_URL);
  assert.equal(await readFile(built.indexPath, 'utf8'), built.json);
  const raw = JSON.parse(await readFile(built.indexPath, 'utf8'));
  assert.deepEqual(raw, built.index);
});

test('官方格式 1：format / title / description / metadata 必填项', () => {
  const raw = built.index;
  assert.equal(raw.format, 1);
  assert.equal(Object.hasOwn(raw, 'redirect'), false, '索引不得表达 redirect 语义（市场抓取固定 redirect:"error"）');

  assert.deepEqual(Object.keys(raw.title).sort(), ['en', 'zh']);
  for (const [language, text] of Object.entries(raw.title)) {
    assert.match(language, /^[a-z]{2,8}(?:-[a-z0-9]{1,8})*$/);
    assert.equal(typeof text, 'string');
    assert.ok(text.trim().length > 0);
  }
  assert.deepEqual(raw.title, entry.title);
  assert.deepEqual(raw.title, raw.plugins[0].title);

  assert.ok(raw.description && typeof raw.description === 'object');
  assert.ok(raw.description.en.trim().length > 0);
  assert.equal(raw.description.zh, entry.description.zh);

  assert.deepEqual(raw.metadata.categories, entry.categories);
  assert.deepEqual(raw.metadata.labels, entry.labels);
  assert.deepEqual(raw.metadata.categories.tools.title, { en: 'Tools', zh: '工具' });
  assert.deepEqual(raw.metadata.labels.windows.title, { en: 'Windows', zh: 'Windows' });
});

test('官方格式 1：插件 name / title / description / author / URL', () => {
  const raw = built.index;
  assert.ok(Array.isArray(raw.plugins) && raw.plugins.length === 1);
  const plugin = raw.plugins[0];

  assert.equal(plugin.name, pkg.name);
  assert.ok(plugin.name.length <= 214);
  assert.match(plugin.name, PACKAGE_NAME);
  assert.match(plugin.name, /^[a-z0-9][a-z0-9._~-]*$/);

  assert.deepEqual(plugin.title, entry.title);
  assert.deepEqual(plugin.description, entry.description);
  assert.equal(plugin.author, entry.author);
  assert.equal(plugin.homepage, entry.homepage);
  assert.equal(new URL(plugin.homepage).protocol, 'https:');
  assert.deepEqual(plugin.repository, {
    url: entry.repository.url,
    directory: entry.repository.directory,
  });
  const repoUrl = new URL(plugin.repository.url);
  assert.equal(repoUrl.protocol, 'https:');
  assert.equal(repoUrl.username, '');
  assert.equal(repoUrl.password, '');
  assert.equal(repoUrl.search, '');
  assert.equal(repoUrl.hash, '');
  assert.match(plugin.repository.directory, /^[A-Za-z0-9._@+-]+$/);

  // icon：与索引同源（官方规则：https 任意来源，或与索引同源的 http）。
  assert.equal(plugin.icon, 'https://underfacet.github.io/dsh-office-suite-for-grid/icon.svg');
  const iconUrl = new URL(plugin.icon);
  assert.equal(iconUrl.protocol, 'https:');
  assert.equal(iconUrl.origin, new URL(EXPECTED_INDEX_URL).origin);

  assert.equal(plugin.metadata.category, entry.category);
  assert.ok(Object.hasOwn(raw.metadata.categories, plugin.metadata.category));
  assert.deepEqual(plugin.metadata.keywords, entry.keywords);
  assert.deepEqual(plugin.metadata.labels, Object.keys(entry.labels));
  assert.equal(new Set(plugin.metadata.labels).size, plugin.metadata.labels.length);
  for (const label of plugin.metadata.labels) assert.ok(Object.hasOwn(raw.metadata.labels, label));
});

test('官方格式 1：versions 条目与 suite/package.json 一致', () => {
  const plugin = built.index.plugins[0];
  assert.ok(Array.isArray(plugin.versions) && plugin.versions.length > 0);
  assert.ok(plugin.versions.some((version) => version.version === pkg.version), '必须包含当前 package.json 版本');

  const seen = new Set();
  for (const versionEntry of plugin.versions) {
    assert.match(versionEntry.version, SEMVER);
    assert.ok(Number.isInteger(versionEntry.publishTimestamp));
    assert.ok(versionEntry.publishTimestamp >= 0 && versionEntry.publishTimestamp < MAX_TIMESTAMP);
    assert.deepEqual(versionEntry.source, { type: 'npm' });
    assert.deepEqual(versionEntry.engines, { dsh: pkg.engines.dsh });
    assert.ok(isValidDshRange(versionEntry.engines.dsh));
    assert.equal(seen.has(versionEntry.version), false, `版本重复：${versionEntry.version}`);
    seen.add(versionEntry.version);
  }
  for (let index = 1; index < plugin.versions.length; index += 1) {
    assert.ok(
      compareSemver(plugin.versions[index - 1].version, plugin.versions[index].version) >= 0,
      'versions 必须按 SemVer 降序',
    );
  }

  // 解析器级别：整体通过（与官方 parseMarketIndex 的必填规则等价）。
  assert.equal(validateIndex(built.index, EXPECTED_INDEX_URL), true);
});

test('build-index --versions：显式历史版本降序 + Pages 三件套', async () => {
  const older = pkg.version === '0.0.1' ? '0.0.0' : '0.0.1';
  const out = join(TMP, 'versions-index.json');
  const pagesDir = join(TMP, 'pages');
  const result = await buildIndex({ out, pagesDir, versions: [older, pkg.version] });

  const raw = JSON.parse(await readFile(out, 'utf8'));
  assert.deepEqual(
    raw.plugins[0].versions.map((versionEntry) => versionEntry.version),
    [pkg.version, older],
  );
  assert.equal(result.pagesDir, pagesDir);
  assert.ok(existsSync(join(pagesDir, 'index.json')));
  assert.ok(existsSync(join(pagesDir, 'icon.svg')));
  assert.ok(existsSync(join(pagesDir, 'index.html')));

  const pagesIndex = JSON.parse(await readFile(join(pagesDir, 'index.json'), 'utf8'));
  assert.deepEqual(pagesIndex, raw);

  const html = await readFile(join(pagesDir, 'index.html'), 'utf8');
  assert.match(html, /^<!doctype html>/);
  assert.match(html, /DSH 办公·电网绘图套件/);
  assert.match(html, /index\.json/);
  assert.match(html, /icon\.svg/);
  assert.match(html, /redirect: "error"|redirect: &quot;error&quot;/);
  assert.match(html, new RegExp(pkg.version.replaceAll('.', '\\.')));
  assert.equal((html.match(/<html/g) ?? []).length, 1);
});

test('build-index 默认版本：并入 docs/releases.json 历史记录并采用其时间戳', async () => {
  const fakeRoot = join(TMP, 'history-root');
  await mkdir(join(fakeRoot, 'suite'), { recursive: true });
  await mkdir(join(fakeRoot, 'docs'), { recursive: true });
  await writeFile(
    join(fakeRoot, 'suite', 'package.json'),
    `${JSON.stringify({ ...pkg, version: '0.2.0' }, null, 2)}\n`,
    'utf8',
  );
  await writeFile(join(fakeRoot, 'market-entry.json'), `${JSON.stringify(entry, null, 2)}\n`, 'utf8');
  await writeFile(
    join(fakeRoot, 'docs', 'releases.json'),
    `${JSON.stringify(
      [
        { version: '0.1.0', publishTimestamp: 1700000000 },
        { version: '0.2.0', publishTimestamp: 1710000000 },
      ],
      null,
      2,
    )}\n`,
    'utf8',
  );

  const result = await buildIndex({ root: fakeRoot, out: join(TMP, 'history-index.json') });
  const versions = result.index.plugins[0].versions;
  assert.deepEqual(versions.map((versionEntry) => versionEntry.version), ['0.2.0', '0.1.0']);
  assert.equal(versions[0].publishTimestamp, 1710000000);
  assert.equal(versions[1].publishTimestamp, 1700000000);
  assert.equal(result.warnings.length, 0);
  assert.equal(versions[0].engines.dsh, pkg.engines.dsh);
});

test('CLI：--out/--versions 可执行，未知参数非零退出', async () => {
  const out = join(TMP, 'cli-index.json');
  const ok = spawnSync(
    process.execPath,
    [join(REPO_ROOT, 'tools', 'build-index.mjs'), '--out', out, '--versions', pkg.version],
    { cwd: REPO_ROOT, encoding: 'utf8' },
  );
  assert.equal(ok.status, 0, ok.stderr || ok.stdout);
  const raw = JSON.parse(await readFile(out, 'utf8'));
  assert.deepEqual(raw.plugins[0].versions.map((versionEntry) => versionEntry.version), [pkg.version]);

  const bad = spawnSync(process.execPath, [join(REPO_ROOT, 'tools', 'build-index.mjs'), '--nope'], {
    cwd: REPO_ROOT,
    encoding: 'utf8',
  });
  assert.notEqual(bad.status, 0);
  assert.match(bad.stderr, /未知参数/);
});

test('负向校验：format / redirect / 包名 / SemVer / 时间戳', () => {
  expectInvalid((index) => {
    index.format = 2;
  }, '$.format');
  expectInvalid((index) => {
    index.redirect = true;
  }, '$.redirect');
  expectInvalid((index) => {
    index.plugins[0].name = 'Bad Name';
  }, '$.plugins[0].name');
  expectInvalid((index) => {
    index.plugins[0].versions[0].version = '1.0';
  }, '$.plugins[0].versions[0].version');
  expectInvalid((index) => {
    index.plugins[0].versions[0].publishTimestamp = 1.5;
  }, '$.plugins[0].versions[0].publishTimestamp');
  expectInvalid((index) => {
    index.plugins[0].versions[0].publishTimestamp = MAX_TIMESTAMP;
  }, '$.plugins[0].versions[0].publishTimestamp');
  expectInvalid((index) => {
    index.plugins[0].versions = [];
  }, '$.plugins[0].versions');
  expectInvalid((index) => {
    index.plugins[0].versions.push(clone(index.plugins[0].versions[0]));
  }, `$.plugins[0].versions[${built.index.plugins[0].versions.length}].version`);
});

test('负向校验：source / engines / 元数据引用', () => {
  expectInvalid((index) => {
    index.plugins[0].versions[0].source = { type: 'http' };
  }, '$.plugins[0].versions[0].source.type');
  expectInvalid((index) => {
    index.plugins[0].versions[0].engines.dsh = 'workspace:*';
  }, '$.plugins[0].versions[0].engines.dsh');
  expectInvalid((index) => {
    index.plugins[0].versions[0].engines.dsh = 'not a range';
  }, '$.plugins[0].versions[0].engines.dsh');
  expectInvalid((index) => {
    index.plugins[0].metadata.labels = ['linux'];
  }, '$.plugins[0].metadata.labels[0]');
  expectInvalid((index) => {
    index.plugins[0].metadata.category = 'office';
  }, '$.plugins[0].metadata.category');
  expectInvalid((index) => {
    index.title = { zh: '缺少 en' };
  }, '$.title');
});

test('负向校验：homepage / repository / icon URL 规则', () => {
  expectInvalid((index) => {
    index.plugins[0].homepage = 'ftp://example.com/home';
  }, '$.plugins[0].homepage');
  expectInvalid((index) => {
    index.plugins[0].repository.url += '?tab=readme';
  }, '$.plugins[0].repository.url');
  expectInvalid((index) => {
    index.plugins[0].repository.directory = '../escape';
  }, '$.plugins[0].repository.directory');
  expectInvalid((index) => {
    index.plugins[0].icon = 'http://evil.example/icon.svg';
  }, '$.plugins[0].icon');
  // origin 含 scheme：https 索引下，同名主机的 http 图标也不是同源（与官方 icon 规则一致）。
  expectInvalid((index) => {
    index.plugins[0].icon = 'http://underfacet.github.io/dsh-office-suite-for-grid/icon.svg';
  }, '$.plugins[0].icon');

  // 官方规则允许任意 https 图标（不要求同源）。
  const httpsIcon = clone(built.index);
  httpsIcon.plugins[0].icon = 'https://cdn.example.com/icon.svg';
  assert.equal(validateIndex(httpsIcon, EXPECTED_INDEX_URL), true);
});

test('SemVer 范围 / 排序辅助函数', () => {
  assert.equal(isValidDshRange('>=0.2.0-rc.0 <0.3.0'), true);
  assert.equal(isValidDshRange('^1.2.3'), true);
  assert.equal(isValidDshRange('~1.2'), true);
  assert.equal(isValidDshRange('1.2.x'), true);
  assert.equal(isValidDshRange('1.2.3 - 2.0.0'), true);
  assert.equal(isValidDshRange('>= 0.2.0 < 0.3.0'), true);
  assert.equal(isValidDshRange('workspace:*'), false);
  assert.equal(isValidDshRange('not-a-range'), false);
  assert.equal(isValidDshRange(''), false);

  assert.equal(compareSemver('1.0.0', '1.0.0-rc.1'), 1);
  assert.equal(compareSemver('1.0.0-alpha', '1.0.0-beta'), -1);
  assert.equal(compareSemver('1.0.0-alpha.1', '1.0.0-alpha'), 1);
  assert.equal(compareSemver('0.2.1', '0.2.0'), 1);
  assert.equal(compareSemver('1.0.0+build.2', '1.0.0+build.1'), 0);
});
