#!/usr/bin/env node
/**
 * P4 市场索引构建器 —— 官方格式 1（单插件），无第三方依赖，纯 Node ESM。
 *
 * 输入：
 *   - market-entry.json          手工维护的插件元信息（根目录）
 *   - suite/package.json         name / version / engines.dsh（权威版本与兼容范围）
 *   - docs/releases.json         可选：历史版本发布时间戳 [{ version, publishTimestamp }]
 *   - 环境变量 SOURCE_DATE_EPOCH 可选：固定本次构建的发布时间（CI 用 tag 提交时间，保证 Pages 与 Releases 产出同一份 index.json）
 *
 * 输出：
 *   - 默认 dist/index.json
 *   - 可选 --pages <dir>：<dir>/index.json + icon.svg + index.html（gh-pages 发布目录）
 *
 * 用法：
 *   node tools/build-index.mjs [--out <path>] [--versions <v1,v2,...>] [--pages <dir>]
 *
 * 校验规则逐条复刻 DSH 官方
 * `@deepseek-ai/dsh-plugin-market-index` 的 parseMarketIndex（源码级核对，2026-10-20）：
 *   format===1；localizedText 必须含 en；npm 包名正则/长度；精确 SemVer；publishTimestamp
 *   整数 ∈ [0,1e11)；versions 非空且不重复；source 为 npm（git 形式也校验，但本项目不用）；
 *   engines.dsh 为 SemVer 范围且禁止 workspace:；homepage 为绝对 http(s)；repository.url
 *   无凭据/query/fragment；icon 为 data: 或 https/同源 URL；metadata 只引用已声明 id。
 *
 * 与官方解析器的两处主动差异（均为更严格，产物本身不受影响）：
 *   1. 顶层 `redirect` 字段：官方解析器丢弃未知字段；由于市场抓取固定 redirect:"error"，
 *      本脚本主动拒绝 `redirect: true`，避免误配置。
 *   2. homepage 额外拒绝含凭据的 URL（官方只查协议）。
 */
import { existsSync } from 'node:fs';
import { copyFile, mkdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));

/** 仓库根目录（tools/ 的上一级）。 */
export const REPO_ROOT = resolve(HERE, '..');

/** 未在 market-entry.json 声明 indexUrl 时的兜底值（必须与 README 完全一致，禁止重定向变体）。 */
export const DEFAULT_INDEX_URL = 'https://underfacet.github.io/dsh-office-suite-for-grid/index.json';

export const LANGUAGE_ID = /^[a-z]{2,8}(?:-[a-z0-9]{1,8})*$/u;
export const PACKAGE_NAME = /^(?:@[a-z0-9][a-z0-9._~-]*\/)?[a-z0-9][a-z0-9._~-]*$/u;
export const PACKAGE_NAME_MAX_LENGTH = 214;
const NUMERIC = '(?:0|[1-9]\\d*)';
const PRERELEASE_PART = `(?:${NUMERIC}|\\d*[a-zA-Z-][0-9a-zA-Z-]*)`;
export const SEMVER = new RegExp(
  `^${NUMERIC}\\.${NUMERIC}\\.${NUMERIC}(?:-${PRERELEASE_PART}(?:\\.${PRERELEASE_PART})*)?(?:\\+[0-9a-zA-Z-]+(?:\\.[0-9a-zA-Z-]+)*)?$`,
  'u',
);
export const COMMIT = /^[0-9a-f]{40}$/u;
export const PATH_SEGMENT = /^[A-Za-z0-9._@+-]+$/u;
export const DATA_ICON = /^data:(?:image\/svg\+xml|image\/png|image\/jpeg|image\/webp);base64,([A-Za-z0-9+/]*={0,2})$/u;
const BASE64 = /^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/u;
export const MAX_INLINE_ICON_BYTES = 32 * 1024;
export const MAX_PUBLISH_TIMESTAMP = 1e11;

/** 官方 MarketIndexError 的等价物：携带 JSON 路径与违规原因。 */
export class MarketIndexError extends Error {
  constructor(path, reason) {
    super(`${path} ${reason}`);
    this.path = path;
    this.reason = reason;
    this.name = 'MarketIndexError';
  }
}

function fail(path, reason) {
  throw new MarketIndexError(path, reason);
}

function objectAt(value, path) {
  if (typeof value !== 'object' || value === null || Array.isArray(value)) fail(path, 'must be an object');
  return value;
}

function stringAt(value, path) {
  if (typeof value !== 'string' || value.trim() === '') fail(path, 'must be a string with a non-whitespace character');
  return value;
}

function localizedText(value, path) {
  if (typeof value === 'string') return stringAt(value, path);
  const object = objectAt(value, path);
  const text = {};
  for (const [language, member] of Object.entries(object)) {
    if (!LANGUAGE_ID.test(language)) fail(`${path}.${language}`, 'must be keyed by a lowercase language id');
    text[language] = stringAt(member, `${path}.${language}`);
  }
  if (!Object.hasOwn(text, 'en')) fail(path, 'must contain en');
  return text;
}

/** 绝对 http(s) URL；默认拒绝凭据、query、fragment（与官方 httpUrl 一致）。 */
function httpUrl(value, path, protocols = ['https:', 'http:']) {
  const text = stringAt(value, path);
  let url;
  try {
    url = new URL(text);
  } catch {
    fail(path, 'must be an absolute URL');
  }
  if (!protocols.includes(url.protocol)) fail(path, `must use ${protocols.join(' or ')}`);
  if (url.username !== '' || url.password !== '') fail(path, 'must not contain credentials');
  if (url.href.includes('?') || url.href.includes('#')) fail(path, 'must not contain a query or fragment');
  return url;
}

/** homepage 允许 query/fragment（官方 homepageUrl），但额外拒绝凭据。 */
function homepageUrl(value, path) {
  const text = stringAt(value, path);
  let url;
  try {
    url = new URL(text);
  } catch {
    fail(path, 'must be an absolute URL');
  }
  if (url.protocol !== 'https:' && url.protocol !== 'http:') fail(path, 'must use https: or http:');
  if (url.username !== '' || url.password !== '') fail(path, 'must not contain credentials');
  return url.href;
}

function packagePath(value, path) {
  const text = stringAt(value, path);
  if (text.split('/').some((segment) => segment === '.' || segment === '..' || !PATH_SEGMENT.test(segment))) {
    fail(path, 'must be a relative directory of letters, digits, and ._@+- segments without . or ..');
  }
  return text;
}

function repository(value, path) {
  const object = objectAt(value, path);
  const url = httpUrl(object.url, `${path}.url`).href;
  const directory = object.directory === undefined ? undefined : packagePath(object.directory, `${path}.directory`);
  return { url, ...(directory === undefined ? {} : { directory }) };
}

function source(value, path) {
  const object = objectAt(value, path);
  switch (object.type) {
    case 'npm': {
      const registry = object.registry === undefined
        ? undefined
        : httpUrl(object.registry, `${path}.registry`).href;
      return { type: 'npm', ...(registry === undefined ? {} : { registry }) };
    }
    case 'git': {
      const commit = stringAt(object.commit, `${path}.commit`);
      if (!COMMIT.test(commit)) fail(`${path}.commit`, 'must be a lowercase 40-hex-digit commit SHA');
      const url = httpUrl(object.url, `${path}.url`, ['https:']);
      if (url.href !== object.url) fail(`${path}.url`, 'must be a canonical URL');
      if (!/^(?:\/[^/]+){2,}$/u.test(url.pathname)) fail(`${path}.url`, 'must name a repository with at least two path segments');
      if (/\.(?:tgz|tar\.gz)$/iu.test(url.pathname)) fail(`${path}.url`, 'must name a repository, not a tarball');
      const packageDir = object.path === undefined ? undefined : packagePath(object.path, `${path}.path`);
      return { type: 'git', url: url.href, commit, ...(packageDir === undefined ? {} : { path: packageDir }) };
    }
    default:
      return fail(`${path}.type`, 'must be "npm" or "git"');
  }
}

/**
 * 轻量 SemVer 范围校验（复刻 semver.validRange 的常用子集）：
 * 支持 `>=1.2.3 <2`、`^1.2.3`、`~1.2`、`1.2.x`、`*`、`1.2.3 - 2.0.0`、`||`。
 * 本项目实际只使用 suite/package.json 的单一范围；不引入 semver 依赖以保持零依赖。
 */
export function isValidDshRange(range) {
  if (typeof range !== 'string') return false;
  const text = range.trim();
  if (text === '' || text.startsWith('workspace:')) return false;
  const alternatives = text.split('||').map((part) => part.trim());
  if (alternatives.some((part) => part === '')) return false;
  for (const alternative of alternatives) {
    // 允许 `>= 1.2.3` 这类操作符与版本之间的空格
    const normalized = alternative.replace(/([<>=~^]+)\s+/g, '$1');
    const hyphen = normalized.match(/^(\S+)\s+-\s+(\S+)$/);
    if (hyphen) {
      if (!isPartialVersion(hyphen[1]) || !isPartialVersion(hyphen[2])) return false;
      continue;
    }
    const tokens = normalized.split(/\s+/);
    if (tokens.some((token) => !isComparatorToken(token))) return false;
  }
  return true;
}

const PARTIAL_VERSION = new RegExp(
  `^v?(?:\\*|[xX]|${NUMERIC})(?:\\.(?:\\*|[xX]|${NUMERIC})){0,2}` +
    '(?:-[0-9A-Za-z-]+(?:\\.[0-9A-Za-z-]+)*)?(?:\\+[0-9A-Za-z-]+(?:\\.[0-9A-Za-z-]+)*)?$',
  'u',
);

function isPartialVersion(text) {
  return PARTIAL_VERSION.test(text);
}

function isComparatorToken(token) {
  if (token === '' || token === '*' || token === 'x' || token === 'X') return true;
  const match = token.match(/^(?:[<>=~^]+)?(.+)$/u);
  if (!match) return false;
  const version = match[1];
  if (version === '' || version === '*' || version === 'x' || version === 'X') return true;
  return isPartialVersion(version);
}

function parseSemver(version) {
  const match = /^(\d+)\.(\d+)\.(\d+)(?:-([0-9A-Za-z.-]+))?(?:\+([0-9A-Za-z.-]+))?$/u.exec(version);
  if (!match) return null;
  return {
    major: Number(match[1]),
    minor: Number(match[2]),
    patch: Number(match[3]),
    pre: match[4] === undefined ? [] : match[4].split('.'),
  };
}

/** SemVer 2.0.0 优先级比较：返回 -1 / 0 / 1。 */
export function compareSemver(left, right) {
  const a = parseSemver(left);
  const b = parseSemver(right);
  if (a === null || b === null) throw new Error(`不是精确 SemVer：${JSON.stringify(a === null ? left : right)}`);
  for (const key of ['major', 'minor', 'patch']) {
    if (a[key] !== b[key]) return a[key] < b[key] ? -1 : 1;
  }
  if (a.pre.length === 0 && b.pre.length === 0) return 0;
  if (a.pre.length === 0) return 1;
  if (b.pre.length === 0) return -1;
  const length = Math.max(a.pre.length, b.pre.length);
  for (let index = 0; index < length; index += 1) {
    const x = a.pre[index];
    const y = b.pre[index];
    if (x === undefined) return -1;
    if (y === undefined) return 1;
    if (x === y) continue;
    const xNumeric = /^\d+$/u.test(x);
    const yNumeric = /^\d+$/u.test(y);
    if (xNumeric && yNumeric) return Number(x) < Number(y) ? -1 : 1;
    if (xNumeric) return -1;
    if (yNumeric) return 1;
    return x < y ? -1 : 1;
  }
  return 0;
}

function engines(value, path) {
  const object = objectAt(value, path);
  const dsh = object.dsh === undefined
    ? undefined
    : (() => {
        const range = stringAt(object.dsh, `${path}.dsh`);
        if (range.startsWith('workspace:')) fail(`${path}.dsh`, 'must not be a workspace: range');
        if (!isValidDshRange(range)) fail(`${path}.dsh`, 'must be a SemVer range');
        return range;
      })();
  return dsh === undefined ? {} : { dsh };
}

function icon(value, path, indexUrl) {
  const text = stringAt(value, path);
  if (text.startsWith('data:')) {
    const payload = DATA_ICON.exec(text)?.[1];
    if (payload === undefined || payload === '' || !BASE64.test(payload)) {
      fail(path, 'must be a base64 SVG, PNG, JPEG, or WebP data URL');
    }
    const decoded = (payload.length / 4) * 3 - (payload.endsWith('==') ? 2 : payload.endsWith('=') ? 1 : 0);
    if (decoded > MAX_INLINE_ICON_BYTES) fail(path, `must not exceed ${String(MAX_INLINE_ICON_BYTES / 1024)} KiB of decoded bytes`);
    return text;
  }
  let resolved;
  try {
    resolved = new URL(text, indexUrl);
  } catch {
    fail(path, 'must be a URL reference');
  }
  const base = new URL(indexUrl);
  if (resolved.protocol !== 'https:' && (resolved.origin === 'null' || resolved.origin !== base.origin)) {
    fail(path, 'must resolve to https or the index origin');
  }
  return resolved.href;
}

function marketVersion(value, path) {
  const object = objectAt(value, path);
  const version = stringAt(object.version, `${path}.version`);
  if (!SEMVER.test(version)) fail(`${path}.version`, 'must be an exact SemVer version');
  const publishTimestamp = object.publishTimestamp;
  if (
    typeof publishTimestamp !== 'number' ||
    !Number.isInteger(publishTimestamp) ||
    publishTimestamp < 0 ||
    publishTimestamp >= MAX_PUBLISH_TIMESTAMP
  ) {
    fail(`${path}.publishTimestamp`, 'must be an integer number of Unix seconds in [0, 10^11)');
  }
  return {
    version,
    publishTimestamp,
    source: source(object.source, `${path}.source`),
    engines: object.engines === undefined ? {} : engines(object.engines, `${path}.engines`),
  };
}

function versions(value, path) {
  if (!Array.isArray(value) || value.length === 0) fail(path, 'must be a non-empty array');
  const seen = new Set();
  return value
    .map((member, position) => {
      const at = `${path}[${String(position)}]`;
      const entry = marketVersion(member, at);
      if (seen.has(entry.version)) fail(`${at}.version`, `duplicates ${entry.version}`);
      seen.add(entry.version);
      return entry;
    })
    .sort((left, right) => compareSemver(right.version, left.version));
}

function declaration(object, path) {
  return { title: localizedText(object.title, `${path}.title`) };
}

function sortKeyDeclaration(object, path) {
  if (object.order !== 'asc' && object.order !== 'desc') fail(`${path}.order`, 'must be "asc" or "desc"');
  return { title: localizedText(object.title, `${path}.title`), order: object.order };
}

function declarations(value, path, read) {
  if (value === undefined) return {};
  const result = {};
  for (const [id, member] of Object.entries(objectAt(value, path))) {
    if (id === '') fail(path, 'must not declare an empty id');
    result[id] = read(objectAt(member, `${path}.${id}`), `${path}.${id}`);
  }
  return result;
}

function indexMetadata(value, path) {
  const object = value === undefined ? {} : objectAt(value, path);
  return {
    categories: declarations(object.categories, `${path}.categories`, declaration),
    labels: declarations(object.labels, `${path}.labels`, declaration),
    sortKeys: declarations(object.sortKeys, `${path}.sortKeys`, sortKeyDeclaration),
  };
}

function declaredId(value, path, declared) {
  const id = stringAt(value, path);
  if (!Object.hasOwn(declared, id)) fail(path, `references undeclared id ${JSON.stringify(id)}`);
  return id;
}

function pluginMetadata(value, path, index) {
  const object = value === undefined ? {} : objectAt(value, path);
  const category = object.category === undefined ? undefined : declaredId(object.category, `${path}.category`, index.categories);
  let keywords = [];
  if (object.keywords !== undefined) {
    if (!Array.isArray(object.keywords)) fail(`${path}.keywords`, 'must be an array');
    keywords = object.keywords.map((keyword, position) => {
      if (typeof keyword !== 'string') fail(`${path}.keywords[${String(position)}]`, 'must be a string');
      return keyword;
    });
  }
  let labels = [];
  if (object.labels !== undefined) {
    if (!Array.isArray(object.labels)) fail(`${path}.labels`, 'must be an array');
    labels = object.labels.map((label, position) => declaredId(label, `${path}.labels[${String(position)}]`, index.labels));
    if (new Set(labels).size !== labels.length) fail(`${path}.labels`, 'must not repeat a label');
  }
  const sortKeys = {};
  if (object.sortKeys !== undefined) {
    for (const [id, sortValue] of Object.entries(objectAt(object.sortKeys, `${path}.sortKeys`))) {
      declaredId(id, `${path}.sortKeys.${id}`, index.sortKeys);
      if (typeof sortValue !== 'number' || !Number.isFinite(sortValue)) fail(`${path}.sortKeys.${id}`, 'must be a finite number');
      sortKeys[id] = sortValue;
    }
  }
  return { ...(category === undefined ? {} : { category }), keywords, labels, sortKeys };
}

function plugin(value, path, index, indexUrl) {
  const object = objectAt(value, path);
  const name = stringAt(object.name, `${path}.name`);
  if (name.length > PACKAGE_NAME_MAX_LENGTH || !PACKAGE_NAME.test(name)) fail(`${path}.name`, 'must be a valid npm package name');
  const title = object.title === undefined ? undefined : localizedText(object.title, `${path}.title`);
  const description = object.description === undefined ? undefined : localizedText(object.description, `${path}.description`);
  const iconUrl = object.icon === undefined ? undefined : icon(object.icon, `${path}.icon`, indexUrl);
  const author = object.author === undefined ? undefined : stringAt(object.author, `${path}.author`);
  const homepage = object.homepage === undefined ? undefined : homepageUrl(object.homepage, `${path}.homepage`);
  const repo = object.repository === undefined ? undefined : repository(object.repository, `${path}.repository`);
  return {
    name,
    ...(title === undefined ? {} : { title }),
    ...(description === undefined ? {} : { description }),
    ...(iconUrl === undefined ? {} : { icon: iconUrl }),
    ...(author === undefined ? {} : { author }),
    ...(homepage === undefined ? {} : { homepage }),
    ...(repo === undefined ? {} : { repository: repo }),
    metadata: pluginMetadata(object.metadata, `${path}.metadata`, index),
    versions: versions(object.versions, `${path}.versions`),
  };
}

/**
 * 校验一份格式 1 索引；违规抛 MarketIndexError{path, reason}。
 * @param {unknown} value 已解析的 JSON 文档
 * @param {string} indexUrl 索引的绝对 URL（icon 相对引用与同源校验用）
 * @returns {true}
 */
export function validateIndex(value, indexUrl) {
  let base;
  try {
    base = new URL(indexUrl);
  } catch {
    fail('$', `index URL ${JSON.stringify(indexUrl)} must be an absolute URL`);
  }
  const object = objectAt(value, '$');
  if (object.format !== 1) fail('$.format', 'must be 1');
  if (Object.hasOwn(object, 'redirect') && object.redirect !== false) {
    fail('$.redirect', 'must not be true: the market fetches indexes with redirect "error"');
  }
  const metadata = indexMetadata(object.metadata, '$.metadata');
  if (!Array.isArray(object.plugins)) fail('$.plugins', 'must be an array');
  const names = new Set();
  object.plugins.forEach((entry, position) => {
    const parsed = plugin(entry, `$.plugins[${String(position)}]`, metadata, base.href);
    if (names.has(parsed.name)) fail(`$.plugins[${String(position)}].name`, `duplicates ${parsed.name}`);
    names.add(parsed.name);
  });
  localizedText(object.title, '$.title');
  if (object.description !== undefined) localizedText(object.description, '$.description');
  return true;
}

/* ────────────────────────────── 构建 ────────────────────────────── */

async function readJson(path) {
  const text = await readFile(path, 'utf8');
  // Windows 编辑器可能写入 UTF-8 BOM；JSON.parse 不接受 BOM，先剥离。
  return JSON.parse(text.charCodeAt(0) === 0xfeff ? text.slice(1) : text);
}

function parseVersionList(input) {
  const parts = (Array.isArray(input) ? input : String(input).split(',')).map((part) => String(part).trim());
  return parts.filter((part) => part !== '');
}

function normalizeReleaseMap(data, path) {
  const map = new Map();
  const entries = Array.isArray(data)
    ? data
    : data && typeof data === 'object'
      ? Object.entries(data).map(([version, publishTimestamp]) => ({ version, publishTimestamp }))
      : null;
  if (entries === null) throw new Error(`${path} 必须是 [{version,publishTimestamp}] 数组或 {version: timestamp} 对象`);
  for (const [index, entry] of entries.entries()) {
    const at = `${path}[${String(index)}]`;
    if (!entry || typeof entry !== 'object') throw new Error(`${at} 必须是对象`);
    const version = entry.version;
    if (typeof version !== 'string' || !SEMVER.test(version)) throw new Error(`${at}.version 不是精确 SemVer`);
    const timestamp = entry.publishTimestamp;
    if (
      typeof timestamp !== 'number' ||
      !Number.isInteger(timestamp) ||
      timestamp < 0 ||
      timestamp >= MAX_PUBLISH_TIMESTAMP
    ) {
      throw new Error(`${at}.publishTimestamp 必须是 [0, 10^11) 的整数 Unix 秒`);
    }
    map.set(version, timestamp);
  }
  return map;
}

async function readReleaseTimestamps(root) {
  const path = join(root, 'docs', 'releases.json');
  if (!existsSync(path)) return new Map();
  return normalizeReleaseMap(await readJson(path), 'docs/releases.json');
}

function resolveBuildTimestamp() {
  const raw = process.env.SOURCE_DATE_EPOCH;
  if (raw !== undefined && raw.trim() !== '') {
    if (!/^\d+$/u.test(raw.trim())) throw new Error(`SOURCE_DATE_EPOCH 必须是整数 Unix 秒：${JSON.stringify(raw)}`);
    const timestamp = Number(raw.trim());
    if (!Number.isSafeInteger(timestamp) || timestamp < 0 || timestamp >= MAX_PUBLISH_TIMESTAMP) {
      throw new Error(`SOURCE_DATE_EPOCH 超出 [0, 10^11)：${raw}`);
    }
    return timestamp;
  }
  return Math.floor(Date.now() / 1000);
}

function normalizeDeclarations(value, path) {
  if (value === undefined) return {};
  const object = objectAt(value, path);
  const result = {};
  for (const [id, member] of Object.entries(object)) {
    if (id === '') fail(path, 'must not declare an empty id');
    const declarationObject = objectAt(member, `${path}.${id}`);
    result[id] = { title: localizedText(declarationObject.title, `${path}.${id}.title`) };
  }
  return result;
}

function normalizeKeywords(value, path) {
  if (value === undefined) return [];
  if (!Array.isArray(value)) fail(path, 'must be an array');
  return value.map((keyword, position) => {
    if (typeof keyword !== 'string') fail(`${path}[${String(position)}]`, 'must be a string');
    return keyword;
  });
}

function assertIndexUrl(value) {
  const text = stringAt(value, 'market-entry.indexUrl');
  let url;
  try {
    url = new URL(text);
  } catch {
    fail('market-entry.indexUrl', 'must be an absolute URL');
  }
  if (url.protocol !== 'https:' && url.protocol !== 'http:') fail('market-entry.indexUrl', 'must use https: or http:');
  if (url.username !== '' || url.password !== '') fail('market-entry.indexUrl', 'must not contain credentials');
  if (url.href.includes('?') || url.href.includes('#')) fail('market-entry.indexUrl', 'must not contain a query or fragment');
  if (url.href !== text) {
    fail('market-entry.indexUrl', 'must be a canonical URL: the market normalizeIndexUrl rejects non-canonical forms');
  }
  return url.href;
}

function resolveIcon(value, indexUrl) {
  const text = stringAt(value, 'market-entry.icon');
  if (text.startsWith('data:')) return icon(text, 'market-entry.icon', indexUrl);
  let resolved;
  try {
    resolved = new URL(text, indexUrl);
  } catch {
    fail('market-entry.icon', 'must be a URL reference');
  }
  const base = new URL(indexUrl);
  if (resolved.protocol !== 'https:' && (resolved.origin === 'null' || resolved.origin !== base.origin)) {
    fail('market-entry.icon', 'must resolve to https or the index origin');
  }
  return resolved.href;
}

/**
 * 构建单插件市场索引。
 * @param {object} [options]
 * @param {string} [options.root] 仓库根（测试用）
 * @param {string} [options.entry] market-entry.json 路径（相对 root 或绝对）
 * @param {string} [options.out] 输出 index.json 路径；默认 <root>/dist/index.json
 * @param {string} [options.pagesDir] 可选 Pages 目录（写 index.json + icon.svg + index.html）
 * @param {string[]|string} [options.versions] 版本列表；默认 docs/releases.json 版本 ∪ package.json 版本
 * @returns {Promise<{index:object,indexPath:string,indexUrl:string,pagesDir:string|null,warnings:string[],json:string}>}
 */
export async function buildIndex(options = {}) {
  const root = options.root === undefined ? REPO_ROOT : resolve(options.root);
  const entryPath = options.entry === undefined ? join(root, 'market-entry.json') : resolve(root, options.entry);
  const pkgPath = join(root, 'suite', 'package.json');
  const pkg = await readJson(pkgPath);
  const entry = await readJson(entryPath);
  const warnings = [];

  if (typeof pkg.name !== 'string' || !PACKAGE_NAME.test(pkg.name) || pkg.name.length > PACKAGE_NAME_MAX_LENGTH) {
    throw new Error(`suite/package.json name 不是合法 npm 包名：${JSON.stringify(pkg.name)}`);
  }
  if (entry.name !== pkg.name) {
    throw new Error(`market-entry.json name (${JSON.stringify(entry.name)}) 必须与 suite/package.json name (${JSON.stringify(pkg.name)}) 一致`);
  }
  if (typeof pkg.version !== 'string' || !SEMVER.test(pkg.version)) {
    throw new Error(`suite/package.json version 不是精确 SemVer：${JSON.stringify(pkg.version)}`);
  }
  const dshRange = pkg.engines === undefined ? undefined : pkg.engines.dsh;
  if (typeof dshRange !== 'string' || !isValidDshRange(dshRange)) {
    throw new Error(`suite/package.json engines.dsh 不是合法 SemVer 范围：${JSON.stringify(dshRange)}`);
  }

  const indexUrl = entry.indexUrl === undefined ? DEFAULT_INDEX_URL : assertIndexUrl(entry.indexUrl);
  const title = localizedText(entry.title, 'market-entry.title');
  const description = localizedText(entry.description, 'market-entry.description');
  const author = stringAt(entry.author, 'market-entry.author');
  const homepage = homepageUrl(entry.homepage, 'market-entry.homepage');
  const repo = repository(entry.repository, 'market-entry.repository');
  const iconUrl = resolveIcon(entry.icon === undefined ? 'icon.svg' : entry.icon, indexUrl);
  const categories = normalizeDeclarations(entry.categories, 'market-entry.categories');
  const labels = normalizeDeclarations(entry.labels, 'market-entry.labels');
  const category = entry.category === undefined ? Object.keys(categories)[0] : entry.category;
  if (typeof category !== 'string' || category === '' || !Object.hasOwn(categories, category)) {
    throw new Error(`market-entry.category 必须引用 market-entry.categories 中已声明的 id：${JSON.stringify(category)}`);
  }
  const keywords = normalizeKeywords(entry.keywords, 'market-entry.keywords');

  const releaseTimestamps = await readReleaseTimestamps(root);
  let versionList;
  if (options.versions !== undefined) {
    versionList = parseVersionList(options.versions);
    if (versionList.length === 0) throw new Error('--versions 不能为空');
    const seen = new Set();
    for (const version of versionList) {
      if (!SEMVER.test(version)) throw new Error(`--versions 含非精确 SemVer：${JSON.stringify(version)}`);
      if (seen.has(version)) throw new Error(`--versions 含重复版本：${version}`);
      seen.add(version);
    }
  } else {
    versionList = [...new Set([...releaseTimestamps.keys(), pkg.version])];
  }
  if (!versionList.includes(pkg.version)) {
    throw new Error(`版本列表必须包含 suite/package.json 当前版本 ${pkg.version}；收到 ${versionList.join(', ')}`);
  }

  const fallbackTimestamp = resolveBuildTimestamp();
  const versionEntries = versionList.map((version) => {
    let publishTimestamp = releaseTimestamps.get(version);
    if (publishTimestamp === undefined) {
      publishTimestamp = fallbackTimestamp;
      warnings.push(
        `版本 ${version} 未在 docs/releases.json 记录 publishTimestamp，使用 ${publishTimestamp}` +
          '（可用 SOURCE_DATE_EPOCH 固定；建议发布后补记 docs/releases.json）',
      );
    }
    return {
      version,
      publishTimestamp,
      source: { type: 'npm' },
      engines: { dsh: dshRange },
    };
  });
  versionEntries.sort((left, right) => compareSemver(right.version, left.version));

  const index = {
    format: 1,
    title,
    description,
    metadata: { categories, labels },
    plugins: [
      {
        name: pkg.name,
        title,
        description,
        icon: iconUrl,
        author,
        homepage,
        repository: repo,
        metadata: { category, keywords, labels: Object.keys(labels) },
        versions: versionEntries,
      },
    ],
  };

  validateIndex(index, indexUrl);

  const outPath = options.out === undefined ? join(root, 'dist', 'index.json') : resolve(options.out);
  const json = `${JSON.stringify(index, null, 2)}\n`;
  await mkdir(dirname(outPath), { recursive: true });
  await writeFile(outPath, json, 'utf8');

  let pagesDir = null;
  if (options.pagesDir !== undefined && options.pagesDir !== null && options.pagesDir !== '') {
    pagesDir = resolve(options.pagesDir);
    await writePages({ root, pagesDir, index, indexUrl, json });
  }

  return { index, indexPath: outPath, indexUrl, pagesDir, warnings, json };
}

async function writePages({ root, pagesDir, index, indexUrl, json }) {
  const iconSource = join(root, 'suite', 'icon.svg');
  if (!existsSync(iconSource)) throw new Error(`缺少图标源文件：${iconSource}`);
  await mkdir(pagesDir, { recursive: true });
  await writeFile(join(pagesDir, 'index.json'), json, 'utf8');
  await copyFile(iconSource, join(pagesDir, 'icon.svg'));
  await writeFile(join(pagesDir, 'index.html'), renderPagesHtml(index, indexUrl), 'utf8');
}

function escapeHtml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

/** 生成 gh-pages 的 README 摘要页（无外部资源、无脚本）。 */
export function renderPagesHtml(index, indexUrl) {
  const plugin = index.plugins[0];
  const zhTitle = index.title.zh ?? index.title.en;
  const enTitle = index.title.en;
  const zhDescription = index.description.zh ?? index.description.en;
  const enDescription = index.description.en;
  const repoUrl = plugin.repository?.url ?? plugin.homepage;
  const releasesUrl = `${repoUrl}/releases`;
  const npmUrl = `https://www.npmjs.com/package/${plugin.name}`;
  const versionRows = plugin.versions
    .map((entry) => {
      const date = new Date(entry.publishTimestamp * 1000).toISOString().slice(0, 10);
      return `        <tr><td><code>${escapeHtml(entry.version)}</code></td><td>${escapeHtml(date)}</td><td><code>${escapeHtml(entry.engines.dsh ?? '')}</code></td></tr>`;
    })
    .join('\n');
  const generated = new Date().toISOString().replace('T', ' ').slice(0, 19);

  return `<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${escapeHtml(zhTitle)} · DSH 市场索引</title>
  <style>
    :root { color-scheme: light dark; --fg: #1f2328; --muted: #57606a; --bg: #ffffff; --card: #f6f8fa; --line: #d0d7de; --accent: #0969da; }
    @media (prefers-color-scheme: dark) { :root { --fg: #e6edf3; --muted: #8b949e; --bg: #0d1117; --card: #161b22; --line: #30363d; --accent: #2f81f7; } }
    * { box-sizing: border-box; }
    body { margin: 0; background: var(--bg); color: var(--fg); font: 16px/1.65 -apple-system, "Segoe UI", "Microsoft YaHei", sans-serif; }
    main { max-width: 880px; margin: 0 auto; padding: 40px 20px 64px; }
    header { display: flex; gap: 18px; align-items: center; border-bottom: 1px solid var(--line); padding-bottom: 22px; }
    header img { width: 72px; height: 72px; }
    h1 { margin: 0 0 4px; font-size: 30px; line-height: 1.25; }
    h1 .en { font-size: 16px; color: var(--muted); font-weight: 500; }
    h2 { margin: 34px 0 10px; font-size: 20px; }
    .lead { margin: 18px 0 6px; }
    .en { color: var(--muted); }
    code { background: var(--card); border: 1px solid var(--line); border-radius: 4px; padding: 1px 5px; font-size: 90%; overflow-wrap: anywhere; }
    section { border-top: 1px solid var(--line); margin-top: 28px; }
    ol, ul { padding-left: 22px; }
    li { margin: 4px 0; }
    table { border-collapse: collapse; width: 100%; margin-top: 8px; }
    th, td { text-align: left; border-bottom: 1px solid var(--line); padding: 7px 10px; vertical-align: top; }
    th { color: var(--muted); font-weight: 600; font-size: 14px; }
    .note { background: var(--card); border: 1px solid var(--line); border-radius: 6px; padding: 10px 12px; color: var(--muted); font-size: 14px; }
    footer { margin-top: 40px; color: var(--muted); font-size: 13px; border-top: 1px solid var(--line); padding-top: 14px; }
    a { color: var(--accent); }
  </style>
</head>
<body>
<main>
  <header>
    <img src="icon.svg" alt="" width="72" height="72">
    <div>
      <h1>${escapeHtml(zhTitle)}<br><span class="en">${escapeHtml(enTitle)}</span></h1>
    </div>
  </header>

  <p class="lead">${escapeHtml(zhDescription)}</p>
  <p class="lead en">${escapeHtml(enDescription)}</p>

  <section>
    <h2>在 DSH 中添加索引</h2>
    <ol>
      <li>打开 DSH「插件 → 市场 → 添加索引」。</li>
      <li>粘贴索引 URL：<code>${escapeHtml(indexUrl)}</code></li>
      <li>在「发现」页找到「${escapeHtml(zhTitle)}」，按 <code>engines.dsh</code> 兼容性一键安装。</li>
    </ol>
    <p class="note">市场抓取索引使用 <code>redirect: "error"</code>：本页 URL 必须直接返回 200，不能是 301 变体（例如 http、大小写主机或 jsDelivr 镜像）。索引缓存 10 分钟。</p>
  </section>

  <section>
    <h2>安装渠道</h2>
    <ul>
      <li><strong>市场索引</strong>：<a href="${escapeHtml(indexUrl)}">${escapeHtml(indexUrl)}</a></li>
      <li><strong>GitHub Releases</strong>：<a href="${escapeHtml(releasesUrl)}">${escapeHtml(releasesUrl)}</a>，下载 <code>${escapeHtml(plugin.name)}-&lt;version&gt;.tgz</code> 后在插件管理中安装。</li>
      <li><strong>npm</strong>：<code>npm install ${escapeHtml(plugin.name)}</code> · <a href="${escapeHtml(npmUrl)}">npmjs.com</a></li>
      <li><strong>本地链接</strong>：克隆 <a href="${escapeHtml(repoUrl)}">${escapeHtml(repoUrl)}</a> 后运行 <code>tools\\install.ps1</code>。</li>
    </ul>
  </section>

  <section>
    <h2>版本</h2>
    <table>
      <thead><tr><th>版本</th><th>发布时间（UTC）</th><th>engines.dsh</th></tr></thead>
      <tbody>
${versionRows}
      </tbody>
    </table>
  </section>

  <section>
    <h2>README 摘要</h2>
    <p>面向电网行业的 DSH 插件：Office 文档/表格阅读、编辑与排版；GB/T 4728 Visio 电网专业图（主接线、配电、拓扑、二次、变电站布置、地理接线、潮流、端子排）；CAD 转 PDF（图框分割、图签目录表、批量队列）。仅 Windows，MIT 协议。</p>
    <ul>
      <li>Office：doc/docx/xls/xlsx/ppt/pptx 读取与导出 PDF，Word 结构化编辑、排版、模板与批量任务。</li>
      <li>Visio：GB/T 4728-2022 符号库，九类图分派，接线校验与结构化错误码。</li>
      <li>CAD：COM / ODA / PDF 三级链路，自动分割、合并、图号排序与图纸目录表。</li>
      <li>分发：市场索引 / GitHub Releases / npm / 本地链接四种渠道；发布资产附 SHA256SUMS。</li>
    </ul>
  </section>

  <footer>
    索引 URL：<code>${escapeHtml(indexUrl)}</code> · 页面生成时间 ${escapeHtml(generated)} UTC · MIT License · ${escapeHtml(plugin.author ?? '')}
  </footer>
</main>
</body>
</html>
`;
}

/* ────────────────────────────── CLI ────────────────────────────── */

function usage() {
  return `用法：node tools/build-index.mjs [选项]

  --out <path>       输出 index.json 路径（默认 dist/index.json）
  --versions <list>  逗号分隔的版本列表（默认 docs/releases.json 的记录 ∪ suite/package.json 版本）
  --pages <dir>      同时生成 gh-pages 目录：index.json + icon.svg + index.html
  -h, --help         显示本帮助

环境变量：
  SOURCE_DATE_EPOCH  固定发布时间的整数 Unix 秒（CI 用 tag 提交时间，保证多 job 产物一致）
`;
}

function parseArgs(argv) {
  const options = {};
  const readValue = (name, index) => {
    const value = argv[index];
    if (value === undefined || value.startsWith('--')) throw new Error(`${name} 缺少取值`);
    return value;
  };
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === '-h' || arg === '--help') {
      options.help = true;
    } else if (arg === '--out') {
      options.out = readValue('--out', (index += 1));
    } else if (arg.startsWith('--out=')) {
      options.out = arg.slice('--out='.length);
    } else if (arg === '--versions') {
      options.versions = readValue('--versions', (index += 1));
    } else if (arg.startsWith('--versions=')) {
      options.versions = arg.slice('--versions='.length);
    } else if (arg === '--pages') {
      options.pagesDir = readValue('--pages', (index += 1));
    } else if (arg.startsWith('--pages=')) {
      options.pagesDir = arg.slice('--pages='.length);
    } else {
      throw new Error(`未知参数：${arg}\n${usage()}`);
    }
  }
  return options;
}

async function main(argv = process.argv.slice(2)) {
  const options = parseArgs(argv);
  if (options.help) {
    process.stdout.write(usage());
    return;
  }
  const result = await buildIndex(options);
  for (const warning of result.warnings) process.stderr.write(`[build-index] 警告：${warning}\n`);
  process.stdout.write(`[build-index] index.json → ${result.indexPath}\n`);
  if (result.pagesDir !== null) {
    process.stdout.write(`[build-index] pages → ${result.pagesDir}（index.json + icon.svg + index.html）\n`);
  }
  process.stdout.write(
    `${JSON.stringify({
      ok: true,
      index: result.indexPath,
      pages: result.pagesDir,
      indexUrl: result.indexUrl,
      plugin: result.index.plugins[0].name,
      versions: result.index.plugins[0].versions.map((entry) => entry.version),
    })}\n`,
  );
}

const invokedPath = process.argv[1] === undefined ? '' : pathToFileURL(resolve(process.argv[1])).href;
const isDirectRun = invokedPath.toLowerCase() === import.meta.url.toLowerCase();
if (isDirectRun) {
  main().catch((error) => {
    process.stderr.write(`[build-index] 失败：${error instanceof Error ? error.message : String(error)}\n`);
    process.exitCode = 1;
  });
}
