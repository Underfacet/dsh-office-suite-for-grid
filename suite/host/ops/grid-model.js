/**
 * 电网图数据模型校验 + 接线连通性分析 + 符号库装载（docs/P2-设计.md §4.1 / §4.2）
 *
 * 职责（Host 侧，绘制前执行，不依赖 Visio）：
 * 1. `validateGridModel(kind, model, options?)`：三类电网图（singleLine / distribution /
 *    networkTopology）统一底座（nodes/links）与类型专属字段（buses/bays/transformers/busTies、
 *    sources/feeders、rings）的结构校验（error 级 → E2021/E2022，阻断绘制）；
 *    以及并查集接线校验（孤岛 / 多电源 / 开关开断 / 接地刀闸合位 / 母线分段 / 环网闭环 /
 *    馈线末端开路 → warnings/infos，不阻断绘制）。
 * 2. `loadSymbols(kind, symbolsDir)`：扫描 `suite/symbols/gb4728.*\*.json`，逐条 `validateSymbol`，
 *    并做 compose 引用闭合/深度检查，返回可直接注入引擎的 `{ [id]: symbol }` 映射。
 * 3. `expandGridLinks(model)`：显式 links + distribution.feeders + 自动连线（母线→间隔首设备、
 *    间隔内串接、变压器两端、母联），按端点对去重（§4.1「默认连线生成规则」）。
 *
 * 统一返回形状：`{ok:true,result}` / `{ok:false,error:{code,message,hint?}}`（错误码见 §8）。
 * 坐标/单位约定：mm，原点左上、x 右 y 下（suite/symbols/README.md §2）。
 */
import { readdirSync, readFileSync } from 'node:fs';
import { basename, dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { validateSymbol } from './symbols.js';

const HERE = dirname(fileURLToPath(import.meta.url));

/** 符号库默认目录（suite/symbols）。 */
export const DEFAULT_SYMBOLS_DIR = resolve(HERE, '..', '..', 'symbols');

/** 支持的图类型。 */
export const GRID_KINDS = Object.freeze(['singleLine', 'distribution', 'networkTopology']);

/** 支持的纸型（与引擎 visio-grid.ps1 一致）。 */
export const GRID_PAPERS = Object.freeze(['A4-landscape', 'A4-portrait', 'A3-landscape', 'A3-portrait']);

/** links[].kind 取值。 */
export const GRID_LINK_KINDS = Object.freeze(['line', 'feeder', 'fiber', 'cable']);

/** 开关类 state 取值。 */
export const GRID_STATES = Object.freeze(['closed', 'open', 'unknown']);

/** 符号实例上限（§7 gridMaxSymbols 默认值）。 */
export const DEFAULT_GRID_MAX_SYMBOLS = 2000;

/** 接线校验（warn/info）代码；engine 与面板按 code 归类展示。 */
export const GRID_VALIDATION_CODES = Object.freeze({
  ISLAND: 'GRID_ISLAND',
  MULTI_SOURCE: 'GRID_MULTI_SOURCE',
  NO_SOURCE: 'GRID_NO_SOURCE',
  SWITCH_OPEN: 'GRID_SWITCH_OPEN',
  GROUND_CLOSED: 'GRID_GROUND_CLOSED',
  BUS_SPLIT: 'GRID_BUS_SPLIT',
  RING_CLOSED: 'GRID_RING_CLOSED',
  RING_OPEN: 'GRID_RING_OPEN',
  FEEDER_OPEN_END: 'GRID_FEEDER_OPEN_END',
  AUTO_LINKS: 'GRID_AUTO_LINKS',
  ID_ALIAS: 'GRID_ID_ALIAS',
});

/** 可作为连线端点的角色（母线/设备/电源/变压器）。busTie 是容器，其 devices 才导电。 */
const CONNECTABLE_ROLES = new Set(['node', 'bus', 'device', 'transformer', 'source']);

/** 末端设备符号：馈线终点为这些类型时不视为「末端开路」。 */
const TERMINAL_LOAD_REFS = new Set(['gb4728.6.T2', 'gb4728.6.G', 'gb4728.7.RING-MAIN', 'gb4728.9.NODE']);

// ───────────────────────────── 小工具 ─────────────────────────────

function isPlainObject(value) {
  return Boolean(value) && typeof value === 'object' && !Array.isArray(value);
}

function isFiniteNumber(value) {
  return typeof value === 'number' && Number.isFinite(value);
}

function isPosNumber(value) {
  return isFiniteNumber(value) && value > 0;
}

function isNonEmptyString(value) {
  return typeof value === 'string' && value.trim().length > 0;
}

function asArray(value) {
  return Array.isArray(value) ? value : [];
}

function failResult(code, message, hint) {
  return { ok: false, error: { code, message, ...(hint ? { hint } : {}) } };
}

function formatErrors(title, errors) {
  const head = errors.slice(0, 6).join('；');
  const tail = errors.length > 6 ? `；…另有 ${errors.length - 6} 项` : '';
  return `${title}（${errors.length} 项）：${head}${tail}`;
}

function summarizeIds(ids) {
  const list = ids.slice(0, 5).join('、');
  return ids.length > 5 ? `${list} 等 ${ids.length} 个` : list;
}

/** 把 symbols（Map / {id:symbol} / symbol[]）归一化为 Map<id, symbol>；无输入返回 null。 */
function toSymbolMap(symbols) {
  if (!symbols) return null;
  const map = new Map();
  if (symbols instanceof Map) {
    for (const [key, value] of symbols) map.set(String(key), value);
    return map;
  }
  if (Array.isArray(symbols)) {
    for (const symbol of symbols) if (isNonEmptyString(symbol?.id)) map.set(symbol.id, symbol);
    return map;
  }
  if (typeof symbols === 'object') {
    for (const [key, value] of Object.entries(symbols)) if (isPlainObject(value)) map.set(String(key), value);
    return map;
  }
  return null;
}

// ───────────────────────────── 默认连线（§4.1） ─────────────────────────────

/**
 * 生成模型默认连线（不含显式 links / feeders）。
 * 规则：母线→bay 首设备；bay 内设备按序串接；变压器两端接 fromBay/toBay 末设备；busTies 接两段母线。
 * @returns {Array<{from:string,to:string,kind:'line'}>}
 */
function buildAutoLinks(model) {
  const autos = [];
  const bays = asArray(model.bays).filter(isPlainObject);
  const baysById = new Map(bays.filter((bay) => isNonEmptyString(bay.id)).map((bay) => [bay.id, bay]));
  const lastDeviceId = (bay) => {
    if (!bay) return null;
    const devices = asArray(bay.devices).filter(isPlainObject);
    const last = devices[devices.length - 1];
    return last && isNonEmptyString(last.id) ? last.id : null;
  };
  for (const bay of bays) {
    const devices = asArray(bay.devices).filter(isPlainObject);
    if (isNonEmptyString(bay.bus) && devices.length && isNonEmptyString(devices[0].id)) {
      autos.push({ from: bay.bus, to: devices[0].id, kind: 'line' });
    }
    for (let index = 0; index + 1 < devices.length; index++) {
      const from = devices[index];
      const to = devices[index + 1];
      if (isNonEmptyString(from.id) && isNonEmptyString(to.id)) autos.push({ from: from.id, to: to.id, kind: 'line' });
    }
  }
  for (const transformer of asArray(model.transformers).filter(isPlainObject)) {
    if (!isNonEmptyString(transformer.id)) continue;
    const fromLast = lastDeviceId(baysById.get(transformer.fromBay));
    const toLast = lastDeviceId(baysById.get(transformer.toBay));
    if (fromLast) autos.push({ from: fromLast, to: transformer.id, kind: 'line' });
    if (toLast) autos.push({ from: transformer.id, to: toLast, kind: 'line' });
  }
  for (const busTie of asArray(model.busTies).filter(isPlainObject)) {
    const between = asArray(busTie.between);
    if (between.length !== 2) continue;
    const devices = asArray(busTie.devices).filter((device) => isPlainObject(device) && isNonEmptyString(device.id));
    if (!devices.length) {
      autos.push({ from: between[0], to: between[1], kind: 'line' });
      continue;
    }
    autos.push({ from: between[0], to: devices[0].id, kind: 'line' });
    for (let index = 0; index + 1 < devices.length; index++) autos.push({ from: devices[index].id, to: devices[index + 1].id, kind: 'line' });
    autos.push({ from: devices[devices.length - 1].id, to: between[1], kind: 'line' });
  }
  return autos;
}

/**
 * 展开模型连线：显式 links → distribution.feeders → 自动连线；端点对（无向）相同只保留一条。
 * @param {object} model §4.1 模型
 * @returns {Array<{id?:string,from:string,to:string,kind:string,label?:string,auto:boolean}>}
 */
export function expandGridLinks(model) {
  if (!isPlainObject(model)) return [];
  const result = [];
  const seen = new Set();
  const pairKey = (a, b) => (a <= b ? `${a}\u0000${b}` : `${b}\u0000${a}`);
  const add = (link) => {
    if (!isPlainObject(link) || !isNonEmptyString(link.from) || !isNonEmptyString(link.to)) return false;
    const key = pairKey(link.from, link.to);
    if (seen.has(key)) return false;
    seen.add(key);
    result.push(link);
    return true;
  };
  for (const link of asArray(model.links)) {
    if (!isPlainObject(link)) continue;
    add({ ...link, kind: isNonEmptyString(link.kind) ? link.kind : 'line', auto: false });
  }
  for (const feeder of asArray(model.feeders)) {
    if (!isPlainObject(feeder)) continue;
    add({ id: isNonEmptyString(feeder.id) ? feeder.id : undefined, from: feeder.from, to: feeder.to, kind: 'feeder', label: feeder.label, auto: false });
  }
  for (const auto of buildAutoLinks(model)) add({ ...auto, auto: true });
  return result;
}

// ───────────────────────────── 模型校验 ─────────────────────────────

/**
 * 校验电网图模型并做接线分析。
 * @param {'singleLine'|'distribution'|'networkTopology'} kind 图类型
 * @param {object} model §4.1 数据模型
 * @param {object} [options]
 * @param {Map<string,object>|Record<string,object>|object[]} [options.symbols] 符号库（提供时校验
 *        symbolRef 存在性与端子名，缺失报 E2022 / 非法端子报 E2021）；缺省只做结构校验
 * @param {number} [options.maxSymbols=2000] 符号实例上限
 * @returns {{ok:true,result:{model:object,warnings:object[],infos:object[],validations:object[]}}
 *          |{ok:false,error:{code:string,message:string,hint?:string}}}
 */
export function validateGridModel(kind, model, options = {}) {
  if (!GRID_KINDS.includes(kind)) {
    return failResult('E2021', `不支持的图类型 kind=${JSON.stringify(kind)}（支持：${GRID_KINDS.join(' / ')}）`, 'kind 取 singleLine / distribution / networkTopology');
  }
  if (!isPlainObject(model)) {
    return failResult('E2021', '电网图模型必须是 JSON 对象', '模型结构见 docs/P2-设计.md §4.1');
  }
  if (model.kind != null && model.kind !== kind) {
    return failResult('E2021', `模型 kind=${JSON.stringify(model.kind)} 与请求 kind=${kind} 不一致`, '模型内 kind 与调用参数需一致，或省略模型内 kind');
  }

  const errors = [];
  const symbolErrors = [];
  const warnings = [];
  const infos = [];
  const error = (message) => errors.push(message);
  const warn = (code, message, extra) => warnings.push({ level: 'warn', code, message, ...(extra ?? {}) });
  const note = (code, message, extra) => infos.push({ level: 'info', code, message, ...(extra ?? {}) });

  // ── meta / title ──
  const meta = model.meta == null ? null : model.meta;
  if (meta != null && !isPlainObject(meta)) error('meta 必须为对象（{ paper, standard, scale }）');
  if (isPlainObject(meta)) {
    if (meta.paper != null && !GRID_PAPERS.includes(meta.paper)) error(`meta.paper 必须为 ${GRID_PAPERS.join(' / ')}，收到 ${JSON.stringify(meta.paper)}`);
    if (meta.scale != null && !isPosNumber(meta.scale)) error('meta.scale 必须为正数');
    if (meta.standard != null && !isNonEmptyString(meta.standard)) error('meta.standard 必须为非空字符串');
  }
  if (model.title != null && typeof model.title !== 'string') error('title 必须为字符串');

  // ── 数组读取（按类型标 required）──
  const arrayOf = (key, required) => {
    const value = model[key];
    if (value == null) {
      if (required) error(`缺少必填数组 ${key}（${kind} 图要求）`);
      return [];
    }
    if (!Array.isArray(value)) {
      error(`${key} 必须为数组`);
      return [];
    }
    if (required && value.length === 0) error(`${key} 不能为空数组（${kind} 图要求）`);
    const list = [];
    value.forEach((item, index) => {
      if (!isPlainObject(item)) error(`${key}[${index}] 必须为对象`);
      else list.push(item);
    });
    return list;
  };

  const nodes = arrayOf('nodes', false);
  const links = arrayOf('links', false);
  const buses = arrayOf('buses', kind === 'singleLine');
  const bays = arrayOf('bays', kind === 'singleLine');
  const transformers = arrayOf('transformers', false);
  const busTies = arrayOf('busTies', false);
  const sources = arrayOf('sources', kind === 'distribution');
  const feeders = arrayOf('feeders', kind === 'distribution');
  const rings = arrayOf('rings', kind === 'networkTopology');

  // ── 元素注册（id 全局唯一；nodes 与 transformers/sources 同 id 同符号按别名放行）──
  const elements = new Map();
  const aliasIds = new Set();
  const aliasNotes = [];
  let instanceCount = 0;
  const checkSymbolRef = (element, where) => {
    if (!isNonEmptyString(element.symbolRef)) error(`${where} 缺少 symbolRef（形如 gb4728.7.QF）`);
    if (!aliasIds.has(element.id)) instanceCount += 1;
  };
  const checkState = (element, where) => {
    if (element.state == null) return;
    if (!GRID_STATES.includes(element.state)) error(`${where}.state 必须为 ${GRID_STATES.join('|')}，收到 ${JSON.stringify(element.state)}`);
  };
  const checkVoltage = (value, where) => {
    if (value == null) return;
    if (isFiniteNumber(value)) {
      if (!isPosNumber(value)) error(`${where} 必须为正数（kV）`);
      return;
    }
    if (Array.isArray(value)) {
      if (!value.length) error(`${where} 不能为空数组`);
      value.forEach((item, index) => {
        if (!isPosNumber(item)) error(`${where}[${index}] 必须为正数（kV）`);
      });
      return;
    }
    error(`${where} 必须为正数或正数数组（kV）`);
  };
  const register = (element, role, where) => {
    const id = element?.id;
    if (!isNonEmptyString(id)) {
      error(`${where} 缺少 id（非空字符串）`);
      return false;
    }
    if (elements.has(id)) {
      const existing = elements.get(id);
      const aliasable = (role === 'node' && (existing.role === 'transformer' || existing.role === 'source'))
        || (existing.role === 'node' && (role === 'transformer' || role === 'source'));
      if (aliasable && isNonEmptyString(existing.symbolRef) && existing.symbolRef === element.symbolRef) {
        aliasNotes.push({ id, roles: [existing.role, role] });
        aliasIds.add(id);
        return true;
      }
      error(`id 重复：${id}（已在 ${existing.where} 定义，${where} 再次使用）`);
      return false;
    }
    elements.set(id, {
      id,
      role,
      where,
      symbolRef: isNonEmptyString(element.symbolRef) ? element.symbolRef : null,
      state: element.state,
      label: element.label,
      tags: isPlainObject(element.tags) ? element.tags : null,
    });
    return true;
  };
  const parseEndpoint = (ref) => {
    const text = String(ref);
    if (elements.has(text)) return { id: text, terminal: null };
    const dot = text.lastIndexOf('.');
    if (dot > 0) {
      const base = text.slice(0, dot);
      if (elements.has(base)) return { id: base, terminal: text.slice(dot + 1) };
    }
    return { id: text, terminal: null };
  };

  // 母线
  for (let index = 0; index < buses.length; index++) {
    const bus = buses[index];
    const where = `buses[${index}]`;
    if (!register(bus, 'bus', where)) continue;
    if (bus.label != null && typeof bus.label !== 'string') error(`${where}.label 必须为字符串`);
    if (bus.section != null && typeof bus.section !== 'string') error(`${where}.section 必须为字符串`);
    checkVoltage(bus.voltage, `${where}.voltage`);
  }
  // 间隔 + 设备
  for (let index = 0; index < bays.length; index++) {
    const bay = bays[index];
    const where = `bays[${index}]`;
    if (!register(bay, 'bay', where)) continue;
    if (!isNonEmptyString(bay.bus)) error(`${where}.bus 必须为母线 id`);
    if (bay.label != null && typeof bay.label !== 'string') error(`${where}.label 必须为字符串`);
    if (bay.devices == null) {
      error(`${where}.devices 必须为设备数组`);
    } else if (!Array.isArray(bay.devices) || bay.devices.length === 0) {
      error(`${where}.devices 不能为空数组`);
    } else {
      bay.devices.forEach((device, deviceIndex) => {
        const deviceWhere = `${where}.devices[${deviceIndex}]`;
        if (!isPlainObject(device)) {
          error(`${deviceWhere} 必须为对象`);
          return;
        }
        if (!register(device, 'device', deviceWhere)) return;
        checkSymbolRef(device, deviceWhere);
        checkState(device, deviceWhere);
        checkVoltage(device.voltage, `${deviceWhere}.voltage`);
      });
    }
  }
  // 变压器
  for (let index = 0; index < transformers.length; index++) {
    const transformer = transformers[index];
    const where = `transformers[${index}]`;
    if (!register(transformer, 'transformer', where)) continue;
    checkSymbolRef(transformer, where);
    checkState(transformer, where);
    checkVoltage(transformer.voltage, `${where}.voltage`);
    if (!isNonEmptyString(transformer.fromBay)) error(`${where}.fromBay 必须为间隔 id`);
    if (!isNonEmptyString(transformer.toBay)) error(`${where}.toBay 必须为间隔 id`);
  }
  // 母联
  for (let index = 0; index < busTies.length; index++) {
    const busTie = busTies[index];
    const where = `busTies[${index}]`;
    if (!register(busTie, 'busTie', where)) continue;
    checkVoltage(busTie.voltage, `${where}.voltage`);
    if (!Array.isArray(busTie.between) || busTie.between.length !== 2) {
      error(`${where}.between 必须为两个母线 id 的数组`);
    }
    if (busTie.devices == null) {
      // 允许无设备（纯母线连接），不报错
    } else if (!Array.isArray(busTie.devices) || busTie.devices.length === 0) {
      error(`${where}.devices 不能为空数组（省略该字段表示无设备）`);
    } else {
      busTie.devices.forEach((device, deviceIndex) => {
        const deviceWhere = `${where}.devices[${deviceIndex}]`;
        if (!isPlainObject(device)) {
          error(`${deviceWhere} 必须为对象`);
          return;
        }
        if (!register(device, 'device', deviceWhere)) return;
        checkSymbolRef(device, deviceWhere);
        checkState(device, deviceWhere);
      });
    }
  }
  // 电源（distribution）
  for (let index = 0; index < sources.length; index++) {
    const source = sources[index];
    const where = `sources[${index}]`;
    if (!register(source, 'source', where)) continue;
    checkSymbolRef(source, where);
    checkState(source, where);
    if (source.tags != null && !isPlainObject(source.tags)) error(`${where}.tags 必须为对象`);
  }
  // 通用节点
  for (let index = 0; index < nodes.length; index++) {
    const node = nodes[index];
    const where = `nodes[${index}]`;
    if (!register(node, 'node', where)) continue;
    checkSymbolRef(node, where);
    checkState(node, where);
    checkVoltage(node.voltage, `${where}.voltage`);
    if (node.tags != null && !isPlainObject(node.tags)) error(`${where}.tags 必须为对象`);
  }
  // links
  for (let index = 0; index < links.length; index++) {
    const link = links[index];
    const where = `links[${index}]`;
    if (isNonEmptyString(link.id)) register(link, 'link', where);
    if (!isNonEmptyString(link.from)) error(`${where}.from 必须为节点 id 或 "节点id.端子id"`);
    if (!isNonEmptyString(link.to)) error(`${where}.to 必须为节点 id 或 "节点id.端子id"`);
    if (link.kind != null && !GRID_LINK_KINDS.includes(link.kind)) error(`${where}.kind 必须为 ${GRID_LINK_KINDS.join('|')}，收到 ${JSON.stringify(link.kind)}`);
    if (link.label != null && typeof link.label !== 'string') error(`${where}.label 必须为字符串`);
  }
  // feeders（distribution）
  for (let index = 0; index < feeders.length; index++) {
    const feeder = feeders[index];
    const where = `feeders[${index}]`;
    if (isNonEmptyString(feeder.id)) register(feeder, 'feeder', where);
    if (!isNonEmptyString(feeder.from)) error(`${where}.from 必须为节点 id`);
    if (!isNonEmptyString(feeder.to)) error(`${where}.to 必须为节点 id`);
    if (feeder.label != null && typeof feeder.label !== 'string') error(`${where}.label 必须为字符串`);
    if (feeder.kind != null && !GRID_LINK_KINDS.includes(feeder.kind)) error(`${where}.kind 必须为 ${GRID_LINK_KINDS.join('|')}，收到 ${JSON.stringify(feeder.kind)}`);
  }
  // rings（networkTopology）
  for (let index = 0; index < rings.length; index++) {
    const ring = rings[index];
    const where = `rings[${index}]`;
    if (!register(ring, 'ring', where)) continue;
    if (ring.label != null && typeof ring.label !== 'string') error(`${where}.label 必须为字符串`);
    if (!Array.isArray(ring.nodes) || ring.nodes.length < 2) error(`${where}.nodes 必须为至少 2 个节点 id 的数组`);
  }

  // 实例上限（§4.3 / §7 gridMaxSymbols）
  const maxSymbols = Number.isFinite(options.maxSymbols) && options.maxSymbols > 0 ? options.maxSymbols : DEFAULT_GRID_MAX_SYMBOLS;
  if (instanceCount > maxSymbols) error(`符号实例数 ${instanceCount} 超过上限 ${maxSymbols}（配置项 gridMaxSymbols）`);

  // ── 引用校验 ──
  const requireRole = (id, roles, where, what) => {
    if (!isNonEmptyString(id)) return;
    const target = elements.get(id);
    if (!target) {
      error(`${where} 引用了不存在的${what}：${id}`);
      return;
    }
    if (!roles.includes(target.role)) error(`${where} 引用的 ${id} 是 ${target.role}，不是${what}`);
  };
  bays.forEach((bay, index) => requireRole(bay.bus, ['bus'], `bays[${index}].bus`, '母线'));
  transformers.forEach((transformer, index) => {
    requireRole(transformer.fromBay, ['bay'], `transformers[${index}].fromBay`, '间隔');
    requireRole(transformer.toBay, ['bay'], `transformers[${index}].toBay`, '间隔');
  });
  busTies.forEach((busTie, index) => {
    asArray(busTie.between).forEach((busId, busIndex) => requireRole(busId, ['bus'], `busTies[${index}].between[${busIndex}]`, '母线'));
  });

  const symbolMap = toSymbolMap(options.symbols);
  for (const element of elements.values()) {
    if (isNonEmptyString(element.symbolRef) && symbolMap && !symbolMap.has(element.symbolRef)) {
      symbolErrors.push(`${element.where}（${element.id}）引用的符号不存在：${element.symbolRef}`);
    }
  }

  const checkEndpoint = (ref, where) => {
    if (!isNonEmptyString(ref)) return;
    const parsed = parseEndpoint(ref);
    const target = elements.get(parsed.id);
    if (!target) {
      error(`${where} 引用了不存在的元素：${ref}`);
      return;
    }
    if (!CONNECTABLE_ROLES.has(target.role)) {
      error(`${where} 不能连接到 ${parsed.id}（角色 ${target.role}；仅母线/设备/电源/变压器/母联设备可作为端点）`);
      return;
    }
    if (parsed.terminal && symbolMap && target.symbolRef) {
      const symbol = symbolMap.get(target.symbolRef);
      if (symbol) {
        const terminalIds = asArray(symbol.terminals).map((terminal) => terminal?.id).filter(isNonEmptyString);
        if (!terminalIds.includes(parsed.terminal)) {
          error(`${where} 引用了不存在的端子：${ref}（符号 ${target.symbolRef} 端子：${terminalIds.join('/') || '无'}）`);
        }
      }
    }
  };
  links.forEach((link, index) => {
    checkEndpoint(link.from, `links[${index}].from`);
    checkEndpoint(link.to, `links[${index}].to`);
  });
  feeders.forEach((feeder, index) => {
    checkEndpoint(feeder.from, `feeders[${index}].from`);
    checkEndpoint(feeder.to, `feeders[${index}].to`);
  });
  rings.forEach((ring, index) => {
    asArray(ring.nodes).forEach((nodeId, nodeIndex) => {
      const where = `rings[${index}].nodes[${nodeIndex}]`;
      if (!isNonEmptyString(nodeId)) {
        error(`${where} 必须为节点 id`);
        return;
      }
      const target = elements.get(nodeId);
      if (!target) error(`${where} 引用了不存在的节点：${nodeId}`);
      else if (!CONNECTABLE_ROLES.has(target.role)) error(`${where} 引用的 ${nodeId} 是 ${target.role}，不能作为环网节点`);
    });
  });

  if (errors.length) {
    return failResult('E2021', formatErrors('电网图模型结构校验未通过', errors), '模型结构见 docs/P2-设计.md §4.1；符号结构见 suite/symbols/README.md');
  }
  if (symbolErrors.length) {
    return failResult('E2022', formatErrors('电网图符号引用校验未通过', symbolErrors), 'symbolRef 需存在于符号库（suite/symbols/**/*.json），可先用 loadSymbols 校验');
  }

  for (const alias of aliasNotes) {
    note(GRID_VALIDATION_CODES.ID_ALIAS, `id ${alias.id} 同时出现在 ${alias.roles.join(' / ')}，按别名处理（不重复绘制）`, { id: alias.id });
  }

  // ───────────────────── 接线校验（并查集；warn/info 不阻断） ─────────────────────
  const connectable = new Map();
  for (const element of elements.values()) {
    if (CONNECTABLE_ROLES.has(element.role)) connectable.set(element.id, element);
  }
  const merged = expandGridLinks(model);
  const autoCount = merged.filter((link) => link.auto).length;

  const parent = new Map();
  for (const id of connectable.keys()) parent.set(id, id);
  const find = (value) => {
    let root = value;
    while (parent.get(root) !== root) root = parent.get(root);
    let cursor = value;
    while (parent.get(cursor) !== root) {
      const next = parent.get(cursor);
      parent.set(cursor, root);
      cursor = next;
    }
    return root;
  };
  const union = (a, b) => {
    const rootA = find(a);
    const rootB = find(b);
    if (rootA !== rootB) parent.set(rootA, rootB);
  };

  const endpointId = (ref) => parseEndpoint(ref).id;
  const openSwitches = new Set();
  const edgePairs = new Set();
  const conductivePairs = new Set();
  const pairKeyOf = (a, b) => (a <= b ? `${a}\u0000${b}` : `${b}\u0000${a}`);
  for (const link of merged) {
    const a = endpointId(link.from);
    const b = endpointId(link.to);
    if (!connectable.has(a) || !connectable.has(b) || a === b) continue;
    const key = pairKeyOf(a, b);
    if (edgePairs.has(key)) continue;
    edgePairs.add(key);
    const elementA = connectable.get(a);
    const elementB = connectable.get(b);
    if (elementA.state === 'open' || elementB.state === 'open') {
      if (elementA.state === 'open') openSwitches.add(a);
      if (elementB.state === 'open') openSwitches.add(b);
      continue;
    }
    conductivePairs.add(key);
    union(a, b);
  }

  if (openSwitches.size) {
    note(GRID_VALIDATION_CODES.SWITCH_OPEN, `开关分位（按断边处理）：${[...openSwitches].join('、')}`, { switches: [...openSwitches] });
  }

  const components = new Map();
  for (const id of connectable.keys()) {
    const root = find(id);
    if (!components.has(root)) components.set(root, []);
    components.get(root).push(id);
  }
  const sourceIds = [];
  for (const element of connectable.values()) {
    if (element.role === 'source' || element.tags?.source === true) sourceIds.push(element.id);
  }
  if (!sourceIds.length) {
    note(GRID_VALIDATION_CODES.NO_SOURCE, '模型未标注电源（tags.source=true 或 distribution.sources[]），跳过电源连通性判定');
  } else {
    const ordered = [...components.values()].sort((left, right) => String(left[0]).localeCompare(String(right[0])));
    for (const ids of ordered) {
      const sourcesIn = ids.filter((id) => sourceIds.includes(id));
      if (!sourcesIn.length) {
        // 全部为分位开关的连通块是「开关断开后的死端」，不算孤岛噪声
        const onlyOpenSwitches = ids.every((id) => connectable.get(id)?.state === 'open');
        if (onlyOpenSwitches) continue;
        warn(GRID_VALIDATION_CODES.ISLAND, `孤岛：${summarizeIds(ids)}（无电源）`, { nodes: ids.slice(0, 20) });
      } else if (sourcesIn.length >= 2) {
        note(GRID_VALIDATION_CODES.MULTI_SOURCE, `多电源并列运行（闭环）：${summarizeIds(sourcesIn)}`, { nodes: sourcesIn.slice(0, 20) });
      }
    }
  }

  // 母线分段运行
  if (buses.length >= 2) {
    const busGroups = new Map();
    for (const bus of buses) {
      if (!connectable.has(bus.id)) continue;
      const root = find(bus.id);
      if (!busGroups.has(root)) busGroups.set(root, []);
      busGroups.get(root).push(bus.id);
    }
    if (busGroups.size > 1) {
      const groups = [...busGroups.values()];
      note(GRID_VALIDATION_CODES.BUS_SPLIT, `母线分段运行：${groups.map((group) => group.join('/')).join(' 与 ')}`, { groups });
    }
  }

  // 接地刀闸合位（检修态）
  for (const element of connectable.values()) {
    if (element.state !== 'closed') continue;
    if (typeof element.symbolRef === 'string' && element.symbolRef.endsWith('.QSE')) {
      warn(GRID_VALIDATION_CODES.GROUND_CLOSED, `接地刀闸合位（检修态）：${element.id}`, { id: element.id });
    }
  }

  // 光纤环闭环：相邻环节点之间必须有直接导电连线（缺段 → warn；环内经其它节点绕行不算闭环）
  for (const ring of rings) {
    const ringNodes = asArray(ring.nodes).filter((id) => connectable.has(id));
    if (ringNodes.length < 3) continue;
    const openPairs = [];
    for (let index = 0; index < ringNodes.length; index++) {
      const a = ringNodes[index];
      const b = ringNodes[(index + 1) % ringNodes.length];
      if (!conductivePairs.has(pairKeyOf(a, b))) openPairs.push(`${a}-${b}`);
    }
    const ringId = isNonEmptyString(ring.id) ? ring.id : '（未命名环）';
    if (openPairs.length) {
      warn(GRID_VALIDATION_CODES.RING_OPEN, `光纤环未闭环：${openPairs.join('、')}（${ringId}）`, { ring: ringId, pairs: openPairs });
    } else {
      note(GRID_VALIDATION_CODES.RING_CLOSED, `光纤环闭环正常：${ringId}（${ringNodes.length} 节点）`, { ring: ringId });
    }
  }

  // 馈线末端开路（distribution）
  if (kind === 'distribution') {
    for (const feeder of feeders) {
      const targetId = endpointId(feeder.to);
      const target = connectable.get(targetId);
      if (!target) continue;
      const hasDownstream = merged.some((link) => endpointId(link.from) === targetId);
      const isTerminal = target.role === 'source' || (target.symbolRef && TERMINAL_LOAD_REFS.has(target.symbolRef));
      if (!hasDownstream && !isTerminal) {
        warn(GRID_VALIDATION_CODES.FEEDER_OPEN_END, `馈线末端开路：${isNonEmptyString(feeder.id) ? feeder.id : feeder.to}（${targetId} 无下游连接）`, { feeder: isNonEmptyString(feeder.id) ? feeder.id : null });
      }
    }
  }

  if (autoCount) {
    note(GRID_VALIDATION_CODES.AUTO_LINKS, `按母线/间隔/端子规则自动生成 ${autoCount} 条连线（显式 links 优先，端点对去重）`, { count: autoCount });
  }

  return {
    ok: true,
    result: {
      model,
      warnings,
      infos,
      validations: [...warnings, ...infos],
    },
  };
}

// ───────────────────────────── 符号库装载 ─────────────────────────────

/** 递归列出目录下的 *.json（仅一层子目录 + 根目录文件，符号库结构固定）。 */
function listSymbolFiles(dir) {
  const files = [];
  let entries;
  try {
    entries = readdirSync(dir, { withFileTypes: true });
  } catch (err) {
    const error = new Error(`符号库目录不可读：${dir}（${err?.message ?? err}）`);
    error.code = 'E1023';
    throw error;
  }
  for (const entry of entries) {
    if (entry.isFile() && entry.name.toLowerCase().endsWith('.json')) files.push(join(dir, entry.name));
  }
  for (const entry of entries) {
    if (!entry.isDirectory()) continue;
    if (!/^gb4728(\.\d+)?$/i.test(entry.name)) continue;
    let children;
    try {
      children = readdirSync(join(dir, entry.name), { withFileTypes: true });
    } catch {
      continue;
    }
    for (const child of children) {
      if (child.isFile() && child.name.toLowerCase().endsWith('.json')) files.push(join(dir, entry.name, child.name));
    }
  }
  files.sort();
  return files;
}

/** compose 引用闭合 / 深度 / 环检查（§4.5 落地规范 1：递归校验，深度 ≤2）。 */
function checkCompose(symbolMap) {
  const errors = [];
  const depthOf = new Map();
  const visiting = new Set();
  const visit = (id, depth) => {
    if (visiting.has(id)) {
      errors.push(`compose 引用成环：${id}`);
      return depth;
    }
    if (depthOf.has(id)) return depthOf.get(id);
    const symbol = symbolMap.get(id);
    if (!symbol) return 0;
    const compose = Array.isArray(symbol.compose) ? symbol.compose : [];
    visiting.add(id);
    let deepest = 0;
    for (let index = 0; index < compose.length; index++) {
      const entry = compose[index];
      if (!isPlainObject(entry) || !isNonEmptyString(entry.symbolRef)) {
        errors.push(`${id} compose[${index}] 缺少 symbolRef`);
        continue;
      }
      if (entry.symbolRef === id) {
        errors.push(`${id} compose 不能引用自身`);
        continue;
      }
      if (!symbolMap.has(entry.symbolRef)) {
        errors.push(`${id} compose[${index}] 引用的符号不存在：${entry.symbolRef}`);
        continue;
      }
      if (!Array.isArray(entry.at) || entry.at.length < 2 || !isFiniteNumber(entry.at[0]) || !isFiniteNumber(entry.at[1])) {
        errors.push(`${id} compose[${index}].at 必须为 [x,y]（mm）`);
      }
      deepest = Math.max(deepest, visit(entry.symbolRef, depth + 1));
    }
    visiting.delete(id);
    const total = deepest + 1;
    depthOf.set(id, total);
    if (total > 3) errors.push(`${id} compose 引用深度 ${total - 1} 超过上限 2`);
    return total;
  };
  for (const id of symbolMap.keys()) visit(id, 0);
  return errors;
}

/**
 * 装载并校验符号库。
 * 说明：所有 gb4728.* 目录的符号都会被载入（不按 kind 过滤——模型可能引用任意部分符号），
 * `kind` 只做图类型合法性校验并回填到结果，供调用方组装引擎 payload 时标注。
 * 示例文件（`example-*.json`）跳过：P1 断路器示例的 id 与正式 `gb4728.7/QF.json` 重复，
 * 且 `clause` 为示意值（§4.5 落地规范 2）。
 * @param {'singleLine'|'distribution'|'networkTopology'|null} kind 图类型（可为 null）
 * @param {string} [symbolsDir=DEFAULT_SYMBOLS_DIR] 符号库根目录（suite/symbols）
 * @returns {{ok:true,result:{symbols:Record<string,object>,entries:object[],warnings:string[],infos:string[],count:number,symbolsDir:string,forKind:string|null}}
 *          |{ok:false,error:{code:string,message:string,hint?:string}}}
 */
export function loadSymbols(kind, symbolsDir = DEFAULT_SYMBOLS_DIR) {
  if (kind != null && !GRID_KINDS.includes(kind)) {
    return failResult('E2021', `不支持的图类型 kind=${JSON.stringify(kind)}（支持：${GRID_KINDS.join(' / ')}）`, 'kind 取 singleLine / distribution / networkTopology');
  }
  const dir = resolve(typeof symbolsDir === 'string' && symbolsDir.trim() ? symbolsDir.trim() : DEFAULT_SYMBOLS_DIR);
  let files;
  try {
    files = listSymbolFiles(dir);
  } catch (err) {
    return failResult('E1023', String(err?.message ?? err), '检查 suite/symbols 是否随包安装，或传入自定义符号库目录');
  }
  if (!files.length) {
    return failResult('E1023', `符号库为空：${dir}`, '符号库应包含 gb4728.<部分号>/*.json');
  }

  const symbols = {};
  const entries = [];
  const warnings = [];
  const infos = [];
  const errors = [];
  const fileById = new Map();
  for (const file of files) {
    const relative = file.slice(dir.length + 1);
    if (basename(file).toLowerCase().startsWith('example-')) {
      infos.push(`跳过示例符号文件：${relative}（正式符号见同名目录下 <id>.json）`);
      continue;
    }
    let symbol;
    try {
      symbol = JSON.parse(readFileSync(file, 'utf8'));
    } catch (err) {
      errors.push(`${relative}: JSON 解析失败（${err?.message ?? err}）`);
      continue;
    }
    const validation = validateSymbol(symbol);
    if (!validation.valid) {
      errors.push(`${relative}: ${validation.errors.join('；')}`);
      continue;
    }
    const id = symbol.id;
    if (fileById.has(id)) {
      errors.push(`符号 id 重复：${id}（${fileById.get(id)} 与 ${relative}）`);
      continue;
    }
    fileById.set(id, relative);
    symbols[id] = symbol;
    for (const warning of validation.warnings) warnings.push(`${id}: ${warning}`);
    entries.push({
      id,
      file: relative,
      name: isPlainObject(symbol.name) ? symbol.name : null,
      category: symbol.category ?? null,
      part: isPlainObject(symbol.standard) ? symbol.standard.part ?? null : null,
      primitives: Array.isArray(symbol.primitives) ? symbol.primitives.length : 0,
      terminals: Array.isArray(symbol.terminals) ? symbol.terminals.length : 0,
      compose: Array.isArray(symbol.compose) ? symbol.compose.length : 0,
    });
  }

  if (errors.length) {
    return failResult('E2022', formatErrors('符号库校验未通过', errors), '逐个修正 suite/symbols/**/*.json 后重新装载；结构见 suite/symbols/README.md');
  }

  const composeErrors = checkCompose(new Map(entries.map((entry) => [entry.id, symbols[entry.id]])));
  if (composeErrors.length) {
    return failResult('E2022', formatErrors('符号库 compose 引用校验未通过', composeErrors), '组合符号只能引用库内符号，深度 ≤2 且不得成环（suite/symbols/README.md §6.5）');
  }

  return {
    ok: true,
    result: {
      symbols,
      entries,
      warnings,
      infos,
      count: entries.length,
      symbolsDir: dir,
      forKind: kind ?? null,
    },
  };
}
