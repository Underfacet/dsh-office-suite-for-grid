/**
 * 电网图数据模型校验 + 接线连通性分析 + 符号库装载（docs/P2-设计.md §4.1 / §4.2、docs/P3-设计.md §4.2）
 *
 * 职责（Host 侧，绘制前执行，不依赖 Visio）：
 * 1. `validateGridModel(kind, model, options?)`：八类电网图（P2：singleLine / distribution /
 *    networkTopology；P3：protection / substationPlan / geoNetwork / powerFlow / terminalStrip）
 *    的结构校验（error 级 → E2021/E2022，阻断绘制）+ 并查集接线校验（warnings/infos，不阻断）。
 *    P2 三类沿用统一底座（nodes/links）与类型专属字段（buses/bays/transformers/busTies、
 *    sources/feeders、rings）；P3 五类按 docs/P3-设计.md §4.4 各自 schema 校验
 *    （protection：circuits/loops/terminalStrips；substationPlan：plot/zones/items；
 *    geoNetwork：viewport/stations/lines/labels；powerFlow：levels/buses/nodes/transformers/flows；
 *    terminalStrip：strips/terminals）。
 *    `cableRoute` 本期未实现（P4）：校验与符号装载均返回 E2024（不静默）。
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

/** P2 图类型（visio-grid.ps1）。 */
export const P2_GRID_KINDS = Object.freeze(['singleLine', 'distribution', 'networkTopology']);

/** P3 图类型（visio-layouts.ps1）。 */
export const P3_GRID_KINDS = Object.freeze(['protection', 'substationPlan', 'geoNetwork', 'powerFlow', 'terminalStrip']);

/** 已定稿 schema 但本期未实现的图类型（P4：cableRoute）。 */
export const GRID_KINDS_PENDING = Object.freeze(['cableRoute']);

/** 支持的图类型（P2 三类 + P3 五类）。 */
export const GRID_KINDS = Object.freeze([...P2_GRID_KINDS, ...P3_GRID_KINDS]);

/** 支持的纸型（与引擎 visio-grid.ps1 / visio-layouts.ps1 一致）。 */
export const GRID_PAPERS = Object.freeze(['A4-landscape', 'A4-portrait', 'A3-landscape', 'A3-portrait']);

/** links[].kind 取值（P3 §4.3 增 'flow'：潮流箭头线）。 */
export const GRID_LINK_KINDS = Object.freeze(['line', 'feeder', 'fiber', 'cable', 'flow']);

/** protection.loops[].kind 取值（触点/信号映射）。 */
export const GRID_LOOP_KINDS = Object.freeze(['signal', 'contact', 'wire']);

/** powerFlow.flows[].style 取值。 */
export const GRID_FLOW_STYLES = Object.freeze(['solid', 'dashed', 'double']);

/** substationPlan.zones[].items[] 图元类型。 */
export const GRID_PLAN_PRIMITIVES = Object.freeze(['rect']);

/** terminalStrip.terminals[].type 取值。 */
export const GRID_TERMINAL_TYPES = Object.freeze(['normal', 'test', 'shorting', 'ground']);

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
  // ── P3 增量（warn/info，不阻断）──
  CIRCUIT_OPEN: 'GRID_CIRCUIT_OPEN',
  PROTECTION_LOOP: 'GRID_PROTECTION_LOOP',
  POWERFLOW_ISLAND: 'GRID_POWERFLOW_ISLAND',
  TRANSFORMER_SAME_LEVEL: 'GRID_TRANSFORMER_SAME_LEVEL',
  TRANSFORMER_CROSS_LEVEL: 'GRID_TRANSFORMER_CROSS_LEVEL',
  PLAN_OUT_OF_PLOT: 'GRID_PLAN_OUT_OF_PLOT',
  GEO_OUTSIDE_VIEWPORT: 'GRID_GEO_OUTSIDE_VIEWPORT',
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
 * @param {'singleLine'|'distribution'|'networkTopology'|'protection'|'substationPlan'|'geoNetwork'|'powerFlow'|'terminalStrip'} kind 图类型
 *        （'cableRoute' 定稿但未实现 → E2024）
 * @param {object} model 对应图类型数据模型（P2 §4.1 / P3 §4.4）
 * @param {object} [options]
 * @param {Map<string,object>|Record<string,object>|object[]} [options.symbols] 符号库（提供时校验
 *        symbolRef 存在性与端子名，缺失报 E2022 / 非法端子报 E2021）；缺省只做结构校验
 * @param {number} [options.maxSymbols=2000] 符号实例上限
 * @returns {{ok:true,result:{model:object,warnings:object[],infos:object[],validations:object[]}}
 *          |{ok:false,error:{code:string,message:string,hint?:string}}}
 */
export function validateGridModel(kind, model, options = {}) {
  if (kind === 'cableRoute') {
    return failResult(
      'E2024',
      'cableRoute（电缆路径图）本期 P3 未实现，schema 已定稿（docs/P3-设计.md §4.4⑥）',
      `P3 可用图类型：${GRID_KINDS.join(' / ')}；cableRoute 计划 P4 提供`,
    );
  }
  if (!GRID_KINDS.includes(kind)) {
    return failResult('E2021', `不支持的图类型 kind=${JSON.stringify(kind)}（支持：${GRID_KINDS.join(' / ')}）`, `kind 取 ${GRID_KINDS.join(' / ')}`);
  }
  if (!isPlainObject(model)) {
    return failResult('E2021', '电网图模型必须是 JSON 对象', '模型结构见 docs/P2-设计.md §4.1 与 docs/P3-设计.md §4.4');
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

  // ── P3 五类：独立 schema 校验分支（P2 三类的既有流程完全不变）──
  if (P3_GRID_KINDS.includes(kind)) {
    return runP3Validation(kind, model, {
      errors,
      warnings,
      infos,
      error,
      warn,
      note,
      symbolMap: toSymbolMap(options.symbols),
      maxSymbols: Number.isFinite(options.maxSymbols) && options.maxSymbols > 0 ? options.maxSymbols : DEFAULT_GRID_MAX_SYMBOLS,
    });
  }

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

// ───────────────────────────── P3 五类校验（§4.2 / §4.4） ─────────────────────────────

/** P3 图类型中文名（错误消息用）。 */
const P3_KIND_LABELS = Object.freeze({
  protection: '二次回路/继保原理图',
  substationPlan: '变电站总平面布置示意图',
  geoNetwork: '地理接线图',
  powerFlow: '潮流图/系统接线图',
  terminalStrip: '端子排图',
});

function isPositiveInteger(value) {
  return typeof value === 'number' && Number.isInteger(value) && value > 0;
}

/** 简易并查集（P3 接线提示共用；warn/info 不阻断）。 */
function makeUnionFind(ids) {
  const parent = new Map();
  for (const id of ids) parent.set(id, id);
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
  return { find, union };
}

/**
 * 解析 powerFlow.flows[].direction："A→B" / "A->B" / "A-B"。
 * 返回 [fromText, toText] 或 null（格式非法）。
 */
function splitDirection(value) {
  const text = String(value ?? '').trim();
  if (!text) return null;
  for (const separator of ['→', '->']) {
    const at = text.indexOf(separator);
    if (at > 0) {
      const left = text.slice(0, at).trim();
      const right = text.slice(at + separator.length).trim();
      return left && right ? [left, right] : null;
    }
  }
  const at = text.indexOf('-');
  if (at > 0 && text.indexOf('-', at + 1) < 0) {
    const left = text.slice(0, at).trim();
    const right = text.slice(at + 1).trim();
    return left && right ? [left, right] : null;
  }
  return null;
}

/**
 * P3 五类图模型校验（被 validateGridModel 在 meta/title 校验之后分派）。
 * 结构错误 → E2021；符号引用错误 → E2022；接线提示 → warnings/infos（不阻断）。
 * @param {'protection'|'substationPlan'|'geoNetwork'|'powerFlow'|'terminalStrip'} kind
 * @param {object} model
 * @param {{errors:string[],warnings:object[],infos:object[],error:Function,warn:Function,note:Function,
 *          symbolMap:Map<string,object>|null,maxSymbols:number}} ctx
 */
function runP3Validation(kind, model, ctx) {
  const { errors, warnings, infos, error, warn, note, symbolMap, maxSymbols } = ctx;
  const elements = new Map();
  const symbolErrors = [];
  let instanceCount = 0;

  const register = (element, role, where) => {
    const id = element?.id;
    if (!isNonEmptyString(id)) {
      error(`${where} 缺少 id（非空字符串）`);
      return false;
    }
    if (elements.has(id)) {
      error(`id 重复：${id}（已在 ${elements.get(id).where} 定义，${where} 再次使用）`);
      return false;
    }
    elements.set(id, {
      id,
      role,
      where,
      symbolRef: isNonEmptyString(element.symbolRef) ? element.symbolRef : null,
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
  const checkSymbolRef = (element, where) => {
    if (!isNonEmptyString(element.symbolRef)) {
      error(`${where} 缺少 symbolRef（形如 gb4728.7.QF）`);
    } else if (symbolMap && !symbolMap.has(element.symbolRef)) {
      symbolErrors.push(`${where}（${element.id ?? '?'}）引用的符号不存在：${element.symbolRef}`);
    }
    instanceCount += 1;
  };
  const checkConnectEndpoint = (ref, where, roles, what) => {
    if (!isNonEmptyString(ref)) return null;
    const parsed = parseEndpoint(ref);
    const target = elements.get(parsed.id);
    if (!target) {
      error(`${where} 引用了不存在的${what}：${ref}`);
      return null;
    }
    if (roles && !roles.includes(target.role)) {
      error(`${where} 引用的 ${parsed.id} 是 ${target.role}，不是${what}`);
      return null;
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
    return { id: parsed.id, terminal: parsed.terminal, element: target };
  };
  const readArray = (key, required, what) => {
    const value = model[key];
    if (value == null) {
      if (required) error(`缺少必填数组 ${key}（${what}）`);
      return [];
    }
    if (!Array.isArray(value)) {
      error(`${key} 必须为数组`);
      return [];
    }
    if (required && value.length === 0) error(`${key} 不能为空数组（${what}）`);
    const list = [];
    value.forEach((item, index) => {
      if (!isPlainObject(item)) error(`${key}[${index}] 必须为对象`);
      else list.push(item);
    });
    return list;
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
  const checkOptionalString = (value, where, name = 'label') => {
    if (value != null && typeof value !== 'string') error(`${where}.${name} 必须为字符串`);
  };

  switch (kind) {
    // ── ① protection：二次回路/继保原理图（§4.4①）──
    case 'protection': {
      const circuits = readArray('circuits', true, '回路');
      const loops = readArray('loops', false);
      const strips = readArray('terminalStrips', false);

      circuits.forEach((circuit, index) => {
        const where = `circuits[${index}]`;
        if (!isNonEmptyString(circuit.id)) {
          error(`${where}.id 必须为非空字符串（回路 id）`);
        } else {
          register(circuit, 'circuit', where);
        }
        checkOptionalString(circuit.label, where, 'label');
        if (circuit.voltage != null && !isNonEmptyString(circuit.voltage)) {
          error(`${where}.voltage 必须为非空字符串（如 DC110V）`);
        }
        const devices = Array.isArray(circuit.devices) ? circuit.devices : null;
        if (!devices || devices.length === 0) {
          error(`${where}.devices 不能为空数组`);
          return;
        }
        devices.forEach((device, deviceIndex) => {
          const deviceWhere = `${where}.devices[${deviceIndex}]`;
          if (!isPlainObject(device)) {
            error(`${deviceWhere} 必须为对象`);
            return;
          }
          if (!register(device, 'device', deviceWhere)) return;
          checkSymbolRef(device, deviceWhere);
          checkOptionalString(device.label, deviceWhere, 'label');
        });
      });

      loops.forEach((loop, index) => {
        const where = `loops[${index}]`;
        if (isNonEmptyString(loop.id)) register(loop, 'loop', where);
        if (!isNonEmptyString(loop.from)) error(`${where}.from 必须为设备 id 或 "设备id.端子id"`);
        if (!isNonEmptyString(loop.to)) error(`${where}.to 必须为设备 id 或 "设备id.端子id"`);
        if (loop.kind != null && !GRID_LOOP_KINDS.includes(loop.kind)) {
          error(`${where}.kind 必须为 ${GRID_LOOP_KINDS.join('|')}，收到 ${JSON.stringify(loop.kind)}`);
        }
        checkOptionalString(loop.label, where, 'label');
        checkConnectEndpoint(loop.from, `${where}.from`, ['device'], '设备');
        checkConnectEndpoint(loop.to, `${where}.to`, ['device'], '设备');
      });

      const terminalKeys = new Set();
      strips.forEach((strip, index) => {
        const where = `terminalStrips[${index}]`;
        if (!register(strip, 'strip', where)) return;
        checkOptionalString(strip.label, where, 'label');
        const side = strip.side == null ? 'top' : strip.side;
        if (side !== 'top' && side !== 'bottom') {
          error(`${where}.side 必须为 top|bottom，收到 ${JSON.stringify(strip.side)}`);
        }
        const terminals = Array.isArray(strip.terminals) ? strip.terminals : null;
        if (!terminals || terminals.length === 0) {
          error(`${where}.terminals 不能为空数组`);
          return;
        }
        terminals.forEach((terminal, terminalIndex) => {
          const terminalWhere = `${where}.terminals[${terminalIndex}]`;
          if (!isPlainObject(terminal)) {
            error(`${terminalWhere} 必须为对象`);
            return;
          }
          if (!isPositiveInteger(terminal.no)) {
            error(`${terminalWhere}.no 必须为正整数（端子号）`);
          } else {
            const key = `${strip.id}\u0000${terminal.no}`;
            if (terminalKeys.has(key)) error(`${terminalWhere}.no 重复：${strip.id} 的端子号 ${terminal.no}`);
            terminalKeys.add(key);
          }
          if (!isNonEmptyString(terminal.ref)) error(`${terminalWhere}.ref 必须为非空回路编号`);
          checkOptionalString(terminal.desc, terminalWhere, 'desc');
          instanceCount += 1;
        });
      });

      // 接线提示：单设备回路 = 回路开断（无法串接），不阻断绘制
      circuits.forEach((circuit, index) => {
        const devices = asArray(circuit.devices).filter(isPlainObject);
        if (devices.length === 1) {
          warn(
            GRID_VALIDATION_CODES.CIRCUIT_OPEN,
            `保护回路只有 1 个设备，无法形成串接：${isNonEmptyString(circuit.id) ? circuit.id : `circuits[${index}]`}`,
            { circuit: isNonEmptyString(circuit.id) ? circuit.id : null },
          );
        }
      });
      // 触点映射 info（端点已存在时）
      loops.forEach((loop, index) => {
        if (!isNonEmptyString(loop.from) || !isNonEmptyString(loop.to)) return;
        const from = parseEndpoint(loop.from);
        const to = parseEndpoint(loop.to);
        if (!elements.has(from.id) || !elements.has(to.id)) return;
        note(
          GRID_VALIDATION_CODES.PROTECTION_LOOP,
          `触点/信号映射：${loop.from} → ${loop.to}${isNonEmptyString(loop.label) ? `（${loop.label}）` : ''}`,
          { loop: isNonEmptyString(loop.id) ? loop.id : null },
        );
      });
      break;
    }

    // ── ② substationPlan：变电站总平面布置示意图（§4.4②）──
    case 'substationPlan': {
      const plot = model.plot;
      const hasPlot = isPlainObject(plot);
      let plotW = 0;
      let plotH = 0;
      if (!hasPlot) {
        error('缺少必填对象 plot（{ widthMm, heightMm, gridMm? }）');
      } else {
        if (!isPosNumber(plot.widthMm)) error('plot.widthMm 必须为正数（mm）');
        else plotW = plot.widthMm;
        if (!isPosNumber(plot.heightMm)) error('plot.heightMm 必须为正数（mm）');
        else plotH = plot.heightMm;
        if (plot.gridMm != null && !isPosNumber(plot.gridMm)) error('plot.gridMm 必须为正数（mm）');
      }
      if (isPlainObject(model.meta) && model.meta.compass != null && typeof model.meta.compass !== 'boolean') {
        error('meta.compass 必须为布尔值');
      }

      const zones = readArray('zones', true, '设备区');
      const zonesById = new Map();
      zones.forEach((zone, index) => {
        const where = `zones[${index}]`;
        if (!register(zone, 'zone', where)) return;
        zonesById.set(zone.id, zone);
        checkOptionalString(zone.label, where, 'label');
        if (!isFiniteNumber(zone.xMm)) error(`${where}.xMm 必须为有限数（mm）`);
        if (!isFiniteNumber(zone.yMm)) error(`${where}.yMm 必须为有限数（mm）`);
        if (!isPosNumber(zone.widthMm)) error(`${where}.widthMm 必须为正数（mm）`);
        if (!isPosNumber(zone.heightMm)) error(`${where}.heightMm 必须为正数（mm）`);
        if (hasPlot && isFiniteNumber(zone.xMm) && isFiniteNumber(zone.yMm) && isPosNumber(zone.widthMm) && isPosNumber(zone.heightMm)) {
          const outside = zone.xMm < 0 || zone.yMm < 0
            || zone.xMm + zone.widthMm > plotW + 1e-6 || zone.yMm + zone.heightMm > plotH + 1e-6;
          if (outside) warn(GRID_VALIDATION_CODES.PLAN_OUT_OF_PLOT, `设备区超出 plot 范围：${zone.id}`, { zone: zone.id });
        }
      });

      const items = readArray('items', false);
      items.forEach((item, index) => {
        const where = `items[${index}]`;
        if (!register(item, 'item', where)) return;
        checkOptionalString(item.label, where, 'label');
        const primitive = item.primitive == null ? null : item.primitive;
        if (primitive != null && !GRID_PLAN_PRIMITIVES.includes(primitive)) {
          error(`${where}.primitive 必须为 ${GRID_PLAN_PRIMITIVES.join('|')}，收到 ${JSON.stringify(item.primitive)}`);
        }
        const isRect = primitive === 'rect';
        if (isRect) {
          if (!isPosNumber(item.widthMm)) error(`${where}.widthMm 必须为正数（rect 图元）`);
          if (!isPosNumber(item.heightMm)) error(`${where}.heightMm 必须为正数（rect 图元）`);
        } else {
          checkSymbolRef(item, where);
        }
        if (item.rot != null && ![0, 90, 180, 270].includes(item.rot)) {
          error(`${where}.rot 必须为 0|90|180|270`);
        }
        if (item.dashed != null && typeof item.dashed !== 'boolean') error(`${where}.dashed 必须为布尔值`);
        const hasX = item.xMm != null;
        const hasY = item.yMm != null;
        if (hasX !== hasY) error(`${where}.xMm/yMm 必须同时提供或同时省略`);
        if (hasX && !isFiniteNumber(item.xMm)) error(`${where}.xMm 必须为有限数（mm）`);
        if (hasY && !isFiniteNumber(item.yMm)) error(`${where}.yMm 必须为有限数（mm）`);
        if (isRect && (!hasX || !hasY)) error(`${where}.xMm/yMm 不能省略（rect 图元需绝对坐标）`);
        if (item.zone != null) {
          if (!isNonEmptyString(item.zone)) error(`${where}.zone 必须为设备区 id`);
          else if (!zonesById.has(item.zone)) error(`${where} 引用了不存在的设备区：${item.zone}`);
        }
        if (hasPlot && hasX && hasY && isFiniteNumber(item.xMm) && isFiniteNumber(item.yMm)) {
          if (item.xMm < 0 || item.yMm < 0 || item.xMm > plotW || item.yMm > plotH) {
            warn(GRID_VALIDATION_CODES.PLAN_OUT_OF_PLOT, `设备/图元超出 plot 范围：${item.id}`, { item: item.id });
          }
        }
      });

      const links = readArray('links', false);
      links.forEach((link, index) => {
        const where = `links[${index}]`;
        if (!isNonEmptyString(link.from)) error(`${where}.from 必须为图元/设备 id`);
        if (!isNonEmptyString(link.to)) error(`${where}.to 必须为图元/设备 id`);
        checkOptionalString(link.label, where, 'label');
        checkConnectEndpoint(link.from, `${where}.from`, ['item'], '图元/设备');
        checkConnectEndpoint(link.to, `${where}.to`, ['item'], '图元/设备');
      });
      break;
    }

    // ── ③ geoNetwork：地理接线图（§4.4③）──
    case 'geoNetwork': {
      const viewport = model.viewport;
      let hasViewport = false;
      let west = 0;
      let south = 0;
      let east = 0;
      let north = 0;
      if (viewport != null) {
        if (!isPlainObject(viewport)) {
          error('viewport 必须为对象（{ west, south, east, north }）');
        } else {
          hasViewport = true;
          for (const key of ['west', 'south', 'east', 'north']) {
            if (!isFiniteNumber(viewport[key])) {
              error(`viewport.${key} 必须为有限数（经纬度）`);
              hasViewport = false;
            }
          }
          if (hasViewport) {
            west = viewport.west;
            south = viewport.south;
            east = viewport.east;
            north = viewport.north;
            if (!(west < east)) error(`viewport 必须满足 west < east（收到 ${west}/${east}）`);
            if (!(south < north)) error(`viewport 必须满足 south < north（收到 ${south}/${north}）`);
          }
        }
      }
      const boundsValid = hasViewport && west < east && south < north;

      const stations = readArray('stations', true, '厂站');
      const stationsById = new Map();
      stations.forEach((station, index) => {
        const where = `stations[${index}]`;
        if (!register(station, 'station', where)) return;
        stationsById.set(station.id, station);
        checkSymbolRef(station, where);
        checkOptionalString(station.label, where, 'label');
        checkVoltage(station.voltage, `${where}.voltage`);
        const hasLng = station.lng != null;
        const hasLat = station.lat != null;
        const hasX = station.xMm != null;
        const hasY = station.yMm != null;
        if (hasLng !== hasLat) error(`${where}.lng/lat 必须同时提供或同时省略`);
        if (hasX !== hasY) error(`${where}.xMm/yMm 必须同时提供或同时省略`);
        if (hasLng && !isFiniteNumber(station.lng)) error(`${where}.lng 必须为有限数`);
        if (hasLat && !isFiniteNumber(station.lat)) error(`${where}.lat 必须为有限数`);
        if (hasX && !isFiniteNumber(station.xMm)) error(`${where}.xMm 必须为有限数（mm）`);
        if (hasY && !isFiniteNumber(station.yMm)) error(`${where}.yMm 必须为有限数（mm）`);
        if (hasViewport) {
          if (!hasLng || !hasLat) error(`${where} 在提供 viewport 时必须给出 lng/lat`);
        } else if (!hasX || !hasY) {
          error(`${where} 缺少坐标：无 viewport 时需提供 xMm/yMm（平面坐标）`);
        }
        if (boundsValid && hasLng && hasLat && isFiniteNumber(station.lng) && isFiniteNumber(station.lat)) {
          if (station.lng < west || station.lng > east || station.lat < south || station.lat > north) {
            warn(GRID_VALIDATION_CODES.GEO_OUTSIDE_VIEWPORT, `厂站超出 viewport 经纬范围：${station.id}`, { station: station.id });
          }
        }
      });

      const lines = readArray('lines', false);
      lines.forEach((line, index) => {
        const where = `lines[${index}]`;
        if (isNonEmptyString(line.id)) register(line, 'line', where);
        if (!isNonEmptyString(line.from)) error(`${where}.from 必须为厂站 id`);
        if (!isNonEmptyString(line.to)) error(`${where}.to 必须为厂站 id`);
        checkConnectEndpoint(line.from, `${where}.from`, ['station'], '厂站');
        checkConnectEndpoint(line.to, `${where}.to`, ['station'], '厂站');
        checkVoltage(line.voltage, `${where}.voltage`);
        checkOptionalString(line.label, where, 'label');
        if (line.points != null) {
          if (!Array.isArray(line.points) || line.points.length === 0) {
            error(`${where}.points 必须为非空数组（[lng,lat] 或 ["站id", null]）`);
          } else {
            line.points.forEach((point, pointIndex) => {
              const pointWhere = `${where}.points[${pointIndex}]`;
              if (!Array.isArray(point) || point.length < 2) {
                error(`${pointWhere} 必须为 [lng,lat] 或 ["站id", null]`);
                return;
              }
              if (typeof point[0] === 'string') {
                if (!isNonEmptyString(point[0])) error(`${pointWhere}[0] 不能为空字符串`);
                else if (!stationsById.has(point[0])) error(`${pointWhere} 引用了不存在的厂站：${point[0]}`);
                if (point[1] != null) error(`${pointWhere} 站引用点的第二项必须为 null`);
              } else if (!isFiniteNumber(point[0]) || !isFiniteNumber(point[1])) {
                error(`${pointWhere} 必须为 [lng,lat]（有限数）或 ["站id", null]`);
              }
            });
          }
        }
      });

      const labels = readArray('labels', false);
      labels.forEach((label, index) => {
        const where = `labels[${index}]`;
        if (isNonEmptyString(label.id)) register(label, 'label', where);
        if (!isNonEmptyString(label.text)) error(`${where}.text 必须为非空字符串`);
        const hasLng = label.lng != null;
        const hasLat = label.lat != null;
        const hasX = label.xMm != null;
        const hasY = label.yMm != null;
        if (hasLng !== hasLat) error(`${where}.lng/lat 必须同时提供或同时省略`);
        if (hasX !== hasY) error(`${where}.xMm/yMm 必须同时提供或同时省略`);
        if (hasLng && !isFiniteNumber(label.lng)) error(`${where}.lng 必须为有限数`);
        if (hasLat && !isFiniteNumber(label.lat)) error(`${where}.lat 必须为有限数`);
        if (hasX && !isFiniteNumber(label.xMm)) error(`${where}.xMm 必须为有限数（mm）`);
        if (hasY && !isFiniteNumber(label.yMm)) error(`${where}.yMm 必须为有限数（mm）`);
        if (!hasLng && !hasX) error(`${where} 缺少坐标：需提供 lng/lat 或 xMm/yMm`);
      });
      break;
    }

    // ── ④ powerFlow：潮流图/系统接线图（§4.4④）──
    case 'powerFlow': {
      const levels = readArray('levels', true, '电压等级');
      const buses = readArray('buses', true, '母线');
      const nodes = readArray('nodes', false);
      const transformers = readArray('transformers', false);
      const flows = readArray('flows', false);
      const links = readArray('links', false);

      const levelsById = new Map();
      levels.forEach((level, index) => {
        const where = `levels[${index}]`;
        if (!register(level, 'level', where)) return;
        levelsById.set(level.id, level);
        if (!isPosNumber(level.voltage)) error(`${where}.voltage 必须为正数（kV）`);
        checkOptionalString(level.label, where, 'label');
      });

      const elementLevel = new Map();
      buses.forEach((bus, index) => {
        const where = `buses[${index}]`;
        if (!register(bus, 'bus', where)) return;
        if (!isNonEmptyString(bus.level)) error(`${where}.level 必须为电压等级 id`);
        else if (!levelsById.has(bus.level)) error(`${where} 引用了不存在的电压等级：${bus.level}`);
        else elementLevel.set(bus.id, bus.level);
        checkOptionalString(bus.label, where, 'label');
        checkVoltage(bus.voltage, `${where}.voltage`);
        instanceCount += 1;
      });

      nodes.forEach((node, index) => {
        const where = `nodes[${index}]`;
        if (!register(node, 'node', where)) return;
        checkSymbolRef(node, where);
        if (node.level != null) {
          if (!isNonEmptyString(node.level)) error(`${where}.level 必须为电压等级 id`);
          else if (!levelsById.has(node.level)) error(`${where} 引用了不存在的电压等级：${node.level}`);
          else elementLevel.set(node.id, node.level);
        }
        if (node.tags != null && !isPlainObject(node.tags)) error(`${where}.tags 必须为对象`);
        checkOptionalString(node.label, where, 'label');
      });

      transformers.forEach((transformer, index) => {
        const where = `transformers[${index}]`;
        if (!register(transformer, 'transformer', where)) return;
        checkSymbolRef(transformer, where);
        if (!isNonEmptyString(transformer.from)) error(`${where}.from 必须为母线/节点 id`);
        if (!isNonEmptyString(transformer.to)) error(`${where}.to 必须为母线/节点 id`);
        checkConnectEndpoint(transformer.from, `${where}.from`, ['bus', 'node'], '母线/节点');
        checkConnectEndpoint(transformer.to, `${where}.to`, ['bus', 'node'], '母线/节点');
        checkOptionalString(transformer.label, where, 'label');
      });

      flows.forEach((flow, index) => {
        const where = `flows[${index}]`;
        if (isNonEmptyString(flow.id)) register(flow, 'flow', where);
        if (!isNonEmptyString(flow.from)) error(`${where}.from 必须为母线/节点/变压器 id`);
        if (!isNonEmptyString(flow.to)) error(`${where}.to 必须为母线/节点/变压器 id`);
        checkConnectEndpoint(flow.from, `${where}.from`, ['bus', 'node', 'transformer'], '母线/节点/变压器');
        checkConnectEndpoint(flow.to, `${where}.to`, ['bus', 'node', 'transformer'], '母线/节点/变压器');
        if (flow.kind != null && flow.kind !== 'flow') error(`${where}.kind 必须为 flow`);
        checkOptionalString(flow.label, where, 'label');
        if (flow.style != null && !GRID_FLOW_STYLES.includes(flow.style)) {
          error(`${where}.style 必须为 ${GRID_FLOW_STYLES.join('|')}，收到 ${JSON.stringify(flow.style)}`);
        }
        if (flow.direction != null) {
          const parts = isNonEmptyString(flow.direction) ? splitDirection(flow.direction) : null;
          if (!parts) {
            error(`${where}.direction 必须为 "A→B" 或 "A-B"，收到 ${JSON.stringify(flow.direction)}`);
          } else {
            const fromId = isNonEmptyString(flow.from) ? parseEndpoint(flow.from).id : null;
            const toId = isNonEmptyString(flow.to) ? parseEndpoint(flow.to).id : null;
            const forward = parts[0] === fromId && parts[1] === toId;
            const reverse = parts[0] === toId && parts[1] === fromId;
            if (!forward && !reverse) error(`${where}.direction 与 from/to 不一致：${flow.direction}`);
            else if (!elements.has(parts[0]) || !elements.has(parts[1])) error(`${where}.direction 引用了不存在的元素：${flow.direction}`);
          }
        }
      });

      links.forEach((link, index) => {
        const where = `links[${index}]`;
        if (isNonEmptyString(link.id)) register(link, 'link', where);
        if (!isNonEmptyString(link.from)) error(`${where}.from 必须为母线/节点/变压器 id`);
        if (!isNonEmptyString(link.to)) error(`${where}.to 必须为母线/节点/变压器 id`);
        if (link.kind != null && !GRID_LINK_KINDS.includes(link.kind)) {
          error(`${where}.kind 必须为 ${GRID_LINK_KINDS.join('|')}，收到 ${JSON.stringify(link.kind)}`);
        }
        checkOptionalString(link.label, where, 'label');
        checkConnectEndpoint(link.from, `${where}.from`, ['bus', 'node', 'transformer'], '母线/节点/变压器');
        checkConnectEndpoint(link.to, `${where}.to`, ['bus', 'node', 'transformer'], '母线/节点/变压器');
      });

      // 变压器跨层 info / 同层 warn
      transformers.forEach((transformer, index) => {
        const where = `transformers[${index}]`;
        const fromId = isNonEmptyString(transformer.from) ? parseEndpoint(transformer.from).id : null;
        const toId = isNonEmptyString(transformer.to) ? parseEndpoint(transformer.to).id : null;
        const fromLevel = fromId ? elementLevel.get(fromId) : null;
        const toLevel = toId ? elementLevel.get(toId) : null;
        if (!fromLevel || !toLevel) return;
        if (fromLevel === toLevel) {
          warn(
            GRID_VALIDATION_CODES.TRANSFORMER_SAME_LEVEL,
            `变压器 ${isNonEmptyString(transformer.id) ? transformer.id : where} 两端接同一电压等级 ${fromLevel}`,
            { transformer: isNonEmptyString(transformer.id) ? transformer.id : null },
          );
          return;
        }
        const fromVoltage = levelsById.get(fromLevel)?.voltage;
        const toVoltage = levelsById.get(toLevel)?.voltage;
        note(
          GRID_VALIDATION_CODES.TRANSFORMER_CROSS_LEVEL,
          `变压器 ${isNonEmptyString(transformer.id) ? transformer.id : where} 跨层：${fromLevel}（${fromVoltage ?? '?'}kV）→ ${toLevel}（${toVoltage ?? '?'}kV）`,
          { transformer: isNonEmptyString(transformer.id) ? transformer.id : null },
        );
      });

      // 潮流连通性（并查集）：孤岛 / 多电源 / 无电源标注
      const connectable = new Map();
      for (const element of elements.values()) {
        if (element.role === 'bus' || element.role === 'node' || element.role === 'transformer') connectable.set(element.id, element);
      }
      if (connectable.size) {
        const uf = makeUnionFind([...connectable.keys()]);
        const edges = [];
        for (const flow of flows) {
          const a = isNonEmptyString(flow.from) ? parseEndpoint(flow.from).id : null;
          const b = isNonEmptyString(flow.to) ? parseEndpoint(flow.to).id : null;
          if (a && b && connectable.has(a) && connectable.has(b) && a !== b) edges.push([a, b]);
        }
        for (const link of links) {
          const a = isNonEmptyString(link.from) ? parseEndpoint(link.from).id : null;
          const b = isNonEmptyString(link.to) ? parseEndpoint(link.to).id : null;
          if (a && b && connectable.has(a) && connectable.has(b) && a !== b) edges.push([a, b]);
        }
        for (const transformer of transformers) {
          const a = isNonEmptyString(transformer.from) ? parseEndpoint(transformer.from).id : null;
          const b = isNonEmptyString(transformer.to) ? parseEndpoint(transformer.to).id : null;
          const tid = isNonEmptyString(transformer.id) ? transformer.id : null;
          if (!tid || !connectable.has(tid)) continue;
          if (a && connectable.has(a) && a !== tid) edges.push([a, tid]);
          if (b && connectable.has(b) && b !== tid) edges.push([b, tid]);
        }
        for (const [a, b] of edges) uf.union(a, b);
        const sources = [...connectable.values()].filter((element) => element.tags?.source === true).map((element) => element.id);
        if (!sources.length) {
          note(GRID_VALIDATION_CODES.NO_SOURCE, '模型未标注电源（nodes[].tags.source=true），跳过潮流孤岛判定');
        } else {
          const components = new Map();
          for (const id of connectable.keys()) {
            const root = uf.find(id);
            if (!components.has(root)) components.set(root, []);
            components.get(root).push(id);
          }
          const ordered = [...components.values()].sort((left, right) => String(left[0]).localeCompare(String(right[0])));
          for (const ids of ordered) {
            const sourcesIn = ids.filter((id) => sources.includes(id));
            if (!sourcesIn.length) {
              warn(GRID_VALIDATION_CODES.POWERFLOW_ISLAND, `潮流孤岛：${summarizeIds(ids)}（无电源）`, { nodes: ids.slice(0, 20) });
            } else if (sourcesIn.length >= 2) {
              note(GRID_VALIDATION_CODES.MULTI_SOURCE, `多电源并列运行：${summarizeIds(sourcesIn)}`, { nodes: sourcesIn.slice(0, 20) });
            }
          }
        }
      }
      break;
    }

    // ── ⑤ terminalStrip：端子排图（§4.4⑤）──
    case 'terminalStrip': {
      if (isPlainObject(model.meta) && model.meta.terminalsPerPage != null) {
        if (!isPositiveInteger(model.meta.terminalsPerPage) || model.meta.terminalsPerPage > 40) {
          error(`meta.terminalsPerPage 必须为 1–40 的整数，收到 ${JSON.stringify(model.meta.terminalsPerPage)}`);
        }
      }
      const strips = readArray('strips', true, '端子排');
      strips.forEach((strip, index) => {
        const where = `strips[${index}]`;
        if (!register(strip, 'strip', where)) return;
        checkOptionalString(strip.label, where, 'label');
        const columns = strip.columns == null ? 1 : strip.columns;
        if (columns !== 1 && columns !== 2) {
          error(`${where}.columns 必须为 1|2，收到 ${JSON.stringify(strip.columns)}`);
        }
        const terminals = Array.isArray(strip.terminals) ? strip.terminals : null;
        if (!terminals || terminals.length === 0) {
          error(`${where}.terminals 不能为空数组`);
          return;
        }
        const nos = new Set();
        const linksToCheck = [];
        terminals.forEach((terminal, terminalIndex) => {
          const terminalWhere = `${where}.terminals[${terminalIndex}]`;
          if (!isPlainObject(terminal)) {
            error(`${terminalWhere} 必须为对象`);
            return;
          }
          if (!isPositiveInteger(terminal.no)) {
            error(`${terminalWhere}.no 必须为正整数（端子号）`);
          } else if (nos.has(terminal.no)) {
            error(`${terminalWhere}.no 重复：${strip.id} 端子号 ${terminal.no}`);
          } else {
            nos.add(terminal.no);
          }
          if (!isNonEmptyString(terminal.ref)) error(`${terminalWhere}.ref 必须为非空回路编号`);
          checkOptionalString(terminal.desc, terminalWhere, 'desc');
          const type = terminal.type == null ? 'normal' : terminal.type;
          if (!GRID_TERMINAL_TYPES.includes(type)) {
            error(`${terminalWhere}.type 必须为 ${GRID_TERMINAL_TYPES.join('|')}，收到 ${JSON.stringify(terminal.type)}`);
          }
          if (terminal.link != null) {
            if (!isPositiveInteger(terminal.link)) error(`${terminalWhere}.link 必须为同排端子号（正整数）或 null`);
            else linksToCheck.push({ where: terminalWhere, no: terminal.no, link: terminal.link });
          }
          instanceCount += 1;
        });
        linksToCheck.forEach(({ where: terminalWhere, no, link }) => {
          if (link === no) error(`${terminalWhere}.link 不能指向自身（端子号 ${no}）`);
          else if (!nos.has(link)) error(`${terminalWhere}.link 引用了本排不存在的端子号：${link}`);
        });
      });
      break;
    }

    default:
      error(`P3 校验缺少 kind 分支：${kind}`);
  }

  if (instanceCount > maxSymbols) error(`符号实例数 ${instanceCount} 超过上限 ${maxSymbols}（配置项 gridMaxSymbols）`);

  if (errors.length) {
    return failResult(
      'E2021',
      formatErrors(`${P3_KIND_LABELS[kind] ?? kind}模型结构校验未通过`, errors),
      `模型结构见 docs/P3-设计.md §4.4（${kind}）；符号结构见 suite/symbols/README.md`,
    );
  }
  if (symbolErrors.length) {
    return failResult(
      'E2022',
      formatErrors(`${P3_KIND_LABELS[kind] ?? kind}符号引用校验未通过`, symbolErrors),
      'symbolRef 需存在于符号库（suite/symbols/**/*.json），可先用 loadSymbols 校验',
    );
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
 * @param {'singleLine'|'distribution'|'networkTopology'|'protection'|'substationPlan'|'geoNetwork'|'powerFlow'|'terminalStrip'|null} kind 图类型（可为 null；cableRoute → E2024）
 * @param {string} [symbolsDir=DEFAULT_SYMBOLS_DIR] 符号库根目录（suite/symbols）
 * @returns {{ok:true,result:{symbols:Record<string,object>,entries:object[],warnings:string[],infos:string[],count:number,symbolsDir:string,forKind:string|null}}
 *          |{ok:false,error:{code:string,message:string,hint?:string}}}
 */
export function loadSymbols(kind, symbolsDir = DEFAULT_SYMBOLS_DIR) {
  if (kind === 'cableRoute') {
    return failResult(
      'E2024',
      'cableRoute（电缆路径图）本期 P3 未实现，schema 已定稿（docs/P3-设计.md §4.4⑥）',
      `P3 可用图类型：${GRID_KINDS.join(' / ')}；cableRoute 计划 P4 提供`,
    );
  }
  if (kind != null && !GRID_KINDS.includes(kind)) {
    return failResult('E2021', `不支持的图类型 kind=${JSON.stringify(kind)}（支持：${GRID_KINDS.join(' / ')}）`, `kind 取 ${GRID_KINDS.join(' / ')}`);
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
