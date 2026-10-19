/**
 * D1：电网图数据模型校验 / 接线连通性 / 符号库装载测试（docs/P2-设计.md §4.1/§4.2/§9）
 *
 * 覆盖：
 * - loadSymbols：17 个正式符号全部通过 validateSymbol、id 唯一、compose 闭合、跳过 P1 示例文件；
 * - validateGridModel：三类模型合法/非法用例（E2021 结构、E2022 符号引用）；
 * - 接线校验：孤岛 / 多电源 / 开关开断 / 母线分段 / 接地刀闸合位 / 环网闭环 / 馈线末端开路；
 * - expandGridLinks：默认连线生成与端点对去重。
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import {
  GRID_VALIDATION_CODES,
  expandGridLinks,
  loadSymbols,
  validateGridModel,
} from '../suite/host/ops/grid-model.js';

/** 17 个正式符号 id（§4.5 清单）。 */
const EXPECTED_SYMBOL_IDS = [
  'gb4728.3.BUS',
  'gb4728.7.QF',
  'gb4728.7.QS',
  'gb4728.7.QSE',
  'gb4728.7.QL',
  'gb4728.7.FU',
  'gb4728.6.T2',
  'gb4728.6.G',
  'gb4728.6.TA',
  'gb4728.6.TV',
  'gb4728.7.F',
  'gb4728.4.C',
  'gb4728.4.L',
  'gb4728.2.PE',
  'gb4728.3.J',
  'gb4728.7.RING-MAIN',
  'gb4728.9.NODE',
];

const codesOf = (list) => list.map((item) => item.code);

/** 干净的 three-bus 主接线模型：单电源、开关全合、母线经母联与主变连通。 */
function makeSingleLine(overrides = {}) {
  return {
    kind: 'singleLine',
    title: '110kV 测试变主接线图',
    meta: { paper: 'A3-landscape', standard: 'GB/T 4728-2022' },
    buses: [
      { id: 'B1', label: '110kV I 母', voltage: 110, section: 'I' },
      { id: 'B2', label: '110kV II 母', voltage: 110, section: 'II' },
    ],
    nodes: [{ id: 'G1', symbolRef: 'gb4728.6.G', label: '1# 发电机', tags: { source: true } }],
    links: [{ id: 'K1', from: 'G1', to: 'B1', kind: 'line', label: '发电机引线' }],
    bays: [
      {
        id: 'X1',
        bus: 'B1',
        label: '1# 主变高压侧',
        devices: [
          { id: 'QS11', symbolRef: 'gb4728.7.QS', state: 'closed' },
          { id: 'QF11', symbolRef: 'gb4728.7.QF', state: 'closed' },
        ],
      },
      {
        id: 'X2',
        bus: 'B2',
        label: '2# 主变高压侧',
        devices: [
          { id: 'QS21', symbolRef: 'gb4728.7.QS', state: 'closed' },
          { id: 'QF21', symbolRef: 'gb4728.7.QF', state: 'closed' },
        ],
      },
    ],
    transformers: [{ id: 'T1', symbolRef: 'gb4728.6.T2', label: '1# 主变', fromBay: 'X1', toBay: 'X2', voltage: [110, 10] }],
    busTies: [{ id: 'BT1', between: ['B1', 'B2'], devices: [{ id: 'QF0', symbolRef: 'gb4728.7.QF', state: 'closed' }] }],
    ...overrides,
  };
}

function makeTopology(overrides = {}) {
  return {
    kind: 'networkTopology',
    title: '通信光纤环网',
    nodes: [
      { id: 'C1', symbolRef: 'gb4728.9.NODE', label: '中心站' },
      { id: 'C2', symbolRef: 'gb4728.9.NODE', label: '变电站 A' },
      { id: 'C3', symbolRef: 'gb4728.9.NODE', label: '变电站 B' },
      { id: 'C4', symbolRef: 'gb4728.9.NODE', label: '变电站 C' },
    ],
    rings: [{ id: 'RING1', label: '光纤环 1', nodes: ['C1', 'C2', 'C3', 'C4'] }],
    links: [
      { id: 'L1', from: 'C1', to: 'C2', kind: 'fiber' },
      { id: 'L2', from: 'C2', to: 'C3', kind: 'fiber' },
      { id: 'L3', from: 'C3', to: 'C4', kind: 'fiber' },
      { id: 'L4', from: 'C4', to: 'C1', kind: 'fiber' },
    ],
    ...overrides,
  };
}

function makeDistribution(overrides = {}) {
  return {
    kind: 'distribution',
    title: '10kV 配电系统图',
    sources: [{ id: 'S1', symbolRef: 'gb4728.3.BUS', label: '110kV 站 10kV 母线' }],
    nodes: [
      { id: 'R1', symbolRef: 'gb4728.7.RING-MAIN', label: '1# 环网柜' },
      { id: 'T1', symbolRef: 'gb4728.6.T2', label: '1# 台区' },
    ],
    feeders: [
      { id: 'F1', label: '馈线 101', from: 'S1', to: 'R1', kind: 'feeder' },
      { id: 'F2', label: '馈线 102', from: 'R1', to: 'T1', kind: 'feeder' },
    ],
    ...overrides,
  };
}

// ───────────────────────────── 符号库 ─────────────────────────────

test('loadSymbols：17 个正式符号全部通过 validateSymbol 且 id 唯一', () => {
  const loaded = loadSymbols('singleLine');
  assert.equal(loaded.ok, true, loaded.ok ? '' : loaded.error.message);
  const { symbols, entries } = loaded.result;
  assert.equal(entries.length, 17);
  assert.equal(loaded.result.count, 17);
  for (const id of EXPECTED_SYMBOL_IDS) {
    assert.ok(symbols[id], `缺少符号 ${id}`);
    assert.ok(Array.isArray(symbols[id].primitives) && symbols[id].primitives.length > 0, `${id} 无图元`);
    assert.ok(Array.isArray(symbols[id].terminals) && symbols[id].terminals.length > 0, `${id} 无端子`);
    assert.equal(symbols[id].standard.clause, '', `${id} 的 clause 必须留空（未核对标准条款）`);
  }
  // P1 示例文件被跳过（其 id 与正式 QF 重复、clause 为示意值）
  assert.ok(loaded.result.infos.some((text) => text.includes('example-qf.json')));
  // 组合符号的 compose 引用闭合（RING-MAIN → BUS/QL×3/FU）
  assert.deepEqual(symbols['gb4728.7.RING-MAIN'].compose.map((entry) => entry.symbolRef).sort(), ['gb4728.3.BUS', 'gb4728.7.FU', 'gb4728.7.QL', 'gb4728.7.QL', 'gb4728.7.QL'].sort());
});

test('loadSymbols：目录不存在 → E1023；kind 非法 → E2021', () => {
  const missing = loadSymbols('singleLine', 'runtime/tmp/__no_such_symbols__');
  assert.equal(missing.ok, false);
  assert.equal(missing.error.code, 'E1023');

  const badKind = loadSymbols('singleLine2');
  assert.equal(badKind.ok, false);
  assert.equal(badKind.error.code, 'E2021');
});

// ───────────────────────────── 三类模型合法用例 ─────────────────────────────

test('validateGridModel：单线图合法模型 → ok（无 warning）', () => {
  const loaded = loadSymbols('singleLine');
  const result = validateGridModel('singleLine', makeSingleLine(), { symbols: loaded.result.symbols });
  assert.equal(result.ok, true, result.ok ? '' : result.error.message);
  assert.deepEqual(result.result.warnings, []);
  assert.equal(result.result.model.kind, 'singleLine');
  assert.deepEqual(result.result.validations, [...result.result.warnings, ...result.result.infos]);
  const infoCodes = codesOf(result.result.infos);
  assert.ok(infoCodes.includes(GRID_VALIDATION_CODES.AUTO_LINKS));
});

test('validateGridModel：配电网合法模型 → ok', () => {
  const loaded = loadSymbols('distribution');
  const result = validateGridModel('distribution', makeDistribution(), { symbols: loaded.result.symbols });
  assert.equal(result.ok, true, result.ok ? '' : result.error.message);
  assert.deepEqual(result.result.warnings, []);
});

test('validateGridModel：拓扑图合法模型 → 环闭环 info', () => {
  const loaded = loadSymbols('networkTopology');
  const result = validateGridModel('networkTopology', makeTopology(), { symbols: loaded.result.symbols });
  assert.equal(result.ok, true, result.ok ? '' : result.error.message);
  assert.ok(codesOf(result.result.infos).includes(GRID_VALIDATION_CODES.RING_CLOSED));
  assert.deepEqual(result.result.warnings, []);
});

// ───────────────────────────── 结构校验（E2021） ─────────────────────────────

test('validateGridModel：kind 与模型不一致 / kind 非法 → E2021', () => {
  const mismatch = validateGridModel('distribution', makeSingleLine());
  assert.equal(mismatch.ok, false);
  assert.equal(mismatch.error.code, 'E2021');
  assert.match(mismatch.error.message, /不一致/);

  const badKind = validateGridModel('singleline', makeSingleLine());
  assert.equal(badKind.ok, false);
  assert.equal(badKind.error.code, 'E2021');
});

test('validateGridModel：单线图缺少 buses/bays → E2021', () => {
  const result = validateGridModel('singleLine', { kind: 'singleLine', nodes: [], links: [] });
  assert.equal(result.ok, false);
  assert.equal(result.error.code, 'E2021');
  assert.match(result.error.message, /buses/);
  assert.match(result.error.message, /bays/);
});

test('validateGridModel：id 全局重复 → E2021', () => {
  const model = makeSingleLine();
  model.nodes.push({ id: 'QS11', symbolRef: 'gb4728.7.QS', label: '重复 id' });
  const result = validateGridModel('singleLine', model);
  assert.equal(result.ok, false);
  assert.equal(result.error.code, 'E2021');
  assert.match(result.error.message, /id 重复：QS11/);
});

test('validateGridModel：links 引用不存在的元素 / bay 作为端点 → E2021', () => {
  const unknown = makeSingleLine({ links: [{ id: 'K9', from: 'G1', to: 'NOPE', kind: 'line' }] });
  const result = validateGridModel('singleLine', unknown);
  assert.equal(result.ok, false);
  assert.equal(result.error.code, 'E2021');
  assert.match(result.error.message, /不存在的元素：NOPE/);

  const bayAsEndpoint = makeSingleLine({ links: [{ id: 'K9', from: 'X1', to: 'B1', kind: 'line' }] });
  const second = validateGridModel('singleLine', bayAsEndpoint);
  assert.equal(second.ok, false);
  assert.equal(second.error.code, 'E2021');
  assert.match(second.error.message, /不能连接到 X1/);
});

test('validateGridModel：state / voltage / bus 引用非法 → E2021', () => {
  const badState = makeSingleLine({
    bays: [{ id: 'X1', bus: 'B1', devices: [{ id: 'QS11', symbolRef: 'gb4728.7.QS', state: 'half' }] }],
  });
  assert.match(validateGridModel('singleLine', badState).error.message, /state 必须为/);

  const badVoltage = makeSingleLine({ buses: [{ id: 'B1', label: '母线', voltage: -110 }] });
  assert.match(validateGridModel('singleLine', badVoltage).error.message, /voltage 必须为正数/);

  const badBusRef = makeSingleLine({ bays: [{ id: 'X1', bus: 'B9', devices: [{ id: 'QS11', symbolRef: 'gb4728.7.QS' }] }] });
  assert.match(validateGridModel('singleLine', badBusRef).error.message, /引用了不存在的母线：B9/);
});

test('validateGridModel：symbolRef 不存在 → E2022；不注入符号库时只做结构校验', () => {
  const loaded = loadSymbols('singleLine');
  const model = makeSingleLine();
  model.nodes[0].symbolRef = 'gb4728.7.XX';
  const withSymbols = validateGridModel('singleLine', model, { symbols: loaded.result.symbols });
  assert.equal(withSymbols.ok, false);
  assert.equal(withSymbols.error.code, 'E2022');
  assert.match(withSymbols.error.message, /gb4728\.7\.XX/);

  const structuralOnly = validateGridModel('singleLine', model);
  assert.equal(structuralOnly.ok, true, structuralOnly.ok ? '' : structuralOnly.error.message);
});

test('validateGridModel：端子不存在 → E2021', () => {
  const loaded = loadSymbols('singleLine');
  const model = makeSingleLine({ links: [{ id: 'K9', from: 'G1.T9', to: 'B1', kind: 'line' }] });
  const result = validateGridModel('singleLine', model, { symbols: loaded.result.symbols });
  assert.equal(result.ok, false);
  assert.equal(result.error.code, 'E2021');
  assert.match(result.error.message, /不存在的端子：G1\.T9/);
});

test('validateGridModel：符号实例超上限 → E2021', () => {
  const result = validateGridModel('singleLine', makeSingleLine(), { maxSymbols: 3 });
  assert.equal(result.ok, false);
  assert.equal(result.error.code, 'E2021');
  assert.match(result.error.message, /超过上限 3/);
});

// ───────────────────────────── 接线校验（warn/info） ─────────────────────────────

test('接线校验：孤立设备 → 孤岛 warning', () => {
  const model = makeSingleLine();
  model.nodes.push({ id: 'J9', symbolRef: 'gb4728.3.J', label: '孤立 T 接点' });
  const result = validateGridModel('singleLine', model);
  assert.equal(result.ok, true, result.ok ? '' : result.error.message);
  const island = result.result.warnings.find((item) => item.code === GRID_VALIDATION_CODES.ISLAND);
  assert.ok(island, '缺少孤岛 warning');
  assert.match(island.message, /孤岛：J9/);
});

test('接线校验：同岛两个电源 → 多电源 info', () => {
  const model = makeSingleLine();
  model.nodes.push({ id: 'G2', symbolRef: 'gb4728.6.G', label: '2# 发电机', tags: { source: true } });
  model.links.push({ id: 'K2', from: 'G2', to: 'B1', kind: 'line' });
  const result = validateGridModel('singleLine', model);
  assert.equal(result.ok, true, result.ok ? '' : result.error.message);
  const multi = result.result.infos.find((item) => item.code === GRID_VALIDATION_CODES.MULTI_SOURCE);
  assert.ok(multi, '缺少多电源 info');
  assert.match(multi.message, /G1/);
  assert.match(multi.message, /G2/);
});

test('接线校验：母联开关分位 → 开关分位 info + 母线分段 info + 孤岛 warning', () => {
  // 无变压器桥接的两段母线：母联分位后 B2 成为无电源电气岛
  const model = makeSingleLine({
    buses: [
      { id: 'B1', label: '110kV I 母', voltage: 110, section: 'I' },
      { id: 'B2', label: '110kV II 母', voltage: 110, section: 'II' },
    ],
    transformers: [],
    bays: [
      { id: 'X1', bus: 'B1', devices: [{ id: 'QS11', symbolRef: 'gb4728.7.QS', state: 'closed' }] },
      { id: 'X2', bus: 'B2', devices: [{ id: 'QS21', symbolRef: 'gb4728.7.QS', state: 'closed' }] },
    ],
    busTies: [{ id: 'BT1', between: ['B1', 'B2'], devices: [{ id: 'QF0', symbolRef: 'gb4728.7.QF', state: 'open' }] }],
  });
  const result = validateGridModel('singleLine', model);
  assert.equal(result.ok, true, result.ok ? '' : result.error.message);
  const infoCodes = codesOf(result.result.infos);
  assert.ok(infoCodes.includes(GRID_VALIDATION_CODES.SWITCH_OPEN));
  assert.ok(infoCodes.includes(GRID_VALIDATION_CODES.BUS_SPLIT));
  const island = result.result.warnings.find((item) => item.code === GRID_VALIDATION_CODES.ISLAND);
  assert.ok(island, '缺少孤岛 warning');
  assert.match(island.message, /B2|QS21/);
});

test('接线校验：接地刀闸合位 → warn', () => {
  const model = makeSingleLine();
  model.bays[0].devices.push({ id: 'QSE1', symbolRef: 'gb4728.7.QSE', state: 'closed' });
  const result = validateGridModel('singleLine', model);
  assert.equal(result.ok, true, result.ok ? '' : result.error.message);
  const ground = result.result.warnings.find((item) => item.code === GRID_VALIDATION_CODES.GROUND_CLOSED);
  assert.ok(ground, '缺少接地刀闸合位 warning');
  assert.match(ground.message, /QSE1/);
});

test('接线校验：光纤环缺一段 → RING_OPEN warn；闭环 → RING_CLOSED info', () => {
  const opened = makeTopology({ links: [
    { id: 'L1', from: 'C1', to: 'C2', kind: 'fiber' },
    { id: 'L2', from: 'C2', to: 'C3', kind: 'fiber' },
    { id: 'L3', from: 'C3', to: 'C4', kind: 'fiber' },
  ] });
  const result = validateGridModel('networkTopology', opened);
  assert.equal(result.ok, true, result.ok ? '' : result.error.message);
  const ringOpen = result.result.warnings.find((item) => item.code === GRID_VALIDATION_CODES.RING_OPEN);
  assert.ok(ringOpen, '缺少 RING_OPEN warning');
  assert.match(ringOpen.message, /C4-C1/);
});

test('接线校验：馈线终点无下游且非终端设备 → FEEDER_OPEN_END warn', () => {
  const model = makeDistribution();
  model.nodes = [{ id: 'J1', symbolRef: 'gb4728.3.J', label: '未接负载的 T 接' }];
  model.feeders = [{ id: 'F1', label: '馈线 101', from: 'S1', to: 'J1', kind: 'feeder' }];
  const result = validateGridModel('distribution', model);
  assert.equal(result.ok, true, result.ok ? '' : result.error.message);
  const openEnd = result.result.warnings.find((item) => item.code === GRID_VALIDATION_CODES.FEEDER_OPEN_END);
  assert.ok(openEnd, '缺少 FEEDER_OPEN_END warning');
  assert.match(openEnd.message, /F1/);
});

test('接线校验：无电源标注 → NO_SOURCE info，不逐岛报孤岛', () => {
  const model = makeTopology();
  const result = validateGridModel('networkTopology', model);
  assert.equal(result.ok, true, result.ok ? '' : result.error.message);
  assert.ok(codesOf(result.result.infos).includes(GRID_VALIDATION_CODES.NO_SOURCE));
  assert.deepEqual(result.result.warnings, []);
});

// ───────────────────────────── 默认连线 ─────────────────────────────

test('expandGridLinks：单线图默认连线（母线→间隔首设备、间隔串接、变压器两端、母联）', () => {
  const links = expandGridLinks(makeSingleLine());
  const pairs = links.map((link) => `${link.from}->${link.to}`);
  assert.ok(pairs.includes('B1->QS11'));
  assert.ok(pairs.includes('QS11->QF11'));
  assert.ok(pairs.includes('QF11->T1'));
  assert.ok(pairs.includes('T1->QF21'));
  assert.ok(pairs.includes('B1->QF0'));
  assert.ok(pairs.includes('QF0->B2'));
  assert.ok(links.filter((link) => link.auto).length >= 8);
});

test('expandGridLinks：显式 links 与 feeders 优先，端点对去重', () => {
  const model = makeDistribution();
  model.links = [{ id: 'L1', from: 'S1', to: 'R1', kind: 'cable' }];
  const links = expandGridLinks(model);
  const same = links.filter((link) => (link.from === 'S1' && link.to === 'R1') || (link.from === 'R1' && link.to === 'S1'));
  assert.equal(same.length, 1);
  assert.equal(same[0].kind, 'cable');
  assert.equal(same[0].auto, false);
  assert.ok(links.some((link) => link.from === 'R1' && link.to === 'T1' && link.kind === 'feeder'));
});

test('expandGridLinks：nodes 与 transformers 同 id 同符号 → 别名放行 + info', () => {
  const model = makeSingleLine();
  model.nodes.push({ id: 'T1', symbolRef: 'gb4728.6.T2', label: '1# 主变', voltage: [110, 10] });
  const result = validateGridModel('singleLine', model);
  assert.equal(result.ok, true, result.ok ? '' : result.error.message);
  assert.ok(codesOf(result.result.infos).includes(GRID_VALIDATION_CODES.ID_ALIAS));
});
