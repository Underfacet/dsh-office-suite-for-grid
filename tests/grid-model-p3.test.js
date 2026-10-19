/**
 * D1：P3 五类电网图模型校验测试（docs/P3-设计.md §4.2 / §4.4 / §8）
 *
 * 覆盖：
 * - GRID_KINDS 扩展（P2 三类 + P3 五类）与 cableRoute → E2024（校验 + 符号装载）；
 * - protection / substationPlan / geoNetwork / powerFlow / terminalStrip 五类合法模型 +
 *   结构非法（E2021）+ 符号引用非法（E2022）+ 符号实例上限；
 * - P3 接线提示：保护回路开断 / 触点映射 / 潮流孤岛 / 多电源 / 变压器跨层同层 /
 *   设备区或图元超出 plot / 厂站超出 viewport（warn/info，不阻断）。
 */
import test from 'node:test';
import assert from 'node:assert/strict';

import {
  GRID_KINDS,
  GRID_LINK_KINDS,
  GRID_VALIDATION_CODES,
  loadSymbols,
  validateGridModel,
} from '../suite/host/ops/grid-model.js';

const codesOf = (list) => list.map((item) => item.code);

/** 载入一次符号库，五类用例共享（与 Host 实际调用一致：校验时注入符号库）。 */
const SYMBOLS = (() => {
  const loaded = loadSymbols(null);
  assert.equal(loaded.ok, true, loaded.ok ? '' : loaded.error.message);
  return loaded.result.symbols;
})();

// ───────────────────────────── 合法模型工厂（§4.4） ─────────────────────────────

function makeProtection(overrides = {}) {
  return {
    kind: 'protection',
    title: '110kV 线路保护电流回路图（示例）',
    meta: { paper: 'A3-landscape', standard: 'GB/T 4728-2022' },
    circuits: [
      {
        id: 'C1',
        label: '电流回路',
        voltage: 'DC110V',
        devices: [
          { id: 'TAa', symbolRef: 'gb4728.6.TA', label: '1TAa' },
          { id: 'KA1', symbolRef: 'gb4728.7.QL', label: 'KA1' },
          { id: 'X1T', symbolRef: 'gb4728.3.J', label: 'X1:1' },
        ],
      },
      {
        id: 'C2',
        label: '电压回路',
        voltage: 'AC100V',
        devices: [
          { id: 'TVa', symbolRef: 'gb4728.6.TV', label: '1TVa' },
          { id: 'KA2', symbolRef: 'gb4728.4.C', label: 'KA2' },
        ],
      },
    ],
    loops: [{ id: 'L1', from: 'KA1', to: 'KA2', kind: 'signal', label: 'KA1-1' }],
    terminalStrips: [
      { id: 'XT1', label: '端子排 1X', side: 'top', terminals: [{ no: 1, ref: 'A411' }, { no: 2, ref: 'B411' }] },
      { id: 'XT2', label: '端子排 2X', side: 'bottom', terminals: [{ no: 1, ref: 'N411' }] },
    ],
    ...overrides,
  };
}

function makeSubstationPlan(overrides = {}) {
  return {
    kind: 'substationPlan',
    title: '110kV 白云变总平面布置示意图',
    meta: { paper: 'A3-landscape', compass: true },
    plot: { widthMm: 800, heightMm: 600, gridMm: 50 },
    zones: [
      { id: 'Z1', label: '110kV 配电装置区', xMm: 0, yMm: 0, widthMm: 300, heightMm: 200 },
      { id: 'Z2', label: '主变区', xMm: 300, yMm: 0, widthMm: 200, heightMm: 200 },
      { id: 'Z3', label: '综合楼及生活区', xMm: 0, yMm: 200, widthMm: 500, heightMm: 200 },
    ],
    items: [
      { id: 'T1', label: '1# 主变', symbolRef: 'gb4728.6.T2', xMm: 350, yMm: 60, rot: 0 },
      { id: 'RD1', label: '进站道路', xMm: 0, yMm: 400, widthMm: 500, heightMm: 24, primitive: 'rect' },
      { id: 'F1', label: '站区围墙', xMm: 0, yMm: 0, widthMm: 800, heightMm: 600, primitive: 'rect', dashed: true },
    ],
    links: [{ id: 'L1', from: 'T1', to: 'RD1', label: '进站道路' }],
    ...overrides,
  };
}

function makeGeoNetwork(overrides = {}) {
  return {
    kind: 'geoNetwork',
    title: '××供电区 35kV 及以上地理接线示意图',
    meta: { paper: 'A3-landscape' },
    viewport: { west: 112.0, south: 30.0, east: 114.0, north: 32.0 },
    stations: [
      { id: 'S1', label: '白云变', symbolRef: 'gb4728.9.NODE', lng: 113.2, lat: 31.1, voltage: 110 },
      { id: 'S2', label: '金亭变', symbolRef: 'gb4728.9.NODE', lng: 113.7, lat: 31.5, voltage: 110 },
    ],
    lines: [
      {
        id: 'LN1',
        from: 'S1',
        to: 'S2',
        voltage: 110,
        label: '白-金 I 回',
        points: [['S1', null], [113.5, 31.3], ['S2', null]],
      },
    ],
    labels: [{ id: 'LBL1', text: '××工业园', lng: 113.1, lat: 30.9 }],
    ...overrides,
  };
}

function makePowerFlow(overrides = {}) {
  return {
    kind: 'powerFlow',
    title: '××电网 110kV 系统接线潮流图（示例）',
    meta: { paper: 'A3-landscape' },
    levels: [
      { id: 'LV110', voltage: 110, label: '110kV' },
      { id: 'LV35', voltage: 35, label: '35kV' },
    ],
    buses: [
      { id: 'B1', level: 'LV110', label: '白云 110kV 母线' },
      { id: 'B2', level: 'LV35', label: '白云 35kV 母线' },
    ],
    nodes: [
      { id: 'G1', symbolRef: 'gb4728.6.G', label: '金亭电厂', level: 'LV110', tags: { source: true } },
      { id: 'S3', symbolRef: 'gb4728.9.NODE', label: '外部系统', level: 'LV110', tags: { source: true } },
    ],
    transformers: [{ id: 'T1', symbolRef: 'gb4728.6.T2', label: '白 1# 主变', from: 'B1', to: 'B2' }],
    flows: [
      { id: 'F1', from: 'B1', to: 'B2', kind: 'flow', direction: 'B1→B2', label: 'P=85.4MW  Q=12.1Mvar', style: 'solid' },
      { id: 'F2', from: 'G1', to: 'B1', kind: 'flow', direction: 'G1→B1', label: 'P=120MW' },
    ],
    links: [{ id: 'L1', from: 'B1', to: 'S3' }],
    ...overrides,
  };
}

function makeTerminalStrip(overrides = {}) {
  return {
    kind: 'terminalStrip',
    title: '端子排图 1X',
    meta: { paper: 'A4-landscape', terminalsPerPage: 20 },
    strips: [
      {
        id: 'XT1',
        label: '端子排 1X（保护屏）',
        columns: 1,
        terminals: [
          { no: 1, ref: 'A411', desc: '1TAa 电流', type: 'normal', link: null },
          { no: 2, ref: 'B411', desc: '1TAb 电流', type: 'shorting', link: 1 },
          { no: 3, ref: 'N411', desc: '1TAn 中性', type: 'ground', link: null },
          { no: 4, ref: 'C411', desc: '试验端子', type: 'test' },
        ],
      },
    ],
    ...overrides,
  };
}

const check = (kind, model, options = {}) => validateGridModel(kind, model, { symbols: SYMBOLS, ...options });

// ───────────────────────────── 类型清单 / cableRoute ─────────────────────────────

test('GRID_KINDS：P2 三类 + P3 五类共 8 类；GRID_LINK_KINDS 含 flow', () => {
  for (const kind of ['singleLine', 'distribution', 'networkTopology', 'protection', 'substationPlan', 'geoNetwork', 'powerFlow', 'terminalStrip']) {
    assert.ok(GRID_KINDS.includes(kind), `缺少图类型 ${kind}`);
  }
  assert.equal(GRID_KINDS.length, 8);
  assert.ok(GRID_LINK_KINDS.includes('flow'));
});

test('cableRoute：校验与符号装载均 → E2024（不静默）', () => {
  const validated = validateGridModel('cableRoute', { kind: 'cableRoute' });
  assert.equal(validated.ok, false);
  assert.equal(validated.error.code, 'E2024');
  assert.match(validated.error.message, /cableRoute/);

  const loaded = loadSymbols('cableRoute');
  assert.equal(loaded.ok, false);
  assert.equal(loaded.error.code, 'E2024');

  const unknown = validateGridModel('singleline', {});
  assert.equal(unknown.ok, false);
  assert.equal(unknown.error.code, 'E2021');
});

// ───────────────────────────── 五类合法模型 ─────────────────────────────

test('protection：合法模型 → ok（触点映射 info，无 warning）', () => {
  const result = check('protection', makeProtection());
  assert.equal(result.ok, true, result.ok ? '' : result.error.message);
  assert.deepEqual(result.result.warnings, []);
  assert.ok(codesOf(result.result.infos).includes(GRID_VALIDATION_CODES.PROTECTION_LOOP));
  assert.deepEqual(result.result.validations, [...result.result.warnings, ...result.result.infos]);
});

test('substationPlan：合法模型 → ok（绝对坐标 + rect 图元 + links）', () => {
  const result = check('substationPlan', makeSubstationPlan());
  assert.equal(result.ok, true, result.ok ? '' : result.error.message);
  assert.deepEqual(result.result.warnings, []);
});

test('geoNetwork：合法模型 → ok（viewport 投影 + 途经点）', () => {
  const result = check('geoNetwork', makeGeoNetwork());
  assert.equal(result.ok, true, result.ok ? '' : result.error.message);
  assert.deepEqual(result.result.warnings, []);
});

test('geoNetwork：无 viewport 时用 xMm/yMm 平面坐标 → ok', () => {
  const result = check('geoNetwork', makeGeoNetwork({
    viewport: undefined,
    stations: [
      { id: 'S1', label: '白云变', symbolRef: 'gb4728.9.NODE', xMm: 40, yMm: 60 },
      { id: 'S2', label: '金亭变', symbolRef: 'gb4728.9.NODE', xMm: 220, yMm: 140 },
    ],
    lines: [{ id: 'LN1', from: 'S1', to: 'S2', points: [['S1', null], [130, 100], ['S2', null]] }],
    labels: [{ id: 'LBL1', text: '××工业园', xMm: 100, yMm: 180 }],
  }));
  assert.equal(result.ok, true, result.ok ? '' : result.error.message);
});

test('powerFlow：合法模型 → ok（跨层 info + 多电源 info，无 warning）', () => {
  const result = check('powerFlow', makePowerFlow());
  assert.equal(result.ok, true, result.ok ? '' : result.error.message);
  assert.deepEqual(result.result.warnings, []);
  const infoCodes = codesOf(result.result.infos);
  assert.ok(infoCodes.includes(GRID_VALIDATION_CODES.TRANSFORMER_CROSS_LEVEL));
  assert.ok(infoCodes.includes(GRID_VALIDATION_CODES.MULTI_SOURCE));
});

test('terminalStrip：合法模型 → ok（shorting link / ground / test 类型）', () => {
  const result = check('terminalStrip', makeTerminalStrip());
  assert.equal(result.ok, true, result.ok ? '' : result.error.message);
});

// ───────────────────────────── protection 非法 / 提示 ─────────────────────────────

test('protection：缺少 circuits / devices 空 → E2021', () => {
  const missing = check('protection', { kind: 'protection', meta: { paper: 'A3-landscape' } });
  assert.equal(missing.ok, false);
  assert.equal(missing.error.code, 'E2021');
  assert.match(missing.error.message, /circuits/);

  const empty = check('protection', makeProtection({ circuits: [{ id: 'C1', devices: [] }] }));
  assert.equal(empty.ok, false);
  assert.equal(empty.error.code, 'E2021');
  assert.match(empty.error.message, /devices 不能为空数组/);
});

test('protection：设备缺 symbolRef / 端子 ref 空 / 端子号重复 → E2021', () => {
  const badDevice = check('protection', makeProtection({
    circuits: [{ id: 'C1', devices: [{ id: 'D1', label: '无符号' }] }],
  }));
  assert.equal(badDevice.ok, false);
  assert.equal(badDevice.error.code, 'E2021');
  assert.match(badDevice.error.message, /symbolRef/);

  const badRef = check('protection', makeProtection({
    terminalStrips: [{ id: 'XT1', side: 'top', terminals: [{ no: 1, ref: '' }] }],
  }));
  assert.equal(badRef.ok, false);
  assert.equal(badRef.error.code, 'E2021');
  assert.match(badRef.error.message, /ref 必须为非空回路编号/);

  const duplicateNo = check('protection', makeProtection({
    terminalStrips: [{ id: 'XT1', side: 'top', terminals: [{ no: 1, ref: 'A' }, { no: 1, ref: 'B' }] }],
  }));
  assert.equal(duplicateNo.ok, false);
  assert.equal(duplicateNo.error.code, 'E2021');
  assert.match(duplicateNo.error.message, /no 重复/);
});

test('protection：loop 端点 / 端子不存在 → E2021；符号不存在 → E2022', () => {
  const badLoop = check('protection', makeProtection({
    loops: [{ id: 'L9', from: 'KA1', to: 'NOPE', kind: 'signal' }],
  }));
  assert.equal(badLoop.ok, false);
  assert.equal(badLoop.error.code, 'E2021');
  assert.match(badLoop.error.message, /不存在的设备：NOPE/);

  const badTerminal = check('protection', makeProtection({
    loops: [{ id: 'L9', from: 'KA1.T9', to: 'KA2', kind: 'signal' }],
  }));
  assert.equal(badTerminal.ok, false);
  assert.equal(badTerminal.error.code, 'E2021');
  assert.match(badTerminal.error.message, /不存在的端子/);

  const badSymbol = check('protection', makeProtection({
    circuits: [{ id: 'C1', devices: [{ id: 'D1', symbolRef: 'gb4728.7.NOPE' }] }],
    loops: [],
    terminalStrips: [],
  }));
  assert.equal(badSymbol.ok, false);
  assert.equal(badSymbol.error.code, 'E2022');
  assert.match(badSymbol.error.message, /gb4728\.7\.NOPE/);
});

test('protection：单设备回路 → CIRCUIT_OPEN warning（不阻断）', () => {
  const model = makeProtection({
    circuits: [{ id: 'C1', devices: [{ id: 'TAa', symbolRef: 'gb4728.6.TA' }] }],
    loops: [],
    terminalStrips: [],
  });
  const result = check('protection', model);
  assert.equal(result.ok, true, result.ok ? '' : result.error.message);
  const open = result.result.warnings.find((item) => item.code === GRID_VALIDATION_CODES.CIRCUIT_OPEN);
  assert.ok(open, '缺少 CIRCUIT_OPEN warning');
  assert.match(open.message, /C1/);
});

// ───────────────────────────── substationPlan 非法 / 提示 ─────────────────────────────

test('substationPlan：plot 缺失 / zones 非法 → E2021', () => {
  const noPlot = check('substationPlan', { kind: 'substationPlan', zones: [] });
  assert.equal(noPlot.ok, false);
  assert.equal(noPlot.error.code, 'E2021');
  assert.match(noPlot.error.message, /plot/);

  const badZone = check('substationPlan', makeSubstationPlan({
    zones: [{ id: 'Z1', xMm: 0, yMm: 0, widthMm: 0, heightMm: 200 }],
  }));
  assert.equal(badZone.ok, false);
  assert.equal(badZone.error.code, 'E2021');
  assert.match(badZone.error.message, /widthMm 必须为正数/);
});

test('substationPlan：rect 图元缺尺寸 / item 未知 zone / 未知 link 端点 → E2021', () => {
  const badRect = check('substationPlan', makeSubstationPlan({
    items: [{ id: 'RD1', primitive: 'rect', xMm: 0, yMm: 0, heightMm: 24 }],
  }));
  assert.equal(badRect.ok, false);
  assert.equal(badRect.error.code, 'E2021');
  assert.match(badRect.error.message, /widthMm 必须为正数/);

  const badZoneRef = check('substationPlan', makeSubstationPlan({
    items: [{ id: 'T1', symbolRef: 'gb4728.6.T2', xMm: 350, yMm: 60, zone: 'Z9' }],
  }));
  assert.equal(badZoneRef.ok, false);
  assert.equal(badZoneRef.error.code, 'E2021');
  assert.match(badZoneRef.error.message, /不存在的设备区：Z9/);

  const badLink = check('substationPlan', makeSubstationPlan({
    links: [{ id: 'L9', from: 'T1', to: 'NOPE' }],
  }));
  assert.equal(badLink.ok, false);
  assert.equal(badLink.error.code, 'E2021');
  assert.match(badLink.error.message, /不存在的图元\/设备：NOPE/);
});

test('substationPlan：item 超出 plot → PLAN_OUT_OF_PLOT warning（不阻断）', () => {
  const result = check('substationPlan', makeSubstationPlan({
    items: [{ id: 'T9', label: '越界主变', symbolRef: 'gb4728.6.T2', xMm: 900, yMm: 700 }],
    links: [],
  }));
  assert.equal(result.ok, true, result.ok ? '' : result.error.message);
  const out = result.result.warnings.find((item) => item.code === GRID_VALIDATION_CODES.PLAN_OUT_OF_PLOT);
  assert.ok(out, '缺少 PLAN_OUT_OF_PLOT warning');
  assert.match(out.message, /T9/);
});

test('substationPlan：item 符号不存在 → E2022', () => {
  const result = check('substationPlan', makeSubstationPlan({
    items: [{ id: 'T1', symbolRef: 'gb4728.6.NOPE', xMm: 350, yMm: 60 }],
    links: [],
  }));
  assert.equal(result.ok, false);
  assert.equal(result.error.code, 'E2022');
});

// ───────────────────────────── geoNetwork 非法 / 提示 ─────────────────────────────

test('geoNetwork：viewport 非法（west >= east / 非数值）→ E2021', () => {
  const bad = check('geoNetwork', makeGeoNetwork({ viewport: { west: 114, south: 30, east: 112, north: 32 } }));
  assert.equal(bad.ok, false);
  assert.equal(bad.error.code, 'E2021');
  assert.match(bad.error.message, /west < east/);

  const nonNumber = check('geoNetwork', makeGeoNetwork({ viewport: { west: 'x', south: 30, east: 114, north: 32 } }));
  assert.equal(nonNumber.ok, false);
  assert.equal(nonNumber.error.code, 'E2021');
  assert.match(nonNumber.error.message, /viewport\.west/);
});

test('geoNetwork：station 缺坐标 / line 端点未知 / points 非法 → E2021', () => {
  const noCoord = check('geoNetwork', makeGeoNetwork({
    stations: [{ id: 'S1', symbolRef: 'gb4728.9.NODE' }],
    lines: [],
    labels: [],
  }));
  assert.equal(noCoord.ok, false);
  assert.equal(noCoord.error.code, 'E2021');
  assert.match(noCoord.error.message, /lng\/lat/);

  const badLine = check('geoNetwork', makeGeoNetwork({
    lines: [{ id: 'LN9', from: 'S1', to: 'S9' }],
  }));
  assert.equal(badLine.ok, false);
  assert.equal(badLine.error.code, 'E2021');
  assert.match(badLine.error.message, /不存在的厂站：S9/);

  const badPoint = check('geoNetwork', makeGeoNetwork({
    lines: [{ id: 'LN1', from: 'S1', to: 'S2', points: [['S1', null], ['NOPE', null]] }],
  }));
  assert.equal(badPoint.ok, false);
  assert.equal(badPoint.error.code, 'E2021');
  assert.match(badPoint.error.message, /不存在的厂站：NOPE/);

  const wrongPoint = check('geoNetwork', makeGeoNetwork({
    lines: [{ id: 'LN1', from: 'S1', to: 'S2', points: [['S1', 1]] }],
  }));
  assert.equal(wrongPoint.ok, false);
  assert.equal(wrongPoint.error.code, 'E2021');
  assert.match(wrongPoint.error.message, /第二项必须为 null/);
});

test('geoNetwork：station 符号不存在 → E2022；超出 viewport → warn', () => {
  const badSymbol = check('geoNetwork', makeGeoNetwork({
    stations: [{ id: 'S1', symbolRef: 'gb4728.9.NOPE', lng: 113, lat: 31 }],
    lines: [],
    labels: [],
  }));
  assert.equal(badSymbol.ok, false);
  assert.equal(badSymbol.error.code, 'E2022');

  const outside = check('geoNetwork', makeGeoNetwork({
    stations: [
      { id: 'S1', symbolRef: 'gb4728.9.NODE', lng: 125, lat: 31.1 },
      { id: 'S2', symbolRef: 'gb4728.9.NODE', lng: 120, lat: 31.5 },
    ],
    lines: [],
    labels: [],
  }));
  assert.equal(outside.ok, true, outside.ok ? '' : outside.error.message);
  const warnOut = outside.result.warnings.find((item) => item.code === GRID_VALIDATION_CODES.GEO_OUTSIDE_VIEWPORT);
  assert.ok(warnOut, '缺少 GEO_OUTSIDE_VIEWPORT warning');
  assert.match(warnOut.message, /S1/);
});

// ───────────────────────────── powerFlow 非法 / 提示 ─────────────────────────────

test('powerFlow：level 引用未知 / level 电压非法 → E2021', () => {
  const badRef = check('powerFlow', makePowerFlow({
    buses: [{ id: 'B1', level: 'LV9' }],
    nodes: [],
    transformers: [],
    flows: [],
    links: [],
  }));
  assert.equal(badRef.ok, false);
  assert.equal(badRef.error.code, 'E2021');
  assert.match(badRef.error.message, /不存在的电压等级：LV9/);

  const badVoltage = check('powerFlow', makePowerFlow({
    levels: [{ id: 'LV110', voltage: -110 }],
    buses: [],
    nodes: [],
    transformers: [],
    flows: [],
    links: [],
  }));
  assert.equal(badVoltage.ok, false);
  assert.equal(badVoltage.error.code, 'E2021');
  assert.match(badVoltage.error.message, /voltage 必须为正数/);
});

test('powerFlow：flow 端点未知 / direction 不一致 / style 非法 → E2021', () => {
  const badEndpoint = check('powerFlow', makePowerFlow({
    flows: [{ id: 'F9', from: 'B1', to: 'NOPE', kind: 'flow' }],
  }));
  assert.equal(badEndpoint.ok, false);
  assert.equal(badEndpoint.error.code, 'E2021');
  assert.match(badEndpoint.error.message, /不存在的母线\/节点\/变压器：NOPE/);

  const badDirection = check('powerFlow', makePowerFlow({
    flows: [{ id: 'F1', from: 'B1', to: 'B2', kind: 'flow', direction: 'G1→B1' }],
  }));
  assert.equal(badDirection.ok, false);
  assert.equal(badDirection.error.code, 'E2021');
  assert.match(badDirection.error.message, /direction 与 from\/to 不一致/);

  const badStyle = check('powerFlow', makePowerFlow({
    flows: [{ id: 'F1', from: 'B1', to: 'B2', kind: 'flow', style: 'thick' }],
  }));
  assert.equal(badStyle.ok, false);
  assert.equal(badStyle.error.code, 'E2021');
  assert.match(badStyle.error.message, /style 必须为/);
});

test('powerFlow：潮流孤岛 → POWERFLOW_ISLAND warning（不阻断）', () => {
  const model = makePowerFlow();
  model.nodes.push({ id: 'LD1', symbolRef: 'gb4728.3.J', label: '未接入负荷', level: 'LV110' });
  const result = check('powerFlow', model);
  assert.equal(result.ok, true, result.ok ? '' : result.error.message);
  const island = result.result.warnings.find((item) => item.code === GRID_VALIDATION_CODES.POWERFLOW_ISLAND);
  assert.ok(island, '缺少 POWERFLOW_ISLAND warning');
  assert.match(island.message, /LD1/);
});

test('powerFlow：无电源标注 → NO_SOURCE info；变压器同层 → TRANSFORMER_SAME_LEVEL warn', () => {
  const noSource = check('powerFlow', makePowerFlow({
    nodes: [{ id: 'G1', symbolRef: 'gb4728.6.G', label: '金亭电厂', level: 'LV110' }],
    flows: [{ id: 'F2', from: 'G1', to: 'B1', kind: 'flow' }],
    links: [],
  }));
  assert.equal(noSource.ok, true, noSource.ok ? '' : noSource.error.message);
  assert.ok(codesOf(noSource.result.infos).includes(GRID_VALIDATION_CODES.NO_SOURCE));

  const sameLevel = check('powerFlow', makePowerFlow({
    transformers: [{ id: 'T1', symbolRef: 'gb4728.6.T2', from: 'B1', to: 'G1' }],
  }));
  assert.equal(sameLevel.ok, true, sameLevel.ok ? '' : sameLevel.error.message);
  assert.ok(sameLevel.result.warnings.some((item) => item.code === GRID_VALIDATION_CODES.TRANSFORMER_SAME_LEVEL));
});

test('powerFlow：变压器符号不存在 → E2022；links kind 非法 → E2021', () => {
  const badSymbol = check('powerFlow', makePowerFlow({
    transformers: [{ id: 'T1', symbolRef: 'gb4728.6.NOPE', from: 'B1', to: 'B2' }],
  }));
  assert.equal(badSymbol.ok, false);
  assert.equal(badSymbol.error.code, 'E2022');

  const badKind = check('powerFlow', makePowerFlow({
    links: [{ id: 'L1', from: 'B1', to: 'S3', kind: 'wireless' }],
  }));
  assert.equal(badKind.ok, false);
  assert.equal(badKind.error.code, 'E2021');
  assert.match(badKind.error.message, /kind 必须为/);
});

// ───────────────────────────── terminalStrip 非法 ─────────────────────────────

test('terminalStrip：strips 缺失 / columns 非法 / type 非法 → E2021', () => {
  const missing = check('terminalStrip', { kind: 'terminalStrip', meta: { paper: 'A4-landscape' } });
  assert.equal(missing.ok, false);
  assert.equal(missing.error.code, 'E2021');
  assert.match(missing.error.message, /strips/);

  const badColumns = check('terminalStrip', makeTerminalStrip({
    strips: [{ id: 'XT1', columns: 3, terminals: [{ no: 1, ref: 'A411' }] }],
  }));
  assert.equal(badColumns.ok, false);
  assert.equal(badColumns.error.code, 'E2021');
  assert.match(badColumns.error.message, /columns 必须为 1\|2/);

  const badType = check('terminalStrip', makeTerminalStrip({
    strips: [{ id: 'XT1', terminals: [{ no: 1, ref: 'A411', type: 'half' }] }],
  }));
  assert.equal(badType.ok, false);
  assert.equal(badType.error.code, 'E2021');
  assert.match(badType.error.message, /type 必须为/);
});

test('terminalStrip：link 指向自身 / 未知端子号 / 重复 no / ref 空 → E2021', () => {
  const selfLink = check('terminalStrip', makeTerminalStrip({
    strips: [{ id: 'XT1', terminals: [{ no: 1, ref: 'A411', link: 1 }] }],
  }));
  assert.equal(selfLink.ok, false);
  assert.equal(selfLink.error.code, 'E2021');
  assert.match(selfLink.error.message, /不能指向自身/);

  const unknownLink = check('terminalStrip', makeTerminalStrip({
    strips: [{ id: 'XT1', terminals: [{ no: 1, ref: 'A411', link: 9 }] }],
  }));
  assert.equal(unknownLink.ok, false);
  assert.equal(unknownLink.error.code, 'E2021');
  assert.match(unknownLink.error.message, /不存在的端子号：9/);

  const duplicateNo = check('terminalStrip', makeTerminalStrip({
    strips: [{ id: 'XT1', terminals: [{ no: 1, ref: 'A411' }, { no: 1, ref: 'B411' }] }],
  }));
  assert.equal(duplicateNo.ok, false);
  assert.equal(duplicateNo.error.code, 'E2021');
  assert.match(duplicateNo.error.message, /no 重复/);

  const emptyRef = check('terminalStrip', makeTerminalStrip({
    strips: [{ id: 'XT1', terminals: [{ no: 1, ref: ' ' }] }],
  }));
  assert.equal(emptyRef.ok, false);
  assert.equal(emptyRef.error.code, 'E2021');
  assert.match(emptyRef.error.message, /ref 必须为非空回路编号/);
});

test('terminalStrip：terminalsPerPage 越界 → E2021', () => {
  const model = makeTerminalStrip();
  model.meta.terminalsPerPage = 41;
  const result = check('terminalStrip', model);
  assert.equal(result.ok, false);
  assert.equal(result.error.code, 'E2021');
  assert.match(result.error.message, /terminalsPerPage 必须为 1–40 的整数/);
});

// ───────────────────────────── 公共：符号实例上限 ─────────────────────────────

test('P3 五类：符号实例超上限 → E2021', () => {
  const cases = [
    ['protection', makeProtection()],
    ['substationPlan', makeSubstationPlan({
      items: [
        { id: 'T1', symbolRef: 'gb4728.6.T2', xMm: 350, yMm: 60 },
        { id: 'T2', symbolRef: 'gb4728.6.T2', xMm: 420, yMm: 60 },
      ],
      links: [],
    })],
    ['geoNetwork', makeGeoNetwork()],
    ['powerFlow', makePowerFlow()],
    ['terminalStrip', makeTerminalStrip()],
  ];
  for (const [kind, model] of cases) {
    const result = validateGridModel(kind, model, { symbols: SYMBOLS, maxSymbols: 1 });
    assert.equal(result.ok, false, `${kind} 应超过上限`);
    assert.equal(result.error.code, 'E2021');
    assert.match(result.error.message, /超过上限 1/);
  }
});
