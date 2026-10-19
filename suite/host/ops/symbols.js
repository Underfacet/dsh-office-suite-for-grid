/**
 * 符号库结构校验占位（docs/P1-设计.md §4.2）
 *
 * P1 只交付校验函数；符号库文件（suite/symbols/gb4728/<系列>/<id>.json）在 P2 生成。
 * 坐标约定与 suite/symbols/README.md 一致：单位 mm，原点取 view 包围盒左上角，
 * x 向右、y 向下（Visio 页面坐标系）。
 * JSON 结构：
 * {
 *   id, name:{zh,en}, category,
 *   standard:{part,clause},
 *   view:{widthMm,heightMm},
 *   primitives:[{type:'polyline'|'line'|'circle'|'arc'|'rect'|'text', ...}],
 *   terminals:[{id, at:[x,y]}]
 * }
 *
 * validateSymbol 返回 { valid, errors, warnings }，不抛异常（便于面板/批量校验逐条报告）。
 */

/** 支持的图元类型（P2 引擎实现时扩展）。 */
export const SYMBOL_PRIMITIVE_TYPES = Object.freeze(['polyline', 'line', 'circle', 'arc', 'rect', 'text']);

function isFiniteNumber(value) {
  return typeof value === 'number' && Number.isFinite(value);
}

function isPoint(value) {
  return Array.isArray(value) && value.length >= 2 && isFiniteNumber(value[0]) && isFiniteNumber(value[1]);
}

/**
 * 校验单个符号定义。
 * @param {any} symbol
 * @returns {{valid:boolean, errors:string[], warnings:string[]}}
 */
export function validateSymbol(symbol) {
  const errors = [];
  const warnings = [];
  const fail = (message) => errors.push(message);

  if (!symbol || typeof symbol !== 'object' || Array.isArray(symbol)) {
    return { valid: false, errors: ['符号必须是 JSON 对象'], warnings };
  }

  if (typeof symbol.id !== 'string' || !symbol.id.trim()) fail('缺少 id（字符串，形如 gb4728.7.QF）');
  if (!symbol.name || typeof symbol.name !== 'object') {
    fail('缺少 name 对象（{ zh, en }）');
  } else {
    if (typeof symbol.name.zh !== 'string' || !symbol.name.zh.trim()) fail('name.zh 必须为非空字符串');
    if (symbol.name.en != null && typeof symbol.name.en !== 'string') fail('name.en 必须为字符串');
  }
  if (symbol.category != null && typeof symbol.category !== 'string') fail('category 必须为字符串');

  if (symbol.standard != null) {
    if (typeof symbol.standard !== 'object' || Array.isArray(symbol.standard)) {
      fail('standard 必须为对象（{ part, clause }）');
    } else {
      if (typeof symbol.standard.part !== 'string' || !symbol.standard.part.trim()) warnings.push('standard.part 缺失（建议填写标准分部，如 GB/T 4728.7-2022）');
      if (symbol.standard.clause != null && typeof symbol.standard.clause !== 'string') fail('standard.clause 必须为字符串');
    }
  } else {
    warnings.push('standard 缺失（建议填写标准编号与条款）');
  }

  if (!symbol.view || typeof symbol.view !== 'object') {
    fail('缺少 view 对象（{ widthMm, heightMm }）');
  } else {
    if (!isFiniteNumber(symbol.view.widthMm) || symbol.view.widthMm <= 0) fail('view.widthMm 必须为正数（mm）');
    if (!isFiniteNumber(symbol.view.heightMm) || symbol.view.heightMm <= 0) fail('view.heightMm 必须为正数（mm）');
  }

  if (!Array.isArray(symbol.primitives) || !symbol.primitives.length) {
    fail('primitives 必须为非空数组');
  } else {
    symbol.primitives.forEach((primitive, index) => {
      const label = `primitives[${index}]`;
      if (!primitive || typeof primitive !== 'object') {
        fail(`${label} 必须为对象`);
        return;
      }
      if (!SYMBOL_PRIMITIVE_TYPES.includes(primitive.type)) {
        fail(`${label}.type 必须为 ${SYMBOL_PRIMITIVE_TYPES.join('|')}，收到: ${JSON.stringify(primitive.type)}`);
        return;
      }
      if ((primitive.type === 'polyline' || primitive.type === 'line') && (!Array.isArray(primitive.points) || primitive.points.length < 2 || !primitive.points.every(isPoint))) {
        fail(`${label}.points 必须为至少 2 个 [x,y] 点`);
      }
      if ((primitive.type === 'circle' || primitive.type === 'arc') && !isPoint(primitive.center)) {
        fail(`${label}.center 必须为 [x,y]`);
      }
      if (primitive.type === 'circle' && !(isFiniteNumber(primitive.radius) && primitive.radius > 0)) {
        fail(`${label}.radius 必须为正数`);
      }
      if (primitive.type === 'rect' && !isPoint(primitive.origin)) fail(`${label}.origin 必须为 [x,y]`);
      if (primitive.type === 'rect' && !(isFiniteNumber(primitive.width) && isFiniteNumber(primitive.height))) {
        fail(`${label}.width/height 必须为数值`);
      }
      if (primitive.type === 'text') {
        if (typeof primitive.text !== 'string' || !primitive.text) fail(`${label}.text 必须为非空字符串`);
        if (!isPoint(primitive.at)) fail(`${label}.at 必须为 [x,y]`);
      }
      if (primitive.stroke != null && !(isFiniteNumber(primitive.stroke) && primitive.stroke > 0)) {
        warnings.push(`${label}.stroke 建议为正数（mm）`);
      }
    });
  }

  if (symbol.terminals != null) {
    if (!Array.isArray(symbol.terminals)) {
      fail('terminals 必须为数组');
    } else {
      symbol.terminals.forEach((terminal, index) => {
        const label = `terminals[${index}]`;
        if (!terminal || typeof terminal !== 'object') {
          fail(`${label} 必须为对象`);
          return;
        }
        if (typeof terminal.id !== 'string' || !terminal.id.trim()) fail(`${label}.id 必须为非空字符串`);
        if (!isPoint(terminal.at)) fail(`${label}.at 必须为 [x,y]`);
      });
    }
  } else {
    warnings.push('terminals 缺失（建议提供接线端子坐标）');
  }

  return { valid: errors.length === 0, errors, warnings };
}

/** 校验失败时抛出（供 P2 引擎使用）；P1 仅保留占位。 */
export function assertValidSymbol(symbol) {
  const result = validateSymbol(symbol);
  if (!result.valid) {
    const error = new Error(`符号结构无效：${result.errors.join('; ')}`);
    error.name = 'SymbolValidationError';
    error.details = result;
    throw error;
  }
  return symbol;
}
