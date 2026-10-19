# 电气符号库（GB/T 4728 系列）—— 结构说明与 P2 规划

本目录是**符号库的落盘格式定义**。P1 交付：目录、本说明、1 个示例 JSON，以及配套的
结构校验函数 `suite/host/ops/symbols.js` 的 `validateSymbol()`（`{valid, errors, warnings}`，
不抛异常）。符号渲染、自动布局与 Visio master 组装在 P2 实现（见 `docs/P1-设计.md` §4.1/§4.2）。

- 当前状态：**结构已定稿，图形渲染未实现（P2）**
- 示例文件：`gb4728.7/example-qf.json`（断路器，`gb4728.7.QF`）
- 校验入口：`validateSymbol(json)` / `assertValidSymbol(json)`（`suite/host/ops/symbols.js`）
- 权威术语/条款号维护表：`docs/符号对照表.md`（本目录只放几何定义）

## 1. 目录与命名

```
symbols/
├─ README.md                     # 本文件
└─ gb4728.7/                     # GB/T 4728.7 开关、控制和保护器件
   └─ example-qf.json            # 断路器示例（P2 起改为正式文件名 QF.json）
```

- 系列目录名 = 标准部分号：`gb4728.6`（电能的发生与转换）、`gb4728.7`（开关、控制和保护器件）、
  `gb4728.8`（测量仪表、灯和信号器件）、`gb4728.11`（建筑安装平面布置图）等。
- 正式符号文件名 = 符号 id 的末段（`gb4728.7.QF` → `gb4728.7/QF.json`）。
- 示例文件保留 `example-` 前缀，避免与后续正式库混淆。
- `id` 规则：`gb4728.<部分号>.<符号代号>`，代号用大写 ASCII（`QF` 断路器、`QS` 隔离开关、
  `FU` 熔断器、`TA` 电流互感器、`TV` 电压互感器、`T` 变压器、`G` 发电机、`M` 电动机、
  `F` 避雷器等），组合图形用连字符（`RING-MAIN` 环网柜）。

## 2. 单个符号 JSON 的字段

| 字段 | 类型 | 必填 | 校验行为（`validateSymbol`） |
|---|---|---|---|
| `id` | string | 是 | 非空字符串，否则报错 |
| `name.zh` / `name.en` | string | 是 | `zh` 必须非空；`en` 可缺省但存在时须为字符串 |
| `category` | string | 否 | 存在时须为字符串（`switchgear` / `machine` / `transformer` / `instrument` / `protective` / `connection` / `installation` 等约定值） |
| `standard.part` | string | 建议 | 缺失只告警；例如 `GB/T 4728.7-2022` |
| `standard.clause` | string | 否 | 存在时须为字符串；**未与原版标准核对前留空**，示例中的值为示意 |
| `view.widthMm` / `view.heightMm` | number | 是 | 必须为正数（mm） |
| `primitives` | array | 是 | 必须为非空数组，逐项按 §3 校验 |
| `terminals` | array | 建议 | 缺失只告警；存在时逐项校验 `id` 与 `at` |

坐标约定：**mm**，原点取 `view` 包围盒左上角，x 向右、y 向下（与 Visio 页面坐标系一致，
P2 的 Visio 引擎按此直接换算，不做 y 轴翻转）。`stroke` 单位同为 mm（`0.35` ≈ 1pt 线）。

## 3. `primitives` 图元（与校验器枚举一致）

允许的 `type`：`polyline`、`line`、`circle`、`arc`、`rect`、`text`。

| `type` | 必填字段 | 可选字段 | 说明 |
|---|---|---|---|
| `polyline` | `points: [x,y][]`（≥2 点） | `stroke`、`closed` | 折线；`closed:true` 表首尾闭合（P2 引擎识别，校验器不检查该字段） |
| `line` | `points`（2 点） | `stroke` | 单段直线，等价于两点 `polyline` |
| `circle` | `center:[x,y]`、`radius>0` | `stroke` | 圆（字段名是 `radius`，不是 `r`） |
| `arc` | `center:[x,y]` | `radius`、`fromDeg`、`toDeg`、`stroke` | 圆弧；角度与半径语义由 P2 引擎实现时细化，校验器目前只强制 `center` |
| `rect` | `origin:[x,y]`、`width`、`height` | `stroke` | 轴对齐矩形（图框、端子盒、仪表框） |
| `text` | `at:[x,y]`、`text` | `sizeMm`、`anchor`、`rotate` | 注释文字；`anchor` ∈ `top-left|top-center|center|bottom-center|bottom-left` |

一个符号的图元按数组顺序绘制（后绘制覆盖先绘制）。**颜色一律不落盘**：渲染色由主题决定
（P2 统一用 Visio 主题色/黑色），因此 JSON 里不出现颜色字段。实心图形（接地符号等）当前用
`rect`/`circle` + 后续 P2 增补的 `fill` 语义表达；校验器暂不校验 `fill`。

## 4. `terminals` 连接锚点

```json
"terminals": [ { "id": "T1", "at": [0, 6] }, { "id": "T2", "at": [16, 6] } ]
```

- `id`：符号内唯一（`T1`、`T2`…），图数据里通过 `{ "symbolRef": "gb4728.7.QF", "terminal": "T1" }` 引用；
- `at`：与 `primitives` 同一坐标系（mm），必须为 `[x,y]`；
- `direction?`（P2 可选扩展）：`left|right|up|down`，提示连线出线方向，未给则由布局按几何自动判定；
  校验器忽略该扩展字段；
- 母线类符号用一条 `polyline` + 多个等距 `terminals` 表达，避免为每个接点造一个符号。

## 5. 与图数据（P2）的关系

`symbolRef` 指向本目录的符号 id，例如主接线图：

```jsonc
{ "devices": [ { "id": "QF1", "symbolRef": "gb4728.7.QF", "label": "断路器", "state": "closed" } ] }
```

P2 的 Visio 引擎（`engines/visio-grid.ps1`）在**内存中组装 master**：读 JSON → 画图元 →
设置 `terminals` 为连接点 → 由布局算法按 `buses/branches` 摆放并连线。**不依赖本机模具文件**，
因此本目录必须自包含（图元全部用坐标表达，不引用外部 stencil 形状）。

## 6. P2 实施清单

1. **已具备（P1）**：`suite/host/ops/symbols.js` 的 `validateSymbol` / `assertValidSymbol`
   （结构校验，含图元枚举与坐标检查）。
2. 目录扫描：`loadSymbols(dir)`（id 去重、按 `standard.part` 分组、逐条跑 `validateSymbol` 汇总报告）。
3. 正式符号集：按 `docs/符号对照表.md` 的电网常用清单补齐（母线、断路器、隔离开关、接地刀闸、
   负荷开关、熔断器、变压器、发电机、CT/PT、避雷器、电容器/电抗器、接地、连接点、环网柜、通信节点）。
4. `clause` 逐条核对纸质/正版标准后填写，并在 `docs/符号对照表.md` 记录核对日期。
5. 组合图形（环网柜/台区）允许引用其他符号：新增 `"compose": [{ "symbolRef": "...", "at": [x,y] }]`
   顶层字段（P2 扩展，需在 `validateSymbol` 中递归检查引用与环）。
6. 单位自备图库（`.vss/.vssx`）导入：同 id 时优先用官方符号，导入件只在缺失时兜底；
   导入不写回本目录。
7. 渲染验证：P2 用 `tests/fixtures` 下的合成图数据比对图元数量/连接点数，不引入真实图纸。

## 7. 许可与来源

几何形状依据 GB/T 4728-2022 系列标准自行绘制为坐标数据（不复制标准原文图形文件）；
`standard.clause` 仅作引用标注。标准原文与印刷版图形不在本仓库分发。
