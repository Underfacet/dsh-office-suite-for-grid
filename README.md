# dsh-office-suite-for-grid

[![ci](https://github.com/Underfacet/dsh-office-suite-for-grid/actions/workflows/ci.yml/badge.svg)](https://github.com/Underfacet/dsh-office-suite-for-grid/actions) [![license](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

DSH 办公·电网绘图套件 —— 面向电网行业的 DeepSeek Harness（DSH）插件：在 DSH Web / Desktop 中一站式完成 Office 文档与表格的阅读、编辑、排版，驱动 Visio 绘制电网专业图表，并将 CAD 图纸智能转换为 PDF（自动分割 / 合并，适配 A4 打印）。

## 功能一览

| 模块 | 能力 |
|---|---|
| Office 文档/表格 | docx/xlsx 阅读（大纲、分页、表格与公式）、结构化编辑、通用排版、公文/红头模板、批量排版、导出 PDF 与渲染预览 |
| Visio 电网绘图 | 电气主接线图、配电系统图、通信网络拓扑、二次/继保图、变电站布置、地理接线图、时间线、流程图、组织结构图、思维导图、泳道图；内置 GB/T 4728-2022 标准符号库，支持挂接单位自备模具库 |
| CAD 转 PDF | 文件夹多张 DWG 自动合并为一个 PDF；单文件合并多图框图纸自动分割为每图一页（默认横向 A4）；图号优先排序；无 CAD 机器经 PDF 级引擎兜底处理 |

## 特性

- **符合 DSH 插件标准**：标准 bundle 清单（`dsh.bundle.patch`、`engines.dsh`、locale、图标、meta），Web 与 Desktop 共用同一套 Client 插件，只使用公开扩展点，多版本自动降级。
- **本地软件版本适配**：自动探测 Word/Excel/Visio、AutoCAD/中望/浩辰、LibreOffice、Python 等软件的安装与版本，按版本选择调用路径，能力分级、优雅降级，诊断页一目了然。
- **独立部署**：全部源码与运行时（venv、日志、备份、缓存、输出）位于项目文件夹内，官方 DSH 程序目录零写入；卸载即删除目录。
- **速度优先**：COM 会话复用、Python 常驻 worker、批量并行（可配置）、探测结果缓存、面板推送不轮询。
- **开源**：MIT 协议；第三方依赖许可证清册见 `THIRD_PARTY_NOTICES.md`。

## 安装（多种方式，Web / Desktop 通用）

1. **市场索引（推荐）**：在 DSH「插件 → 市场」添加本仓库 Pages 索引地址 `https://underfacet.github.io/dsh-office-suite-for-grid/index.json`，随后在发现页一键安装与升级。
2. **GitHub Releases**：下载 `dsh-office-suite-for-grid-x.y.z.tgz`，在插件管理中选择该文件安装。
3. **npm**：`npm install` 后经插件管理按包名安装（`dsh-office-suite-for-grid`）。
4. **源码/本地目录**：克隆本仓库后运行 `tools\install.ps1`（链接安装，代码留在本目录）。

## 目录结构

```
suite\        插件本体（安装对象）
runtime\      用户机运行时（venv、日志、备份、缓存）——自动生成，不入库
output\       产物：diagrams\（.vsdx）、pdf\（合并/分割 PDF）
docs\         设计文档、用户手册、兼容矩阵、符号对照表
tests\        脱敏公开样例与冒烟脚本
tests-local\  真实样例图纸（本地保密，不入库）
tools\        环境探测、安装/升级、发布脚本
```

## 协议与维护

- 协议：MIT（见 LICENSE）
- 分支：`main`（稳定）/ `dev`（开发）/ `feature\*`；发版打 `release-x.y.z` 标签，变更记录见 CHANGELOG.md
