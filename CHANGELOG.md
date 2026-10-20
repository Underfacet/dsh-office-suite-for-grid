# 更新日志

本项目遵循[语义化版本](https://semver.org/lang/zh-CN/)。

## [1.0.0] - 2026-10-20

首个正式版本：Office 读写编辑排版（COM+Python 双引擎、备份、预览导出）、Visio 九类图（含八类电网专业图与 GB/T 4728 符号库）、CAD 图纸智能转 PDF（分割/合并/目录表/批量队列）、四页签面板、市场索引与 CI 发布基建。

## [未发布]

### 0.1.0（P2 开发中）

- P0：仓库骨架、设计文档、环境探测、GitHub 建库与自动推送
- P1：套件骨架（Host：3 工具 + 11 命令 + 引擎诊断；Client：四页面板）
- P1：Office 读取（Word 大纲/正文/表格、Excel 工作表/样例，COM 引擎）
- P1：Visio 流程图（COM 引擎，自动分列布局，.vsdx）
- P1：CAD 转 PDF 后端：COM/ODA/PDF 三级链路、PDF 级分割与合并引擎、图框识别、图号提取排序（13 项单元测试）
- P1：边界校验（2026 目录）与任务进度/取消
- P2：Office 编辑/排版（office_edit/office_format 工具、9 个新命令、COM+Python venv 双引擎、自动备份 3 份轮换、样式集/模板、批量排版任务）
- P2：Visio 电网图（主接线图/配电系统图/网络拓扑图，接线校验，17 项 GB/T 4728 符号库，本机三图冒烟通过）
- P2：suite.preview 契约修正、默认样式集、配置白名单扩展
- P3：CAD 目录表（suite.cad.catalog，图号/图名/图幅/比例/张数 → xlsx/CSV）、批量队列（inputs[] 父任务聚合）、Office 导出（office_export/suite.export，走内置 officeToPdf，真 PDF 实测）
- P3：五类新电网图（protection/substationPlan/geoNetwork/powerFlow/terminalStrip）+ visio-layouts 引擎（本机冒烟 5/5，terminalStrip 多页）
- P3：visio_draw 九类分派、命令 21 个、cad-batch 引擎预判修复


