# 第三方依赖与许可清册

本套件计划引入以下第三方组件（最终清单以 `suite\package.json` 与发布包为准，随发版同步本文件）：

| 组件 | 用途 | 协议 |
|---|---|---|
| pdf-lib | PDF 合并、页面裁剪/缩放/拆分 | MIT |
| pdfjs-dist | PDF 矢量解析与图框检测 | Apache-2.0 |
| python-docx | Word 文档读写（Python 兜底引擎） | MIT |
| openpyxl | Excel 工作簿读写（Python 兜底引擎） | MIT |
| python-pptx | 演示文稿读写（后期模块） | MIT |
| lxml | python-docx 依赖 | BSD-3-Clause |

- Microsoft Office / Visio / AutoCAD / 中望 / 浩辰 / LibreOffice 为**用户本机已安装软件**，本套件仅通过其公开 COM/命令行接口调用，不随包分发其任何组件。
- 电气符号库依据公开国家标准 GB/T 4728-2022《电气简图用图形符号》系列整理，符号定义文件标注标准条款编号，供核对。
- 图标为原创作品（SVG）。
