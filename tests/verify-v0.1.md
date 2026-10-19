# v0.1 真机验证记录

- 日期：2026-10-20（凌晨空闲时段）
- 环境：DSH 0.2.1-rc.2（desktop 配置，Web GUI）、Windows 11、Word/Excel/PowerPoint/Visio 16、Python 3.14、无 CAD/LibreOffice/ODA
- 方法：插件经插件管理器链接安装后，由独立子智能体在新会话中通过对话工具活体调用。

## 结果

| 项目 | 方式 | 结果 |
|---|---|---|
| 工具注册 | Tool 目录检查 | suite_env / office_read / visio_draw 均已注册且 schema 正确 |
| suite_env | 对话工具调用 | 返回完整探测 JSON（含 visioInstalledVersions=[16.0]、officeInstalledVersions=[16.0]） |
| office_read (xlsx) | 对话工具调用 | 正确识别「台账」工作表 2 行 3 列，表头与数据样例无误 |
| office_read (docx) | 引擎直测 | 大纲（标题 1×2）、正文预览、表格内容正确 |
| visio_draw | 引擎直测 | 生成 sample-flowchart.vsdx（15023 字节，5 节点 4 连线） |
| 引擎脚本编码 | 检查 | 三个 .ps1 均为 UTF-8 带 BOM（PS 5.1 兼容），pwsh→powershell 回退链生效 |

## 结论与限制

- v0.1 全链路（工具注册 → 子进程 → PowerShell → COM → JSON 结果）可用。
- 本机无 AutoCAD/中望/浩辰、LibreOffice、ODA：CAD 转 PDF 与渲染/PDF 路径在本机不可用，按设计走降级与后续 PDF 级引擎；相关功能在有对应软件的机器上验收。
- Visio CurVer=11 为残留注册项，真实安装为 16.0（probe 已双重检测）。
