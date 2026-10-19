# 安装前检查与指引脚本（PowerShell 5.1 兼容）
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File tools\install.ps1
$ErrorActionPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$repo = Split-Path -Parent $PSScriptRoot
Write-Host '==== dsh-office-suite-for-grid 安装检查 ===='

function Check-Cmd($name) {
  $c = Get-Command $name -ErrorAction SilentlyContinue
  if ($c) { Write-Host ("[OK]  {0,-16} {1}" -f $name, $c.Source) } else { Write-Host ("[--]  {0,-16} 未找到（可选/按需）" -f $name) }
}
Check-Cmd powershell
Check-Cmd pwsh
Check-Cmd git
Check-Cmd node
Check-Cmd python
Check-Cmd py

$apps = [ordered]@{
  'Word'        = 'HKLM:\SOFTWARE\Classes\Word.Application'
  'Excel'       = 'HKLM:\SOFTWARE\Classes\Excel.Application'
  'Visio'       = 'HKLM:\SOFTWARE\Classes\Visio.Application'
  'AutoCAD'     = 'HKLM:\SOFTWARE\Classes\AutoCAD.Application'
  '中望 CAD'    = 'HKLM:\SOFTWARE\Classes\ZWCAD.Application'
  '浩辰 CAD'    = 'HKLM:\SOFTWARE\Classes\GstarCAD.Application'
}
foreach ($k in $apps.Keys) {
  if (Test-Path $apps[$k]) { Write-Host "[OK]  已安装: $k" } else { Write-Host "[--]  未安装: $k（对应引擎自动降级）" }
}

Write-Host ''
Write-Host '安装方式（任选其一）：'
Write-Host '  1) 市场索引：DSH 插件 -> 市场 -> 添加索引 https://underfacet.github.io/dsh-office-suite-for-grid/index.json'
Write-Host '  2) GitHub Releases：下载 dsh-office-suite-for-grid-x.y.z.tgz 后在插件管理中安装'
Write-Host '  3) 本地链接安装：让 DSH 的智能体执行“用插件管理器把以下目录链接安装：' + (Join-Path $repo 'suite') + '”'
Write-Host '安装后运行：/office-env 或在对话中调用 suite_env 工具查看引擎诊断。'
