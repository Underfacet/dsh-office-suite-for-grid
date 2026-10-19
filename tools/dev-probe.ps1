# 开发用：一键环境探测（调用随包 probe 引擎）
# 用法：powershell -NoProfile -ExecutionPolicy Bypass -File tools\dev-probe.ps1
$repo = Split-Path -Parent $PSScriptRoot
Set-ExecutionPolicy -Scope Process Bypass -Force | Out-Null
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
& (Join-Path $repo 'suite\engines\probe.ps1')
