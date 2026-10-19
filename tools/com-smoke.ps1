[CmdletBinding()]
param([string]$Json = '{}')

# ============================================================================
# COM 引擎冒烟脚本（docs/P1-设计.md §5.2 D5）
#
# 在有 AutoCAD / 中望CAD / 浩辰CAD 的机器上执行，用于验证 cad-convert.ps1：
#   1) 不带 -dwg：输出可复制的用例说明（JSON），不执行任何转换；
#   2) 带 -dwg：  构造 payload 调用 cad-convert.ps1 场景 A（分割），
#                 并回报退出码、原始输出、产出的 PDF 清单。
#
# 例：
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\com-smoke.ps1 -Json '{}'
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\com-smoke.ps1 -Json '{"dwg":"D:\demo\多图框.dwg","outputDir":"D:\demo\out"}'
#
# 注意：本机无 CAD，脚本未在真机联调（P3/D5 遗留项）。
# ============================================================================

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Write-Json($value) {
  $value | ConvertTo-Json -Depth 10
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$convertScript = Join-Path $repoRoot 'suite\engines\cad-convert.ps1'
$probeScript = Join-Path $repoRoot 'suite\engines\probe.ps1'

if (-not (Test-Path -LiteralPath $convertScript)) {
  Write-Json ([ordered]@{ ok = $false; error = [ordered]@{ code = 'E1099'; message = "未找到引擎脚本: $convertScript" } })
  exit 2
}

$argObj = $null
if ($Json -and $Json.Trim()) {
  try { $argObj = $Json | ConvertFrom-Json } catch { $argObj = $null }
}

# 读取引擎探测（CAD / ODA 是否可用），失败不阻塞冒烟说明输出
$probe = $null
if (Test-Path -LiteralPath $probeScript) {
  try {
    $probeText = & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $probeScript -Json '{}' 2>$null
    if ($probeText) { $probe = ($probeText -join "`n") | ConvertFrom-Json }
  } catch { $probe = $null }
}
$cadSummary = $null
if ($probe) {
  $cadSummary = [ordered]@{
    best      = $probe.cad.best
    platforms = @($probe.cad.platforms | ForEach-Object { [ordered]@{ id = $_.id; present = $_.present; curVer = $_.curVer } })
    odaFound  = [bool]$probe.oda.found
    odaPath   = $probe.oda.converter
  }
}

$dwg = ''
if ($argObj -and $argObj.dwg) { $dwg = [string]$argObj.dwg }
$outputDir = ''
if ($argObj -and $argObj.outputDir) { $outputDir = [string]$argObj.outputDir }
if (-not $outputDir) { $outputDir = Join-Path $repoRoot 'runtime\tmp\com-smoke' }

$cases = @(
  [ordered]@{
    name        = '场景 A：单文件多图框 → 每图一页 PDF'
    description = '验证块名+矩形尺寸识别图框、SetWindowToPlot、PlotToFile（A4 横向）'
    json        = (@{ mode = 'split'; input = $dwg; outputDir = $outputDir; paper = 'A4-landscape'; sortMode = 'drawingNo'; pc3 = 'DWG To PDF.pc3' } | ConvertTo-Json -Compress)
    command     = "powershell -NoProfile -ExecutionPolicy Bypass -File `"$convertScript`" -Json '<json>'"
  },
  [ordered]@{
    name        = '场景 B：单文件 extents 打印（供 Node 合并）'
    description = '验证模型空间 Layout、PlotType=extents、PlotToFile 单文件 PDF'
    json        = (@{ mode = 'merge'; input = $dwg; output = (Join-Path $outputDir 'merged-sample.pdf'); paper = 'A4-landscape'; pc3 = 'DWG To PDF.pc3' } | ConvertTo-Json -Compress)
    command     = "powershell -NoProfile -ExecutionPolicy Bypass -File `"$convertScript`" -Json '<json>'"
  }
)

if (-not $dwg) {
  Write-Json ([ordered]@{
    ok     = $true
    result = [ordered]@{
      note          = '未传入 dwg，仅输出用例说明（本机可安全执行）。传入 -Json 参数并在其中提供 dwg 与 outputDir 即可执行场景 A 冒烟。'
      machine       = [ordered]@{ psVersion = $PSVersionTable.PSVersion.ToString(); os = [System.Environment]::OSVersion.VersionString }
      cad           = $cadSummary
      convertScript = $convertScript
      cases         = $cases
    }
  })
  exit 0
}

# ───────── 执行场景 A 冒烟 ─────────
if (-not (Test-Path -LiteralPath $dwg)) {
  Write-Json ([ordered]@{ ok = $false; error = [ordered]@{ code = 'E1023'; message = "输入 DWG 不存在: $dwg" } })
  exit 2
}
if (-not (Test-Path -LiteralPath $outputDir)) { New-Item -ItemType Directory -Force -Path $outputDir | Out-Null }

$payload = [ordered]@{
  mode       = 'split'
  input      = $dwg
  outputDir  = $outputDir
  paper      = 'A4-landscape'
  sortMode   = 'drawingNo'
  pc3        = 'DWG To PDF.pc3'
  frameBlockNames = @('*图框*', 'A0', 'A1', 'A2', 'A3', 'A4', '*TK*')
} | ConvertTo-Json -Compress -Depth 6

$rawOutput = @()
$exitCode = 0
try {
  $rawOutput = & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $convertScript -Json $payload 2>&1
  $exitCode = $LASTEXITCODE
} catch {
  $rawOutput = @($_.Exception.Message)
  $exitCode = 1
}

$rawText = ($rawOutput | ForEach-Object { [string]$_ }) -join "`n"
$parsed = $null
try {
  $startIndex = $rawText.IndexOf('{')
  $endIndex = $rawText.LastIndexOf('}')
  if ($startIndex -ge 0 -and $endIndex -gt $startIndex) {
    $parsed = $rawText.Substring($startIndex, $endIndex - $startIndex + 1) | ConvertFrom-Json
  }
} catch { $parsed = $null }

$produced = @()
if (Test-Path -LiteralPath $outputDir) {
  $produced = @(Get-ChildItem -LiteralPath $outputDir -Filter '*.pdf' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
}

Write-Json ([ordered]@{
  ok     = ($exitCode -eq 0)
  result = [ordered]@{
    exitCode   = $exitCode
    input      = $dwg
    outputDir  = $outputDir
    payload    = $payload
    engine     = $parsed
    rawOutput  = $rawText
    produced   = $produced
    cad        = $cadSummary
    note       = $(if ($exitCode -eq 0) { '冒烟完成：请人工打开输出 PDF 核对图框裁剪/图号命名/A4 横向' } else { '冒烟失败：请结合 engine.rawOutput 与 CAD 版本定位（pc3/媒体名/COM 方法差异）' })
  }
})
exit $exitCode
