# grid-smoke.ps1 —— visio-grid.ps1 冒烟脚本（三类图各一最小用例）
#
# 用法（Windows PowerShell 5.1）：
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\grid-smoke.ps1
#   powershell -NoProfile -ExecutionPolicy Bypass -File tools\grid-smoke.ps1 -OutDir runtime\tmp\diagrams
#
# 行为：
# 1. 从 suite\symbols 读取符号库（跳过 P1 示例 example-*.json），按用例注入所需符号；
# 2. 构造 singleLine / distribution / networkTopology 三个最小模型；
# 3. 以独立子进程调用 suite\engines\visio-grid.ps1（-Json 载荷写临时文件，避免命令行长度限制）；
# 4. 断言返回 ok=true、输出 .vsdx 存在且大小 > 0，并打印统计表；任一用例失败则 exit 1。
#
# 产物位于 runtime\tmp\diagrams（git 忽略）；不写入仓库受版本控制文件。

[CmdletBinding()]
param(
  [string]$OutDir = '',
  [string]$SymbolsDir = '',
  [string]$EnginePath = ''
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $OutDir) { $OutDir = Join-Path $repoRoot 'runtime\tmp\diagrams' }
if (-not $SymbolsDir) { $SymbolsDir = Join-Path $repoRoot 'suite\symbols' }
if (-not $EnginePath) { $EnginePath = Join-Path $repoRoot 'suite\engines\visio-grid.ps1' }
$payloadDir = Join-Path $repoRoot 'runtime\tmp\grid-smoke'

if (-not (Test-Path -LiteralPath $EnginePath)) { throw "引擎脚本不存在：$EnginePath" }
if (-not (Test-Path -LiteralPath $SymbolsDir)) { throw "符号库目录不存在：$SymbolsDir" }
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }
if (-not (Test-Path -LiteralPath $payloadDir)) { New-Item -ItemType Directory -Force -Path $payloadDir | Out-Null }

function Read-SymbolLibrary {
  param([string]$Root)
  $library = @{}
  foreach ($directory in @(Get-ChildItem -LiteralPath $Root -Directory)) {
    if ($directory.Name -notmatch '^gb4728(\.\d+)?$') { continue }
    foreach ($file in @(Get-ChildItem -LiteralPath $directory.FullName -Filter '*.json' -File)) {
      if ($file.Name.ToLowerInvariant().StartsWith('example-')) { continue }
      $symbol = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
      $symbolId = [string]$symbol.id
      if ($symbolId) { $library[$symbolId] = $symbol }
    }
  }
  return $library
}

function Select-Symbols {
  param($Library, [string[]]$Ids)
  $selected = @{}
  foreach ($id in $Ids) {
    if (-not $Library.ContainsKey($id)) { throw "符号库缺少符号：$id" }
    $selected[$id] = $Library[$id]
  }
  return $selected
}

function Invoke-GridCase {
  param([string]$Name, [string]$Kind, $Model, [string[]]$SymbolIds)

  $symbols = Select-Symbols -Library $script:Library -Ids $SymbolIds
  $outputPath = Join-Path $OutDir ("{0}-smoke.vsdx" -f $Name)
  if (Test-Path -LiteralPath $outputPath) { Remove-Item -LiteralPath $outputPath -Force }

  $payload = @{
    kind = $Kind
    title = "冒烟：$Name"
    model = $Model
    symbols = $symbols
    outputPath = $outputPath
    validations = @()
  }
  $payloadPath = Join-Path $payloadDir ("{0}.json" -f $Name)
  [System.IO.File]::WriteAllText($payloadPath, ($payload | ConvertTo-Json -Depth 20), (New-Object System.Text.UTF8Encoding($false)))

  $command = "& '$EnginePath' -Json (Get-Content -LiteralPath '$payloadPath' -Raw -Encoding UTF8)"
  $stdout = & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command $command 2>&1
  $text = ($stdout | Out-String)
  $parsed = $null
  try { $parsed = $text | ConvertFrom-Json } catch { $parsed = $null }

  $record = [pscustomobject]@{
    Case = $Name
    Kind = $Kind
    Ok = $false
    Symbols = 0
    Shapes = 0
    Links = 0
    Bytes = 0
    Output = $outputPath
    Error = ''
  }
  if ($null -eq $parsed) {
    $record.Error = "引擎输出非 JSON：$($text.Trim())"
    return $record
  }
  if (-not $parsed.ok) {
    $record.Error = ("{0}: {1}" -f $parsed.error.code, $parsed.error.message)
    return $record
  }
  $record.Ok = $true
  $record.Symbols = $parsed.result.stats.symbols
  $record.Shapes = $parsed.result.stats.shapes
  $record.Links = $parsed.result.stats.links
  if (Test-Path -LiteralPath $parsed.result.outputPath) {
    $record.Bytes = (Get-Item -LiteralPath $parsed.result.outputPath).Length
  }
  if (-not (Test-Path -LiteralPath $parsed.result.outputPath)) { $record.Ok = $false; $record.Error = '输出文件不存在' }
  elseif ($record.Bytes -le 0) { $record.Ok = $false; $record.Error = '输出文件为空' }
  return $record
}

# ── 最小模型（§4.1） ──
$singleLineModel = @{
  kind = 'singleLine'
  title = '冒烟：110kV 单母分段主接线'
  meta = @{ paper = 'A3-landscape'; standard = 'GB/T 4728-2022' }
  buses = @(
    @{ id = 'B1'; label = '110kV I 母'; voltage = 110; section = 'I' },
    @{ id = 'B2'; label = '110kV II 母'; voltage = 110; section = 'II' }
  )
  nodes = @(
    @{ id = 'G1'; symbolRef = 'gb4728.6.G'; label = '1# 发电机'; tags = @{ source = $true } }
  )
  links = @(
    @{ id = 'K1'; from = 'G1'; to = 'B1'; kind = 'line'; label = '110kV 引线' }
  )
  bays = @(
    @{ id = 'X1'; bus = 'B1'; label = '1# 主变高压侧'; devices = @(
        @{ id = 'QS11'; symbolRef = 'gb4728.7.QS'; state = 'closed' },
        @{ id = 'QF11'; symbolRef = 'gb4728.7.QF'; state = 'closed' },
        @{ id = 'QS12'; symbolRef = 'gb4728.7.QS'; state = 'closed' }
      ) },
    @{ id = 'X2'; bus = 'B2'; label = '2# 主变高压侧'; devices = @(
        @{ id = 'QS21'; symbolRef = 'gb4728.7.QS'; state = 'closed' },
        @{ id = 'QF21'; symbolRef = 'gb4728.7.QF'; state = 'closed' }
      ) }
  )
  transformers = @(
    @{ id = 'T1'; symbolRef = 'gb4728.6.T2'; label = '1# 主变（110/10kV）'; fromBay = 'X1'; toBay = 'X2'; voltage = @(110, 10) }
  )
  busTies = @(
    @{ id = 'BT1'; between = @('B1', 'B2'); devices = @(
        @{ id = 'QF0'; symbolRef = 'gb4728.7.QF'; state = 'open' }
      ) }
  )
}

$distributionModel = @{
  kind = 'distribution'
  title = '冒烟：10kV 配电系统'
  meta = @{ paper = 'A3-landscape' }
  sources = @(
    @{ id = 'S1'; symbolRef = 'gb4728.3.BUS'; label = '110kV 站 10kV 母线'; tags = @{ source = $true } }
  )
  nodes = @(
    @{ id = 'R1'; symbolRef = 'gb4728.7.RING-MAIN'; label = '1# 环网柜' },
    @{ id = 'R2'; symbolRef = 'gb4728.7.RING-MAIN'; label = '2# 环网柜' },
    @{ id = 'T1'; symbolRef = 'gb4728.6.T2'; label = '1# 台区（10/0.4kV）' }
  )
  feeders = @(
    @{ id = 'F1'; label = '馈线 101'; from = 'S1'; to = 'R1'; kind = 'feeder' },
    @{ id = 'F2'; label = '馈线 102'; from = 'R1'; to = 'R2'; kind = 'feeder' },
    @{ id = 'F3'; label = '馈线 103'; from = 'R2'; to = 'T1'; kind = 'feeder' }
  )
}

$topologyModel = @{
  kind = 'networkTopology'
  title = '冒烟：通信光纤环网'
  meta = @{ paper = 'A3-landscape' }
  nodes = @(
    @{ id = 'C1'; symbolRef = 'gb4728.9.NODE'; label = '中心站' },
    @{ id = 'C2'; symbolRef = 'gb4728.9.NODE'; label = '变电站 A' },
    @{ id = 'C3'; symbolRef = 'gb4728.9.NODE'; label = '变电站 B' },
    @{ id = 'C4'; symbolRef = 'gb4728.9.NODE'; label = '变电站 C' }
  )
  links = @(
    @{ id = 'L1'; from = 'C1'; to = 'C2'; kind = 'fiber' },
    @{ id = 'L2'; from = 'C2'; to = 'C3'; kind = 'fiber' },
    @{ id = 'L3'; from = 'C3'; to = 'C4'; kind = 'fiber' },
    @{ id = 'L4'; from = 'C4'; to = 'C1'; kind = 'fiber' }
  )
  rings = @(
    @{ id = 'RING1'; label = '光纤环 1'; nodes = @('C1', 'C2', 'C3', 'C4') }
  )
}

$script:Library = Read-SymbolLibrary -Root $SymbolsDir
Write-Host ("[grid-smoke] 符号库：{0} 个符号（{1}）" -f $script:Library.Count, $SymbolsDir)
Write-Host ("[grid-smoke] 引擎：{0}" -f $EnginePath)
Write-Host ("[grid-smoke] 输出目录：{0}" -f $OutDir)

$results = @()
$results += Invoke-GridCase -Name 'singleLine' -Kind 'singleLine' -Model $singleLineModel -SymbolIds @('gb4728.3.BUS', 'gb4728.6.G', 'gb4728.6.T2', 'gb4728.7.QS', 'gb4728.7.QF')
$results += Invoke-GridCase -Name 'distribution' -Kind 'distribution' -Model $distributionModel -SymbolIds @('gb4728.3.BUS', 'gb4728.6.T2', 'gb4728.7.RING-MAIN', 'gb4728.7.QL', 'gb4728.7.FU')
$results += Invoke-GridCase -Name 'networkTopology' -Kind 'networkTopology' -Model $topologyModel -SymbolIds @('gb4728.9.NODE')

$results | Format-Table Case, Kind, Ok, Symbols, Shapes, Links, Bytes, Error -AutoSize | Out-String -Width 200 | Write-Host

$failed = @($results | Where-Object { -not $_.Ok })
if ($failed.Count -gt 0) {
  Write-Host ("[grid-smoke] 失败 {0}/{1}" -f $failed.Count, $results.Count)
  exit 1
}
Write-Host ("[grid-smoke] 全部通过（{0} 个用例，输出 .vsdx）" -f $results.Count)
exit 0
