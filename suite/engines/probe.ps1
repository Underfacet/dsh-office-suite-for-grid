[CmdletBinding()]
param([string]$Json = '{}')
$ErrorActionPreference = 'SilentlyContinue'
$result = [ordered]@{}

function CurVer($n) { (Get-ItemProperty "HKLM:\SOFTWARE\Classes\$n\CurVer").'(default)' }

function Get-RegDefault($path) {
  $item = Get-ItemProperty -LiteralPath $path -ErrorAction SilentlyContinue
  if ($null -eq $item) { return $null }
  return $item.'(default)'
}

# ───────────────────────── v0.1 既有字段（保持不变） ─────────────────────────
foreach ($n in 'Word.Application','Excel.Application','PowerPoint.Application','Visio.Application','AutoCAD.Application','ZWCAD.Application','GstarCAD.Application') {
  $prog = Test-Path "HKLM:\SOFTWARE\Classes\$n"
  $result[$n] = [ordered]@{
    present = [bool]$prog
    curVer  = if ($prog) { (CurVer $n) } else { $null }
  }
}

# CurVer 可能指向残留的旧版本注册项，这里补查真实安装键（x64 与 WOW64 两个 hive）
$visioVersions = @()
$officeVersions = @()
foreach ($v in '16.0','15.0','14.0','12.0','11.0') {
  if ((Test-Path "HKLM:\SOFTWARE\Microsoft\Office\$v\Visio") -or (Test-Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\$v\Visio")) { $visioVersions += $v }
  if ((Test-Path "HKLM:\SOFTWARE\Microsoft\Office\$v\Word\InstallRoot") -or (Test-Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\$v\Word\InstallRoot")) { $officeVersions += $v }
}
$result['visioInstalledVersions'] = $visioVersions
$result['officeInstalledVersions'] = $officeVersions

$py = Get-Command python -ErrorAction SilentlyContinue
$result['python'] = [ordered]@{
  found   = [bool]$py
  version = ((python --version 2>&1) -join ' ').Trim()
}
$pyl = Get-Command py -ErrorAction SilentlyContinue
$result['pyLauncher'] = [ordered]@{
  found = [bool]$pyl
  list  = ((py -0p 2>&1 | Out-String)).Trim()
}

$lo = @('C:\Program Files\LibreOffice\program\soffice.exe','C:\Program Files (x86)\LibreOffice\program\soffice.exe') | Where-Object { Test-Path $_ }
$result['libreOffice'] = [ordered]@{ found = ($lo.Count -gt 0); paths = @($lo) }

$oda = @((Get-ChildItem 'C:\Program Files\ODA','C:\Program Files (x86)\ODA' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName))

# ───────────────────────── P1 新增：CAD（AutoCAD / 中望 / 浩辰） ─────────────────────────
$appPathRoots = @(
  'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths',
  'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths',
  'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths',
  'HKCU:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths'
)

$platformDefs = @(
  [ordered]@{
    id = 'autocad'; name = 'AutoCAD'; progId = 'AutoCAD.Application'; exeNames = @('acad.exe')
    vendorRoots = @('SOFTWARE\Autodesk\AutoCAD')
    locationProps = @('AcadLocation','Location','InstallPath')
    appDataDirs = @('Autodesk')
  },
  [ordered]@{
    id = 'zwcad'; name = '中望CAD'; progId = 'ZWCAD.Application'; exeNames = @('zwcad.exe')
    vendorRoots = @('SOFTWARE\ZWSOFT\ZWCAD','SOFTWARE\ZWSOFT\ZWCAD+','SOFTWARE\ZWSOFT')
    locationProps = @('Location','InstallPath','Path','AcadLocation')
    appDataDirs = @('ZWSOFT')
  },
  [ordered]@{
    id = 'gstarcad'; name = '浩辰CAD'; progId = 'GstarCAD.Application'; exeNames = @('gcad.exe','GstarCAD.exe')
    vendorRoots = @('SOFTWARE\Gstarsoft\GstarCAD','SOFTWARE\GstarCAD','SOFTWARE\Gstarsoft')
    locationProps = @('Location','InstallPath','Path','AcadLocation')
    appDataDirs = @('Gstarsoft','GstarCAD')
  }
)

$cadPlatforms = @()
foreach ($def in $platformDefs) {
  $progPresent = [bool]((Test-Path -LiteralPath "HKLM:\SOFTWARE\Classes\$($def.progId)") -or (Test-Path -LiteralPath "HKCU:\SOFTWARE\Classes\$($def.progId)"))
  $curVerValue = $null
  if ($progPresent) { $curVerValue = CurVer $def.progId }

  $locations = @()
  $versions = @()
  foreach ($hive in @('HKLM:','HKCU:')) {
    foreach ($vendorRoot in $def.vendorRoots) {
      $rootPath = "$hive\$vendorRoot"
      if (-not (Test-Path -LiteralPath $rootPath)) { continue }
      foreach ($sub in @(Get-ChildItem -LiteralPath $rootPath -ErrorAction SilentlyContinue)) {
        $versions += [string]$sub.PSChildName
        $subProps = Get-ItemProperty -LiteralPath $sub.PSPath -ErrorAction SilentlyContinue
        foreach ($prop in $def.locationProps) {
          if ($subProps -and $subProps.$prop) { $locations += [string]$subProps.$prop }
        }
        foreach ($child in @(Get-ChildItem -LiteralPath $sub.PSPath -ErrorAction SilentlyContinue)) {
          $versions += [string]$child.PSChildName
          $childProps = Get-ItemProperty -LiteralPath $child.PSPath -ErrorAction SilentlyContinue
          foreach ($prop in $def.locationProps) {
            if ($childProps -and $childProps.$prop) { $locations += [string]$childProps.$prop }
          }
        }
      }
    }
  }
  foreach ($root in $appPathRoots) {
    foreach ($exeName in $def.exeNames) {
      $key = "$root\$exeName"
      if (Test-Path -LiteralPath $key) {
        $value = Get-RegDefault $key
        if ($value) { $locations += [string]$value }
      }
    }
  }
  $locations = @($locations | Where-Object { $_ } | Sort-Object -Unique)
  $versions = @($versions | Where-Object { $_ } | Sort-Object -Unique)
  $hasInstallPath = $false
  foreach ($location in $locations) {
    if (Test-Path -LiteralPath $location) { $hasInstallPath = $true; break }
  }
  # 仅注册了残留版本键（例如看图器/旧安装）不算已安装；需 ProgID 或可访问的安装路径
  $present = [bool]($progPresent -or $hasInstallPath)
  $cadPlatforms += [ordered]@{
    id               = $def.id
    name             = $def.name
    progId           = $def.progId
    present          = $present
    curVer           = if ($progPresent) { $curVerValue } else { $null }
    installedVersions = @($versions)
    installPaths     = @($locations)
  }
}

$pc3Files = @()
foreach ($dirName in @('Autodesk','ZWSOFT','Gstarsoft','GstarCAD')) {
  $base = Join-Path $env:APPDATA $dirName
  if ($base -and (Test-Path -LiteralPath $base)) {
    $pc3Files += @(Get-ChildItem -LiteralPath $base -Filter '*.pc3' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 40 -ExpandProperty FullName)
  }
}
$pc3Files = @($pc3Files | Where-Object { $_ } | Sort-Object -Unique | Select-Object -First 60)

$cadBest = $null
foreach ($platform in $cadPlatforms) {
  if ($platform.present) { $cadBest = $platform.id; break }
}

$result['cad'] = [ordered]@{
  any       = [bool]($cadPlatforms | Where-Object { $_.present })
  best      = $cadBest
  platforms = @($cadPlatforms)
  pc3       = @($pc3Files)
}

# ───────────────────────── P1 新增：ODAFileConverter ─────────────────────────
$odaRoots = @()
if ($env:ProgramFiles) { $odaRoots += (Join-Path $env:ProgramFiles 'ODA') }
if (${env:ProgramFiles(x86)}) { $odaRoots += (Join-Path ${env:ProgramFiles(x86)} 'ODA') }
if ($env:LOCALAPPDATA) { $odaRoots += (Join-Path $env:LOCALAPPDATA 'Programs\ODA') }
$odaRoots += 'C:\ODA'
$odaRoots = @($odaRoots | Where-Object { $_ } | Sort-Object -Unique)

$odaPaths = @()
foreach ($root in $odaRoots) {
  if (Test-Path -LiteralPath $root) {
    $odaPaths += @(Get-ChildItem -LiteralPath $root -Filter 'ODAFileConverter.exe' -Recurse -Depth 3 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
  }
}
$odaPaths = @($odaPaths | Where-Object { $_ } | Sort-Object -Unique)
$odaConverter = $null
if ($odaPaths.Count -gt 0) { $odaConverter = $odaPaths[0] }

$result['odaConverter'] = [ordered]@{
  found     = ($oda.Count -gt 0 -or $odaPaths.Count -gt 0)
  dirs      = @($oda)
  converter = $odaConverter
  paths     = @($odaPaths)
}

$result['oda'] = [ordered]@{
  found     = ($odaPaths.Count -gt 0)
  converter = $odaConverter
  paths     = @($odaPaths)
  roots     = @($odaRoots)
}

$result['os'] = [System.Environment]::OSVersion.VersionString

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$result | ConvertTo-Json -Depth 6
