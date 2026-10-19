[CmdletBinding()]
param([string]$Json = '{}')
$ErrorActionPreference = 'SilentlyContinue'
$result = [ordered]@{}

function CurVer($n) { (Get-ItemProperty "HKLM:\SOFTWARE\Classes\$n\CurVer").'(default)' }

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
$result['odaConverter'] = [ordered]@{ found = ($oda.Count -gt 0); dirs = $oda }

$result['os'] = [System.Environment]::OSVersion.VersionString

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$result | ConvertTo-Json -Depth 4
