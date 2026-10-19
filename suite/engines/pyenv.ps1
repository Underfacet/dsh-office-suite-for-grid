# P2 venv 创建与依赖安装辅助（Windows PowerShell 5.1 兼容，UTF-8 BOM）
# 输入：-Json '{"venvPath":"...","requirements":"...","python":"python","wheelsDir":"...","suiteVersion":"0.2.0"}'
# 输出：{ok:true,result:{python,pythonVersion,packages:{docx,openpyxl},created,installMode,warnings}}
#       {ok:false,error:{code:'E2002'|'E2003',message,hint?}}
# 说明：仅负责创建/修复 venv 与安装锁定依赖；锁与升级判定在 host/ops/venv.js。
[CmdletBinding()]
param([Parameter(Mandatory = $true)][string]$Json)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$env:PYTHONIOENCODING = 'utf-8'
$env:PIP_DISABLE_PIP_VERSION_CHECK = '1'

function Write-Result($obj, [int]$exitCode) {
  $obj | ConvertTo-Json -Depth 8 -Compress
  exit $exitCode
}

function Write-Failure([string]$code, [string]$message, [string]$hint) {
  $errorHash = [ordered]@{ code = $code; message = $message }
  if ($hint) { $errorHash['hint'] = $hint }
  Write-Result ([ordered]@{ ok = $false; error = $errorHash }) 1
}

function Get-Tail([string]$text, [int]$maxLength) {
  if (-not $text) { return '' }
  $trimmed = $text.Trim()
  if ($trimmed.Length -le $maxLength) { return $trimmed }
  return $trimmed.Substring($trimmed.Length - $maxLength)
}

try {
  $arg = $Json | ConvertFrom-Json
  $venvPath = [string]$arg.venvPath
  $requirements = [string]$arg.requirements
  $wheelsDir = [string]$arg.wheelsDir
  $suiteVersion = [string]$arg.suiteVersion
  $pythonArg = [string]$arg.python

  if (-not $venvPath) { Write-Failure 'E2002' 'venvPath 为空' '' }
  if (-not $requirements) { Write-Failure 'E2002' 'requirements 路径为空' '' }
  if (-not (Test-Path -LiteralPath $requirements)) {
    Write-Failure 'E2002' "requirements 文件不存在: $requirements" '请确认插件包内 suite\requirements-office.txt 存在'
  }

  # ── 1) 解析基础 Python ──
  $python = $null
  if ($pythonArg) {
    if (Test-Path -LiteralPath $pythonArg) {
      $python = (Resolve-Path -LiteralPath $pythonArg).Path
    } else {
      $cmd = Get-Command $pythonArg -ErrorAction SilentlyContinue
      if ($cmd) { $python = $cmd.Source }
    }
  } else {
    $cmd = Get-Command python -ErrorAction SilentlyContinue
    if ($cmd) { $python = $cmd.Source }
  }
  if (-not $python) {
    Write-Failure 'E2002' '未找到可用的 Python（python.exe）' '请安装 Python 3.9–3.13，或用配置 venvPython 指向便携版 python.exe'
  }

  $warnings = New-Object System.Collections.ArrayList
  $created = $false
  $venvPython = Join-Path $venvPath 'Scripts\python.exe'

  # ── 2) 创建 / 修复 venv ──
  if (-not (Test-Path -LiteralPath $venvPython)) {
    if (Test-Path -LiteralPath $venvPath) {
      # 损坏自愈：改名保留现场
      $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
      $broken = "$venvPath.broken-$stamp"
      try {
        Rename-Item -LiteralPath $venvPath -NewName (Split-Path -Leaf $broken) -ErrorAction Stop
        [void]$warnings.Add("检测到不完整的 venv，已重命名为 $broken")
      } catch {
        try { Remove-Item -LiteralPath $venvPath -Recurse -Force -ErrorAction SilentlyContinue } catch { }
      }
    }
    $parent = Split-Path -Parent $venvPath
    if ($parent -and -not (Test-Path -LiteralPath $parent)) { [void](New-Item -ItemType Directory -Path $parent -Force) }

    $createOutput = ''
    try {
      $createOutput = (& $python -m venv $venvPath 2>&1 | Out-String)
    } catch {
      $createOutput = $_.Exception.Message
    }
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $venvPython)) {
      Write-Failure 'E2002' "创建 venv 失败（$python -m venv）" (Get-Tail $createOutput 800)
    }
    $created = $true
  } else {
    [void]$warnings.Add('复用已存在的 venv 目录')
  }

  if (-not (Test-Path -LiteralPath $venvPython)) {
    Write-Failure 'E2002' "venv 创建后未找到 python.exe: $venvPython" '请检查磁盘权限与杀毒软件拦截'
  }

  # ── 3) 安装锁定依赖（离线 wheel 优先） ──
  $installMode = 'online'
  $pipArgs = @('-m', 'pip', 'install', '--no-input', '--prefer-binary')
  if ($wheelsDir -and (Test-Path -LiteralPath $wheelsDir)) {
    $wheelFiles = @(Get-ChildItem -LiteralPath $wheelsDir -Filter '*.whl' -ErrorAction SilentlyContinue)
    if ($wheelFiles.Count -gt 0) {
      $pipArgs += @('--find-links', $wheelsDir)
      $installMode = 'offline-wheels'
    }
  }
  $pipArgs += @('-r', $requirements)

  $previousEap = $ErrorActionPreference
  $ErrorActionPreference = 'Continue'
  $pipOutput = ''
  try {
    $pipOutput = (& $venvPython @pipArgs 2>&1 | Out-String)
  } catch {
    $pipOutput = $_.Exception.Message
  } finally {
    $ErrorActionPreference = $previousEap
  }
  if ($LASTEXITCODE -ne 0) {
    Write-Failure 'E2003' (
      "pip 安装依赖失败（exit $LASTEXITCODE）: $requirements"
    ) (
      (Get-Tail $pipOutput 900) + "`n" +
      '可尝试：① 把 python-docx/openpyxl/lxml 的 win_amd64 wheel 放入 runtime\wheels 后重试；' +
      '② 安装便携 Python 3.12 并用配置 venvPython 指定其 python.exe；③ 检查网络/代理（PyPI 可达性）'
    )
  }

  # ── 4) 健康校验 + 版本回读 ──
  $checkCode = "import sys,json; import docx,openpyxl; print(json.dumps({'python':sys.version.split()[0],'docx':getattr(docx,'__version__','?'),'openpyxl':openpyxl.__version__}))"
  $checkOutput = ''
  try {
    $checkOutput = (& $venvPython -c $checkCode 2>&1 | Out-String)
  } catch {
    $checkOutput = $_.Exception.Message
  }
  if ($LASTEXITCODE -ne 0) {
    Write-Failure 'E2003' '依赖安装后 import docx/openpyxl 校验失败' (Get-Tail $checkOutput 800)
  }
  $checkJson = $null
  try {
    $start = $checkOutput.IndexOf('{')
    $end = $checkOutput.LastIndexOf('}')
    if ($start -ge 0 -and $end -gt $start) {
      $checkJson = $checkOutput.Substring($start, $end - $start + 1) | ConvertFrom-Json
    }
  } catch { $checkJson = $null }
  if ($null -eq $checkJson) {
    Write-Failure 'E2003' '无法解析依赖版本校验输出' (Get-Tail $checkOutput 800)
  }

  # ── 5) 写版本标记 ──
  if ($suiteVersion) {
    try {
      [System.IO.File]::WriteAllText((Join-Path $venvPath 'SUITE_VERSION'), $suiteVersion, (New-Object System.Text.UTF8Encoding($false)))
    } catch {
      [void]$warnings.Add("写入 SUITE_VERSION 失败: $($_.Exception.Message)")
    }
  }

  $result = [ordered]@{
    ok            = $true
    result        = [ordered]@{
      python        = $venvPython
      pythonVersion = [string]$checkJson.python
      packages      = [ordered]@{
        docx     = [string]$checkJson.docx
        openpyxl = [string]$checkJson.openpyxl
      }
      created       = [bool]$created
      installMode   = $installMode
      suiteVersion  = $suiteVersion
      warnings      = @($warnings)
    }
  }
  Write-Result $result 0
}
catch {
  Write-Failure 'E2002' "venv 辅助脚本异常: $($_.Exception.Message)" $_.ScriptStackTrace
}
