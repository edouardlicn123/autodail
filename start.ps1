# 手机号拨测工具启动脚本（Windows PowerShell，含环境自举）
# 用法（在 start.ps1 所在目录打开 PowerShell）:
#   powershell -ExecutionPolicy Bypass -File .\start.ps1                  # 检查环境并启动
#   powershell -ExecutionPolicy Bypass -File .\start.ps1 -Setup           # 自动安装缺失环境后启动
#   powershell -ExecutionPolicy Bypass -File .\start.ps1 -Setup 号码.txt 9000
#   powershell -ExecutionPolicy Bypass -File .\start.ps1 9000             # 只指定端口
# 等效的批处理入口: start.bat（双击可用）

$ErrorActionPreference = "Continue"
Set-Location -LiteralPath $PSScriptRoot

# ---------- 参数解析 ----------
$Setup = $false
$NumFile = ""
$Port = 8765
foreach ($a in $args) {
    if ($a -in @('--setup', '-setup', '-Setup', '-SETUP')) { $Setup = $true }
    elseif ($a -match '^\d+$') { $Port = [int]$a }
    else { $NumFile = $a }
}

function Info($m) { Write-Host "[环境] $m" }
function Die($m) { Write-Host "[错误] $m" -ForegroundColor Red; exit 1 }

if ($NumFile -and -not (Test-Path -LiteralPath $NumFile)) { Die "号码文件不存在: $NumFile" }

# ---------- Python ----------
function Find-Python {
    foreach ($c in @('python3', 'python', 'py')) {
        $cmd = Get-Command $c -ErrorAction SilentlyContinue
        if ($cmd) {
            & $cmd.Source -c "import sys; sys.exit(0 if sys.version_info>=(3,8) else 1)" 2>$null
            if ($LASTEXITCODE -eq 0) { return $cmd.Source }
        }
    }
    $p = Get-ChildItem "$env:LOCALAPPDATA\Programs\Python\Python3*\python.exe" -ErrorAction SilentlyContinue |
        Sort-Object FullName -Descending | Select-Object -First 1
    if ($p) { return $p.FullName }
    return $null
}

$PY = Find-Python
if (-not $PY) {
    if ($Setup) {
        Info "未找到 Python，开始自动安装 3.12 ..."
        if (Get-Command winget -ErrorAction SilentlyContinue) {
            Info "使用 winget 安装 ..."
            winget install -e --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements
            $PY = Find-Python
        }
        if (-not $PY) {
            Info "改用官方安装包静默安装（用户级，自动加入 PATH）..."
            try {
                Invoke-WebRequest -Uri "https://www.python.org/ftp/python/3.12.8/amd64.exe" -OutFile "python-installer.exe"
                Start-Process -FilePath ".\python-installer.exe" -ArgumentList "/quiet", "InstallAllUsers=0", "PrependPath=1", "Include_test=0" -Wait
                Remove-Item "python-installer.exe" -ErrorAction SilentlyContinue
                $PY = Find-Python
            } catch {
                Die "下载/安装 Python 失败：$_"
            }
        }
        if (-not $PY) { Die "Python 已安装但当前会话探测不到，请重新打开 PowerShell 再运行" }
    } else {
        Die "未找到 python（需 3.8+）。加 -Setup 参数运行可自动安装"
    }
}
$ver = & $PY -c "import sys;print('%d.%d'%sys.version_info[:2])"
Info "使用 Python: $PY ($ver)"

# ---------- openpyxl ----------
& $PY -c "import openpyxl" 2>$null
if ($LASTEXITCODE -ne 0) {
    if ($Setup) {
        Info "安装 openpyxl ..."
        & $PY -m pip install --user openpyxl
        if ($LASTEXITCODE -ne 0) { Die "openpyxl 安装失败，请手动: $PY -m pip install --user openpyxl" }
        Info "openpyxl 安装完成"
    } else {
        Info "提示: 未安装 openpyxl，网页导入/回写 xlsx 将不可用。加 -Setup 参数可自动安装"
    }
}

# ---------- adb ----------
$ADB = $null
if (Get-Command adb -ErrorAction SilentlyContinue) { $ADB = (Get-Command adb).Source }
elseif (Test-Path "$PSScriptRoot\platform-tools\adb.exe") { $ADB = "$PSScriptRoot\platform-tools\adb.exe" }

if (-not $ADB) {
    if ($Setup) {
        Info "下载 Google platform-tools（含 adb.exe）..."
        try {
            Invoke-WebRequest -Uri "https://dl.google.com/android/repository/platform-tools-latest-windows.zip" -OutFile "platform-tools.zip"
            Expand-Archive -Force -Path "platform-tools.zip" -DestinationPath $PSScriptRoot
        } catch {
            Die "下载/解压 platform-tools 失败：$_"
        }
        Remove-Item "platform-tools.zip" -ErrorAction SilentlyContinue
        if (-not (Test-Path "$PSScriptRoot\platform-tools\adb.exe")) { Die "解压后未找到 adb.exe" }
        $ADB = "$PSScriptRoot\platform-tools\adb.exe"
    } else {
        Die "未安装 adb。加 -Setup 参数运行可自动安装"
    }
}
Info "使用 adb: $ADB"
$env:ADB_PATH = $ADB

# ---------- 设备检查 ----------
& $ADB get-state *> $null
if ($LASTEXITCODE -ne 0) {
    Write-Host "[警告] 未检测到手机，请检查:"
    Write-Host "       1. 数据线已连接（或 adb connect 手机IP:5555）"
    Write-Host "       2. 手机已开启 USB 调试并在弹窗中允许授权"
    $ans = Read-Host "仍要继续启动吗? [y/N]"
    if ($ans -notin @('y', 'Y')) { exit 1 }
}

if ($NumFile) { Info "预载号码文件: $NumFile" }
Info "端口: $Port"
$launchArgs = @()
if ($NumFile) { $launchArgs += $NumFile }
$launchArgs += "$Port"
& $PY phone_dial_checker.py @launchArgs
