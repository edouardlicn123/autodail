@echo off
rem 手机号拨测工具启动脚本（Windows / cmd）
rem 用法:
rem   start.bat                        检查环境后启动
rem   start.bat --setup                自动安装缺失环境(python/openpyxl/adb)后启动
rem   start.bat 号码.txt 9000           预载号码文件和端口
rem   start.bat --setup 号码.txt 9000
rem   start.bat 9000                   只指定端口
chcp 65001 >nul
setlocal EnableDelayedExpansion
cd /d "%~dp0"

set "SETUP=0"
set "NUM_FILE="
set "PORT=8765"
if /i "%~1"=="--setup" (
    set "SETUP=1"
    shift
)

if not "%~1"=="" (
    echo %~1| findstr /r "^[0-9][0-9]*$" >nul 2>&1 && (
        set "PORT=%~1"
    ) || (
        set "NUM_FILE=%~1"
        if not "%~2"=="" set "PORT=%~2"
    )
)

if defined NUM_FILE if not exist "%NUM_FILE%" (
    echo [错误] 号码文件不存在: %NUM_FILE%
    exit /b 1
)

rem ---------- Python ----------
set "PY="
for %%C in (python3 python py) do if not defined PY (
    %%C -c "import sys; sys.exit(0 if sys.version_info>=(3,8) else 1)" >nul 2>&1 && set "PY=%%C"
)
if not defined PY (
    if "%SETUP%"=="1" (
        call :install_python
    ) else (
        echo [错误] 未找到 python（需 3.8 以上版本）。运行 start.bat --setup 自动安装
        exit /b 1
    )
)
if not defined PY exit /b 1
for /f "tokens=*" %%v in ('%PY% -c "import sys;print('%d.%d'%sys.version_info[:2])"') do set "PYVER=%%v"
echo [环境] 使用 Python: %PY% [!PYVER!]

rem ---------- openpyxl ----------
"%PY%" -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    if "%SETUP%"=="1" (
        echo [环境] 安装 openpyxl ...
        "%PY%" -m pip install --user openpyxl
        "%PY%" -c "import openpyxl" >nul 2>&1 || (
            echo [错误] openpyxl 安装失败，请手动执行: "%PY%" -m pip install --user openpyxl
            exit /b 1
        )
    ) else (
        echo [环境] 提示: 未安装 openpyxl，网页导入/回写 xlsx 将不可用。用 --setup 可自动安装
    )
)

rem ---------- adb ----------
set "ADBEXE="
where adb >nul 2>&1 && set "ADBEXE=adb"
if not defined ADBEXE if exist "platform-tools\adb.exe" set "ADBEXE=%CD%\platform-tools\adb.exe"
if not defined ADBEXE (
    if "%SETUP%"=="1" (
        echo [环境] 下载 Google platform-tools（含 adb.exe）...
        curl -L -o platform-tools.zip https://dl.google.com/android/repository/platform-tools-latest-windows.zip
        if not exist platform-tools.zip (
            echo [错误] 下载失败，请检查网络（需能访问 dl.google.com）
            exit /b 1
        )
        powershell -NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive -Force 'platform-tools.zip' ."
        del platform-tools.zip >nul 2>&1
        if exist "platform-tools\adb.exe" (
            set "ADBEXE=%CD%\platform-tools\adb.exe"
        ) else (
            echo [错误] 解压后未找到 adb.exe
            exit /b 1
        )
    ) else (
        echo [错误] 未安装 adb。运行 start.bat --setup 自动安装
        exit /b 1
    )
)
set "ADB_PATH=%ADBEXE%"
echo [环境] 使用 adb: %ADBEXE%

rem ---------- 设备检查 ----------
"%ADBEXE%" get-state >nul 2>&1
if errorlevel 1 (
    echo [警告] 未检测到手机，请检查:
    echo        1. 数据线已连接（或 adb connect 手机IP:5555）
    echo        2. 手机已开启 USB 调试并在弹窗中允许授权
    set "ANS="
    set /p "ANS=仍要继续启动吗? [y/N] "
    if /i not "!ANS!"=="y" exit /b 1
)

if defined NUM_FILE echo [环境] 预载号码文件: %NUM_FILE%
echo [环境] 端口: %PORT%
"%PY%" phone_dial_checker.py %NUM_FILE% %PORT%
goto :eof

rem ---------- 子过程: 安装 Python ----------
:install_python
where winget >nul 2>&1 && (
    echo [环境] 使用 winget 安装 Python 3.12 ...
    winget install -e --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements
)
call :find_py_local
if defined PY goto :eof
echo [环境] 改用官方安装包静默安装（用户级，自动加入 PATH）...
curl -L -o python-installer.exe https://www.python.org/ftp/python/3.12.8/amd64.exe
if not exist python-installer.exe (
    echo [错误] 下载失败，请到 https://www.python.org/downloads 手动安装
    goto :eof
)
.\python-installer.exe /quiet InstallAllUsers=0 PrependPath=1 Include_test=0
del python-installer.exe >nul 2>&1
call :find_py_local
if defined PY goto :eof
echo [错误] Python 已安装但当前终端探测不到，请重新打开 cmd 后再运行 start.bat
goto :eof

:find_py_local
for /d %%P in ("%LOCALAPPDATA%\Programs\Python\Python3*") do if exist "%%P\python.exe" set "PY=%%P\python.exe"
goto :eof
