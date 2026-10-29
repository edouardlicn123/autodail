#!/usr/bin/env bash
# 手机号拨测工具启动脚本（含环境自举）
#
# 用法:
#   ./start.sh                  # 检查环境后启动；缺环境时给出提示
#   ./start.sh --setup          # 自动安装缺失环境（python3 / openpyxl / adb），完成后启动
#   ./start.sh 号码.txt 9000     # 预载号码文件和端口
#   ./start.sh --setup 号码.txt 9000
#   ./start.sh 9000             # 只指定端口
#
# 环境说明:
#   - Linux: 需要 sudo 权限安装系统包（apt/dnf/yum/pacman 自动识别）
#   - Windows: 请在 Git Bash 或 WSL 中运行本脚本；--setup 会用 winget 或
#     python.org 官方安装包装 Python，并下载 Google platform-tools 作为 adb
#
cd "$(dirname "$0")" || exit 1

SETUP=0
if [ "$1" = "--setup" ]; then
    SETUP=1
    shift
fi

# 参数解析: [号码文件] [端口] 或 [端口]
if [ -n "$1" ]; then
    if [ "$1" -eq "$1" ] 2>/dev/null; then
        PORT="$1"; NUM_FILE=""
    else
        NUM_FILE="$1"; PORT="${2:-8765}"
    fi
else
    NUM_FILE=""; PORT=8765
fi

OS_TYPE="$(uname -s)"
IS_WIN=0
case "$OS_TYPE" in
    MINGW*|MSYS*|CYGWIN*) IS_WIN=1 ;;
esac

log(){ echo "[环境] $*"; }
die(){ echo "[错误] $*" >&2; exit 1; }
have(){ command -v "$1" >/dev/null 2>&1; }

if [ -n "$NUM_FILE" ] && [ ! -f "$NUM_FILE" ]; then
    die "号码文件不存在: $NUM_FILE"
fi

# ---------------- Python ----------------

detect_python(){
    for c in python3 python py; do
        if have "$c"; then
            if "$c" -c 'import sys; sys.exit(0 if sys.version_info>=(3,8) else 1)' >/dev/null 2>&1; then
                PY="$c"
                return 0
            fi
        fi
    done
    return 1
}

install_python(){
    if [ "$IS_WIN" = 1 ]; then
        if have winget; then
            log "Windows: 使用 winget 安装 Python 3.12 ..."
            winget install -e --id Python.Python.3.12 --scope user --silent \
                --accept-package-agreements --accept-source-agreements \
                || log "winget 安装失败，尝试官方安装包..."
        fi
        if ! detect_python; then
            log "下载 Python 官方安装包并静默安装（用户级，自动加入 PATH）..."
            curl -L -o python-installer.exe https://www.python.org/ftp/python/3.12.8/amd64.exe \
                || die "下载 Python 安装包失败，请到 https://www.python.org/downloads 手动安装"
            ./python-installer.exe //quiet InstallAllUsers=0 PrependPath=1 Include_test=0 \
                || die "Python 静默安装失败，请手动运行 python-installer.exe"
            rm -f python-installer.exe
            # Git Bash 不会刷新 PATH，直接探测常见安装位置
            for p in "$LOCALAPPDATA/Programs/Python/"Python3*/python.exe /c/Python3*/python.exe; do
                if [ -x "$p" ]; then PY="$p"; return 0; fi
            done
            die "Python 已安装但当前终端看不到，请关闭并重新打开 Git Bash 后再运行 ./start.sh"
        fi
    else
        if have apt-get; then
            log "Debian/Ubuntu: 安装 python3 与 pip ..."
            sudo apt-get update && sudo apt-get install -y python3 python3-pip \
                || die "python3 安装失败，请手动安装后重试"
        elif have dnf; then
            sudo dnf install -y python3 python3-pip || die "python3 安装失败"
        elif have yum; then
            sudo yum install -y python3 python3-pip || die "python3 安装失败"
        elif have pacman; then
            sudo pacman -Sy --noconfirm python python-pip || die "python3 安装失败"
        else
            die "未识别的包管理器，请手动安装 python3 后重试"
        fi
        detect_python || die "安装后仍未检测到 python3"
    fi
}

if ! detect_python; then
    if [ "$SETUP" = 1 ]; then
        install_python
    else
        die "未找到 python3（需 >= 3.8）。运行 ./start.sh --setup 自动安装，或手动安装后重试"
    fi
fi
log "使用 Python: $PY ($("$PY" -c 'import sys;print("%d.%d"%sys.version_info[:2])'))"

# ---------------- openpyxl（xlsx 导入/回写需要） ----------------

if "$PY" -c "import openpyxl" >/dev/null 2>&1; then
    log "openpyxl 已就绪"
else
    if [ "$SETUP" = 1 ]; then
        log "安装 openpyxl ..."
        "$PY" -m pip install --user openpyxl \
            || "$PY" -m pip install --user --break-system-packages openpyxl \
            || die "openpyxl 安装失败，请手动执行: $PY -m pip install --user openpyxl"
        "$PY" -c "import openpyxl" >/dev/null 2>&1 || die "openpyxl 安装后仍不可用"
        log "openpyxl 安装完成"
    else
        log "提示: 未安装 openpyxl，网页导入/回写 xlsx 将不可用。运行 ./start.sh --setup 可自动安装"
    fi
fi

# ---------------- adb ----------------

ADB_LOCAL=""
if [ -x "./platform-tools/adb.exe" ]; then ADB_LOCAL="$(pwd)/platform-tools/adb.exe"; fi
if [ -x "./platform-tools/adb" ]; then ADB_LOCAL="$(pwd)/platform-tools/adb"; fi

install_adb(){
    if [ "$IS_WIN" = 1 ]; then
        log "下载 Google platform-tools（含 adb.exe）..."
        curl -L -o platform-tools.zip https://dl.google.com/android/repository/platform-tools-latest-windows.zip \
            || die "下载失败，请检查网络（需能访问 dl.google.com）"
        if have unzip; then unzip -oq platform-tools.zip; else
            powershell -NoProfile -Command "Expand-Archive -Force 'platform-tools.zip' ." || die "解压失败"
        fi
        rm -f platform-tools.zip
        [ -x "./platform-tools/adb.exe" ] || die "platform-tools 解压后未找到 adb.exe"
    else
        if have apt-get; then
            log "Debian/Ubuntu: 安装 adb ..."
            sudo apt-get update && sudo apt-get install -y adb
        elif have dnf; then sudo dnf install -y android-tools
        elif have yum; then sudo yum install -y android-tools
        elif have pacman; then sudo pacman -Sy --noconfirm android-tools
        fi
        if ! have adb; then
            log "包管理器未提供 adb，改用官方 platform-tools ..."
            curl -L -o platform-tools.zip https://dl.google.com/android/repository/platform-tools-latest-linux.zip \
                || die "下载失败，请检查网络"
            unzip -oq platform-tools.zip || die "解压失败（需要 unzip）"
            rm -f platform-tools.zip
        fi
        [ -x "./platform-tools/adb" ] || have adb || die "adb 安装失败"
    fi
}

if ! have adb && [ -z "$ADB_LOCAL" ]; then
    if [ "$SETUP" = 1 ]; then
        install_adb
        if [ -x "./platform-tools/adb.exe" ]; then ADB_LOCAL="$(pwd)/platform-tools/adb.exe"; fi
        if [ -x "./platform-tools/adb" ]; then ADB_LOCAL="$(pwd)/platform-tools/adb"; fi
    else
        die "未安装 adb。运行 ./start.sh --setup 自动安装，或手动安装（如 sudo apt install adb）"
    fi
fi
if [ -n "$ADB_LOCAL" ]; then
    export ADB_PATH="$ADB_LOCAL"
    log "使用 adb: $ADB_PATH"
else
    log "使用系统 adb"
fi

# ---------------- 设备检查 ----------------

ADB_PROBE="${ADB_PATH:-adb}"
if ! "$ADB_PROBE" get-state >/dev/null 2>&1; then
    echo "[警告] 未检测到手机，请检查:"
    echo "       1. 数据线已连接（或 adb connect 手机IP:5555）"
    echo "       2. 手机已开启 USB 调试并在弹窗中允许授权"
    read -r -p "仍要继续启动吗? [y/N] " ans
    [ "$ans" = "y" ] || exit 1
fi

[ -n "$NUM_FILE" ] && echo "预载号码文件: $NUM_FILE"
echo "端口: $PORT"
exec "$PY" phone_dial_checker.py ${NUM_FILE:+"$NUM_FILE"} "$PORT"
