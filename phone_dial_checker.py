#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
手机号半自动拨测工具（空号人工验证）

流程：电脑通过 ADB 控制安卓手机自动拨号 → 人耳听回铃音 → 在网页上点击判定结果（可填备注）→ 间隔倒计时后自动拨下一个。

用法：
    python3 phone_dial_checker.py [号码文件] [端口]

    号码文件     可选。每行一个号码的 txt（启动后也可在网页上从 daillist 目录导入 xlsx/csv/txt）
    端口         默认 8765（可只传端口：python3 phone_dial_checker.py 9000）

网页功能：
    - 双栏布局：左侧进度/拨测/设置，右侧导入/导出，底部全宽结果列表
    - 从 daillist/ 导入：选文件 → 工作表 → 表头行 → 电话列 / 姓名列 / 地址列 / 输出列（后三列可选）
    - 标记结果时可附备注，一并记录拨打日期时间（YYYY-MM-DD HH:MM:SS）
    - 「输出结果到表格」：把 时间+结果+备注 写回源表格的输出列，另存为 <源文件名>_结果.xlsx（不改动原文件）
    - 设置：默认电话间隔（默认 30 秒，0=标记后立即拨打下一个），持久化到 settings.json
    - 清空已测结果：设置卡片内一键清除所有标记，号码列表保留
结果自动保存到 results.csv（含姓名、地址、备注、完整时间），中途退出可续测。

可选环境变量：
    ADB_PATH     adb 可执行文件路径（默认 "adb"）
    RESULTS_CSV  结果文件路径（默认脚本目录下 results.csv）

依赖：Python 标准库；读取/写出 xlsx 需要 openpyxl（pip3 install --user --break-system-packages openpyxl）。
注意：仅连接一台手机；请控制拨号频率，避免触发运营商风控。
"""
import base64
import csv
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import time
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

try:
    import openpyxl
except ImportError:
    openpyxl = None

ADB = os.environ.get("ADB_PATH", "adb")
HOST = "127.0.0.1"
RESULT_LABELS = ["实号确认", "空号", "关机", "停机", "无法接通", "通话中", "跳过"]
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CSV_PATH = os.environ.get("RESULTS_CSV", os.path.join(BASE_DIR, "results.csv"))
DAILLIST_DIR = os.path.join(BASE_DIR, "daillist")
SETTINGS_PATH = os.path.join(BASE_DIR, "settings.json")
META_PATH = os.path.join(BASE_DIR, "import_meta.json")
TABLE_EXTS = (".xlsx", ".xlsm", ".csv", ".txt")
DEFAULT_INTERVAL = 8  # 默认电话间隔（秒）

_lock = threading.Lock()
_numbers = []   # 有序号码列表
_results = {}   # 号码 -> {"result": str, "note": str, "time": "YYYY-MM-DD HH:MM:SS"}
_cursor = None  # 手动跳转的当前指针（列表下标）；None=自动指向第一个未测号码
_context = {}   # 号码 -> {"name": 姓名, "addr": 地址}（来自导入的可选列）
IMPORT_META = {}  # 最近一次导入的 {file, sheet, header_row, phone_col, name_col, addr_col, out_col}


def load_settings():
    d = {"interval": DEFAULT_INTERVAL,
         "voice": {"dir": "/sdcard/autodail", "name": "", "auto_speaker": False,
                   "speaker": True, "sx": None, "sy": None, "ssw": None, "ssh": None}}
    try:
        with open(SETTINGS_PATH, encoding="utf-8") as f:
            v = json.load(f)
        if isinstance(v.get("interval"), (int, float)) and 0 <= v["interval"] <= 3600:
            d["interval"] = int(v["interval"])
        if isinstance(v.get("voice"), dict):
            d["voice"].update({k: v["voice"][k] for k in d["voice"] if k in v["voice"]})
    except Exception:
        pass
    return d


def save_settings():
    try:
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(SETTINGS, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


SETTINGS = load_settings()
VOICE = SETTINGS["voice"]  # {enabled, dir, name, seconds, speaker}


def save_import_meta():
    try:
        with open(META_PATH, "w", encoding="utf-8") as f:
            json.dump(IMPORT_META, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def load_import_meta():
    global IMPORT_META
    try:
        with open(META_PATH, encoding="utf-8") as f:
            v = json.load(f)
        if isinstance(v, dict):
            IMPORT_META = v
    except Exception:
        pass


# ---------------- ADB ----------------

def adb(*args, timeout=8):
    """执行 adb 命令；adb 不存在返回 None，超时返回空串。"""
    try:
        r = subprocess.run([ADB, *args], capture_output=True, text=True, timeout=timeout)
        out = (r.stdout or "").strip()
        err = (r.stderr or "").strip()
        return (out + ("\n" + err if err else "")).strip()
    except FileNotFoundError:
        return None
    except subprocess.TimeoutExpired:
        return ""


def extract_number(line):
    """从一段文本中提取电话号码（优先 11 位手机号）。"""
    text = str(line).replace("-", "").replace(" ", "").replace("\t", "")
    cands = re.findall(r"\d{7,15}", text)
    for c in cands:
        if len(c) == 11 and c.startswith("1"):
            return c
    return cands[0] if cands else None


def load_numbers(path):
    nums, seen = [], set()
    with open(path, encoding="utf-8-sig") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            n = extract_number(line)
            if n and n not in seen:
                seen.add(n)
                nums.append(n)
    return nums


# ---------------- 表格读取与导入 ----------------

def list_dial_files():
    """daillist 目录下可导入的文件（排除临时锁文件与导出副本）。"""
    if not os.path.isdir(DAILLIST_DIR):
        return []
    out = []
    for name in sorted(os.listdir(DAILLIST_DIR)):
        if name.startswith(".") or name.startswith("~"):
            continue
        if os.path.splitext(name)[1].lower() in TABLE_EXTS:
            out.append(name)
    return out


def read_table(path, sheet=None, max_rows=None):
    """读取表格，返回 (sheet名列表, 行列表)。csv/txt 的 sheet 列表为 [None]。"""
    ext = os.path.splitext(path)[1].lower()
    rows = []
    if ext in (".xlsx", ".xlsm"):
        if openpyxl is None:
            raise RuntimeError("读取 xlsx 需要 openpyxl：pip3 install --user --break-system-packages openpyxl")
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            sheets = wb.sheetnames
            ws = wb[sheet] if sheet in sheets else wb[sheets[0]]
            for row in ws.iter_rows(values_only=True):
                rows.append(["" if c is None else str(c).strip() for c in row])
                if max_rows and len(rows) >= max_rows:
                    break
        finally:
            wb.close()
        return sheets, rows
    if ext == ".csv":
        with open(path, encoding="utf-8-sig", errors="replace", newline="") as f:
            rows = list(csv.reader(f))
    else:  # txt：每行一列
        with open(path, encoding="utf-8-sig", errors="replace") as f:
            rows = [[ln.strip()] for ln in f if ln.strip() and not ln.startswith("#")]
    if max_rows:
        rows = rows[:max_rows]
    return [None], rows


def sniff_header_row(rows, scan=10):
    """在前 scan 行里找最像表头的行（0-based）。"""
    best, best_score = 0, -1
    for i, row in enumerate(rows[:scan]):
        nonempty = sum(1 for c in row if str(c).strip())
        joined = "".join(str(c) for c in row)
        bonus = 5 if any(k in joined for k in ("电话", "手机", "号码", "联系", "姓名")) else 0
        score = nonempty + bonus
        if score > best_score:
            best, best_score = i, score
    return best


def guess_phone_col(headers, data_rows):
    """猜电话列：优先'联系电话/手机'，'号码'次之（排除'身份证号码'），最后扫描数据。"""
    best, best_score = 0, 0
    for i, h in enumerate(headers):
        h = str(h)
        s = 0
        if "联系电话" in h or "手机" in h:
            s = 4
        elif "电话" in h:
            s = 3
        elif "号码" in h and "身份证" not in h:
            s = 2
        if s > best_score:
            best, best_score = i, s
    if best_score:
        return best
    for row in data_rows[:5]:
        for i, c in enumerate(row):
            v = extract_number(c)
            if v and len(v) == 11 and v.startswith("1"):
                return i
    return 0


def guess_col_by_keys(headers, rules):
    """按 关键词->得分 规则猜列，返回得分最高的列索引或 None。"""
    best, best_score = None, 0
    for i, h in enumerate(headers):
        h = str(h).strip()
        s = 0
        for keys, score in rules:
            if any(k in h for k in keys):
                s = score
                break
        if s > best_score:
            best, best_score = i, s
    return best


def guess_name_col(headers):
    return guess_col_by_keys(headers, [
        (("姓名",), 4), (("产权人",), 3), (("名字", "称呼"), 2), (("联系人",), 1),
    ])


def guess_addr_col(headers):
    return guess_col_by_keys(headers, [
        (("住址", "详址"), 4), (("地址",), 3), (("房号", "门牌"), 2),
    ])


def guess_out_col(headers):
    return guess_col_by_keys(headers, [
        (("结果", "状态"), 4), (("回访", "拨打", "核实"), 3), (("备注",), 2),
    ])


def do_import(fname, sheet, header_row, phone_col, name_col, addr_col, out_col):
    """按用户选择导入号码列表（去重），替换当前列表，并记录导入元数据。"""
    safe = os.path.basename(fname)
    path = os.path.join(DAILLIST_DIR, safe)
    if not os.path.isfile(path):
        return False, "文件不存在：" + safe
    sheets, rows = read_table(path, sheet)
    if not rows or header_row >= len(rows):
        return False, "表头行超出文件范围"
    nums, seen, ctx = [], set(), {}
    for row in rows[header_row + 1:]:
        if phone_col >= len(row):
            continue
        n = extract_number(row[phone_col])
        if not n or n in seen:
            continue
        seen.add(n)
        nums.append(n)
        rec = {}
        if name_col is not None and name_col < len(row):
            nm = str(row[name_col]).strip()
            if nm and nm != n:
                rec["name"] = nm
        if addr_col is not None and addr_col < len(row):
            ad = str(row[addr_col]).strip()
            if ad:
                rec["addr"] = ad
        if rec:
            ctx[n] = rec
    if not nums:
        return False, "没有提取到任何号码，请检查表头行和电话列的选择"
    _numbers.clear()
    _numbers.extend(nums)
    _context.clear()
    _context.update(ctx)
    IMPORT_META.clear()
    IMPORT_META.update({
        "file": safe,
        "sheet": sheet if (sheet and sheet in sheets) else (sheets[0] if sheets else None),
        "header_row": header_row,
        "phone_col": phone_col,
        "name_col": name_col,
        "addr_col": addr_col,
        "out_col": out_col,
    })
    save_import_meta()
    return True, "已导入 %d 个号码" % len(nums)


def do_export():
    """把 时间+结果+备注 写回源表格的输出列，另存为 <源文件名>_结果.xlsx。"""
    if openpyxl is None:
        return False, "写出 xlsx 需要 openpyxl", None
    meta = IMPORT_META
    if not meta.get("file"):
        return False, "尚未从 daillist 导入文件", None
    if meta.get("out_col") is None:
        return False, "导入时未选择输出列，请重新导入并勾选输出列", None
    src = os.path.join(DAILLIST_DIR, meta["file"])
    if not os.path.isfile(src):
        return False, "源文件不存在：" + meta["file"], None
    if os.path.splitext(src)[1].lower() not in (".xlsx", ".xlsm"):
        return False, "结果回写仅支持 xlsx 源文件", None
    wb = openpyxl.load_workbook(src)  # 完整加载以保留格式
    try:
        ws = wb[meta["sheet"]] if meta.get("sheet") in wb.sheetnames else wb.worksheets[0]
        header_row = int(meta.get("header_row", 0))
        phone_col = int(meta.get("phone_col", 0))
        out_col = int(meta["out_col"])
        written = 0
        for row in ws.iter_rows(min_row=header_row + 2):
            try:
                pv = row[phone_col].value
                target = row[out_col]
            except IndexError:
                continue
            n = extract_number(pv)
            if not n:
                continue
            r = _results.get(n)
            if not r:
                continue
            parts = [r["time"], r["result"]]
            if r.get("note"):
                parts.append(r["note"])
            target.value = " ".join(parts)
            written += 1
        stem, ext = os.path.splitext(src)
        out_path = stem + "_结果" + ext
        wb.save(out_path)
    finally:
        wb.close()
    return True, "已写出 %d 条到 %s" % (written, os.path.basename(out_path)), out_path


def load_prev_results():
    """续测：加载上次的结果（兼容旧格式）。"""
    if not os.path.exists(CSV_PATH):
        return
    try:
        with open(CSV_PATH, encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                n = (row.get("号码") or "").strip()
                res = (row.get("结果") or "").strip()
                if not n or not res:
                    continue
                _results[n] = {
                    "result": res,
                    "note": (row.get("备注") or "").strip(),
                    "time": (row.get("记录时间") or "").strip(),
                }
                rec = _context.setdefault(n, {})
                nm = (row.get("姓名") or "").strip()
                ad = (row.get("地址") or "").strip()
                if nm:
                    rec["name"] = nm
                if ad:
                    rec["addr"] = ad
    except Exception:
        pass


def save_csv():
    with open(CSV_PATH, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["序号", "号码", "姓名", "地址", "结果", "备注", "记录时间"])
        for i, n in enumerate(_numbers, 1):
            r = _results.get(n)
            if r:
                info = _context.get(n, {})
                w.writerow([i, n, info.get("name", ""), info.get("addr", ""),
                            r["result"], r.get("note", ""), r["time"]])


def current_index():
    global _cursor
    if _cursor is not None and 0 <= _cursor < len(_numbers):
        return _cursor
    for i, n in enumerate(_numbers):
        if n not in _results:
            return i
    return None


def dial(number):
    """直接拨出（ACTION_CALL）；失败则回退为打开拨号盘由人确认。"""
    out = adb("shell", "am", "start", "-a", "android.intent.action.CALL", "-d", "tel:" + number)
    if out is None:
        return False, "本机未安装 adb，请先安装并加入 PATH"
    low = out.lower()
    if "error" in low or "exception" in low or "not found" in low:
        out2 = adb("shell", "am", "start", "-a", "android.intent.action.DIAL", "-d", "tel:" + number)
        if out2 is None:
            return False, "本机未安装 adb，请先安装并加入 PATH"
        if "error" in out2.lower() or "exception" in out2.lower():
            return False, "拨号失败，请检查手机连接（adb devices）"
        return True, "已打开拨号盘，请在手机上点击呼叫"
    return True, "已拨出 " + number


def hangup():
    adb("shell", "input", "keyevent", "6")  # KEYCODE_ENDCALL
    return True, "已发送挂断指令"


def device_connected():
    out = adb("get-state", timeout=5)
    return out is not None and out.strip() == "device"


def call_state():
    """0=空闲 1=响铃/拨出中 2=通话中 -1=未知 -2=无adb"""
    out = adb("shell", "dumpsys", "telephony.registry")
    if out is None:
        return -2
    m = re.search(r"mCallState=(\d)", out)
    return int(m.group(1)) if m else -1


# ---------------- 语音播放 ----------------
# 说明：Android 通话上行被系统隔离，普通媒体流对方听不到。
# 可行做法 = 通话中打开免提，媒体流外放、麦克风拾音传给对方，
# 用户语音对话不受影响。播放由用户在网页按键触发，不做自动播放。

_last_call_state = -1      # 用于检测"接通"瞬间（自动开免提）
_as_thread = None          # 自动开免提线程，防止重复触发
VOICE_STATUS = ""  # 最近一次语音/免提动作的状态回显

AUDIO_EXT = (".mp3", ".wav", ".amr", ".m4a", ".aac", ".ogg", ".flac", ".3gp")


def set_status(msg):
    global VOICE_STATUS
    VOICE_STATUS = time.strftime("%H:%M:%S ") + msg


def adb_binary(*args, timeout=20):
    try:
        r = subprocess.run([ADB, *args], capture_output=True, timeout=timeout)
        return r.stdout or b""
    except Exception:
        return b""


def get_screen_size():
    out = adb("shell", "wm", "size", timeout=8) or ""
    m = re.search(r"(\d+)x(\d+)", out)
    return (int(m.group(1)), int(m.group(2))) if m else (None, None)


def ensure_call_ui():
    """若通话界面不在前台，尝试拉回（HONOR 等机型通话界面可能被按 HOME 缩到后台）。"""
    out = adb("shell", "dumpsys", "window", timeout=8) or ""
    for line in out.splitlines():
        if "mFocusedApp" in line:
            if "incallui" in line:
                return
            break
    adb("shell", "am", "start", "-n", "com.android.incallui/.InCallActivity", timeout=8)


def list_voice_files(directory):
    out = adb("shell", "ls", "-1", directory, timeout=8)
    if out is None:
        return None, "本机未安装 adb"
    if "No such file" in out:
        adb("shell", "mkdir", "-p", directory, timeout=8)  # 目录不存在则自动创建
        out = adb("shell", "ls", "-1", directory, timeout=8) or ""
    elif "Permission denied" in out:
        return None, "手机上无法读取该目录（检查路径，如 /sdcard/autodail）"
    files = sorted(l.strip() for l in out.splitlines()
                   if l.strip() and l.strip().lower().endswith(AUDIO_EXT))
    return files, None


def resolve_voice_uri(directory, name):
    """优先用 MediaStore 的 content:// URI（适配分区存储），失败退回 file://。"""
    q = adb("shell", "content", "query", "--uri", "content://media/external/audio/media",
            "--projection", "_id,_display_name", timeout=12)
    if q:
        pat = re.compile(r"_id=(\d+),\s*_display_name=" + re.escape(name) + r"(\s|$|,)")
        m = pat.search(q)
        if m:
            return "content://media/external/audio/media/" + m.group(1)
    return '"file://' + (directory.rstrip("/") + "/" + name) + '"'


def set_media_volume():
    g = adb("shell", "media", "volume", "--stream", "3", "--get", timeout=8)
    m = re.search(r"max=(\d+)", g or "")
    if m:
        adb("shell", "media", "volume", "--stream", "3", "--set", str(int(m.group(1))), timeout=8)


def toggle_speaker():
    """打开/关闭免提。
    优先用已标定坐标直接点击（部分机型如 HONOR 通话界面持续动画，
    uiautomator dump 报 could not get idle state，无法用控件定位）；
    未标定时尝试 uiautomator 关键字定位作为兜底。"""
    if VOICE.get("sx") is not None and VOICE.get("sy") is not None:
        adb("shell", "cmd", "statusbar", "collapse", timeout=6)
        ensure_call_ui()
        w, h = get_screen_size()
        w = w or VOICE.get("ssw") or 1080
        h = h or VOICE.get("ssh") or 2400
        x = int(VOICE["sx"] * w)
        y = int(VOICE["sy"] * h)
        adb("shell", "input", "tap", str(x), str(y), timeout=8)
        return True, "已点击免提键标定点 (%d,%d)" % (x, y)
    # 未标定：uiautomator 关键字兜底
    adb("shell", "uiautomator", "dump", "/sdcard/window_dump.xml", timeout=15)
    xml = adb("shell", "cat", "/sdcard/window_dump.xml", timeout=8)
    adb("shell", "rm", "/sdcard/window_dump.xml", timeout=6)
    if not xml or "could not get idle" in xml.lower():
        return False, "通话界面无法自动定位（该机型需先在网页上标定免提键位置）"
    for m in re.finditer(r"<node[^>]*>", xml):
        node = m.group(0)
        if not re.search(r"(免提|扬声器|speaker|loudspeaker)", node, re.I):
            continue
        bm = re.search(r'bounds="\[(\d+),(\d+)\]\[(\d+),(\d+)\]"', node)
        if not bm:
            continue
        x = (int(bm.group(1)) + int(bm.group(3))) // 2
        y = (int(bm.group(2)) + int(bm.group(4))) // 2
        adb("shell", "input", "tap", str(x), str(y), timeout=8)
        return True, "已点击免提键 (%d,%d)" % (x, y)
    return False, "通话界面未找到免提按钮，请在网页上标定免提键位置"


def stop_voice():
    adb("shell", "input", "keyevent", "127", timeout=8)  # KEYCODE_MEDIA_PAUSE
    return True, "已发送暂停媒体指令"


def _playcall_worker():
    """通话中按键播放：开免提 → 媒体音量最大 → 播放所选语音。"""
    try:
        if VOICE.get("speaker", True):
            ok, msg = toggle_speaker()
            set_status("开免提 " + ("成功：" if ok else "失败：") + msg)
        set_media_volume()
        name = VOICE.get("name") or ""
        if not name:
            set_status("未选择语音文件，请先在语音卡片中加载列表并选择")
            return
        uri = resolve_voice_uri(VOICE.get("dir", "/sdcard/autodail"), name)
        out = adb("shell", "am", "start", "-a", "android.intent.action.VIEW",
                  "-d", uri, "-t", "audio/*", timeout=10)
        low = (out or "").lower()
        ok = out is not None and "error" not in low and "exception" not in low
        set_status(("已播放: " if ok else "播放失败：") + name
                   + ("" if ok else "（手机未弹出播放器或文件不可读）"))
    except Exception as e:
        set_status("播放异常: %s" % e)


def play_during_call():
    threading.Thread(target=_playcall_worker, daemon=True).start()
    return True, "正在开免提并播放，结果见状态行"


def _autospeaker_worker():
    try:
        time.sleep(1.0)  # 等通话界面稳定
        ok, msg = toggle_speaker()
        set_status("拨出自动开免提 " + ("成功：" if ok else "失败：") + msg)
    except Exception as e:
        set_status("自动开免提异常: %s" % e)


def maybe_autospeaker(prev_state, now_state):
    global _as_thread
    if not VOICE.get("auto_speaker"):
        return
    # 拨出（state 1）即触发；若轮询跨过了响铃（0→2）也补一次
    hit = (now_state == 1 and prev_state != 1) or (now_state == 2 and prev_state == 0)
    if not hit:
        return
    if _as_thread and _as_thread.is_alive():
        return
    _as_thread = threading.Thread(target=_autospeaker_worker, daemon=True)
    _as_thread.start()


def export_ready():
    m = IMPORT_META
    return bool(m.get("file") and m.get("out_col") is not None
                and os.path.isfile(os.path.join(DAILLIST_DIR, m["file"]))
                and os.path.splitext(m["file"])[1].lower() in (".xlsx", ".xlsm"))


# ---------------- HTTP ----------------

INDEX_HTML = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>拨打通知电话平台</title>
<style>
:root{--bg:#0f1420;--card:#1a2233;--line:#2a3550;--txt:#e8edf7;--sub:#8b97b0;--blue:#4d7cfe;--green:#2ecc71;--red:#ff5c5c;--orange:#ffa94d;--gray:#5b6680}
*{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--txt);font-family:-apple-system,"PingFang SC","Microsoft YaHei",sans-serif;min-height:100vh;display:flex;justify-content:center;padding:24px 16px}
.wrap{width:100%;max-width:1400px}
h1{font-size:20px;margin-bottom:4px}
.sub{color:var(--sub);font-size:13px;margin-bottom:16px}
.badges{display:flex;gap:8px;margin-bottom:14px;flex-wrap:wrap}
.badge{font-size:12px;padding:4px 10px;border-radius:99px;border:1px solid var(--line);color:var(--sub)}
.badge.ok{color:#7bed9f;border-color:#2d5a3d}
.badge.bad{color:#ff9f9f;border-color:#5a2d2d}
.badge.call{color:#ffd479;border-color:#5a4a2d}
.cols2{display:grid;grid-template-columns:1.05fr .95fr;gap:14px;align-items:start}
.col{display:flex;flex-direction:column;gap:14px;min-width:0}
@media(max-width:920px){.cols2{grid-template-columns:1fr}}
.card{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:18px}
.progress-info{display:flex;justify-content:space-between;font-size:13px;color:var(--sub);margin-bottom:8px}
.bar{height:8px;background:#232d45;border-radius:99px;overflow:hidden}
.bar>div{height:100%;background:linear-gradient(90deg,#4d7cfe,#38b6ff);transition:width .3s}
.cur-label{font-size:13px;color:var(--sub);margin-bottom:6px}
.cur-num{font-size:34px;font-weight:700;letter-spacing:2px;font-variant-numeric:tabular-nums}
.cur-name{font-size:17px;color:var(--sub);font-weight:400;letter-spacing:0}
.cur-addr{font-size:13px;color:#7f8cab;font-weight:400;letter-spacing:0}
.cur-state{font-size:13px;margin-top:6px;color:var(--sub)}
.cur-state.cd{color:#ffd479}
.btns{display:flex;gap:10px;margin-top:14px;flex-wrap:wrap;align-items:center}
button{border:none;border-radius:10px;padding:11px 18px;font-size:15px;cursor:pointer;color:#fff;background:#2a3550;transition:filter .15s, transform .05s}
button:hover{filter:brightness(1.15)}button:active{transform:scale(.97)}
button:disabled{opacity:.4;cursor:not-allowed}
.call{background:var(--blue)}.hang{background:#8d6e63}.warn{background:#a94442}.export{background:#1f8a5a}
.rbtn{flex:1;min-width:0;padding:14px 2px;font-size:15px;font-weight:600}
.rbtn.primary{flex:1.9;padding:18px 2px;font-size:20px;letter-spacing:1px}
.r0{background:var(--red)}.r1{background:var(--green)}.r2{background:#e67e22}.r3{background:#d35400}.r4{background:#7f8c8d}.r5{background:#34495e}.r6{background:#3f6ad8}
.kbd{font-size:11px;opacity:.75;margin-left:4px}
label.auto{display:flex;align-items:center;gap:8px;font-size:14px;color:var(--sub);margin-top:12px;cursor:pointer}
table{width:100%;border-collapse:collapse;font-size:14px}
th,td{padding:8px 10px;text-align:left;border-bottom:1px solid var(--line)}
th{color:var(--sub);font-weight:500;font-size:12px}
td.num{font-variant-numeric:tabular-nums;letter-spacing:1px}
td.addr,td.note{max-width:110px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#aab4c8;font-size:12px}
td.time{font-size:11px;color:#8b97b0;white-space:nowrap}
th,td{padding:7px 8px}
.chip{padding:2px 10px;border-radius:99px;font-size:12px;display:inline-block}
.c-空号{background:#3d1f24;color:#ff9f9f}.c-实号,.c-实号确认{background:#1f3d29;color:#7bed9f}
.c-关机,.c-停机{background:#3d2f1f;color:#ffd479}.c-无法接通{background:#262b38;color:#aab4c8}.c-通话中{background:#1f2b4d;color:#8fb0ff}.c-跳过{background:#262b38;color:#aab4c8}
.retest{background:transparent;color:var(--sub);font-size:12px;padding:3px 8px;border:1px solid var(--line);border-radius:6px}
#toast{position:fixed;bottom:28px;left:50%;transform:translateX(-50%);background:#2a3550;color:#fff;padding:10px 18px;border-radius:10px;font-size:14px;opacity:0;transition:opacity .3s;pointer-events:none;max-width:90vw}
#toast.show{opacity:1}
.all-done{text-align:center;padding:30px 0;color:#7bed9f;font-size:18px;font-weight:600}
/* 导入与设置 */
select,input[type=number],input[type=text]{background:#232d45;color:var(--txt);border:1px solid var(--line);border-radius:8px;padding:8px 10px;font-size:14px;color-scheme:dark}
select{max-width:280px}
input[type=number]{width:110px}
input[type=text]{flex:1;min-width:180px}
input[type=text]::placeholder{color:#5b6680}
.field-row{display:flex;gap:14px;flex-wrap:wrap;margin-bottom:6px}
.field{display:flex;flex-direction:column;gap:6px;font-size:12px;color:var(--sub)}
.cols{font-size:13px;margin-top:10px}
.cols th,.cols td{padding:5px 8px}
.cols input[type=radio]{accent-color:var(--blue)}
.cols th label{font-weight:400;color:#8b97b0;white-space:nowrap}
.prev-wrap{max-height:150px;overflow:auto;margin-top:8px;border:1px solid var(--line);border-radius:8px}
.prev-wrap table{font-size:12px}
.prev-wrap td{padding:4px 8px;white-space:nowrap;color:#aab4c8}
.imp-hint{font-size:12px;color:var(--sub)}
#settings-card .cur-label{cursor:pointer;user-select:none}
#voice-card .cur-label{cursor:pointer;user-select:none}
.num-link{cursor:pointer;color:#8fb0ff;text-decoration:none}
.num-link:hover{text-decoration:underline}
#calib-overlay{position:fixed;inset:0;background:rgba(0,0,0,.78);z-index:50;display:none;align-items:center;justify-content:center}
#calib-box{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:14px;max-height:94vh;overflow:auto;max-width:92vw}
#calib-imgwrap{position:relative;display:inline-block;line-height:0}
#calib-img{max-height:62vh;display:block;border-radius:8px;cursor:crosshair}
#calib-dot{position:absolute;width:16px;height:16px;border:2px solid var(--red);border-radius:50%;transform:translate(-50%,-50%);pointer-events:none;display:none;background:rgba(255,92,92,.25)}
</style>
</head>
<body>
<div class="wrap">
  <h1>拨打通知电话平台</h1>
  <div class="sub">电脑拨号 · 人耳判断 · 点击记录 · 间隔自动下一个 ｜ 结果保存于 results.csv</div>

  <div class="badges">
    <span id="dev-badge" class="badge">设备检测中…</span>
    <span id="call-badge" class="badge">通话状态：—</span>
    <span class="badge">快捷键：1-7 标记 · D 拨打 · H 挂断</span>
  </div>

  <div class="cols2">
    <!-- 左列 -->
    <div class="col">
      <div class="card">
        <div class="progress-info"><span id="prog-text">—</span><span id="prog-pct"></span></div>
        <div class="bar"><div id="prog-bar" style="width:0%"></div></div>
      </div>

      <div class="card">
        <div id="cur-area">
          <div class="cur-label">当前号码</div>
          <div class="cur-num"><span id="cur-num">—</span> <span class="cur-name" id="cur-name"></span> <span class="cur-addr" id="cur-addr"></span></div>
          <div class="cur-state" id="cur-state"></div>
          <div class="btns">
            <button class="call" id="btn-dial" onclick="dial()">拨打 <span class="kbd">D</span></button>
            <button class="hang" id="btn-hang" onclick="hangup()">挂断 <span class="kbd">H</span></button>
            <span style="margin-left:auto;display:flex;gap:10px">
              <button onclick="voicePlayCall()" style="background:#3f6ad8;color:#fff;border:none;border-radius:8px;padding:12px 18px;cursor:pointer">播放语音</button>
              <button onclick="voiceSpeaker()">切换免提</button>
            </span>
          </div>
        </div>
        <div id="done-area" class="all-done" style="display:none">全部号码已测完 🎉</div>
        <div class="btns" id="result-btns"></div>
        <input type="text" id="note-input" placeholder="备注（选填）：随结果一起保存，导出时一并写入表格" style="width:100%;margin-top:12px">
        <div style="display:flex;flex-wrap:wrap;gap:6px 22px;margin-top:12px;align-items:center">
          <label class="auto" style="margin-top:0"><input type="checkbox" id="auto-next" checked> 挂断后按设定间隔自动打下一个</label>
          <label class="auto" style="margin-top:0"><input type="checkbox" id="auto-spk" onchange="autoSpeaker(this)"> 拨出后自动开免提（未接通即开）</label>
        </div>
      </div>

      <div class="card" id="voice-card">
        <div class="cur-label" onclick="toggleVoice()">语音播放（接通后播语音，同时保持通话） ▸</div>
        <div id="voice-body" style="display:none">
          <div class="field-row">
            <div class="field">手机里的语音目录
              <div style="display:flex;gap:6px">
                <input type="text" id="voice-dir" placeholder="/sdcard/autodail" style="min-width:150px">
                <button onclick="loadVoiceList()" style="padding:8px 12px">加载列表</button>
              </div>
            </div>
            <div class="field">语音文件
              <select id="voice-file" style="max-width:200px"><option value="">（先加载列表）</option></select>
            </div>
          </div>
          <div class="btns">
            <button class="call" onclick="voicePlay()">试播</button>
            <button class="hang" onclick="voiceStop()">停止播放</button>
            <button onclick="openCalib()">标定免提键</button>
            <span class="imp-hint" id="calib-state"></span>
          </div>
          <div class="imp-hint" style="margin-top:6px" id="voice-status"></div>
          <div class="imp-hint" style="margin-top:4px">说明：通话中点「播放语音」= 确保通话界面在前台 → 开免提 → 媒体音量最大 → 播放所选语音，可随时点「停止播放」；免提键标定一次即可（换手机或改分辨率后需重标）。</div>
        </div>
      </div>

      <div class="card" id="settings-card">
        <div class="cur-label" onclick="toggleSettings()">设置 ▸</div>
        <div id="settings-body" style="display:none">
          <div class="field-row">
            <div class="field">默认电话间隔（秒，0 = 立即拨打下一个）
              <input type="number" id="set-interval" min="0" max="3600" step="1">
            </div>
            <button onclick="saveSettings()" style="align-self:flex-end">保存设置</button>
          </div>
          <div class="btns" style="margin-top:6px">
            <button class="warn" onclick="clearResults()">清空已测结果</button>
            <span class="imp-hint">清除所有标记记录，号码列表与导入信息保留</span>
          </div>
        </div>
      </div>

      <div class="card">
        <div class="cur-label" onclick="toggleFilter()">过滤 <span id="filter-summary" style="color:#5b6680;font-weight:400;font-size:12px"></span> ▸</div>
        <div id="filter-body" style="display:none">
          <div id="filter-box" style="display:flex;flex-wrap:wrap;gap:6px 18px;padding:8px 0 2px"></div>
        </div>
        <div style="overflow-x:auto">
          <table style="min-width:520px">
            <thead><tr><th style="width:40px">序号</th><th>号码</th><th>姓名</th><th>地址</th><th>结果</th><th>备注</th><th>拨打时间</th><th style="width:60px"></th></tr></thead>
            <tbody id="tbody"></tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- 右列 -->
    <div class="col">
      <div class="card">
        <div class="cur-label">导出结果</div>
        <div class="btns" style="margin-top:0">
          <button class="export" id="btn-export" onclick="doExport()">输出结果到表格</button>
          <span class="imp-hint" id="export-hint"></span>
        </div>
      </div>

      <div class="card">
        <div class="cur-label">导入号码（来自 daillist/ 目录）</div>
        <div class="field-row">
          <div class="field">文件
            <div style="display:flex;gap:6px">
              <select id="imp-file" onchange="onFileChange()"></select>
              <button onclick="loadFiles()" style="padding:8px 12px">刷新</button>
            </div>
          </div>
          <div class="field" id="sheet-field" style="display:none">工作表
            <select id="imp-sheet" onchange="preview()"></select>
          </div>
          <div class="field" id="hdr-field" style="display:none">表头行
            <select id="imp-hdr" onchange="preview()"></select>
          </div>
        </div>
        <div id="imp-cols"></div>
        <div class="prev-wrap" id="imp-prev" style="display:none"></div>
        <div class="btns">
          <button class="call" id="btn-import" onclick="doImport()">导入</button>
          <span class="imp-hint" id="imp-hint"></span>
        </div>
      </div>
    </div>
  </div>
</div>
<div id="toast"></div>
<div id="calib-overlay">
  <div id="calib-box">
    <div style="display:flex;justify-content:space-between;align-items:center;gap:12px;margin-bottom:10px">
      <b>标定免提键：点击截图上「免提」按钮的位置</b>
      <button onclick="closeCalib()" style="padding:6px 14px">关闭</button>
    </div>
    <div id="calib-imgwrap"><img id="calib-img"><div id="calib-dot"></div></div>
    <div class="btns">
      <button class="call" onclick="saveCalib()">保存标定位置</button>
      <span class="imp-hint" id="calib-hint">请先把手机停在通话界面再打开本窗口</span>
    </div>
  </div>
</div>

<script>
const COLORS = {"实号确认":"r1","空号":"r0","关机":"r2","停机":"r3","无法接通":"r4","通话中":"r6","跳过":"r5"};
let S = null, PREVIEW = null, toastTimer = null, cdTimer = null;
let FILTER = null;

function esc(s){ return String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }
function toast(msg){
  const t = document.getElementById('toast');
  t.textContent = msg; t.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(()=>t.classList.remove('show'), 3000);
}
async function post(url, body){
  const r = await fetch(url, {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(body||{})});
  return await r.json();
}
async function refresh(){
  try{ S = await (await fetch('/api/state')).json(); }catch(e){ return; }
  render();
}

/* ---------- 导入 ---------- */
async function loadFiles(){
  let d;
  try{ d = await (await fetch('/api/files')).json(); }catch(e){ return; }
  const sel = document.getElementById('imp-file');
  const cur = sel.value;
  sel.innerHTML = d.files.map(f=>`<option value="${esc(f)}">${esc(f)}</option>`).join('');
  document.getElementById('imp-hint').textContent = d.files.length ? '' : 'daillist/ 目录下没有可导入的文件（支持 xlsx/csv/txt）';
  if(!d.files.length){ document.getElementById('imp-cols').innerHTML=''; return; }
  if(d.files.includes(cur)) sel.value = cur;
  onFileChange();
}
function onFileChange(){
  document.getElementById('imp-hdr').value = '';
  document.getElementById('imp-sheet').value = '';
  preview();
}
async function preview(){
  const f = document.getElementById('imp-file').value;
  if(!f) return;
  const hdrEl = document.getElementById('imp-hdr');
  const body = {file: f};
  const sheet = document.getElementById('imp-sheet').value;
  if(sheet) body.sheet = sheet;
  if(hdrEl.value !== '') body.header_row = parseInt(hdrEl.value);
  try{ PREVIEW = await post('/api/preview', body); }catch(e){ return; }
  if(!PREVIEW.ok){ toast(PREVIEW.msg||'预览失败'); return; }
  const sf = document.getElementById('sheet-field');
  if(PREVIEW.sheets && PREVIEW.sheets.length > 1){
    sf.style.display = '';
    document.getElementById('imp-sheet').innerHTML =
      PREVIEW.sheets.map(s=>`<option value="${esc(s)}" ${s===PREVIEW.sheet?'selected':''}>${esc(s)}</option>`).join('');
  } else { sf.style.display = 'none'; }
  hdrEl.innerHTML = '';
  for(let i=0;i<10;i++){
    hdrEl.innerHTML += `<option value="${i}" ${i===PREVIEW.header_row?'selected':''}>第 ${i+1} 行</option>`;
  }
  hdrEl.parentElement.style.display = '';
  renderCols();
}
function colRadio(name, val, checked){
  return `<input type="radio" name="${name}" value="${val}" ${checked?'checked':''}>`;
}
function renderCols(){
  const box = document.getElementById('imp-cols');
  const pv = document.getElementById('imp-prev');
  const hs = PREVIEW.headers || [];
  if(!hs.length){ box.innerHTML=''; pv.style.display='none'; return; }
  let html = '<table class="cols"><thead><tr><th>列名</th><th style="width:64px">电话列</th>'
    + `<th>姓名列 <label>${colRadio('ncol', -1, PREVIEW.guess_name_col===null)}不选</label></th>`
    + `<th>地址列 <label>${colRadio('acol', -1, PREVIEW.guess_addr_col===null)}不选</label></th>`
    + `<th>输出列 <label>${colRadio('ocol', -1, PREVIEW.guess_out_col===null)}不选</label></th></tr></thead><tbody>`;
  hs.forEach((h,i)=>{
    const label = esc(String(h).replace(/\\n/g,' ').trim()) || `<span style="color:#5b6680">(空列名 ${i})</span>`;
    html += `<tr><td>${label}</td>`
         +  `<td>${colRadio('pcol', i, i===PREVIEW.guess_phone_col)}</td>`
         +  `<td>${colRadio('ncol', i, i===PREVIEW.guess_name_col)}</td>`
         +  `<td>${colRadio('acol', i, i===PREVIEW.guess_addr_col)}</td>`
         +  `<td>${colRadio('ocol', i, i===PREVIEW.guess_out_col)}</td></tr>`;
  });
  html += '</tbody></table>';
  box.innerHTML = html;
  const samples = PREVIEW.samples || [];
  if(samples.length){
    pv.style.display = '';
    let phtml = '<table><tbody>';
    samples.forEach(r=>{
      phtml += '<tr>' + hs.map((_,i)=>`<td>${esc((r[i]??''))}</td>`).join('') + '</tr>';
    });
    pv.innerHTML = phtml + '</tbody></table>';
  } else pv.style.display='none';
}
async function doImport(){
  const f = document.getElementById('imp-file').value;
  if(!f) return;
  const p = document.querySelector('input[name=pcol]:checked');
  const n = document.querySelector('input[name=ncol]:checked');
  const a = document.querySelector('input[name=acol]:checked');
  const o = document.querySelector('input[name=ocol]:checked');
  if(!p){ toast('请先选择电话列'); return; }
  const body = {
    file: f,
    sheet: document.getElementById('imp-sheet').value || null,
    header_row: parseInt(document.getElementById('imp-hdr').value),
    phone_col: parseInt(p.value),
    name_col: n && n.value !== '-1' ? parseInt(n.value) : null,
    addr_col: a && a.value !== '-1' ? parseInt(a.value) : null,
    out_col: o && o.value !== '-1' ? parseInt(o.value) : null,
  };
  document.getElementById('btn-import').disabled = true;
  const r = await post('/api/import', body);
  document.getElementById('btn-import').disabled = false;
  if(r.ok){
    toast(r.msg + (body.out_col!==null ? '（可随时点「输出结果到表格」回写）' : ''));
    refresh();
  } else toast(r.msg||'导入失败');
}

/* ---------- 导出 ---------- */
async function doExport(){
  document.getElementById('btn-export').disabled = true;
  const r = await post('/api/export');
  document.getElementById('btn-export').disabled = false;
  toast(r.ok ? r.msg : (r.msg||'导出失败'));
  refresh();
}

/* ---------- 设置 ---------- */
function toggleSettings(){
  const b = document.getElementById('settings-body');
  const card = document.querySelector('#settings-card .cur-label');
  const show = b.style.display === 'none';
  b.style.display = show ? '' : 'none';
  card.textContent = show ? '设置 ▾' : '设置 ▸';
}
async function saveSettings(){
  const iv = parseInt(document.getElementById('set-interval').value);
  if(isNaN(iv) || iv < 0 || iv > 3600){ toast('间隔需为 0-3600 的整数秒'); return; }
  const r = await post('/api/settings', {interval: iv});
  if(r.ok) toast('已保存：间隔 ' + r.interval + ' 秒');
  else toast(r.msg||'保存失败');
  refresh();
}
async function clearResults(){
  if(!confirm('确定清空所有已测结果吗？')) return;
  const r = await post('/api/clear_results');
  if(r.ok) toast('已清空 ' + r.cleared + ' 条结果');
  refresh();
}

/* ---------- 语音播放 ---------- */
function toggleVoice(){
  const b = document.getElementById('voice-body');
  b.style.display = (b.style.display==='none') ? '' : 'none';
}
async function loadVoiceList(){
  const dir = document.getElementById('voice-dir').value.trim() || '/sdcard/autodail';
  const r = await post('/api/voice', {action:'list', dir});
  if(!r.ok){ toast(r.msg); return; }
  const sel = document.getElementById('voice-file');
  const keep = (S&&S.voice&&S.voice.name) || '';
  sel.innerHTML = r.files.length
    ? r.files.map(f=>`<option value="${esc(f)}">${esc(f)}</option>`).join('')
    : '<option value="">（目录中没有音频文件）</option>';
  if(r.files.includes(keep)) sel.value = keep;
  toast('已加载 ' + r.files.length + ' 个语音文件');
}
async function voicePlay(){
  const name = document.getElementById('voice-file').value;
  const r = await post('/api/voice', {action:'play', name});
  toast(r.msg||'');
}
async function voiceStop(){
  const r = await post('/api/voice', {action:'stop'});
  toast(r.msg||'');
}
async function voiceSpeaker(){
  const r = await post('/api/voice', {action:'speaker'});
  toast(r.msg||'');
}
let CALIB = null;
async function openCalib(){
  document.getElementById('calib-overlay').style.display = 'flex';
  document.getElementById('calib-hint').textContent = '正在截取手机屏幕…';
  CALIB = {};
  const r = await post('/api/voice', {action:'shot'});
  if(!r.ok){ document.getElementById('calib-hint').textContent = r.msg; return; }
  CALIB.w = r.w; CALIB.h = r.h;
  document.getElementById('calib-img').src = 'data:image/png;base64,' + r.image;
  const v = S && S.voice || {};
  if(v.sx != null){
    const dot = document.getElementById('calib-dot');
    dot.style.display = 'block';
    dot.style.left = (v.sx*100)+'%'; dot.style.top = (v.sy*100)+'%';
    document.getElementById('calib-hint').textContent = '已显示当前标定点，点击新位置可覆盖';
  } else {
    document.getElementById('calib-hint').textContent = '点击截图中「免提」按钮的位置';
  }
}
document.getElementById('calib-imgwrap').addEventListener('click', e=>{
  if(e.target.id !== 'calib-img' || !CALIB) return;
  const rect = e.target.getBoundingClientRect();
  const x = (e.clientX-rect.left)/rect.width, y = (e.clientY-rect.top)/rect.height;
  CALIB.x = x; CALIB.y = y;
  const dot = document.getElementById('calib-dot');
  dot.style.display = 'block';
  dot.style.left = (x*100)+'%'; dot.style.top = (y*100)+'%';
  document.getElementById('calib-hint').textContent =
    '已选择 (' + Math.round(x*100) + '%, ' + Math.round(y*100) + '%)，点击「保存标定位置」生效';
});
async function saveCalib(){
  if(!(CALIB && CALIB.x != null)){ toast('请先在截图上点击免提按钮位置'); return; }
  const r = await post('/api/voice', {action:'calibrate', x:CALIB.x, y:CALIB.y});
  toast(r.msg||'');
  if(r.ok) closeCalib();
}
function closeCalib(){ document.getElementById('calib-overlay').style.display='none'; }
async function voicePlayCall(){
  const name = document.getElementById('voice-file').value;
  if(!name){ toast('请先在语音卡片中加载列表并选择语音文件'); return; }
  post('/api/voice', {action:'save', name:name, dir:document.getElementById('voice-dir').value.trim()});
  const r = await post('/api/voice', {action:'playcall'});
  toast(r.msg||'');
}
async function autoSpeaker(cb){
  const r = await post('/api/voice', {action:'save', auto_speaker: cb.checked});
  if(!r.ok){ cb.checked = false; toast(r.msg||'设置失败'); return; }
  toast(cb.checked ? '已开启：拨出后立即开免提（需已标定免提键）' : '已关闭拨出自动开免提');
}

/* ---------- 拨测 ---------- */
function stopCountdown(){
  if(cdTimer){ clearInterval(cdTimer); cdTimer = null; }
  const el = document.getElementById('cur-state');
  el.textContent = ''; el.classList.remove('cd');
}
function startCountdown(sec){
  stopCountdown();
  const el = document.getElementById('cur-state');
  el.classList.add('cd');
  let left = sec;
  const tick = ()=>{
    if(left <= 0){ stopCountdown(); dial(); return; }
    el.textContent = left + ' 秒后自动拨打下一个（点「拨打」立即开始）';
    left--;
  };
  tick();
  cdTimer = setInterval(tick, 1000);
}
function render(){
  const done = Object.keys(S.results).length;
  const total = S.total;
  document.getElementById('prog-text').textContent = `已完成 ${done} / ${total}`;
  const pct = total ? Math.round(done/total*100) : 0;
  document.getElementById('prog-pct').textContent = pct + '%';
  document.getElementById('prog-bar').style.width = pct + '%';

  const dev = document.getElementById('dev-badge');
  dev.textContent = S.device_connected ? '设备已连接' : '设备未连接（检查 USB 调试 / adb devices）';
  dev.className = 'badge ' + (S.device_connected ? 'ok' : 'bad');
  const cs = {'0':'空闲','1':'拨出/响铃中','2':'通话中','-1':'未知','-2':'无 adb'};
  const cb = document.getElementById('call-badge');
  cb.textContent = '通话状态：' + (cs[S.call_state] ?? '—');
  cb.className = 'badge ' + (S.call_state===0 ? '' : 'call');

  const iv = document.getElementById('set-interval');
  if(document.activeElement !== iv && S.interval !== undefined) iv.value = S.interval;
  const v = S.voice || {};
  const vd = document.getElementById('voice-dir');
  if(document.activeElement !== vd && v.dir) vd.value = v.dir;
  document.getElementById('auto-spk').checked = !!v.auto_speaker;
  document.getElementById('calib-state').textContent =
    (v.sx != null) ? '免提键已标定 (' + Math.round(v.sx*100) + '%, ' + Math.round(v.sy*100) + '%)' : '免提键未标定';
  document.getElementById('voice-status').textContent = S.voice_status || '';

  const eb = document.getElementById('btn-export');
  const eh = document.getElementById('export-hint');
  eb.disabled = !S.export_ready;
  eh.textContent = S.export_ready
    ? `把「拨打时间 + 结果 + 备注」写入 ${esc(S.export_file)} 的输出列，另存为 _结果.xlsx`
    : '导入时选择输出列后可用（回写到源表格副本，不改原文件）';

  const cur = S.current_number;
  document.getElementById('cur-area').style.display = cur ? '' : 'none';
  document.getElementById('done-area').style.display = (!cur && total>0) ? '' : 'none';
  document.getElementById('cur-num').textContent = cur || '—';
  const info = cur && S.context ? (S.context[cur]||{}) : {};
  document.getElementById('cur-name').textContent = info.name ? '（'+info.name+'）' : '';
  document.getElementById('cur-addr').textContent = info.addr ? '📍'+info.addr : '';

  const rb = document.getElementById('result-btns');
  rb.innerHTML = '';
  S.labels.forEach((label,i)=>{
    const b = document.createElement('button');
    b.className = 'rbtn ' + (COLORS[label]||'r5') + (i===0?' primary':'');
    b.innerHTML = label + ' <span class="kbd">' + (i+1) + '</span>';
    b.disabled = !cur;
    b.onclick = ()=>mark(label);
    rb.appendChild(b);
  });
  document.getElementById('btn-dial').disabled = !cur;
  document.getElementById('btn-hang').disabled = !S.device_connected;

  if(!FILTER){
    FILTER = loadFilter() || defaultFilter();
    renderFilter();
  }
  renderTable();
}
function defaultFilter(){
  const f = {'未测':true};
  S.labels.forEach(l=>f[l]=true);
  f['实号确认']=false;
  return f;
}
function loadFilter(){
  try{
    const v = JSON.parse(localStorage.getItem('dl_filter_v1'));
    if(v && typeof v==='object'){ const m = defaultFilter(); Object.keys(v).forEach(k=>m[k]=v[k]); return m; }
  }catch(e){}
  return null;
}
function saveFilter(){ try{ localStorage.setItem('dl_filter_v1', JSON.stringify(FILTER)); }catch(e){} }
function toggleFilter(){
  const b = document.getElementById('filter-body');
  b.style.display = (b.style.display==='none') ? '' : 'none';
}
function renderFilter(){
  const cats = [...S.labels, '未测'];
  document.getElementById('filter-box').innerHTML = cats.map(c=>{
    const on = FILTER[c] !== false;
    return `<label class="auto" style="margin-top:0"><input type="checkbox" ${on?'checked':''} onchange="setFilter('${esc(c)}', this.checked)"> ${esc(c)}</label>`;
  }).join('');
  const hidden = cats.filter(c=>FILTER[c]===false);
  document.getElementById('filter-summary').textContent = hidden.length ? '已隐藏: '+hidden.join('、') : '';
}
function setFilter(cat, on){
  FILTER[cat] = on; saveFilter(); renderFilter(); renderTable();
}
function renderTable(){
  const tb = document.getElementById('tbody');
  tb.innerHTML = '';
  if(!S || !S.ordered) return;
  S.ordered.forEach((n,i)=>{
    const r = S.results[n];
    const cat = r ? r.result : '未测';
    if(FILTER && FILTER[cat]===false) return;
    const info = (S.context && S.context[n]) || {};
    const tr = document.createElement('tr');
    tr.innerHTML = `<td>${i+1}</td><td class="num"><a class="num-link" title="点击定位到该号码，点「拨打」开始" onclick="jumpTo('${esc(n)}')">${esc(n)}</a></td><td style="color:#aab4c8">${esc(info.name||'')}</td><td class="addr" title="${esc(info.addr||'')}">${esc(info.addr||'')}</td>
      <td>${r?`<span class="chip c-${esc(r.result)}">${esc(r.result)}</span>`:'<span style="color:#5b6680">未测</span>'}</td>
      <td class="note" title="${esc(r&&r.note||'')}">${esc(r&&r.note||'')}</td>
      <td class="time">${r?esc(r.time):''}</td>
      <td>${r?`<button class="retest" onclick="unmark('${esc(n)}')">重测</button>`:''}</td>`;
    if(n===S.current_number) tr.style.background = 'rgba(77,124,254,.08)';
    tb.appendChild(tr);
  });
}
async function dial(number){
  stopCountdown();
  const r = await post('/api/dial', number?{number}:{});
  if(!r.ok) toast(r.msg);
  refresh();
}
async function hangup(){
  await post('/api/hangup'); refresh();
}
async function mark(result){
  if(!S.current_number) return;
  const note = document.getElementById('note-input').value.trim();
  const r = await post('/api/mark', {result, note});
  if(!r.ok){ toast(r.msg||'记录失败'); refresh(); return; }
  document.getElementById('note-input').value = '';
  if(document.getElementById('auto-next').checked && r.next_number){
    await post('/api/hangup');
    const iv = (S && S.interval !== undefined) ? S.interval : 30;
    if(iv > 0) startCountdown(iv);
    else setTimeout(()=>dial(), 500);
  }
  refresh();
}
async function unmark(number){
  await post('/api/unmark', {number});
  refresh();
}
async function jumpTo(number){
  const r = await post('/api/jump', {number});
  toast(r.ok ? ('已定位到 ' + number + '，点「拨打」开始') : (r.msg || '跳转失败'));
}
document.addEventListener('keydown', e=>{
  const tag = (e.target && e.target.tagName) || '';
  if(tag === 'INPUT' || tag === 'SELECT' || tag === 'TEXTAREA') return; // 输入框内不触发快捷键
  if(!S) return;
  const i = parseInt(e.key);
  if(i>=1 && i<=S.labels.length && S.current_number) mark(S.labels[i-1]);
  else if(e.key==='d'||e.key==='D'){ if(S.current_number) dial(); }
  else if(e.key==='h'||e.key==='H') hangup();
});
loadFiles();
refresh();
setInterval(refresh, 1500);
</script>
</body>
</html>
"""


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _json(self, obj, code=200):
        data = json.dumps(obj, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _body(self):
        n = int(self.headers.get("Content-Length") or 0)
        if n <= 0:
            return {}
        try:
            return json.loads(self.rfile.read(n).decode())
        except Exception:
            return {}

    def do_GET(self):
        global _last_call_state
        if self.path == "/api/state":
            with _lock:
                idx = current_index()
                cs_now = call_state()
                cs_prev = _last_call_state
                _last_call_state = cs_now
                payload = {
                    "labels": RESULT_LABELS,
                    "total": len(_numbers),
                    "ordered": list(_numbers),
                    "context": _context,
                    "results": _results,
                    "current_index": idx,
                    "current_number": _numbers[idx] if idx is not None else None,
                    "device_connected": device_connected(),
                    "call_state": cs_now,
                    "interval": SETTINGS["interval"],
                    "voice": dict(VOICE),
                    "voice_status": VOICE_STATUS,
                    "export_ready": export_ready(),
                    "export_file": IMPORT_META.get("file"),
                }
            maybe_autospeaker(cs_prev, cs_now)
            self._json(payload)
        elif self.path == "/api/files":
            self._json({"files": list_dial_files()})
        elif self.path == "/":
            data = INDEX_HTML.encode()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        else:
            self._json({"error": "not found"}, 404)

    def do_POST(self):
        body = self._body()
        if self.path == "/api/preview":
            fname = os.path.basename((body.get("file") or "").strip())
            path = os.path.join(DAILLIST_DIR, fname)
            if not fname or not os.path.isfile(path):
                self._json({"ok": False, "msg": "文件不存在"})
                return
            try:
                sheets, rows = read_table(path, body.get("sheet") or None, max_rows=12)
            except Exception as e:
                self._json({"ok": False, "msg": "读取失败：%s" % e})
                return
            hdr = body.get("header_row")
            hdr = int(hdr) if hdr is not None else sniff_header_row(rows)
            hdr = max(0, min(hdr, max(0, len(rows) - 1)))
            headers = rows[hdr] if rows else []
            data = rows[hdr + 1:]
            self._json({
                "ok": True,
                "sheets": sheets,
                "sheet": body.get("sheet") or (sheets[0] if sheets else None),
                "header_row": hdr,
                "headers": headers,
                "samples": data[:3],
                "guess_phone_col": guess_phone_col(headers, data),
                "guess_name_col": guess_name_col(headers),
                "guess_addr_col": guess_addr_col(headers),
                "guess_out_col": guess_out_col(headers),
            })
        elif self.path == "/api/import":
            fname = os.path.basename((body.get("file") or "").strip())
            try:
                header_row = int(body.get("header_row") or 0)
                phone_col = int(body.get("phone_col") or 0)

                def opt(v):
                    return int(v) if v is not None else None

                name_col = opt(body.get("name_col"))
                addr_col = opt(body.get("addr_col"))
                out_col = opt(body.get("out_col"))
            except (TypeError, ValueError):
                self._json({"ok": False, "msg": "参数错误"})
                return
            with _lock:
                try:
                    ok, msg = do_import(fname, body.get("sheet") or None, header_row,
                                        phone_col, name_col, addr_col, out_col)
                except Exception as e:
                    ok, msg = False, "导入失败：%s" % e
                if ok:
                    save_csv()
                    globals()["_cursor"] = None  # 新导入从头开始
                pending = sum(1 for n in _numbers if n not in _results)
            self._json({"ok": ok, "msg": msg, "total": len(_numbers), "pending": pending})
        elif self.path == "/api/export":
            with _lock:
                try:
                    ok, msg, out_path = do_export()
                except Exception as e:
                    ok, msg, out_path = False, "导出失败：%s" % e, None
            self._json({"ok": ok, "msg": msg, "path": out_path})
        elif self.path == "/api/settings":
            try:
                iv = int(body.get("interval"))
            except (TypeError, ValueError):
                self._json({"ok": False, "msg": "间隔需为整数秒"})
                return
            if iv < 0 or iv > 3600:
                self._json({"ok": False, "msg": "间隔需在 0-3600 秒之间"})
                return
            SETTINGS["interval"] = iv
            save_settings()
            self._json({"ok": True, "interval": iv})
        elif self.path == "/api/clear_results":
            with _lock:
                cleared = len(_results)
                _results.clear()
                save_csv()
                globals()["_cursor"] = None
            self._json({"ok": True, "cleared": cleared})
        elif self.path == "/api/dial":
            with _lock:
                idx = current_index()
                number = (body.get("number") or "").strip() or (_numbers[idx] if idx is not None else None)
            if not number:
                self._json({"ok": False, "msg": "没有待测号码，请先在上方导入"})
                return
            ok, msg = dial(number)
            self._json({"ok": ok, "msg": msg, "number": number})
        elif self.path == "/api/hangup":
            ok, msg = hangup()
            self._json({"ok": ok, "msg": msg})
        elif self.path == "/api/mark":
            result = (body.get("result") or "").strip()
            note = str(body.get("note") or "").strip()[:200]
            with _lock:
                idx = current_index()
                if idx is None:
                    self._json({"ok": False, "msg": "已全部测完"})
                    return
                if result not in RESULT_LABELS:
                    self._json({"ok": False, "msg": "无效结果"})
                    return
                number = _numbers[idx]
                _results[number] = {"result": result, "note": note,
                                    "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
                # 指针推进到下一个未测号码
                nxt_cursor = None
                for i, n in enumerate(_numbers):
                    if n not in _results:
                        nxt_cursor = i
                        break
                globals()["_cursor"] = nxt_cursor
                save_csv()
                nxt = current_index()
                next_number = _numbers[nxt] if nxt is not None else None
            self._json({"ok": True, "number": number, "next_number": next_number})
        elif self.path == "/api/jump":
            with _lock:
                target = str(body.get("number") or "").strip()
                if target in _numbers:
                    globals()["_cursor"] = _numbers.index(target)
                    self._json({"ok": True, "number": target})
                    return
                try:
                    i = int(body.get("index"))
                except (TypeError, ValueError):
                    self._json({"ok": False, "msg": "无效的跳转目标"})
                    return
                if 0 <= i < len(_numbers):
                    globals()["_cursor"] = i
                    self._json({"ok": True, "number": _numbers[i]})
                else:
                    self._json({"ok": False, "msg": "序号超出范围"})
        elif self.path == "/api/unmark":
            with _lock:
                number = (body.get("number") or "").strip()
                _results.pop(number, None)
                save_csv()
            self._json({"ok": True})
        elif self.path == "/api/voice":
            action = body.get("action")
            if action == "list":
                d = (body.get("dir") or VOICE.get("dir") or "/sdcard/autodail").strip()
                files, err = list_voice_files(d)
                if err:
                    self._json({"ok": False, "msg": err})
                    return
                VOICE["dir"] = d
                save_settings()
                self._json({"ok": True, "files": files, "dir": d})
            elif action == "play":
                name = (body.get("name") or "").strip()
                if not name:
                    self._json({"ok": False, "msg": "请先加载列表并选择语音文件"})
                    return
                VOICE["name"] = name
                save_settings()
                set_media_volume()
                uri = resolve_voice_uri(VOICE.get("dir", "/sdcard/autodail"), name)
                out = adb("shell", "am", "start", "-a", "android.intent.action.VIEW",
                          "-d", uri, "-t", "audio/*", timeout=10)
                low = (out or "").lower()
                ok = out is not None and "error" not in low and "exception" not in low
                self._json({"ok": ok, "msg": "已在手机上播放（媒体音量已调到最大）" if ok
                            else "播放失败：手机未弹出播放器或文件不可读"})
            elif action == "stop":
                ok, msg = stop_voice()
                self._json({"ok": ok, "msg": msg})
            elif action == "speaker":
                ok, msg = toggle_speaker()
                set_status(("手动开免提成功：" if ok else "手动开免提失败：") + msg)
                self._json({"ok": ok, "msg": msg})
            elif action == "shot":
                png = adb_binary("exec-out", "screencap", "-p", timeout=25)
                ok = png.startswith(b"\x89PNG")
                if ok:
                    w = int.from_bytes(png[16:20], "big")
                    h = int.from_bytes(png[20:24], "big")
                    self._json({"ok": True, "image": base64.b64encode(png).decode(), "w": w, "h": h})
                else:
                    self._json({"ok": False, "msg": "截图失败，请检查手机连接"})
            elif action == "calibrate":
                try:
                    sx = float(body["x"]); sy = float(body["y"])
                except (KeyError, TypeError, ValueError):
                    self._json({"ok": False, "msg": "坐标参数错误"})
                    return
                if not (0 <= sx <= 1 and 0 <= sy <= 1):
                    self._json({"ok": False, "msg": "坐标需在 0-1 之间"})
                    return
                w, h = get_screen_size()
                VOICE.update(sx=sx, sy=sy, ssw=w, ssh=h)
                save_settings()
                set_status("免提键已标定: (%.1f%%, %.1f%%)" % (sx * 100, sy * 100))
                self._json({"ok": True, "msg": "已保存免提键位置", "voice": dict(VOICE)})
            elif action == "save":
                if "name" in body:
                    VOICE["name"] = (body.get("name") or "").strip()
                if "dir" in body and (body.get("dir") or "").strip():
                    VOICE["dir"] = body["dir"].strip()
                if "auto_speaker" in body:
                    VOICE["auto_speaker"] = bool(body["auto_speaker"])
                save_settings()
                self._json({"ok": True, "voice": dict(VOICE)})
            elif action == "playcall":
                if VOICE.get("speaker", True):
                    ensure_call_ui()
                ok, msg = play_during_call()
                self._json({"ok": ok, "msg": msg})
            else:
                self._json({"ok": False, "msg": "未知操作"})
        else:
            self._json({"error": "not found"}, 404)


def main():
    args = [a for a in sys.argv[1:]]
    path, port = None, 8765
    if args:
        if args[0].isdigit():
            port = int(args[0])
            path = args[1] if len(args) > 1 else None
        else:
            path = args[0]
            port = int(args[1]) if len(args) > 1 else 8765
    if path:
        if not os.path.exists(path):
            print("号码文件不存在:", path)
            sys.exit(1)
        with _lock:
            _numbers.extend(load_numbers(path))

    with _lock:
        load_import_meta()
        if not path and IMPORT_META.get("file"):
            # 无预载文件时，按上次导入元数据恢复号码列表与导出能力
            try:
                do_import(IMPORT_META.get("file"), IMPORT_META.get("sheet"),
                          int(IMPORT_META.get("header_row", 0)),
                          int(IMPORT_META.get("phone_col", 0)),
                          IMPORT_META.get("name_col"), IMPORT_META.get("addr_col"),
                          IMPORT_META.get("out_col"))
            except Exception as e:
                print("[提示] 恢复上次导入失败：%s" % e)
        load_prev_results()
        save_csv()

    if shutil.which(ADB) is None:
        print("[提示] 未找到 adb，请先安装（如 Ubuntu: sudo apt install adb）")
        print("       或设置环境变量 ADB_PATH 指向 adb 可执行文件")
    if openpyxl is None:
        print("[提示] 未安装 openpyxl，无法导入/回写 xlsx：pip3 install --user --break-system-packages openpyxl")

    print(f"共 {len(_numbers)} 个号码，已完成 {len(_results)} 个，待测 {len(_numbers) - len(_results)} 个")
    print(f"电话间隔: {SETTINGS['interval']} 秒（网页设置中可修改）")
    print(f"结果文件: {CSV_PATH}")
    print(f"请在浏览器打开 http://{HOST}:{port} 开始拨测 (Ctrl+C 退出)")

    srv = ThreadingHTTPServer((HOST, port), Handler)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("\n已退出，结果保存在", CSV_PATH)


if __name__ == "__main__":
    main()
