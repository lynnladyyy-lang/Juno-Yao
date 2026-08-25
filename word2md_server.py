# -*- coding: utf-8 -*-
"""
word2md_server.py —— Word → Markdown 转换的本地 Web 服务。

双击 word2md.bat 会启动本服务并自动打开浏览器，在网页中：
  右侧选择 Word 文件 → 选择输出目录 → 点击「开始转换」→ 左侧查看结果。

核心转换逻辑复用 word2md.py（Word COM → 过滤 HTML → markdown）。
"""

import os
import re
import sys
import shutil
import socket
import tempfile
import threading
import webbrowser
import ctypes
from ctypes import wintypes
from queue import Queue

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from flask import Flask, request, jsonify, send_from_directory

from word2md import convert_files

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(os.path.expanduser("~"), "Downloads", "word2md_output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

ALLOWED_EXTS = (".doc", ".docx", ".rtf")


class _BROWSEINFO(ctypes.Structure):
    _fields_ = [
        ("hwndOwner", wintypes.HWND),
        ("pidlRoot", ctypes.c_void_p),
        ("pszDisplayName", wintypes.LPWSTR),
        ("lpszTitle", wintypes.LPCWSTR),
        ("ulFlags", wintypes.UINT),
        ("lpfn", ctypes.c_void_p),
        ("lParam", wintypes.LPARAM),
        ("iImage", ctypes.c_int),
    ]


def choose_directory(title="选择输出目录"):
    """弹出 Windows 原生「选择文件夹」对话框，返回所选路径（取消返回 None）。

    注意：SHBrowseForFolderW 是 Shell/COM 调用，必须在 STA 线程中运行，
    所以这里新开一个 STA daemon 线程执行对话框，主线程 join 等结果。
    """
    result_q = Queue()

    def _sta_run():
        shell32 = ctypes.windll.shell32
        ole32 = ctypes.windll.ole32
        user32 = ctypes.windll.user32

        # 64 位关键修复：显式声明返回/参数类型，避免指针被截断为 32 位
        shell32.SHBrowseForFolderW.restype = ctypes.c_void_p
        shell32.SHBrowseForFolderW.argtypes = [ctypes.POINTER(_BROWSEINFO)]
        shell32.SHGetPathFromIDListW.restype = wintypes.BOOL
        shell32.SHGetPathFromIDListW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p]
        ole32.CoTaskMemFree.argtypes = [ctypes.c_void_p]
        user32.GetForegroundWindow.restype = wintypes.HWND

        ole32.CoInitializeEx(None, 0x2)  # COINIT_APARTMENTTHREADED = STA
        try:
            bi = _BROWSEINFO()
            bi.hwndOwner = user32.GetForegroundWindow()  # 对话框置顶/模态到浏览器
            bi.lpszTitle = title
            # BIF_RETURNONLYFSDIRS(0x1) | BIF_NEWDIALOGSTYLE(0x40) | BIF_EDITBOX(0x10)
            bi.ulFlags = 0x00000001 | 0x00000040 | 0x00000010
            pidl = shell32.SHBrowseForFolderW(ctypes.byref(bi))
            if not pidl:
                result_q.put(None)
                return
            buf = ctypes.create_unicode_buffer(1024)
            ok = shell32.SHGetPathFromIDListW(pidl, buf)
            ole32.CoTaskMemFree(pidl)
            result_q.put(buf.value if ok else None)
        except Exception as e:
            result_q.put(e)
        finally:
            ole32.CoUninitialize()

    t = threading.Thread(target=_sta_run, daemon=True)
    t.start()
    t.join()
    val = result_q.get()
    if isinstance(val, Exception):
        raise val
    return val


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 512 * 1024 * 1024  # 单次请求上限 512MB


@app.route("/")
def index():
    return send_from_directory(BASE_DIR, "index.html")


@app.route("/config", methods=["GET"])
def config():
    return jsonify({"default_dir": OUTPUT_DIR})


@app.route("/convert", methods=["POST"])
def convert():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "未收到任何文件"}), 400

    out_dir = (request.form.get("out_dir") or "").strip() or OUTPUT_DIR
    out_dir = os.path.abspath(out_dir)
    os.makedirs(out_dir, exist_ok=True)

    tmpdir = tempfile.mkdtemp(prefix="word2md_upload_")
    results = []
    try:
        src_paths = []
        for f in files:
            name = os.path.basename(f.filename or "untitled")
            if not name.lower().endswith(ALLOWED_EXTS):
                results.append({
                    "name": name,
                    "ok": False,
                    "message": "不支持的文件类型（仅支持 .doc / .docx / .rtf）",
                })
                continue
            dest = os.path.join(tmpdir, name)
            f.save(dest)
            src_paths.append(dest)

        if src_paths:
            conv = convert_files(src_paths, out_dir=out_dir)
            for (ok, msg, mdpath), src in zip(conv, src_paths):
                results.append({
                    "name": os.path.basename(src),
                    "ok": bool(ok),
                    "message": msg,
                    "md": mdpath if ok else None,
                })
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    return jsonify({"results": results, "output_dir": out_dir})


@app.route("/choose_dir", methods=["POST"])
def choose_dir():
    try:
        path = choose_directory()
        return jsonify({"dir": path})
    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "traceback": traceback.format_exc()}), 500


@app.route("/set_dir", methods=["POST"])
def set_dir():
    data = request.get_json(silent=True) or {}
    raw = (data.get("dir") or "").strip()
    if not raw:
        return jsonify({"error": "路径为空"}), 400
    if not re.match(r"^[A-Za-z]:[\\/]|^\\\\", raw):
        return jsonify({"error": "请输入绝对路径，如 D:\\输出 或 C:\\Users\\Lynn\\Desktop\\xx"}), 400
    try:
        path = os.path.abspath(raw)
        os.makedirs(path, exist_ok=True)
        return jsonify({"dir": path})
    except Exception as e:
        return jsonify({"error": f"无法创建目录：{e}"}), 500


@app.route("/open_folder", methods=["POST"])
def open_folder():
    data = request.get_json(silent=True) or {}
    target = (data.get("dir") or "").strip() or OUTPUT_DIR
    try:
        os.startfile(target)
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/shutdown", methods=["POST"])
def shutdown():
    def _exit():
        os._exit(0)
    threading.Timer(0.3, _exit).start()
    return jsonify({"ok": True})


def _free_port():
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port


if __name__ == "__main__":
    port = _free_port()
    url = f"http://127.0.0.1:{port}/"
    threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    print(f"word2md 服务已启动：{url}  默认输出目录：{OUTPUT_DIR}")
    app.run(host="127.0.0.1", port=port, debug=False, threaded=True)
