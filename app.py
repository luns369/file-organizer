# -*- coding: utf-8 -*-
"""
AI 文件整理工具 —— 图形界面（本地网页版）

用法：
    python app.py

启动后会自动打开浏览器访问 http://127.0.0.1:<端口>，
页面上可一键执行：分析生成总表、按分类归档、撤销归档。

说明：
  - 服务只在本机运行（127.0.0.1），不会暴露到局域网
  - API key 保存在本机 config.json，不经过任何外部服务器
"""

import contextlib
import json
import os
import socket
import sys
import threading
import traceback
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

# 数据目录：源码运行时用脚本所在目录；打包成 exe 后自动改用 exe 所在目录
if getattr(sys, "frozen", False):
    DATA_DIR = Path(sys.executable).resolve().parent
else:
    DATA_DIR = Path(__file__).resolve().parent
CONFIG_FILE = DATA_DIR / "config.json"
LOG_DIR = DATA_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)
# 打包成 exe 后，gui.html 被打进 exe 内部（_MEIPASS），从那里读取
if getattr(sys, "_MEIPASS", None):
    GUI_FILE = Path(sys._MEIPASS) / "gui.html"
else:
    GUI_FILE = DATA_DIR / "gui.html"

# 可执行的任务名
TASKS = {"analyze", "organize", "undo"}

# 全局运行状态
_lock = threading.Lock()
_state = {
    "running": None,      # 当前正在运行的任务名
    "exit_code": None,    # 最近一次任务的退出码
    "log_path": None,     # 当前任务的日志文件
}


def mask_key(key):
    """API key 掩码显示：sk-****后4位，避免明文回传浏览器。"""
    if not key:
        return ""
    return f"{key[:3]}****{key[-4:]}"


def load_config():
    if CONFIG_FILE.exists():
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    return {}


def save_config(data):
    """合并保存配置；api_key 传空字符串表示保留原值。"""
    cfg = load_config()
    for k in ("source_folder", "api_key", "base_url", "model", "excel_name"):
        if data.get(k) not in (None, ""):
            cfg[k] = str(data[k]).strip()
    # custom_rules（自定义分类要求）允许清空：只要提交了该字段就覆盖
    if "custom_rules" in data:
        cfg["custom_rules"] = str(data["custom_rules"]).strip()
    CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    return cfg


def start_task(name):
    """后台启动一个任务。返回 (是否成功, 提示语)。"""
    if name not in TASKS:
        return False, "未知任务"
    with _lock:
        if _state["running"]:
            return False, f"已有任务「{_state['running']}」在运行，请等它完成。"
        log_path = LOG_DIR / f"{name}.log"
        log_path.write_text("", encoding="utf-8")
        _state["running"] = name
        _state["exit_code"] = None
        threading.Thread(target=_wait_task, args=(name,), daemon=True).start()
        return True, f"「{name}」已开始运行，日志实时刷新中……"


def _run_task(name):
    """在后台线程里运行任务，stdout/stderr 重定向到日志文件。返回退出码。"""
    log_path = LOG_DIR / f"{name}.log"
    with open(log_path, "a", encoding="utf-8", buffering=1) as fp:  # 行缓冲，保证日志实时写入
        with contextlib.redirect_stdout(fp), contextlib.redirect_stderr(fp):
            code = 0
            try:
                if name == "analyze":
                    import main as _m
                    _m.main()
                elif name == "organize":
                    import organize as _o
                    _o.do_organize()
                else:
                    import organize as _o
                    _o.do_undo()
            except SystemExit as e:
                # 脚本内部用 sys.exit 提示错误
                msg = e.code
                if msg and not isinstance(msg, int):
                    print(msg)
                code = msg if isinstance(msg, int) else 1
            except Exception:
                traceback.print_exc()
                code = 1
    return code


def _wait_task(name):
    code = _run_task(name)
    with _lock:
        _state["exit_code"] = code
        _state["running"] = None


def read_log(name, offset):
    """读取日志文件自 offset 起的增量行。返回 (新增文本, 新offset, 任务状态)。"""
    log_path = LOG_DIR / f"{name}.log"
    raw = b""
    if log_path.exists():
        try:
            with open(log_path, "rb") as f:
                f.seek(offset)
                raw = f.read()
        except OSError as e:
            raw = f"<读取日志失败: {e}>".encode("utf-8")
    text = raw.decode("utf-8", errors="replace")
    with _lock:
        running = _state["running"] == name
        done = not running
        code = _state["exit_code"] if not running else None
    return text, offset + len(raw), running, done, code


class Handler(BaseHTTPRequestHandler):
    def _send(self, code, body, ctype="application/json; charset=utf-8"):
        data = body if isinstance(body, bytes) else body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, *args):
        pass  # 关闭默认的访问日志，保持控制台干净

    # ---------- 页面 ----------
    def do_GET(self):
        if self.path == "/" or self.path == "/index.html":
            html = GUI_FILE.read_bytes()
            self._send(200, html, "text/html; charset=utf-8")
        elif self.path == "/api/config":
            cfg = load_config()
            self._send(200, json.dumps({
                "source_folder": cfg.get("source_folder", ""),
                "api_key": mask_key(cfg.get("api_key", "")),
                "base_url": cfg.get("base_url", "https://api.deepseek.com"),
                "model": cfg.get("model", "deepseek-chat"),
                "excel_name": cfg.get("excel_name", "文件整理总表.xlsx"),
                "custom_rules": cfg.get("custom_rules", ""),
                "running": _state["running"],
            }, ensure_ascii=False))
        elif self.path.startswith("/api/log/"):
            name = self.path.split("/")[-1].split("?")[0]  # 去掉查询字符串
            offset = int(self._query().get("offset", "0"))
            text, new_off, running, done, code = read_log(name, offset)
            self._send(200, json.dumps({
                "text": text, "offset": new_off,
                "running": running, "done": done, "exit_code": code,
            }, ensure_ascii=False))
        elif self.path == "/api/status":
            self._send(200, json.dumps({"running": _state["running"]}))
        elif self.path == "/api/open_excel":
            cfg = load_config()
            folder = Path(cfg.get("source_folder", ""))
            excel = folder / cfg.get("excel_name", "文件整理总表.xlsx")
            if excel.exists():
                os.startfile(str(excel))  # noqa: 用默认程序打开
                self._send(200, json.dumps({"ok": True}))
            else:
                self._send(200, json.dumps({"ok": False, "msg": f"总表不存在：{excel}"}, ensure_ascii=False))
        elif self.path == "/api/open_folder":
            cfg = load_config()
            folder = Path(cfg.get("source_folder", ""))
            if folder.exists():
                os.startfile(str(folder))  # noqa
                self._send(200, json.dumps({"ok": True}))
            else:
                self._send(200, json.dumps({"ok": False, "msg": f"文件夹不存在：{folder}"}, ensure_ascii=False))
        else:
            self._send(404, json.dumps({"ok": False, "msg": "404"}))

    def do_POST(self):
        if self.path == "/api/config":
            length = int(self.headers.get("Content-Length", 0))
            data = json.loads(self.rfile.read(length) or b"{}")
            cfg = save_config(data)
            self._send(200, json.dumps({
                "ok": True,
                "source_folder": cfg.get("source_folder", ""),
                "api_key": mask_key(cfg.get("api_key", "")),
            }, ensure_ascii=False))
        elif self.path.startswith("/api/run/"):
            name = self.path.split("/")[-1]
            ok, msg = start_task(name)
            self._send(200, json.dumps({"ok": ok, "msg": msg}, ensure_ascii=False))
        elif self.path == "/api/shutdown":
            self._send(200, json.dumps({"ok": True, "msg": "服务已关闭，可以关闭这个窗口。"}, ensure_ascii=False))
            threading.Timer(0.5, self.server.shutdown).start()
        else:
            self._send(404, json.dumps({"ok": False, "msg": "404"}))

    def _query(self):
        from urllib.parse import parse_qs, urlsplit
        return {k: v[0] for k, v in parse_qs(urlsplit(self.path).query).items()}


def find_port(start=8501, end=8601):
    for port in range(start, end):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", port)) != 0:
                return port
    raise OSError("端口 8501-8600 都被占用，无法启动。")



def main():
    # 打包成 exe 且无控制台窗口运行时，stdout/stderr 为 None，print 会报错，这里兜底
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")
    if not GUI_FILE.exists():
        print(f"缺少界面文件：{GUI_FILE}")
        sys.exit(1)
    port = find_port()
    server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = f"http://127.0.0.1:{port}"
    print(f"AI 文件整理助手已启动：{url}")
    print("浏览器未自动打开的话，请手动复制上面地址访问。")
    print("关闭服务：在页面点「关闭服务」，或直接关闭本窗口。\n")
    webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
