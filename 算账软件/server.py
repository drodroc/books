# -*- coding: utf-8 -*-
"""产品月账本地服务。只监听本机。"""

from __future__ import annotations

import json
import mimetypes
import re
import sys
import threading
import webbrowser
from email.parser import BytesParser
from email.policy import default
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse

import ledger

ROOT = Path(__file__).resolve().parent
STATIC = ROOT / "static"
HOST = "127.0.0.1"
MAX_BODY = 80 * 1024 * 1024


def db_call(fn):
    with ledger._lock:
        conn = ledger.connect()
        try:
            return fn(conn)
        finally:
            conn.close()


def parse_multipart(content_type, body):
    raw = (
        b"Content-Type: "
        + content_type.encode("utf-8", "replace")
        + b"\r\nMIME-Version: 1.0\r\n\r\n"
        + body
    )
    msg = BytesParser(policy=default).parsebytes(raw)
    fields = {}
    files = []
    if not msg.is_multipart():
        return fields, files
    for part in msg.iter_parts():
        name = part.get_param("name", header="content-disposition")
        if not name:
            continue
        payload = part.get_payload(decode=True) or b""
        filename = part.get_filename()
        if filename:
            files.append({"name": name, "filename": filename, "data": payload})
        else:
            fields[name] = payload.decode("utf-8", "replace")
    return fields, files


class App(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt, *args):
        sys.stderr.write("%s\n" % (fmt % args))

    def _url(self):
        return urlparse(self.path)

    def _query(self):
        return parse_qs(self._url().query)

    def _read_body(self):
        length = int(self.headers.get("Content-Length") or 0)
        if length > MAX_BODY:
            raise ValueError("文件太大了。")
        return self.rfile.read(length) if length else b""

    def _json_body(self):
        raw = self._read_body()
        if not raw:
            return {}
        return json.loads(raw.decode("utf-8"))

    def _send(self, status, body, content_type, filename=None):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        if filename:
            self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(filename)}")
        self.end_headers()
        self.wfile.write(body)

    def _json(self, data, status=200):
        self._send(status, json.dumps(data, ensure_ascii=False), "application/json; charset=utf-8")

    def _error(self, message, status=400):
        self._json({"error": message}, status)

    def _static(self, rel):
        rel = unquote(rel).lstrip("/")
        path = (STATIC / rel).resolve()
        if not str(path).startswith(str(STATIC.resolve())) or not path.is_file():
            self._error("没有这个文件。", 404)
            return
        mime = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        if path.suffix in (".js",):
            mime = "text/javascript; charset=utf-8"
        elif path.suffix == ".css":
            mime = "text/css; charset=utf-8"
        elif path.suffix == ".html":
            mime = "text/html; charset=utf-8"
        self._send(200, path.read_bytes(), mime)

    def do_GET(self):
        try:
            url = self._url()
            path = url.path
            if path in ("/", "/index.html"):
                self._static("index.html")
                return
            if path.startswith("/static/"):
                self._static(path[len("/static/") :])
                return
            if path == "/api/state":
                query = self._query()
                month = (query.get("month") or [""])[0]
                basis = (query.get("basis") or ["confirmed"])[0]
                self._json(db_call(lambda conn: ledger.build_state(conn, month, basis)))
                return
            if path == "/api/export":
                query = self._query()
                month = (query.get("month") or [""])[0]
                basis = (query.get("basis") or ["confirmed"])[0]
                if not ledger.valid_month(month):
                    self._error("先选择一个月份再导出。")
                    return
                data = db_call(lambda conn: ledger.export_workbook(conn, month, basis))
                self._send(
                    200,
                    data,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    f"月账-{month}.xlsx",
                )
                return
            if path == "/api/template":
                self._send(
                    200,
                    ledger.expense_template(),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "支出模板.xlsx",
                )
                return
            if path == "/api/backup":
                data = db_call(ledger.backup_data)
                raw = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
                self._send(200, raw, "application/json; charset=utf-8", "月账备份.json")
                return
            self._error("没有这个地址。", 404)
        except ValueError as exc:
            self._error(str(exc))
        except Exception as exc:
            self._error(f"出错了：{exc}", 500)

    def do_POST(self):
        try:
            path = self._url().path
            if path == "/api/preview":
                ctype = self.headers.get("Content-Type", "")
                _fields, files = parse_multipart(ctype, self._read_body())
                if not files:
                    self._error("没有收到文件。")
                    return
                upload = files[0]
                token, saved = ledger.save_upload(upload["filename"], upload["data"])
                try:
                    info = ledger.analyze_file(saved)
                except Exception:
                    saved.unlink(missing_ok=True)
                    (ledger.TMP / f"{token}.json").unlink(missing_ok=True)
                    raise
                info["token"] = token
                info["filename"] = upload["filename"]
                self._json(info)
                return
            if path == "/api/commit":
                body = self._json_body()
                result = db_call(
                    lambda conn: ledger.commit_import(
                        conn,
                        body.get("token", ""),
                        body.get("kind") or "",
                        body.get("month") or "",
                        bool(body.get("include_costs", True)),
                        body.get("mapping") or None,
                        (body.get("track") or "").strip(),
                        bool(body.get("allow_overlap")),
                        bool(body.get("assign_undated")),
                        bool(body.get("costs_only")),
                    )
                )
                self._json(result)
                return
            if path == "/api/expenses":
                body = self._json_body()
                expense_id = db_call(
                    lambda conn: ledger.add_expense(
                        conn,
                        body.get("name", ""),
                        body.get("amount"),
                        body.get("month", ""),
                        body.get("category", ""),
                        body.get("note", ""),
                        body.get("spent_on", ""),
                        body.get("title", ""),
                    )
                )
                self._json({"id": expense_id})
                return
            if path == "/api/allocations":
                body = self._json_body()
                db_call(
                    lambda conn: ledger.allocate(
                        conn,
                        int(body["expense_id"]),
                        body.get("title", ""),
                        body.get("amount"),
                        body.get("attr") or "",
                    )
                )
                self._json({"ok": True})
                return
            if path == "/api/flags":
                body = self._json_body()
                db_call(lambda conn: ledger.set_no_cost(conn, body["title"], body["month"], bool(body.get("no_cost"))))
                self._json({"ok": True})
                return
            if path == "/api/meta":
                body = self._json_body()
                db_call(lambda conn: ledger.set_product_meta(conn, body["title"], body.get("track"), body.get("note")))
                self._json({"ok": True})
                return
            if path == "/api/demo":
                month = db_call(ledger.load_demo)
                self._json({"month": month})
                return
            if path == "/api/restore":
                ctype = self.headers.get("Content-Type", "")
                if "multipart/form-data" in ctype:
                    _fields, files = parse_multipart(ctype, self._read_body())
                    if not files:
                        self._error("没有收到备份文件。")
                        return
                    payload = json.loads(files[0]["data"].decode("utf-8"))
                else:
                    payload = self._json_body()
                db_call(lambda conn: ledger.restore_data(conn, payload))
                self._json({"ok": True})
                return
            self._error("没有这个地址。", 404)
        except ValueError as exc:
            self._error(str(exc))
        except KeyError as exc:
            self._error(f"缺少字段 {exc}")
        except json.JSONDecodeError:
            self._error("请求不是合法的 JSON。")
        except Exception as exc:
            self._error(f"出错了：{exc}", 500)

    def do_PATCH(self):
        try:
            path = self._url().path
            body = self._json_body()
            expense = re.fullmatch(r"/api/expenses/(\d+)", path)
            alloc = re.fullmatch(r"/api/allocations/(\d+)", path)
            if expense:
                db_call(lambda conn: ledger.update_expense(conn, int(expense.group(1)), body))
                self._json({"ok": True})
                return
            if alloc:
                db_call(
                    lambda conn: ledger.update_allocation(
                        conn,
                        int(alloc.group(1)),
                        body.get("amount"),
                        body.get("title") if "title" in body else None,
                        body.get("attr") if "attr" in body else None,
                    )
                )
                self._json({"ok": True})
                return
            self._error("没有这个地址。", 404)
        except ValueError as exc:
            self._error(str(exc))
        except Exception as exc:
            self._error(f"出错了：{exc}", 500)

    def do_DELETE(self):
        try:
            path = self._url().path
            expense = re.fullmatch(r"/api/expenses/(\d+)", path)
            alloc = re.fullmatch(r"/api/allocations/(\d+)", path)
            imp = re.fullmatch(r"/api/imports/(\d+)", path)
            if expense:
                db_call(lambda conn: ledger.delete_expense(conn, int(expense.group(1))))
            elif alloc:
                db_call(lambda conn: ledger.delete_allocation(conn, int(alloc.group(1))))
            elif imp:
                db_call(lambda conn: ledger.delete_import(conn, int(imp.group(1))))
            else:
                self._error("没有这个地址。", 404)
                return
            self._json({"ok": True})
        except ValueError as exc:
            self._error(str(exc))
        except Exception as exc:
            self._error(f"出错了：{exc}", 500)


def open_window(url):
    if sys.platform == "win32":
        candidates = [
            Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
            Path(r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
            Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
            Path(r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"),
        ]
        for candidate in candidates:
            if candidate.exists():
                try:
                    import subprocess

                    subprocess.Popen([str(candidate), f"--app={url}"])
                    return
                except OSError:
                    continue
    webbrowser.open(url)


def serve(open_browser=True):
    if hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass
    ledger.sweep_tmp()
    server = None
    port = 17321
    for candidate in range(17321, 17340):
        try:
            server = ThreadingHTTPServer((HOST, candidate), App)
            port = candidate
            break
        except OSError:
            continue
    if server is None:
        raise SystemExit("17321 到 17339 端口都被占用了，请关掉占用的程序再打开。")
    url = f"http://{HOST}:{port}/"
    print("产品月账已经打开：" + url)
    print("关掉这个窗口，软件就退出。数据在「算账软件\\data\\ledger.db」。")
    if open_browser:
        threading.Timer(0.4, lambda: open_window(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已退出。")
    finally:
        server.server_close()


if __name__ == "__main__":
    serve("--no-browser" not in sys.argv)
