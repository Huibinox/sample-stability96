"""Serve the shared-folder stability dashboard on the current Windows PC."""
import hashlib
import json
import os
import tempfile
import time
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from stability96 import STAMP, SharedStore, atomic_json, category, deadline, replay

ROOT = Path(__file__).resolve().parent
CONFIG = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "SampleStability96" / "config.json"
MAX_UPLOAD = 30 * 1024 * 1024


class Dashboard:
    def __init__(self, config=CONFIG):
        self.config = config
        self.seen = {}

    def settings(self):
        try:
            return json.loads(self.config.read_text(encoding="utf-8"))
        except FileNotFoundError:
            return {}

    def configure(self, data):
        actor = data["actor"].strip()
        shared = Path(data["shared"].strip())
        if not actor or not shared.is_absolute() or not shared.is_dir():
            raise ValueError("请填写姓名或工号，并输入可访问的共享文件夹完整路径。")
        self.config.parent.mkdir(parents=True, exist_ok=True)
        atomic_json(self.config, dict(actor=actor, shared=str(shared)))

    def connected(self):
        settings = self.settings()
        if not settings.get("shared") or not settings.get("actor"):
            raise ValueError("请先设置姓名和共享文件夹。")
        return SharedStore(settings["shared"]), settings["actor"]

    def samples(self):
        store, actor = self.connected()
        if not store.root.is_dir():
            raise OSError("共享文件夹不可访问。")
        notes = []
        for directory in (store.root, store.root / "待导入"):
            if not directory.is_dir():
                continue
            for path in sorted(directory.iterdir()):
                if path.name.startswith("~$") or path.suffix.lower() not in (".xlsx", ".xlsm") or not path.is_file():
                    continue
                try:
                    stat = path.stat()
                    signature = (stat.st_size, stat.st_mtime_ns)
                    if time.time() - stat.st_mtime < 3 or self.seen.get(path) == signature:
                        continue
                    store.import_file(path, actor)
                    self.seen[path] = signature
                except (OSError, ValueError) as exc:
                    notes.append(f"{path.name}：{exc}")
        samples, reports = replay(store.read_events())
        now = datetime.now()
        rows = []
        for sample in samples.values():
            due = deadline(sample)
            rows.append(dict(sample, state=category(sample, now),
                             due=due.strftime(STAMP) if due else "",
                             seconds=int((due - now).total_seconds()) if due else None))
        return dict(rows=rows, reports=notes + reports, synced=now.strftime(STAMP))

    def upload(self, name, body):
        store, actor = self.connected()
        name = Path(name).name
        source = Path(name)
        if source.suffix.lower() not in (".xlsx", ".xlsm"):
            raise ValueError("请选择 .xlsx 或 .xlsm 文件。")
        if not store.root.is_dir():
            raise OSError("共享文件夹不可访问。")
        directory = store.root / "待导入"
        directory.mkdir(exist_ok=True)
        digest = hashlib.sha256(body).hexdigest()[:12]
        target = directory / f"{source.stem}-{digest}{source.suffix.lower()}"
        fd, temp = tempfile.mkstemp(dir=directory, suffix=".tmp")
        try:
            with os.fdopen(fd, "wb") as handle:
                handle.write(body)
            if not target.exists():
                os.replace(temp, target)
        finally:
            Path(temp).unlink(missing_ok=True)
        return store.import_file(target, actor)

    def action(self, data):
        store, actor = self.connected()
        kind = data["kind"]
        if kind not in ("complete", "undo", "edit"):
            raise ValueError("不支持的操作。")
        value = None
        if kind == "edit":
            project, stamp = data["project"].strip(), data["collected"].strip()
            parsed = datetime.fromisoformat(stamp)
            if not project or ":" not in stamp or parsed.tzinfo is not None:
                raise ValueError("请填写项目号及完整的当地采样时间。")
            value = dict(project=project, collected=parsed.strftime(STAMP))
        store.action(kind, data["bas"], actor, data["expected"], value)


class Handler(BaseHTTPRequestHandler):
    def response(self, code, data, content_type="application/json; charset=utf-8"):
        body = data if isinstance(data, bytes) else json.dumps(data, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def dispatch(self, write=False):
        host = f"127.0.0.1:{self.server.server_port}"
        if self.headers.get("Host") != host or (write and self.headers.get("Origin") != f"http://{host}"):
            return self.response(403, {"error": "请求来源无效。"})
        app = self.server.dashboard
        try:
            if not write:
                if self.path == "/":
                    return self.response(200, self.server.page, "text/html; charset=utf-8")
                if self.path == "/api/settings":
                    return self.response(200, app.settings())
                if self.path == "/api/samples":
                    return self.response(200, app.samples())
            else:
                size = int(self.headers.get("Content-Length", 0))
                if not 0 < size <= MAX_UPLOAD:
                    raise ValueError("请求为空或文件超过 30 MB。")
                body = self.rfile.read(size)
                if self.path == "/api/upload":
                    from urllib.parse import unquote
                    result = app.upload(unquote(self.headers.get("X-Filename", "")), body)
                    return self.response(200, {"imported": result[0] if result else 0})
                data = json.loads(body)
                if self.path == "/api/settings":
                    app.configure(data)
                    return self.response(200, {"ok": True})
                if self.path == "/api/action":
                    app.action(data)
                    return self.response(200, {"ok": True})
        except (ValueError, KeyError, TypeError, AttributeError) as exc:
            return self.response(400, {"error": str(exc)})
        except OSError as exc:
            return self.response(503, {"error": str(exc)})
        except Exception as exc:
            return self.response(500, {"error": str(exc)})
        return self.response(404, {"error": "页面不存在。"})

    def do_GET(self):
        self.dispatch()

    def do_POST(self):
        self.dispatch(write=True)


def make_server(config=CONFIG):
    server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    server.dashboard = Dashboard(config)
    server.page = (ROOT / "web.html").read_bytes()
    return server


if __name__ == "__main__":
    with make_server() as server:
        url = f"http://127.0.0.1:{server.server_port}/"
        print(f"Sample Stability 96h: {url}", flush=True)
        print("Keep this window open while using the dashboard.", flush=True)
        webbrowser.open(url)
        server.serve_forever()
