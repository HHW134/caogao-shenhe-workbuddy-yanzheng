# -*- coding: utf-8 -*-
"""
草案审核 WorkBuddy 验证 —— 本地 Web 服务
========================================

零第三方依赖（仅 Python 标准库）的轻量 HTTP 服务，为前端提供：

    GET  /                      前端页面
    GET  /static/<file>         前端静态资源
    GET  /api/modules           审核模块清单
    GET  /api/rules             规则库概览
    GET  /api/rules/<key>       某个规则文件的规则明细
    POST /api/review            上传 docx 并执行审核，返回统一问题表
    GET  /api/runs              历史审核记录
    GET  /api/download/<run>/<name>  下载产物（批注 docx / 问题清单 xlsx）

启动：
    python server.py            # 默认 http://127.0.0.1:8000
    python server.py --port 8800 --open
"""

from __future__ import annotations

import argparse
import json
import mimetypes
import re
import sys
import traceback
import uuid
from datetime import datetime
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse

BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))

from review_engine import MODULES, run_review  # noqa: E402

FRONTEND_DIR = BASE_DIR / "frontend"
RUNS_DIR = BASE_DIR / "runs"
RULES_DIR = BASE_DIR / "rules"
RUNS_DIR.mkdir(exist_ok=True)

MAX_UPLOAD = 60 * 1024 * 1024  # 60MB

RULE_FILES = {
    "terminology": {"file": "terminology_rules.json", "name": "术语和定义规则", "desc": "术语条目、引导语、章节结构等 17 条规则"},
    "a4guifan": {"file": "a4guifan_rules.json", "name": "规范性引用文件规则", "desc": "A4R001~A4R015，引用章节与正文引用一致性"},
    "linkage": {"file": "form_validation_term_ref_rules.json", "name": "术语与引用联动规则", "desc": "FVT001~FVT008，两章节交叉校验"},
    "fulldoc": {"file": "fulldoc_rules.json", "name": "全文档审核规则汇总", "desc": "目次/范围/前言/缩略语/正文等全模块规则明细"},
}


# ---------------------------------------------------------------------------
# 规则读取与归一化
# ---------------------------------------------------------------------------

def _load_json(path: Path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def normalize_rules(key: str) -> list:
    """把四种不同格式的规则文件统一成同一份表结构。"""
    info = RULE_FILES.get(key)
    if not info:
        return []
    path = RULES_DIR / info["file"]
    if not path.exists():
        return []
    data = _load_json(path)
    rows = []

    if key == "fulldoc":
        details = data.get("规则明细", []) if isinstance(data, dict) else []
        for i, r in enumerate(details, 1):
            rows.append({
                "编号": r.get("规则编号") or f"FD{i:03d}",
                "分类": r.get("所属模块", ""),
                "审核要点": r.get("审核要点", ""),
                "问题": r.get("审核问题", ""),
                "修改建议": r.get("如何修改", ""),
                "级别": r.get("严重级别", ""),
                "启用": "启用" if "active" in str(r.get("启用状态", "")).lower() else "未启用",
                "来源": r.get("源码位置", ""),
            })
        return rows

    if key == "terminology":
        for i, r in enumerate(data, 1):
            rows.append({
                "编号": f"T{i:03d}",
                "分类": r.get("审核要点", ""),
                "审核要点": r.get("规则描述", ""),
                "问题": r.get("审核问题", ""),
                "修改建议": r.get("如何修改", ""),
                "级别": "",
                "启用": "启用",
                "来源": "术语章节审核器",
            })
        return rows

    for r in data:  # a4guifan / linkage
        rows.append({
            "编号": r.get("rule_id", ""),
            "分类": r.get("category", ""),
            "审核要点": r.get("rule_name", ""),
            "问题": r.get("comment_text", ""),
            "修改建议": r.get("how_to_fix", ""),
            "级别": r.get("severity", ""),
            "启用": "启用",
            "来源": r.get("source_function", ""),
        })
    return rows


def rules_overview() -> list:
    out = []
    for key, info in RULE_FILES.items():
        rows = normalize_rules(key)
        out.append({
            "key": key,
            "名称": info["name"],
            "说明": info["desc"],
            "文件": f"rules/{info['file']}",
            "规则数": len(rows),
            "启用数": sum(1 for r in rows if r["启用"] == "启用"),
        })
    return out


# ---------------------------------------------------------------------------
# multipart/form-data 解析（标准库，Python 3.13 已移除 cgi 模块）
# ---------------------------------------------------------------------------

def parse_multipart(body: bytes, boundary: bytes):
    fields, files = {}, {}
    delim = b"--" + boundary
    for seg in body.split(delim):
        if not seg or seg in (b"--", b"--\r\n", b"\r\n"):
            continue
        if seg.startswith(b"\r\n"):
            seg = seg[2:]
        if seg.endswith(b"\r\n"):
            seg = seg[:-2]
        if seg.endswith(b"--"):
            seg = seg[:-2].rstrip(b"\r\n")
        head, sep, data = seg.partition(b"\r\n\r\n")
        if not sep:
            continue
        headers = head.decode("utf-8", "replace")
        m_name = re.search(r'name="([^"]*)"', headers)
        if not m_name:
            continue
        name = m_name.group(1)
        m_file = re.search(r'filename="([^"]*)"', headers)
        if m_file and m_file.group(1):
            files[name] = (m_file.group(1), data)
        else:
            fields[name] = data.decode("utf-8", "replace")
    return fields, files


# ---------------------------------------------------------------------------
# 审核执行
# ---------------------------------------------------------------------------

def execute_review(filename: str, raw: bytes, modules: list) -> dict:
    run_id = datetime.now().strftime("%Y%m%d-%H%M%S-") + uuid.uuid4().hex[:6]
    run_dir = RUNS_DIR / run_id
    run_dir.mkdir(parents=True, exist_ok=True)

    safe_name = re.sub(r'[\\/:*?"<>|]', "_", Path(filename).name) or "input.docx"
    if not safe_name.lower().endswith(".docx"):
        safe_name += ".docx"
    src = run_dir / safe_name
    src.write_bytes(raw)

    summary = run_review(src, modules=modules, out_dir=run_dir,
                         write_docx=True, write_xlsx=True)

    downloads = []
    for art in summary.get("产物", []):
        p = Path(art["路径"])
        if p.exists():
            downloads.append({
                "类型": art["类型"],
                "模块": art["模块"],
                "文件名": p.name,
                "大小": p.stat().st_size,
                "url": f"/api/download/{run_id}/{p.name}",
            })
    summary["产物"] = downloads
    summary["run_id"] = run_id

    with open(run_dir / "result.json", "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    return summary


def list_runs(limit: int = 30) -> list:
    out = []
    for d in sorted(RUNS_DIR.iterdir(), reverse=True):
        if not d.is_dir():
            continue
        rj = d / "result.json"
        if not rj.exists():
            continue
        try:
            data = _load_json(rj)
        except Exception:
            continue
        out.append({
            "run_id": d.name,
            "文件名": data.get("文件名", ""),
            "审核时间": data.get("审核时间", ""),
            "问题总数": data.get("问题总数", 0),
            "级别统计": data.get("级别统计", {}),
            "执行模块": data.get("执行模块", []),
        })
        if len(out) >= limit:
            break
    return out


# ---------------------------------------------------------------------------
# HTTP Handler
# ---------------------------------------------------------------------------

class ReviewHandler(SimpleHTTPRequestHandler):
    server_version = "DraftReview/1.0"

    def log_message(self, fmt, *args):  # 精简日志
        sys.stderr.write("[%s] %s\n" % (datetime.now().strftime("%H:%M:%S"), fmt % args))

    # -- 工具 ---------------------------------------------------------------
    def _send(self, code: int, body: bytes, ctype: str = "application/json; charset=utf-8",
              extra: dict | None = None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def _json(self, data, code: int = 200):
        self._send(code, json.dumps(data, ensure_ascii=False).encode("utf-8"))

    def _file(self, path: Path, download: bool = False):
        if not path.exists() or not path.is_file():
            return self._json({"error": "文件不存在", "path": path.name}, 404)
        ctype = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        extra = {}
        if download:
            ascii_name = re.sub(r"[^A-Za-z0-9._-]", "_", path.name) or "download"
            extra["Content-Disposition"] = (
                f"attachment; filename=\"{ascii_name}\"; "
                f"filename*=UTF-8''{quote_name(path.name)}")
        self._send(200, path.read_bytes(), ctype, extra)

    def _route(self) -> str:
        """解析请求路径，兼容 percent-encoded 与原始 UTF-8 字节两种中文路径。"""
        route = unquote(urlparse(self.path).path, encoding="utf-8", errors="replace")
        try:  # http.server 以 latin-1 解码请求行，未编码的中文需要还原
            route = route.encode("latin-1").decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            pass
        return route

    # -- GET ----------------------------------------------------------------
    def do_GET(self):
        route = self._route()
        try:
            if route in ("/", "/index.html"):
                return self._file(FRONTEND_DIR / "index.html")
            if route.startswith("/static/"):
                target = (FRONTEND_DIR / route[len("/static/"):]).resolve()
                if FRONTEND_DIR.resolve() in target.parents:
                    return self._file(target)
                return self._json({"error": "非法路径"}, 403)
            if route == "/api/modules":
                return self._json({
                    "modules": [{"key": k, **v} for k, v in MODULES.items()],
                    "rules": rules_overview(),
                })
            if route == "/api/rules":
                return self._json({"rules": rules_overview()})
            if route.startswith("/api/rules/"):
                return self._json({"rows": normalize_rules(route.split("/")[-1])})
            if route == "/api/runs":
                return self._json({"runs": list_runs()})
            if route.startswith("/api/download/"):
                parts = route.split("/")
                if len(parts) >= 5:
                    run_id, name = parts[3], parts[4]
                    target = (RUNS_DIR / run_id / name).resolve()
                    if RUNS_DIR.resolve() in target.parents:
                        return self._file(target, download=True)
                return self._json({"error": "参数错误"}, 400)
            if route.startswith("/api/result/"):
                run_id = route.split("/")[-1]
                return self._file(RUNS_DIR / run_id / "result.json")
            return self._json({"error": "未知接口"}, 404)
        except Exception:
            traceback.print_exc()
            return self._json({"error": "服务器内部错误", "detail": traceback.format_exc(limit=3)}, 500)

    # -- POST ---------------------------------------------------------------
    def do_POST(self):
        route = self._route()
        if route != "/api/review":
            return self._json({"error": "未知接口"}, 404)

        try:
            length = int(self.headers.get("Content-Length") or 0)
            if length <= 0:
                return self._json({"error": "空请求"}, 400)
            if length > MAX_UPLOAD:
                return self._json({"error": f"文件过大（上限 {MAX_UPLOAD // 1024 // 1024}MB）"}, 413)

            ctype = self.headers.get("Content-Type", "")
            m = re.search(r"boundary=([^;]+)", ctype)
            if not m:
                return self._json({"error": "请求不是 multipart/form-data"}, 400)
            boundary = m.group(1).strip('"').encode("utf-8")

            body = self.rfile.read(length)
            fields, files = parse_multipart(body, boundary)
            if "file" not in files:
                return self._json({"error": "未收到文件"}, 400)

            filename, raw = files["file"]
            if not filename.lower().endswith(".docx"):
                return self._json({"error": "仅支持 .docx 文件"}, 400)

            modules = [m_.strip() for m_ in (fields.get("modules") or "chapters").split(",") if m_.strip()]
            summary = execute_review(filename, raw, modules)
            return self._json(summary)
        except Exception:
            traceback.print_exc()
            return self._json({"error": "审核失败", "detail": traceback.format_exc(limit=5)}, 500)


def quote_name(name: str) -> str:
    from urllib.parse import quote
    return quote(name, safe="")


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="草案审核 WorkBuddy 验证：本地 Web 服务")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", type=int, default=8000)
    parser.add_argument("--open", action="store_true", help="启动后自动打开浏览器")
    args = parser.parse_args(argv)

    httpd = ThreadingHTTPServer((args.host, args.port), ReviewHandler)
    url = f"http://{args.host}:{args.port}"
    print("草案审核 WorkBuddy 验证 已启动")
    print(f"  访问地址：{url}")
    print(f"  产物目录：{RUNS_DIR}")
    print("  Ctrl+C 退出")
    if args.open:
        import webbrowser
        webbrowser.open(url)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n已停止")
    finally:
        httpd.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
