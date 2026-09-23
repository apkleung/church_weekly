# -*- coding: utf-8 -*-
"""Vercel Python serverless function.
   前端 POST JSON -> 行 make_program_tpl.py 生成 .docx -> 傳返落嚟。
   Vercel 直接用呢個檔做 function，唔使執 handler 名。
"""
from http.server import BaseHTTPRequestHandler
import io, json, os, sys, tempfile, traceback

_here = os.path.dirname(os.path.abspath(__file__))
_root = os.path.dirname(_here)
sys.path.insert(0, _root)
import make_program_tpl as gen


class handler(BaseHTTPRequestHandler):
    """Vercel 經典 Python handler：BaseHTTPRequestHandler 子類。"""

    def do_POST(self):
        try:
            length = int(self.headers.get("Content-Length", 0))
            raw = self.rfile.read(length) if length else b"{}"
            data = json.loads(raw.decode("utf-8"))
            docx = self._generate(data)
            self.send_response(200)
            self.send_header("Content-Type",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            self.send_header("Content-Disposition",
                "attachment; filename=\"program-%s.docx\"" % str(data.get("date", "table")))
            self.send_header("Content-Length", str(len(docx)))
            self.end_headers()
            self.wfile.write(docx)
        except Exception as e:
            body = json.dumps({"error": str(e), "trace": traceback.format_exc()}).encode("utf-8")
            self.send_response(500)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

    def _generate(self, data):
        os.chdir(_root)  # 模板 .docx 喺項目根
        fd, out = tempfile.mkstemp(suffix=".docx"); os.close(fd)
        try:
            gen.build(data, out)
            with open(out, "rb") as f:
                return f.read()
        finally:
            try: os.remove(out)
            except OSError: pass

    def log_message(self, *a):  # 唔想 Vercel log 噪音
        pass