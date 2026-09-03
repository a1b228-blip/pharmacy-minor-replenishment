#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
藥劑科小藥撥補單 - 本地專用 HTTP 伺服器 (靜態檔案 + 高畫質向量 PDF 生成 API)
"""

import http.server
import socketserver
import json
import os
import io
import sys
import tempfile

from convert import generate_single_page_pdf, TARGET_COLUMN_MAP

PORT = 8765

class PharmacyHandler(http.server.SimpleHTTPRequestHandler):
    def do_POST(self):
        if self.path == '/api/pdf':
            try:
                content_length = int(self.headers.get('Content-Length', 0))
                post_data = self.rfile.read(content_length)
                req_json = json.loads(post_data.decode('utf-8'))

                raw_rows = req_json.get('rows', [])

                # 對齊主管 14 欄標準格式
                headers = [h for _, h in TARGET_COLUMN_MAP]
                col_keys = [k for k, _ in TARGET_COLUMN_MAP]

                formatted_rows = []
                for row_dict in raw_rows:
                    curr_row = []
                    for k in col_keys:
                        if k is None:
                            val = row_dict.get('單位簽收人/日期', '')
                        else:
                            val = row_dict.get(k, '')
                        curr_row.append(str(val or ''))
                    formatted_rows.append(curr_row)

                # 生成高品質單頁橫式 PDF
                with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
                    tmp_pdf_path = tmp.name

                generate_single_page_pdf(headers, formatted_rows, tmp_pdf_path)

                with open(tmp_pdf_path, 'rb') as f:
                    pdf_bytes = f.read()

                os.remove(tmp_pdf_path)

                self.send_response(200)
                self.send_header('Content-Type', 'application/pdf')
                self.send_header('Content-Length', str(len(pdf_bytes)))
                self.send_header('Content-Disposition', 'attachment; filename="single_page.pdf"')
                self.end_headers()
                self.wfile.write(pdf_bytes)
                return

            except Exception as e:
                print("API /api/pdf Error:", e)
                self.send_response(500)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.end_headers()
                self.wfile.write(json.dumps({'error': str(e)}).encode('utf-8'))
                return

        self.send_response(404)
        self.end_headers()

def run_server():
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    socketserver.TCPServer.allow_reuse_address = True
    with socketserver.TCPServer(('127.0.0.1', PORT), PharmacyHandler) as httpd:
        print(f"✅ 藥劑科小藥撥補單伺服器已啟動：http://127.0.0.1:{PORT}")
        httpd.serve_forever()

if __name__ == '__main__':
    run_server()
