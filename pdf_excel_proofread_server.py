#!/usr/bin/env python3
from __future__ import annotations

import html
import io
import json
import os
import re
import uuid
import zlib
from dataclasses import dataclass
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import urlparse
from xml.etree import ElementTree as ET
from zipfile import ZipFile


HOST = os.environ.get("HOST", "0.0.0.0")
PORT = int(os.environ.get("PORT", "8765"))
BASE_DIR = Path(__file__).resolve().parent
INDEX_FILE = BASE_DIR / "pdf_excel_proofread_web.html"

PDF_STORE: dict[str, bytes] = {}

XML_NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_NS = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}


def _json_response(handler: BaseHTTPRequestHandler, status: int, payload: dict[str, Any]) -> None:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.send_header("Cache-Control", "no-store")
    handler.end_headers()
    handler.wfile.write(body)


def _read_text(zip_file: ZipFile, path: str) -> str:
    return zip_file.read(path).decode("utf-8")


def _column_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha()).upper()
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return max(idx - 1, 0)


def _normalize_space(text: str) -> str:
    return re.sub(r"\s+", "", text).strip().lower()


def _dedupe_headers(headers: list[str]) -> list[str]:
    used: dict[str, int] = {}
    result: list[str] = []
    for i, header in enumerate(headers):
        base = header.strip() or f"列{i + 1}"
        count = used.get(base, 0)
        if count:
            name = f"{base}({count + 1})"
        else:
            name = base
        used[base] = count + 1
        result.append(name)
    return result


def _shared_strings(zip_file: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zip_file.namelist():
        return []
    root = ET.fromstring(_read_text(zip_file, "xl/sharedStrings.xml"))
    values: list[str] = []
    for si in root.findall("x:si", XML_NS):
        text = "".join((t.text or "") for t in si.findall(".//x:t", XML_NS))
        values.append(text)
    return values


def _sheet_path(zip_file: ZipFile, preferred_name: str = "总览表_录入表") -> str:
    workbook = ET.fromstring(_read_text(zip_file, "xl/workbook.xml"))
    sheets = workbook.find("x:sheets", XML_NS)
    if sheets is None:
        raise ValueError("Excel 缺少 sheets 节点。")

    target_rel_id = None
    fallback_rel_id = None
    for sheet in sheets.findall("x:sheet", XML_NS):
        rel_id = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        if rel_id and fallback_rel_id is None:
            fallback_rel_id = rel_id
        if sheet.attrib.get("name") == preferred_name and rel_id:
            target_rel_id = rel_id
            break
    rel_id = target_rel_id or fallback_rel_id
    if not rel_id:
        raise ValueError("Excel 未找到可用工作表。")

    rel_root = ET.fromstring(_read_text(zip_file, "xl/_rels/workbook.xml.rels"))
    for rel in rel_root.findall("r:Relationship", REL_NS):
        if rel.attrib.get("Id") == rel_id:
            target = rel.attrib.get("Target", "")
            target = target.lstrip("/")
            if not target.startswith("xl/"):
                target = f"xl/{target}"
            return target
    raise ValueError("Excel 工作表关系解析失败。")


def parse_excel_keep_headers(excel_bytes: bytes) -> tuple[list[str], list[dict[str, Any]]]:
    with ZipFile(io.BytesIO(excel_bytes)) as zf:
        shared = _shared_strings(zf)
        sheet_xml_path = _sheet_path(zf)
        sheet_root = ET.fromstring(_read_text(zf, sheet_xml_path))

    rows = sheet_root.findall(".//x:sheetData/x:row", XML_NS)
    if not rows:
        return [], []

    matrix: list[dict[int, str]] = []
    max_col = 0
    for row in rows:
        row_data: dict[int, str] = {}
        for c in row.findall("x:c", XML_NS):
            ref = c.attrib.get("r", "")
            col_idx = _column_index(ref)
            ctype = c.attrib.get("t", "")
            value = ""
            if ctype == "s":
                v = c.find("x:v", XML_NS)
                if v is not None and (v.text or "").isdigit():
                    sidx = int(v.text or "0")
                    if 0 <= sidx < len(shared):
                        value = shared[sidx]
            elif ctype == "inlineStr":
                value = "".join((t.text or "") for t in c.findall(".//x:t", XML_NS))
            else:
                v = c.find("x:v", XML_NS)
                value = (v.text or "") if v is not None else ""
            row_data[col_idx] = value
            max_col = max(max_col, col_idx)
        matrix.append(row_data)

    header_row = matrix[0] if matrix else {}
    raw_headers = [header_row.get(i, "").strip() for i in range(max_col + 1)]
    headers = _dedupe_headers(raw_headers)

    records: list[dict[str, Any]] = []
    for ridx, row_data in enumerate(matrix[1:], start=2):
        cells: dict[str, str] = {}
        has_value = False
        for cidx in range(max_col + 1):
            header = headers[cidx]
            value = row_data.get(cidx, "")
            if value != "":
                has_value = True
            cells[header] = value
        if has_value:
            records.append({"row_number": ridx, "cells": cells})
    return headers, records


def _unescape_pdf_literal(raw: str) -> str:
    out = bytearray()
    i = 0
    while i < len(raw):
        ch = raw[i]
        if ch != "\\":
            out.append(ord(ch))
            i += 1
            continue
        i += 1
        if i >= len(raw):
            break
        esc = raw[i]
        mapping = {"n": 10, "r": 13, "t": 9, "b": 8, "f": 12, "\\": 92, "(": 40, ")": 41}
        if esc in mapping:
            out.append(mapping[esc])
            i += 1
            continue
        if esc.isdigit():
            oct_part = esc
            i += 1
            for _ in range(2):
                if i < len(raw) and raw[i].isdigit():
                    oct_part += raw[i]
                    i += 1
                else:
                    break
            try:
                out.append(int(oct_part, 8) & 0xFF)
            except ValueError:
                pass
            continue
        out.append(ord(esc))
        i += 1
    for enc in ("utf-8", "utf-16-be", "latin1"):
        try:
            return out.decode(enc)
        except Exception:
            pass
    return out.decode("latin1", errors="ignore")


def extract_pdf_text_rough(pdf_bytes: bytes, *, max_streams: int = 900, max_chars: int = 2_000_000) -> str:
    text_parts: list[str] = []
    total_chars = 0
    stream_re = re.compile(rb"stream\r?\n(.*?)\r?\nendstream", re.S)
    stream_count = 0
    for match in stream_re.finditer(pdf_bytes):
        stream_count += 1
        if stream_count > max_streams:
            break
        raw_stream = match.group(1)
        candidates: list[bytes] = []
        candidates.append(raw_stream)
        try:
            candidates.append(zlib.decompress(raw_stream))
        except Exception:
            pass
        for content in candidates:
            if b"Tj" not in content and b"TJ" not in content and b"BT" not in content:
                continue
            s = content.decode("latin1", errors="ignore")
            for lit in re.findall(r"\((.*?)(?<!\\)\)\s*Tj", s, flags=re.S):
                val = _unescape_pdf_literal(lit)
                text_parts.append(val)
                total_chars += len(val)
                if total_chars >= max_chars:
                    break
            for arr in re.findall(r"\[(.*?)\]\s*TJ", s, flags=re.S):
                for lit in re.findall(r"\((.*?)(?<!\\)\)", arr, flags=re.S):
                    val = _unescape_pdf_literal(lit)
                    text_parts.append(val)
                    total_chars += len(val)
                    if total_chars >= max_chars:
                        break
                if total_chars >= max_chars:
                    break
            for hx in re.findall(r"<([0-9A-Fa-f]+)>\s*Tj", s):
                try:
                    b = bytes.fromhex(hx)
                    for enc in ("utf-16-be", "utf-8", "latin1"):
                        try:
                            val = b.decode(enc)
                            text_parts.append(val)
                            total_chars += len(val)
                            if total_chars >= max_chars:
                                break
                            break
                        except Exception:
                            continue
                except Exception:
                    continue
                if total_chars >= max_chars:
                    break
            if total_chars >= max_chars:
                break
        if total_chars >= max_chars:
            break
    combined = "\n".join(part for part in text_parts if part.strip())
    if combined.strip():
        return combined
    return pdf_bytes.decode("latin1", errors="ignore")


FOCUS_HEADERS = ["大条目", "中条目", "小条目", "条目释义", "例句原文块", "条目解释", "PDF页码"]


def _field_candidates(value: str) -> list[str]:
    value = (value or "").strip()
    if not value:
        return []
    lines = [ln.strip() for ln in value.splitlines() if ln.strip()]
    if not lines:
        return []
    if len(lines) == 1:
        return [lines[0][:80]]
    return [lines[0][:80], lines[1][:80]]


def compare_rows_with_pdf(headers: list[str], rows: list[dict[str, Any]], pdf_text: str) -> list[dict[str, Any]]:
    pdf_norm = _normalize_space(pdf_text)
    enriched: list[dict[str, Any]] = []
    for row in rows:
        cells: dict[str, str] = row["cells"]
        target_headers = [h for h in FOCUS_HEADERS if h in headers]
        if not target_headers:
            target_headers = [h for h in headers if (cells.get(h, "").strip())][:6]
        total = 0
        found: list[str] = []
        missing: list[str] = []
        for header in target_headers:
            value = cells.get(header, "")
            snippets = _field_candidates(value)
            if not snippets:
                continue
            total += 1
            ok = any(_normalize_space(s) and _normalize_space(s) in pdf_norm for s in snippets)
            if ok:
                found.append(header)
            else:
                missing.append(header)
        score = (len(found) / total) if total else 0.0
        if score >= 0.6:
            level = "high"
        elif score >= 0.3:
            level = "medium"
        else:
            level = "low"
        enriched.append(
            {
                "row_number": row["row_number"],
                "cells": cells,
                "compare": {
                    "score": round(score, 3),
                    "level": level,
                    "found_fields": found,
                    "missing_fields": missing,
                },
            }
        )
    return enriched


@dataclass
class UploadPayload:
    excel_bytes: bytes
    excel_name: str
    pdf_bytes: bytes | None
    pdf_name: str | None


def parse_upload(handler: BaseHTTPRequestHandler) -> UploadPayload:
    content_type = handler.headers.get("Content-Type", "")
    if "multipart/form-data" not in content_type:
        raise ValueError("请求不是 multipart/form-data。")
    match = re.search(r'boundary="?([^";]+)"?', content_type)
    if not match:
        raise ValueError("未找到 multipart boundary。")
    boundary = match.group(1).encode("utf-8")

    length = int(handler.headers.get("Content-Length", "0"))
    raw = handler.rfile.read(length)
    if not raw:
        raise ValueError("上传内容为空。")

    parts = raw.split(b"--" + boundary)
    files: dict[str, tuple[str, bytes]] = {}
    for part in parts:
        part = part.strip()
        if not part or part == b"--":
            continue
        if b"\r\n\r\n" not in part:
            continue
        header_blob, body = part.split(b"\r\n\r\n", 1)
        body = body.rstrip(b"\r\n")
        headers = header_blob.decode("utf-8", errors="ignore")
        disp = re.search(
            r'Content-Disposition:\s*form-data;\s*name="([^"]+)"(?:;\s*filename="([^"]*)")?',
            headers,
            flags=re.I,
        )
        if not disp:
            continue
        field_name = disp.group(1)
        filename = disp.group(2) or ""
        files[field_name] = (filename, body)

    if "excel" not in files:
        raise ValueError("请至少上传 Excel。")
    excel_name, excel_bytes = files["excel"]
    pdf_name = None
    pdf_bytes = None
    if "pdf" in files:
        pdf_name, pdf_bytes = files["pdf"]
    if not excel_bytes:
        raise ValueError("上传 Excel 内容为空。")
    return UploadPayload(
        excel_bytes=excel_bytes,
        excel_name=(excel_name or "excel.xlsx"),
        pdf_bytes=pdf_bytes if pdf_bytes else None,
        pdf_name=(pdf_name or "document.pdf") if pdf_bytes else None,
    )


class Handler(BaseHTTPRequestHandler):
    server_version = "PDFExcelProofread/1.0"

    def log_message(self, fmt: str, *args: Any) -> None:
        return

    def do_GET(self) -> None:
        path = urlparse(self.path).path
        if path == "/":
            if not INDEX_FILE.exists():
                self.send_error(404, "Missing index file")
                return
            body = INDEX_FILE.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        if path.startswith("/api/pdf/"):
            token = path.split("/api/pdf/", 1)[1]
            data = PDF_STORE.get(token)
            if data is None:
                self.send_error(404, "PDF not found")
                return
            self.send_response(200)
            self.send_header("Content-Type", "application/pdf")
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Cache-Control", "no-store")
            self.end_headers()
            self.wfile.write(data)
            return
        self.send_error(404, "Not found")

    def do_POST(self) -> None:
        path = urlparse(self.path).path
        if path != "/api/compare":
            self.send_error(404, "Not found")
            return
        try:
            payload = parse_upload(self)
            headers, rows = parse_excel_keep_headers(payload.excel_bytes)
            if not headers or not rows:
                raise ValueError("Excel 未解析出数据，请确认第一行为表头。")
            compared = [
                {
                    "row_number": row["row_number"],
                    "cells": row["cells"],
                    "compare": {
                        "score": None,
                        "level": "na",
                        "found_fields": [],
                        "missing_fields": [],
                    },
                }
                for row in rows
            ]
            token = None
            if payload.pdf_bytes:
                token = uuid.uuid4().hex
                PDF_STORE[token] = payload.pdf_bytes
            _json_response(
                self,
                200,
                {
                    "ok": True,
                    "excel_name": payload.excel_name,
                    "pdf_name": payload.pdf_name,
                    "headers": headers,
                    "rows": compared,
                    "pdf_token": token,
                    "row_count": len(compared),
                },
            )
        except Exception as exc:
            _json_response(self, 400, {"ok": False, "error": html.escape(str(exc))})


def main() -> None:
    if not INDEX_FILE.exists():
        raise SystemExit(f"Missing file: {INDEX_FILE}")
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    print(f"Open: http://{HOST}:{PORT}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
