#!/usr/bin/env python3
from __future__ import annotations

import json
from pathlib import Path

from pdf_excel_proofread_server import DEFAULT_DICT_FILE, parse_excel_keep_headers


def main() -> None:
    if not DEFAULT_DICT_FILE.exists():
        raise SystemExit(f"Missing source file: {DEFAULT_DICT_FILE}")

    headers, rows = parse_excel_keep_headers(DEFAULT_DICT_FILE.read_bytes())
    payload = {
        "ok": True,
        "headers": headers,
        "rows": rows,
        "row_count": len(rows),
        "source_file": DEFAULT_DICT_FILE.name,
    }

    out = Path(__file__).resolve().parent / "dictionary.json"
    out.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    print(f"Wrote {out} with {len(rows)} rows")


if __name__ == "__main__":
    main()
