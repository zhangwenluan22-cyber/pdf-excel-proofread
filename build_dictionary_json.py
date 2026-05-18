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

    root = Path(__file__).resolve().parent
    json_text = json.dumps(payload, ensure_ascii=False)

    out = root / "dictionary.json"
    out.write_text(json_text, encoding="utf-8")

    js_out = root / "dictionary_data.js"
    js_out.write_text(
        "window.__GRAMMAR_DICTIONARY_PAYLOAD__ = " + json_text + ";\n",
        encoding="utf-8",
    )

    print(f"Wrote {out} with {len(rows)} rows")
    print(f"Wrote {js_out} with {len(rows)} rows")


if __name__ == "__main__":
    main()
