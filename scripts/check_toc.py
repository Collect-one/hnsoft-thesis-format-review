#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from docx_common import DocxModel, ParagraphInfo, find_body_start


def norm(s: str) -> str:
    return re.sub(r"\s+", "", s or "")


def main() -> int:
    ap = argparse.ArgumentParser(description="Compare cached TOC page numbers with rendered page breaks.")
    ap.add_argument("docx")
    ap.add_argument("--json", dest="json_path")
    args = ap.parse_args()

    model = DocxModel(args.docx)
    body_start = find_body_start(model)
    body_start_page = body_start.inferred_page if body_start else 1

    headings = model.headings()
    heading_by_title = {norm(p.text): p for p in headings if p.text.strip()}
    toc = model.toc_entries()

    findings: list[dict] = []
    compared = 0
    for entry in toc:
        title = entry["title"]
        page_text = entry["page"]
        if not re.fullmatch(r"\d+", page_text):
            continue  # 摘要/Abstract 的罗马页码
        p = heading_by_title.get(norm(title))
        if p is None:
            continue
        expected = p.inferred_page - body_start_page + 1
        actual = int(page_text)
        compared += 1
        if expected != actual:
            findings.append({
                "severity": "important",
                "kind": "toc_page_mismatch",
                "title": title,
                "toc_page": actual,
                "rendered_page": expected,
                "message": f"目录写 {actual}，按文档分页标记推算为 {expected}。TOC 缓存可能已过期。",
            })


    # TOC entry numbers and Latin text should use Times New Roman.
    for p in model.paragraphs:
        if not p.style_id.startswith("TOC"):
            continue
        for run in p.runs:
            if not re.search(r"[A-Za-z0-9]", run.text):
                continue
            fonts = (run.rpr or {}).get("fonts", {}) or {}
            ascii_font = fonts.get("ascii") or fonts.get("hAnsi")
            if ascii_font and "Times New Roman" not in ascii_font:
                findings.append({
                    "severity": "suggestion",
                    "kind": "toc_number_font",
                    "paragraph": p.index,
                    "message": f"目录条目 {p.text.strip()[:40]!r} 的数字/英文不是 Times New Roman（ascii={ascii_font}）。",
                })
                break

    if not model.has_update_fields:
        findings.append({
            "severity": "suggestion",
            "kind": "update_fields",
            "message": "settings.xml 未设置 updateFields；打开 Word 时不会自动更新目录域，提交前需手动更新整个目录。",
        })

    data = {
        "file": str(Path(args.docx).resolve()),
        "toc_entries": len(toc),
        "compared_numeric_entries": compared,
        "update_fields": model.has_update_fields,
        "findings": findings,
    }

    print(f"TOC entries: {len(toc)}  compared: {compared}  Findings: {len(findings)}")
    for f in findings:
        print(f"  [{f['severity']}] {f['kind']}: {f['message']}")
    if args.json_path:
        Path(args.json_path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"JSON written: {args.json_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
