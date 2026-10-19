#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from docx_common import DocxModel, ParagraphInfo, find_body_start, find_body_start

CAPTION_RE = re.compile(r"^(图|表|续表)\s*\d+\.\d+\s+\S")
HEADING_STYLE = {"1": 1, "2": 2, "3": 3}


def first_fmt(p: ParagraphInfo) -> dict:
    return p.runs[0].rpr if p.runs else {}


def font_of(fmt: dict) -> dict:
    return fmt.get("fonts", {}) or {}


def is_caption(p: ParagraphInfo) -> bool:
    return bool(CAPTION_RE.match(p.text.strip()))


def is_code(p: ParagraphInfo) -> bool:
    name = (p.style_name or "").lower()
    return p.style_id in {"af1", "HTML"} or "code" in name


def is_body_candidate(p: ParagraphInfo) -> bool:
    t = p.text.strip()
    if len(t) < 15 or p.images or is_caption(p) or is_code(p):
        return False
    if p.style_id in {"1", "2", "3", "TOC", "TOC1", "TOC2", "TOC3"}:
        return False
    if p.style_id.startswith("TOC"):
        return False
    if t in {"核心代码如下所示：", "参考文献", "致谢", "结  论"}:
        return False
    return True


def body_issues(p: ParagraphInfo) -> list[str]:
    issues: list[str] = []
    ppr = p.ppr or {}
    ind = ppr.get("ind", {}) or {}
    first_chars = ind.get("firstLineChars")
    first = ind.get("firstLine")
    try:
        first_chars_i = int(first_chars) if first_chars is not None else None
    except ValueError:
        first_chars_i = None
    try:
        first_i = int(first) if first is not None else None
    except ValueError:
        first_i = None
    if not ((first_chars_i is not None and first_chars_i >= 200) or (first_i is not None and first_i >= 480)):
        issues.append(f"首行缩进不足（firstLineChars={first_chars}, firstLine={first}）")

    sp = ppr.get("spacing", {}) or {}
    if str(sp.get("line", "")) != "440" or sp.get("lineRule") != "exact":
        issues.append(f"正文行距不是固定22磅（line={sp.get('line')}, lineRule={sp.get('lineRule')}）")

    fmt = first_fmt(p)
    if str(fmt.get("sz", "")) != "24":
        issues.append(f"正文字号不是小四（sz={fmt.get('sz')}）")
    fonts = font_of(fmt)
    east = fonts.get("eastAsia", "")
    ascii_font = fonts.get("ascii", "") or fonts.get("hAnsi", "")
    if east and "宋体" not in east:
        issues.append(f"正文中文字体不是宋体（eastAsia={east}）")
    if ascii_font and "Times New Roman" not in ascii_font:
        issues.append(f"正文英文/数字字体不是Times New Roman（ascii={ascii_font}）")
    return issues


def heading_issues(p: ParagraphInfo) -> list[str]:
    level = HEADING_STYLE.get(p.style_id)
    if level is None:
        try:
            level = int(p.ppr.get("outlineLvl", -1)) + 1
        except (TypeError, ValueError):
            level = 0
    expected_sz = {1: "36", 2: "30", 3: "24"}.get(level)
    issues: list[str] = []
    fmt = first_fmt(p)
    if expected_sz and str(fmt.get("sz", "")) != expected_sz:
        issues.append(f"标题字号不符（level={level}, sz={fmt.get('sz')}, 期望={expected_sz}）")
    fonts = font_of(fmt)
    east = fonts.get("eastAsia", "")
    if east and "黑体" not in east:
        issues.append(f"标题中文字体不是黑体（eastAsia={east}）")
    sp = (p.ppr or {}).get("spacing", {}) or {}
    if str(sp.get("line", "")) != "440" or sp.get("lineRule") != "exact":
        issues.append(f"标题行距不是固定22磅（line={sp.get('line')}, lineRule={sp.get('lineRule')}）")
    return issues


def caption_issues(p: ParagraphInfo) -> list[str]:
    issues: list[str] = []
    fmt = first_fmt(p)
    if str(fmt.get("sz", "")) != "21":
        issues.append(f"图/表题字号不是五号（sz={fmt.get('sz')}）")
    fonts = font_of(fmt)
    east = fonts.get("eastAsia", "")
    if east and "黑体" not in east:
        issues.append(f"图/表题中文字体不是黑体（eastAsia={east}）")
    ppr = p.ppr or {}
    is_continued = p.text.strip().startswith("续表")
    expected_jc = "right" if is_continued else "center"
    if ppr.get("jc") != expected_jc:
        issues.append(f"图/表题对齐不符（jc={ppr.get('jc')}, 期望={expected_jc}）")
    sp = ppr.get("spacing", {}) or {}
    if str(sp.get("line", "")) != "400" or sp.get("lineRule") != "exact":
        issues.append(f"图/表题行距不是固定20磅（line={sp.get('line')}, lineRule={sp.get('lineRule')}）")
    return issues


def code_issues(p: ParagraphInfo) -> list[str]:
    issues: list[str] = []
    for run in p.runs:
        fmt = run.rpr
        if str(fmt.get("sz", "")) != "21":
            issues.append(f"代码字号不是五号（sz={fmt.get('sz')}）")
            break
    fonts = font_of(first_fmt(p))
    ascii_font = fonts.get("ascii", "") or fonts.get("hAnsi", "")
    if ascii_font and "Times New Roman" not in ascii_font:
        issues.append(f"代码英文/数字字体不是Times New Roman（ascii={ascii_font}）")
    return issues


def main() -> int:
    ap = argparse.ArgumentParser(description="Inspect DOCX structure, styles, paragraphs and runs.")
    ap.add_argument("docx", help="path to .docx")
    ap.add_argument("--json", dest="json_path", help="write full JSON report")
    ap.add_argument("--max-findings", type=int, default=80)
    args = ap.parse_args()

    model = DocxModel(args.docx)
    body_start = find_body_start(model)
    body_start_idx = body_start.index if body_start else 0
    body_start_page = body_start.inferred_page if body_start else 1

    def body_page(p: ParagraphInfo) -> int:
        return p.inferred_page - body_start_page + 1

    ref_start = next((p.index for p in model.paragraphs if p.style_id == "1" and p.text.strip() == "参考文献"), None)
    ref_end = next((p.index for p in model.paragraphs if p.style_id == "1" and p.text.strip() == "致谢"), None)
    findings: list[dict] = []
    for p in model.paragraphs:
        if p.index < body_start_idx:
            continue
        if ref_start is not None and p.index >= ref_start and (ref_end is None or p.index < ref_end):
            continue
        t = p.text.strip()
        if not t:
            continue
        page = body_page(p)
        if is_caption(p):
            for msg in caption_issues(p):
                findings.append({"kind": "caption", "paragraph": p.index, "page": page, "text": t[:80], "issue": msg})
        elif is_code(p):
            for msg in code_issues(p):
                findings.append({"kind": "code", "paragraph": p.index, "page": page, "text": t[:80], "issue": msg})
        elif p.style_id in HEADING_STYLE or p.ppr.get("outlineLvl") in {"0", "1", "2"}:
            for msg in heading_issues(p):
                findings.append({"kind": "heading", "paragraph": p.index, "page": page, "text": t[:80], "issue": msg})
        elif is_body_candidate(p):
            for msg in body_issues(p):
                findings.append({"kind": "body", "paragraph": p.index, "page": page, "text": t[:80], "issue": msg})

    headings = [p for p in model.headings() if p.index >= body_start_idx]
    data = {
        "file": str(Path(args.docx).resolve()),
        "paragraphs": len(model.paragraphs),
        "tables": len(model.tables),
        "headings": len(headings),
        "toc_entries": len(model.toc_entries()),
        "update_fields": model.has_update_fields,
        "body_start_paragraph": body_start_idx,
        "findings": findings,
        "heading_list": [
            {
                "paragraph": p.index,
                "style": p.style_id,
                "page": body_page(p),
                "text": p.text.strip(),
                "sz": first_fmt(p).get("sz"),
                "fonts": font_of(first_fmt(p)),
                "spacing": (p.ppr or {}).get("spacing", {}),
            }
            for p in headings
        ],
    }

    print(f"File: {data['file']}")
    print(f"Paragraphs: {data['paragraphs']}  Tables: {data['tables']}  Headings: {data['headings']}  TOC entries: {data['toc_entries']}")
    print(f"updateFields: {data['update_fields']}  body starts at top-level paragraph #{body_start_idx}")
    print("\nHeadings:")
    for h in data["heading_list"]:
        print(f"  p{h['paragraph']:>4}  page~{h['page']:<3}  style={h['style']:<3}  sz={h['sz']!s:<4}  {h['text'][:60]}")
    print(f"\nFormat findings: {len(findings)}")
    for f in findings[: args.max_findings]:
        print(f"  [{f['kind']}] p{f['paragraph']} page~{f['page']} {f['text']!r}: {f['issue']}")
    if len(findings) > args.max_findings:
        print(f"  ... {len(findings) - args.max_findings} more")

    if args.json_path:
        Path(args.json_path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"\nJSON written: {args.json_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

