#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from datetime import date
from pathlib import Path

from docx_common import DocxModel, ParagraphInfo


def has_numbering(p: ParagraphInfo) -> bool:
    ppr = p.element.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")
    return ppr is not None and ppr.find("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr") is not None


def main() -> int:
    ap = argparse.ArgumentParser(description="Check reference list numbering, years, punctuation and EB/OL fields.")
    ap.add_argument("docx")
    ap.add_argument("--json", dest="json_path")
    args = ap.parse_args()

    model = DocxModel(args.docx)
    def norm_heading(text: str) -> str:
        return re.sub(r"\s+", "", text or "")

    start = next(
        (
            p.index
            for p in model.paragraphs
            if not p.style_id.startswith("TOC") and norm_heading(p.text) == "参考文献"
        ),
        None,
    )
    if start is None:
        data = {
            "file": str(Path(args.docx).resolve()),
            "count": 0,
            "entries": [],
            "findings": [{
                "severity": "suggestion",
                "kind": "no_references",
                "message": "未找到“参考文献”标题，跳过参考文献检查。",
            }],
        }
        print("Reference heading not found; skipped.")
        if args.json_path:
            Path(args.json_path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        return 0
    end = next(
        (
            p.index
            for p in model.paragraphs
            if p.index > start and not p.style_id.startswith("TOC")
            and (norm_heading(p.text).startswith("致谢") or norm_heading(p.text).startswith("附录"))
        ),
        None,
    )
    entries = [
        p for p in model.paragraphs
        if p.index > start and (end is None or p.index < end)
        and p.text.strip() and p.style_id != "1"
        and re.search(r"\[[A-Z/]+\]", p.text)
    ]

    findings: list[dict] = []
    current_year = date.today().year
    text_entries = []

    for idx, p in enumerate(entries, start=1):
        t = p.text.strip()
        explicit = re.match(r"^\[(\d+)\]\s*", t)
        number = int(explicit.group(1)) if explicit else idx
        body = re.sub(r"^\[\d+\]\s*", "", t)
        text_entries.append({"number": number, "text": t, "paragraph": p.index, "auto_numbered": has_numbering(p)})

        if explicit and number != idx:
            findings.append({
                "severity": "important",
                "kind": "numbering",
                "paragraph": p.index,
                "message": f"第 {idx} 条的显式编号是 [{number}]，编号不连续。",
            })

        if re.search(r"等\.[^\s\[]", body):
            findings.append({
                "severity": "suggestion",
                "kind": "spacing",
                "paragraph": p.index,
                "message": "“等.”后缺空格。",
            })
        if re.search(r"\[[A-Z/]+\]\.[^\s]", body):
            findings.append({
                "severity": "suggestion",
                "kind": "spacing",
                "paragraph": p.index,
                "message": "文献类型标识（如 [J].）后缺空格。",
            })
        if re.search(r":\d", body):
            findings.append({
                "severity": "suggestion",
                "kind": "spacing",
                "paragraph": p.index,
                "message": "页码前的冒号后缺空格。",
            })
        if "[EB/OL]" in body and not re.search(r"https?://|DOI:", body, re.I):
            findings.append({
                "severity": "suggestion",
                "kind": "ebol_missing_url",
                "paragraph": p.index,
                "message": "EB/OL 条目缺 URL 或 DOI。",
            })
        if "[J/OL]" in body and "网络首发" in body and not re.search(r"\d+\s*\(\s*\d+\s*\)", body):
            findings.append({
                "severity": "suggestion",
                "kind": "online_first",
                "paragraph": p.index,
                "message": "J/OL 网络首发条目缺卷(期)或页码，建议补 DOI/URL 或引用日期。",
            })
        if "[M]" in body and not re.search(r"[:：]\s*\S+,\s*\d{4}", body):
            findings.append({
                "severity": "suggestion",
                "kind": "book_format",
                "paragraph": p.index,
                "message": "专著 [M] 条目建议按“出版地: 出版者, 出版年”补齐。",
            })
        years = [int(y) for y in re.findall(r"(?<!\d)(20\d{2})(?!\d)", body)]
        for y in years:
            if y > current_year:
                findings.append({
                    "severity": "suggestion",
                    "kind": "future_year",
                    "paragraph": p.index,
                    "message": f"文献年份 {y} 晚于当前年份 {current_year}。",
                })

    # duplicates by first 30 chars
    seen = {}
    for e in text_entries:
        key = re.sub(r"\s+", "", e["text"])[:40]
        if key in seen:
            findings.append({
                "severity": "important",
                "kind": "duplicate",
                "paragraph": e["paragraph"],
                "message": f"疑似与第 {seen[key]} 条重复。",
            })
        else:
            seen[key] = e["number"]

    data = {
        "file": str(Path(args.docx).resolve()),
        "count": len(text_entries),
        "entries": text_entries,
        "findings": findings,
    }

    print(f"Reference entries: {len(text_entries)}  Findings: {len(findings)}")
    for f in findings:
        print(f"  [{f['severity']}] {f['kind']}: {f['message']}")
    if args.json_path:
        Path(args.json_path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"JSON written: {args.json_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
