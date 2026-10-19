#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from docx_common import DocxModel, ParagraphInfo, find_body_start


def ind_is_top(p: ParagraphInfo) -> bool:
    ind = (p.ppr or {}).get("ind", {}) or {}
    fc = ind.get("firstLineChars")
    fl = ind.get("firstLine")
    return fc in (None, "0") and fl in (None, "0")


def font_has(fmt: dict, key: str, needle: str) -> bool:
    fonts = fmt.get("fonts", {}) or {}
    return needle in (fonts.get(key, "") or "")


def main() -> int:
    ap = argparse.ArgumentParser(description="Check front matter labels and conclusion/reference/acknowledgement heading spacing.")
    ap.add_argument("docx")
    ap.add_argument("--json", dest="json_path")
    args = ap.parse_args()

    model = DocxModel(args.docx)
    body_start = find_body_start(model)
    body_idx = body_start.index if body_start else len(model.paragraphs) + 1
    findings: list[dict] = []

    def find_before(pred, before=body_idx):
        for p in model.paragraphs:
            if p.index >= before:
                break
            if pred(p):
                return p
        return None

    abstract = find_before(lambda p: p.text.strip().startswith("摘要："))
    if abstract:
        if not abstract.runs:
            findings.append({"severity": "important", "kind": "missing_run", "paragraph": abstract.index, "message": "摘要段落没有文本 run。"})
        else:
            fmt = abstract.runs[0].rpr
            if not font_has(fmt, "eastAsia", "黑体"):
                findings.append({"severity": "important", "kind": "abstract_label_font", "paragraph": abstract.index, "message": "“摘要：”不是黑体。"})
            if str(fmt.get("sz", "")) != "24":
                findings.append({"severity": "important", "kind": "abstract_label_size", "paragraph": abstract.index, "message": f"“摘要：”字号不是小四（sz={fmt.get('sz')}）。"})
        if not ind_is_top(abstract):
            findings.append({"severity": "important", "kind": "abstract_label_indent", "paragraph": abstract.index, "message": "“摘要：”没有顶头。"})

    keywords_cn = find_before(lambda p: p.text.strip().startswith("关键词："))
    if keywords_cn:
        if keywords_cn.runs:
            fmt = keywords_cn.runs[0].rpr
            if not font_has(fmt, "eastAsia", "黑体"):
                findings.append({"severity": "important", "kind": "keywords_label_font", "paragraph": keywords_cn.index, "message": "中文“关键词：”不是黑体。"})
            if str(fmt.get("sz", "")) != "24":
                findings.append({"severity": "important", "kind": "keywords_label_size", "paragraph": keywords_cn.index, "message": f"中文“关键词：”字号不是小四（sz={fmt.get('sz')}）。"})
        if not ind_is_top(keywords_cn):
            findings.append({"severity": "important", "kind": "keywords_label_indent", "paragraph": keywords_cn.index, "message": "中文“关键词：”没有顶头。"})

    abstract_en = find_before(lambda p: p.text.strip().startswith("Abstract:"))
    if abstract_en:
        title = None
        for p in reversed(model.paragraphs):
            if p.index >= abstract_en.index:
                continue
            if p.text.strip() and not p.text.strip().startswith("摘要："):
                title = p
                break
        if title and title.runs:
            fmt = title.runs[0].rpr
            if not (font_has(fmt, "ascii", "Times New Roman") or font_has(fmt, "hAnsi", "Times New Roman")):
                findings.append({"severity": "important", "kind": "english_title_font", "paragraph": title.index, "message": "英文题目不是 Times New Roman。"})
            if str(fmt.get("sz", "")) != "36":
                findings.append({"severity": "important", "kind": "english_title_size", "paragraph": title.index, "message": f"英文题目字号不是小二（sz={fmt.get('sz')}）。"})
            if fmt.get("b") not in ("1", True):
                findings.append({"severity": "important", "kind": "english_title_bold", "paragraph": title.index, "message": "英文题目没有加粗。"})
            if (title.ppr or {}).get("jc") != "center":
                findings.append({"severity": "important", "kind": "english_title_align", "paragraph": title.index, "message": "英文题目没有居中。"})
            sp = (title.ppr or {}).get("spacing", {}) or {}
            if str(sp.get("line", "")) != "440" or sp.get("lineRule") != "exact":
                findings.append({"severity": "important", "kind": "english_title_spacing", "paragraph": title.index, "message": f"英文题目行距不是固定22磅（line={sp.get('line')}, lineRule={sp.get('lineRule')}）。"})

    keywords_en = find_before(lambda p: p.text.strip().startswith("Keywords:"))
    if keywords_en:
        if keywords_en.runs:
            fmt = keywords_en.runs[0].rpr
            if not (font_has(fmt, "ascii", "Times New Roman") or font_has(fmt, "hAnsi", "Times New Roman")):
                findings.append({"severity": "important", "kind": "keywords_en_font", "paragraph": keywords_en.index, "message": "英文 Keywords: 不是 Times New Roman。"})
            if str(fmt.get("sz", "")) != "24":
                findings.append({"severity": "important", "kind": "keywords_en_size", "paragraph": keywords_en.index, "message": f"英文 Keywords: 字号不是小四（sz={fmt.get('sz')}）。"})
            if fmt.get("b") not in ("1", True):
                findings.append({"severity": "important", "kind": "keywords_en_bold", "paragraph": keywords_en.index, "message": "英文 Keywords: 没有加粗。"})
        if not ind_is_top(keywords_en):
            findings.append({"severity": "important", "kind": "keywords_en_indent", "paragraph": keywords_en.index, "message": "英文 Keywords: 没有顶头。"})

    for title_text in ("结论", "参考文献", "致谢"):
        heading = next(
            (
                p for p in model.paragraphs
                if p.style_id == "1" and re.sub(r"\s+", "", p.text) == title_text
            ),
            None,
        )
        if not heading:
            continue
        sp = (heading.ppr or {}).get("spacing", {}) or {}
        before_lines = sp.get("beforeLines")
        after_lines = sp.get("afterLines")
        before = sp.get("before")
        after = sp.get("after")
        before_ok = before_lines == "100" or (before is not None and int(before) >= 300)
        after_ok = after_lines == "100" or (after is not None and int(after) >= 300)
        if not (before_ok and after_ok):
            findings.append({
                "severity": "important",
                "kind": "heading_spacing",
                "paragraph": heading.index,
                "message": f"{title_text}标题段前/段后不是各1行（beforeLines={before_lines}, afterLines={after_lines}, before={before}, after={after}）。",
            })

    data = {
        "file": str(Path(args.docx).resolve()),
        "findings": findings,
    }
    print(f"Front matter findings: {len(findings)}")
    for f in findings:
        print(f"  [{f['severity']}] {f['kind']}: {f['message']}")
    if args.json_path:
        Path(args.json_path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"JSON written: {args.json_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
