#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
from collections import defaultdict
from pathlib import Path

from docx_common import DocxModel, ParagraphInfo, find_body_start, find_body_start

CAPTION_RE = re.compile(r"^(图|表|续表)\s*(\d+)\.(\d+)\s+(.+)$")
REF_RE = re.compile(r"(图|表)\s?(\d+)\.(\d+)")


def caption_entries(model: DocxModel) -> list[dict]:
    out = []
    for p in model.paragraphs:
        t = p.text.strip()
        m = CAPTION_RE.match(t)
        if not m:
            continue
        kind, major, minor, title = m.group(1), m.group(2), m.group(3), m.group(4).strip()
        logical_kind = "表" if kind == "续表" else kind
        out.append(
            {
                "paragraph": p.index,
                "page": p.inferred_page,
                "kind": kind,
                "logical_kind": logical_kind,
                "number": f"{major}.{minor}",
                "title": title,
                "text": t,
            }
        )
    return out


def references(model: DocxModel) -> list[dict]:
    out = []
    for p in model.paragraphs:
        t = p.text.strip()
        if not t or CAPTION_RE.match(t):
            continue
        if p.style_id.startswith("TOC"):
            continue
        for m in REF_RE.finditer(t):
            out.append(
                {
                    "paragraph": p.index,
                    "page": p.inferred_page,
                    "kind": m.group(1),
                    "number": f"{m.group(2)}.{m.group(3)}",
                    "text": t[:120],
                }
            )
    return out


def nearest_image(model: DocxModel, caption: ParagraphInfo) -> tuple[list[str], int | None]:
    for p in reversed(model.paragraphs):
        if p.index >= caption.index:
            continue
        if caption.index - p.index > 8:
            break
        if p.images:
            return p.images, p.index
    return [], None


def main() -> int:
    ap = argparse.ArgumentParser(description="Check figure/table captions, references and image reuse.")
    ap.add_argument("docx")
    ap.add_argument("--json", dest="json_path")
    args = ap.parse_args()

    model = DocxModel(args.docx)
    caps = caption_entries(model)
    refs = references(model)

    # body page normalization
    body_start = find_body_start(model)
    body_start_page = body_start.inferred_page if body_start else 1

    def body_page(p: int) -> int:
        return p - body_start_page + 1 if body_start else p

    findings: list[dict] = []
    cap_keys = {(c["logical_kind"], c["number"]) for c in caps}
    ref_keys = {(r["kind"], r["number"]) for r in refs}

    # missing captions / unreferenced captions
    for kind, number in sorted(ref_keys - cap_keys):
        findings.append({
            "severity": "important",
            "kind": "missing_caption",
            "number": f"{kind}{number}",
            "message": f"正文引用了 {kind}{number}，但没有找到对应题注。",
        })
    for kind, number in sorted(cap_keys - ref_keys):
        # 续表 not separately referenced
        if any(c["kind"] == "续表" and c["logical_kind"] == kind and c["number"] == number for c in caps):
            continue
        findings.append({
            "severity": "suggestion",
            "kind": "unreferenced_caption",
            "number": f"{kind}{number}",
            "message": f"{kind}{number} 有题注，但正文没有出现对应引用。",
        })

    # duplicate captions (except 续表)
    grouped = defaultdict(list)
    for c in caps:
        if c["kind"] != "续表":
            grouped[(c["logical_kind"], c["number"])].append(c)
    for (kind, number), items in grouped.items():
        if len(items) > 1:
            findings.append({
                "severity": "important",
                "kind": "duplicate_caption",
                "number": f"{kind}{number}",
                "message": f"{kind}{number} 出现 {len(items)} 次题注。",
                "paragraphs": [x["paragraph"] for x in items],
            })

    # figure image mapping and reuse
    figure_images: dict[str, list[dict]] = defaultdict(list)
    for c in caps:
        if c["logical_kind"] != "图":
            continue
        p = next((x for x in model.paragraphs if x.index == c["paragraph"]), None)
        if p is None:
            continue
        rids, image_paragraph = nearest_image(model, p)
        if not rids:
            findings.append({
                "severity": "important",
                "kind": "missing_image",
                "number": f"图{c['number']}",
                "message": f"图{c['number']} 题注前没有找到图片。",
                "paragraph": c["paragraph"],
            })
            continue
        for rid in rids:
            target = model.rel_target(rid)
            figure_images[target].append({
                "number": c["number"],
                "title": c["title"],
                "caption_paragraph": c["paragraph"],
                "image_paragraph": image_paragraph,
                "page": body_page(c["page"]),
                "rid": rid,
            })

    for target, items in figure_images.items():
        numbers = sorted({x["number"] for x in items}, key=lambda s: tuple(int(v) for v in s.split(".")))
        if len(numbers) > 1:
            findings.append({
                "severity": "suggestion",
                "kind": "image_reused",
                "target": target,
                "message": "以下图使用了同一张图片，请确认是否为同一页面的不同功能展示：" + "、".join(f"图{n}" for n in numbers),
                "figures": items,
            })

    # numbering gaps
    for kind in ("图", "表"):
        nums = sorted({tuple(int(v) for v in c["number"].split(".")) for c in caps if c["logical_kind"] == kind})
        if not nums:
            continue
        majors = sorted({n[0] for n in nums})
        for major in majors:
            minors = sorted(n[1] for n in nums if n[0] == major)
            if minors and minors != list(range(1, max(minors) + 1)):
                findings.append({
                    "severity": "important",
                    "kind": "numbering_gap",
                    "number": f"{kind}{major}.x",
                    "message": f"{kind}{major} 的编号不连续：{minors}",
                })

    data = {
        "file": str(Path(args.docx).resolve()),
        "captions": caps,
        "references": refs,
        "findings": findings,
        "figure_images": {k: v for k, v in figure_images.items()},
    }

    print(f"Captions: {len(caps)}  References: {len(refs)}  Findings: {len(findings)}")
    for f in findings:
        print(f"  [{f['severity']}] {f['kind']}: {f['message']}")
    if args.json_path:
        Path(args.json_path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"JSON written: {args.json_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
