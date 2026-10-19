#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path

from docx_common import DocxModel, W, TableInfo, lname


def cell_borders(tr) -> list[str]:
    out: list[str] = []
    for tc in tr.findall(W + "tc"):
        tcpr = tc.find(W + "tcPr")
        if tcpr is None:
            continue
        borders = tcpr.find(W + "tcBorders")
        if borders is None:
            continue
        for b in borders:
            out.append(f"{lname(b.tag)}:{b.get(W + 'val', '')}:sz={b.get(W + 'sz', '')}")
    return out


def first_cell_paragraph(tbl):
    tc = tbl.find(".//" + W + "tc")
    if tc is None:
        return None
    return tc.find(W + "p")


def main() -> int:
    ap = argparse.ArgumentParser(description="Check three-line table structure, width, captions and row integrity.")
    ap.add_argument("docx")
    ap.add_argument("--json", dest="json_path")
    args = ap.parse_args()

    model = DocxModel(args.docx)
    findings: list[dict] = []
    table_rows = []

    for t in model.tables:
        cm = t.grid_width_twips / 567.0 if t.grid_width_twips else 0.0
        row_cells = [len(r.cell_texts) for r in t.rows]
        header_cells = row_cells[0] if row_cells else 0
        multi_para_rows = []
        for idx, r in enumerate(t.rows):
            counts = [len(c) for c in r.cells]
            # A row where most cells contain multiple paragraphs often means two logical rows were merged.
            if counts and sum(1 for c in counts if c >= 2) >= max(3, header_cells - 1):
                multi_para_rows.append({"row": idx, "paragraphs_per_cell": counts, "text": " | ".join(r.cell_texts)})

        info = {
            "index": t.index,
            "caption": t.caption,
            "rows": len(t.rows),
            "grid_width_twips": t.grid_width_twips,
            "width_cm": round(cm, 2),
            "tbl_width": t.tbl_width,
            "header_repeat": t.header_repeat,
            "borders": t.borders,
            "row_cell_counts": row_cells,
            "multi_paragraph_rows": multi_para_rows,
        }
        table_rows.append(info)

        label = t.caption or f"表#{t.index}"
        if not t.caption:
            findings.append({"severity": "important", "kind": "missing_caption", "table": t.index, "message": f"{label} 没有找到表题。"})
        if cm and abs(cm - 16.0) > 0.2:
            findings.append({"severity": "suggestion", "kind": "table_width", "table": t.index, "message": f"{label} 宽度约 {cm:.2f}cm，建议 16cm。"})
        top = t.borders.get("top", {})
        bottom = t.borders.get("bottom", {})
        for name, border in (("上边线", top), ("下边线", bottom)):
            if not border or border.get("val") != "single":
                findings.append({"severity": "important", "kind": "table_border", "table": t.index, "message": f"{label} 缺少{name}或{name}不是单实线。"})
            else:
                try:
                    sz = int(border.get("sz", "0"))
                except ValueError:
                    sz = 0
                if sz < 12:
                    findings.append({"severity": "important", "kind": "table_border", "table": t.index, "message": f"{label} {name}线宽 {sz/8:.2f}pt，建议 1.5pt。"})
        if len(set(row_cells)) > 1:
            findings.append({"severity": "important", "kind": "row_cell_count", "table": t.index, "message": f"{label} 各行列数不一致：{row_cells}"})
        if multi_para_rows:
            findings.append({
                "severity": "important",
                "kind": "merged_row_suspect",
                "table": t.index,
                "message": f"{label} 有 {len(multi_para_rows)} 行每个单元格含多段，可能把两个逻辑行合并了；请检查。",
                "rows": multi_para_rows,
            })
        if len(t.rows) >= 8 and not t.header_repeat:
            findings.append({"severity": "suggestion", "kind": "header_repeat", "table": t.index, "message": f"{label} 行数较多但未设置标题行重复。"})

        # font / spacing sample
        p = first_cell_paragraph(t.element)
        if p is not None:
            ppr = p.find(W + "pPr")
            sp = ppr.find(W + "spacing") if ppr is not None else None
            if sp is not None and (sp.get(W + "line") not in (None, "360")):
                findings.append({"severity": "suggestion", "kind": "table_line_spacing", "table": t.index, "message": f"{label} 表内行距不是 1.5 倍（line={sp.get(W + 'line')}）。"})
            sizes = [sz.get(W + "val") for sz in t.element.iter(W + "sz") if sz.get(W + "val")]
            bad_sizes = sorted({s for s in sizes if s not in {"21", "18"}})
            if bad_sizes:
                findings.append({"severity": "suggestion", "kind": "table_font_size", "table": t.index, "message": f"{label} 表内存在非五号字号：{bad_sizes}"})

    data = {
        "file": str(Path(args.docx).resolve()),
        "tables": table_rows,
        "findings": findings,
    }

    print(f"Tables: {len(model.tables)}  Findings: {len(findings)}")
    for f in findings:
        print(f"  [{f['severity']}] {f['kind']}: {f['message']}")
    if args.json_path:
        Path(args.json_path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"JSON written: {args.json_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
