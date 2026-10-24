#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent


def run_json(script: str, docx: str, out_json: Path) -> tuple[dict, str]:
    cmd = [sys.executable, "-B", str(SCRIPT_DIR / script), docx, "--json", str(out_json)]
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    proc = subprocess.run(
        cmd, cwd=str(SCRIPT_DIR), capture_output=True, text=True,
        encoding="utf-8", errors="replace", env=env, timeout=180,
    )
    if proc.returncode != 0:
        raise RuntimeError(f"{script} failed:\nSTDOUT:\n{proc.stdout}\nSTDERR:\n{proc.stderr}")
    data = json.loads(out_json.read_text(encoding="utf-8")) if out_json.exists() else {}
    return data, proc.stdout


def fmt_para_list(paragraphs: list[int], limit: int = 12) -> str:
    ps = sorted({int(p) for p in paragraphs if p is not None})
    if not ps:
        return ""
    shown = "、".join(str(p) for p in ps[:limit])
    if len(ps) > limit:
        shown += f" 等 {len(ps)} 处"
    return shown


def main() -> int:
    ap = argparse.ArgumentParser(description="Run all thesis format checks and write a Markdown report.")
    ap.add_argument("docx")
    ap.add_argument("--out", required=True, help="output directory")
    ap.add_argument("--title", default="论文审查报告")
    args = ap.parse_args()

    docx = str(Path(args.docx).resolve())
    out = Path(args.out).resolve()
    out.mkdir(parents=True, exist_ok=True)

    specs = [
        ("inspect_docx.py", "inspect"),
        ("check_figures.py", "figures"),
        ("check_tables.py", "tables"),
        ("check_toc.py", "toc"),
        ("check_refs.py", "refs"),
        ("check_front_matter.py", "front"),
    ]
    data: dict[str, dict] = {}
    logs: dict[str, str] = {}
    for script, key in specs:
        data[key], logs[key] = run_json(script, docx, out / f"{key}.json")

    all_findings: list[dict] = []
    for key in ("inspect", "figures", "tables", "toc", "refs", "front"):
        for f in data.get(key, {}).get("findings", []):
            item = dict(f)
            item["section"] = key
            all_findings.append(item)

    severity_order = {"blocking": 0, "important": 1, "suggestion": 2}
    all_findings.sort(key=lambda f: (severity_order.get(f.get("severity", "suggestion"), 9), f.get("section", "")))

    counts = Counter(f.get("severity", "suggestion") for f in all_findings)
    lines: list[str] = []
    lines.append(f"# {args.title}")
    lines.append("")
    lines.append(f"- 论文文件：`{docx}`")
    lines.append(f"- 生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"- 文档规模：{data['inspect'].get('paragraphs', '?')} 个顶层段落，{data['inspect'].get('tables', '?')} 张表，{data['inspect'].get('headings', '?')} 个标题，{data['inspect'].get('toc_entries', '?')} 条目录")
    lines.append(f"- 问题统计：阻断 {counts.get('blocking', 0)}，重要 {counts.get('important', 0)}，建议 {counts.get('suggestion', 0)}")
    lines.append("")

    if counts.get("blocking", 0) or counts.get("important", 0):
        lines.append("## 需要优先处理")
        for f in all_findings:
            if f.get("severity") not in {"blocking", "important"}:
                continue
            lines.append(f"- **[{f.get('severity')}] {f.get('message', f.get('kind'))}**")
            for key in ("title", "toc_page", "rendered_page", "number", "paragraph"):
                if key in f:
                    lines.append(f"  - {key}: {f[key]}")
            if f.get("rows"):
                lines.append("  - 多段行：")
                for row in f["rows"]:
                    lines.append(f"    - 第 {row['row']} 行：{row['text'][:120]}")
        lines.append("")
    else:
        lines.append("## 需要优先处理")
        lines.append("未发现阻断或重要问题。")
        lines.append("")

    # Front matter
    lines.append("## 前置部分")
    front_findings = [f for f in all_findings if f.get("section") == "front"]
    if not front_findings:
        lines.append("摘要、关键词、英文摘要和结论/参考文献/致谢标题未发现脚本范围内的格式异常。")
    else:
        for f in front_findings:
            lines.append(f"- p{f.get('paragraph')}：{f.get('message')}")
    lines.append("")

    # Format
    lines.append("## 正文字体与版式")
    fmt = [f for f in all_findings if f.get("section") == "inspect"]
    if not fmt:
        lines.append("脚本范围内未发现正文/标题/图题/代码的字体、字号、行距异常。")
    else:
        for f in fmt[:60]:
            lines.append(f"- p{f.get('paragraph')}（第 {f.get('page')} 页）{f.get('text','')!r}：{f.get('issue')}")
        if len(fmt) > 60:
            lines.append(f"- ……另有 {len(fmt)-60} 条，见 `inspect.json`。")
    lines.append("")

    # Figures
    lines.append("## 图表")
    fig_findings = [f for f in all_findings if f.get("section") == "figures"]
    table_findings = [f for f in all_findings if f.get("section") == "tables"]
    if not fig_findings:
        lines.append("- 图题、正文引用存在性和编号未发现缺失或重复；引用句与图片的实际页内位置、图片比例及图片内容仍需人工复核。")
    else:
        for f in fig_findings:
            lines.append(f"- **{f.get('message')}**")
            if f.get("figures"):
                for x in f["figures"]:
                    lines.append(f"  - 图{x.get('number')}：{x.get('title')}（题注段落 {x.get('caption_paragraph')}）")
    if not table_findings:
        lines.append("- 表格宽度、上下边线、表题和行结构未发现异常；表内分隔线和续表题仍需人工复核。")
    else:
        for f in table_findings:
            lines.append(f"- {f.get('message')}")
    lines.append("")

    # TOC
    lines.append("## 目录与分页")
    toc_findings = [f for f in all_findings if f.get("section") == "toc"]
    if not toc_findings:
        lines.append("目录缓存页码与分页标记一致。")
    else:
        for f in toc_findings:
            lines.append(f"- {f.get('message')}")
    lines.append("")

    # References
    lines.append("## 参考文献")
    refs = data.get("refs", {})
    lines.append(f"- 条目数：{refs.get('count', '?')}")
    ref_findings = [f for f in all_findings if f.get("section") == "refs"]
    if not ref_findings:
        lines.append("- 未发现编号、年份、标点或著录项问题。")
    else:
        grouped: dict[str, list[int]] = defaultdict(list)
        for f in ref_findings:
            grouped[f.get("message", f.get("kind", "问题"))].append(f.get("paragraph", 0))
        for msg, paras in grouped.items():
            lines.append(f"- {msg}：{len(paras)} 处（段落 {fmt_para_list(paras)}）")
    lines.append("")

    # Manual review
    lines.append("## 需要人工复核")
    lines.append("- 摘要、第 5 章和结论中的用例数、指标、策略名和结论是否一致。")
    lines.append("- 第 2/3 章提出的功能，第 4 章是否有实现，第 5 章是否有测试。")
    lines.append("- 代码块与正文描述的接口、变量、数据流和返回字段是否对应。")
    lines.append("- 图题描述与截图内容是否一致；同一图片被复用是否确实展示的是同一页面的不同功能。")
    lines.append("- 日期、姓名、导师、企业导师和封面信息是否一致。")
    if any(f.get("kind") == "image_reused" for f in fig_findings):
        lines.append("- 脚本检测到图片复用：请确认是否为同一页面的不同功能展示；如果是，正文要说明展示的是哪一部分。")
    if any(f.get("kind") == "toc_page_mismatch" for f in toc_findings):
        lines.append("- 目录存在页码不一致：在 Word 中更新整个目录域并保存，再重新导出 PDF 核对。")
    lines.append("")

    lines.append("## 原始数据")
    for key in ("inspect", "figures", "tables", "toc", "refs", "front"):
        lines.append(f"- `{key}.json`")
    lines.append("")

    report = "\n".join(lines)
    (out / "audit_report.md").write_text(report, encoding="utf-8")
    (out / "audit_data.json").write_text(json.dumps({"data": data, "logs": logs}, ensure_ascii=False, indent=2), encoding="utf-8")
    print(report)
    print(f"\nReport written: {out / 'audit_report.md'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
