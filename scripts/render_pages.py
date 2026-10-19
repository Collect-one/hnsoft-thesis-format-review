#!/usr/bin/env python3
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path


def parse_pages(spec: str | None, count: int) -> list[int]:
    if not spec:
        return list(range(1, count + 1))
    pages: list[int] = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            pages.extend(range(int(a), int(b) + 1))
        else:
            pages.append(int(part))
    return sorted({p for p in pages if 1 <= p <= count})


def main() -> int:
    ap = argparse.ArgumentParser(description="Render PDF pages to PNG for visual QA.")
    ap.add_argument("--pdf", help="input PDF")
    ap.add_argument("--docx", help="input DOCX; a sibling PDF or exported PDF will be used")
    ap.add_argument("--out", required=True, help="output directory")
    ap.add_argument("--dpi", type=int, default=120)
    ap.add_argument("--pages", help="page spec, e.g. 1-3,7")
    args = ap.parse_args()

    out = Path(args.out).resolve()
    out.mkdir(parents=True, exist_ok=True)
    pdf = Path(args.pdf).resolve() if args.pdf else None
    if pdf is None and args.docx:
        docx = Path(args.docx).resolve()
        sibling = docx.with_suffix(".pdf")
        if sibling.exists():
            pdf = sibling
        else:
            exported = out / (docx.stem + ".pdf")
            ps1 = Path(__file__).resolve().parent / "export_pdf.ps1"
            if sys.platform.startswith("win") and ps1.exists():
                cmd = ["powershell", "-ExecutionPolicy", "Bypass", "-File", str(ps1), "-Docx", str(docx), "-Pdf", str(exported)]
                subprocess.run(cmd, check=True)
                pdf = exported
    if pdf is None or not pdf.exists():
        raise SystemExit("PDF not found. Pass --pdf, or export the DOCX to PDF first.")

    try:
        import pymupdf as fitz
    except ImportError:
        try:
            import fitz  # type: ignore
        except ImportError:
            fitz = None

    if fitz is not None:
        doc = fitz.open(str(pdf))
        pages = parse_pages(args.pages, doc.page_count)
        for n in pages:
            pix = doc[n - 1].get_pixmap(dpi=args.dpi)
            dest = out / f"page-{n:03d}.png"
            pix.save(str(dest))
        print(f"Rendered {len(pages)} page(s) from {pdf} to {out}")
        return 0

    pdftoppm = shutil.which("pdftoppm")
    if pdftoppm:
        prefix = out / "page"
        cmd = [pdftoppm, "-png", "-r", str(args.dpi), str(pdf), str(prefix)]
        subprocess.run(cmd, check=True)
        print(f"Rendered pages with pdftoppm to {out}")
        return 0

    raise SystemExit("PyMuPDF and pdftoppm are both unavailable. Install pymupdf or poppler.")


if __name__ == "__main__":
    raise SystemExit(main())
