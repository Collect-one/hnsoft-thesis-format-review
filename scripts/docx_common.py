from __future__ import annotations

import copy
import re
import zipfile
from dataclasses import dataclass, field
from typing import Any, Iterator
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
V = "{urn:schemas-microsoft-com:vml}"


def lname(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _val(el: ET.Element | None, default: str | None = None) -> str | None:
    if el is None:
        return default
    return el.get(W + "val", default)


def _parse_rpr(rpr: ET.Element | None) -> dict[str, Any]:
    out: dict[str, Any] = {}
    if rpr is None:
        return out
    rf = rpr.find(W + "rFonts")
    if rf is not None:
        out["fonts"] = {lname(k): v for k, v in rf.attrib.items()}
    for tag in ("sz", "szCs", "b", "bCs", "color", "kern", "i", "u"):
        el = rpr.find(W + tag)
        if el is not None:
            out[tag] = _val(el, "1")
    return out


def _parse_ppr(ppr: ET.Element | None) -> dict[str, Any]:
    out: dict[str, Any] = {}
    if ppr is None:
        return out
    sp = ppr.find(W + "spacing")
    if sp is not None:
        out["spacing"] = {lname(k): v for k, v in sp.attrib.items()}
    ind = ppr.find(W + "ind")
    if ind is not None:
        out["ind"] = {lname(k): v for k, v in ind.attrib.items()}
    jc = ppr.find(W + "jc")
    if jc is not None:
        out["jc"] = _val(jc)
    outline = ppr.find(W + "outlineLvl")
    if outline is not None:
        out["outlineLvl"] = _val(outline)
    for tag in ("keepNext", "keepLines", "pageBreakBefore"):
        el = ppr.find(W + tag)
        if el is not None:
            out[tag] = _val(el, "1")
    return out


def _merge(dst: dict[str, Any], src: dict[str, Any]) -> None:
    for key, value in src.items():
        if isinstance(value, dict) and isinstance(dst.get(key), dict):
            dst[key].update(copy.deepcopy(value))
        else:
            dst[key] = copy.deepcopy(value)


@dataclass
class StyleInfo:
    style_id: str
    name: str = ""
    based_on: str | None = None
    rpr: dict[str, Any] = field(default_factory=dict)
    ppr: dict[str, Any] = field(default_factory=dict)


@dataclass
class RunInfo:
    text: str
    rpr: dict[str, Any]


@dataclass
class ParagraphInfo:
    index: int
    block_index: int
    text: str
    style_id: str
    style_name: str
    ppr: dict[str, Any]
    runs: list[RunInfo]
    images: list[str]
    inferred_page: int
    element: ET.Element = field(repr=False, default=None)


@dataclass
class TableRowInfo:
    cells: list[list[str]]
    cell_texts: list[str]


@dataclass
class TableInfo:
    index: int
    block_index: int
    caption: str
    rows: list[TableRowInfo]
    grid_width_twips: int
    tbl_width: str
    borders: dict[str, dict[str, str]]
    header_repeat: bool
    inferred_page: int
    element: ET.Element = field(repr=False, default=None)


class DocxModel:
    def __init__(self, path: str):
        self.path = path
        self.z = zipfile.ZipFile(path)
        self.document = ET.fromstring(self.z.read("word/document.xml"))
        self.body = self.document.find(W + "body")
        self.styles_root = ET.fromstring(self.z.read("word/styles.xml"))
        self.rels_root = ET.fromstring(self.z.read("word/_rels/document.xml.rels"))
        self.settings_root = ET.fromstring(self.z.read("word/settings.xml"))
        self.rels: dict[str, str] = {
            rel.get("Id", ""): rel.get("Target", "") for rel in self.rels_root
        }
        self.styles: dict[str, StyleInfo] = {}
        self.default_paragraph_style_id: str = ""
        self.default_rpr: dict[str, Any] = {}
        self.default_ppr: dict[str, Any] = {}
        self._load_styles()
        self.paragraphs: list[ParagraphInfo] = []
        self.tables: list[TableInfo] = []
        self.block_pages: dict[int, int] = {}
        self._load_body()

    def _load_styles(self) -> None:
        dd = self.styles_root.find(W + "docDefaults")
        if dd is not None:
            self.default_rpr = _parse_rpr(dd.find(".//" + W + "rPr"))
            self.default_ppr = _parse_ppr(dd.find(".//" + W + "pPr"))
        for st in self.styles_root.findall(W + "style"):
            sid = st.get(W + "styleId", "")
            if st.get(W + "type") == "paragraph" and st.get(W + "default") == "1":
                self.default_paragraph_style_id = sid
            name_el = st.find(W + "name")
            based = st.find(W + "basedOn")
            self.styles[sid] = StyleInfo(
                style_id=sid,
                name=_val(name_el, "") or "",
                based_on=_val(based) if based is not None else None,
                rpr=_parse_rpr(st.find(W + "rPr")),
                ppr=_parse_ppr(st.find(W + "pPr")),
            )

    def _style_chain(self, style_id: str) -> list[StyleInfo]:
        chain: list[StyleInfo] = []
        seen: set[str] = set()
        sid = style_id
        while sid and sid in self.styles and sid not in seen:
            seen.add(sid)
            st = self.styles[sid]
            chain.append(st)
            sid = st.based_on or ""
        return list(reversed(chain))

    def resolved_style_rpr(self, style_id: str) -> dict[str, Any]:
        out = copy.deepcopy(self.default_rpr)
        for st in self._style_chain(style_id):
            _merge(out, st.rpr)
        return out

    def resolved_style_ppr(self, style_id: str) -> dict[str, Any]:
        out = copy.deepcopy(self.default_ppr)
        for st in self._style_chain(style_id):
            _merge(out, st.ppr)
        return out

    @staticmethod
    def paragraph_text(p: ET.Element) -> str:
        parts: list[str] = []
        for node in p.iter():
            if node.tag == W + "t":
                parts.append(node.text or "")
            elif node.tag == W + "tab":
                parts.append("\t")
            elif node.tag in (W + "br", W + "cr"):
                parts.append("\n")
            elif node.tag == W + "drawing":
                parts.append("[[IMAGE]]")
            elif node.tag == W + "pict":
                parts.append("[[IMAGE]]")
        return "".join(parts)

    @staticmethod
    def paragraph_images(p: ET.Element) -> list[str]:
        rids: list[str] = []
        for node in p.iter():
            if node.tag == V + "imagedata":
                rid = node.get(R + "id")
                if rid:
                    rids.append(rid)
            elif lname(node.tag) == "blip":
                rid = node.get(R + "embed")
                if rid:
                    rids.append(rid)
        return list(dict.fromkeys(rids))

    def paragraph_style_id(self, p: ET.Element) -> str:
        ppr = p.find(W + "pPr")
        if ppr is None:
            return ""
        st = ppr.find(W + "pStyle")
        return _val(st, "") or self.default_paragraph_style_id

    def paragraph_style_name(self, style_id: str) -> str:
        return self.styles.get(style_id, StyleInfo(style_id)).name

    def paragraph_ppr(self, p: ET.Element) -> dict[str, Any]:
        sid = self.paragraph_style_id(p)
        out = self.resolved_style_ppr(sid)
        _merge(out, _parse_ppr(p.find(W + "pPr")))
        return out

    def run_rpr(self, p: ET.Element, r: ET.Element) -> dict[str, Any]:
        sid = self.paragraph_style_id(p)
        out = self.resolved_style_rpr(sid)
        _merge(out, _parse_rpr(r.find(W + "rPr")))
        return out

    def _load_body(self) -> None:
        page = 1
        p_index = 0
        t_index = 0
        for block_index, child in enumerate(list(self.body), start=1):
            self.block_pages[block_index] = page
            if child.tag == W + "p":
                p_index += 1
                sid = self.paragraph_style_id(child)
                runs: list[RunInfo] = []
                for r in child.findall(W + "r"):
                    text = "".join(t.text or "" for t in r.iter(W + "t"))
                    if not text:
                        continue
                    runs.append(RunInfo(text=text, rpr=self.run_rpr(child, r)))
                marker_count = sum(1 for e in child.iter() if e.tag == W + "lastRenderedPageBreak")
                inferred = page + (1 if marker_count else 0)
                self.paragraphs.append(
                    ParagraphInfo(
                        index=p_index,
                        block_index=block_index,
                        text=self.paragraph_text(child),
                        style_id=sid,
                        style_name=self.paragraph_style_name(sid),
                        ppr=self.paragraph_ppr(child),
                        runs=runs,
                        images=self.paragraph_images(child),
                        inferred_page=inferred,
                        element=child,
                    )
                )
                page += marker_count
            elif child.tag == W + "tbl":
                t_index += 1
                rows: list[TableRowInfo] = []
                for tr in child.findall(W + "tr"):
                    cells: list[list[str]] = []
                    cell_texts: list[str] = []
                    for tc in tr.findall(W + "tc"):
                        paras = [
                            self.paragraph_text(p).strip()
                            for p in tc.findall(W + "p")
                            if self.paragraph_text(p).strip()
                        ]
                        cells.append(paras)
                        cell_texts.append(" / ".join(paras))
                    rows.append(TableRowInfo(cells=cells, cell_texts=cell_texts))
                grid = child.find(W + "tblGrid")
                cols = grid.findall(W + "gridCol") if grid is not None else []
                width = sum(int(c.get(W + "w", "0") or 0) for c in cols)
                tblpr = child.find(W + "tblPr")
                tblw_el = tblpr.find(W + "tblW") if tblpr is not None else None
                tblw = f"{tblw_el.get(W + 'w', '?')}/{tblw_el.get(W + 'type', '?')}" if tblw_el is not None else "?"
                borders: dict[str, dict[str, str]] = {}
                if tblpr is not None:
                    tb = tblpr.find(W + "tblBorders")
                    if tb is not None:
                        for b in tb:
                            borders[lname(b.tag)] = {lname(k): v for k, v in b.attrib.items()}
                header_repeat = False
                if child.findall(W + "tr"):
                    trpr = child.findall(W + "tr")[0].find(W + "trPr")
                    header_repeat = trpr is not None and trpr.find(W + "tblHeader") is not None
                caption = ""
                for prev in reversed(self.paragraphs):
                    if prev.block_index >= block_index:
                        continue
                    t = prev.text.strip()
                    if re.match(r"^表\s*\d+\.\d+", t):
                        caption = t
                        break
                    if prev.block_index < block_index - 8:
                        break
                marker_count = sum(1 for e in child.iter() if e.tag == W + "lastRenderedPageBreak")
                self.tables.append(
                    TableInfo(
                        index=t_index,
                        block_index=block_index,
                        caption=caption,
                        rows=rows,
                        grid_width_twips=width,
                        tbl_width=tblw,
                        borders=borders,
                        header_repeat=header_repeat,
                        inferred_page=page,
                        element=child,
                    )
                )
                page += marker_count

    @property
    def has_update_fields(self) -> bool:
        return self.settings_root.find(W + "updateFields") is not None

    def rel_target(self, rid: str) -> str:
        return self.rels.get(rid, "")

    def headings(self) -> list[ParagraphInfo]:
        out = []
        for p in self.paragraphs:
            sid = p.style_id
            outline = str(p.ppr.get("outlineLvl", ""))
            if sid in {"1", "2", "3"} or outline in {"0", "1", "2"}:
                out.append(p)
        return out

    def toc_entries(self) -> list[dict[str, str]]:
        entries: list[dict[str, str]] = []
        for p in self.paragraphs:
            if not p.style_id.startswith("TOC"):
                continue
            text = p.text.strip()
            if not text:
                continue
            m = re.match(r"^(.*?)\t\s*([ivxlcdmIVXLCDM]+|\d+)\s*$", text)
            if m:
                entries.append({"title": m.group(1).strip(), "page": m.group(2).strip(), "style": p.style_id})
        return entries

    def find_paragraphs(self, pattern: str) -> list[ParagraphInfo]:
        rx = re.compile(pattern)
        return [p for p in self.paragraphs if rx.search(p.text)]



def find_body_start(model: "DocxModel") -> ParagraphInfo | None:
    for p in model.paragraphs:
        if p.style_id.startswith("TOC"):
            continue
        t = re.sub(r"\s+", "", p.text or "")
        if t.startswith("第1章") or t.startswith("第一章"):
            return p
    for p in model.headings():
        if not p.style_id.startswith("TOC"):
            return p
    return None
