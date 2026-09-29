#!/usr/bin/env python3
"""Extract text + equations (in document order) from a .docx file.

python-docx's `paragraph.text` silently returns "" for Word equations (they
are OMML XML, not `w:t` runs), so this walks each paragraph's XML directly
and inlines any `<m:oMath>` as a linearized "[EQ: ...]" expression in its
position relative to the surrounding text (see lib/omml.py).

Top-level body content is walked by `body_blocks`, which annotate.py and
audit_graded.py share so the ids always agree:
    pN  top-level paragraph N (== python-docx `document.paragraphs[N]`)
    tN  top-level table N (== `document.tables[N]`); nested tables are
        flattened into their cell's text
    sN  body-level content control (w:sdt) N — its text would otherwise be
        invisible, since python-docx skips it

This module is normally used through extract.py (cached, with images
exported); run standalone it just prints the record.

Usage:
    .venv/Scripts/python scripts/extract_docx.py <file.docx>
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import _venv  # noqa: F401  — must precede third-party imports (re-runs under .venv)

import docx
from docx.table import Table
from docx.text.paragraph import Paragraph
from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parent))
from lib import cache  # noqa: E402
from lib.images import export_image  # noqa: E402
from lib.views import SOURCE_TOKEN  # noqa: E402
from lib.ooxml_extras import chart_summary  # noqa: E402
from lib.omml import M_NS, linearize_omath  # noqa: E402

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
IMAGE_REF_NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "v": "urn:schemas-microsoft-com:vml",
}
IMAGE_XPATH = ".//w:drawing | .//w:pict | .//w:object"
CHART_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"


def element_text(el) -> tuple[str, list[str]]:
    """Text of any element in document order, equations inlined as
    [EQ: …]; paragraphs inside it (a content control, a nested table)
    are separated by newlines. Returns (text, equations)."""
    parts: list[str] = []
    equations: list[str] = []

    def walk(node):
        q = etree.QName(node)
        if q.namespace == M_NS and q.localname == "oMath":
            expr = linearize_omath(node)
            equations.append(expr)
            parts.append(f"[EQ: {expr}]")
            return
        if q.namespace == W_NS:
            if q.localname == "t":
                parts.append(node.text or "")
                return
            if q.localname == "tab":
                parts.append("\t")
                return
            if q.localname in ("br", "cr"):
                parts.append("\n")
                return
        for child in node:
            walk(child)
        if q.namespace == W_NS and q.localname in ("p", "tc") and node is not el:
            parts.append("\n")

    walk(el)
    return "".join(parts).strip("\n") if len(el) else "", equations


def has_image(el) -> bool:
    # Explicit namespaces: `el` may be a raw lxml element (a content control),
    # not a python-docx element with its built-in prefix map.
    return bool(etree.ElementBase.xpath(el, IMAGE_XPATH, namespaces={"w": W_NS}))


def body_blocks(document):
    """Yield (id, kind, obj) for top-level body content in order: ("pN", "p",
    Paragraph), ("tN", "t", Table), ("sN", "s", <w:sdt> element)."""
    counts = {"p": 0, "t": 0, "s": 0}
    for child in document.element.body.iterchildren():
        tag = etree.QName(child).localname
        kind = {"p": "p", "tbl": "t", "sdt": "s"}.get(tag)
        if kind is None:
            continue
        obj = Paragraph(child, document) if kind == "p" else Table(child, document) if kind == "t" else child
        yield f"{kind}{counts[kind]}", kind, obj
        counts[kind] += 1


def table_rows(table: Table) -> list[list[str]]:
    """Cell text per row (grid order, as `table.cell(r, c)` addresses it);
    nested tables are flattened into their cell's text."""
    return [[element_text(cell._tc)[0] for cell in row.cells] for row in table.rows]


def export_element_images(el, document, image_dir: Path) -> list[str]:
    """Export the pictures inside `el` (DrawingML blips and legacy VML
    imagedata, e.g. an equation a .doc stored as a picture); return paths."""
    paths = []
    rids = etree.ElementBase.xpath(el, ".//a:blip/@r:embed | .//v:imagedata/@r:id",
                                   namespaces=IMAGE_REF_NS)
    for rid in rids:
        rel = document.part.rels.get(rid)
        if rel is None or rel.is_external:
            continue
        part = rel.target_part
        paths.append(export_image(part.blob, Path(str(part.partname)).name, image_dir))
    return paths


def element_charts(el, document) -> list[str]:
    """Native Word charts inside `el`, each summarized from its chart XML
    (type, title, every series' data) — a chart has no picture to export."""
    rids = etree.ElementBase.xpath(el, ".//c:chart/@r:id",
                                   namespaces={"c": CHART_NS, "r": IMAGE_REF_NS["r"]})
    out = []
    for rid in rids:
        rel = document.part.rels.get(rid)
        if rel is not None and not rel.is_external:
            out.append(chart_summary(rel.target_part.blob))
    return out


def count_images(document) -> int:
    """Embedded images via the document's part relationships — inline and
    floating alike, unlike `document.inline_shapes`."""
    return sum(1 for rel in document.part.rels.values() if "image" in rel.reltype)


def extract_docx(path: Path, image_dir: Path | None = None) -> dict:
    """`image_dir`: when given, pictures in paragraphs and content controls
    are exported there and listed in the block's `images` (lib/images.py)."""
    document = docx.Document(str(path))
    blocks = []
    equations_found = 0
    for bid, kind, obj in body_blocks(document):
        if kind == "t":
            blocks.append({"type": "table", "id": bid, "rows": table_rows(obj)})
            continue
        el = obj._p if kind == "p" else obj
        text, eqs = element_text(el)
        equations_found += len(eqs)
        block = {"type": "paragraph" if kind == "p" else "content_control", "id": bid, "text": text}
        if kind == "p" and obj.style is not None and obj.style.name != "Normal":
            block["style"] = obj.style.name
        if eqs:
            block["equations"] = eqs
        charts = element_charts(el, document)
        if charts:
            block["charts"] = charts
        drawings = len(etree.ElementBase.xpath(el, IMAGE_XPATH, namespaces={"w": W_NS}))
        if drawings > len(charts):  # a picture/object beyond the charts' own frames
            block["has_image"] = True
            if image_dir is not None:
                exported = export_element_images(el, document, image_dir)
                if exported:
                    block["images"] = exported
        blocks.append(block)

    image_count = count_images(document)
    # A conversion / PDF base (content-neutral name under .cache/converted) is
    # named literally; a student's own file by token, filled in per view.
    in_cache = Path(path).resolve().is_relative_to(cache.CONVERTED_DIR.resolve())
    unzip_source = str(path) if in_cache else SOURCE_TOKEN
    inspect_dir = f".cache/inspect/{cache.sha256_of(path)[:12]}"
    counts = {t: sum(1 for b in blocks if b["type"] == t)
              for t in ("paragraph", "table", "content_control")}
    return {
        "type": "docx",
        "blocks": blocks,
        "summary": {
            "paragraph_count": counts["paragraph"],
            "table_count": counts["table"],
            "content_controls": counts["content_control"],
            "equations_found": equations_found,
            "chart_count": sum(len(b.get("charts", [])) for b in blocks),
            "non_empty_paragraphs": sum(1 for b in blocks
                                        if b["type"] == "paragraph" and b["text"].strip()),
            "image_count": image_count,
            "image_warning": (
                f"{image_count} embedded image(s) — their content is NOT in the text (no OCR). "
                "Pictures in a paragraph are shown with their exported file in the view's "
                "`(image: …)` tag — Read it there. For an image anywhere else (inside a table, "
                f"header or text box): `unzip -o \"{unzip_source}\" -d \"{inspect_dir}\"` "
                f"and view \"{inspect_dir}/word/media/\"*."
            ) if image_count else None,
        },
    }


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) != 2 or not Path(sys.argv[1]).exists():
        print(json.dumps({"error": "usage: extract_docx.py <existing file.docx>"}))
        sys.exit(1)
    print(json.dumps(extract_docx(Path(sys.argv[1])), indent=2, ensure_ascii=False))


if __name__ == "__main__":
    main()
