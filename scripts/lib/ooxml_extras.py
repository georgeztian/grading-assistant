"""Content that python-docx / openpyxl don't expose: charts and text boxes.

- A native Office chart is XML (title, series names, categories, values),
  not a picture, so it has no image to export. `chart_summary` turns that XML
  into one readable line — the chart's actual data — so an agent can grade a
  chart answer without seeing it rendered.
- Spreadsheet text boxes / shapes live in drawing parts openpyxl ignores
  entirely (it drops them on save, too). `xlsx_drawing_objects` reads them
  straight from the package, per sheet, with the cell each is anchored at.
  lib/xlsx_shapes.py puts them back into a graded copy.
"""
from __future__ import annotations

import posixpath
import zipfile

from lxml import etree

NS = {
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}
MAX_POINTS = 40


def _texts(el, path: str) -> list[str]:
    return [t for t in el.xpath(path, namespaces=NS) if t is not None]


def chart_summary(xml: bytes) -> str:
    """One line: chart type, title, and every series' category=value pairs."""
    root = etree.fromstring(xml)
    title = "".join(_texts(root, "c:chart/c:title//a:t/text()")).strip()
    plot = root.find("c:chart/c:plotArea", NS)
    kinds, series = [], []
    if plot is not None:
        for child in plot:
            name = etree.QName(child).localname
            if not name.endswith("Chart"):
                continue
            kinds.append(name[:-5])
            for ser in child.findall("c:ser", NS):
                label = " ".join(_texts(ser, "c:tx//c:v/text()")) or "series"
                cats = _texts(ser, "(c:cat|c:xVal)//c:pt/c:v/text()")
                vals = _texts(ser, "(c:val|c:yVal)//c:pt/c:v/text()")
                pairs = [f"{c}={v}" for c, v in zip(cats, vals)] if cats else vals
                more = f", … (+{len(pairs) - MAX_POINTS})" if len(pairs) > MAX_POINTS else ""
                series.append(f"{label}: {', '.join(pairs[:MAX_POINTS])}{more}")
    head = f"{'/'.join(kinds) or 'chart'} chart"
    if title:
        head += f" {title!r}"
    return head + (" — " + "; ".join(series) if series else " (no cached data)")


def _rels(zf: zipfile.ZipFile, part: str) -> dict[str, tuple[str, str]]:
    """rId -> (type, absolute target part) for `part`."""
    folder, name = posixpath.split(part)
    rels_name = posixpath.join(folder, "_rels", name + ".rels")
    if rels_name not in zf.namelist():
        return {}
    out = {}
    for rel in etree.fromstring(zf.read(rels_name)).findall("pr:Relationship", NS):
        if rel.get("TargetMode") == "External":
            continue
        target = posixpath.normpath(posixpath.join(folder, rel.get("Target")))
        out[rel.get("Id")] = (rel.get("Type"), target.lstrip("/"))
    return out


def sheet_parts(zf: zipfile.ZipFile) -> dict[str, str]:
    """Sheet name -> worksheet part name."""
    wb = etree.fromstring(zf.read("xl/workbook.xml"))
    rels = _rels(zf, "xl/workbook.xml")
    out = {}
    for sheet in wb.findall("m:sheets/m:sheet", NS):
        rid = sheet.get(f"{{{NS['r']}}}id")
        if rid in rels:
            out[sheet.get("name")] = rels[rid][1]
    return out


def drawing_part(zf: zipfile.ZipFile, sheet_part: str) -> str | None:
    for rtype, target in _rels(zf, sheet_part).values():
        if rtype.endswith("/drawing"):
            return target
    return None


def anchor_cell(anchor) -> str:
    from openpyxl.utils.cell import get_column_letter
    start = anchor.find("xdr:from", NS)
    if start is None:
        return "absolute"
    col = int(start.findtext("xdr:col", "0", NS))
    row = int(start.findtext("xdr:row", "0", NS))
    return f"{get_column_letter(col + 1)}{row + 1}"


def shape_text(sp) -> str:
    paras = []
    for p in sp.iterfind(".//xdr:txBody/a:p", NS):
        paras.append("".join(p.xpath(".//a:t/text()", namespaces=NS)))
    return "\n".join(paras).strip()


def xlsx_drawing_objects(path) -> dict[str, dict[str, list[dict]]]:
    """{sheet: {"text_boxes": [{anchor, text}], "charts": [{anchor, summary}]}}
    for every sheet that has any."""
    out: dict[str, dict[str, list[dict]]] = {}
    with zipfile.ZipFile(path) as zf:
        for sheet, part in sheet_parts(zf).items():
            drawing = drawing_part(zf, part)
            if drawing is None or drawing not in zf.namelist():
                continue
            rels = _rels(zf, drawing)
            root = etree.fromstring(zf.read(drawing))
            found = {"text_boxes": [], "charts": []}
            for anchor in root:
                cell = anchor_cell(anchor)
                for sp in anchor.iterfind(".//xdr:sp", NS):
                    text = shape_text(sp)
                    if text:
                        found["text_boxes"].append({"anchor": cell, "text": text})
                for chart in anchor.iterfind(".//c:chart", NS):
                    rid = chart.get(f"{{{NS['r']}}}id")
                    if rid in rels and rels[rid][1] in zf.namelist():
                        found["charts"].append({"anchor": cell,
                                                "summary": chart_summary(zf.read(rels[rid][1]))})
            if found["text_boxes"] or found["charts"]:
                out[sheet] = found
    return out
