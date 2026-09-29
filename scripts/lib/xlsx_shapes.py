"""Put a workbook's shapes (text boxes, arrows, callouts) back after openpyxl.

openpyxl keeps cells, formulas, styles, images and charts when it saves, but
silently drops every drawing *shape* — so a student's answer typed into a
text box would vanish from the graded copy. After annotate.py saves, this
copies each shape anchor from the original package's drawing parts into the
graded package, sheet by sheet (creating the sheet's drawing part and
relationship when openpyxl wrote none).

Only self-contained shapes are copied: an anchor whose XML references other
package parts (r:embed / r:id — pictures, charts) is skipped, since openpyxl
already carried those over itself.
"""
from __future__ import annotations

import posixpath
import zipfile
from pathlib import Path

from lxml import etree

from lib.ooxml_extras import NS, drawing_part, sheet_parts

DRAWING_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
DRAWING_CT = "application/vnd.openxmlformats-officedocument.drawing+xml"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
SHAPE_TAGS = {"sp", "grpSp", "cxnSp"}
# Worksheet children that must come after <drawing> (CT_Worksheet order).
AFTER_DRAWING = {"legacyDrawing", "legacyDrawingHF", "drawingHF", "picture", "oleObjects",
                 "controls", "webPublishItems", "tableParts", "extLst"}


def _shape_anchors(zf: zipfile.ZipFile, drawing: str) -> list:
    root = etree.fromstring(zf.read(drawing))
    out = []
    for anchor in root:
        tags = {etree.QName(el).localname for el in anchor}
        refs = anchor.xpath(".//@r:embed | .//@r:id | .//@r:link", namespaces=NS)
        if tags & SHAPE_TAGS and not refs:
            out.append(anchor)
    return out


def _rels_name(part: str) -> str:
    folder, name = posixpath.split(part)
    return posixpath.join(folder, "_rels", name + ".rels")


def restore_shapes(original: Path, graded: Path) -> int:
    """Copy shapes from `original` into `graded` (in place). Returns how many."""
    with zipfile.ZipFile(original) as zo:
        wanted = {}
        for sheet, part in sheet_parts(zo).items():
            drawing = drawing_part(zo, part)
            if drawing and drawing in zo.namelist():
                anchors = _shape_anchors(zo, drawing)
                if anchors:
                    wanted[sheet] = anchors
    if not wanted:
        return 0

    with zipfile.ZipFile(graded) as zg:
        files = {name: zg.read(name) for name in zg.namelist()}
        graded_sheets = sheet_parts(zg)
        graded_drawings = {s: drawing_part(zg, p) for s, p in graded_sheets.items()}

    copied = 0
    for sheet, anchors in wanted.items():
        part = graded_sheets.get(sheet)
        if part is None:
            continue
        drawing = graded_drawings.get(sheet)
        if drawing is None or drawing not in files:
            drawing = _new_drawing(files, part)
        root = etree.fromstring(files[drawing])
        for anchor in anchors:
            root.append(anchor)
            copied += 1
        files[drawing] = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    tmp = graded.with_name(graded.name + ".shapes.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zw:
        for name, data in files.items():
            zw.writestr(name, data)
    tmp.replace(graded)
    return copied


def _new_drawing(files: dict, sheet_part: str) -> str:
    """Create an empty drawing part for `sheet_part` and wire it up."""
    n = 1
    while f"xl/drawings/drawing{n}.xml" in files:
        n += 1
    drawing = f"xl/drawings/drawing{n}.xml"
    files[drawing] = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                      b'<xdr:wsDr xmlns:xdr="' + NS["xdr"].encode() + b'" xmlns:a="'
                      + NS["a"].encode() + b'"/>')

    # content type
    ct = etree.fromstring(files["[Content_Types].xml"])
    override = etree.SubElement(ct, f"{{{CT_NS}}}Override")
    override.set("PartName", "/" + drawing)
    override.set("ContentType", DRAWING_CT)
    files["[Content_Types].xml"] = etree.tostring(ct, xml_declaration=True, encoding="UTF-8",
                                                  standalone=True)

    # sheet -> drawing relationship
    rels_name = _rels_name(sheet_part)
    if rels_name in files:
        rels = etree.fromstring(files[rels_name])
    else:
        rels = etree.Element(f"{{{NS['pr']}}}Relationships", nsmap={None: NS["pr"]})
    used = {r.get("Id") for r in rels}
    rid = next(f"rId{i}" for i in range(1, 10_000) if f"rId{i}" not in used)
    rel = etree.SubElement(rels, f"{{{NS['pr']}}}Relationship")
    rel.set("Id", rid)
    rel.set("Type", DRAWING_REL)
    rel.set("Target", posixpath.relpath(drawing, posixpath.dirname(sheet_part)))
    files[rels_name] = etree.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)

    # <drawing r:id=…/> in the sheet, at its schema position
    sheet = etree.fromstring(files[sheet_part])
    el = etree.Element(f"{{{NS['m']}}}drawing", nsmap={"r": NS["r"]})
    el.set(f"{{{NS['r']}}}id", rid)
    later = next((i for i, child in enumerate(sheet)
                  if etree.QName(child).localname in AFTER_DRAWING), None)
    if later is None:
        sheet.append(el)
    else:
        sheet.insert(later, el)
    files[sheet_part] = etree.tostring(sheet, xml_declaration=True, encoding="UTF-8", standalone=True)
    return drawing
