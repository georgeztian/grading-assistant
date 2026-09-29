"""Compact, line-oriented rendering of extraction records.

The cached JSON records are complete but verbose (indented, and every
paragraph repeats fields like `"equations": []`). Agents read this view
instead — one line per paragraph/table row/cell, empty paragraphs omitted,
default fields dropped — so each agent carries far fewer tokens in context on
every turn. The JSON record stays the source of truth for scripts.

Every addressable unit has a stable id, used identically by the view, the
rubric (rubric.py) and annotation anchors (annotate.py):

- documents (.docx/.doc/.pdf): `p<N>` = top-level paragraph N, `t<N>` =
  table N, `s<N>` = body-level content control N (see extract_docx.body_blocks);
  for a PDF these index the base .docx rebuilt from it (lib/pdf_docx.py)
- workbooks (.xlsx/.xls): `<Sheet>!<Cell>`, e.g. `WACC!B5`
"""
from __future__ import annotations

import json
from pathlib import Path

from lib.cache import rel_path

# Stands for "the file this view is for" inside shared record text (e.g. an
# unzip command) — the record is shared by identical files, so it can't hold
# any one student's path; render_view substitutes the requesting file's.
SOURCE_TOKEN = "{this file}"


def split_sheet_ref(spec: str) -> tuple[str, str]:
    """'Sheet!B5' -> ('Sheet', 'B5'). The view writes sheet names unquoted
    (`Ex TN5 WACC!G6`); Excel's quoted form (`'O''Brien'!B5`) is accepted
    too. The last "!" separates, so a sheet name may itself contain one."""
    sheet, ref = spec.rsplit("!", 1)
    if len(sheet) >= 2 and sheet[0] == sheet[-1] == "'":  # Excel never allows a name to start/end with '
        sheet = sheet[1:-1].replace("''", "'")
    return sheet, ref


def is_sheet_record(record: dict) -> bool:
    return record["type"] in ("xlsx", "xls")


def doc_blocks(record: dict) -> list[dict]:
    if record["type"] == "pdf":
        return record.get("annotation_blocks", [])
    return record["blocks"]


def block_nonempty(block: dict) -> bool:
    if block["type"] == "table":
        return any(c.strip() for row in block["rows"] for c in row)
    return bool(block["text"].strip()) or bool(block.get("has_image")) or bool(block.get("charts"))


SHEET_OBJECTS = (("text_boxes", "textbox"), ("charts", "chart"))


def cell_nonempty(cell: dict) -> bool:
    value = cell.get("value")
    return (value is not None and value != "") or "formula" in cell or bool(cell.get("comment"))


def units(record: dict) -> list[dict]:
    """Every addressable unit in document/workbook order."""
    out = []
    if is_sheet_record(record):
        for sheet in record["sheet_order"]:
            info = record["sheets"][sheet]
            for coord, cell in info["cells"].items():
                out.append({"id": f"{sheet}!{coord}", "sheet": sheet, "coord": coord,
                            "cell": cell, "nonempty": cell_nonempty(cell)})
            # Text boxes / charts: units of their own (id Sheet!textbox1, Sheet!chart1).
            for key, prefix in SHEET_OBJECTS:
                for n, obj in enumerate(info.get(key, []), 1):
                    out.append({"id": f"{sheet}!{prefix}{n}", "sheet": sheet, "coord": None,
                                "object": (f"{prefix}{n}", obj), "nonempty": True})
    else:
        for block in doc_blocks(record):
            out.append({"id": block["id"], "block": block,
                        "page": block.get("page"), "nonempty": block_nonempty(block)})
    return out


def fmt_value(value) -> str:
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float):
        return format(value, ".12g")
    if isinstance(value, int):
        return str(value)
    return json.dumps(value, ensure_ascii=False, default=str)


def render_block(block: dict, indent: str = "") -> list[str]:
    bid = block["id"]
    if block["type"] != "table":
        tags = []
        if block["type"] == "content_control":
            tags.append("content control")
        if block.get("style"):
            tags.append(block["style"])
        if block.get("images"):
            tags.append("image: " + ", ".join(block["images"]))
        elif block.get("has_image"):
            tags.append("image")
        for chart in block.get("charts", []):
            tags.append(f"chart: {chart}")
        head = f"{indent}[{bid}]" + (f" ({', '.join(tags)})" if tags else "")
        text_lines = block["text"].split("\n")
        out = [f"{head} {text_lines[0]}".rstrip()]
        out += [f"{indent}    {t}" for t in text_lines[1:]]
        return out
    rows = block["rows"]
    ncols = max((len(r) for r in rows), default=0)
    out = [f"{indent}[{bid}] (table {len(rows)}x{ncols})"]
    for r, row in enumerate(rows):
        if any(c.strip() for c in row):
            out.append(f"{indent}    r{r}: " + " | ".join(c.replace("\n", " / ") for c in row))
    return out


def render_object(name: str, obj: dict, indent: str = "") -> str:
    body = obj.get("summary") or json.dumps(obj.get("text", ""), ensure_ascii=False)
    return f"{indent}{name} (at {obj['anchor']}): {body}"


def render_unit(u: dict, indent: str = "") -> list[str]:
    if "object" in u:
        return [render_object(*u["object"], indent)]
    if "cell" in u:
        return [render_cell(u["coord"], u["cell"], indent)]
    return render_block(u["block"], indent)


def render_cell(coord: str, cell: dict, indent: str = "") -> str:
    if "formula" in cell:
        value = fmt_value(cell["value"]) if "value" in cell else "(no cached value)"
        line = f"{indent}{coord}: {cell['formula']} -> {value}"
    elif "value" in cell:
        line = f"{indent}{coord}: {fmt_value(cell['value'])}"
    else:
        line = f"{indent}{coord}: (empty)"
    if cell.get("comment"):
        line += f"  [comment: {json.dumps(cell['comment'], ensure_ascii=False)}]"
    return line


def sheet_header(record: dict, sheet: str) -> str:
    info = record["sheets"][sheet]
    extras = []
    if info.get("hidden"):
        extras.append("HIDDEN")
    if info.get("chartsheet"):
        extras.append("chart sheet: no cells — annotate its data cells on another tab")
    for key in ("hidden_rows", "hidden_columns", "merged_ranges"):
        if info.get(key):
            extras.append(f"{key}={info[key]}")
    for im in info.get("images", []):
        extras.append(f"image at {im['anchor']}: {im['path']}")
    return f"## Sheet {json.dumps(sheet, ensure_ascii=False)}" + (f"  ({'; '.join(extras)})" if extras else "")


def page_header(record: dict, page_index: int) -> str:
    page = record["pages"][page_index]
    line = f"=== page {page_index + 1} ==="
    if page.get("flagged_reasons"):
        line += (f"  FLAGGED {','.join(page['flagged_reasons'])} — read the page render, not "
                 f"the text: {page.get('render_path')}")
    return line


def warnings(record: dict) -> list[str]:
    out = []
    summary = record.get("summary") or {}
    for key in ("image_warning", "note"):
        if summary.get(key):
            out.append(summary[key])
    out += record.get("warnings", [])
    return out


def render_units(record: dict, selected: list[dict], indent: str = "") -> list[str]:
    """Render a subset of units (in the given order) verbatim, with sheet/page
    context lines and an ellipsis wherever non-empty content is skipped."""
    all_nonempty = [u["id"] for u in units(record) if u["nonempty"]]
    rank = {uid: i for i, uid in enumerate(all_nonempty)}
    out: list[str] = []
    prev_rank = None
    current_group = object()
    for u in selected:
        if not u["nonempty"]:
            continue
        group = u.get("sheet") if is_sheet_record(record) else u.get("page")
        if group != current_group:
            if is_sheet_record(record):
                out.append(f"{indent}[Sheet {json.dumps(group, ensure_ascii=False)}]")
            elif group is not None and record["type"] == "pdf":
                out.append(f"{indent}[page {group + 1}]")
            current_group = group
            prev_rank = None
        r = rank[u["id"]]
        if prev_rank is not None and r != prev_rank + 1:
            out.append(f"{indent}  ⋯")
        prev_rank = r
        out += render_unit(u, indent + "  ")
    return out


def render_view(record: dict, label: str) -> str:
    kind = record["type"]
    if record.get("converted_from"):
        kind += f" (converted from {Path(record['converted_from']).suffix.lower()})"
    lines = [f"# Extraction view: {label}  [{kind}]"]
    if is_sheet_record(record):
        lines.append("# ids: <Sheet>!<Cell>. Cells shown as `A1: value` or `A1: =formula -> cached value`. "
                     "Blank cells omitted. Text boxes / charts are listed as textboxN / chartN "
                     "(at their anchor cell) — an answer in a text box is anchored at that cell.")
    else:
        lines.append("# ids: [pN] = paragraph N, [tN] = table N (rows r0, r1, …), [sN] = content control N. "
                     "Empty paragraphs omitted. Use these ids as annotation anchors.")
        if kind == "pdf":
            lines.append("# PDF: paragraphs are from the base .docx rebuilt from this PDF "
                         f"({rel_path(record['base_docx_path'])}); the graded copy is written from it.")
    summary = record.get("summary") or {}
    scalars = {k: v for k, v in summary.items()
               if k not in ("image_warning", "note") and v not in (None, "", [], {})}
    if scalars:
        lines.append("# summary: " + ", ".join(f"{k}={v}" for k, v in scalars.items()))
    for w in warnings(record):
        lines.append(f"! {w.replace(SOURCE_TOKEN, label)}")
    lines.append("")

    if is_sheet_record(record):
        by_sheet: dict[str, list[dict]] = {}
        for u in units(record):
            by_sheet.setdefault(u["sheet"], []).append(u)
        for sheet in record["sheet_order"]:
            lines.append(sheet_header(record, sheet))
            for u in by_sheet.get(sheet, []):
                if u["nonempty"]:
                    lines += render_unit(u)
            lines.append("")
    elif record["type"] == "pdf":
        by_page: dict[int, list[dict]] = {}
        for block in doc_blocks(record):
            by_page.setdefault(block.get("page", 0), []).append(block)
        for i in range(len(record["pages"])):
            lines.append(page_header(record, i))
            for block in by_page.get(i, []):
                if block_nonempty(block):
                    lines += render_block(block)
            lines.append("")
    else:
        for block in doc_blocks(record):
            if block_nonempty(block):
                lines += render_block(block)
    return "\n".join(lines).rstrip() + "\n"
