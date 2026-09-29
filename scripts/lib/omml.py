"""Linearize Word OMML (<m:oMath>) equation XML into readable plain text.

python-docx's `paragraph.text` silently returns "" for equation paragraphs —
it never parses OMML. This walks the math XML directly so equations show up
as readable inline expressions instead of vanishing from the extracted text.
Not full LaTeX — just enough structure (fractions, sub/superscript, roots,
n-ary operators) to compare against a solution's wording.
"""
from __future__ import annotations

from lxml import etree

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def _tag(el) -> str:
    return etree.QName(el).localname


def _m(el, local_name: str):
    return el.find(f"{{{M_NS}}}{local_name}")


def _m_all(el, local_name: str):
    return el.findall(f"{{{M_NS}}}{local_name}")


def _text_of(e_el) -> str:
    """Render an <m:e> (or similar) base element's direct text content."""
    if e_el is None:
        return ""
    return linearize_element(e_el)


def linearize_element(el) -> str:
    """Recursively render one OMML element (and its children) as text."""
    tag = _tag(el)

    if tag == "t":
        return el.text or ""

    if tag == "r":
        return "".join(linearize_element(c) for c in el if _tag(c) != "rPr")

    if tag == "f":
        num = _m(el, "num")
        den = _m(el, "den")
        return f"({_text_of(num)})/({_text_of(den)})"

    if tag in ("sSub", "sSup", "sSubSup"):
        base = _text_of(_m(el, "e"))
        parts = [base]
        sub = _m(el, "sub")
        sup = _m(el, "sup")
        if sub is not None:
            parts.append(f"_{{{_text_of(sub)}}}")
        if sup is not None:
            parts.append(f"^{{{_text_of(sup)}}}")
        return "".join(parts)

    if tag == "rad":
        deg = _m(el, "deg")
        base = _text_of(_m(el, "e"))
        deg_text = _text_of(deg) if deg is not None else ""
        if deg_text.strip():
            return f"root({deg_text}, {base})"
        return f"sqrt({base})"

    if tag == "nary":
        nary_pr = _m(el, "naryPr")
        chr_el = _m(nary_pr, "chr") if nary_pr is not None else None
        # Per the OMML spec an n-ary operator with no m:chr is an integral.
        symbol = chr_el.get(f"{{{M_NS}}}val") if chr_el is not None else "∫"
        sub = _text_of(_m(el, "sub"))
        sup = _text_of(_m(el, "sup"))
        base = _text_of(_m(el, "e"))
        bounds = ""
        if sub or sup:
            bounds = f"_{{{sub}}}^{{{sup}}}"
        return f"{symbol}{bounds}({base})"

    if tag == "d":
        # Delimiter wrapping one or more m:e children. Its characters default
        # to "(", ")" and "|" but may be brackets, bars, or empty.
        d_pr = _m(el, "dPr")

        def char(name: str, default: str) -> str:
            node = _m(d_pr, name) if d_pr is not None else None
            return node.get(f"{{{M_NS}}}val", "") if node is not None else default

        parts = [_text_of(e) for e in _m_all(el, "e")]
        return char("begChr", "(") + char("sepChr", "|").join(parts) + char("endChr", ")")

    if tag == "eqArr":
        # A stack of equations (one per m:e) — keep them apart.
        return "; ".join(_text_of(e) for e in _m_all(el, "e"))

    if tag in ("limLow", "limUpp"):
        mark = "_" if tag == "limLow" else "^"
        return f"{_text_of(_m(el, 'e'))}{mark}{{{_text_of(_m(el, 'lim'))}}}"

    if tag in ("num", "den", "e", "sub", "sup", "deg", "lim"):
        return "".join(linearize_element(c) for c in el)

    if tag.endswith("Pr"):  # property elements (dPr, fPr, ctrlPr, …) hold no text
        return ""

    # Fallback: containers (oMath, oMathPara, acc, groupChr, m, ...) — just
    # concatenate whatever text/known children they hold, in document order.
    return "".join(linearize_element(c) for c in el)


def linearize_omath(omath_el) -> str:
    """Entry point: render one <m:oMath> element as a linear expression."""
    return linearize_element(omath_el).strip()
