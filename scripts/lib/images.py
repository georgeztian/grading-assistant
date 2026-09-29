"""Export embedded images to viewable files, once per source file.

The extraction has no OCR, so an image's content (a pasted screenshot of
work, an equation a legacy .doc stores as a picture, a figure) is never in
the text. Rather than make every agent unzip the file and guess which image
belongs where, the extractors export each image here and the view shows its
path right next to the paragraph/sheet it belongs to — one Read away.

Files go to `.cache/renders/<sha256 of source>/`. Vector formats that image
viewers can't open (EMF/WMF — common for equations in files that went through
the legacy .doc format) are converted to PNG: Pillow first, then (Windows)
.NET's System.Drawing via PowerShell. If neither works the original file is
kept and its path still listed.
"""
from __future__ import annotations

import os
import subprocess
from pathlib import Path

from lib.cache import rel_path

VECTOR_EXTS = {".emf", ".wmf"}


def _ps_quote(path: Path) -> str:
    """A PowerShell single-quoted string literal ('' escapes a quote)."""
    return "'" + str(path).replace("'", "''") + "'"


def _vector_to_png(src: Path, dst: Path) -> bool:
    try:
        from PIL import Image
        with Image.open(src) as im:
            im.load()
            im.save(dst)
        return True
    except Exception:
        pass
    if os.name != "nt":
        return False
    script = ("Add-Type -AssemblyName System.Drawing; "
              f"$i=[System.Drawing.Image]::FromFile({_ps_quote(src)}); "
              "$b=New-Object System.Drawing.Bitmap($i.Width,$i.Height); "
              "$g=[System.Drawing.Graphics]::FromImage($b); $g.Clear([System.Drawing.Color]::White); "
              "$g.DrawImage($i,0,0,$i.Width,$i.Height); "
              f"$b.Save({_ps_quote(dst)},[System.Drawing.Imaging.ImageFormat]::Png); $g.Dispose(); $b.Dispose(); $i.Dispose()")
    try:
        subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
                       capture_output=True, timeout=60, check=True)
        return dst.exists()
    except (OSError, subprocess.SubprocessError):
        return False


def export_image(blob: bytes, name: str, out_dir: Path) -> str:
    """Write `blob` as `out_dir/name` (PNG-converted if vector) and return
    its repo-relative path."""
    out_dir.mkdir(parents=True, exist_ok=True)
    target = out_dir / name
    # Rewrite whenever the bytes differ: a re-extraction (new schema, --force)
    # can produce a different picture under the same name, and a stale file
    # would silently show the grader the wrong image.
    unchanged = target.exists() and target.read_bytes() == blob
    if not unchanged:
        target.write_bytes(blob)
    if target.suffix.lower() in VECTOR_EXTS:
        png = target.with_suffix(".png")
        if unchanged and png.exists():
            return rel_path(png)
        if _vector_to_png(target, png):
            return rel_path(png)
    return rel_path(target)


def image_ext(fmt: str | None, fallback: str = ".png") -> str:
    fmt = (fmt or "").lower().lstrip(".")
    return f".{'jpg' if fmt == 'jpeg' else fmt}" if fmt else fallback

