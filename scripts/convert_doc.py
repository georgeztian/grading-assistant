#!/usr/bin/env python3
"""Convert legacy .doc / .xls files to .docx / .xlsx.

python-docx cannot open .doc and openpyxl cannot open .xls, and xlrd (the
only pure-Python .xls reader) exposes cached values but no formulas — so the
toolkit converts legacy files to their OOXML equivalents first and then
extracts/annotates those exactly like native .docx/.xlsx files: formulas,
formatting, merged cells and images all survive into the graded copy.

Converters, in order of preference (first one installed wins):
1. Microsoft Word / Excel via COM (Windows) — the native applications, so
   the conversion is exact. Runs in a child process with a timeout (a stuck
   dialog can't hang the caller), with macros force-disabled and files
   opened read-only. Needs pywin32 (in requirements.txt on Windows).
2. LibreOffice (`soffice --headless`) — on PATH or in its standard install
   location; the option on macOS/Linux.
If neither is available this raises a clear `no_converter` error instead of
guessing. (`pandoc` is not an option: it cannot read the binary formats.)

Output goes to `<out-dir>/<sha256 of the source>/converted.<ext>` — keyed
and named by content, not by any student's file name — so a file is
converted only once (later calls reuse the existing output), and two
students' identical files share it without either's name showing. Conversions are
serialized across processes with a lock file, since several agents may ask
at the same time and Office automation does not like concurrent sessions.

Usage:
    .venv/Scripts/python scripts/convert_doc.py <file.doc|file.xls> [--out-dir DIR]
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

import _venv  # noqa: F401  — must precede third-party imports (re-runs under .venv)

sys.path.insert(0, str(Path(__file__).resolve().parent))
from lib import cache  # noqa: E402

TARGET_EXT = {".doc": "docx", ".xls": "xlsx"}
OFFICE_TIMEOUT = 180
SOFFICE_TIMEOUT = 300
LOCK_TIMEOUT = 600

# LibreOffice is often installed without being put on PATH (always so on
# Windows), so also check its standard install locations.
_SOFFICE_CANDIDATES = [
    Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
    Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
    Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
]
_OFFICE_PROGID = {"docx": "Word.Application", "xlsx": "Excel.Application"}


def find_soffice() -> str | None:
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    for candidate in _SOFFICE_CANDIDATES:
        if candidate.exists():
            return str(candidate)
    return None


def office_available(target_ext: str) -> bool:
    """Is Microsoft Word/Excel registered for COM automation (and pywin32
    importable)?"""
    if os.name != "nt":
        return False
    try:
        import winreg
        import win32com.client  # noqa: F401
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, _OFFICE_PROGID[target_ext] + r"\CLSID"):
            return True
    except (ImportError, OSError):
        return False


def available_converters(target_ext: str) -> list[str]:
    out = []
    if office_available(target_ext):
        out.append("office")
    if find_soffice():
        out.append("libreoffice")
    return out


class _Lock:
    """Cross-process lock via an exclusively created file (stale after
    LOCK_TIMEOUT, in case a holder crashed)."""

    def __init__(self, path: Path):
        self.path = path

    def __enter__(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        deadline = time.time() + LOCK_TIMEOUT
        while True:
            try:
                fd = os.open(str(self.path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                os.write(fd, str(os.getpid()).encode())
                os.close(fd)
                return self
            except FileExistsError:
                try:
                    if time.time() - self.path.stat().st_mtime > LOCK_TIMEOUT:
                        self.path.unlink()
                        continue
                except OSError:
                    continue
                if time.time() > deadline:
                    raise RuntimeError("conversion_failed: timed out waiting for the conversion lock")
                time.sleep(0.5)

    def __exit__(self, *exc):
        try:
            self.path.unlink()
        except OSError:
            pass


# ------------------------------------------------------------ Office (COM)

def _office_worker(src: str, dst: str, target_ext: str) -> None:
    """Runs in a child process: open `src` read-only in Word/Excel with
    macros disabled and save it as OOXML to `dst`."""
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        if target_ext == "docx":
            app = win32com.client.DispatchEx("Word.Application")
            try:
                app.Visible = False
                app.DisplayAlerts = 0  # wdAlertsNone
                app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
                doc = app.Documents.Open(FileName=src, ConfirmConversions=False, ReadOnly=True,
                                         AddToRecentFiles=False, Visible=False,
                                         PasswordDocument="__no_password__", NoEncodingDialog=True)
                try:
                    doc.SaveAs2(FileName=dst, FileFormat=16)  # wdFormatDocumentDefault (.docx)
                finally:
                    doc.Close(SaveChanges=0)
            finally:
                app.Quit()
        else:
            app = win32com.client.DispatchEx("Excel.Application")
            try:
                app.Visible = False
                app.DisplayAlerts = False
                app.AskToUpdateLinks = False
                app.EnableEvents = False
                app.AutomationSecurity = 3
                wb = app.Workbooks.Open(Filename=src, UpdateLinks=0, ReadOnly=True,
                                        Password="__no_password__", IgnoreReadOnlyRecommended=True,
                                        AddToMru=False)
                try:
                    wb.SaveAs(Filename=dst, FileFormat=51)  # xlOpenXMLWorkbook (.xlsx)
                finally:
                    wb.Close(SaveChanges=False)
            finally:
                app.Quit()
    finally:
        pythoncom.CoUninitialize()


def _convert_office(path: Path, tmp: Path, target_ext: str) -> None:
    cmd = [sys.executable, str(Path(__file__).resolve()), "--office-worker",
           str(path.resolve()), str(tmp.resolve()), target_ext]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=OFFICE_TIMEOUT)
    except subprocess.TimeoutExpired:
        raise RuntimeError(f"conversion_failed: Microsoft Office timed out on {path.name} "
                           "(a dialog, password or repair prompt?)")
    if proc.returncode != 0 or not tmp.exists():
        detail = (proc.stderr or proc.stdout or "").strip().splitlines()[-1:] or ["no output"]
        raise RuntimeError(f"conversion_failed: Microsoft Office could not convert {path.name}: "
                           f"{detail[0]}")


# -------------------------------------------------------------- LibreOffice

def _convert_soffice(path: Path, tmp_dir: Path, target_ext: str) -> Path:
    converter = find_soffice()
    try:
        subprocess.run([converter, "--headless", "--convert-to", target_ext,
                        "--outdir", str(tmp_dir), str(path)],
                       check=True, capture_output=True, timeout=SOFFICE_TIMEOUT)
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired) as exc:
        raise RuntimeError(f"conversion_failed: {converter} failed on {path.name}: {exc}") from exc
    produced = tmp_dir / f"{path.stem}.{target_ext}"
    if not produced.exists():
        raise RuntimeError(f"conversion_failed: {converter} did not produce {produced.name}")
    return produced


# ------------------------------------------------------------------ entry

def convert(path: Path, out_dir: Path, target_ext: str | None = None) -> Path:
    """Convert `path` (.doc/.xls) to OOXML under `<out_dir>/<sha256>/`,
    reusing an earlier conversion of the same content. Raises RuntimeError
    (`no_converter` / `conversion_failed`) rather than guessing."""
    target_ext = target_ext or TARGET_EXT.get(path.suffix.lower())
    if target_ext not in ("docx", "xlsx"):
        raise ValueError(f"unsupported conversion: {path.suffix} -> {target_ext}")
    target_dir = out_dir / cache.sha256_of(path)
    # Content-neutral name: the conversion is shared by every file with this
    # content, so it must not carry one student's file name.
    target = target_dir / f"converted.{target_ext}"
    if target.exists() and target.stat().st_size > 0:
        return target

    converters = available_converters(target_ext)
    if not converters:
        app = "Word" if target_ext == "docx" else "Excel"
        raise RuntimeError(
            f"no_converter: neither Microsoft {app} (Windows) nor LibreOffice (soffice) is "
            f"installed — flag {path.name} for manual conversion rather than guessing at its "
            "content."
        )

    errors = []
    with _Lock(out_dir / ".convert.lock"):
        if target.exists() and target.stat().st_size > 0:  # converted while we waited
            return target
        target_dir.mkdir(parents=True, exist_ok=True)
        for name in converters:
            work = target_dir / f".tmp-{os.getpid()}-{name}"
            work.mkdir(exist_ok=True)
            try:
                if name == "office":
                    produced = work / f"{path.stem}.{target_ext}"
                    _convert_office(path, produced, target_ext)
                else:
                    produced = _convert_soffice(path, work, target_ext)
                cache.replace_with_retry(produced, target)
                return target
            except RuntimeError as exc:
                errors.append(str(exc))
            finally:
                shutil.rmtree(work, ignore_errors=True)
    raise RuntimeError("; ".join(errors))


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) == 5 and sys.argv[1] == "--office-worker":
        _office_worker(sys.argv[2], sys.argv[3], sys.argv[4])
        return
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("file", type=Path)
    parser.add_argument("--out-dir", type=Path, default=cache.CONVERTED_DIR)
    args = parser.parse_args()

    if not args.file.exists():
        print(json.dumps({"error": f"file not found: {args.file}"}))
        sys.exit(1)
    try:
        result_path = convert(args.file, args.out_dir)
    except (RuntimeError, ValueError) as exc:
        print(json.dumps({"error": str(exc)}))
        sys.exit(1)
    print(json.dumps({"converted_path": str(result_path),
                      "converters_available": available_converters(TARGET_EXT[args.file.suffix.lower()])}))


if __name__ == "__main__":
    main()
