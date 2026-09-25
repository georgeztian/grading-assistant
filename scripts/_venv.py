"""Guarantee the toolkit runs inside the project's private `.venv/`.

Every entry-point script imports this module before any third-party import.
If the current interpreter isn't the project venv's, the script is re-run
under `.venv`'s interpreter (so a bare `python scripts/extract.py ...` still
lands in the right environment); if `.venv/` doesn't exist yet, it exits
with instructions to run `scripts/setup_env.py` instead of silently using
whatever packages happen to be installed globally.
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
VENV_DIR = ROOT / ".venv"
_RELAUNCH_FLAG = "GRADING_ASSISTANT_VENV_RELAUNCHED"


def venv_python() -> Path:
    if os.name == "nt":
        return VENV_DIR / "Scripts" / "python.exe"
    return VENV_DIR / "bin" / "python"


def in_project_venv() -> bool:
    try:
        return Path(sys.prefix).resolve() == VENV_DIR.resolve()
    except OSError:
        return False


def ensure() -> None:
    if in_project_venv():
        return
    py = venv_python()
    if not py.exists():
        sys.exit(
            "error: project Python environment not found at .venv/ - create it once with:\n"
            "    python scripts/setup_env.py"
        )
    if os.environ.get(_RELAUNCH_FLAG):
        sys.exit(f"error: relaunched under {py} but still not inside {VENV_DIR}")
    if not Path(sys.argv[0]).is_file():
        # Imported from `python -c`/stdin/REPL — nothing to re-run.
        sys.exit(f"error: not running in the project environment - use {py.relative_to(ROOT)}")
    env = dict(os.environ, **{_RELAUNCH_FLAG: "1"})
    # subprocess rather than os.execv: on Windows execv detaches from the
    # console, so the caller would see the script "finish" before its output.
    sys.exit(subprocess.call([str(py), *sys.argv], env=env))


# Only relaunch when a script is being run directly; setup_env.py imports
# this for its path helpers and must work from any interpreter.
if Path(sys.argv[0]).name != "setup_env.py":
    ensure()
