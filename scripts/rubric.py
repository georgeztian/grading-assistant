#!/usr/bin/env python3
"""Per-solutions-file answer key ("rubric"): built once, independently
verified once, then read by every grader and checker for that homework
instead of each of them reading the raw solutions file.

The rubric is a *reorganized, complete copy* of the solutions file, not a
summary. The rubric builder agent writes only a map — which solution content
(by view id) belongs to which question, plus key points / restated answers —
and this script copies that content into the rubric **verbatim** from the
extraction record. The script also refuses any map that leaves a non-empty
paragraph/table/cell/text box of the solutions file unassigned, so nothing
can be silently dropped.

Verification: an independent rubric checker compares the rubric against the
raw solutions and records a per-question review (review.json) — `approve`
refuses without one covering every question. Only an approved rubric is
served to graders (`path` fails otherwise); any change to the map, the
rubric, the solutions file or the extraction version afterwards makes it
stale until rebuilt (and re-reviewed if its text changed).

A concern about the *official solution itself* (not the rubric) puts the
rubric `on_hold`: nothing is graded against it until the user decides and
the orchestrator runs `release`.

Files live in .cache/rubrics/<sha256 of solutions file>/:
    map.json      written by the rubric builder
    rubric.md     generated here — what graders/checkers read
    status.json   new -> unverified -> verified | rejected | on_hold
    review_round<N>.json / issues_round<N>.json / concerns_<N>.json   evidence

Usage:
    .venv/Scripts/python scripts/rubric.py init    <solutions>
    .venv/Scripts/python scripts/rubric.py build   <solutions>
    .venv/Scripts/python scripts/rubric.py approve <solutions> <review.json>    # rubric checker
    .venv/Scripts/python scripts/rubric.py reject  <solutions> <issues.json>    # rubric checker
    .venv/Scripts/python scripts/rubric.py hold    <solutions> <concerns.json> [<review.json>]
    .venv/Scripts/python scripts/rubric.py release <solutions> --note "<the user's decision>"
    .venv/Scripts/python scripts/rubric.py status  <solutions>
    .venv/Scripts/python scripts/rubric.py path    <solutions>    # rubric.md, only if verified
    .venv/Scripts/python scripts/rubric.py graded  <solutions> <graded_dir>   # outdated outputs
"""
from __future__ import annotations

import argparse
import datetime
import hashlib
import json
import re
import sys
from pathlib import Path

import _venv  # noqa: F401  — must precede third-party imports (re-runs under .venv)

from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries

sys.path.insert(0, str(Path(__file__).resolve().parent))
from lib import cache, views  # noqa: E402
import extract  # noqa: E402

ANSWER_TYPES = ("numeric", "multiple_choice", "short_answer", "conceptual",
                "formula", "table", "mixed")
QUESTION_KEYS = {"id", "title", "blocks", "answer_type", "final_answer", "tolerance",
                 "numeric_tolerance", "key_points", "image_notes", "grading_notes"}
REVIEW_CHECKS = ("complete_and_bounded", "restated_answer_ok", "key_points_ok",
                 "tolerance_ok", "image_notes_ok")
ISSUE_PROBLEMS = ("omission", "misassignment", "wrong_boundary", "key_point_error",
                  "restatement_error", "image_note_error", "exclusion_error",
                  "answer_type_error", "other")
MAX_ROUNDS = 3
CELL_REF_RE = re.compile(r"^\$?[A-Za-z]{1,3}\$?\d+$")


class MapError(Exception):
    pass


class RubricError(Exception):
    """The rubric cannot be used (not verified, stale, on hold)."""


def now() -> str:
    return datetime.datetime.now().isoformat(timespec="seconds")


def file_sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def rubric_dir(solutions: Path) -> Path:
    return cache.RUBRIC_DIR / cache.sha256_of(solutions)


def load_status(d: Path) -> dict:
    f = d / "status.json"
    if f.exists():
        return json.loads(f.read_text(encoding="utf-8"))
    return {"state": "new", "round": 0, "history": []}


def save_status(d: Path, status: dict) -> None:
    cache.write_text_atomic(d / "status.json", json.dumps(status, indent=2, ensure_ascii=False))


def emit(obj: dict, code: int = 0) -> None:
    print(json.dumps(obj, indent=2, ensure_ascii=False))
    sys.exit(code)


# ---------------------------------------------------------------- map parsing

def resolve_spec(spec: str, record: dict, units: list[dict], pos: dict[str, int]) -> list[int]:
    if not isinstance(spec, str) or not spec.strip():
        raise MapError(f"block spec must be a non-empty string, got {spec!r}")
    spec = spec.strip()
    if views.is_sheet_record(record):
        if "!" not in spec:
            raise MapError(f"{spec!r}: workbook specs look like 'Sheet!B5', 'Sheet!A1:D20', "
                           "'Sheet!textbox1' or 'Sheet!*'")
        sheet, ref = views.split_sheet_ref(spec)
        if sheet not in record["sheets"]:
            raise MapError(f"{spec!r}: no sheet named {sheet!r} (sheets: {record['sheet_order']})")
        if ref == "*":
            return [i for i, u in enumerate(units) if u["sheet"] == sheet]
        if ":" in ref:
            try:
                min_col, min_row, max_col, max_row = range_boundaries(ref)
            except ValueError as exc:
                raise MapError(f"{spec!r}: bad range ({exc})")
            # Whole columns (A:C) / rows (2:5) leave the other axis open.
            min_col, min_row = min_col or 1, min_row or 1
            max_col, max_row = max_col or 10 ** 6, max_row or 10 ** 7
            out = []
            for i, u in enumerate(units):
                if u["sheet"] != sheet or u["coord"] is None:  # text boxes/charts: by id
                    continue
                row, col = coordinate_to_tuple(u["coord"])
                if min_row <= row <= max_row and min_col <= col <= max_col:
                    out.append(i)
            return out
        if CELL_REF_RE.match(ref):  # b5 / $B$5 -> B5, as the view writes it
            ref = ref.replace("$", "").upper()
        uid = f"{sheet}!{ref}"
        if uid not in pos:
            raise MapError(f"{spec!r}: no content there in the solutions file (cells, "
                           "or text boxes/charts as Sheet!textbox1 / Sheet!chart1)")
        return [pos[uid]]

    parts = [part.strip() for part in spec.split("-")]
    if len(parts) > 2:
        raise MapError(f"{spec!r}: document specs look like 'p12', 't3', 's0' or 'p12-p20'")
    for part in parts:
        if part not in pos:
            raise MapError(f"{spec!r}: no block {part!r} in the solutions file")
    start, end = pos[parts[0]], pos[parts[-1]]
    if start > end:
        raise MapError(f"{spec!r}: range runs backwards")
    return list(range(start, end + 1))


def resolve_list(specs, where: str, record, units, pos, errors: list[str]) -> list[int]:
    if not isinstance(specs, list) or not specs:
        errors.append(f"{where}: 'blocks' must be a non-empty list of block specs")
        return []
    out: list[int] = []
    for spec in specs:
        try:
            out += resolve_spec(spec, record, units, pos)
        except MapError as exc:
            errors.append(f"{where}: {exc}")
    return sorted(set(out))


def preview(unit: dict, limit: int = 70) -> str:
    text = " ".join(views.render_unit(unit))
    return text if len(text) <= limit else text[:limit - 1] + "…"


def validate_map(m: dict, record: dict) -> tuple[dict, list[str], list[str]]:
    """Return (resolved, errors, warnings)."""
    units = views.units(record)
    pos = {u["id"]: i for i, u in enumerate(units)}
    errors: list[str] = []
    warnings: list[str] = []
    owners: dict[int, list[str]] = {}

    questions = m.get("questions")
    if not isinstance(questions, list) or not questions:
        errors.append("'questions' must be a non-empty list")
        questions = []
    resolved_questions = []
    seen_ids = set()
    for n, q in enumerate(questions):
        where = f"questions[{n}]"
        if not isinstance(q, dict):
            errors.append(f"{where}: must be an object")
            continue
        qid = q.get("id")
        if not isinstance(qid, str) or not qid.strip():
            errors.append(f"{where}: 'id' (e.g. \"Q1\", \"Q2b\") is required")
            qid = f"#{n}"
        elif qid in seen_ids:
            errors.append(f"{where}: duplicate question id {qid!r}")
        seen_ids.add(qid)
        where = f"question {qid}"
        unknown = set(q) - QUESTION_KEYS
        if unknown:
            warnings.append(f"{where}: ignoring unknown field(s) {sorted(unknown)}")
        if q.get("answer_type") not in ANSWER_TYPES:
            errors.append(f"{where}: 'answer_type' must be one of {list(ANSWER_TYPES)}")
        key_points = q.get("key_points", [])
        if not isinstance(key_points, list) or not all(isinstance(k, str) and k.strip() for k in key_points):
            errors.append(f"{where}: 'key_points' must be a list of non-empty strings")
            key_points = []
        if q.get("answer_type") == "conceptual" and not key_points:
            errors.append(f"{where}: conceptual questions need 'key_points' (every key word/point "
                          "the solution's answer relies on)")
        for field in ("title", "final_answer", "tolerance", "image_notes", "grading_notes"):
            if field in q and not isinstance(q[field], str):
                errors.append(f"{where}: '{field}' must be a string")
        tol = q.get("numeric_tolerance")
        if tol is not None and not (isinstance(tol, dict) and len(tol) == 1
                                    and set(tol) <= {"abs", "rel"}
                                    and isinstance(next(iter(tol.values())), (int, float))
                                    and next(iter(tol.values())) >= 0):
            errors.append(f"{where}: 'numeric_tolerance' must be {{\"abs\": x}} or {{\"rel\": x}}")
        idx = resolve_list(q.get("blocks"), where, record, units, pos, errors)
        if idx and not any(units[i]["nonempty"] for i in idx):
            errors.append(f"{where}: its blocks contain no content")
        for i in idx:
            owners.setdefault(i, []).append(qid)
        resolved_questions.append({**q, "_idx": idx})

    shared_specs = m.get("shared_blocks", [])
    shared = resolve_list(shared_specs, "shared_blocks", record, units, pos, errors) if shared_specs else []
    for i in shared:
        owners.setdefault(i, []).append("shared")

    excluded = []
    for n, ex in enumerate(m.get("excluded_blocks", [])):
        where = f"excluded_blocks[{n}]"
        if not isinstance(ex, dict) or not isinstance(ex.get("reason"), str) or not ex["reason"].strip():
            errors.append(f"{where}: needs 'blocks' and a non-empty 'reason'")
            continue
        idx = resolve_list(ex.get("blocks"), where, record, units, pos, errors)
        for i in idx:
            if units[i]["nonempty"] and owners.get(i):
                errors.append(f"{where}: {units[i]['id']} is excluded but also assigned to "
                              f"{owners[i]} — pick one")
            owners.setdefault(i, []).append("excluded")
        excluded.append({"reason": ex["reason"], "_idx": idx, "blocks": ex.get("blocks")})

    uncovered = [u for i, u in enumerate(units) if u["nonempty"] and i not in owners]
    for u in uncovered:
        errors.append(f"uncovered content: {u['id']} {preview(u)!r} — assign it to a question, "
                      "shared_blocks, or excluded_blocks (with a reason)")
    multi = [f"{units[i]['id']}→{o}" for i, o in owners.items()
             if units[i]["nonempty"] and len([x for x in o if x != "shared"]) > 1]
    if multi:
        warnings.append("content assigned to more than one question (fine if deliberate): "
                        + ", ".join(multi[:20]) + (" …" if len(multi) > 20 else ""))
    return {"questions": resolved_questions, "shared": shared, "excluded": excluded,
            "units": units}, errors, warnings


# ------------------------------------------------------------------ rendering

def image_lines(record: dict, selected: list[dict]) -> list[str]:
    out = []
    if record["type"] == "pdf":
        pages = sorted({u["page"] for u in selected if u["nonempty"] and u.get("page") is not None})
        renders = [record["pages"][p]["render_path"] for p in pages
                   if record["pages"][p].get("render_path")]
        if renders:
            out.append("Rendered page(s) to read visually (equations/images — the text layer is "
                       "unreliable there): " + "; ".join(renders))
    elif not views.is_sheet_record(record):
        if any(u["nonempty"] and u["block"].get("has_image") for u in selected):
            out.append("Contains image(s): Read the files in the `(image: …)` tags below — "
                       "their content is not in the text; Image notes describe them.")
    return out


def render_rubric(resolved: dict, record: dict, label: str, sol_sha: str, round_no: int) -> str:
    units = resolved["units"]
    lines = [
        f"# Answer key: {label}",
        "<!-- generated by scripts/rubric.py from map.json — never edit by hand -->",
        f"Solutions file: {label}  (sha256 {sol_sha[:16]}…, map round {round_no})",
        "",
        "How to read: lines starting with an id — [p12], [t3], or a cell like `B5:` under a "
        "[Sheet \"…\"] line — are the solutions file's own content, copied verbatim by script "
        "(authoritative). \"Restated answer\", \"Key points\", \"Image notes\" and \"Grading "
        "notes\" were written by the rubric builder and verified by an independent rubric "
        "checker against the raw solutions. \"⋯\" marks skipped content that belongs to "
        "another section.",
    ]
    source_warnings = views.warnings(record)
    if source_warnings:
        lines += ["", "Source warnings:"] + [f"! {w}" for w in source_warnings]

    if resolved["shared"]:
        lines += ["", "## Shared context (applies to every question)"]
        lines += render_section_body(record, [units[i] for i in resolved["shared"]])

    for q in resolved["questions"]:
        title = f" — {q['title']}" if q.get("title") else ""
        lines += ["", f"## {q['id']}{title}"]
        meta = f"Answer type: {q.get('answer_type')}"
        if q.get("tolerance"):
            meta += f" | Tolerance: {q['tolerance']}"
        if q.get("numeric_tolerance"):
            (kind, value), = q["numeric_tolerance"].items()
            meta += (f" | Numeric tolerance ({'absolute' if kind == 'abs' else 'relative'}, "
                     f"used by compare_xlsx.py): {value:g}")
        lines.append(meta)
        if q.get("final_answer"):
            lines.append(f"Restated answer: {q['final_answer']}")
        if q.get("key_points"):
            lines.append("Key points — a complete answer must cover every one:")
            lines += [f"  {n}. {k}" for n, k in enumerate(q["key_points"], 1)]
        if q.get("image_notes"):
            lines.append(f"Image notes: {q['image_notes']}")
        if q.get("grading_notes"):
            lines.append(f"Grading notes: {q['grading_notes']}")
        lines += render_section_body(record, [units[i] for i in q["_idx"]])

    if resolved["excluded"]:
        lines += ["", "## Excluded from grading (not answer content)"]
        for ex in resolved["excluded"]:
            shown = [units[i] for i in ex["_idx"] if units[i]["nonempty"]]
            first = preview(shown[0], 60) if shown else ""
            lines.append(f"- {', '.join(ex['blocks'])} — {ex['reason']} (first: {first!r})")
    return "\n".join(lines) + "\n"


def render_section_body(record: dict, selected: list[dict]) -> list[str]:
    return image_lines(record, selected) + ["Solution (verbatim):"] + views.render_units(record, selected)


# ------------------------------------------------------------------- commands

def cmd_init(solutions: Path) -> None:
    record, _cache_path, view, _hit = extract.get_record(solutions)
    d = rubric_dir(solutions)
    d.mkdir(parents=True, exist_ok=True)
    status = load_status(d)
    emit({
        "rubric_dir": str(d),
        "map_path": str(d / "map.json"),
        "view_path": str(view),
        "state": status["state"],
        "round": status["round"],
        "nonempty_units": sum(1 for u in views.units(record) if u["nonempty"]),
        "id_format": ("'Sheet!B5', 'Sheet!A1:D20', 'Sheet!textbox1', 'Sheet!chart1', 'Sheet!*'"
                      if views.is_sheet_record(record)
                      else "'p12', 't3', 's0', 'p12-p20' (range in document order)"),
    })


def cmd_build(solutions: Path) -> None:
    d = rubric_dir(solutions)
    map_file = d / "map.json"
    if not map_file.exists():
        emit({"error": f"no map at {map_file} — run `rubric.py init` and write it first"}, 1)
    try:
        m = json.loads(map_file.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        emit({"error": f"map.json is not valid JSON: {exc}"}, 1)
    record, *_ = extract.get_record(solutions)
    resolved, errors, warnings = validate_map(m, record)
    if errors:
        emit({"result": "invalid", "errors": errors, "warnings": warnings,
              "note": "rubric.md not written — fix map.json and rebuild"}, 1)

    status = load_status(d)
    sol_sha = cache.sha256_of(solutions)
    map_sha = file_sha(map_file)
    # A new round: first build, a correction after rejection, or a change made
    # after the user's decision on a hold (all need a fresh review).
    if status["state"] in ("new", "rejected", "on_hold"):
        status["round"] += 1
    round_no = max(status["round"], 1)
    text = render_rubric(resolved, record, cache.rel_path(solutions), sol_sha, round_no)
    rubric_file = d / "rubric.md"
    rubric_sha = hashlib.sha256(text.encode("utf-8")).hexdigest()

    if (status["state"] == "verified" and status.get("rubric_sha256") == rubric_sha
            and status.get("map_sha256") == map_sha):
        # Same content as approved — (re)write it in case the file on disk
        # was altered, and adopt the current extraction version: the approval
        # still applies byte for byte.
        if not rubric_file.exists() or file_sha(rubric_file) != rubric_sha:
            cache.write_text_atomic(rubric_file, text)
        if status.get("schema_version") != cache.SCHEMA_VERSION:
            status["schema_version"] = cache.SCHEMA_VERSION
            save_status(d, status)
        emit({"result": "unchanged", "state": "verified", "rubric_path": str(rubric_file)})
    if status["state"] == "verified":
        status["round"] += 1  # a changed, previously approved rubric needs a new review
        round_no = status["round"]
        text = render_rubric(resolved, record, cache.rel_path(solutions), sol_sha, round_no)
        rubric_sha = hashlib.sha256(text.encode("utf-8")).hexdigest()

    cache.write_text_atomic(rubric_file, text)
    status.update({
        "state": "unverified", "solutions_file": cache.rel_path(solutions),
        "solutions_sha256": sol_sha, "map_sha256": map_sha, "rubric_sha256": rubric_sha,
        "schema_version": cache.SCHEMA_VERSION, "built_at": now(),
    })
    status.setdefault("history", []).append({"round": round_no, "event": "built", "at": now()})
    save_status(d, status)
    emit({"result": "built", "state": "unverified", "round": round_no,
          "rubric_path": str(rubric_file), "questions": len(resolved["questions"]),
          "warnings": warnings,
          "next": "an independent rubric-checker must review it (approve/reject) before grading"})


def _check_unchanged(d: Path, status: dict, solutions: Path) -> str | None:
    rubric_file = d / "rubric.md"
    if not rubric_file.exists():
        return "rubric.md is missing"
    if file_sha(rubric_file) != status.get("rubric_sha256"):
        return "rubric.md changed after it was built (edited by hand?) — rebuild it"
    if not (d / "map.json").exists():
        return "map.json is missing — rebuild it (rubric-builder), then re-review"
    if file_sha(d / "map.json") != status.get("map_sha256"):
        return "map.json changed after the last build — rebuild, then re-review"
    if cache.sha256_of(solutions) != status.get("solutions_sha256"):
        return "solutions file changed"
    if status.get("schema_version") != cache.SCHEMA_VERSION:
        return ("extraction changed since the build (schema "
                f"{status.get('schema_version')} -> {cache.SCHEMA_VERSION}) — rebuild; it stays "
                "verified only if its text comes out identical")
    return None


def load_map(d: Path) -> dict:
    return json.loads((d / "map.json").read_text(encoding="utf-8"))


def verified_rubric(solutions: Path) -> dict:
    """The usable rubric for `solutions`, or RubricError. Returns
    {path, rubric_sha256, solutions_file, solutions_sha256, round, question_ids}."""
    d = rubric_dir(solutions)
    status = load_status(d)
    if status["state"] == "on_hold":
        raise RubricError(f"rubric for {solutions} is ON HOLD pending the user's decision on a "
                          f"solution concern ({status.get('last_concerns')}) — do not grade")
    if status["state"] != "verified":
        raise RubricError(f"rubric for {solutions} is not verified (state: {status['state']!r}) — "
                          "do not grade against it; the orchestrator must build and verify it first")
    problem = _check_unchanged(d, status, solutions)
    if problem:
        raise RubricError(f"rubric is stale: {problem} — do not grade against it")
    return {"path": str(d / "rubric.md"), "rubric_sha256": status["rubric_sha256"],
            "solutions_file": cache.rel_path(solutions),
            "solutions_sha256": status["solutions_sha256"], "round": status["round"],
            "question_ids": [q["id"] for q in load_map(d)["questions"]]}


def validate_review(review_file: Path, d: Path, status: dict) -> dict:
    """A rubric checker's per-question evidence: every question in the map,
    every check answered (true, or null where the question has nothing to
    check), all true — written after this build."""
    try:
        review = json.loads(review_file.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise RubricError(f"cannot read review file: {exc}")
    if review_file.stat().st_mtime < datetime.datetime.fromisoformat(status["built_at"]).timestamp() - 1:
        raise RubricError("review file predates the current build — review the rubric as built now")
    questions = {q["id"]: q for q in load_map(d)["questions"]}
    per_q = review.get("questions") if isinstance(review, dict) else None
    if not isinstance(per_q, dict):
        raise RubricError('review must be {"questions": {"<id>": {checks…}}, "shared_ok": …, '
                          '"exclusions_ok": …, "inventory_ok": …}')
    problems = []
    missing = [qid for qid in questions if qid not in per_q]
    if missing:
        problems.append(f"no review for question(s) {missing}")
    extra = [qid for qid in per_q if qid not in questions]
    if extra:
        problems.append(f"review names unknown question(s) {extra}")
    applies = {"restated_answer_ok": "final_answer", "key_points_ok": "key_points",
               "tolerance_ok": ("tolerance", "numeric_tolerance"), "image_notes_ok": "image_notes"}
    for qid, checks in per_q.items():
        if qid not in questions or not isinstance(checks, dict):
            continue
        for check in REVIEW_CHECKS:
            value = checks.get(check, "missing")
            fields = applies.get(check)
            fields = (fields,) if isinstance(fields, str) else fields
            relevant = fields is None or any(questions[qid].get(f) for f in fields)
            if value is None and not relevant:
                continue
            if value is not True:
                problems.append(f"{qid}.{check} is {value!r} (must be true"
                                + ("" if relevant else ", or null — not applicable") + ")")
    for check in ("inventory_ok", "shared_ok", "exclusions_ok"):
        if review.get(check) is not True:
            problems.append(f"{check} must be true")
    if problems:
        raise RubricError("review incomplete or not passing — " + "; ".join(problems)
                          + " (anything not true is a rejection: use `reject` with the issues)")
    return review


def cmd_approve(solutions: Path, review_file: Path) -> None:
    d = rubric_dir(solutions)
    status = load_status(d)
    if status["state"] != "unverified":
        emit({"error": f"can only approve a freshly built rubric (state is {status['state']!r})"}, 1)
    problem = _check_unchanged(d, status, solutions)
    if problem:
        emit({"error": problem}, 1)
    try:
        review = validate_review(review_file, d, status)
    except RubricError as exc:
        emit({"error": str(exc)}, 1)
    target = d / f"review_round{status['round']}.json"
    cache.write_text_atomic(target, json.dumps(review, indent=2, ensure_ascii=False))
    status.update({"state": "verified", "verified_at": now(), "last_review": str(target)})
    status["history"].append({"round": status["round"], "event": "approved", "at": now()})
    save_status(d, status)
    emit({"result": "verified", "rubric_path": str(d / "rubric.md"), "round": status["round"]})


def cmd_reject(solutions: Path, issues_file: Path) -> None:
    d = rubric_dir(solutions)
    status = load_status(d)
    if status["state"] != "unverified":
        emit({"error": f"can only reject a freshly built rubric (state is {status['state']!r})"}, 1)
    try:
        issues = json.loads(issues_file.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        emit({"error": f"cannot read issues file: {exc}"}, 1)
    items = issues.get("issues") if isinstance(issues, dict) else None
    if not isinstance(items, list) or not items:
        emit({"error": "issues file must be {\"issues\": [ {question, problem, detail, fix}, … ]} "
                       "with at least one issue"}, 1)
    for n, it in enumerate(items):
        if not isinstance(it, dict) or it.get("problem") not in ISSUE_PROBLEMS or not it.get("detail"):
            emit({"error": f"issues[{n}] needs 'problem' (one of {list(ISSUE_PROBLEMS)}) and 'detail'"}, 1)
    target = d / f"issues_round{status['round']}.json"
    cache.write_text_atomic(target, json.dumps(issues, indent=2, ensure_ascii=False))
    status.update({"state": "rejected", "last_issues": str(target)})
    status["history"].append({"round": status["round"], "event": "rejected",
                              "issues": len(items), "at": now()})
    save_status(d, status)
    escalate = status["round"] >= MAX_ROUNDS
    emit({"result": "rejected", "issues_path": str(target), "round": status["round"],
          "escalate_to_user": escalate,
          "next": ("stop: rounds exhausted — show the user the issues and ask how to proceed"
                   if escalate else
                   "re-invoke the rubric-builder with this issues file, rebuild, then a fresh "
                   "rubric-checker review")})


def cmd_hold(solutions: Path, concerns_file: Path, review_file: Path | None) -> None:
    """A suspected error in the official solution (or a grader/checker
    concern the rubric can't settle): stop all grading on this rubric until
    the user decides. From `unverified` the rubric checker must also supply
    its passing review, so a release can go straight to verified. A further
    concern while already on hold is added to the outstanding ones."""
    d = rubric_dir(solutions)
    status = load_status(d)
    if status["state"] not in ("unverified", "verified", "on_hold"):
        emit({"error": f"can only hold an unverified, verified or on-hold rubric (state is "
                       f"{status['state']!r})"}, 1)
    try:
        concerns = json.loads(concerns_file.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        emit({"error": f"cannot read concerns file: {exc}"}, 1)
    items = concerns.get("concerns") if isinstance(concerns, dict) else None
    if not isinstance(items, list) or not items or not all(
            isinstance(c, dict) and c.get("question") and c.get("detail") for c in items):
        emit({"error": 'concerns file must be {"concerns": [{"question", "detail", "source"}, …]}'}, 1)
    if status["state"] == "unverified":
        if review_file is None:
            emit({"error": "holding an unverified rubric needs the checker's passing review.json "
                           "too (the rubric itself must be faithful; only the solution is in doubt)"}, 1)
        try:
            review = validate_review(review_file, d, status)
        except RubricError as exc:
            emit({"error": str(exc)}, 1)
        review_target = d / f"review_round{status['round']}.json"
        cache.write_text_atomic(review_target, json.dumps(review, indent=2, ensure_ascii=False))
        status["last_review"] = str(review_target)
    elif status["state"] == "on_hold" and status.get("last_concerns"):
        try:  # keep every outstanding concern in the latest file the user is shown
            earlier = json.loads(Path(status["last_concerns"]).read_text(encoding="utf-8"))
            concerns = {"concerns": earlier.get("concerns", []) + items}
        except (OSError, json.JSONDecodeError, AttributeError):
            pass
    n = sum(1 for e in status["history"] if e["event"] == "held") + 1
    target = d / f"concerns_{n}.json"
    cache.write_text_atomic(target, json.dumps(concerns, indent=2, ensure_ascii=False))
    status.update({"state": "on_hold", "last_concerns": str(target)})
    status["history"].append({"round": status["round"], "event": "held", "at": now()})
    save_status(d, status)
    emit({"result": "on_hold", "concerns_path": str(target),
          "next": "stop grading this homework; show the user the concerns and ask for a decision"})


def cmd_release(solutions: Path, note: str) -> None:
    """The user decided the solution stands as is: back to verified. (If the
    user instead corrects the solutions file, or the rubric needs an
    accepted alternative, rebuild instead — that needs a fresh review.)"""
    d = rubric_dir(solutions)
    status = load_status(d)
    if status["state"] != "on_hold":
        emit({"error": f"nothing to release (state is {status['state']!r})"}, 1)
    if not note.strip():
        emit({"error": "--note must record the user's decision"}, 1)
    problem = _check_unchanged(d, status, solutions)
    if problem:
        emit({"error": f"{problem} — rebuild and re-review instead of releasing"}, 1)
    status.update({"state": "verified", "released_at": now()})
    status["history"].append({"round": status["round"], "event": "released", "note": note, "at": now()})
    save_status(d, status)
    emit({"result": "verified", "rubric_path": str(d / "rubric.md")})


def cmd_status(solutions: Path) -> None:
    d = rubric_dir(solutions)
    status = load_status(d)
    out = {"rubric_dir": str(d), **{k: v for k, v in status.items() if k != "history"}}
    if status["state"] == "verified":
        problem = _check_unchanged(d, status, solutions)
        if problem:
            out["state"] = "stale"
            out["problem"] = problem
    emit(out)


def cmd_path(solutions: Path) -> None:
    try:
        info = verified_rubric(solutions)
    except RubricError as exc:
        emit({"error": str(exc)}, 1)
    emit({"rubric_path": info["path"], "state": "verified", "round": info["round"],
          "question_ids": info["question_ids"]})


def cmd_graded(solutions: Path, graded_dir: Path) -> None:
    """Graded files made from this solutions file whose recorded rubric is
    not the current verified one — they were graded against an older (or
    since-held) answer key, or an earlier version of this solutions file
    (same path, since corrected), and must be regraded or reviewed by hand."""
    from lib import marks
    sol_sha = cache.sha256_of(solutions)
    sol_file = cache.rel_path(solutions)
    try:
        current = verified_rubric(solutions)["rubric_sha256"]
    except RubricError as exc:
        current, why = None, str(exc)
    outdated, current_ok, unrecorded = [], [], []
    for f in sorted(graded_dir.rglob("*_Graded.*")):
        if f.suffix.lower() not in (".docx", ".xlsx") or f.name.startswith("~$"):
            continue
        meta = marks.read_meta(f)
        if meta is None:
            unrecorded.append(cache.rel_path(f))
            continue
        if meta.get("solutions_sha256") != sol_sha:
            if meta.get("solutions_file") == sol_file:  # graded against the file before a correction
                outdated.append(cache.rel_path(f))
            continue
        (current_ok if meta.get("rubric_sha256") == current else outdated).append(cache.rel_path(f))
    out = {"current_rubric": current[:16] if current else None, "outdated": outdated,
           "up_to_date": len(current_ok),
           "no_record": unrecorded}
    if current is None:
        out["note"] = f"no usable rubric right now: {why}"
    emit(out, 1 if outdated else 0)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("command", choices=["init", "build", "approve", "reject", "hold",
                                            "release", "status", "path", "graded"])
    parser.add_argument("solutions", type=Path)
    parser.add_argument("files", type=Path, nargs="*",
                        help="approve: review.json · reject: issues.json · "
                             "hold: concerns.json [review.json] · graded: graded dir")
    parser.add_argument("--note", default="", help="release: the user's decision, recorded")
    args = parser.parse_args()

    if not args.solutions.exists():
        emit({"error": f"file not found: {args.solutions}"}, 1)
    needs = {"approve": 1, "reject": 1, "hold": 1, "graded": 1}
    if len(args.files) < needs.get(args.command, 0):
        emit({"error": f"{args.command} needs: " + parser._actions[3].help}, 1)
    simple = {"init": cmd_init, "build": cmd_build, "status": cmd_status, "path": cmd_path}
    try:
        if args.command in simple:
            simple[args.command](args.solutions)
        elif args.command == "approve":
            cmd_approve(args.solutions, args.files[0])
        elif args.command == "reject":
            cmd_reject(args.solutions, args.files[0])
        elif args.command == "hold":
            cmd_hold(args.solutions, args.files[0], args.files[1] if len(args.files) > 1 else None)
        elif args.command == "release":
            cmd_release(args.solutions, args.note)
        else:
            cmd_graded(args.solutions, args.files[0])
    except Exception as exc:  # extraction/conversion failure, corrupt file, …
        emit({"error": f"{type(exc).__name__}: {exc}"}, 1)


if __name__ == "__main__":
    main()
