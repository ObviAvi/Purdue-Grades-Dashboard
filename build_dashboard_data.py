"""
Parse grades.xlsx (all sheets — Purdue EM grade distribution exports) and emit
grades_data.json for purdue_grade_dashboard.html.

Run: python build_dashboard_data.py

Sheets use different layouts (contiguous % columns, interleaved Students/%,
or compact headers). Older builds only read the first sheet, which dropped
Fall 2021–Spring 2025 tabs.
"""
from __future__ import annotations

import json
import math
import re
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent
XLSX = ROOT / "grades.xlsx"
OUT_JSON = ROOT / "grades_data.json"


def sem_sort_key(desc: str) -> float:
    if not desc or not isinstance(desc, str):
        return 0.0
    parts = desc.strip().split()
    if len(parts) != 2:
        return 0.0
    term, yr_s = parts[0], parts[1]
    try:
        y = int(yr_s)
    except ValueError:
        return 0.0
    off = {"Spring": 0.0, "Summer": 0.25, "Fall": 0.5}.get(term, 0.1)
    return y + off


def normalize_course_num(x) -> str | None:
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return None
    if isinstance(x, (int, float)) and not isinstance(x, bool):
        if isinstance(x, float) and math.isnan(x):
            return None
        if float(x) == int(float(x)):
            return str(int(float(x)))
        return str(x).strip()
    s = str(x).strip()
    if re.fullmatch(r"\d+\.0", s):
        return s[:-2]
    return s


def map_excel_letter_to_bucket(letter: str) -> str | None:
    if letter is None or (isinstance(letter, float) and math.isnan(letter)):
        return None
    letter = str(letter).strip()
    if letter in ("A", "A-", "A+", "B", "B-", "B+", "C", "C-", "C+", "D", "D-", "D+", "F", "W"):
        return letter
    if letter == "E":
        return "F"
    if letter in ("AU", "I", "IF", "N", "NS", "P", "PI", "S", "SI", "U"):
        return "O"
    if letter in ("FN",):
        return "F"
    if letter in ("WF", "WN", "WU", "WIP"):
        return "W"
    return None


def find_subject_header_row(df: pd.DataFrame) -> int | None:
    for r in range(min(len(df), 60)):
        v = df.iloc[r, 0]
        if pd.notna(v) and str(v).strip() == "Subject":
            return r
    return None


def detect_grade_layout(df: pd.DataFrame, h: int) -> tuple[str, list[str], list[int]] | None:
    """Return (mode, letters, pct_column_indices) or None."""
    row = df.iloc[h]
    c9 = row[9] if len(row) > 9 else None
    s9 = str(c9).strip() if pd.notna(c9) else ""

    if s9 == "% of Total":
        letters_row = h - 1
        letters: list[str] = []
        pct_cols: list[int] = []
        for c in range(9, len(row)):
            lab = df.iloc[letters_row, c]
            if lab is None or (isinstance(lab, float) and math.isnan(lab)):
                break
            sl = str(lab).strip()
            if sl in ("% of Total", "Students"):
                break
            letters.append(sl)
            pct_cols.append(c)
        if letters:
            return ("contiguous", letters, pct_cols)
        return None

    if s9 == "Students":
        pr = df.iloc[h - 1]
        letters = []
        pct_cols = []
        c = 9
        while c < len(pr):
            lab = pr[c]
            if lab is None or (isinstance(lab, float) and math.isnan(lab)):
                c += 1
                continue
            sl = str(lab).strip()
            if sl in ("% of Total", "Students"):
                c += 1
                continue
            if re.match(r"^-?\d+\.?\d*$", sl):
                c += 1
                continue
            letters.append(sl)
            pct_cols.append(c + 1)
            c += 2
        if len(letters) >= 3:
            return ("interleaved", letters, pct_cols)
        return None

    if s9 and s9 not in ("% of Total", "Students") and map_excel_letter_to_bucket(s9) is not None:
        letters = []
        pct_cols = []
        c = 9
        while c < len(row):
            lab = row[c]
            if lab is None or (isinstance(lab, float) and math.isnan(lab)):
                break
            sl = str(lab).strip()
            if sl in ("% of Total", "Students"):
                break
            letters.append(sl)
            pct_cols.append(c)
            c += 1
        if letters:
            return ("same_line", letters, pct_cols)
    return None


def build_grade_dict(
    row: list,
    letters: list[str],
    pct_cols: list[int],
) -> dict[str, float]:
    out: dict[str, float] = {}
    for lab, col in zip(letters, pct_cols):
        if col >= len(row):
            break
        v = row[col]
        if v is None or (isinstance(v, float) and math.isnan(v)):
            continue
        try:
            pct = float(v)
        except (TypeError, ValueError):
            continue
        if pct <= 0:
            continue
        bucket = map_excel_letter_to_bucket(lab)
        if bucket is None:
            continue
        out[bucket] = out.get(bucket, 0.0) + pct
    s = sum(out.values())
    if s <= 0:
        return {}
    if abs(s - 1.0) > 0.02:
        for k in list(out.keys()):
            out[k] = out[k] / s
    return out


def parse_sheet(df: pd.DataFrame, sheet_name: str, course_titles: dict[str, str]) -> list[tuple]:
    h = find_subject_header_row(df)
    if h is None:
        print(f"  skip {sheet_name}: no Subject header", file=sys.stderr)
        return []

    layout = detect_grade_layout(df, h)
    if layout is None:
        print(f"  skip {sheet_name}: unknown grade layout", file=sys.stderr)
        return []

    _mode, letters, pct_cols = layout
    data_start = h + 1

    subj = None
    course = None
    title = None
    last_sem: str | None = None
    rows_out: list[tuple] = []

    for r in range(data_start, len(df)):
        row = df.iloc[r]
        c0 = row[0]
        c2 = row[2]
        c3 = row[3]
        c5 = row[5]
        c8 = row[8]

        if pd.notna(c0):
            subj = str(c0).strip()

        if pd.notna(c2):
            cn = normalize_course_num(c2)
            if cn:
                course = cn
                if pd.notna(c3):
                    title = str(c3).strip()
                key = f"{subj}|{course}"
                if title:
                    course_titles[key] = title

        if pd.isna(c2) and pd.notna(c3) and pd.notna(row[6]):
            title = str(c3).strip()
            if subj and course:
                course_titles[f"{subj}|{course}"] = title

        sem = None
        if pd.notna(c5):
            sem = str(c5).strip()
            last_sem = sem
        elif last_sem:
            sem = last_sem

        inst = None
        if pd.notna(c8):
            inst = str(c8).strip()

        if not subj or not course or not sem:
            continue

        grades = build_grade_dict(row.tolist(), letters, pct_cols)
        if not grades:
            continue

        rows_out.append((subj, course, sem, grades, inst))

    print(f"  {sheet_name}: {len(rows_out)} section rows ({layout[0]})")
    return rows_out


def main() -> None:
    if not XLSX.exists():
        print("Missing", XLSX, file=sys.stderr)
        sys.exit(1)

    xl = pd.ExcelFile(XLSX)
    all_rows: list[tuple] = []
    course_titles: dict[str, str] = {}
    for name in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=name, header=None)
        all_rows.extend(parse_sheet(df, name, course_titles))

    semesters_set: set[str] = set()
    subjects_set: set[str] = set()

    for subj, course, sem, _grades, _inst in all_rows:
        semesters_set.add(sem)
        subjects_set.add(subj)

    semesters = sorted(semesters_set, key=sem_sort_key)
    subjects = sorted(subjects_set)
    subj_idx = {s: i for i, s in enumerate(subjects)}
    sem_idx = {s: i for i, s in enumerate(semesters)}

    compact_records: list[list] = []
    for subj, course, sem, grades, inst in all_rows:
        si = subj_idx[subj]
        mi = sem_idx[sem]
        compact_records.append([si, course, mi, grades, inst])

    raw = {
        "subjects": subjects,
        "semesters": semesters,
        "courseTitles": course_titles,
        "records": compact_records,
    }

    OUT_JSON.write_text(json.dumps(raw, separators=(",", ":")), encoding="utf-8")
    print(
        f"Wrote {OUT_JSON.name}: {len(compact_records)} records, "
        f"{len(subjects)} subjects, {len(semesters)} semesters"
    )


if __name__ == "__main__":
    main()
