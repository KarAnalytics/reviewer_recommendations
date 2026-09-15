"""Populate Handling editor and Reviewer 1/2(/3) on Submissions from the
optional Assignments and handlingEditor_assignment sheets -- data these
scripts already read for other purposes (exclusion, workload in
suggest_reviewers.py) but weren't writing back to Submissions until now.

Only fills currently-BLANK cells -- never overwrites anything already
there, manual or otherwise -- so it's safe to re-run, including after
co-chairs have started filling things in by hand.

- Handling editor: the name from handlingEditor_assignment (Member Name,
  Submission #) for that paper.
- Reviewer 1/2(/3) (however many "Reviewer N"-style columns exist -- see
  common.find_reviewer_slot_columns, so a renamed header like "Reviewer 3
  (use only two per paper)" still works -- that text is treated as a
  recommendation, not a hard limit; every available Reviewer-N column
  gets filled from actual data): from Assignments (#, Subreviewer "Name
  <email>", Status). For each paper, non-declined candidates are ranked
  "review added to easychair" > "accepted" > "submission accessed" >
  "submission not accessed" and fill blank Reviewer-N slots first,
  written as "Name (status)" so it's clear whether that's a confirmed
  reviewer or still a pending invitation. Anyone who *denied* is then
  written into any Reviewer-N slots still blank afterward, as
  "Name (denied)" with a pink cell fill -- visible at a glance so you
  don't accidentally re-invite them for another paper.

Usage:
    python sync_assignments.py
    python sync_assignments.py --dry-run     # show what would change, don't write
    python sync_assignments.py --workbook path\to\file.xlsx
"""
from __future__ import annotations

import argparse
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from common import (WORKBOOK_PATH, GOOGLE_SHEET_ID, load_workbook, load_assignments,
                     load_handling_editors, find_reviewer_slot_columns, set_cell_fill, start_logging)

SUB_SHEET = "Submissions"
DENIED_FILL = "FFC0CB"  # pink

_STATUS_RANK = {
    "review added to easychair": 0,
    "accepted": 1,
    "submission accessed": 2,
    "submission not accessed": 3,
}


def find_columns(ws, required: list[str]) -> dict[str, int]:
    header = {}
    for cell in ws[1]:
        if cell.value:
            header[str(cell.value).strip()] = cell.column
    missing = [c for c in required if c not in header]
    if missing:
        raise SystemExit(f"'{ws.title}' is missing expected column(s): {missing}")
    return header


def main() -> None:
    start_logging("sync_assignments")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--workbook", default=str(WORKBOOK_PATH) if WORKBOOK_PATH else None,
                     required=WORKBOOK_PATH is None and not GOOGLE_SHEET_ID,
                     help="path to the .xlsx (auto-detected if there's exactly one "
                          "next to reviewer_tools/; otherwise required, or set "
                          "WORKBOOK_PATH in .env)")
    ap.add_argument("--dry-run", action="store_true", help="show what would change, don't write anything")
    ap.add_argument("--max-reviewers", type=int, default=None,
                     help="only fill up to this many Reviewer-N slots per paper, even if more "
                          "Reviewer-N columns exist (default: fill all of them)")
    args = ap.parse_args()

    wb = load_workbook(args.workbook)
    if SUB_SHEET not in wb.sheetnames:
        raise SystemExit(f"No '{SUB_SHEET}' sheet in {args.workbook}")
    sub_ws = wb[SUB_SHEET]
    sub_col = find_columns(sub_ws, ["#", "Title", "paper", "Handling editor"])

    reviewer_slots = find_reviewer_slot_columns(sub_col)
    slot_cols = [reviewer_slots[k] for k in sorted(reviewer_slots, key=lambda s: int(s.split()[-1]))]

    editors = load_handling_editors(wb)
    assignment_records = load_assignments(wb)
    if not editors and not assignment_records:
        print("No handlingEditor_assignment or Assignments sheet found (or both are empty) "
              "-- nothing to sync.")
        return

    by_submission: dict[str, list[dict]] = {}
    for rec in assignment_records:
        by_submission.setdefault(rec["submission"], []).append(rec)

    editor_filled = 0
    reviewer_filled = 0
    for r in range(2, sub_ws.max_row + 1):
        title = sub_ws.cell(r, sub_col["Title"]).value
        if not title or not sub_ws.cell(r, sub_col["paper"]).value:
            continue
        sub_num = sub_ws.cell(r, sub_col["#"]).value
        if sub_num is None:
            continue
        key = str(sub_num).strip()

        he_col = sub_col["Handling editor"]
        if not sub_ws.cell(r, he_col).value and key in editors:
            value = editors[key]
            print(f"  #{key}: Handling editor -> {value!r}" + (" (dry-run)" if args.dry_run else ""))
            if not args.dry_run:
                sub_ws.cell(r, he_col).value = value
            editor_filled += 1

        records = by_submission.get(key, [])
        active = sorted((rec for rec in records if not rec["declined"]),
                         key=lambda rec: _STATUS_RANK.get(rec["status"], 99))
        denied = [rec for rec in records if rec["declined"]]
        if args.max_reviewers is not None:
            active = active[: args.max_reviewers]

        remaining_slots = [col for col in slot_cols if not sub_ws.cell(r, col).value]
        ai = di = 0
        for col in remaining_slots:
            if ai < len(active):
                rec = active[ai]
                ai += 1
                value = f"{rec['name']} ({rec['status']})"
                print(f"  #{key}: col {col} -> {value!r}" + (" (dry-run)" if args.dry_run else ""))
                if not args.dry_run:
                    sub_ws.cell(r, col).value = value
                reviewer_filled += 1
            elif di < len(denied):
                rec = denied[di]
                di += 1
                value = f"{rec['name']} (denied)"
                print(f"  #{key}: col {col} -> {value!r} [pink]" + (" (dry-run)" if args.dry_run else ""))
                if not args.dry_run:
                    sub_ws.cell(r, col).value = value
                    set_cell_fill(sub_ws, r, col, DENIED_FILL)
                reviewer_filled += 1

    print(f"\nHandling editor cells filled: {editor_filled}")
    print(f"Reviewer slot cells filled: {reviewer_filled}")

    if args.dry_run:
        print("\n--dry-run: nothing written.")
        return
    if editor_filled == 0 and reviewer_filled == 0:
        print("Nothing to write.")
        return

    wb.save(args.workbook)
    print(f"Saved {args.workbook}")


if __name__ == "__main__":
    main()
