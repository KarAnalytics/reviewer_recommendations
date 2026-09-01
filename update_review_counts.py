"""(Re)install the No_reviews_assigned column on ReviewerList.

This writes a live Excel COUNTIF formula per reviewer row -- it counts how
many times that person's name appears across Submissions' Reviewer 1,
Reviewer 2, and Reviewer 3 columns. Because it's a formula, not a static
number, the count updates itself in Excel the instant you type/change a
name in Reviewer 1/2/3 -- you do NOT need to re-run this after every
assignment.

You only need to (re)run this script when:
  - The column doesn't exist yet (first-time setup), or
  - You've added new rows to ReviewerList and want the formula filled
    down into them.

Note: COUNTIF does an exact (case-insensitive) text match, so a reviewer's
name in Reviewer 1/2/3 needs to be spelled exactly as it appears in
ReviewerList's Author column for the count to pick it up -- copying from
AISuggestedReviewers or ReviewerList itself keeps this accurate.

Usage:
    python update_review_counts.py
    python update_review_counts.py --workbook path\to\file.xlsx
"""
from __future__ import annotations

import argparse

from common import WORKBOOK_PATH, GOOGLE_SHEET_ID, REVIEW_COUNT_COLUMN, load_workbook, sync_review_counts, start_logging


def main() -> None:
    start_logging("update_review_counts")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--workbook", default=str(WORKBOOK_PATH) if WORKBOOK_PATH else None,
                     required=WORKBOOK_PATH is None and not GOOGLE_SHEET_ID,
                     help="path to the .xlsx (auto-detected if there's exactly one "
                          "next to reviewer_tools/; otherwise required, or set "
                          "WORKBOOK_PATH in .env)")
    args = ap.parse_args()

    wb = load_workbook(args.workbook)
    n = sync_review_counts(wb)
    if n == 0:
        raise SystemExit(
            "Nothing written -- check that the workbook has a 'Submissions' sheet with "
            "Reviewer 1/Reviewer 2/Reviewer 3 columns and a 'ReviewerList' sheet with an "
            "Author column."
        )

    wb.save(args.workbook)
    print(f"Wrote live '{REVIEW_COUNT_COLUMN}' formulas for {n} reviewer(s).")
    print("This now updates automatically in Excel whenever Reviewer 1/2/3 change --")
    print("re-run this script only after adding new rows to ReviewerList.")
    print(f"Saved {args.workbook}")


if __name__ == "__main__":
    main()
