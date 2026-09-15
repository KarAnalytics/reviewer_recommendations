"""Add this year's submission authors to ReviewerList as candidate
reviewers, pulling from an Authorlist sheet (Submission #, First name,
Last name, Email, Country, Affiliation, Web page, Person #,
Corresponding?) if the workbook has one.

For each unique person in Authorlist not already in ReviewerList (fuzzy
name match -- same logic used everywhere else in these scripts, honorifics
like "Dr." ignored), adds a new ReviewerList row with Author/Email/
Affiliation prefilled from Authorlist (and Website too, if Authorlist's
Web page column has one) -- Position/Interests are left blank for
enrich_reviewers.py to fill in as usual.

Safe to re-run: only adds people who are still missing, so running this
again after adding more papers/authors only adds the new ones.

Usage:
    python import_authors.py
    python import_authors.py --dry-run     # show what would be added, don't write
    python import_authors.py --workbook path\to\file.xlsx
"""
from __future__ import annotations

import argparse
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

from common import WORKBOOK_PATH, GOOGLE_SHEET_ID, load_workbook, names_match, start_logging

AUTHORLIST_SHEET = "Authorlist"
REV_SHEET = "ReviewerList"


def find_columns(ws, required: list[str]) -> dict[str, int]:
    header = {}
    for cell in ws[1]:
        if cell.value:
            header[str(cell.value).strip()] = cell.column
    missing = [c for c in required if c not in header]
    if missing:
        raise SystemExit(f"'{ws.title}' is missing expected column(s): {missing}")
    return header


def collect_authors(al_ws, al_col: dict[str, int]) -> list[dict]:
    """One entry per unique person in Authorlist -- the same person can
    appear on multiple submissions (or with inconsistent name formatting,
    e.g. "Hadi Karimikia" vs. "Dr. Hadi Karimikia"), so this dedupes by
    exact email match first (a stronger signal when available), then by
    fuzzy name match, preferring a row marked Corresponding when one
    exists for that person."""
    result: list[dict] = []
    for r in range(2, al_ws.max_row + 1):
        first = al_ws.cell(r, al_col["First name"]).value
        last = al_ws.cell(r, al_col["Last name"]).value
        if not first and not last:
            continue
        full_name = " ".join(p for p in (str(first or "").strip(), str(last or "").strip()) if p)
        if not full_name:
            continue
        entry = {
            "name": full_name,
            "email": str(al_ws.cell(r, al_col["Email"]).value or "").strip(),
            "affiliation": al_ws.cell(r, al_col["Affiliation"]).value or "",
            "website": al_ws.cell(r, al_col["Web page"]).value if "Web page" in al_col else "",
            "corresponding": bool(al_ws.cell(r, al_col["Corresponding?"]).value) if "Corresponding?" in al_col else False,
        }

        match_idx = None
        for i, existing in enumerate(result):
            same_email = entry["email"] and entry["email"].lower() == existing["email"].lower()
            if same_email or names_match(entry["name"], existing["name"]):
                match_idx = i
                break

        if match_idx is None:
            result.append(entry)
        elif entry["corresponding"] and not result[match_idx]["corresponding"]:
            result[match_idx] = entry

    return result


def main() -> None:
    start_logging("import_authors")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--workbook", default=str(WORKBOOK_PATH) if WORKBOOK_PATH else None,
                     required=WORKBOOK_PATH is None and not GOOGLE_SHEET_ID,
                     help="path to the .xlsx (auto-detected if there's exactly one "
                          "next to reviewer_tools/; otherwise required, or set "
                          "WORKBOOK_PATH in .env)")
    ap.add_argument("--dry-run", action="store_true", help="show what would be added, don't write anything")
    args = ap.parse_args()

    wb = load_workbook(args.workbook)
    if AUTHORLIST_SHEET not in wb.sheetnames:
        raise SystemExit(f"No '{AUTHORLIST_SHEET}' sheet in this workbook -- nothing to import.")
    if REV_SHEET not in wb.sheetnames:
        raise SystemExit(f"No '{REV_SHEET}' sheet in this workbook.")

    al_ws = wb[AUTHORLIST_SHEET]
    rev_ws = wb[REV_SHEET]

    al_col = find_columns(al_ws, ["First name", "Last name", "Email", "Affiliation"])
    rev_col = find_columns(rev_ws, ["Author", "Email", "Affiliation"])
    # Website/Position/Interests/# may or may not already exist as columns.
    rev_header = {str(c.value).strip(): c.column for c in rev_ws[1] if c.value}

    existing_names = []
    max_num = 0
    for r in range(2, rev_ws.max_row + 1):
        name = rev_ws.cell(r, rev_col["Author"]).value
        if name and str(name).strip():
            existing_names.append(str(name).strip())
        if "#" in rev_header:
            n = rev_ws.cell(r, rev_header["#"]).value
            if isinstance(n, (int, float)):
                max_num = max(max_num, int(n))

    authors = collect_authors(al_ws, al_col)
    to_add = [
        a for a in authors
        if not any(names_match(a["name"], existing) for existing in existing_names)
    ]

    print(f"Authorlist: {len(authors)} unique people. Already in ReviewerList: "
          f"{len(authors) - len(to_add)}. New: {len(to_add)}.")

    if not to_add:
        print("Nothing to add.")
        return

    for a in to_add:
        print(f"  + {a['name']}" + (f" ({a['affiliation']})" if a["affiliation"] else ""))

    if args.dry_run:
        print("\n--dry-run: nothing written.")
        return

    next_row = rev_ws.max_row + 1
    for a in to_add:
        r = next_row
        if "#" in rev_header:
            max_num += 1
            rev_ws.cell(r, rev_header["#"]).value = max_num
        rev_ws.cell(r, rev_col["Author"]).value = a["name"]
        rev_ws.cell(r, rev_col["Email"]).value = a["email"]
        rev_ws.cell(r, rev_col["Affiliation"]).value = a["affiliation"]
        if a["website"] and "Website" in rev_header:
            rev_ws.cell(r, rev_header["Website"]).value = a["website"]
        next_row += 1

    wb.save(args.workbook)
    print(f"\nAdded {len(to_add)} new reviewer candidate(s). Saved.")
    print("Next: run enrich_reviewers.py to fill in Position/Interests for the new rows.")


if __name__ == "__main__":
    main()
