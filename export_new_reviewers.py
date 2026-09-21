r"""Compare our ReviewerList against an external EasyChair reviewer-pool
export (Name, Email, ... columns) and write everyone who's in ours but
not in theirs into a new tab in that same EasyChair file -- ready to hand
back to EasyChair to register as reviewers/PC members.

Matches by email (exact, case-insensitive) first, then by fuzzy name
match (same logic used everywhere else in these scripts), so formatting
differences between the two sources don't produce false "missing"
entries.

The new tab is fully replaced each run (not appended to), so it's safe
to re-run against a fresher EasyChair export later -- it always reflects
the current comparison, not an accumulation of past runs.

Usage:
    python export_new_reviewers.py --easychair-file "C:\path\to\EasyChair_reviewer_pool.xlsx"
    python export_new_reviewers.py --easychair-file ... --dry-run
    python export_new_reviewers.py --easychair-file ... --sheet-name "New Reviewers"
    python export_new_reviewers.py --easychair-file ... --workbook path\to\our_file.xlsx
"""
from __future__ import annotations

import argparse
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import openpyxl

from common import WORKBOOK_PATH, GOOGLE_SHEET_ID, load_workbook, names_match, start_logging

REV_SHEET = "ReviewerList"
DEFAULT_NEW_SHEET_NAME = "New Reviewers"


def find_columns(ws, required: list[str]) -> dict[str, int]:
    header = {}
    for cell in ws[1]:
        if cell.value:
            header[str(cell.value).strip()] = cell.column
    missing = [c for c in required if c not in header]
    if missing:
        raise SystemExit(f"'{ws.title}' is missing expected column(s): {missing}")
    return header


def load_easychair_pool(path: str):
    wb = openpyxl.load_workbook(path)
    ws = wb.worksheets[0]  # the pool listing -- first sheet, whatever it's named
    header = {str(c.value).strip(): c.column for c in ws[1] if c.value}
    if "Name" not in header or "Email" not in header:
        raise SystemExit(
            f"{path} doesn't look like an EasyChair reviewer pool export (expected "
            f"Name/Email columns on its first sheet, found {list(header.keys())})."
        )
    names, emails = [], set()
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, header["Name"]).value
        email = ws.cell(r, header["Email"]).value
        if not name:
            continue
        names.append(str(name).strip())
        if email:
            emails.add(str(email).strip().lower())
    return wb, names, emails


def main() -> None:
    start_logging("export_new_reviewers")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--easychair-file", required=True,
                     help="path to the EasyChair reviewer-pool .xlsx (Name, Email, ... columns)")
    ap.add_argument("--workbook", default=str(WORKBOOK_PATH) if WORKBOOK_PATH else None,
                     required=WORKBOOK_PATH is None and not GOOGLE_SHEET_ID,
                     help="path to OUR .xlsx (auto-detected if there's exactly one next to "
                          "reviewer_tools/; otherwise required, or set WORKBOOK_PATH in .env) "
                          "-- ignored if GOOGLE_SHEET_ID is set")
    ap.add_argument("--sheet-name", default=DEFAULT_NEW_SHEET_NAME,
                     help=f"name for the new tab written into the EasyChair file "
                          f"(default {DEFAULT_NEW_SHEET_NAME!r})")
    ap.add_argument("--dry-run", action="store_true", help="show who would be added, don't write anything")
    args = ap.parse_args()

    wb = load_workbook(args.workbook)
    if REV_SHEET not in wb.sheetnames:
        raise SystemExit(f"No '{REV_SHEET}' sheet found.")
    rev_ws = wb[REV_SHEET]
    rev_col = find_columns(rev_ws, ["Author", "Email", "Affiliation"])

    ec_wb, ec_names, ec_emails = load_easychair_pool(args.easychair_file)
    print(f"EasyChair pool ({args.easychair_file}): {len(ec_names)} reviewers.")

    our_reviewers = []
    for r in range(2, rev_ws.max_row + 1):
        name = rev_ws.cell(r, rev_col["Author"]).value
        if not name or not str(name).strip():
            continue
        our_reviewers.append({
            "name": str(name).strip(),
            "email": str(rev_ws.cell(r, rev_col["Email"]).value or "").strip(),
            "affiliation": str(rev_ws.cell(r, rev_col["Affiliation"]).value or "").strip(),
            "position": str(rev_ws.cell(r, rev_col["Position"]).value or "").strip() if "Position" in rev_col else "",
            "interests": str(rev_ws.cell(r, rev_col["Interests"]).value or "").strip() if "Interests" in rev_col else "",
        })
    print(f"Our ReviewerList: {len(our_reviewers)} reviewers.")

    missing = []
    for rv in our_reviewers:
        if rv["email"] and rv["email"].lower() in ec_emails:
            continue
        if any(names_match(rv["name"], n) for n in ec_names):
            continue
        missing.append(rv)

    print(f"\nMissing from EasyChair pool: {len(missing)}")
    for m in missing:
        extra = f" -- {m['affiliation']}" if m["affiliation"] else ""
        print(f"  + {m['name']} ({m['email'] or 'no email'}){extra}")

    if args.dry_run:
        print("\n--dry-run: nothing written.")
        return
    if not missing:
        print("Nothing to write.")
        return

    if args.sheet_name in ec_wb.sheetnames:
        del ec_wb[args.sheet_name]
    new_ws = ec_wb.create_sheet(args.sheet_name)
    new_ws.append(["Name", "Email", "Affiliation", "Position", "Interests"])
    for m in missing:
        new_ws.append([m["name"], m["email"], m["affiliation"], m["position"], m["interests"]])

    ec_wb.save(args.easychair_file)
    print(f"\nWrote {len(missing)} reviewer(s) to sheet {args.sheet_name!r} in {args.easychair_file}")


if __name__ == "__main__":
    main()
