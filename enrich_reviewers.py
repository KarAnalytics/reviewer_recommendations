"""Fill in Position / Interests / Website for each person on the
ReviewerList sheet, using web search + an LLM to summarize what it finds.

Provider is picked via LLM_PROVIDER in .env (default 'ollama' -- your
locally-signed-in Ollama app, no key needed). 'ollama'/'ollama_cloud' do a
two-step search-then-summarize using Ollama Cloud's search API (needs
OLLAMA_API_KEY regardless of LLM_PROVIDER, since it's the only standalone
search endpoint wired up here); 'anthropic'/'openai' do it in one call via
that provider's own built-in web-search tool, using LLM_API_KEY. See
common.research_person() for details.

Usage:
    python enrich_reviewers.py                 # fill every blank row
    python enrich_reviewers.py --limit 5        # just the first 5 (dry run)
    python enrich_reviewers.py --force          # re-look-up EVERY row, even
                                                  # ones already filled in
    python enrich_reviewers.py --only "Pei Siang Goh,Xiaobai Li"  # redo just these people
    python enrich_reviewers.py --workbook path\to\file.xlsx
"""
from __future__ import annotations

import argparse
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import openpyxl

from common import WORKBOOK_PATH, research_person, start_logging

SHEET = "ReviewerList"
SAVE_EVERY = 1  # autosave cadence, so a crash mid-run doesn't lose progress


def find_columns(ws) -> dict[str, int]:
    header = {}
    for cell in ws[1]:
        if cell.value:
            header[str(cell.value).strip()] = cell.column
    required = ["Author", "Affiliation", "Email", "Position", "Interests", "Website"]
    missing = [c for c in required if c not in header]
    if missing:
        raise SystemExit(f"ReviewerList is missing expected column(s): {missing}")
    return header


def main() -> None:
    start_logging("enrich_reviewers")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--workbook", default=str(WORKBOOK_PATH) if WORKBOOK_PATH else None,
                     required=WORKBOOK_PATH is None,
                     help="path to the .xlsx (auto-detected if there's exactly one "
                          "next to reviewer_tools/; otherwise required, or set "
                          "WORKBOOK_PATH in .env)")
    ap.add_argument("--limit", type=int, default=None, help="only process the first N eligible rows")
    ap.add_argument("--force", action="store_true", help="re-look-up rows even if already filled in")
    ap.add_argument("--only", default=None,
                     help="comma-separated exact reviewer name(s) to (re-)look-up, "
                          "ignoring the already-filled-in check")
    args = ap.parse_args()
    only_names = None
    if args.only:
        only_names = {n.strip().lower() for n in args.only.split(",") if n.strip()}

    wb = openpyxl.load_workbook(args.workbook)
    if SHEET not in wb.sheetnames:
        raise SystemExit(f"No '{SHEET}' sheet in {args.workbook}")
    ws = wb[SHEET]
    col = find_columns(ws)

    todo = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, col["Author"]).value
        if not name or not str(name).strip():
            continue
        if only_names is not None:
            if str(name).strip().lower() in only_names:
                todo.append(r)
            continue
        pos = ws.cell(r, col["Position"]).value
        interests = ws.cell(r, col["Interests"]).value
        website = ws.cell(r, col["Website"]).value
        already_done = any(v not in (None, "") for v in (pos, interests, website))
        if already_done and not args.force:
            continue
        todo.append(r)

    if args.limit:
        todo = todo[: args.limit]

    if not todo:
        print("Nothing to do -- every reviewer already has Position/Interests/Website filled in.")
        print("(Pass --force to re-look-up everyone anyway.)")
        return

    print(f"Looking up {len(todo)} reviewer(s)...")
    t0 = time.time()
    for i, r in enumerate(todo, 1):
        name = str(ws.cell(r, col["Author"]).value).strip()
        affiliation = ws.cell(r, col["Affiliation"]).value or ""
        email = ws.cell(r, col["Email"]).value or ""

        print(f"[{i}/{len(todo)}] {name} ...", end=" ", flush=True)
        try:
            info = research_person(name, str(affiliation), str(email))
        except RuntimeError as exc:
            print(f"\nFATAL: {exc}")
            sys.exit(1)
        except Exception as exc:  # noqa: BLE001
            print(f"error ({exc}), skipping")
            continue

        ws.cell(r, col["Position"]).value = info["position"]
        ws.cell(r, col["Interests"]).value = info["interests"]
        ws.cell(r, col["Website"]).value = info["website"]
        print(info["position"] or "-")

        if i % SAVE_EVERY == 0:
            wb.save(args.workbook)

    wb.save(args.workbook)
    elapsed = time.time() - t0
    print(f"Done. Saved {args.workbook} ({elapsed:.0f}s, {elapsed/max(len(todo),1):.1f}s/reviewer avg).")


if __name__ == "__main__":
    main()
