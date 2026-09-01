"""Populate the AISuggestedReviewers column on the Submissions sheet: for
each paper, ask the local Ollama model to pick the best-matching reviewers
from ReviewerList, automatically excluding the paper's own authors (and
anyone already entered as Reviewer 1/2/3 for that paper).

Usage:
    python suggest_reviewers.py                  # fill blank AISuggestedReviewers cells only
    python suggest_reviewers.py --force           # recompute every paper's suggestions
    python suggest_reviewers.py --top-n 5         # how many reviewers to suggest (default 5)
    python suggest_reviewers.py --workbook path\to\file.xlsx

Run this any time you add new papers (or re-run enrich_reviewers.py) --
it only (re)computes rows that need it, so it's safe/cheap to re-run.

Requires the Ollama app running locally with the model set in .env
(OLLAMA_MODEL, default gemma4:31b-cloud). No OLLAMA_API_KEY needed for
this script -- that's only for enrich_reviewers.py's web search.
"""
from __future__ import annotations

import argparse
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import openpyxl

from common import WORKBOOK_PATH, CONFERENCE_NAME, ollama_chat, extract_json, split_authors, names_match

SUB_SHEET = "Submissions"
REV_SHEET = "ReviewerList"
SAVE_EVERY = 5


def find_columns(ws, required: list[str]) -> dict[str, int]:
    header = {}
    for cell in ws[1]:
        if cell.value:
            header[str(cell.value).strip()] = cell.column
    missing = [c for c in required if c not in header]
    if missing:
        raise SystemExit(f"'{ws.title}' is missing expected column(s): {missing}")
    return header


def load_reviewer_pool(ws, col: dict[str, int]) -> list[dict]:
    pool = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, col["Author"]).value
        if not name or not str(name).strip():
            continue
        pool.append({
            "name": str(name).strip(),
            "position": str(ws.cell(r, col["Position"]).value or "").strip(),
            "interests": str(ws.cell(r, col["Interests"]).value or "").strip(),
            "affiliation": str(ws.cell(r, col["Affiliation"]).value or "").strip(),
        })
    return pool


def build_own_submission_index(sub_ws, sub_col: dict[str, int]) -> list[dict]:
    """Every submitted paper with its row + authors, so we can tell a
    candidate reviewer "you yourself submitted a paper on X" -- a much
    stronger interest signal than a bio blurb, and it's pulled fresh from
    Submissions each run rather than baked into ReviewerList.
    """
    papers = []
    for r in range(2, sub_ws.max_row + 1):
        title = sub_ws.cell(r, sub_col["Title"]).value
        if not title or not str(title).strip():
            continue
        if not sub_ws.cell(r, sub_col["paper"]).value:
            continue
        papers.append({
            "row": r,
            "title": str(title).strip(),
            "keywords": str(sub_ws.cell(r, sub_col["Keywords"]).value or "").strip(),
            "authors": split_authors(str(sub_ws.cell(r, sub_col["Authors"]).value or "")),
        })
    return papers


def own_papers_for(name: str, all_papers: list[dict], exclude_row: int) -> list[dict]:
    return [
        p for p in all_papers
        if p["row"] != exclude_row and any(names_match(name, a) for a in p["authors"])
    ]


def pick_reviewers(title: str, keywords: str, abstract: str,
                    candidates: list[dict], top_n: int) -> list[str]:
    if not candidates:
        return []

    lines = []
    for i, c in enumerate(candidates):
        line = (f"{i+1}. {c['name']} -- {c['position'] or 'position unknown'} "
                f"({c['affiliation'] or 'affiliation unknown'}). "
                f"Interests: {c['interests'] or 'unknown'}")
        if c.get("own_papers"):
            own = "; ".join(
                f"\"{p['title']}\" (keywords: {p['keywords'] or 'n/a'})"
                for p in c["own_papers"]
            )
            line += f" | Also a {CONFERENCE_NAME} author, on: {own}"
        lines.append(line)
    listing = "\n".join(lines)

    prompt = f"""You are helping a conference track co-chair assign peer reviewers.

PAPER
Title: {title}
Keywords: {keywords}
Abstract: {abstract}

CANDIDATE REVIEWERS (numbered list; you may ONLY choose from this list):
{listing}

Pick the {top_n} candidates whose research interests/position best match this
paper's topic. Topical fit comes first. When a candidate's stated Interests
are unknown/thin but they are listed as a {CONFERENCE_NAME} author on a paper
covering a related topic, treat that paper's title/keywords as a strong
signal of their expertise too.

Among candidates with comparably good topical fit, prefer ones more likely
to actually respond to and complete a review request -- PhD students,
postdocs, and assistant/associate professors -- over senior
leadership/administrative roles (dean, vice dean, department head/chair,
director, provost, president, CEO, or a named/distinguished/endowed chair
professorship), who tend to be busier and slower to respond. Do not let
this seniority preference override a clearly better topical match, and
don't penalize a senior person who is genuinely the best topical fit --
use it only as a tiebreaker.

Reply with ONLY a JSON array of exactly {top_n} objects (fewer only if the
list has fewer candidates), ranked best-fit first:
[{{"name": "<exact name as listed above>", "reason": "<one short clause, <12 words>"}}]"""

    reply = ollama_chat([{"role": "user", "content": prompt}])
    try:
        data = extract_json(reply)
    except Exception:
        return []

    by_name = {c["name"].lower(): c["name"] for c in candidates}
    picks = []
    for item in data:
        if not isinstance(item, dict):
            continue
        name = str(item.get("name", "")).strip()
        canonical = by_name.get(name.lower())
        if not canonical:
            # loose fallback match in case the model paraphrased slightly
            for cand in candidates:
                if names_match(name, cand["name"]):
                    canonical = cand["name"]
                    break
        if canonical and canonical not in [p[0] for p in picks]:
            picks.append((canonical, str(item.get("reason", "")).strip()))

    return picks[:top_n]


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--workbook", default=str(WORKBOOK_PATH) if WORKBOOK_PATH else None,
                     required=WORKBOOK_PATH is None,
                     help="path to the .xlsx (auto-detected if there's exactly one "
                          "next to reviewer_tools/; otherwise required, or set "
                          "WORKBOOK_PATH in .env)")
    ap.add_argument("--top-n", type=int, default=5)
    ap.add_argument("--force", action="store_true", help="recompute every paper, not just blank ones")
    ap.add_argument("--limit", type=int, default=None, help="only process the first N eligible papers")
    args = ap.parse_args()

    wb = openpyxl.load_workbook(args.workbook)
    for sheet in (SUB_SHEET, REV_SHEET):
        if sheet not in wb.sheetnames:
            raise SystemExit(f"No '{sheet}' sheet in {args.workbook}")

    sub_ws = wb[SUB_SHEET]
    rev_ws = wb[REV_SHEET]

    sub_col = find_columns(sub_ws, [
        "Authors", "Title", "paper", "Keywords", "Abstract",
        "Reviewer 1", "Reviewer 2", "Reviewer 3", "AISuggestedReviewers",
    ])
    rev_col = find_columns(rev_ws, ["Author", "Affiliation", "Position", "Interests"])

    pool = load_reviewer_pool(rev_ws, rev_col)
    if not pool:
        raise SystemExit("ReviewerList has no reviewers to suggest from.")

    all_papers = build_own_submission_index(sub_ws, sub_col)

    todo = []
    for r in range(2, sub_ws.max_row + 1):
        title = sub_ws.cell(r, sub_col["Title"]).value
        if not title or not str(title).strip():
            continue
        paper_flag = sub_ws.cell(r, sub_col["paper"]).value
        if not paper_flag:
            continue  # withdrawn / not a real submission
        existing = sub_ws.cell(r, sub_col["AISuggestedReviewers"]).value
        if existing and not args.force:
            continue
        todo.append(r)

    if args.limit:
        todo = todo[: args.limit]

    if not todo:
        print("Nothing to do -- every submitted paper already has AISuggestedReviewers filled in.")
        print("(Pass --force to recompute everyone anyway.)")
        return

    print(f"Suggesting reviewers for {len(todo)} paper(s)...")
    t0 = time.time()
    for i, r in enumerate(todo, 1):
        title = str(sub_ws.cell(r, sub_col["Title"]).value or "").strip()
        keywords = str(sub_ws.cell(r, sub_col["Keywords"]).value or "").strip()
        abstract = str(sub_ws.cell(r, sub_col["Abstract"]).value or "").strip()
        authors_field = sub_ws.cell(r, sub_col["Authors"]).value or ""
        paper_authors = split_authors(str(authors_field))

        already_assigned = [
            str(sub_ws.cell(r, sub_col[c]).value).strip()
            for c in ("Reviewer 1", "Reviewer 2", "Reviewer 3")
            if sub_ws.cell(r, sub_col[c]).value
        ]

        excluded_names = paper_authors + already_assigned
        candidates = []
        for c in pool:
            if any(names_match(c["name"], ex) for ex in excluded_names):
                continue
            c = dict(c, own_papers=own_papers_for(c["name"], all_papers, exclude_row=r))
            candidates.append(c)

        print(f"[{i}/{len(todo)}] #{sub_ws.cell(r, sub_col['Title']).value and r} {title[:60]!r} "
              f"({len(candidates)} eligible reviewers) ...", end=" ", flush=True)

        try:
            picks = pick_reviewers(title, keywords, abstract, candidates, args.top_n)
        except Exception as exc:  # noqa: BLE001
            print(f"error ({exc}), skipping")
            continue

        if not picks:
            print("no suggestions returned")
            continue

        cell_text = "\n".join(
            f"{j}. {name}" + (f" -- {reason}" if reason else "")
            for j, (name, reason) in enumerate(picks, 1)
        )
        sub_ws.cell(r, sub_col["AISuggestedReviewers"]).value = cell_text
        print(f"{len(picks)} suggested")

        if i % SAVE_EVERY == 0:
            wb.save(args.workbook)

    wb.save(args.workbook)
    elapsed = time.time() - t0
    print(f"Done. Saved {args.workbook} ({elapsed:.0f}s, {elapsed/max(len(todo),1):.1f}s/paper avg).")


if __name__ == "__main__":
    main()
