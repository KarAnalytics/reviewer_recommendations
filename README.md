# Reviewer recommendation tools

Two scripts for conference track chairs: one fills in each reviewer's
Position/Interests/Website from a web search, the other uses that (plus
each paper's Keywords/Abstract) to suggest well-matched, non-conflicted
reviewers for every submission.

Works on any conference's spreadsheet as long as it matches the format
below -- nothing here is tied to a specific conference. Both scripts are
safe to re-run: they only fill in blank cells unless you pass `--force`,
so adding a handful of new papers/reviewers and re-running only does work
for the new rows.

## Input spreadsheet format

A single `.xlsx` workbook with two sheets (exact names, case-sensitive):

**`Submissions`** -- one row per paper. Required columns:

| Column | Used for |
|---|---|
| `Authors` | Free text, e.g. `"A. Smith, B. Jones and C. Lee"` -- parsed to exclude co-authors from their own paper's suggestions |
| `Title` | Shown to the model; also used to detect real submission rows (blank Title = skipped) |
| `paper` | Any truthy value (e.g. `✔`) marks a row as an actual submission to process; blank/0 rows are skipped (e.g. withdrawn papers) |
| `Keywords` | Fed to the model for topical matching |
| `Abstract` | Fed to the model for topical matching |
| `Reviewer 1`, `Reviewer 2`, `Reviewer 3` | If already filled in, those names are also excluded from that paper's suggestions |
| `AISuggestedReviewers` | Output column -- `suggest_reviewers.py` writes here |

**`ReviewerList`** -- one row per candidate reviewer. Required columns:

| Column | Used for |
|---|---|
| `Author` | The reviewer's name -- must match how they appear in `Submissions.Authors` closely enough for name matching (exact, or same last name + first initial) |
| `Affiliation` | Fed into the web-search query and shown to the model |
| `Email` | Optional context shown to the model (helps disambiguate) |
| `Position` | Output column -- `enrich_reviewers.py` writes here |
| `Interests` | Output column -- `enrich_reviewers.py` writes here |
| `Website` | Output column -- `enrich_reviewers.py` writes here |

Extra columns anywhere are ignored, so you can keep whatever else you
already track (notes, review-request dates, etc.).

## One-time setup

1. Put your workbook (`.xlsx`) in the folder next to `reviewer_tools/`. If
   it's the only `.xlsx` there, both scripts find it automatically;
   otherwise pass `--workbook path\to\file.xlsx` or set `WORKBOOK_PATH` in
   `.env`.
2. **Close the file in Excel** before running either script (Excel locks
   it for writing, so saving will fail while it's open).
3. Install the two extra packages (openpyxl, requests):
   ```
   pip install -r requirements.txt
   ```
4. Install the **Ollama** desktop app (<https://ollama.com/download>),
   sign in, and pull a cloud-capable model, e.g.:
   ```
   ollama pull gemma4:31b-cloud
   ```
   Once signed in, chat calls from these scripts route through your local
   app with no separate API key needed.
5. Copy `.env.example` to `.env` and paste in an API key from
   <https://ollama.com/settings/keys> (sign in with the same account
   Ollama is signed in with). This key is only needed for
   `enrich_reviewers.py`'s web searches -- `suggest_reviewers.py` just
   talks to your already-signed-in local Ollama app and needs no key.
6. Optional, in `.env`: set `CONFERENCE_NAME` (e.g. `"MyConf 2027"`) and
   `CONFERENCE_FIELD` (e.g. `"human-computer interaction"`) to sharpen the
   model's name-collision check and sense of topical fit. Generic defaults
   work without these, just less precisely.

## 1. `enrich_reviewers.py` -- fill in Position / Interests / Website

For every reviewer on the `ReviewerList` sheet with all three columns blank,
this web-searches their name + affiliation, then asks the model to extract
a job title, a short list of research interests, and a best-guess website
(personal/faculty page > Google Scholar > LinkedIn), and writes those into
the sheet. It never invents info that isn't backed by a search result --
rows it can't confidently identify get `Position = "Unknown"` instead of a
guess, and it's specifically instructed to watch for name collisions
(a search result about a different person who happens to share the name).

```
python enrich_reviewers.py --limit 5    # try it on 5 people first
python enrich_reviewers.py              # then do everyone left
python enrich_reviewers.py --only "Jane Doe,John Smith"  # redo just these people
```

Saves progress after every row, so it's safe to stop and resume. Ollama
Cloud's web-search API has a rate limit/quota; if you hit it the script
stops with a clear message instead of hanging -- just re-run it a bit
later (or the next day, if it's a daily quota) and it'll pick up where it
left off.

## 2. `suggest_reviewers.py` -- populate AISuggestedReviewers

For every submitted paper with a blank `AISuggestedReviewers` cell, this:

- Parses the `Authors` field and excludes anyone on that list (plus anyone
  already in `Reviewer 1/2/3`) from the candidate pool -- co-authors are
  never suggested to review their own paper.
- Sends the paper's Title/Keywords/Abstract plus the remaining candidates'
  Position/Interests to the model and asks it to rank the best-fit
  reviewers, choosing only from that candidate list (so it can't suggest
  someone who isn't actually in ReviewerList).
- For any candidate who is themselves an author on a *different* submitted
  paper, that paper's title/keywords are pulled in live from the
  Submissions sheet and given to the model as an extra signal -- their own
  current submission is a much stronger, fresher indicator of expertise
  than a web-search bio blurb, especially for reviewers whose
  Position/Interests came back blank/"Unknown". This is computed fresh
  each run, not stored in ReviewerList.
- As a *tiebreaker* among comparably good topical matches, it prefers
  reviewers more likely to actually respond -- PhD students, postdocs,
  assistant/associate professors -- over senior leadership/administrative
  roles (dean, department head/chair, director, president/CEO, named
  chair professorships), who tend to be slower to respond to review
  requests. Topical fit always comes first; a senior person who's clearly
  the best topical match still gets suggested.
- Writes a numbered list (name + one-line reason) into `AISuggestedReviewers`.

```
python suggest_reviewers.py             # default: top 5 per paper
python suggest_reviewers.py --top-n 3
python suggest_reviewers.py --limit 2   # try it on 2 papers first
```

Run this again whenever you add new papers -- it'll only compute
suggestions for the new rows. Run it with `--force` if you want to
recompute everyone (e.g. after re-running `enrich_reviewers.py` and getting
better interest data).

## Notes

- `common.py` holds shared logic (`.env` loading, workbook auto-detection,
  the Ollama calls, and the author-name parsing/matching) -- no need to
  run it directly.
- These are AI-generated *suggestions* to speed up manual reviewer
  selection, not final assignments -- spot-check a few before relying on
  them, especially early on, and especially for reviewers with common
  names (name-collision mismatches are the main failure mode to watch for).
