# Reviewer recommendation tools

Three scripts for conference track chairs: one fills in each reviewer's
Position/Interests/Website from a web search, one uses that (plus each
paper's Keywords/Abstract) to suggest well-matched, non-conflicted, load-
balanced reviewers for every submission, and one keeps a live
"how many reviews is this person already on" count.

Works on any conference's spreadsheet as long as it matches the format
below -- nothing here is tied to a specific conference. Both scripts are
safe to re-run: they only fill in blank cells unless you pass `--force`,
so adding a handful of new papers/reviewers and re-running only does work
for the new rows.

Every run also writes its own timestamped, numbered log
(`logs/<script>_run<N>_<timestamp>.log`) to a `logs/` folder next to this
one -- kept outside `reviewer_tools/` and out of git (see `.gitignore` one
level up) since logs can contain reviewer PII and paper content.

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
| `No_reviews_assigned` | Output column, auto-created if missing -- `update_review_counts.py` (and `suggest_reviewers.py`, opportunistically) write a live formula here |

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
4. Copy `.env.example` to `.env`, then pick **one** of:
   - **Local Ollama (default, no key to manage)**: install the Ollama
     desktop app (<https://ollama.com/download>), sign in, pull a
     cloud-capable model (`ollama pull gemma4:31b-cloud`), and leave
     `LLM_PROVIDER=ollama` in `.env` as-is -- chat calls route through your
     signed-in app automatically.
   - **Bring your own key**: set `LLM_PROVIDER` to `anthropic`, `openai`,
     `groq`, `openrouter`, or `ollama_cloud` in `.env`, and paste your own
     API key for that provider into `LLM_API_KEY`. See "LLM providers"
     below for what each needs and what it doesn't cover.
5. Either way, `enrich_reviewers.py`'s web-*search* step (finding a
   reviewer's position/interests/website) needs its own key:
   `OLLAMA_API_KEY` from <https://ollama.com/settings/keys>, **unless**
   you're using `LLM_PROVIDER=anthropic` or `openai`, which search the web
   themselves and don't need it. See "LLM providers" below.
6. Optional, in `.env`: set `CONFERENCE_NAME` (e.g. `"MyConf 2027"`) and
   `CONFERENCE_FIELD` (e.g. `"human-computer interaction"`) to sharpen the
   model's name-collision check and sense of topical fit. Generic defaults
   work without these, just less precisely.

## LLM providers -- bring your own key

`LLM_PROVIDER` (+ `LLM_API_KEY`, `LLM_MODEL`) in `.env` controls who
handles **ranking/suggestion calls** (`suggest_reviewers.py`, and
`enrich_reviewers.py`'s summarizing step). Each person running this can
set their own -- there's no shared secret to manage for this part.

| `LLM_PROVIDER` | Needs `LLM_API_KEY`? | Notes |
|---|---|---|
| `ollama` (default) | No | Your locally-signed-in Ollama desktop app. |
| `ollama_cloud` | Yes | Ollama Cloud directly, no local app needed. |
| `anthropic` | Yes | `LLM_MODEL` default: `claude-sonnet-4-5`. |
| `openai` | Yes | `LLM_MODEL` default: `gpt-4o-mini`. |
| `groq` | Yes | `LLM_MODEL` default: `llama-3.3-70b-versatile`. |
| `openrouter` | Yes | `LLM_MODEL` default: `openrouter/auto`. |

**`enrich_reviewers.py`'s web-*search* step is a separate concern** from
the table above -- it's what actually finds a reviewer's info, not just
ranks candidates:

- `LLM_PROVIDER=ollama`/`ollama_cloud`: does the search in a dedicated
  call to Ollama Cloud's search API, then a second call summarizes the
  results. Needs `OLLAMA_API_KEY` for that search call **regardless** of
  which `LLM_PROVIDER` you set for ranking (Ollama Cloud's is the only
  standalone "give me raw search results" endpoint wired up here).
- `LLM_PROVIDER=anthropic`/`openai`: searches and extracts in a *single*
  call using that provider's own built-in web-search tool, authenticated
  with `LLM_API_KEY` -- no `OLLAMA_API_KEY` needed in this case.
- `groq`/`openrouter`: fine for ranking calls, but have no web-search-tool
  path built here, so `enrich_reviewers.py` will error if you pick one of
  these -- use `ollama`/`ollama_cloud`/`anthropic`/`openai` for enrichment.

The Anthropic/OpenAI web-search paths are newer and less exercised here
than the Ollama path (which has been run against 245 real reviewers) --
smoke-test with `--limit 5` before trusting either at scale.

Adding another OpenAI-wire-compatible provider (e.g. a different one you
already use) is a few lines in `common.py`'s `_OPENAI_COMPATIBLE_PROVIDERS`
dict -- no new function needed for the ranking-call side.

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
- Among candidates who are a *reasonable* topical fit (not among everyone
  indiscriminately), ranks by these priorities, in order:
  1. **Topical fit is a floor, not the top sort key.** A candidate has to
     plausibly connect to the paper's topic to be considered at all.
  2. **Primary: early-career reviewers first.** PhD students, postdocs,
     and assistant/associate professors are preferred over senior
     leadership/administrative roles (dean, department head/chair,
     director, president/CEO, named chair professorships) -- juniors
     typically respond to and complete review requests faster; people in
     those senior roles are usually stretched thin with editorial boards
     and other service. A senior/leadership candidate is only suggested
     when they're clearly and substantially the best topical match and no
     reasonably-fitting junior candidate covers the topic -- seniors are a
     fallback, not a coequal option.
  3. **Secondary tiebreak, among similar seniority + fit:** prefers
     reviewers currently assigned to **fewer** papers (real Reviewer 1/2/3
     assignments, read fresh from Submissions each run) -- spreads actual
     review workload instead of piling onto people already reviewing
     several papers.
  4. A clearly better topical match still wins even if it means a more
     senior or busier person -- priorities 2-3 break ties, they don't
     override an obviously better fit.

  **How "junior vs. senior" is actually decided:** there's no Python-side
  seniority score -- `Position` is passed to the model as raw text (e.g.
  `"Dean"`, `"Assistant Professor"`, `"PhD Candidate"`), and the model
  judges seniority and topical fit together in one call, guided by the
  prompt language above. That's a deliberate choice: it handles arbitrary
  or foreign job titles (e.g. *"Wissenschaftlicher Mitarbeiter"*, *"Vice
  Dean for Research"*) without needing an exhaustive keyword list, but it
  also means the exact ordering is inference-time behavior, not a rule
  you can point to in code -- spot-check suggestions periodically rather
  than assuming the ordering is guaranteed.
- Writes a numbered list (name + one-line reason) into `AISuggestedReviewers`.
- Caps how many *papers* any one reviewer can be suggested for across the
  whole run (default 5, `--max-per-reviewer`) -- once someone hits the
  cap they're dropped from the candidate pool for the rest of the run, so
  a handful of obviously well-matched people don't end up suggested for
  nearly every paper. The count is seeded from suggestions already sitting
  in `AISuggestedReviewers`, so the cap holds correctly even across
  repeated runs, not just within one. Pass `--max-per-reviewer 0` to
  disable it.
- Also refreshes `ReviewerList.No_reviews_assigned` (see below) on every
  run, so it stays in sync even if you never run `update_review_counts.py`
  directly.

```
python suggest_reviewers.py             # default: top 5 per paper
python suggest_reviewers.py --top-n 3
python suggest_reviewers.py --max-per-reviewer 8
python suggest_reviewers.py --limit 2   # try it on 2 papers first
```

Run this again whenever you add new papers -- it'll only compute
suggestions for the new rows. Run it with `--force` if you want to
recompute everyone (e.g. after re-running `enrich_reviewers.py` and getting
better interest data; this also fully recomputes the `--max-per-reviewer`
counts from scratch across all papers, rather than adding on top of what
was already there).

Note the cap is about how many papers a reviewer gets *suggested* for --
it's a different number from `No_reviews_assigned`, which counts how many
papers a reviewer is *actually assigned* to (Reviewer 1/2/3). Being
suggested doesn't use up review capacity; being assigned does.

## 3. `update_review_counts.py` -- keep No_reviews_assigned in sync

Writes a live Excel formula into each reviewer's `No_reviews_assigned`
cell: `=COUNTIF(Submissions!<Reviewer 1 col>,...)` summed across Reviewer
1/2/3. Because it's a formula, not a static number, it recalculates
*itself* in Excel the instant a track chair types a name into Reviewer
1/2/3 -- there's nothing to re-run after every assignment.

```
python update_review_counts.py
```

You only need to run this for first-time setup (adding the column) or
after adding new rows to `ReviewerList` (to fill the formula into them) --
`suggest_reviewers.py` also does this automatically on every run, so in
practice you rarely need to call it directly.

Note: `COUNTIF` does an exact (case-insensitive) text match, so a name
typed into Reviewer 1/2/3 needs to be spelled exactly as it appears in
`ReviewerList.Author` to be counted -- copying from `AISuggestedReviewers`
or `ReviewerList` keeps this accurate.

## Notes

- `common.py` holds shared logic (`.env` loading, workbook auto-detection,
  the multi-provider LLM calls, and the author-name parsing/matching) --
  no need to run it directly.
- These are AI-generated *suggestions* to speed up manual reviewer
  selection, not final assignments -- spot-check a few before relying on
  them, especially early on, and especially for reviewers with common
  names (name-collision mismatches are the main failure mode to watch for).
