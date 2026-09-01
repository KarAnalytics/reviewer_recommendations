# SIGDSA reviewer tools

Two scripts that work on `..\SIGDSA26_2026-09-01_ReviewerList.xlsx` in place.
Both are safe to re-run: they only fill in blank cells unless you pass
`--force`, so adding a handful of new papers/reviewers and re-running only
does work for the new rows.

## One-time setup

1. **Close the Excel file** before running either script (Excel locks it for
   writing, so `wb.save()` will fail while it's open).
2. Install the two extra packages (openpyxl, requests):
   ```
   pip install -r requirements.txt
   ```
3. Make sure the **Ollama** desktop app is running and signed in (it already
   is on this machine -- `ollama list` should show `gemma4:31b-cloud`).
4. Copy `.env.example` to `.env` and paste in an API key from
   <https://ollama.com/settings/keys> (sign in with the same account Ollama
   is already signed in with on this machine). This key is only needed for
   `enrich_reviewers.py`'s web searches -- `suggest_reviewers.py` just talks
   to your already-signed-in local Ollama app and needs no key.

## 1. `enrich_reviewers.py` -- fill in Position / Interests / Website

For every reviewer on the `ReviewerList` sheet with all three columns blank,
this web-searches their name + affiliation, then asks the model to extract
a job title, a short list of research interests, and a best-guess website
(personal/faculty page > Google Scholar > LinkedIn), and writes those into
the sheet. It never invents info that isn't backed by a search result --
rows it can't confidently identify get `Position = "Unknown"` instead of a
guess.

```
python enrich_reviewers.py --limit 5    # try it on 5 people first
python enrich_reviewers.py              # then do everyone left
```

Saves progress every 5 rows, so it's safe to stop and resume.

## 2. `suggest_reviewers.py` -- populate AISuggestedReviewers

For every submitted paper with a blank `AISuggestedReviewers` cell, this:

- Parses the `Authors` field and excludes anyone on that list (plus anyone
  already in `Reviewer 1/2/3`) from the candidate pool -- co-authors are
  never suggested to review their own paper.
- Sends the paper's Title/Keywords/Abstract plus the remaining candidates'
  Position/Interests to the model and asks it to rank the best-fit
  reviewers, choosing only from that candidate list (so it can't suggest
  someone who isn't actually in ReviewerList).
- For any candidate who is themselves a SIGDSA26 author on a *different*
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
```

Run this again whenever you add new papers -- it'll only compute
suggestions for the new rows. Run it with `--force` if you want to
recompute everyone (e.g. after re-running `enrich_reviewers.py` and getting
better interest data).

## Notes

- `common.py` holds shared logic (`.env` loading, the Ollama calls, and the
  author-name parsing/matching) -- no need to run it directly.
- These are AI-generated *suggestions* to speed up manual reviewer
  selection, not final assignments -- spot-check a few before relying on
  them, especially early on.
