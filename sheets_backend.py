"""A Google Sheets-backed drop-in for the small slice of openpyxl's API
the three scripts actually use: wb.sheetnames, wb[sheet_name], ws.max_row,
ws[1] (header row iteration), ws.cell(row, col).value (get/set), wb.save().

This lets enrich_reviewers.py / suggest_reviewers.py / update_review_counts.py
run unchanged against either a local .xlsx (openpyxl) or a shared Google
Sheet (this module) -- see common.load_workbook(), which picks based on
whether GOOGLE_SHEET_ID is set.

Reads the whole sheet in one batched API call up front; writes are batched
in memory and only sent to Google on .save() -- this matters because the
Sheets API is rate-limited (unlike a local file), and the existing scripts
already call .save() periodically (not after every single cell), so this
lines up naturally with no changes needed to their save cadence.
"""
from __future__ import annotations

import json
import os
from pathlib import Path

import gspread
from google.oauth2.service_account import Credentials

_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
_HERE = Path(__file__).resolve().parent


def _resolve(path_str: str) -> Path:
    """Resolve relative to reviewer_tools/, not the current working
    directory, so this works the same whether you run the scripts from
    inside reviewer_tools/ or elsewhere."""
    p = Path(path_str)
    return p if p.is_absolute() else _HERE / p


def _load_client() -> gspread.Client:
    """Authorize and return a gspread Client. Two modes:

    - Service account (GOOGLE_SERVICE_ACCOUNT_FILE / _JSON set): one shared
      credential, same access for everyone who has the key file. Simple,
      but every edit shows up under the service account's identity in the
      Sheet's revision history, and access can't be revoked per-person.

    - Per-person OAuth (neither of the above set): each person signs in
      with their OWN Google account via a one-time browser consent screen
      (gspread caches the resulting token locally after that, so it's
      only interactive once per person per machine). Edits are attributed
      to each real person, and access is revoked the normal way -- remove
      them from the Sheet's sharing list. Reads the OAuth client
      credential from GOOGLE_OAUTH_CLIENT_FILE, or gspread's own default
      config location if that's not set (see gspread.oauth()'s docs).
    """
    raw_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if raw_json:
        info = json.loads(raw_json)
        creds = Credentials.from_service_account_info(info, scopes=_SCOPES)
        return gspread.authorize(creds)

    file_path = os.environ.get("GOOGLE_SERVICE_ACCOUNT_FILE", "").strip()
    if file_path:
        creds = Credentials.from_service_account_file(str(_resolve(file_path)), scopes=_SCOPES)
        return gspread.authorize(creds)

    oauth_kwargs = {"scopes": _SCOPES}
    oauth_client_file = os.environ.get("GOOGLE_OAUTH_CLIENT_FILE", "").strip()
    if oauth_client_file:
        oauth_kwargs["credentials_filename"] = str(_resolve(oauth_client_file))
    try:
        return gspread.oauth(**oauth_kwargs)
    except FileNotFoundError as exc:
        raise RuntimeError(
            "GOOGLE_SHEET_ID is set but no Google credentials were found. Either "
            "set GOOGLE_SERVICE_ACCOUNT_FILE/GOOGLE_SERVICE_ACCOUNT_JSON in .env "
            "(shared service account), or set up per-person OAuth: get an OAuth "
            "client credential (Google Cloud Console -> APIs & Services -> "
            "Credentials -> Create Credentials -> OAuth client ID -> Desktop app), "
            "save it locally, and set GOOGLE_OAUTH_CLIENT_FILE to its path -- "
            "you'll then get a one-time browser sign-in prompt. See README.md's "
            "'Google Sheets backend' section."
        ) from exc


class _GCell:
    __slots__ = ("_ws", "row", "column")

    def __init__(self, ws: "_GWorksheet", row: int, column: int):
        self._ws = ws
        self.row = row
        self.column = column

    @property
    def value(self):
        return self._ws._get(self.row, self.column)

    @value.setter
    def value(self, v):
        self._ws._set(self.row, self.column, v)


class _GWorksheet:
    def __init__(self, gs_worksheet):
        self._gs = gs_worksheet
        self.title = gs_worksheet.title
        # One batched read of the whole sheet. FORMULA render option gives
        # back "=..." for formula cells and the plain value otherwise --
        # matches openpyxl's cell.value semantics closely enough for our
        # read patterns (mostly: is this blank? what does it say?).
        self._grid: list[list] = gs_worksheet.get_values(value_render_option="FORMULA")
        self._pending: dict[tuple[int, int], object] = {}

    @property
    def max_row(self) -> int:
        return len(self._grid)

    def _get(self, row: int, col: int):
        r, c = row - 1, col - 1
        if 0 <= r < len(self._grid) and 0 <= c < len(self._grid[r]):
            v = self._grid[r][c]
            return v if v != "" else None
        return None

    def _set(self, row: int, col: int, value) -> None:
        self._pending[(row, col)] = value
        # Mirror into the in-memory grid too, so a read immediately after a
        # write in the same run (before save()) sees the new value --
        # several call sites in the scripts do exactly this (e.g. reading
        # a header cell right after possibly creating it).
        while len(self._grid) < row:
            self._grid.append([])
        r = self._grid[row - 1]
        while len(r) < col:
            r.append("")
        r[col - 1] = "" if value is None else value

    def __getitem__(self, row_num: int) -> list[_GCell]:
        """Mimic openpyxl's ws[1] -- a list of cell objects for that row,
        covering however many columns that row actually has data in."""
        ncols = len(self._grid[row_num - 1]) if row_num - 1 <= len(self._grid) - 1 else 0
        return [_GCell(self, row_num, c) for c in range(1, ncols + 1)]

    def cell(self, row: int, column: int) -> _GCell:
        return _GCell(self, row, column)

    def flush(self) -> None:
        if not self._pending:
            return
        updates = [
            {"range": gspread.utils.rowcol_to_a1(r, c), "values": [["" if v is None else v]]}
            for (r, c), v in self._pending.items()
        ]
        self._gs.batch_update(updates, value_input_option="USER_ENTERED")
        self._pending.clear()


class GoogleSheetWorkbook:
    """Opens a Google Sheet by ID and wraps each tab as a _GWorksheet."""

    def __init__(self, sheet_id: str):
        gc = _load_client()
        try:
            self._sh = gc.open_by_key(sheet_id)
        except gspread.exceptions.APIError as exc:
            raise RuntimeError(
                f"Couldn't open Google Sheet {sheet_id!r}: {exc}. Check that "
                f"GOOGLE_SHEET_ID is correct and that the sheet is shared "
                f"(Editor access) with whichever identity is authenticating -- "
                f"the service account's client_email, or your own Google "
                f"account if using per-person OAuth."
            ) from exc
        self._worksheets = {ws.title: _GWorksheet(ws) for ws in self._sh.worksheets()}

    @property
    def sheetnames(self) -> list[str]:
        return list(self._worksheets)

    def __getitem__(self, name: str) -> _GWorksheet:
        return self._worksheets[name]

    def save(self, _path=None) -> None:
        for ws in self._worksheets.values():
            ws.flush()
