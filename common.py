"""Shared helpers for the SIGDSA reviewer-list scripts.

Handles: .env loading, talking to the local Ollama app (chat completions)
and to Ollama Cloud's web-search API, and parsing author names out of the
Submissions sheet's "Authors" free-text field.
"""
from __future__ import annotations

import json
import os
import re
import time
from pathlib import Path

import requests

HERE = Path(__file__).resolve().parent
WORKBOOK_PATH = HERE.parent / "SIGDSA26_2026-09-01_ReviewerList.xlsx"


# ---------------------------------------------------------------------------
# .env loading (no external dependency -- just enough to read KEY=VALUE lines)
# ---------------------------------------------------------------------------
def load_env(env_path: Path | None = None) -> None:
    env_path = env_path or (HERE / ".env")
    if not env_path.exists():
        return
    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


load_env()

OLLAMA_HOST = os.environ.get("OLLAMA_HOST", "http://localhost:11434").rstrip("/")
OLLAMA_MODEL = os.environ.get("OLLAMA_MODEL", "gemma4:31b-cloud")
OLLAMA_API_KEY = os.environ.get("OLLAMA_API_KEY", "").strip()


# ---------------------------------------------------------------------------
# Ollama chat (local app -- already signed in, no API key needed)
# ---------------------------------------------------------------------------
def ollama_chat(messages: list[dict], *, model: str | None = None,
                 temperature: float = 0.2, retries: int = 3) -> str:
    """Call the local Ollama app's /api/chat and return the reply text."""
    url = f"{OLLAMA_HOST}/api/chat"
    payload = {
        "model": model or OLLAMA_MODEL,
        "messages": messages,
        "stream": False,
        "options": {"temperature": temperature},
    }
    last_err = None
    for attempt in range(1, retries + 1):
        try:
            resp = requests.post(url, json=payload, timeout=120)
            resp.raise_for_status()
            data = resp.json()
            return data["message"]["content"]
        except Exception as exc:  # noqa: BLE001
            last_err = exc
            time.sleep(min(2 ** attempt, 15))
    raise RuntimeError(
        f"Ollama chat call failed after {retries} attempts. Is the Ollama "
        f"app running? ({OLLAMA_HOST}) Last error: {last_err}"
    )


def extract_json(text: str) -> dict | list:
    """Pull the first JSON object/array out of a model reply."""
    text = text.strip()
    text = re.sub(r"^```(json)?|```$", "", text, flags=re.MULTILINE).strip()
    match = re.search(r"[\[{].*[\]}]", text, flags=re.DOTALL)
    if not match:
        raise ValueError(f"No JSON found in model reply: {text[:200]!r}")
    return json.loads(match.group(0))


# ---------------------------------------------------------------------------
# Ollama Cloud web search (needs OLLAMA_API_KEY)
# ---------------------------------------------------------------------------
MAX_RATE_LIMIT_WAIT = 30  # seconds -- never block longer than this on a single retry


def ollama_web_search(query: str, *, max_results: int = 5, retries: int = 3) -> list[dict]:
    if not OLLAMA_API_KEY:
        raise RuntimeError(
            "OLLAMA_API_KEY is not set. Copy .env.example to .env and paste "
            "a key from https://ollama.com/settings/keys into it."
        )
    url = "https://ollama.com/api/web_search"
    headers = {"Authorization": f"Bearer {OLLAMA_API_KEY}"}
    payload = {"query": query, "max_results": max_results}
    last_err = None
    for attempt in range(1, retries + 1):
        try:
            resp = requests.post(url, json=payload, headers=headers, timeout=30)
            if resp.status_code == 401:
                raise RuntimeError("OLLAMA_API_KEY was rejected (401 Unauthorized). "
                                    "Check the key in .env.")
            if resp.status_code == 429:
                # Rate-limited. Respect Retry-After if given, but never block
                # longer than MAX_RATE_LIMIT_WAIT on any single attempt --
                # a large/daily-quota Retry-After should surface as an error
                # (so the run can be resumed later), not hang the script.
                try:
                    retry_after = float(resp.headers.get("Retry-After", 10 * attempt))
                except ValueError:
                    retry_after = 10 * attempt
                if retry_after > MAX_RATE_LIMIT_WAIT:
                    raise RuntimeError(
                        f"Ollama Cloud web search is rate-limited and asked to wait "
                        f"{retry_after:.0f}s (likely an hourly/daily search quota) -- "
                        f"stopping here so you can resume later. Already-filled rows "
                        f"were saved; just re-run the script once the quota resets."
                    )
                last_err = "429 Too Many Requests"
                time.sleep(retry_after)
                continue
            resp.raise_for_status()
            data = resp.json()
            return data.get("results", [])
        except RuntimeError:
            raise
        except Exception as exc:  # noqa: BLE001
            last_err = exc
            time.sleep(min(2 ** attempt, 15))
    raise RuntimeError(f"Ollama web_search failed after {retries} attempts: {last_err}")


# ---------------------------------------------------------------------------
# Author-name parsing / matching
# ---------------------------------------------------------------------------
def split_authors(authors_field: str) -> list[str]:
    """Turn 'A, B, C and D' (possibly with embedded newlines) into
    ['A', 'B', 'C', 'D'].
    """
    if not authors_field:
        return []
    text = re.sub(r"\s+", " ", str(authors_field)).strip()
    text = re.sub(r"\s+and\s+", ", ", text, flags=re.IGNORECASE)
    parts = [p.strip(" .") for p in text.split(",")]
    return [p for p in parts if p]


def normalize_name(name: str) -> str:
    name = re.sub(r"\s+", " ", str(name)).strip().lower()
    name = re.sub(r"[^\w\s'-]", "", name)
    return name


def names_match(a: str, b: str) -> bool:
    na, nb = normalize_name(a), normalize_name(b)
    if not na or not nb:
        return False
    if na == nb:
        return True
    # last-name + first-initial fallback (handles "J. Smith" vs "John Smith")
    pa, pb = na.split(), nb.split()
    if pa and pb and pa[-1] == pb[-1] and pa[0][:1] == pb[0][:1]:
        return True
    return False
