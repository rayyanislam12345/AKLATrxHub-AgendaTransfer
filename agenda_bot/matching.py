"""Fuzzy matching of agenda items to hub transactions and deliverables."""

from __future__ import annotations

import re
from difflib import SequenceMatcher

STOPWORDS = {
    "the", "a", "an", "of", "for", "to", "and", "in", "on", "with", "by", "from", "at",
    "re", "draft", "revised", "finalized", "finalised", "final", "finalization",
    "finalisation", "review", "reviewed", "amendments", "amendment", "updated", "update",
    "preparation", "prepare", "prepared", "comments", "circulation", "lenders", "document",
}


def norm(text: str) -> str:
    text = str(text or "").lower().replace("&", " and ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def compact(text: str) -> str:
    """'M-6 Motorway' -> 'm6motorway' so 'M6' and 'M-6' compare equal."""
    return norm(text).replace(" ", "")


def _stem(word: str) -> str:
    for suffix in ("ies", "es", "s"):
        if len(word) > 4 and word.endswith(suffix):
            return word[: -len(suffix)] + ("y" if suffix == "ies" else "")
    return word


def keywords(text: str) -> set[str]:
    return {_stem(w) for w in norm(text).split() if w not in STOPWORDS and len(w) > 1}


def transaction_score(agenda_name: str, hub_name: str, aliases: list[str] | None = None) -> float:
    """Score in [0, 1] for how well an agenda transaction matches a hub transaction."""
    candidates = [hub_name, *(aliases or [])]
    best = 0.0
    a_norm, a_compact = norm(agenda_name), compact(agenda_name)
    if not a_compact:
        return 0.0
    for cand in candidates:
        c_norm, c_compact = norm(cand), compact(cand)
        if not c_compact:
            continue
        if a_compact == c_compact:
            return 1.0
        # Short code appearing as a whole token, e.g. "PTQ" in "PTQ - Port Qasim PPP".
        c_tokens = c_norm.split()
        a_tokens = a_norm.split()
        if a_tokens and _contains_run(c_tokens, a_tokens):
            best = max(best, 0.95)
        elif c_tokens and _contains_run(a_tokens, c_tokens):
            best = max(best, 0.9)
        # Handle "M6" vs "M-6 ..." once tokens are compacted.
        if c_compact.startswith(a_compact) and len(a_compact) >= 2 and _boundary(cand, a_compact):
            best = max(best, 0.92)
        # Initials, e.g. "PIDG" for "Private Infrastructure Development Group".
        if len(a_compact) >= 3 and a_compact == initials(cand):
            best = max(best, 0.9)
        best = max(best, SequenceMatcher(None, a_norm, c_norm).ratio())
    return best


def initials(text: str) -> str:
    text = re.sub(r"\((private|pvt)\)|\blimited\b|\bltd\b|\bpvt\b", " ", str(text or ""), flags=re.I)
    return "".join(w[0] for w in norm(text).split() if w not in {"of", "the", "and", "for", "a", "an"})


def _contains_run(haystack: list[str], needle: list[str]) -> bool:
    n = len(needle)
    return any(haystack[i:i + n] == needle for i in range(len(haystack) - n + 1))


def _boundary(hub_name: str, code: str) -> bool:
    """True when `code` ends on a word boundary inside hub_name (so 'M6' != 'M60')."""
    consumed = ""
    for idx, ch in enumerate(hub_name.lower()):
        if ch.isalnum():
            consumed += ch
        if consumed == code:
            rest = hub_name[idx + 1:]
            return not rest or not rest[0].isalnum()
        if not code.startswith(consumed):
            return False
    return False


def deliverable_score(agenda_texts: list[str], hub_text: str) -> float:
    """Best similarity between any agenda description and a hub deliverable cell."""
    hub_kw = keywords(hub_text)
    hub_norm = norm(hub_text)
    if not hub_norm:
        return 0.0
    best = 0.0
    for text in agenda_texts:
        t_norm = norm(text)
        if not t_norm:
            continue
        if t_norm == hub_norm:
            return 1.0
        ratio = SequenceMatcher(None, t_norm, hub_norm).ratio()
        t_kw = keywords(text)
        overlap = 0.0
        if t_kw and hub_kw:
            common = len(t_kw & hub_kw)
            overlap = common / min(len(t_kw), len(hub_kw))
            # Penalise one-word coincidences between long descriptions.
            if common == 1 and max(len(t_kw), len(hub_kw)) > 3:
                overlap *= 0.6
        # Keyword overlap alone never beats an exact match, so "Risk Allocation Matrix"
        # prefers that row over "Presentation on Risk Allocation Matrix".
        best = max(best, ratio, overlap * (0.9 + 0.09 * ratio))
    return best
