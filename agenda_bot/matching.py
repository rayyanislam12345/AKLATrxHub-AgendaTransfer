"""Fuzzy matching of agenda items to hub transactions and deliverables."""

from __future__ import annotations

import re
from difflib import SequenceMatcher

STOPWORDS = {
    "the", "a", "an", "of", "for", "to", "and", "in", "on", "with", "by", "from", "at",
    "re", "draft", "revised", "finalized", "finalised", "final", "finalization",
    "finalisation", "review", "reviewed", "amendments", "amendment", "updated", "update",
    "preparation", "prepare", "prepared", "comments", "circulation", "lenders", "document",
    "finalise", "finalize", "first", "version", "versions", "complete", "completed",
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


SYNONYMS = {"memo": "memorandum", "dd": "diligence", "i": "1", "ii": "2", "iii": "3",
            "iv": "4", "v": "5", "vol": "volume", "agreements": "agreement"}

# Words too common in transaction names to identify one on their own.
GENERIC_TX_WORDS = {
    "due", "diligence", "dd", "project", "matter", "acquisition", "privatisation",
    "privatization", "proposal", "transaction", "limited", "ltd", "private", "pvt",
    "company", "co", "advisory", "legal", "services", "report", "workstream", "work",
    "the", "of", "and", "for", "a", "an", "on", "in", "sb", "sahib", "group",
}


def keywords(text: str) -> set[str]:
    words = (SYNONYMS.get(w, w) for w in norm(text).split())
    return {_stem(w) for w in words if w not in STOPWORDS and (len(w) > 1 or w.isdigit())}


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
        # Acronyms either way round: "PIDG" on the agenda for "Private Infrastructure
        # Development Group" in the hub, or "NSCL" in the hub for the full name.
        if acronym_match(agenda_name, cand):
            best = max(best, 0.9)
        best = max(best, SequenceMatcher(None, a_norm, c_norm).ratio(),
                   _distinctive_word_score(agenda_name, cand))
    return best


_ACRONYM_SKIP = {"of", "the", "and", "for", "a", "an", "&"}
_CORPORATE = {"limited", "ltd", "pvt", "private", "company", "co", "plc", "inc"}


def acronyms(text: str) -> set[str]:
    """Acronyms written in a name: all-caps words of 3+ characters ("PIDG", "NSCL",
    "(FSC)"), or the whole name when it is one short word ("Pidg")."""
    raw = str(text or "")
    found = {m.lower() for m in re.findall(r"\b[A-Z][A-Z0-9&]{2,}\b", raw)}
    found = {f.replace("&", "") for f in found}
    whole = compact(raw)
    if whole.isalpha() and 3 <= len(whole) <= 5 and len(norm(raw).split()) == 1:
        found.add(whole)
    return found


def _segments(text: str) -> list[list[str]]:
    """Word lists an acronym could stand for: the whole name (with and without any
    bracketed part) and each piece between dashes, slashes or brackets."""
    raw = str(text or "")
    pieces = [raw, re.sub(r"\([^)]*\)", " ", raw)]
    pieces += re.split(r"\s[-–—]\s|[–—|/,()]", raw)
    out = []
    for piece in pieces:
        words = [w for w in norm(piece).split() if w not in _ACRONYM_SKIP]
        if words:
            out.append(words)
    return out


def initials(text: str) -> set[str]:
    """Initials a name could be abbreviated to. For "National Steel Complex Limited":
    NSCL, NSC (without the corporate suffix) and NS, N (leading words), so "NSC"
    and "NSCL" both work."""
    variants: set[str] = set()
    for words in _segments(text):
        core = [w for w in words if w not in _CORPORATE]
        for ws in (words, core):
            letters = "".join(w[0] for w in ws)
            # Every leading run of words ("Sapphire Wind Power" -> SWP).
            variants.update(letters[:n] for n in range(2, len(letters) + 1))
    # "NSCL" / "AMPL" for a name the hub writes without Limited / Private Limited.
    variants |= {v + "l" for v in variants} | {v + "pl" for v in variants}
    return variants


def acronym_match(a: str, b: str) -> bool:
    """True when one name contains an acronym of the other."""
    return bool(acronyms(a) & initials(b)) or bool(acronyms(b) & initials(a))


def _distinctive_word_score(agenda_name: str, hub_name: str) -> float:
    """Agendas often write 'Riali – Due Diligence' or 'Artistic-DISCOS Privatisation'.

    Score by how many of the agenda's distinctive words (not generic ones such as
    'Due Diligence' or 'Project') appear in the hub name or its initials.
    """
    agenda_words = {_stem(w) for w in norm(agenda_name).split()} - GENERIC_TX_WORDS
    agenda_words = {w for w in agenda_words if len(w) > 1}
    if not agenda_words:
        return 0.0
    if acronym_match(agenda_name, hub_name):   # "NSCL - Gas Sale Matter"
        return 0.9
    hub_words = {_stem(w) for w in norm(hub_name).split()}
    found = len(agenda_words & hub_words)
    if not found:
        return 0.0
    share = found / len(agenda_words)
    return 0.8 + 0.15 * share if share >= 0.5 else 0.0


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
    """Best similarity between any agenda description and a hub deliverable cell.

    Based on shared keywords (cosine of the two keyword sets), so "Presentation on the
    USP memorandum" prefers "Presentation on the Memorandum on the USP" over
    "Memorandum on the USP". Plain string similarity only counts when it is high.
    """
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
        t_kw = keywords(text)
        if t_kw and hub_kw:
            if t_kw == hub_kw:
                best = max(best, 0.99)
            common = len(t_kw & hub_kw)
            best = max(best, 0.98 * common / (len(t_kw) * len(hub_kw)) ** 0.5)
            if t_kw <= hub_kw:                   # "The Proposal" -> "Proposal for ..."
                best = max(best, 0.6)
        ratio = SequenceMatcher(None, t_norm, hub_norm).ratio()
        if ratio >= 0.8:
            best = max(best, ratio * 0.98)
    return best
