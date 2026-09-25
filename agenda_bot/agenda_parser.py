"""Parse AKLA "Daily Agenda" Word documents into structured agenda items.

Expected layout (labels are matched case-insensitively and tolerate small
variations):

    Daily Agenda
    Date: September 02, 2026
    Senior Counsel: Areen Ali Shah
    Deliverables
    Transaction: M6.
    Scope of Task: ...
    Complete Deliverable: ...
    Day Deliverable: ...
    Current Deliverable Status: ...
    Transaction: PTQ.
    ...
    Daily Timesheet            <- everything from here on is ignored
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import BinaryIO, Iterable

import docx
from docx.document import Document as _Document
from docx.table import Table
from docx.text.paragraph import Paragraph

# Normalised label -> field name on AgendaItem.
ITEM_LABELS = {
    "transaction": "transaction",
    "transaction name": "transaction",
    "matter": "transaction",
    "project": "transaction",
    "scope of task": "scope",
    "scope of tasks": "scope",
    "scope": "scope",
    "task": "scope",
    "complete deliverable": "complete_deliverable",
    "overall deliverable": "complete_deliverable",
    "final deliverable": "complete_deliverable",
    "day deliverable": "day_deliverable",
    "days deliverable": "day_deliverable",
    "todays deliverable": "day_deliverable",
    "current deliverable status": "status",
    "deliverable status": "status",
    "current status": "status",
    "status": "status",
}

# Header labels whose value is the employee's name.
NAME_LABELS = (
    "senior counsel", "counsel", "associate", "senior associate", "junior associate",
    "partner", "managing partner", "trainee", "trainee lawyer", "intern", "paralegal",
    "lawyer", "name", "employee", "prepared by",
)

# Any label ending in one of these ("Trainee Associate", "Associate Lawyer") is a name line.
NAME_ROLE_WORDS = {"counsel", "associate", "partner", "trainee", "intern", "lawyer",
                   "paralegal", "advocate", "name", "employee", "consultant", "officer"}

# Section headings that end the agenda part of the document.
STOP_HEADINGS = ("daily timesheet", "timesheet")

# Transactions that are not client matters and should never hit the hub.
NON_TRANSACTIONS = {"internal", "admin", "administrative", "n/a", "na", "none", "-", "general"}

_LABEL_RE = re.compile(r"^\s*([A-Za-z][A-Za-z’' &/.-]{0,40}?)\s*:\s*(.*)$", re.S)
_FILENAME_RE = re.compile(
    r"daily[\s_-]*agenda[\s_-]*-?[\s_-]*(?P<name>.+?)[\s_\-\[(]+(?P<date>[A-Za-z]+[\s_-]+\d{1,2}[\s_,-]+\d{4})",
    re.I,
)


@dataclass
class AgendaItem:
    transaction: str = ""
    scope: str = ""
    complete_deliverable: str = ""
    day_deliverable: str = ""
    status: str = ""

    def is_empty(self) -> bool:
        return not any((self.transaction, self.scope, self.complete_deliverable,
                        self.day_deliverable, self.status))


@dataclass
class Agenda:
    employee: str
    agenda_date: date | None
    items: list[AgendaItem] = field(default_factory=list)
    source: str = ""
    warnings: list[str] = field(default_factory=list)


def normalise_label(label: str) -> str:
    label = label.lower().replace("’", "").replace("'", "")
    label = re.sub(r"[^a-z ]+", " ", label)
    return re.sub(r"\s+", " ", label).strip()


def clean_value(value: str) -> str:
    value = re.sub(r"\s+", " ", value.replace("\t", " ")).strip()
    return value.rstrip(".;, ").strip()


def parse_date(text: str) -> date | None:
    text = re.sub(r"[_]+", " ", text).strip().rstrip(".")
    text = re.sub(r"(\d)(st|nd|rd|th)\b", r"\1", text)
    text = re.sub(r"\s*,\s*", ", ", text)
    text = re.sub(r"\s+", " ", text)
    for fmt in ("%B %d, %Y", "%B %d %Y", "%b %d, %Y", "%b %d %Y", "%d %B %Y", "%d %B, %Y",
                "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _iter_block_text(doc: _Document) -> Iterable[str]:
    """Yield paragraph texts in document order, descending into tables."""
    body = doc.element.body
    for child in body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            yield Paragraph(child, doc).text
        elif tag == "tbl":
            table = Table(child, doc)
            for row in table.rows:
                seen = set()
                for cell in row.cells:
                    # Merged cells are returned once per grid column; skip repeats.
                    if id(cell._tc) in seen:
                        continue
                    seen.add(id(cell._tc))
                    for para in cell.paragraphs:
                        yield para.text


def _name_date_from_filename(filename: str) -> tuple[str, date | None]:
    stem = Path(filename).stem
    stem = re.sub(r"^[0-9a-f]{6,}-", "", stem)  # upload prefixes like "acad96c8-"
    match = _FILENAME_RE.search(stem)
    if not match:
        return "", None
    name = re.sub(r"[_]+", " ", match.group("name")).strip(" -_")
    return name, parse_date(match.group("date"))


def parse_agenda(source: str | Path | BinaryIO | bytes, filename: str | None = None) -> Agenda:
    """Parse one Daily Agenda .docx (path, file object or raw bytes)."""
    if isinstance(source, bytes):
        source = io.BytesIO(source)
    if filename is None:
        filename = str(source) if isinstance(source, (str, Path)) else ""
    doc = docx.Document(source)

    employee = ""
    agenda_date: date | None = None
    items: list[AgendaItem] = []
    current: AgendaItem | None = None
    pending_field: str | None = None  # label seen with its value on the next line
    warnings: list[str] = []

    for raw in _iter_block_text(doc):
        text = raw.strip()
        if not text:
            continue
        if normalise_label(text) in STOP_HEADINGS:
            break

        match = _LABEL_RE.match(text)
        label = normalise_label(match.group(1)) if match else ""
        value = clean_value(match.group(2)) if match else ""

        if match and label == "date":
            if agenda_date is None:
                agenda_date = parse_date(value)
            pending_field = None
            continue
        if match and not employee and label not in ITEM_LABELS and (
                label in NAME_LABELS or label.split()[-1:] and label.split()[-1] in NAME_ROLE_WORDS):
            employee = value
            pending_field = None
            continue
        if match and label in ITEM_LABELS:
            field_name = ITEM_LABELS[label]
            if field_name == "transaction" or current is None:
                current = AgendaItem()
                items.append(current)
            if value:
                setattr(current, field_name, value)
                pending_field = None
            else:
                pending_field = field_name
            continue

        if pending_field and current is not None:
            setattr(current, pending_field, clean_value(text))
            pending_field = None

    fn_name, fn_date = _name_date_from_filename(filename) if filename else ("", None)
    if not employee:
        employee = fn_name
        if not employee:
            warnings.append("could not find the employee name")
    if agenda_date is None:
        agenda_date = fn_date
        if agenda_date is None:
            warnings.append("could not find the agenda date")

    kept = []
    for item in items:
        if item.is_empty():
            continue
        if not item.transaction:
            warnings.append(f"deliverable without a transaction skipped: {item.day_deliverable or item.scope!r}")
            continue
        if item.transaction.lower() in NON_TRANSACTIONS:
            continue
        kept.append(item)

    return Agenda(employee=employee, agenda_date=agenda_date, items=kept,
                  source=Path(filename).name if filename else "", warnings=warnings)
