"""Backend-independent logic: link agenda items to hub rows and build the new columns."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

from .agenda_parser import Agenda
from .config import Config
from .matching import deliverable_score, norm, transaction_score

LOG_HEADERS = [
    "Agenda Date", "Employee", "Agenda File", "Agenda Transaction", "Matched Transaction",
    "Transaction Score", "Day Deliverable", "Complete Deliverable", "Matched Hub Deliverable",
    "Deliverable Score", "Hub Row(s)", "Status", "Note",
]


@dataclass
class HubTable:
    """A worksheet's used range: `rows[i]` is the sheet row `first_row + i` (1-based)."""
    rows: list[list]
    first_row: int = 1
    first_col: int = 1


@dataclass
class ColumnWrite:
    header: str
    column: int             # 1-based sheet column
    header_row: int         # 1-based sheet row of the header cell
    values: list            # one per data row, starting at header_row + 1


@dataclass
class UpdatePlan:
    header_row: int
    columns: list[ColumnWrite]
    log_rows: list[list]
    active_transactions: list[str] = field(default_factory=list)
    unmatched: list[list] = field(default_factory=list)


def _cell(row: list, idx: int):
    return row[idx] if idx < len(row) else None


def _text(value) -> str:
    return "" if value is None else str(value).strip()


def _find_column(headers: list, candidates: list[str], exact_only: bool = False) -> int | None:
    wanted = [norm(c) for c in candidates]
    normed = [norm(_text(h)) for h in headers]
    for cand in wanted:                       # exact match, in preference order
        if cand in normed:
            return normed.index(cand)
    if exact_only:
        return None
    for cand in wanted:                       # header that starts with a candidate
        for idx, header in enumerate(normed):
            if header and header.startswith(cand):
                return idx
    return None


def locate_header(table: HubTable, cfg: Config) -> tuple[int, int, int | None]:
    """Return (row index in table, transaction col index, deliverable col index or None).

    A header row must have at least three filled cells, so title rows such as a merged
    "TRANSACTIONS HUB" banner are never mistaken for it. Rows where both the transaction
    and deliverable headers match exactly win over looser matches.
    """
    if cfg.hub.header_row:
        indices = [cfg.hub.header_row - table.first_row]
    else:
        indices = range(min(25, len(table.rows)))
    best = None
    for idx in indices:
        if idx < 0 or idx >= len(table.rows):
            continue
        headers = table.rows[idx]
        if sum(1 for h in headers if _text(h)) < 3:
            continue
        for rank, exact in ((0, True), (1, False)):
            tx_col = _find_column(headers, cfg.hub.transaction_columns, exact)
            if tx_col is None:
                continue
            dl_col = _find_column(headers, cfg.hub.deliverable_columns, exact)
            score = (rank, 0 if dl_col is not None else 1, idx)
            if best is None or score < best[0]:
                best = (score, idx, tx_col, dl_col)
            break
    if best is None:
        raise ValueError(
            "could not find the Transaction column header; set [hub] header_row and "
            "transaction_columns in the config to match the sheet"
        )
    return best[1], best[2], best[3]


def build_update(table: HubTable, agendas: list[Agenda], cfg: Config,
                 run_date: date | None = None) -> UpdatePlan:
    out = cfg.output
    hdr_idx, tx_col, dl_col = locate_header(table, cfg)
    headers = table.rows[hdr_idx]
    data = table.rows[hdr_idx + 1:]
    header_row = table.first_row + hdr_idx

    # Columns this bot owns; ignored when deciding whether a hub row is blank.
    out_names = {norm(n) for n in (out.working_column, out.status_column, out.active_column,
                                   out.date_column) if n}
    own = {i for i, h in enumerate(headers) if norm(_text(h)) in out_names}

    def has_content(row: list) -> bool:
        return any(_text(v) for i, v in enumerate(row) if i not in own)

    # Transaction name for every data row (fill down merged / grouped cells).
    row_tx: list[str] = []
    last = ""
    for row in data:
        name = _text(_cell(row, tx_col))
        if name:
            last = name
        elif cfg.hub.fill_down_transaction and has_content(row):
            name = last
        row_tx.append(name)

    tx_rows: dict[str, list[int]] = {}
    for i, name in enumerate(row_tx):
        if name:
            tx_rows.setdefault(name, []).append(i)

    workers: list[list[str]] = [[] for _ in data]
    statuses: list[list[str]] = [[] for _ in data]
    active_tx: set[str] = set()
    log_rows: list[list] = []
    unmatched: list[list] = []

    def add(i: int, employee: str, status: str) -> None:
        if employee and employee not in workers[i]:
            workers[i].append(employee)
        if status:
            entry = f"{employee}: {status}" if employee else status
            if entry not in statuses[i]:
                statuses[i].append(entry)

    for agenda in agendas:
        for item in agenda.items:
            scored = [(transaction_score(item.transaction, name, cfg.aliases_for(name)), name)
                      for name in tx_rows]
            tx_score, tx_name = max(scored, default=(0.0, ""))
            log = [agenda.agenda_date.isoformat() if agenda.agenda_date else "", agenda.employee,
                   agenda.source, item.transaction, "", round(tx_score, 2), item.day_deliverable,
                   item.complete_deliverable, "", "", "", item.status, ""]
            if tx_score < cfg.matching.transaction_threshold:
                log[4] = f"(closest: {tx_name})" if tx_name else ""
                log[12] = "TRANSACTION NOT FOUND in hub - add an alias in config.toml or check the name"
                log_rows.append(log)
                unmatched.append(log)
                continue

            log[4] = tx_name
            active_tx.add(tx_name)
            rows = tx_rows[tx_name]
            target_rows = rows
            note = ""
            if dl_col is not None:
                texts = [item.day_deliverable, item.complete_deliverable, item.scope]
                dl_scored = [(deliverable_score(texts, _text(_cell(data[i], dl_col))), i) for i in rows]
                dl_score, best_row = max(dl_scored, default=(0.0, -1))
                log[9] = round(dl_score, 2)
                if dl_score >= cfg.matching.deliverable_threshold:
                    target_rows = [best_row]
                    log[8] = _text(_cell(data[best_row], dl_col))
                else:
                    note = "deliverable not found in hub - marked against every row of the transaction"
                    unmatched.append(log)
            for i in target_rows:
                add(i, agenda.employee, item.status)
            log[10] = ", ".join(str(header_row + 1 + i) for i in target_rows)
            log[12] = note or "matched"
            log_rows.append(log)

    dates = sorted({a.agenda_date for a in agendas if a.agenda_date})
    date_label = (run_date or (dates[-1] if dates else date.today())).isoformat()

    blank = [not has_content(row) for row in data]
    working = [None if blank[i] else "; ".join(workers[i]) for i in range(len(data))]
    status = [None if blank[i] else "; ".join(statuses[i]) for i in range(len(data))]
    active = [None if blank[i] or not row_tx[i]
              else (out.active_yes if row_tx[i] in active_tx else out.active_no)
              for i in range(len(data))]
    dated = [None if blank[i] else date_label for i in range(len(data))]

    wanted = [(out.working_column, working), (out.status_column, status),
              (out.active_column, active)]
    if out.date_column:
        wanted.append((out.date_column, dated))

    # Reuse existing output columns; otherwise append after the last column that holds
    # anything, in any row, so existing data can never be overwritten.
    width = max((len(r) for r in table.rows), default=0)
    last_used = max((i for r in table.rows[hdr_idx:] for i, v in enumerate(r) if _text(v)),
                    default=-1)
    next_free = max(last_used, len(headers) - 1, width - 1) + 1
    columns = []
    for name, values in wanted:
        idx = next((i for i, h in enumerate(headers) if norm(_text(h)) == norm(name)), None)
        if idx is None:
            idx = next_free
            next_free += 1
        occupied = [r for r in data if _text(_cell(r, idx))]
        if idx not in own and (_text(_cell(headers, idx)) or occupied):
            raise ValueError(f"refusing to write '{name}' into column {idx + 1}: it already has data")
        columns.append(ColumnWrite(header=name, column=table.first_col + idx,
                                   header_row=header_row, values=values))

    return UpdatePlan(header_row=header_row, columns=columns, log_rows=log_rows,
                      active_transactions=sorted(active_tx), unmatched=unmatched)
