"""Read agendas from a folder and update a local copy of the hub workbook (openpyxl)."""

from __future__ import annotations

from copy import copy
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries

from .agenda_parser import Agenda, parse_agenda
from .hub import LOG_HEADERS, HubTable, UpdatePlan


def load_agendas_from_folder(folder: str | Path, on_date: date | None = None) -> list[Agenda]:
    agendas = []
    for path in sorted(Path(folder).rglob("*.docx")):
        if path.name.startswith("~$"):  # Word lock files
            continue
        agenda = parse_agenda(path)
        if on_date and agenda.agenda_date and agenda.agenda_date != on_date:
            continue
        agendas.append(agenda)
    return agendas


def read_table(ws) -> HubTable:
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    return HubTable(rows=rows, first_row=1, first_col=1)


def pick_sheet(wb, name: str):
    if name:
        return wb[name]
    return wb.worksheets[0]


def apply_plan(wb, ws, plan: UpdatePlan, log_sheet: str) -> None:
    for col in plan.columns:
        header_cell = ws.cell(row=col.header_row, column=col.column, value=col.header)
        # Style new headers like the header cell to their left.
        if col.column > 1:
            left = ws.cell(row=col.header_row, column=col.column - 1)
            if left.has_style and not header_cell.has_style:
                header_cell._style = copy(left._style)
        for offset, value in enumerate(col.values, start=1):
            ws.cell(row=col.header_row + offset, column=col.column, value=value)

    if log_sheet:
        if log_sheet in wb.sheetnames:
            del wb[log_sheet]
        log = wb.create_sheet(log_sheet)
        log.append(LOG_HEADERS)
        for row in plan.log_rows:
            log.append(row)


def _extend_filter(ws, plan: UpdatePlan) -> None:
    """Make sure the sheet's filter covers the new columns so 'Active Today' can be filtered."""
    if not plan.columns:
        return
    last_col = max(c.column for c in plan.columns)
    last_row = plan.header_row + len(plan.columns[0].values)
    min_col = 1
    if ws.auto_filter.ref:
        min_col, _, max_col, max_row = range_boundaries(ws.auto_filter.ref)
        last_col, last_row = max(last_col, max_col), max(last_row, max_row)
    ws.auto_filter.ref = (f"{get_column_letter(min_col)}{plan.header_row}:"
                          f"{get_column_letter(last_col)}{last_row}")


def update_workbook(hub_path: str | Path, out_path: str | Path, build, sheet_name: str,
                    log_sheet: str) -> UpdatePlan:
    """`build` is a callable HubTable -> UpdatePlan."""
    keep_vba = str(hub_path).lower().endswith(".xlsm")
    wb = openpyxl.load_workbook(hub_path, keep_vba=keep_vba)
    ws = pick_sheet(wb, sheet_name)
    plan = build(read_table(ws))
    apply_plan(wb, ws, plan, log_sheet)
    _extend_filter(ws, plan)
    wb.save(out_path)
    return plan
