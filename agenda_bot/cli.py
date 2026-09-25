"""Command line entry point.

    # Try it locally on a downloaded copy of the hub:
    python -m agenda_bot local --agendas ./agendas --hub "Transaction Hub.xlsx" --out updated.xlsx

    # Run against SharePoint (what the scheduled job does):
    python -m agenda_bot sharepoint --config config.toml
"""

from __future__ import annotations

import argparse
import sys
from datetime import date, datetime
from functools import partial
from zoneinfo import ZoneInfo

from .config import Config, load_config
from .hub import build_update


def _today() -> date:
    return datetime.now(ZoneInfo("Asia/Karachi")).date()


def _report(agendas, plan) -> None:
    print(f"Agendas read: {len(agendas)}")
    for a in agendas:
        when = a.agenda_date.isoformat() if a.agenda_date else "no date"
        print(f"  - {a.employee or '(unknown)'} [{when}] {a.source}: {len(a.items)} item(s)")
        for w in a.warnings:
            print(f"      warning: {w}")
    print(f"Active transactions today: {', '.join(plan.active_transactions) or 'none'}")
    if plan.unmatched:
        print("Needs attention:")
        for row in plan.unmatched:
            print(f"  - {row[1]}: '{row[3]}' / '{row[6]}' -> {row[12]}")


def _check_dates(agendas, on_date: date | None) -> None:
    if on_date is None:
        return
    undated = [a.source for a in agendas if a.agenda_date is None]
    if undated:
        print(f"warning: no date found in {', '.join(undated)}; included anyway", file=sys.stderr)


def run_local(args, cfg: Config) -> int:
    from .local_backend import load_agendas_from_folder, update_workbook

    on_date = args.date  # local runs use every agenda in the folder unless --date is given
    agendas = load_agendas_from_folder(args.agendas, on_date)
    _check_dates(agendas, on_date)
    plan = update_workbook(args.hub, args.out or args.hub,
                           partial(build_update, agendas=agendas, cfg=cfg, run_date=on_date),
                           cfg.hub.worksheet, cfg.output.log_sheet)
    _report(agendas, plan)
    print(f"Saved: {args.out or args.hub}")
    return 0


def run_sharepoint(args, cfg: Config) -> int:
    from .graph_backend import GraphClient, GraphWorkbook, load_agendas_from_sharepoint

    sp = cfg.sharepoint
    if not sp.hub_url or not sp.agenda_folder_url:
        print("error: set hub_url and agenda_folder_url (config or HUB_URL / AGENDA_FOLDER_URL)",
              file=sys.stderr)
        return 2
    on_date = args.date or _today()
    client = GraphClient(sp.tenant_id, sp.client_id, sp.client_secret)
    agendas = load_agendas_from_sharepoint(client, sp.agenda_folder_url, on_date, args.lookback_days)
    _check_dates(agendas, on_date)
    workbook = GraphWorkbook(client, sp.hub_url)
    try:
        sheet = cfg.hub.worksheet or workbook.sheet_names()[0]
        plan = build_update(workbook.read_table(sheet), agendas, cfg, run_date=on_date)
        if args.dry_run:
            print("Dry run: hub not modified")
        else:
            workbook.apply_plan(sheet, plan, cfg.output.log_sheet)
    finally:
        workbook.close()
    _report(agendas, plan)
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="agenda_bot", description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--config", help="path to config.toml")
    parser.add_argument("--date", type=date.fromisoformat,
                        help="agenda date to process (YYYY-MM-DD); sharepoint default: today in "
                             "Pakistan, local default: every agenda in the folder")
    sub = parser.add_subparsers(dest="mode", required=True)

    local = sub.add_parser("local", help="update a local .xlsx copy of the hub")
    local.add_argument("--agendas", required=True, help="folder containing agenda .docx files")
    local.add_argument("--hub", required=True, help="hub workbook (.xlsx)")
    local.add_argument("--out", help="where to save (default: overwrite --hub)")

    sp = sub.add_parser("sharepoint", help="update the hub on SharePoint via Microsoft Graph")
    sp.add_argument("--dry-run", action="store_true", help="match and report without writing")
    sp.add_argument("--lookback-days", type=int, default=3,
                    help="only open agenda files modified in the last N days (default 3)")

    args = parser.parse_args(argv)
    cfg = load_config(args.config)
    if args.mode == "local":
        return run_local(args, cfg)
    return run_sharepoint(args, cfg)


if __name__ == "__main__":
    sys.exit(main())
