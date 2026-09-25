"""Configuration loading (TOML) with defaults that fit the Transaction Hub sheet."""

from __future__ import annotations

import os
import tomllib
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class HubConfig:
    worksheet: str = ""              # blank = first worksheet
    header_row: int = 0              # 1-based; 0 = auto-detect
    transaction_columns: list[str] = field(default_factory=lambda: [
        "Transaction", "Transaction Name", "Transactions", "Matter", "Project", "Deal",
        "Client / Transaction", "Transaction / Matter",
    ])
    deliverable_columns: list[str] = field(default_factory=lambda: [
        "Deliverable", "Deliverables", "Specific Deliverable", "Deliverable Name",
        "Task", "Tasks", "Document", "Workstream",
    ])
    # Optional client column; agendas may name the client instead of the project.
    client_columns: list[str] = field(default_factory=lambda: ["Client", "Client Name"])
    # The hub's own status column, read for waiting_statuses below.
    status_columns: list[str] = field(default_factory=lambda: ["Status"])
    fill_down_transaction: bool = True  # transaction only written on first row of a group


@dataclass
class OutputConfig:
    working_column: str = "Working On Today"
    status_column: str = "Deliverable Status (Today)"
    active_column: str = "Active Today"
    date_column: str = "Agenda Date"   # blank string disables this column
    active_yes: str = "Yes"
    active_no: str = "No"
    log_sheet: str = "Agenda Log"      # blank string disables the log sheet
    # Hub statuses meaning the deliverable sits with someone outside the team: those rows
    # are always Active Today, and Working On Today / Deliverable Status stay blank.
    waiting_statuses: list[str] = field(default_factory=lambda: [
        "Deliverable With Boss", "Deliverable With Client",
    ])


@dataclass
class MatchingConfig:
    transaction_threshold: float = 0.8
    deliverable_threshold: float = 0.5


@dataclass
class SharePointConfig:
    hub_url: str = ""
    agenda_folder_url: str = ""
    tenant_id: str = ""
    client_id: str = ""
    client_secret: str = ""


@dataclass
class Config:
    hub: HubConfig = field(default_factory=HubConfig)
    output: OutputConfig = field(default_factory=OutputConfig)
    matching: MatchingConfig = field(default_factory=MatchingConfig)
    sharepoint: SharePointConfig = field(default_factory=SharePointConfig)
    aliases: dict[str, list[str]] = field(default_factory=dict)

    def aliases_for(self, hub_transaction: str) -> list[str]:
        """Aliases are keyed by the hub's transaction name (case-insensitive)."""
        key = hub_transaction.strip().lower()
        for name, values in self.aliases.items():
            if name.strip().lower() == key:
                return [values] if isinstance(values, str) else list(values)
        return []


def _apply(section, data: dict) -> None:
    for key, value in data.items():
        if not hasattr(section, key):
            raise ValueError(f"unknown config key: {key}")
        setattr(section, key, value)


def load_config(path: str | Path | None) -> Config:
    cfg = Config()
    if path:
        with open(path, "rb") as fh:
            data = tomllib.load(fh)
        for name in ("hub", "output", "matching", "sharepoint"):
            if name in data:
                _apply(getattr(cfg, name), data[name])
        cfg.aliases = data.get("aliases", {})

    sp = cfg.sharepoint
    sp.tenant_id = os.environ.get("AZURE_TENANT_ID") or sp.tenant_id
    sp.client_id = os.environ.get("AZURE_CLIENT_ID") or sp.client_id
    sp.client_secret = os.environ.get("AZURE_CLIENT_SECRET") or sp.client_secret
    sp.hub_url = os.environ.get("HUB_URL") or sp.hub_url
    sp.agenda_folder_url = os.environ.get("AGENDA_FOLDER_URL") or sp.agenda_folder_url
    return cfg
