"""Microsoft Graph backend: read agendas from SharePoint and edit the hub workbook in place.

Edits go through the Excel workbook API, so the file stays in SharePoint, keeps its
formatting and version history, and people can keep it open while the bot runs.

Needs an Entra ID (Azure AD) app registration with the *application* permission
Sites.ReadWrite.All (or Sites.Selected granted on the akla-internal site) and admin
consent. Credentials come from AZURE_TENANT_ID / AZURE_CLIENT_ID / AZURE_CLIENT_SECRET.
"""

from __future__ import annotations

import base64
import re
import time
from datetime import date, datetime, timedelta, timezone
from urllib.parse import quote

import requests

from .agenda_parser import Agenda, parse_agenda
from . import style
from .config import Config
from .hub import LOG_HEADERS, HubTable, UpdatePlan

GRAPH = "https://graph.microsoft.com/v1.0"

# Rows below the last hub row whose inherited formatting is cleared as well.
CLEAR_BELOW = 1000


def share_id(url: str) -> str:
    """Encode a SharePoint sharing link for the /shares endpoint."""
    encoded = base64.urlsafe_b64encode(url.strip().encode()).decode().rstrip("=")
    return "u!" + encoded


def col_letter(n: int) -> str:
    letters = ""
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def col_number(letters: str) -> int:
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n


class GraphClient:
    def __init__(self, tenant_id: str, client_id: str, client_secret: str):
        if not (tenant_id and client_id and client_secret):
            raise ValueError("AZURE_TENANT_ID, AZURE_CLIENT_ID and AZURE_CLIENT_SECRET must be set")
        resp = requests.post(
            f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
            data={"grant_type": "client_credentials", "client_id": client_id,
                  "client_secret": client_secret,
                  "scope": "https://graph.microsoft.com/.default"},
            timeout=30,
        )
        resp.raise_for_status()
        self.session = requests.Session()
        self.session.headers["Authorization"] = f"Bearer {resp.json()['access_token']}"

    def request(self, method: str, path: str, **kwargs) -> requests.Response:
        url = path if path.startswith("http") else GRAPH + path
        for attempt in range(5):
            resp = self.session.request(method, url, timeout=60, **kwargs)
            if resp.status_code in (429, 503, 504) and attempt < 4:
                time.sleep(int(resp.headers.get("Retry-After", 2 ** (attempt + 1))))
                continue
            if not resp.ok:
                raise RuntimeError(f"Graph {method} {path} failed: {resp.status_code} {resp.text[:500]}")
            return resp
        raise RuntimeError("unreachable")

    def get(self, path: str, **kw) -> dict:
        return self.request("GET", path, **kw).json()

    def resolve_share(self, url: str) -> dict:
        return self.get(f"/shares/{share_id(url)}/driveItem")


def load_agendas_from_sharepoint(client: GraphClient, folder_url: str, on_date: date | None,
                                 lookback_days: int = 3) -> list[Agenda]:
    folder = client.resolve_share(folder_url)
    drive_id = folder["parentReference"]["driveId"]
    cutoff = datetime.now(timezone.utc) - timedelta(days=lookback_days)
    agendas: list[Agenda] = []

    def walk(item_id: str) -> None:
        url = f"/drives/{drive_id}/items/{item_id}/children?$top=200"
        while url:
            page = client.get(url)
            for child in page.get("value", []):
                if "folder" in child:
                    walk(child["id"])
                    continue
                name = child.get("name", "")
                if not name.lower().endswith(".docx") or name.startswith("~$"):
                    continue
                modified = datetime.fromisoformat(child["lastModifiedDateTime"].replace("Z", "+00:00"))
                if modified < cutoff:
                    continue
                content = client.request("GET", f"/drives/{drive_id}/items/{child['id']}/content").content
                agenda = parse_agenda(content, filename=name)
                if on_date and agenda.agenda_date and agenda.agenda_date != on_date:
                    continue
                agendas.append(agenda)
            url = page.get("@odata.nextLink")

    walk(folder["id"])
    return agendas


class GraphWorkbook:
    def __init__(self, client: GraphClient, hub_url: str):
        self.client = client
        item = client.resolve_share(hub_url)
        self.base = f"/drives/{item['parentReference']['driveId']}/items/{item['id']}/workbook"
        session = client.request("POST", f"{self.base}/createSession",
                                 json={"persistChanges": True}).json()
        self.headers = {"workbook-session-id": session["id"]}

    def close(self) -> None:
        try:
            self.client.request("POST", f"{self.base}/closeSession", headers=self.headers)
        except RuntimeError:
            pass

    def _sheet(self, name: str) -> str:
        return f"{self.base}/worksheets('{quote(name.replace(chr(39), chr(39) * 2), safe='')}')"

    def sheet_names(self) -> list[str]:
        sheets = self.client.get(f"{self.base}/worksheets?$select=name,position", headers=self.headers)
        return [s["name"] for s in sorted(sheets["value"], key=lambda s: s["position"])]

    def read_table(self, sheet: str) -> HubTable:
        rng = self.client.get(f"{self._sheet(sheet)}/usedRange(valuesOnly=true)?$select=address,values",
                              headers=self.headers)
        start = rng["address"].split("!")[-1].split(":")[0].replace("$", "")
        m = re.match(r"([A-Z]+)(\d+)", start)
        return HubTable(rows=rng["values"], first_row=int(m.group(2)), first_col=col_number(m.group(1)))

    def write(self, sheet: str, top_left_col: int, top_row: int, values: list[list]) -> None:
        if not values:
            return
        width = max(len(r) for r in values)
        values = [[("" if v is None else v) for v in r] + [""] * (width - len(r)) for r in values]
        address = (f"{col_letter(top_left_col)}{top_row}:"
                   f"{col_letter(top_left_col + width - 1)}{top_row + len(values) - 1}")
        self.client.request("PATCH", f"{self._sheet(sheet)}/range(address='{address}')",
                            headers=self.headers, json={"values": values})

    def _table_at(self, sheet: str, header_row: int) -> tuple[str, int, int, int] | None:
        """(name, first col, last col, last row) of the Excel Table whose header is header_row."""
        tables = self.client.get(f"{self._sheet(sheet)}/tables?$select=name", headers=self.headers)
        for table in tables.get("value", []):
            name = quote(table["name"], safe="")
            rng = self.client.get(f"{self.base}/tables('{name}')/range?$select=address",
                                  headers=self.headers)
            ref = rng["address"].split("!")[-1].replace("$", "")
            m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", ref)
            if m and int(m.group(2)) == header_row:
                return table["name"], col_number(m.group(1)), col_number(m.group(3)), int(m.group(4))
        return None

    def apply_plan(self, sheet: str, plan: UpdatePlan, log_sheet: str, cfg: Config | None = None) -> None:
        table = self._table_at(sheet, plan.header_row)
        for col in sorted(plan.columns, key=lambda c: c.column):
            values = [[col.header]] + [[v] for v in col.values]
            if table and col.column == table[2] + 1:
                # Add as a new column of the Excel Table so its filter covers it.
                name, first_col, last_col, last_row = table
                rows = last_row - plan.header_row + 1
                values = (values + [[None]] * rows)[:rows]
                values = [["" if v is None else v for v in r] for r in values]
                self.client.request("POST", f"{self.base}/tables('{quote(name, safe='')}')/columns/add",
                                    headers=self.headers, json={"index": None, "values": values})
                table = (name, first_col, last_col + 1, last_row)
            else:
                self.write(sheet, col.column, col.header_row, values)
        if log_sheet:
            if log_sheet in self.sheet_names():
                self.client.request("POST", f"{self._sheet(log_sheet)}/range(address='A1:Z5000')/clear",
                                    headers=self.headers, json={"applyTo": "Contents"})
            else:
                self.client.request("POST", f"{self.base}/worksheets/add",
                                    headers=self.headers, json={"name": log_sheet})
            self.write(log_sheet, 1, 1, [LOG_HEADERS] + plan.log_rows)
        self.format_plan(sheet, plan, log_sheet, cfg or Config())

    # ---- formatting -------------------------------------------------------------

    def _fmt(self, sheet: str, address: str, part: str, body: dict) -> None:
        self.client.request("PATCH", f"{self._sheet(sheet)}/range(address='{address}')/format{part}",
                            headers=self.headers, json=body)

    def _borders(self, sheet: str, address: str, sides: list[str], color: str,
                 weight: str = "Thin") -> None:
        for side in sides:
            self._fmt(sheet, address, f"/borders/{side}",
                      {"style": "Continuous", "weight": weight, "color": f"#{color}"})

    def _style_header(self, sheet: str, address: str) -> None:
        self._fmt(sheet, address, "", {"horizontalAlignment": "Center",
                                       "verticalAlignment": "Center", "wrapText": True})
        self._fmt(sheet, address, "/fill", {"color": f"#{style.HEADER_FILL}"})
        self._fmt(sheet, address, "/font", {"name": style.FONT, "size": style.SIZE, "bold": True,
                                            "italic": False, "color": f"#{style.HEADER_FONT}"})
        self._borders(sheet, address, ["EdgeLeft", "EdgeRight", "InsideVertical"], style.HEADER_DIVIDER)
        self._borders(sheet, address, ["EdgeTop"], style.HEADER_TOP, "Medium")

    def _style_data(self, sheet: str, address: str, align: str) -> None:
        self.client.request("POST", f"{self._sheet(sheet)}/range(address='{address}')/format/fill/clear",
                            headers=self.headers)
        self._fmt(sheet, address, "", {"horizontalAlignment": align.capitalize(),
                                       "verticalAlignment": "Top", "wrapText": True})
        self._fmt(sheet, address, "/font", {"name": style.FONT, "size": style.SIZE, "bold": False,
                                            "italic": False, "color": "#000000"})
        self._borders(sheet, address, ["EdgeLeft", "EdgeRight", "EdgeTop", "EdgeBottom",
                                       "InsideVertical", "InsideHorizontal"], style.DATA_BORDER)

    def format_plan(self, sheet: str, plan: UpdatePlan, log_sheet: str, cfg: Config) -> None:
        """Make the bot's columns look like the rest of the hub (no cell colours)."""
        # Excel copies the neighbouring column's formatting, including the STATUS
        # column's conditional formatting, into columns added to the table. Clear all
        # formats below the headers (conditional ones included) before restyling.
        cols = sorted(c.column for c in plan.columns)
        if cols:
            span = (f"{col_letter(cols[0])}{plan.header_row + 1}:"
                    f"{col_letter(cols[-1])}{max(plan.last_row, plan.header_row + 1) + CLEAR_BELOW}")
            self.client.request("POST", f"{self._sheet(sheet)}/range(address='{span}')/clear",
                                headers=self.headers, json={"applyTo": "Formats"})
        for col in plan.columns:
            letter = col_letter(col.column)
            width, align = style.layout_for(col.header, cfg)
            self._style_header(sheet, f"{letter}{plan.header_row}")
            self._fmt(sheet, f"{letter}{plan.header_row}", "", {"columnWidth": width * 7})
            if plan.last_row > plan.header_row:
                self._style_data(sheet, f"{letter}{plan.header_row + 1}:{letter}{plan.last_row}", align)
        if log_sheet:
            last = col_letter(len(LOG_HEADERS))
            self._style_header(log_sheet, f"A1:{last}1")
            self._fmt(log_sheet, f"A1:{last}1", "", {"columnWidth": style.LOG_WIDTH * 7})
            if plan.log_rows:
                self._style_data(log_sheet, f"A2:{last}{len(plan.log_rows) + 1}", "left")
