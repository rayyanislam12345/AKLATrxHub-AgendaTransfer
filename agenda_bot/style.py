"""House style of the Transactions Hub, applied to the columns the bot adds.

Header cells: navy fill, bold gold Arial 10, centred, thin white dividers, medium
gold top edge. Data cells: no fill, Arial 10, top-aligned and wrapped, thin light
grey borders on every side.
"""

FONT = "Arial"
SIZE = 10
HEADER_FILL = "002060"
HEADER_FONT = "FFC000"
HEADER_DIVIDER = "FFFFFF"
HEADER_TOP = "FFC000"
DATA_BORDER = "BFBFBF"

# Width (Excel characters) and horizontal alignment per output column.
COLUMN_LAYOUT = {
    "working": (24, "left"),
    "status": (30, "left"),
    "active": (12, "center"),
    "date": (18, "center"),
}
LOG_WIDTH = 24


def layout_for(header: str, cfg) -> tuple[int, str]:
    out = cfg.output
    key = {out.working_column: "working", out.status_column: "status",
           out.active_column: "active", out.date_column: "date"}.get(header, "working")
    return COLUMN_LAYOUT[key]


def date_text(d) -> str:
    """Same style as the hub's comments: 'September 25, 2026'."""
    return f"{d:%B} {d.day:02d}, {d.year}"
