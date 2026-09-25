# AKLA Agenda → Transaction Hub bot

Reads everyone's **Daily Agenda** (.docx), links each item to a transaction and a
specific deliverable on the **Transaction Hub** sheet, and fills four columns:

| Column | What it shows |
|---|---|
| **Working On Today** | Who has that deliverable on today's agenda, e.g. `Areen Ali Shah; Sara Khan` |
| **Deliverable Status (Today)** | Each person's "Current Deliverable Status", e.g. `Areen Ali Shah: Pending with me` |
| **Active Today** | `Yes` only on rows where someone is named in Working On Today, otherwise `No`. Filter on this to hide idle work |
| **Agenda Date** | The agenda date these values come from |

It also writes an **Agenda Log** tab listing every agenda item, where it was linked,
and anything it could not match.

## How it reads an agenda

It follows the current template: the `Date:`, the name line (`Senior Counsel:`,
`Associate:`, …), then one block per `Transaction:` with `Scope of Task`,
`Complete Deliverable`, `Day Deliverable` and `Current Deliverable Status`.
Everything from **Daily Timesheet** down is ignored, because that part is the previous day.
`Internal` items are skipped. Labels can be in paragraphs or in a table, and a value
can sit on the line below its label.

## How it matches

* **Transaction**: `M6` matches `M-6 Motorway…` and `PTQ` matches `PTQ – …`, but `M6`
  does not match `M60`. For nicknames the hub doesn't contain, add an alias in `config.toml`.
* **Deliverable**: compares the agenda's Day Deliverable, Complete Deliverable and
  Scope with each deliverable row of that transaction. If nothing is close enough, the
  person is marked on every row of the transaction and the item is flagged in the log.
* Several people on the same deliverable are listed together.
* Unknown transactions are never added to the hub. They are listed under
  "Needs attention" and in the Agenda Log.

The bot finds the header row and the Transaction/Deliverable columns on its own.
If the hub uses different header names, list them in `config.toml` (see
`config.example.toml`). It reuses the output columns if they already exist, and
adds them after the last column the first time it runs.

## Try it locally

```bash
pip install -r requirements.txt
# download a copy of the hub, put today's agendas in ./agendas
python -m agenda_bot local --agendas ./agendas --hub "Transaction Hub.xlsx" --out "Hub (updated).xlsx"
# only agendas for one date:
python -m agenda_bot --date 2026-09-02 local --agendas ./agendas --hub hub.xlsx --out out.xlsx
```

Write to a new file (`--out`) and don't upload it over the live hub.
openpyxl can drop things it doesn't support (charts, some conditional formatting).
The SharePoint mode below edits the live file in place and doesn't have this problem.

## Run it automatically on SharePoint

1. **Agenda folder**: everyone saves their agenda into one SharePoint folder
   (sub-folders per person are fine). The bot opens `.docx` files changed in the last
   3 days and keeps the ones whose `Date:` is today.
2. **App registration** (IT / M365 admin, once): in Entra ID, create an app registration,
   add the Microsoft Graph *application* permission `Sites.ReadWrite.All`, or
   `Sites.Selected` granted on the *akla-internal* site only, and grant admin
   consent. Then create a client secret.
3. **GitHub secrets** (repo → Settings → Secrets → Actions):
   `AZURE_TENANT_ID`, `AZURE_CLIENT_ID`, `AZURE_CLIENT_SECRET`,
   `HUB_URL` (sharing link to the hub workbook), `AGENDA_FOLDER_URL` (sharing link to the agenda folder).
4. Optional: copy `config.example.toml` to `config.toml` to set header names, aliases and thresholds.
5. `.github/workflows/daily-agenda-sync.yml` runs at 11:00 and 15:00 PKT, Monday to Friday.
   You can also run it from the Actions tab and pick a date. Use **dry run** to
   see the matches without writing anything.

The hub is edited through the Excel workbook API, so it stays in place, keeps
its formatting and history, and people can keep it open while the bot runs. Each
run rewrites the four columns, so yesterday's values never carry over.

Run it by hand: `python -m agenda_bot sharepoint --dry-run` (with the environment variables above).

## Tests

```bash
pip install pytest && pytest -q
```
