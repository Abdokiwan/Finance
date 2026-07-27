# Finance

Kiwan's Finance Tracker — a single-page web app (`kiwan-finance-app.html`)
that talks over JSONP to a Google Apps Script backend bound to the
`kiwan_finance_tracker` Google Sheet. The Apps Script source (`Code.gs` /
`FinanceEntry.gs`) is maintained directly in the Apps Script editor and is
**not tracked in this repository** — the notes below describe how that
backend is organized and behaves, for reference when editing it there.

## Architecture

- **Two files, one job each.**
  - `FinanceEntry.gs` is the web app: the **only** `doGet`/`doPost` in
    the project, all balance/totals calculations, the Finance-sheet
    label lookups, and the Gemini AI ask/voice/memory endpoints.
  - `Code.gs` is sheet-side automation only: the `onEdit` trigger (auto
    date/month fill, Cash↔Dollar sync rows) and the Dashboard Tools menu.
    It has no `doGet`/`doPost` and calls shared helpers
    (`getCashColMap_`, `getDollarColMap_`, etc.) that live in
    `FinanceEntry.gs` — every `.gs` file in an Apps Script project
    shares one global scope, so this works without any import, as long
    as a given function name is defined in exactly one of the two files.
  - **Never let both files define `doGet`/`doPost`** (or duplicate any
    other function name). That's exactly what used to make the API
    silently serve the wrong JSON schema, depending on which file Apps
    Script happened to load last.
- **Single source of truth**: the `Cash` and `Dollar` sheets (raw ledgers).
  All balances, income/expense totals, and breakdowns are computed from
  these rows — never hardcoded.
- **Summary figures** (Available Balance, Net Worth, Assets, Loans/Debt,
  Dollar price, account balances) are read from the `Finance` and
  `Net worth` sheets by searching for their label text (e.g. "NET WORTH",
  "NBE/Insta"), not a fixed cell address, so the API keeps working if the
  sheet layout shifts — with the bottom-up ledger computation as a
  fallback if a label can't be found.
- **Account aliases** (`ACCOUNT_ALIASES` in `FinanceEntry.gs`) are the one
  place that merges account-name variants in the Cash sheet (`NBE`,
  `NBE/Insta`, `Insta` all merge to the `Insta` display key, since NBE and
  Insta are the same physical account).
- **Month sync**: opening the app with no month explicitly selected shows
  whatever `Dashboard!A7` currently says; picking a month in the app (or
  the sheet) keeps the other side in sync.

## Setup (Apps Script)

1. Open the Apps Script project bound to the spreadsheet.
2. Make sure the project has exactly two files, `Code.gs` and
   `FinanceEntry.gs` (organized as described above), with no other copy
   of `doGet`/`doPost` anywhere.
3. Project Settings ▸ **Script properties** ▸ add:
   - `GEMINI_API_KEY` — your real Gemini API key (never commit this).
   - `GEMINI_MODEL` — optional, defaults to `gemini-2.5-flash`.
   - `MONTHLY_LIMIT` — optional, your monthly budget in EGP (defaults to `15000`).
4. Deploy ▸ Manage deployments ▸ New version (or New deployment) ▸ Web app.
   - Execute as: Me
   - Who has access: Anyone with the link
5. Copy the `/exec` URL into `WEB_APP_URL` in `kiwan-finance-app.html` if
   you created a **new** deployment (a new *version* of an existing
   deployment keeps the same URL).

## Testing

From the Apps Script editor, run these functions and check the log
(View ▸ Logs):

- `testBalances()` — account balances from the Cash sheet.
- `testWealthSummary()` — balances, month totals, dollar holdings/price,
  net worth, and loans in one shot.
- `testGemini()` — sanity-checks the Gemini key/connection.

## Known spreadsheet-side discrepancies

See the audit notes in the project history for a full write-up, but in
short: the `Finance` and `Dashboard` tabs' own "Total Expenses" figures
differ from a bottom-up sum of the Cash ledger by an amount that matches
the "Calls" category almost exactly, in at least one month. This looks
like a pre-existing quirk in the sheet's own formulas rather than a bug in
this code — worth double-checking the SUMIFS/QUERY formula behind that
cell if you want the two to match to the cent.
