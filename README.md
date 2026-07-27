# Finance

Kiwan's Finance Tracker — a Google Apps Script backend (`Code.gs`) bound to
the `kiwan_finance_tracker` Google Sheet, and a single-page web app
(`kiwan-finance-app.html`) that talks to it over JSONP.

## Architecture

- **Single source of truth**: the `Cash` and `Dollar` sheets (raw ledgers).
  All balances, income/expense totals, and breakdowns are computed from
  these rows — never hardcoded.
- **Summary figures** (Available Balance, Net Worth, Assets, Loans/Debt,
  Dollar price) are read from the `Finance` and `Net worth` sheets by
  searching for their label text (e.g. "NET WORTH"), not a fixed cell
  address, so the API keeps working if the sheet layout shifts.
- **One `doGet` / one `doPost`.** There used to be a second, conflicting
  copy of both in `FinanceEntry.gs`; that file has been merged into
  `Code.gs` and removed. Do not add a second `doGet`/`doPost` anywhere —
  Apps Script silently lets the last-loaded one win, which is what caused
  the API to intermittently return the wrong schema.
- **Account aliases** (`ACCOUNT_ALIASES` in `Code.gs`) are the one place
  that merges account-name variants in the Cash sheet (`NBE`, `NBE/Insta`,
  `Insta` all merge to the `Insta` display key, since NBE and Insta are
  the same physical account).

## Setup (Apps Script)

1. Open the Apps Script project bound to the spreadsheet.
2. Make sure `Code.gs` is the **only** file with the ledger logic (delete
   any other copy of `doGet`/`doPost` if one exists).
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
