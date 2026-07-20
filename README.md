# QBExtract — QuickBooks Desktop Data Extractor

A free Python tool that exports QuickBooks Desktop / Enterprise data to a single JSON file via the QuickBooks SDK (QBXML). Built for ERP migrations — used in production to migrate real distributors off QuickBooks Desktop with 60K+ invoice histories, including a corrupt source file with 2,861 unfixable internal link errors.

Outputs **21 entity types** in one bundle:

**Masters**
- Chart of Accounts (with hierarchy via `parent`)
- Customers (with addresses, terms, contact info, **including sub-customers/jobs with `parent`**)
- Items — Inventory, Non-Inventory, **Service**, **OtherCharge**, **Discount** (all 5 SDK item types)
- Price Levels (all levels — first 5 typically map to A-E pricing tiers)
- Quantity Discounts
- Vendors
- Sales Reps
- Payment Terms (Standard + DateDriven)
- **Ship Methods**
- **Payment Methods**
- **Sales Tax Codes** (per-line taxable / non-taxable markers)
- **Sales Tax Items** (per-jurisdiction tax rate items)

**Transactions**
- Invoices (headers + lines + sales tax fields)
- Customer Payments (with invoice application details)
- Credit Memos (with sales tax fields)
- Sales Orders (open + closed)
- Purchase Orders
- Vendor Bills (**with `line_type` distinguishing expense vs item lines**)
- Bill Payments
- Open AR (current open balances)

**Ledger**
- **General Ledger** — every posting line, pulled through QuickBooks' own report engine so QBD generates the implicit balancing entries. Accrual basis by default; QBD can also compute cash basis on request (`--gl-basis cash`/`both`). This is the grain that feeds a Profit & Loss report and transaction-level drill-down.

> **Default scope:** a normal run extracts only the **chart of accounts + General Ledger** — what a Profit & Loss needs. Pass **`--full`** to also pull all the other masters and transaction tables listed above (the complete ERP-migration bundle). The trimmed default runs far fewer SDK queries, so it's faster and much less exposed to the SDK's per-open hang risk.

No Intuit cloud account or API key required. Uses the QB SDK, which talks to a locally-running QuickBooks Desktop / Enterprise instance.

## What's new in v3

- **General Ledger extraction.** The posting-line ledger, pulled via QuickBooks' report engine (`GeneralDetailReportQueryRq` with `GeneralDetailReportType = GeneralLedger`) rather than by re-deriving double-entry from raw transactions. Letting QBD produce the GL means it generates the implicit balancing entries and — with `<ReportBasis>Cash</ReportBasis>` — can compute cash-basis itself, which is otherwise the hardest part of P&L replication. The extractor runs **Accrual basis by default** (`--gl-basis cash`/`both` to also pull the QBD-computed cash basis) and tags every row with its basis. See [General Ledger](#general-ledger) below.
- **GL amounts are exact decimal strings, never floats.** Money display strings (`"1,234.56"`, `"(1,234.56)"`) are parsed to `Decimal` and stored as strings so the GL reconciles to the penny downstream — binary float can't represent decimal cents exactly.
- **`--probe-gl`** — a structure probe for the GL report. The report response is a `<ReportRet>` shape, not the entity `…Ret` blocks the other extractors parse, and whether it exposes each line's internal TxnID is version-dependent. Run the probe against the real company file first: it dumps the raw report XML and reports which columns come back and whether TxnID (needed for drill-down and attachment linking) is present.

## What's new in v2

The v1 release worked for clean QB files but lost data on corrupt files and dropped several entity types entirely. v2 addresses both:

- **ASCII normalization of every text field.** Real QB files contain Windows-1252 smart punctuation, Latin-1 supplement chars, and stray high-bit bytes that the SDK silently rejects on round-trip. v2 decodes XML entities, remaps CP1252 control-range chars to ASCII, and strips anything outside 7-bit ASCII.
- **Full Chart of Accounts.** v1 dropped accounts entirely; v2 extracts the full account hierarchy with `parent` FullName.
- **All item types.** v1 only pulled Inventory + Non-Inventory. v2 also pulls Service, OtherCharge, and Discount items — transactions reference these (Fuel Surcharge, Shipping, Pick-Up Credit, etc.) and silently fail on import if missing.
- **Customer jobs preserved.** v1 dropped sub-customers; v2 keeps them with a `parent` field pointing at the parent's FullName.
- **Ship Methods, Payment Methods, Sales Tax Codes, Sales Tax Items** — new master tables. Transactions reference them; without them, import rejects with "invalid reference" errors.
- **Sales tax fields on invoices + credit memos** — subtotal, sales tax total, tax item, tax code, per-line tax codes.
- **Bill line typing** — `line_type` field distinguishes expense vs item lines (separate XML schemas on push).
- **Date range filters** — `--from-date`, `--to-date`, `--year` for surgical re-extracts.
- **Name-range chunking** for Customers and Items. A corrupt record in one alphabetic bucket only loses that bucket instead of failing the whole query.
- **`--corrupt-safe`** flag for files with internal corruption — narrows queries to non-history fields to avoid SDK calls that walk corrupt transaction ledgers.

## Why this exists

Every QuickBooks Desktop migration hits the same wall: getting your data out cleanly. Intuit's own tools are limited to UI-based exports (CSV per entity, no relationships preserved). Third-party migration services charge $5K-$25K to do what's mostly a scripting problem. Open-source SDK examples exist but require Python + pywin32 + COM glue and never work on the first try.

This tool is what we wished existed when we started doing distribution-ERP migrations. It compiles to a self-contained `.exe` so the customer's machine doesn't need Python installed.

## Requirements

- **Windows** (QB Desktop is Windows-only; this tool inherits that constraint)
- **QuickBooks Desktop / Enterprise 2019 or newer** with the company file you want to export, **opened**
- **QuickBooks SDK** — free download from Intuit. We use [QBSDK 17.0](https://developer.intuit.com/app/developer/qbdesktop/docs/get-started/download-qbdesktop-sdk) or newer. Install on the same machine as QuickBooks.
- **Python 3.11+** (only if running from source — the compiled `.exe` has no Python dependency)

The script uses the QB SDK's COM `RequestProcessor` interface via VBScript, so it works on any Windows install without `pywin32`.

## Usage

### Option 1: Run the EXE (no Python install needed)

Download `QBExtract.exe` from the [Releases](../../releases) page. Run it from any folder; it writes the JSON into the current directory.

### Option 2: Run from source

```bash
git clone https://github.com/josephsprei-lang/quickbooks-desktop-extractor.git
cd quickbooks-desktop-extractor
python QBExtract.py
```

### Command-line options

```
QBExtract.py                                  # interactive
QBExtract.py --years 3                        # last 3 years of transactions
QBExtract.py --years 0                        # all history (chunked by year)
QBExtract.py --year 2024                      # just 2024
QBExtract.py --from-date 2024-01-01 --to-date 2024-06-30
QBExtract.py --corrupt-safe --years 0         # full history, corrupt file
QBExtract.py --output mybundle.json           # custom output path

QBExtract.py --probe-gl --year 2024           # probe GL report structure, then exit
QBExtract.py --year 2024                      # accounts + GL for 2024 (DEFAULT)
QBExtract.py --full --years 0                 # complete ERP bundle (all masters + txns)
QBExtract.py --gl-basis cash                  # GL on QBD-computed cash basis (default: accrual)
QBExtract.py --gl-basis both                  # GL on both accrual and cash
QBExtract.py --years 0 --gl-granularity year  # ALL history, fewest SDK calls (one report/year)
QBExtract.py --gl-granularity quarter         # chunk GL reports by quarter (default: month)
QBExtract.py --no-gl --full                   # ERP bundle without the General Ledger
QBExtract.py --company-file "C:\QB\Company.QBW"   # open a specific file directly
QBExtract.py --year 2024 --no-pause               # no "Press ENTER to exit" — for scripting
```

`--no-pause` suppresses the end-of-run "Press ENTER to exit" prompt so the extractor can be run from a script or scheduled task without blocking. (The prompt otherwise appears only on a bare double-click of the EXE, to keep the console window open.)

### Which company file gets exported

By default the extractor runs against **whatever company file QuickBooks currently has open** — open your file in QuickBooks, then run the tool. QuickBooks only ever has one file open at a time, and the SDK cannot switch it.

`--company-file "C:\path\to\Company.QBW"` opens a specific file directly, but only works when **either**:
- QuickBooks is **closed** and this app has been granted **unattended access** (QuickBooks → Edit → Preferences → Integrated Applications → this app → "Allow access even when QuickBooks is not running"), and the file opens without an interactive login; **or**
- QuickBooks is already **open with that same file** (identical to the default).

If QuickBooks is open with a *different* file and you pass a path, `BeginSession` errors — close and reopen the right file, or use the default (no `--company-file`) with the file already open. The path is most useful for unattended/scheduled runs.

### Either way

1. Open QuickBooks Desktop and load your company file
2. Run `QBExtract.exe` (or `python QBExtract.py`)
3. QuickBooks shows an authorization prompt — **click the link in the body to enable the Yes button**, then click Yes
4. Optionally check "Whenever this company file is open" so it doesn't prompt again
5. Choose how many years of transaction history to pull (`0` = all, `3` = recommended for typical migrations)
6. Wait — large files (60K+ invoices) take 5-30 minutes depending on chunking
7. Output: `<Company Name>_export_<YYYYMMDD>.json` in the current directory

## Migrating corrupt company files

If `Verify Data` / `Rebuild Data` reports unfixable link errors in your source file:

1. Run with `--corrupt-safe` to avoid SDK calls that walk corrupt transaction ledgers
2. Use `--year YYYY` to extract one year at a time — limits the blast radius of any single failed chunk
3. Watch the output for chunks that report `FAILED` or `ERROR`; those are recoverable by re-running just that range
4. Compare totals on QB's own **Sales by Item Summary** report to the JSON sums per year to spot dropped data

## General Ledger

The General Ledger is the one extractor that does **not** read entity `…Ret` blocks. It calls QuickBooks' report engine — `GeneralDetailReportQueryRq` with `GeneralDetailReportType = GeneralLedger` — and parses the report response.

**Why the report engine instead of rebuilding double-entry?** QuickBooks stores transactions (checks, deposits, journal entries, inventory adjustments, …), not a flat ledger. You could pull every posting transaction type and re-implement QBD's posting logic to derive the GL yourself, but that means re-deriving the implicit balancing entries *and* cash-basis conversion — months of work that rarely reconciles exactly. Letting QBD produce the GL means the report already contains the balancing entries, and running it with `<ReportBasis>Cash</ReportBasis>` gets QBD's own cash-basis numbers for free.

**How it runs:**
- Accrual basis by default. `--gl-basis cash` pulls QBD-computed cash basis instead; `--gl-basis both` pulls both, tagging each posting line with its `basis`.
- **By default the extractor pulls only the chart of accounts plus the General Ledger** — exactly what a P&L needs — and skips every other master and transaction table. Because customers and items alone are 70+ name-range queries, this cuts the SDK query count (and thus the number of session opens that can hang) by an order of magnitude. Pass **`--full`** for the complete ERP-migration bundle when you also need those tables.
- Chunked by calendar period (`--gl-granularity month` by default, or `quarter`/`year`). Each chunk is one `ReportPeriod`; a failed period is logged and skipped rather than losing the whole ledger — the same failure-isolation idea as the transaction extractors' year chunks. **Chunking is a transport concern only — it does not affect reporting.** Every row is emitted flat with its own date/account/amount/basis, so you re-aggregate into any period downstream regardless of chunk size; a month-chunked and a year-chunked pull produce identical posting lines. The one exception is `running_balance` (per-chunk, not cumulative across the whole history) — irrelevant for a P&L, which sums `amount`. Smaller chunks isolate failures more finely; larger chunks mean far fewer SDK calls (fewer `BeginSession` opens that can hang), so **use `year` for all-history pulls**.
- Amounts are parsed from the report's display strings to `Decimal` and stored as strings — never `float()`.

**Run the probe first.** The report response format is version-dependent in one important way: whether it exposes each posting line's internal **TxnID**. Downstream drill-down (P&L → transaction detail → attached PDF) and attachment linking need that GUID. Before relying on it, run:

```bash
QBExtract.py --probe-gl --year 2024
```

against the real company file. The probe runs one small GL report, writes the raw response to `<Company>_gl_probe_<YYYYMMDD>.xml`, and prints:
- the report's columns (`ColType` / `ColTitle` → the field each maps to), so you can confirm the column mapping matches your QB version,
- the row-type counts, and
- **whether any `DataRow` carries a `<TxnID>`** — the key question. If it does, drill-down can link directly. If it doesn't, plan for RefNumber-based linking or a supplementary raw-transaction join on `txn_type` + `ref_number` + `date` + `amount`.

The extractor captures TxnID when present and leaves it empty when not, so it works either way — but knowing which case you're in determines how the downstream drill-down is built.

**Getting TxnIDs when the report omits them (`--probe-txn`).** The report has no TxnID column and no flag to add one. To link GL lines to transactions (and thence to attachments), use qbXML's generic `TransactionQueryRq`, which returns `TxnID` + `TxnType` + `RefNumber` + `TxnDate` + `Amount` across all posting types in one query; join each GL line to a transaction on `(RefNumber, Date)` and carry the TxnID onto the row (one TxnID covers all of a transaction's GL lines — the grain an attachment attaches to). Run `QBExtract.py --probe-txn --year 2024` first: it runs one `TransactionQueryRq` plus a GL report for the same window, dumps the raw transaction XML, and reports the empirical join match rate (clean / ambiguous / no-ref) plus the distinct type vocabularies on each side, so you can see how solid the join is and build the type-normalization map from real data before investing in the transaction-index + attachment layer. Note: mapping a TxnID to the actual attachment **file** is a separate, custom problem — the files live in the `<Company> Attach` folder and qbXML does not expose that mapping (it's reachable via the QODBC driver or by parsing the Attach folder).

> **Correctness gate:** before trusting the GL, reconcile one full period to a native QuickBooks P&L (Reports → Company & Financial → Profit & Loss) to the penny. Run `reconcile_pnl.py` on the export (see [Reconciling to QuickBooks](#reconciling-to-quickbooks)) and compare its NET INCOME and section totals against QuickBooks' P&L for the same period and basis.

## Exporting all history (one file per year)

For a full-history GL pull, run **one invocation per year** rather than one giant run. Each invocation is its own QuickBooks session and just 2 SDK calls (accounts + one year-granularity GL report), so it's light on QuickBooks and a hang on one year doesn't cost you the others — you re-run just that year. The output filename is tagged with the period (`<Company>_export_2024.json`), so per-year runs don't overwrite each other.

`export_years.ps1` does the loop for you:

```powershell
# QuickBooks must already be OPEN with the company file loaded. Do NOT use --company-file.
.\export_years.ps1 -StartYear 2015                 # 2015..this year, one file per year
.\export_years.ps1 -StartYear 2018 -EndYear 2024   # explicit range
```

It writes `exports\qbgl_<year>.json` per year, kills and skips any year that hangs past a timeout (default 30 min), and prints a summary with the exact command to re-run any failed year. Or do it by hand:

```powershell
python QBExtract.py --year 2024 --gl-granularity year --no-pause
python QBExtract.py --year 2023 --gl-granularity year --no-pause
...
```

Each yearly file includes the chart of accounts, so `reconcile_pnl.py` works on any of them independently. Chunking is transport-only — a per-year set of files carries the same posting lines as one combined pull, and you re-aggregate across years downstream.

## Reconciling to QuickBooks

`reconcile_pnl.py` reads an export bundle and rebuilds a QuickBooks-style Profit & Loss from the `general_ledger` posting lines, so you can compare it — to the penny — against QuickBooks' own report. It touches no QuickBooks and needs no Windows (it only reads the JSON), so run it anywhere, including on a Mac:

```bash
python reconcile_pnl.py "Switch Broker Network Inc__export_20260720.json"
python reconcile_pnl.py bundle.json --basis cash        # if you pulled cash basis
python reconcile_pnl.py bundle.json --totals-only       # section totals only
python reconcile_pnl.py bundle.json --csv pnl.csv       # also write accounts x basis CSV
```

It prints Income / COGS / Gross Profit / Expense / Net Ordinary Income / Other Income / Other Expense / **Net Income**, grouped and indented by account hierarchy, plus a **Ledger balances** check (the sum of *all* posting amounts must be `0.00` — if it isn't, a chunk failed or rows are missing) and a count of any GL lines whose account type couldn't be classified.

To reconcile: in QuickBooks run **Reports → Company & Financial → Profit & Loss**, set the date range to the export's period and the basis to match, and compare the section totals and Net Income. They should tie exactly.

**Sign convention:** QuickBooks' GL Amount column is debit-positive / credit-negative, so income accounts sum negative and expenses sum positive in the raw ledger. `reconcile_pnl.py` negates income so the P&L reads normally (income and expenses both positive, Net Income = Income − Expenses). If your totals come out sign-flipped, the convention is isolated in one `DISPLAY_SIGN` table at the top of the script.

## Output format

Single JSON file:

```json
{
  "meta": {
    "company": "Example Corp",
    "exported_at": "2026-05-25T14:23:11",
    "years_back": 3,
    "date_range": null,
    "corrupt_safe": false,
    "full": false,
    "gl": true,
    "gl_basis": "accrual",
    "gl_granularity": "month",
    "extractor": "QBExtract v3.0"
  },
  "accounts":           [ /* chart of accounts with parent hierarchy */ ],
  "terms":              [ /* payment terms */ ],
  "ship_methods":       [ /* shipping methods */ ],
  "payment_methods":    [ /* payment methods */ ],
  "sales_tax_codes":    [ /* per-line taxable / non-taxable markers */ ],
  "sales_tax_items":    [ /* per-jurisdiction tax rate items */ ],
  "sales_reps":         [ /* sales reps */ ],
  "price_levels":       [ /* price levels with per-item overrides */ ],
  "quantity_discounts": [ /* discount items */ ],
  "vendors":            [ /* vendors */ ],
  "items":              [ /* all 5 item types: INV NON SVC OCH DSC */ ],
  "customers":          [ /* customers + jobs (jobs have parent field) */ ],
  "invoices":           [ /* invoice headers + lines + tax fields */ ],
  "payments":           [ /* customer payments + invoice application */ ],
  "credit_memos":       [ /* CM headers + lines + tax fields */ ],
  "sales_orders":       [ /* SO headers + lines */ ],
  "purchase_orders":    [ /* PO headers + lines */ ],
  "bills":              [ /* bills with line_type=expense|item */ ],
  "bill_payments":      [ /* bill payment checks */ ],
  "open_ar":            [ /* current unpaid invoices snapshot */ ],
  "general_ledger":     [ /* posting lines, per basis — see below */ ]
}
```

Each `general_ledger` element is one posting line:

```json
{
  "txn_id": "1000-9999999999",
  "txn_type": "Check",
  "ref_number": "1042",
  "date": "2024-03-14",
  "account_full_name": "Rent Expense",
  "account_number": "6000",
  "account_list_id": "80000042-2222222222",
  "account_type": "Expense",
  "name": "Acme Supplies",
  "memo": "Office rent",
  "split": "Checking",
  "class": "",
  "amount": "1200.00",
  "running_balance": "1200.00",
  "basis": "Accrual"
}
```

- `amount` is an **exact decimal string** (never a float) so downstream sums reconcile to the penny. `basis` is `"Accrual"` or `"Cash"`; by default only accrual is pulled, and with `--gl-basis both` every posting line appears once per basis.
- `account_full_name` / `account_number` / `account_list_id` / `account_type` are enriched by joining the report's account back to the extracted chart of accounts — this is what feeds P&L classification (Income / COGS / Expense) and sub-account roll-up. When account numbers are on, the GL labels accounts as `"<number> . <name>"`; the join matches on the parsed account number (then FullName, then Name), so numbered accounts classify correctly.
- `txn_id` is the internal transaction GUID **when the report exposes it** (see the probe note below); otherwise it is empty and downstream must link on `txn_type` + `ref_number` + `date` + `amount`.

Each entity preserves its QB-side relationships via `FullName` references (the QB-native ID system). Customer FullName is unique within a company file; items use FullName for hierarchical groups; accounts likewise.

Run the extractor on your data and inspect the JSON to see exact field-level structure for each entity type.

## Building the EXE yourself

```bash
pip install pyinstaller
pyinstaller --onefile --console QBExtract.py
```

The `build_exe.bat` script does the same. Output appears in `dist/QBExtract.exe`.

## Limitations

- **QuickBooks Desktop / Enterprise only.** Does not work with QuickBooks Online — that uses a completely different API.
- **Read-only.** This tool only reads from QuickBooks. It does not modify, delete, or write back. Safe to run against a production company file.
- **Authorization is per-machine.** The first run of the script triggers a QB authorization dialog where the user grants the script access. Subsequent runs use the saved authorization until QB is reinstalled or the auth is revoked.
- **Sales reps cannot be pushed back via SDK.** QB Desktop's SDK exposes `SalesRepQuery` for reading but has no `SalesRepAdd`. If you're round-tripping the data into a new QB company file, the underlying `OtherNames` entities can be pushed via SDK and then promoted to Sales Reps via an IIF import.
- **Large invoice histories are chunked by year automatically** because QB rejects single-query result sets above ~5,000 records on some configurations.

## What about importing?

This repo is the **extraction half**. Once you have JSON, importing it into a destination system is a separate problem with system-specific schema mapping. We use it to import into PostgreSQL for our ERP (Ask the Ledger), but the JSON is generic enough that you can import it into any system you want — Postgres, MySQL, SQLite, Snowflake, even just Excel for inspection.

If you're round-tripping into another QuickBooks company file (e.g., to escape a corrupt source), be aware of these gotchas:
- **Push parents before children** for accounts, items (Groups), and customers (Jobs). The SDK validates `ParentRef.FullName` at insert time.
- **Push masters before transactions.** Every `*Ref` on a transaction must resolve.
- **Push terms / ship methods / payment methods / sales tax codes / sales tax items** before invoices / CMs / payments — these are silent-failure references.
- **Truncate field values BEFORE escaping**, not after. `xml_escape(value)[:N]` can chop a multi-byte entity (`&amp;` becomes `&am`) and overflow QB's 41-char Name limit.
- **Transactions don't dedupe by RefNumber.** Only Customer / Vendor / Item / Account / Other masters dedupe by Name. If you re-run a transaction push without first deleting prior attempts, you'll get duplicate invoices.
- **`IncludeRetElement`** narrows responses — handy for corrupt files where the default response walks broken ledger entries. Trade-off: any field you exclude returns empty.

If you want guidance on importing into a specific system, open an issue.

## Troubleshooting

- **QB authorization dialog Yes button is greyed out:** click the link in the body of the dialog to enable the Yes button, then optionally check "Whenever this company file is open" so it stops prompting.
- **QB authorization dialog won't respond at all:** restart the PC, reopen QB, try again. This is a known QB SDK quirk with no fix.
- **"Cannot create RequestProcessor":** the QB SDK isn't installed. Download from Intuit (link above) and install on the same machine as QuickBooks.
- **0 invoices exported:** the company file may not have invoices, or the year filter excluded them all. Try with `0` (all history).
- **Export times out / hangs on a corrupt file:** add `--corrupt-safe` to skip SDK calls that walk transaction ledgers. Also try `--year YYYY` to limit the blast radius.
- **"Invalid argument" on push of round-tripped data:** check for non-ASCII characters in fields you didn't normalize. v2's `clean()` handles this on extract; if you're pushing JSON from another source, run it through a similar filter first.

## License

MIT — see [LICENSE](LICENSE).

This is provided as-is, no warranty. Don't use this on a company file without a backup.

## Origin

Built by [Joseph Sprei](https://asktheledger.com) at [Ask the Ledger](https://asktheledger.com), an on-premise ERP for wholesale distributors. We open-sourced the extractor because the import side is where the real value sits — getting data out of QuickBooks shouldn't be a paid consulting engagement.

Issues and PRs welcome. No guarantees on response time — solo project.
