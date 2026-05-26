# QBExtract — QuickBooks Desktop Data Extractor

A free Python tool that exports QuickBooks Desktop / Enterprise data to a single JSON file via the QuickBooks SDK (QBXML). Built for ERP migrations — used in production to migrate real distributors off QuickBooks Desktop with 60K+ invoice histories, including a corrupt source file with 2,861 unfixable internal link errors.

Outputs **20 entity types** in one bundle:

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

No Intuit cloud account or API key required. Uses the QB SDK, which talks to a locally-running QuickBooks Desktop / Enterprise instance.

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
```

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
    "extractor": "QBExtract v2.0"
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
  "open_ar":            [ /* current unpaid invoices snapshot */ ]
}
```

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
