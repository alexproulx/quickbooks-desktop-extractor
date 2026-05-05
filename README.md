# QBExtract — QuickBooks Desktop Data Extractor

A free Python tool that exports QuickBooks Desktop / Enterprise data to a single JSON file via the QuickBooks SDK (QBXML). Built for ERP migrations — used in production to migrate real distributors off QuickBooks Desktop with 60K+ invoice histories.

Outputs **15 entity types** in one bundle:

- Customers (with addresses, terms, contact info)
- Items (with cost, on-hand, sales price, multiple price levels)
- Price Levels (all levels — first 5 typically map to A-E pricing tiers)
- Quantity Discounts
- Vendors
- Sales Reps
- Payment Terms
- Invoices (headers + all line items)
- Payments (received from customers)
- Credit Memos
- Sales Orders (open)
- Purchase Orders (open)
- Bills (vendor invoices)
- Bill Payments
- Open AR (current open balances)

No Intuit cloud account or API key required. Uses the QB SDK, which talks to a locally-running QuickBooks Desktop / Enterprise instance.

## Why this exists

Every QuickBooks Desktop migration hits the same wall: getting your data out cleanly. Intuit's own tools are limited to UI-based exports (CSV per entity, no relationships preserved). Third-party migration services charge $5K-$25K to do what's mostly a scripting problem. Open-source SDK examples exist but require Python + pywin32 + COM glue and never work on the first try.

This tool is what we wished existed when we started doing distribution-ERP migrations. It's been used on three real client migrations (bakeries and electronics distributors). It compiles to a self-contained `.exe` so the customer's machine doesn't need Python installed.

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

### Either way

1. Open QuickBooks Desktop and load your company file
2. Run `QBExtract.exe` (or `python QBExtract.py`)
3. QuickBooks shows an authorization prompt — click **Yes / Allow**
4. Choose how many years of invoice history to pull (`0` = all, `3` = recommended for typical migrations)
5. Wait — large files (60K+ invoices) take 5-10 minutes
6. Output: `<Company Name>_export_<YYYYMMDD>.json` in the current directory

## Output format

Single JSON file with this top-level structure:

```json
{
  "company_name": "Example Corp",
  "exported_at": "2026-05-05T14:23:11",
  "qb_version": "2024",
  "customers": [ /* ... */ ],
  "items": [ /* ... */ ],
  "price_levels": [ /* ... */ ],
  "quantity_discounts": [ /* ... */ ],
  "vendors": [ /* ... */ ],
  "sales_reps": [ /* ... */ ],
  "terms": [ /* ... */ ],
  "invoices": [ /* ... */ ],
  "payments": [ /* ... */ ],
  "credit_memos": [ /* ... */ ],
  "sales_orders": [ /* ... */ ],
  "purchase_orders": [ /* ... */ ],
  "bills": [ /* ... */ ],
  "bill_payments": [ /* ... */ ],
  "open_ar": [ /* ... */ ]
}
```

Each entity preserves its QB-side relationships via `FullName` references (the QB-native ID system). Customer FullName is unique within a company file; items use FullName for hierarchical groups; etc.

Run the extractor on your data and inspect the JSON to see exact field-level structure for each entity type.

## Building the EXE yourself

```bash
pip install pyinstaller
pyinstaller --onefile --console QBExtract.py
```

The `build_exe.bat` script does the same. Output appears in `dist/QBExtract.exe`.

## Limitations

- **QuickBooks Desktop / Enterprise only.** Does not work with QuickBooks Online — that uses a completely different API. (If demand exists, a QBO version is feasible but is a separate project.)
- **Read-only.** This tool only reads from QuickBooks. It does not modify, delete, or write back. Safe to run against a production company file.
- **Authorization is per-machine.** The first run of the script triggers a QB authorization dialog where the user grants the script access. Subsequent runs use the saved authorization until QB is reinstalled or the auth is revoked.
- **Some entities are limited by QB SDK constraints.** Notably, large invoice histories are chunked by year automatically because QB rejects single-query result sets above ~5,000 records.

## What about importing?

This repo is the **extraction half**. Once you have JSON, importing it into a destination system is a separate problem with system-specific schema mapping. We use it to import into PostgreSQL for our ERP (Ask the Ledger), but the JSON is generic enough that you can import it into any system you want — Postgres, MySQL, SQLite, Snowflake, even just Excel for inspection.

If you want guidance on importing into a specific system, open an issue.

## Troubleshooting

- **QB authorization dialog won't respond:** restart the PC, reopen QB, try again. This is a known QB SDK quirk with no fix.
- **"Cannot create RequestProcessor":** the QB SDK isn't installed. Download from Intuit (link above) and install on the same machine as QuickBooks.
- **0 invoices exported:** the company file may not have invoices, or the year filter excluded them all. Try with `0` (all history).
- **Export times out / hangs:** large QB files with 50K+ invoices are chunked by year automatically. If a single year still times out, the chunk is skipped and logged. You can re-run with a smaller year range.

## License

MIT — see [LICENSE](LICENSE).

This is provided as-is, no warranty. Don't use this on a company file without a backup.

## Origin

Built by [Joseph Sprei](https://asktheledger.com) at [Ask the Ledger](https://asktheledger.com), an on-premise ERP for wholesale distributors. We open-sourced the extractor because the import side is where the real value sits — getting data out of QuickBooks shouldn't be a paid consulting engagement.

Issues and PRs welcome. No guarantees on response time — solo project.
