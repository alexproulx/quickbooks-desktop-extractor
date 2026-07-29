# Handoff: Build a P&L UI from the QBExtract JSON bundle

## Your task
Build a UI that reads the JSON bundle produced by `QBExtract.py` and reconstructs
**Profit & Loss statements** with:
- account-hierarchy roll-up (QuickBooks-style: Income → COGS → Gross Profit → Expense → Net Ordinary Income → Other Income/Expense → Net Income),
- a **period selector** (year / quarter / month / custom range),
- **period comparison** (e.g. 2024 vs 2023, or month-over-month columns),
- **drill-down**: P&L line → the account's posting lines for the period → individual transaction detail. (A "view the source PDF" step is a *future* capability — design for it, see "TxnID & attachments" below — but it is not wired up yet.)

There is already a reference implementation of the P&L math in **`reconcile_pnl.py`** (Python) in this repo. It reads the same bundle and prints a correct QuickBooks-matching P&L. **Use it as your oracle**: run it on the same file and make your UI's section totals and Net Income match it to the penny. Read it before you start — it encodes every rule below.

**A sample bundle to develop against ships with this repo: `docs/sample_bundle.json`.**
It is synthetic but schema-identical to a real export: 1,244 GL lines across
**3 years (2022–2024), every month populated**, a numbered chart with
sub-accounts, all five P&L account types plus balance-sheet accounts, correct
double-entry (balances to 0.00), both amount signs, ~380 `-SPLIT-` lines, ~310
blank `ref_number` lines, a mix of populated and empty `txn_id`, and four
`class` values. It has a deliberate year-over-year growth trend (Net Income
≈ 84k → 174k → 292k) so period-comparison views show meaningful movement.
Validate against it with `python reconcile_pnl.py docs/sample_bundle.json`
(Net Income for 2022–2024 = 550,657.01; ledger balances to 0.00).

---

## Where the data comes from (context)
QuickBooks Desktop has no cloud API and its file is proprietary, so this data is
pulled through a running QuickBooks instance via Intuit's SDK. The **General
Ledger** is the source of truth: it's the flat list of every posting line, and it
already reconciles to QuickBooks' own P&L to the penny. You are NOT re-deriving
accounting — you are aggregating an already-correct ledger.

The bundle is generated per basis (usually **Accrual**) and may be split
**one file per year** (`qbgl_2020.json`, `qbgl_2021.json`, …) to keep load off
QuickBooks. Your loader should accept one or many files and concatenate their
`general_ledger` arrays (see "Multiple files").

---

## File structure (top level)
```jsonc
{
  "meta": { ... },              // run metadata (see below)
  "accounts": [ ... ],          // chart of accounts (classification + hierarchy)
  "general_ledger": [ ... ],    // THE data: one object per posting line
  // These exist but are EMPTY unless the extract was run with --full (P&L doesn't need them):
  "customers": [], "vendors": [], "items": [], "invoices": [], "payments": [],
  "credit_memos": [], "sales_orders": [], "purchase_orders": [], "bills": [],
  "bill_payments": [], "open_ar": [], "terms": [], "ship_methods": [],
  "payment_methods": [], "sales_tax_codes": [], "sales_tax_items": [],
  "sales_reps": [], "price_levels": [], "quantity_discounts": []
}
```
For the P&L you need exactly two arrays: **`accounts`** and **`general_ledger`**.

### `meta`
```jsonc
{
  "company": "Switch Broker Network Inc.",
  "company_file": "",              // path if opened directly
  "exported_at": "2026-07-20T...", // ISO timestamp of the extract
  "date_range": ["2024-01-01","2024-12-31"], // or null when --years was used
  "years_back": 0,                 // 0 = all history
  "gl_basis": "accrual",           // "accrual" | "cash" | "both"
  "gl_granularity": "year",        // extraction chunk size — IRRELEVANT to you
  "full": false,                   // whether the ERP tables were included
  "extractor": "QBExtract v3.0"
}
```
`gl_granularity` is a transport detail of the extract; it does **not** affect the
data. Rows are flat and dated; aggregate into any period you like.

### `accounts` — the chart of accounts (classification + hierarchy)
```jsonc
{
  "list_id": "80000012-1234567890",
  "name": "Consulting Income",          // short name
  "full_name": "Consulting Income",     // hierarchical: "Parent:Child:Grandchild"
  "parent": "",                         // parent's full_name (empty if top-level)
  "account_type": "Income",             // QB AccountType enum — see below
  "special_type": "",                   // SpecialAccountType, usually empty
  "account_number": "4000",             // may be "" if numbers aren't used
  "description": "",
  "active": true
}
```

### `general_ledger` — the posting lines (THE data)
```jsonc
{
  "txn_id": "",                         // internal transaction GUID — see caveat
  "txn_type": "Cheque",                 // QB display type (LOCALIZED — see gotcha)
  "ref_number": "1042",                 // the "Num" column; often "" (deposits, JEs)
  "date": "2024-03-14",                 // YYYY-MM-DD — aggregate by this
  "account_full_name": "Consulting Income", // clean master name (number stripped)
  "account_number": "4000",             // the account this line posts to
  "account_list_id": "",                // may be "" (this report gives no ListID)
  "account_type": "Income",             // ENRICHED QB AccountType — classify on this
  "name": "Acme Corp",                  // entity/payee display name (for drill-down)
  "memo": "March retainer",
  "split": "1010 . Operating",          // the "Split" column: other side, or "-SPLIT-"
  "class": "",                          // QB class, if used (optional P&L dimension)
  "amount": "-1000.00",                 // EXACT DECIMAL STRING — see gotchas
  "running_balance": "",                // per-chunk, NOT cumulative — IGNORE
  "basis": "Accrual"                    // "Accrual" or "Cash"
}
```

---

## Critical gotchas (read these or your P&L will be wrong)

1. **`amount` is a decimal STRING, not a number.** It is stored as a string on
   purpose so it reconciles to the penny. Do **not** `parseFloat` and sum — you
   will accumulate binary-float error across ~20k+ rows. Sum in **integer cents**
   (`Math.round(parseFloat(a)*100)` per row, add integers, divide at the end) or
   use a decimal library. Empty string `""` means no amount → treat as 0.

2. **Sign convention: debit-positive / credit-negative.** This is QuickBooks'
   GL convention. Consequence for the P&L:
   - **Income / OtherIncome** accounts sum **negative** in the raw ledger.
   - **Expense / COGS / OtherExpense** accounts sum **positive**.
   To display a normal P&L (everything positive, Net Income = Income − Expenses),
   **negate income**. The exact per-type display sign (from `reconcile_pnl.py`):
   ```
   Income: -1   OtherIncome: -1   CostOfGoodsSold: +1   Expense: +1   OtherExpense: +1
   display_amount = sign[account_type] * raw_sum
   ```

3. **Classify on `account_type`, which is already on each GL row.** The five
   P&L types are `Income`, `CostOfGoodsSold`, `Expense`, `OtherIncome`,
   `OtherExpense`. Every other type (`Bank`, `AccountsReceivable`,
   `AccountsPayable`, `CreditCard`, `Equity`, `FixedAsset`, `OtherAsset`,
   `OtherCurrentAsset`, `OtherCurrentLiability`, `LongTermLiability`) is a balance
   sheet account — **exclude it from the P&L**. `NonPosting` never appears in the GL.

4. **Amounts are line-level; roll up sub-accounts into parents.** `full_name`
   uses colon notation (`"Auto:Fuel"`), and `parent` points at the parent's
   `full_name`. QuickBooks' P&L shows the parent's total *including* its children,
   with children indented beneath. Build the account tree from `accounts` and
   aggregate child totals up to parents.

5. **Join GL rows to `accounts` by `account_number` first.** When a company uses
   account numbers (this one does), match `gl.account_number` →
   `accounts[].account_number`, then fall back to `full_name`/`name`. The GL rows
   are already enriched with `account_type`, so you usually don't need the join
   for classification — but you need `accounts` for the **hierarchy** (parent/child).

6. **`basis`: never mix.** Filter to one basis (usually `"Accrual"`). If
   `meta.gl_basis` is `"both"`, every posting line appears **twice** (once per
   basis) — always filter to a single basis before summing, and offer a basis
   toggle only when both are present.

7. **`txn_type` is a localized display string.** You'll see values like
   `"Cheque"`, `"Bill Pmt -Check"`, `"General Journal"`, `"Deposit"`,
   `"Invoice"`. Use it for display/grouping, but don't hard-code a canonical
   list — read the distinct values from the data.

8. **`ref_number` is often empty** (deposits, journal entries, many bank
   transactions). Don't use it as a primary key.

9. **`running_balance` is per-extraction-chunk, not cumulative.** Ignore it for
   the P&L; if you ever show a running balance, recompute it yourself from
   ordered `amount`s.

10. **Data-integrity check:** the sum of **all** `amount`s for a single basis
    across the whole file ≈ `0.00` (double-entry). If it isn't, a chunk failed
    during extraction or rows are missing — surface a warning.

---

## Building the P&L (mirror `reconcile_pnl.py`)
For a chosen basis and date range:
1. Filter `general_ledger` to `basis == chosen` and `from <= date <= to`.
2. Group by account (`account_number` or `account_full_name`); sum `amount` (in cents).
3. Apply the display sign per `account_type` (step 2 above) to get each account's P&L amount.
4. Aggregate into sections and roll sub-accounts up to parents.
5. Compute:
   ```
   Total Income        = Σ display(Income accounts)
   Total COGS          = Σ display(CostOfGoodsSold accounts)
   Gross Profit        = Total Income − Total COGS
   Total Expense       = Σ display(Expense accounts)
   Net Ordinary Income = Gross Profit − Total Expense
   Total Other Income  = Σ display(OtherIncome accounts)
   Total Other Expense = Σ display(OtherExpense accounts)
   Net Income          = Net Ordinary Income + Total Other Income − Total Other Expense
   ```
Validate: run `python reconcile_pnl.py <file>` and match its output.

---

## Drill-down design
The GL row already *is* the transaction-line detail, so drill-down is just
progressive filtering of `general_ledger`:

1. **P&L summary** — section/account totals (above).
2. **Account detail** — click an account line → all GL rows for that
   `account_number` in the selected period. Columns map directly:
   Date=`date`, Type=`txn_type`, Num=`ref_number`, Name=`name`, Memo=`memo`,
   Split=`split`, Amount=`amount`.
3. **Transaction detail** — a single posting event. Since one QuickBooks
   transaction spans multiple GL lines (both sides of the entry), group rows that
   belong to the same transaction. **Today there is no reliable per-line key for
   this** (see below), so group approximately on `(date, txn_type, ref_number,
   name)` when you need a transaction-level view, and treat it as best-effort.

### TxnID & attachments (future — design for it now)
`txn_id` is currently **empty**: QuickBooks' report engine doesn't expose the
internal transaction GUID. A follow-up extractor feature (`--txn-index`, built on
`TransactionQueryRq`) will populate `txn_id` on each GL row by joining on
RefNumber+Date. The end goal is a "View source PDF" action that resolves a
transaction to its attached document. So:
- Key your transaction-detail view on `txn_id` **when present**, falling back to
  the `(date, txn_type, ref_number, name)` heuristic when it's empty.
- Leave a placeholder "View attachment" affordance that lights up once `txn_id`
  is populated. (The attachment *file* mapping is a separate backend workstream —
  the UI just needs a stable `txn_id` to request against.)

---

## Period selector & comparison
Every row has a `date` (`YYYY-MM-DD`), so all of this is client-side date
filtering + group-by — no extraction changes needed:
- **Period selector:** presets (This Year, Last Year, This Quarter, custom range)
  → filter rows by `date`. Also support fiscal-year offset if the company's fiscal
  year isn't January (ask; SBN's is likely calendar).
- **Comparison:** render N periods as columns. For each period, run the same
  aggregation over that period's date filter; align rows by account. Common views:
  year-over-year, quarter-by-quarter, month-by-month, actual vs prior.
- **Class dimension (optional):** if `class` is populated, you can offer a
  class/department P&L by adding `class` to the group-by.

---

## Multiple files
History is likely delivered as one file per year (`qbgl_<year>.json`). Loader:
- Concatenate every file's `general_ledger` into one array.
- `accounts` is a current snapshot and is (nearly) identical in every file —
  dedupe by `list_id`/`account_number` and keep one copy (prefer the newest file's).
- `meta` differs per file; keep the set or the newest.
After merge you have a single flat ledger spanning all years to slice by period.

## Performance
Expect ~20k GL rows per year (~100k–200k for a decade). That's fine to hold and
aggregate in the browser, but:
- Parse amounts to integer cents once at load, not per render.
- Pre-index rows by `(basis, account_number)` and by year/month for fast
  period slicing.
- Memoize per-(period, basis) aggregations for the comparison view.

## Validation checklist before you ship
- [ ] Your Net Income for a full year matches `reconcile_pnl.py` (and QuickBooks) to the penny.
- [ ] Sum of all amounts for one basis ≈ 0.00 (integrity check surfaced).
- [ ] Sub-account totals roll into parents; indentation matches QuickBooks.
- [ ] Switching basis (if both present) never double-counts.
- [ ] Period comparison columns each reconcile independently.
- [ ] Drill-down from a P&L line reaches the exact posting rows that sum to that line.
