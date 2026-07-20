"""
reconcile_pnl.py — P&L reconciliation helper for QBExtract bundles
==================================================================
Reads a QBExtract JSON bundle and prints a QuickBooks-style Profit & Loss
built from the `general_ledger` posting lines, so you can compare it — to the
penny — against QuickBooks' own report (Reports > Company & Financial > Profit
& Loss) for the same period and basis.

This is the correctness gate for the GL extract. It touches no QuickBooks and
needs no Windows: it only reads the exported JSON, so run it anywhere (Mac,
Linux, Windows).

Usage:
  python reconcile_pnl.py "Company_export_20260720.json"
  python reconcile_pnl.py bundle.json --basis accrual      # one basis
  python reconcile_pnl.py bundle.json --totals-only        # section totals only
  python reconcile_pnl.py bundle.json --csv pnl.csv        # also write a CSV

Sign convention:
  QuickBooks' GL Amount column is debit-positive / credit-negative, so income
  accounts (credit-natural) sum NEGATIVE and expense accounts sum POSITIVE in
  the raw ledger. This tool negates income so the P&L reads the familiar way
  (income and expenses both shown positive; Net Income = Income - Expenses).
  If your totals come out sign-flipped vs QuickBooks, tell me — the convention
  is isolated in DISPLAY_SIGN below.
"""

import argparse
import json
import sys
from decimal import Decimal, InvalidOperation


# QB AccountType values that belong on a P&L, in report order, with the sign
# that converts the raw (debit-positive) GL sum into a P&L display amount.
DISPLAY_SIGN = {
    'Income':           -1,
    'CostOfGoodsSold':   1,
    'Expense':           1,
    'OtherIncome':      -1,
    'OtherExpense':      1,
}
PNL_TYPES = set(DISPLAY_SIGN)


def _dec(s):
    """Parse a GL amount string to Decimal; '' / None / junk -> 0."""
    if s is None:
        return Decimal('0')
    if isinstance(s, (int, float)):
        # Tolerate a numeric amount, though GL amounts are exported as strings.
        return Decimal(str(s))
    s = s.strip()
    if not s:
        return Decimal('0')
    try:
        return Decimal(s)
    except InvalidOperation:
        return Decimal('0')


def _fmt(d):
    """QB-style money: 2 decimals, thousands separators, negatives in parens."""
    q = d.quantize(Decimal('0.01'))
    body = f"{abs(q):,.2f}"
    return f"({body})" if q < 0 else body


def _index_accounts(accounts):
    """full_name/name -> account dict, for classifying GL rows by account type."""
    by_full, by_name = {}, {}
    for a in accounts or []:
        if a.get('full_name'):
            by_full.setdefault(a['full_name'], a)
        if a.get('name'):
            by_name.setdefault(a['name'], a)
    return by_full, by_name


def _classify(row, by_full, by_name):
    """Return (account_full_name, account_type) for a GL row, preferring the
    accounts master over the row's own (possibly un-enriched) fields."""
    full = row.get('account_full_name', '')
    acct = by_full.get(full) or by_name.get(full)
    if acct:
        return acct.get('full_name') or full, acct.get('account_type', '')
    # Fall back to whatever the extractor enriched onto the row.
    return full, row.get('account_type', '')


def build_pnl(bundle, basis):
    """Aggregate the bundle's general_ledger for one basis into per-account and
    per-type totals. Returns a dict with account sums, type sums, and the
    all-accounts control total (should be ~0 for a complete double-entry set)."""
    by_full, by_name = _index_accounts(bundle.get('accounts'))
    gl = [r for r in bundle.get('general_ledger', []) if r.get('basis') == basis]

    per_account = {}      # full_name -> {'type':.., 'raw': Decimal}
    control_total = Decimal('0')   # sum of ALL raw amounts (all account types)
    unclassified = {'count': 0, 'raw': Decimal('0')}

    for r in gl:
        amt = _dec(r.get('amount'))
        control_total += amt
        full, atype = _classify(r, by_full, by_name)
        if atype not in PNL_TYPES:
            if not atype:
                unclassified['count'] += 1
                unclassified['raw'] += amt
            continue
        slot = per_account.setdefault(full, {'type': atype, 'raw': Decimal('0')})
        slot['raw'] += amt

    return {
        'basis': basis,
        'lines': len(gl),
        'per_account': per_account,
        'control_total': control_total,
        'unclassified': unclassified,
    }


def _section(per_account, atype):
    """(sorted [(full_name, display_amount)], total_display) for one account type."""
    sign = DISPLAY_SIGN[atype]
    rows = [(full, sign * v['raw'])
            for full, v in per_account.items() if v['type'] == atype]
    rows.sort(key=lambda t: t[0].lower())
    total = sum((amt for _, amt in rows), Decimal('0'))
    return rows, total


def print_pnl(pnl, totals_only=False, out=sys.stdout):
    """Print a QB-style P&L for one basis."""
    pa = pnl['per_account']
    W = 62

    def line(label, amount=None, indent=0):
        if amount is None:
            print(f"{'  ' * indent}{label}", file=out)
        else:
            text = f"{'  ' * indent}{label}"
            print(f"{text:<{W}}{_fmt(amount):>18}", file=out)

    def section(title, atype):
        rows, total = _section(pa, atype)
        if not rows and total == 0:
            return total
        print(f"\n  {title}", file=out)
        if not totals_only:
            for full, amt in rows:
                depth = full.count(':')
                short = full.split(':')[-1]
                line(short, amt, indent=2 + depth)
        line(f"Total {title}", total, indent=1)
        return total

    print("=" * (W + 18), file=out)
    print(f"  PROFIT & LOSS — {pnl['basis']} basis", file=out)
    print(f"  ({pnl['lines']} GL posting lines)", file=out)
    print("=" * (W + 18), file=out)

    income = section('Income', 'Income')
    cogs = section('Cost of Goods Sold', 'CostOfGoodsSold')
    gross = income - cogs
    print(file=out)
    line('Gross Profit', gross, indent=0)

    expense = section('Expense', 'Expense')
    net_ordinary = gross - expense
    print(file=out)
    line('Net Ordinary Income', net_ordinary, indent=0)

    other_inc = section('Other Income', 'OtherIncome')
    other_exp = section('Other Expense', 'OtherExpense')
    net_other = other_inc - other_exp
    if other_inc or other_exp:
        print(file=out)
        line('Net Other Income', net_other, indent=0)

    net_income = net_ordinary + net_other
    print("\n" + " " * 2 + "-" * (W + 16 - 2), file=out)
    line('NET INCOME', net_income, indent=0)
    print("=" * (W + 18), file=out)

    # Reconciliation diagnostics.
    print("\n  Checks:", file=out)
    ctrl = pnl['control_total'].quantize(Decimal('0.01'))
    ok = ctrl == 0
    print(f"    Ledger balances (sum of ALL posting amounts): {_fmt(ctrl)} "
          f"{'OK' if ok else '<-- NOT zero: missing rows or a failed chunk'}",
          file=out)
    unc = pnl['unclassified']
    if unc['count']:
        print(f"    Unclassified GL lines (account type unknown): {unc['count']} "
              f"totaling {_fmt(unc['raw'])} — excluded from P&L, investigate", file=out)
    # Net income should equal the negative of all P&L raw sums, which equals the
    # balance-sheet net movement (since the whole ledger nets to zero).
    print(f"    Compare NET INCOME above to QuickBooks' P&L "
          f"({pnl['basis']} basis) for the same period.", file=out)


def write_csv(pnls, path):
    """Write per-account P&L display amounts (one column per basis) to CSV."""
    import csv
    bases = [p['basis'] for p in pnls]
    # union of accounts across bases, in report-type order then name
    type_order = ['Income', 'CostOfGoodsSold', 'Expense', 'OtherIncome', 'OtherExpense']
    keys = {}
    for p in pnls:
        for full, v in p['per_account'].items():
            keys[full] = v['type']
    def sort_key(full):
        return (type_order.index(keys[full]) if keys[full] in type_order else 99,
                full.lower())
    with open(path, 'w', newline='') as f:
        w = csv.writer(f)
        w.writerow(['account_full_name', 'account_type'] + [f'{b}' for b in bases])
        for full in sorted(keys, key=sort_key):
            atype = keys[full]
            sign = DISPLAY_SIGN.get(atype, 1)
            amts = []
            for p in pnls:
                v = p['per_account'].get(full)
                amts.append(str((sign * v['raw']).quantize(Decimal('0.01')))
                            if v else '')
            w.writerow([full, atype] + amts)
    print(f"\nWrote per-account CSV: {path}")


def main():
    ap = argparse.ArgumentParser(
        description="Reconcile a QBExtract bundle's General Ledger into a "
                    "QuickBooks-style P&L for comparison.")
    ap.add_argument('json_file', help='Path to the QBExtract export JSON bundle.')
    ap.add_argument('--basis', choices=['accrual', 'cash', 'both'], default='accrual',
                    help='Which basis to report (default: accrual).')
    ap.add_argument('--totals-only', action='store_true',
                    help='Print section totals only, not individual accounts.')
    ap.add_argument('--csv', default=None,
                    help='Also write a per-account CSV (accounts x basis).')
    args = ap.parse_args()

    try:
        with open(args.json_file, encoding='utf-8') as f:
            bundle = json.load(f)
    except (OSError, json.JSONDecodeError) as e:
        print(f"ERROR: could not read {args.json_file}: {e}", file=sys.stderr)
        sys.exit(1)

    gl = bundle.get('general_ledger')
    if not gl:
        print("This bundle has no `general_ledger` array. Re-run QBExtract "
              "without --no-gl (GL is on by default).", file=sys.stderr)
        sys.exit(1)

    meta = bundle.get('meta', {})
    print(f"Company: {meta.get('company', '?')}")
    period = meta.get('date_range') or f"years_back={meta.get('years_back', '?')}"
    print(f"Period:  {period}")
    available = sorted({r.get('basis') for r in gl if r.get('basis')})
    print(f"Bases in bundle: {', '.join(available) or '(none)'}")

    want = {'accrual': ['Accrual'], 'cash': ['Cash'],
            'both': ['Accrual', 'Cash']}[args.basis]
    pnls = []
    for basis in want:
        if basis not in available:
            print(f"\n(skipping {basis} — not present in bundle)")
            continue
        pnl = build_pnl(bundle, basis)
        print()
        print_pnl(pnl, totals_only=args.totals_only)
        pnls.append(pnl)

    if args.csv and pnls:
        write_csv(pnls, args.csv)


if __name__ == '__main__':
    main()
