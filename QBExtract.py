"""
QBExtract.py — QuickBooks Desktop Data Extractor
=================================================
Reads data directly from a running QuickBooks Desktop/Enterprise company file
via the QB SDK (QBXML) and produces a single JSON bundle for import.

What's new in v3 (2026-07):
  - **General Ledger** — the posting-line ledger, pulled via QuickBooks'
    own report engine (`GeneralDetailReportQueryRq` / `GeneralLedger`) rather
    than by re-deriving double-entry from raw transactions. QBD generates the
    implicit balancing entries and, with `<ReportBasis>Cash</ReportBasis>`,
    computes cash-basis itself. Runs Accrual basis by default (add `--gl-basis
    cash` or `both` for the QBD-computed cash basis) and tags every row with its
    basis. See `--gl-basis` / `--no-gl`. GL amounts
    are stored as exact decimal STRINGS (never float) so they reconcile to the
    penny downstream.
  - **`--probe-gl`** — a structure probe for the GL report. Run it against the
    real company file first: it dumps the raw report XML and reports which
    columns come back and whether each line's internal TxnID is exposed (the
    one version-dependent unknown that drill-down/attachment linking depends on).

What's new in v2 (2026-05):
  - **ASCII normalization** of every text field. Real-world QB files have
    Windows-1252 smart punctuation, Latin-1 supplement chars, and stray
    high-bit bytes that the SDK silently rejects on round-trip. The extractor
    now decodes XML entities, remaps CP1252 control-range chars to ASCII, and
    strips anything outside 7-bit ASCII. Necessary for migrating corrupt files.
  - **Full Chart of Accounts** — Account hierarchy with ParentRef preserved.
    Previous version omitted accounts entirely.
  - **All item types** — Service, OtherCharge, and Discount items in addition
    to Inventory + Non-Inventory. Transactions reference these (Fuel Surcharge,
    Shipping, Pick-Up Credit, etc.) and silently fail on import if missing.
  - **Customer jobs** — sub-customers now retained with `parent` field
    pointing at the parent's FullName. v1 dropped them entirely.
  - **Ship Methods + Payment Methods + Sales Tax Codes + Sales Tax Items** —
    new master tables. Transactions reference them; without them, import
    rejects with "invalid reference" errors.
  - **Sales tax fields on invoices + credit memos** — subtotal, sales tax
    total, tax item (jurisdiction), tax code, per-line tax codes. The push
    side needs these to reconstruct tax accurately.
  - **Bill line typing** — expense lines vs item lines are now distinguished
    (`line_type` field). They're separate XML schemas on push.
  - **Date range filters** — `--from-date YYYY-MM-DD`, `--to-date YYYY-MM-DD`,
    `--year YYYY` for surgical re-extracts when patching a specific period.
  - **Name-range chunking** for Customers and Items. A corrupt record in one
    alphabetic bucket only loses that bucket instead of failing the whole
    query. Works on clean files too; no flag needed.

Requirements:
  - QuickBooks Desktop/Enterprise must be OPEN with the company file loaded
  - QuickBooks SDK must be installed (comes with QB or download from Intuit)
  - No Python or pywin32 needed when compiled to EXE (uses VBScript for COM)

Usage:
  python QBExtract.py                    # interactive, prompts for years
  python QBExtract.py --years 3          # last 3 years
  python QBExtract.py --years 0          # all history
  python QBExtract.py --year 2024        # only 2024
  python QBExtract.py --from-date 2024-01-01 --to-date 2024-06-30
  (or compiled to QBExtract.exe via PyInstaller)

Output:
  <CompanyName>_export_<YYYYMMDD>.json  in the current directory

Compile to EXE:
  pip install pyinstaller
  pyinstaller --onefile --console QBExtract.py
"""

import argparse
import datetime
import html
import json
import os
import re
import subprocess
import sys
import tempfile
import traceback
from decimal import Decimal, InvalidOperation


# ============================================================================
# QB SESSION — uses VBScript for COM (100% reliable on any Windows)
# ============================================================================

# VBScript template that handles all QB SDK COM calls.
# Python writes the QBXML request to a file, calls this script,
# and reads the response from another file.
VBS_TEMPLATE = r'''
' QBExtract VBScript bridge — called by Python
' Args: action, request_file, response_file, [company_file]
Dim action, reqFile, resFile
action  = WScript.Arguments(0)
reqFile = WScript.Arguments(1)
resFile = WScript.Arguments(2)

' Optional 4th arg: full path to a .QBW company file to open directly. Empty
' string (the default) means "use the company file QuickBooks currently has
' open." A specific path only works when QB is closed (and this app has
' unattended access) or already open with that same file.
Dim qbFile
If WScript.Arguments.Count > 3 Then
    qbFile = WScript.Arguments(3)
Else
    qbFile = ""
End If

Dim rp
On Error Resume Next
Set rp = CreateObject("QBXMLRP2.RequestProcessor")
If Err.Number <> 0 Then
    Err.Clear
    Set rp = CreateObject("QBXMLRP.RequestProcessor")
    If Err.Number <> 0 Then
        WriteFile resFile, "ERROR: Could not create QB RequestProcessor: " & Err.Description
        WScript.Quit 1
    End If
End If
On Error GoTo 0

If action = "connect" Then
    On Error Resume Next
    rp.OpenConnection2 "", "QBExtract - ERP Importer", 1
    If Err.Number <> 0 Then
        Err.Clear
        rp.OpenConnection "", "QBExtract - ERP Importer"
        If Err.Number <> 0 Then
            WriteFile resFile, "ERROR: OpenConnection failed: " & Err.Description
            WScript.Quit 1
        End If
    End If
    On Error GoTo 0

    Dim ticket
    On Error Resume Next
    ticket = rp.BeginSession(qbFile, 2)
    If Err.Number <> 0 Then
        Err.Clear
        ticket = rp.BeginSession(qbFile, 0)
        If Err.Number <> 0 Then
            WriteFile resFile, "ERROR: BeginSession failed: " & Err.Description
            rp.CloseConnection
            WScript.Quit 1
        End If
    End If
    On Error GoTo 0

    WriteFile resFile, "OK:" & ticket
    rp.EndSession
    rp.CloseConnection

ElseIf action = "query" Then
    Dim reqXml, envelope, response

    reqXml = ReadFile(reqFile)

    On Error Resume Next
    rp.OpenConnection2 "", "QBExtract - ERP Importer", 1
    If Err.Number <> 0 Then
        Err.Clear
        rp.OpenConnection "", "QBExtract - ERP Importer"
    End If
    On Error GoTo 0

    On Error Resume Next
    ticket = rp.BeginSession(qbFile, 2)
    If Err.Number <> 0 Then
        Err.Clear
        ticket = rp.BeginSession(qbFile, 0)
    End If
    On Error GoTo 0

    envelope = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
               "<?qbxml version=""13.0""?>" & _
               "<QBXML><QBXMLMsgsRq onError=""continueOnError"">" & _
               reqXml & _
               "</QBXMLMsgsRq></QBXML>"

    On Error Resume Next
    response = rp.ProcessRequest(ticket, envelope)
    If Err.Number <> 0 Then
        WriteFile resFile, "ERROR: ProcessRequest failed: " & Err.Description
        rp.EndSession
        rp.CloseConnection
        WScript.Quit 1
    End If
    On Error GoTo 0

    WriteFile resFile, response
    rp.EndSession
    rp.CloseConnection
End If

Sub WriteFile(path, content)
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.CreateTextFile(path, True, True)
    f.Write content
    f.Close
End Sub

Function ReadFile(path)
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(path, 1, False, -1)
    ReadFile = f.ReadAll
    f.Close
End Function
'''


def _connect_error_hint(msg, company_file=''):
    """Append actionable guidance to known connect-stage failures. The
    auto-login/permission case is the common footgun: it means QB was asked to
    open the file itself (QB closed, or --company-file given) without unattended
    access granted, and a failed attempt can leave QuickBooks running headless
    holding the file lock."""
    low = msg.lower()
    if any(s in low for s in ('log in', 'log into', 'automatic', 'permission',
                              'administrator')):
        return (
            msg + "\n\n"
            "  This app tried to open the company file AUTOMATICALLY "
            + ("(--company-file was given)" if company_file
               else "(QuickBooks was not open with the file)") + ",\n"
            "  but it is not granted auto-login.\n\n"
            "  EASIEST FIX: open QuickBooks, log in, and load the company file,\n"
            "  then run this WITHOUT --company-file. An attached live session\n"
            "  needs no auto-login permission.\n\n"
            "  For unattended runs: QuickBooks > Edit > Preferences > Integrated\n"
            "  Applications > Company Preferences > select 'QBExtract - ERP\n"
            "  Importer' > check 'Allow this application to log in automatically'\n"
            "  and choose a login user. If already granted, revoke it, Save, then\n"
            "  grant again (a known QB quirk).\n\n"
            "  NOTE: a failed auto-login can leave QuickBooks running in the\n"
            "  background holding the file open. If QB then says the file is\n"
            "  'already open', end the QBW32.EXE / QuickBooks.exe process in Task\n"
            "  Manager (or reboot) before retrying."
        )
    return msg


class QBSession:
    """Manages connection to QuickBooks via VBScript COM bridge."""

    def __init__(self):
        self.company_name = ''
        self.company_file = ''
        self._vbs_path = None
        self._req_path = None
        self._res_path = None

    def connect(self, company_file=''):
        """Connect to QuickBooks. `company_file` is an optional full path to a
        .QBW file to open directly; empty (the default) uses the file QB
        currently has open. A path only works when QB is closed and this app
        has unattended access, or QB is already open with that same file — the
        SDK cannot switch QB from a different open file."""
        self.company_file = company_file or ''
        # Write VBScript bridge to temp file
        self._vbs_path = os.path.join(tempfile.gettempdir(), 'qbextract_bridge.vbs')
        self._req_path = os.path.join(tempfile.gettempdir(), 'qbextract_req.xml')
        self._res_path = os.path.join(tempfile.gettempdir(), 'qbextract_res.xml')

        with open(self._vbs_path, 'w', encoding='utf-8') as f:
            f.write(VBS_TEMPLATE)

        # Test connection. Note: connect is deliberately NOT retried — an
        # auto-login/permission failure must fail fast, not relaunch QB
        # repeatedly.
        try:
            result = self._run_vbs('connect')
        except RuntimeError as e:
            raise RuntimeError(_connect_error_hint(str(e), self.company_file)) from None
        if result.startswith('ERROR:'):
            raise RuntimeError(_connect_error_hint(result, self.company_file))

        # Get company name
        xml = self._send(self._company_query())
        self.company_name = self._parse_company_name(xml)
        print(f"  Connected to: {self.company_name}")

    def disconnect(self):
        # Clean up temp files
        for p in (self._vbs_path, self._req_path, self._res_path):
            if p and os.path.exists(p):
                try:
                    os.remove(p)
                except Exception:
                    pass

    def _run_vbs(self, action):
        """Run the VBScript bridge and return the response."""
        if not os.path.exists(self._req_path):
            with open(self._req_path, 'w') as f:
                f.write('')

        with open(self._res_path, 'w') as f:
            f.write('')

        try:
            proc = subprocess.run(
                ['cscript', '//Nologo', self._vbs_path,
                 action, self._req_path, self._res_path,
                 self.company_file or ''],
                capture_output=True, text=True, timeout=600
            )
        except FileNotFoundError:
            raise RuntimeError("cscript.exe not found — Windows Script Host may be disabled")

        if not os.path.exists(self._res_path):
            stderr = proc.stderr.strip() if proc.stderr else 'No output'
            raise RuntimeError(f"VBScript bridge failed: {stderr}")

        # VBScript FileSystemObject may write UTF-16LE; try that first, fall back to utf-8-sig
        try:
            with open(self._res_path, 'r', encoding='utf-16') as f:
                result = f.read()
        except (UnicodeError, UnicodeDecodeError):
            with open(self._res_path, 'r', encoding='utf-8-sig') as f:
                result = f.read()

        if result.startswith('ERROR:'):
            raise RuntimeError(result)

        return result

    def _send(self, request_xml):
        """Send QBXML request via VBScript bridge, return response XML.

        No automatic retry: each call is a full connect -> BeginSession ->
        ProcessRequest -> EndSession -> CloseConnection cycle, and re-firing it
        immediately after a failure hammers QuickBooks with a fresh connection
        while it is still unstable — which can crash QB. A failed request is
        logged by the caller and its chunk skipped; re-run that period instead."""
        with open(self._req_path, 'w', encoding='utf-16') as f:
            f.write(request_xml)

        result = self._run_vbs('query')
        return result

    def _company_query(self):
        return '<CompanyQueryRq requestID="1"></CompanyQueryRq>'

    def _parse_company_name(self, xml):
        """Extract company name from CompanyQueryRs."""
        m = re.search(r'<CompanyName>(.*?)</CompanyName>', xml)
        return m.group(1) if m else 'UnknownCompany'


# ============================================================================
# XML HELPERS
# ============================================================================

def xml_val(xml, tag, default=''):
    """Extract first value of a tag from XML string."""
    m = re.search(rf'<{tag}>(.*?)</{tag}>', xml, re.DOTALL)
    return m.group(1).strip() if m else default


def xml_all(xml, tag):
    """Extract all occurrences of a tag."""
    return re.findall(rf'<{tag}>(.*?)</{tag}>', xml, re.DOTALL)


def xml_blocks(xml, tag):
    """Extract all blocks between opening and closing tags."""
    return re.findall(rf'<{tag}[\s>].*?</{tag}>', xml, re.DOTALL)


def xml_ref(xml, ref_tag, sub_tag='FullName'):
    """Extract a sub-tag from a Ref block, e.g. CustomerRef -> FullName."""
    m = re.search(rf'<{ref_tag}>(.*?)</{ref_tag}>', xml, re.DOTALL)
    if not m:
        return ''
    inner = m.group(1)
    m2 = re.search(rf'<{sub_tag}>(.*?)</{sub_tag}>', inner, re.DOTALL)
    return m2.group(1).strip() if m2 else ''


# ============================================================================
# ASCII NORMALIZATION
# ============================================================================
#
# Real-world QB files accumulate text data from many sources over decades:
# Windows clipboards, Word/Excel imports, OCR, scanner inputs, etc. By the
# time it lands in QB, the strings may contain:
#   * Windows-1252 control-range chars (0x80-0x9F): smart quotes, em dash, etc.
#   * Already-decoded Unicode general punctuation (U+2013, U+2018, etc.)
#   * Latin-1 supplement (0xA0-0xFF): non-breaking space, degree sign, accents
#
# QB SDK qbXML v13 silently rejects entire batches containing any of these.
# The error message is generic ("Invalid argument") and unrelated to the
# offending field. We have to normalize every string to 7-bit ASCII before
# either writing JSON or pushing back.

_ASCII_REMAP = {
    # Windows-1252 control range (`&#150;` etc. decode here first)
    '\x80': 'EUR', '\x82': ',',  '\x83': 'f',  '\x84': '"',
    '\x85': '...','\x86': '+',   '\x87': '++', '\x88': '^',
    '\x89': '%o', '\x8A': 'S',   '\x8B': '<',  '\x8C': 'OE',
    '\x8E': 'Z',  '\x91': "'",   '\x92': "'",  '\x93': '"',
    '\x94': '"',  '\x95': '*',   '\x96': '-',  '\x97': '--',
    '\x98': '~',  '\x99': '(TM)','\x9A': 's',  '\x9B': '>',
    '\x9C': 'oe', '\x9E': 'z',   '\x9F': 'Y',
    # Common Unicode general-punctuation already-decoded equivalents
    '–': '-', '—': '--', '‘': "'", '’': "'",
    '‚': ',', '“': '"',  '”': '"', '„': '"',
    '•': '*', '…': '...','‰': '%o','‹': '<',
    '›': '>', '™': '(TM)',
    # Latin-1 supplement chars QB also rejects
    ' ': ' ', '·': '.', '°': 'deg', '´': "'",
}


def clean(s):
    """Strip whitespace, decode XML entities, ASCII-normalize. Three layered
    sanitizations: html.unescape converts `&apos;`/`&#150;` to real chars
    (otherwise re-escape on push doubles `&apos;` -> `&amp;apos;` and overflows
    QB's 41-char Name); _ASCII_REMAP swaps known smart-punctuation for ASCII
    equivalents; final strip drops any remaining char outside 7-bit ASCII (QB
    SDK 13 rejects whole batches containing them, even valid UTF-8)."""
    if not s:
        return ''
    out = html.unescape(s.strip())
    for k, v in _ASCII_REMAP.items():
        if k in out:
            out = out.replace(k, v)
    return ''.join(ch for ch in out if 0x20 <= ord(ch) < 0x7F or ch in '\t\n\r')


def to_float(s):
    try:
        return float(s.strip()) if s and s.strip() else 0.0
    except (ValueError, AttributeError):
        return 0.0


# ============================================================================
# CHUNKING HELPERS
# ============================================================================

def _name_range_chunks():
    """Alphabetic name-range chunks for Customer/Item queries. Each chunk uses
    inclusive FromName + an upper bound that's the next prefix + 'zzzzz', so
    names like 'Acme', 'Adams' all fall into the 'A' bucket.

    Why: a single CustomerQuery on a corrupt file may abort partway through
    because of one bad record. Chunking by name range isolates damage to a
    single alphabetic bucket — you lose 1/36th of the master at worst, instead
    of everything after the corrupt record."""
    ranges = []
    for c in '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ':
        ranges.append((c, c + 'zzzzz'))
    return ranges


def _build_date_chunks(years_back, date_range):
    """Return list of (from_date, to_date) pairs to iterate. If date_range is
    supplied, returns exactly one window covering that range. Otherwise builds
    backward year chunks."""
    if date_range is not None:
        return [date_range]
    today = datetime.date.today()
    max_years = 30 if years_back == 0 else years_back
    chunks = []
    for i in range(max_years):
        chunk_end   = today - datetime.timedelta(days=365 * i)
        chunk_start = today - datetime.timedelta(days=365 * (i + 1))
        chunks.append((chunk_start.strftime('%Y-%m-%d'),
                       chunk_end.strftime('%Y-%m-%d')))
    return chunks


def _txn_date_filter(from_d, to_d):
    """Build a TxnDateRangeFilter block. Either endpoint may be None/empty."""
    parts = ['      <TxnDateRangeFilter>']
    if from_d:
        parts.append(f'        <FromTxnDate>{from_d}</FromTxnDate>')
    if to_d:
        parts.append(f'        <ToTxnDate>{to_d}</ToTxnDate>')
    parts.append('      </TxnDateRangeFilter>')
    return '\n'.join(parts)


# ============================================================================
# EXTRACTORS — MASTERS
# ============================================================================

def extract_accounts(session):
    """Pull the chart of accounts. Single query — typical company has <500
    accounts. Captures the hierarchy via `parent` so the import can rebuild
    nested accounts in the right order."""
    print("  Extracting chart of accounts...")
    request = '<AccountQueryRq requestID="100"><ActiveStatus>All</ActiveStatus></AccountQueryRq>'
    xml = session._send(request)
    out = []
    for b in xml_blocks(xml, 'AccountRet'):
        out.append({
            'list_id':        clean(xml_val(b, 'ListID')),
            'name':           clean(xml_val(b, 'Name')),
            'full_name':      clean(xml_val(b, 'FullName')),
            'parent':         clean(xml_ref(b, 'ParentRef', 'FullName')),
            'account_type':   clean(xml_val(b, 'AccountType')),
            'special_type':   clean(xml_val(b, 'SpecialAccountType')),
            'account_number': clean(xml_val(b, 'AccountNumber')),
            'description':    clean(xml_val(b, 'Desc')),
            'active':         xml_val(b, 'IsActive') == 'true',
        })
    print(f"    {len(out)} accounts found")
    return out


def _parse_customer_blocks(blocks):
    """Parse <CustomerRet> blocks into customer dicts. Sub-customers (jobs)
    are included; their `parent` field carries the parent's FullName so the
    push can recreate the hierarchy with ParentRef."""
    customers = []
    for b in blocks:
        bill_blocks = xml_blocks(b, 'BillAddress')
        bill = bill_blocks[0] if bill_blocks else ''
        customers.append({
            'list_id':      clean(xml_val(b, 'ListID')),
            'name':         clean(xml_val(b, 'Name')),
            'parent':       clean(xml_ref(b, 'ParentRef', 'FullName')),
            'company':      clean(xml_val(b, 'CompanyName')),
            'contact':      clean(xml_val(b, 'FirstName') + ' ' + xml_val(b, 'LastName')).strip(),
            'phone':        clean(xml_val(b, 'Phone')),
            'fax':          clean(xml_val(b, 'Fax')),
            'email':        clean(xml_val(b, 'Email')),
            'addr1':        clean(xml_val(bill, 'Addr1')),
            'addr2':        clean(xml_val(bill, 'Addr2')),
            'city':         clean(xml_val(bill, 'City')),
            'state':        clean(xml_val(bill, 'State')),
            'zip':          clean(xml_val(bill, 'PostalCode')),
            'country':      clean(xml_val(bill, 'Country')),
            'terms':        clean(xml_ref(b, 'TermsRef', 'FullName')),
            'salesman':     clean(xml_ref(b, 'SalesRepRef', 'FullName')),
            'price_level':  clean(xml_ref(b, 'PriceLevelRef', 'FullName')),
            'tax_code':     clean(xml_ref(b, 'CustomerTaxCodeRef', 'FullName')),
            'credit_limit': to_float(xml_val(b, 'CreditLimit')),
            'balance':      to_float(xml_val(b, 'Balance')),
            'active':       xml_val(b, 'IsActive') == 'true',
            'notes':        clean(xml_val(b, 'Notes')),
            'account_no':   clean(xml_val(b, 'AccountNumber')),
        })
    return customers


def extract_customers(session, corrupt_safe=False):
    """Pull all customers via name-range chunks (parents AND jobs).

    `corrupt_safe`: when True, narrows the response to non-history fields only.
    Excluding Balance keeps the SDK from walking each customer's transaction
    ledger to compute the balance — which is what trips on corrupt records.
    The trade-off is `balance` comes back as 0 in the JSON; recompute it
    later from invoices - payments if needed."""
    print("  Extracting customers (chunked by name range)...")
    if corrupt_safe:
        include_elements = ''.join(
            f"      <IncludeRetElement>{e}</IncludeRetElement>\n" for e in [
                'ListID', 'Name', 'FullName', 'IsActive', 'ParentRef',
                'CompanyName', 'FirstName', 'LastName', 'Phone', 'Fax', 'Email',
                'BillAddress', 'ShipAddress', 'TermsRef', 'SalesRepRef',
                'PriceLevelRef', 'CustomerTaxCodeRef', 'CreditLimit',
                'Notes', 'AccountNumber', 'JobStatus',
            ]
        )
    else:
        include_elements = ''

    all_custs = []
    for from_name, to_name in _name_range_chunks():
        # qbxml v13 schema order: ActiveStatus, NameRangeFilter, IncludeRetElement, OwnerID
        request = f"""
    <CustomerQueryRq requestID="2">
      <ActiveStatus>All</ActiveStatus>
      <NameRangeFilter>
        <FromName>{from_name}</FromName>
        <ToName>{to_name}</ToName>
      </NameRangeFilter>
{include_elements}      <OwnerID>0</OwnerID>
    </CustomerQueryRq>
    """
        try:
            xml = session._send(request)
            chunk = _parse_customer_blocks(xml_blocks(xml, 'CustomerRet'))
            if chunk:
                print(f"    {from_name}*: {len(chunk)} customers")
            all_custs.extend(chunk)
        except Exception as e:
            print(f"    {from_name}*: FAILED -- {str(e)[:100]}")

    print(f"    {len(all_custs)} customers found total")
    return all_custs


def _parse_item_blocks(blocks, item_type):
    """Parse Item*Ret blocks. Same shape regardless of inventory/service/etc."""
    items = []
    for b in blocks:
        items.append({
            'list_id':      clean(xml_val(b, 'ListID')),
            'item':         clean(xml_val(b, 'Name')),
            'description':  clean(xml_val(b, 'SalesDesc') or xml_val(b, 'Desc')),
            'item_type':    item_type,
            'active':       xml_val(b, 'IsActive') == 'true',
            'price':        to_float(xml_val(b, 'SalesPrice')),
            'cost':         to_float(xml_val(b, 'PurchaseCost')),
            'on_hand':      to_float(xml_val(b, 'QuantityOnHand')),
            'on_order':     to_float(xml_val(b, 'QuantityOnOrder')),
            'reorder_pt':   to_float(xml_val(b, 'ReorderPoint')),
            'unitms':       clean(xml_ref(b, 'UnitOfMeasureSetRef', 'FullName')),
            'income_acct':  clean(xml_ref(b, 'IncomeAccountRef', 'FullName')),
            'asset_acct':   clean(xml_ref(b, 'AssetAccountRef', 'FullName')),
            'cogs_acct':    clean(xml_ref(b, 'COGSAccountRef', 'FullName')),
            'item_group':   clean(xml_ref(b, 'ParentRef', 'FullName')),
            'barcode':      clean(xml_val(b, 'BarCodeValue')),
        })
    return items


def extract_items(session, corrupt_safe=False):
    """Pull every item type the SDK exposes: Inventory, Non-Inventory, Service,
    OtherCharge, Discount. All five share the same JSON shape via `item_type`.

    Special-type items (Service/OtherCharge/Discount) are pulled in a single
    query each (typical count <50). Inventory + Non-Inventory use name-range
    chunking for corrupt-file safety.

    `corrupt_safe`: when True, omits QuantityOnHand/QuantityOnOrder on
    Inventory queries — those fields force the SDK to walk the inventory
    ledger, which is where corrupt transaction records live."""
    print("  Extracting items (chunked by name range)...")

    if corrupt_safe:
        inv_includes = ''.join(
            f"      <IncludeRetElement>{e}</IncludeRetElement>\n" for e in [
                'ListID', 'Name', 'FullName', 'IsActive', 'ParentRef',
                'SalesDesc', 'SalesPrice', 'PurchaseCost', 'ReorderPoint',
                'UnitOfMeasureSetRef', 'IncomeAccountRef', 'AssetAccountRef',
                'COGSAccountRef', 'BarCodeValue',
            ]
        )
        non_includes = ''.join(
            f"      <IncludeRetElement>{e}</IncludeRetElement>\n" for e in [
                'ListID', 'Name', 'FullName', 'IsActive', 'ParentRef',
                'SalesDesc', 'SalesPrice', 'PurchaseCost',
                'UnitOfMeasureSetRef', 'IncomeAccountRef', 'BarCodeValue',
            ]
        )
    else:
        inv_includes = non_includes = ''

    all_items = []
    for item_type, query_tag, ret_tag, includes in [
        ('INV', 'ItemInventoryQueryRq',    'ItemInventoryRet',    inv_includes),
        ('NON', 'ItemNonInventoryQueryRq', 'ItemNonInventoryRet', non_includes),
        ('SVC', 'ItemServiceQueryRq',      'ItemServiceRet',      ''),
        ('OCH', 'ItemOtherChargeQueryRq',  'ItemOtherChargeRet',  ''),
        ('DSC', 'ItemDiscountQueryRq',     'ItemDiscountRet',     ''),
    ]:
        # Special-type items: single query, no name-range chunking
        if item_type in ('SVC', 'OCH', 'DSC'):
            request = f'<{query_tag} requestID="3"><ActiveStatus>All</ActiveStatus></{query_tag}>'
            try:
                xml = session._send(request)
                chunk = _parse_item_blocks(xml_blocks(xml, ret_tag), item_type)
                if chunk:
                    print(f"    {item_type}: {len(chunk)} items")
                all_items.extend(chunk)
            except Exception as e:
                print(f"    {item_type}: FAILED -- {str(e)[:100]}")
            continue

        # Inventory + Non-Inventory: name-range chunked
        for from_name, to_name in _name_range_chunks():
            request = f"""
    <{query_tag} requestID="3">
      <ActiveStatus>All</ActiveStatus>
      <NameRangeFilter>
        <FromName>{from_name}</FromName>
        <ToName>{to_name}</ToName>
      </NameRangeFilter>
{includes}      <OwnerID>0</OwnerID>
    </{query_tag}>
    """
            try:
                xml = session._send(request)
                chunk = _parse_item_blocks(xml_blocks(xml, ret_tag), item_type)
                if chunk:
                    print(f"    {item_type} {from_name}*: {len(chunk)} items")
                all_items.extend(chunk)
            except Exception as e:
                print(f"    {item_type} {from_name}*: FAILED -- {str(e)[:100]}")

    print(f"    {len(all_items)} items found total")
    return all_items


def extract_price_levels(session):
    """Pull all price levels (Enterprise feature)."""
    print("  Extracting price levels...")
    request = '<PriceLevelQueryRq requestID="5"><ActiveStatus>All</ActiveStatus></PriceLevelQueryRq>'
    xml = session._send(request)
    price_levels = []
    for b in xml_blocks(xml, 'PriceLevelRet'):
        fixed_prices = []
        for ib in xml_blocks(b, 'PriceLevelPerItemRet'):
            fixed_prices.append({
                'item_list_id':  clean(xml_ref(ib, 'ItemRef', 'ListID')),
                'item_name':     clean(xml_ref(ib, 'ItemRef', 'FullName')),
                'custom_price':  to_float(xml_val(ib, 'CustomPrice')),
                'custom_pct':    to_float(xml_val(ib, 'CustomPricePercent')),
            })
        price_levels.append({
            'name':         clean(xml_val(b, 'Name')),
            'type':         clean(xml_val(b, 'PriceLevelType')),
            'pct':          to_float(xml_val(b, 'PriceLevelFixedPct')),
            'fixed_prices': fixed_prices,
        })
    print(f"    {len(price_levels)} price levels found")
    return price_levels


def extract_quantity_discounts(session):
    """Pull discount items. Note: same query as Discount items in extract_items;
    kept here for backward compatibility with v1 JSON consumers."""
    print("  Extracting quantity discounts...")
    request = '<ItemDiscountQueryRq requestID="11"><ActiveStatus>All</ActiveStatus></ItemDiscountQueryRq>'
    xml = session._send(request)
    discounts = []
    for b in xml_blocks(xml, 'ItemDiscountRet'):
        discounts.append({
            'list_id':       clean(xml_val(b, 'ListID')),
            'item_name':     clean(xml_val(b, 'Name')),
            'description':   clean(xml_val(b, 'ItemDesc')),
            'discount_rate': to_float(xml_val(b, 'DiscountRate')),
            'discount_pct':  to_float(xml_val(b, 'DiscountRatePercent')),
            'account':       clean(xml_ref(b, 'AccountRef', 'FullName')),
            'active':        xml_val(b, 'IsActive') == 'true',
        })
    print(f"    {len(discounts)} discount items found")
    return discounts


def extract_vendors(session):
    """Pull all vendors."""
    print("  Extracting vendors...")
    request = '<VendorQueryRq requestID="6"><ActiveStatus>All</ActiveStatus></VendorQueryRq>'
    xml = session._send(request)
    vendors = []
    for b in xml_blocks(xml, 'VendorRet'):
        vendors.append({
            'list_id':   clean(xml_val(b, 'ListID')),
            'name':      clean(xml_val(b, 'Name')),
            'company':   clean(xml_val(b, 'CompanyName')),
            'contact':   clean(xml_val(b, 'FirstName') + ' ' + xml_val(b, 'LastName')).strip(),
            'phone':     clean(xml_val(b, 'Phone')),
            'fax':       clean(xml_val(b, 'Fax')),
            'email':     clean(xml_val(b, 'Email')),
            'addr1':     clean(xml_val(b, 'Addr1')),
            'addr2':     clean(xml_val(b, 'Addr2')),
            'city':      clean(xml_val(b, 'City')),
            'state':     clean(xml_val(b, 'State')),
            'zip':       clean(xml_val(b, 'PostalCode')),
            'terms':     clean(xml_ref(b, 'TermsRef', 'FullName')),
            'balance':   to_float(xml_val(b, 'Balance')),
            'active':    xml_val(b, 'IsActive') == 'true',
            'notes':     clean(xml_val(b, 'Notes')),
            'account_no':clean(xml_val(b, 'AccountNumber')),
        })
    print(f"    {len(vendors)} vendors found")
    return vendors


def extract_sales_reps(session):
    """Pull sales rep list."""
    print("  Extracting sales reps...")
    request = '<SalesRepQueryRq requestID="8"><ActiveStatus>All</ActiveStatus></SalesRepQueryRq>'
    xml = session._send(request)
    reps = []
    for b in xml_blocks(xml, 'SalesRepRet'):
        reps.append({
            'initial': clean(xml_val(b, 'Initial')),
            'name':    clean(xml_ref(b, 'SalesRepEntityRef', 'FullName')),
            'active':  xml_val(b, 'IsActive') == 'true',
        })
    print(f"    {len(reps)} sales reps found")
    return reps


def extract_terms(session):
    """Pull payment terms (Standard + DateDriven)."""
    print("  Extracting payment terms...")
    request = '<TermsQueryRq requestID="9"><ActiveStatus>All</ActiveStatus></TermsQueryRq>'
    xml = session._send(request)
    terms = []
    for b in xml_blocks(xml, 'StandardTermsRet'):
        terms.append({
            'name':      clean(xml_val(b, 'Name')),
            'type':      'Standard',
            'net_days':  int(xml_val(b, 'NetDays') or '0'),
            'disc_days': int(xml_val(b, 'DiscountDays') or '0'),
            'disc_pct':  to_float(xml_val(b, 'DiscountPct')),
        })
    for b in xml_blocks(xml, 'DateDrivenTermsRet'):
        terms.append({
            'name':      clean(xml_val(b, 'Name')),
            'type':      'DateDriven',
            'net_days':  int(xml_val(b, 'NetDays') or '0'),
            'disc_days': 0,
            'disc_pct':  to_float(xml_val(b, 'DiscountPct')),
        })
    print(f"    {len(terms)} terms found")
    return terms


def extract_ship_methods(session):
    """Pull ship methods. Invoices and SOs reference ShipMethodRef and silently
    fail on push if the method doesn't exist on the destination."""
    print("  Extracting ship methods...")
    request = '<ShipMethodQueryRq requestID="101"><ActiveStatus>All</ActiveStatus></ShipMethodQueryRq>'
    xml = session._send(request)
    out = []
    for b in xml_blocks(xml, 'ShipMethodRet'):
        out.append({
            'list_id': clean(xml_val(b, 'ListID')),
            'name':    clean(xml_val(b, 'Name')),
            'active':  xml_val(b, 'IsActive') == 'true',
        })
    print(f"    {len(out)} ship methods found")
    return out


def extract_payment_methods(session):
    """Pull payment methods. ReceivePayment / BillPaymentCheck reference these."""
    print("  Extracting payment methods...")
    request = '<PaymentMethodQueryRq requestID="102"><ActiveStatus>All</ActiveStatus></PaymentMethodQueryRq>'
    xml = session._send(request)
    out = []
    for b in xml_blocks(xml, 'PaymentMethodRet'):
        out.append({
            'list_id':      clean(xml_val(b, 'ListID')),
            'name':         clean(xml_val(b, 'Name')),
            'payment_type': clean(xml_val(b, 'PaymentMethodType')),
            'active':       xml_val(b, 'IsActive') == 'true',
        })
    print(f"    {len(out)} payment methods found")
    return out


def extract_sales_tax_codes(session):
    """Pull sales tax codes (per-line on invoices/CMs to mark taxable vs not)."""
    print("  Extracting sales tax codes...")
    request = '<SalesTaxCodeQueryRq requestID="200"><ActiveStatus>All</ActiveStatus></SalesTaxCodeQueryRq>'
    xml = session._send(request)
    out = []
    for b in xml_blocks(xml, 'SalesTaxCodeRet'):
        out.append({
            'list_id':     clean(xml_val(b, 'ListID')),
            'name':        clean(xml_val(b, 'Name')),
            'description': clean(xml_val(b, 'Desc')),
            'is_taxable':  xml_val(b, 'IsTaxable') == 'true',
            'active':      xml_val(b, 'IsActive') == 'true',
        })
    print(f"    {len(out)} sales tax codes found")
    return out


def extract_sales_tax_items(session):
    """Pull sales tax items (per-jurisdiction tax rate items, e.g. NY/NJ Sales Tax)."""
    print("  Extracting sales tax items...")
    request = '<ItemSalesTaxQueryRq requestID="201"><ActiveStatus>All</ActiveStatus></ItemSalesTaxQueryRq>'
    xml = session._send(request)
    out = []
    for b in xml_blocks(xml, 'ItemSalesTaxRet'):
        out.append({
            'list_id':     clean(xml_val(b, 'ListID')),
            'name':        clean(xml_val(b, 'Name')),
            'description': clean(xml_val(b, 'ItemDesc')),
            'tax_rate':    to_float(xml_val(b, 'TaxRate')),
            'tax_vendor':  clean(xml_ref(b, 'TaxVendorRef', 'FullName')),
            'active':      xml_val(b, 'IsActive') == 'true',
        })
    print(f"    {len(out)} sales tax items found")
    return out


# ============================================================================
# EXTRACTORS — TRANSACTIONS
# ============================================================================

def _parse_invoice_blocks(blocks):
    """Parse InvoiceRet XML blocks into dicts."""
    invoices = []
    for b in blocks:
        lines = []
        line_num = 1
        for lb in xml_blocks(b, 'InvoiceLineRet'):
            lines.append({
                'line_no':      line_num,
                'item':         clean(xml_ref(lb, 'ItemRef', 'FullName')),
                'item_list_id': clean(xml_ref(lb, 'ItemRef', 'ListID')),
                'description':  clean(xml_val(lb, 'Desc')),
                'qty':          to_float(xml_val(lb, 'Quantity')),
                'unitms':       clean(xml_val(lb, 'UnitOfMeasure')),
                'price':        to_float(xml_val(lb, 'Rate')),
                'ext_price':    to_float(xml_val(lb, 'Amount')),
                'tax_code':     clean(xml_ref(lb, 'SalesTaxCodeRef', 'FullName')),
            })
            line_num += 1

        invoices.append({
            'txn_id':       clean(xml_val(b, 'TxnID')),
            'inv_no':       clean(xml_val(b, 'RefNumber')),
            'inv_date':     clean(xml_val(b, 'TxnDate')),
            'due_date':     clean(xml_val(b, 'DueDate')),
            'cust_name':    clean(xml_ref(b, 'CustomerRef', 'FullName')),
            'cust_list_id': clean(xml_ref(b, 'CustomerRef', 'ListID')),
            'po_num':       clean(xml_val(b, 'PONumber')),
            'salesman':     clean(xml_ref(b, 'SalesRepRef', 'FullName')),
            'terms':        clean(xml_ref(b, 'TermsRef', 'FullName')),
            'ship_date':    clean(xml_val(b, 'ShipDate')),
            'ship_via':     clean(xml_ref(b, 'ShipMethodRef', 'FullName')),
            'inv_amt':      to_float(xml_val(b, 'Subtotal')),
            'tax_amt':      to_float(xml_val(b, 'SalesTaxTotal')),
            'tax_item':     clean(xml_ref(b, 'ItemSalesTaxRef', 'FullName')),
            'tax_code':     clean(xml_ref(b, 'CustomerSalesTaxCodeRef', 'FullName')),
            'total_amt':    to_float(xml_val(b, 'TotalAmount')),
            'balance':      to_float(xml_val(b, 'BalanceRemaining')),
            'memo':         clean(xml_val(b, 'Memo')),
            'is_paid':      xml_val(b, 'IsPaid') == 'true',
            'lines':        lines,
        })
    return invoices


def extract_invoices(session, years_back=3, date_range=None):
    """Pull invoice headers and lines. Chunked by year unless date_range given."""
    if date_range is not None:
        print(f"  Extracting invoices ({date_range[0] or '...'} to {date_range[1] or '...'})...")
    elif years_back == 0:
        print("  Extracting invoices (all history, chunked by year)...")
    else:
        print(f"  Extracting invoices (last {years_back} years)...")

    chunks = _build_date_chunks(years_back, date_range)
    all_invoices = []
    for from_date, to_date in chunks:
        request = f"""
    <InvoiceQueryRq requestID="7">
{_txn_date_filter(from_date, to_date)}
      <IncludeLineItems>true</IncludeLineItems>
      <OwnerID>0</OwnerID>
    </InvoiceQueryRq>
    """
        try:
            xml = session._send(request)
            batch = _parse_invoice_blocks(xml_blocks(xml, 'InvoiceRet'))
            if batch:
                all_invoices.extend(batch)
                print(f"    {from_date} to {to_date}: {len(batch)} invoices")
        except Exception as e:
            err = str(e)
            if 'timeout' in err.lower():
                print(f"    {from_date} to {to_date}: TIMEOUT -- skipping chunk")
            else:
                print(f"    {from_date} to {to_date}: ERROR -- {err[:100]}")

    print(f"    {len(all_invoices)} invoices found total")
    return all_invoices


def extract_payments(session, years_back=3, date_range=None):
    """Pull received customer payments."""
    if date_range is not None:
        print(f"  Extracting payments ({date_range[0] or '...'} to {date_range[1] or '...'})...")
    elif years_back == 0:
        print("  Extracting payments (all history, chunked by year)...")
    else:
        print(f"  Extracting payments (last {years_back} years)...")

    chunks = _build_date_chunks(years_back, date_range)
    all_payments = []
    for from_date, to_date in chunks:
        request = f"""
    <ReceivePaymentQueryRq requestID="20">
{_txn_date_filter(from_date, to_date)}
      <IncludeLineItems>true</IncludeLineItems>
      <OwnerID>0</OwnerID>
    </ReceivePaymentQueryRq>
    """
        try:
            xml = session._send(request)
            blocks = xml_blocks(xml, 'ReceivePaymentRet')
            for b in blocks:
                applied = []
                for ab in xml_blocks(b, 'AppliedToTxnRet'):
                    applied.append({
                        'inv_no':   clean(xml_ref(ab, 'TxnID', 'TxnID') or xml_val(ab, 'RefNumber')),
                        'ref_no':   clean(xml_val(ab, 'RefNumber')),
                        'amount':   to_float(xml_val(ab, 'Amount')),
                        'disc_amt': to_float(xml_val(ab, 'DiscountAmount')),
                    })
                all_payments.append({
                    'txn_id':       clean(xml_val(b, 'TxnID')),
                    'ref_no':       clean(xml_val(b, 'RefNumber')),
                    'date':         clean(xml_val(b, 'TxnDate')),
                    'cust_name':    clean(xml_ref(b, 'CustomerRef', 'FullName')),
                    'cust_list_id': clean(xml_ref(b, 'CustomerRef', 'ListID')),
                    'total_amt':    to_float(xml_val(b, 'TotalAmount')),
                    'method':       clean(xml_ref(b, 'PaymentMethodRef', 'FullName')),
                    'memo':         clean(xml_val(b, 'Memo')),
                    'applied':      applied,
                })
            if blocks:
                print(f"    {from_date} to {to_date}: {len(blocks)} payments")
        except Exception as e:
            print(f"    {from_date} to {to_date}: ERROR -- {str(e)[:100]}")

    print(f"    {len(all_payments)} payments found total")
    return all_payments


def extract_credit_memos(session, years_back=3, date_range=None):
    """Pull credit memos. Tax fields (subtotal, sales_tax_total, tax_item,
    tax_code, per-line tax_code) are critical — without them the push side
    can't reconstruct sales tax and the totals come out wrong."""
    if date_range is not None:
        print(f"  Extracting credit memos ({date_range[0] or '...'} to {date_range[1] or '...'})...")
    elif years_back == 0:
        print("  Extracting credit memos (all history, chunked by year)...")
    else:
        print(f"  Extracting credit memos (last {years_back} years)...")

    chunks = _build_date_chunks(years_back, date_range)
    all_memos = []
    for from_date, to_date in chunks:
        request = f"""
    <CreditMemoQueryRq requestID="21">
{_txn_date_filter(from_date, to_date)}
      <IncludeLineItems>true</IncludeLineItems>
      <OwnerID>0</OwnerID>
    </CreditMemoQueryRq>
    """
        try:
            xml = session._send(request)
            blocks = xml_blocks(xml, 'CreditMemoRet')
            for b in blocks:
                lines = []
                line_num = 1
                for lb in xml_blocks(b, 'CreditMemoLineRet'):
                    lines.append({
                        'line_no':      line_num,
                        'item':         clean(xml_ref(lb, 'ItemRef', 'FullName')),
                        'item_list_id': clean(xml_ref(lb, 'ItemRef', 'ListID')),
                        'description':  clean(xml_val(lb, 'Desc')),
                        'qty':          to_float(xml_val(lb, 'Quantity')),
                        'price':        to_float(xml_val(lb, 'Rate')),
                        'ext_price':    to_float(xml_val(lb, 'Amount')),
                        'tax_code':     clean(xml_ref(lb, 'SalesTaxCodeRef', 'FullName')),
                    })
                    line_num += 1
                all_memos.append({
                    'txn_id':          clean(xml_val(b, 'TxnID')),
                    'ref_no':          clean(xml_val(b, 'RefNumber')),
                    'date':            clean(xml_val(b, 'TxnDate')),
                    'cust_name':       clean(xml_ref(b, 'CustomerRef', 'FullName')),
                    'cust_list_id':    clean(xml_ref(b, 'CustomerRef', 'ListID')),
                    'subtotal':        to_float(xml_val(b, 'Subtotal')),
                    'sales_tax_total': to_float(xml_val(b, 'SalesTaxTotal')),
                    'sales_tax_pct':   to_float(xml_val(b, 'SalesTaxPercentage')),
                    'tax_item':        clean(xml_ref(b, 'ItemSalesTaxRef', 'FullName')),
                    'tax_code':        clean(xml_ref(b, 'CustomerSalesTaxCodeRef', 'FullName')),
                    'total_amt':       to_float(xml_val(b, 'TotalAmount')),
                    'balance':         to_float(xml_val(b, 'CreditRemaining')),
                    'memo':            clean(xml_val(b, 'Memo')),
                    'lines':           lines,
                })
            if blocks:
                print(f"    {from_date} to {to_date}: {len(blocks)} credit memos")
        except Exception as e:
            print(f"    {from_date} to {to_date}: ERROR -- {str(e)[:100]}")

    print(f"    {len(all_memos)} credit memos found total")
    return all_memos


def extract_sales_orders(session, date_range=None):
    """Pull sales orders. Optional date_range filters by TxnDate."""
    if date_range is not None:
        print(f"  Extracting sales orders ({date_range[0] or '...'} to {date_range[1] or '...'})...")
        filt = _txn_date_filter(date_range[0], date_range[1])
    else:
        print("  Extracting sales orders...")
        filt = ''
    request = f"""
    <SalesOrderQueryRq requestID="22">
{filt}
      <IncludeLineItems>true</IncludeLineItems>
      <OwnerID>0</OwnerID>
    </SalesOrderQueryRq>
    """
    try:
        xml = session._send(request)
    except Exception as e:
        if 'not enabled' in str(e).lower() or 'not supported' in str(e).lower():
            print("    Sales Orders not available (QB Pro feature)")
            return []
        raise
    orders = []
    for b in xml_blocks(xml, 'SalesOrderRet'):
        lines = []
        line_num = 1
        for lb in xml_blocks(b, 'SalesOrderLineRet'):
            lines.append({
                'line_no':      line_num,
                'item':         clean(xml_ref(lb, 'ItemRef', 'FullName')),
                'item_list_id': clean(xml_ref(lb, 'ItemRef', 'ListID')),
                'description':  clean(xml_val(lb, 'Desc')),
                'qty':          to_float(xml_val(lb, 'Quantity')),
                'price':        to_float(xml_val(lb, 'Rate')),
                'ext_price':    to_float(xml_val(lb, 'Amount')),
                'qty_invoiced': to_float(xml_val(lb, 'Invoiced')),
            })
            line_num += 1
        orders.append({
            'txn_id':       clean(xml_val(b, 'TxnID')),
            'ref_no':       clean(xml_val(b, 'RefNumber')),
            'date':         clean(xml_val(b, 'TxnDate')),
            'cust_name':    clean(xml_ref(b, 'CustomerRef', 'FullName')),
            'cust_list_id': clean(xml_ref(b, 'CustomerRef', 'ListID')),
            'po_num':       clean(xml_val(b, 'PONumber')),
            'ship_date':    clean(xml_val(b, 'ShipDate')),
            'ship_via':     clean(xml_ref(b, 'ShipMethodRef', 'FullName')),
            'salesman':     clean(xml_ref(b, 'SalesRepRef', 'FullName')),
            'total_amt':    to_float(xml_val(b, 'TotalAmount')),
            'is_fully_inv': xml_val(b, 'IsFullyInvoiced') == 'true',
            'memo':         clean(xml_val(b, 'Memo')),
            'lines':        lines,
        })
    print(f"    {len(orders)} sales orders found")
    return orders


def extract_purchase_orders(session, date_range=None):
    """Pull purchase orders. Optional date_range filters by TxnDate."""
    if date_range is not None:
        print(f"  Extracting purchase orders ({date_range[0] or '...'} to {date_range[1] or '...'})...")
        filt = _txn_date_filter(date_range[0], date_range[1])
    else:
        print("  Extracting purchase orders...")
        filt = ''
    request = f"""
    <PurchaseOrderQueryRq requestID="23">
{filt}
      <IncludeLineItems>true</IncludeLineItems>
      <OwnerID>0</OwnerID>
    </PurchaseOrderQueryRq>
    """
    try:
        xml = session._send(request)
    except Exception as e:
        if 'not enabled' in str(e).lower() or 'not supported' in str(e).lower():
            print("    Purchase Orders not available")
            return []
        raise
    orders = []
    for b in xml_blocks(xml, 'PurchaseOrderRet'):
        lines = []
        line_num = 1
        for lb in xml_blocks(b, 'PurchaseOrderLineRet'):
            lines.append({
                'line_no':      line_num,
                'item':         clean(xml_ref(lb, 'ItemRef', 'FullName')),
                'item_list_id': clean(xml_ref(lb, 'ItemRef', 'ListID')),
                'description':  clean(xml_val(lb, 'Desc')),
                'qty':          to_float(xml_val(lb, 'Quantity')),
                'price':        to_float(xml_val(lb, 'Rate')),
                'ext_price':    to_float(xml_val(lb, 'Amount')),
                'qty_received': to_float(xml_val(lb, 'ReceivedQuantity')),
            })
            line_num += 1
        orders.append({
            'txn_id':        clean(xml_val(b, 'TxnID')),
            'ref_no':        clean(xml_val(b, 'RefNumber')),
            'date':          clean(xml_val(b, 'TxnDate')),
            'expected_date': clean(xml_val(b, 'ExpectedDate')),
            'vend_name':     clean(xml_ref(b, 'VendorRef', 'FullName')),
            'vend_list_id':  clean(xml_ref(b, 'VendorRef', 'ListID')),
            'total_amt':     to_float(xml_val(b, 'TotalAmount')),
            'is_fully_rcvd': xml_val(b, 'IsFullyReceived') == 'true',
            'memo':          clean(xml_val(b, 'Memo')),
            'lines':         lines,
        })
    print(f"    {len(orders)} purchase orders found")
    return orders


def extract_bills(session, years_back=3, date_range=None):
    """Pull vendor bills (AP). Each line gets `line_type` = 'expense' or 'item'
    so the push side can pick ExpenseLineAdd vs ItemLineAdd (different schemas).
    A bill can mix both line types in the same transaction."""
    if date_range is not None:
        print(f"  Extracting vendor bills ({date_range[0] or '...'} to {date_range[1] or '...'})...")
    elif years_back == 0:
        print("  Extracting vendor bills (all history, chunked by year)...")
    else:
        print(f"  Extracting vendor bills (last {years_back} years)...")

    chunks = _build_date_chunks(years_back, date_range)
    all_bills = []
    for from_date, to_date in chunks:
        request = f"""
    <BillQueryRq requestID="24">
{_txn_date_filter(from_date, to_date)}
      <IncludeLineItems>true</IncludeLineItems>
      <OwnerID>0</OwnerID>
    </BillQueryRq>
    """
        try:
            xml = session._send(request)
            blocks = xml_blocks(xml, 'BillRet')
            for b in blocks:
                lines = []
                line_num = 1
                # Expense lines: gl_code is an Account FullName
                for lb in xml_blocks(b, 'ExpenseLineRet'):
                    lines.append({
                        'line_no':     line_num,
                        'line_type':   'expense',
                        'description': clean(xml_val(lb, 'Memo')),
                        'amount':      to_float(xml_val(lb, 'Amount')),
                        'gl_code':     clean(xml_ref(lb, 'AccountRef', 'FullName')),
                    })
                    line_num += 1
                # Item lines: gl_code is an Item FullName (NOT an account!)
                for lb in xml_blocks(b, 'ItemLineRet'):
                    lines.append({
                        'line_no':     line_num,
                        'line_type':   'item',
                        'description': clean(xml_val(lb, 'Desc')),
                        'amount':      to_float(xml_val(lb, 'Amount')),
                        'qty':         to_float(xml_val(lb, 'Quantity')),
                        'gl_code':     clean(xml_ref(lb, 'ItemRef', 'FullName')),
                    })
                    line_num += 1
                all_bills.append({
                    'txn_id':       clean(xml_val(b, 'TxnID')),
                    'ref_no':       clean(xml_val(b, 'RefNumber')),
                    'date':         clean(xml_val(b, 'TxnDate')),
                    'due_date':     clean(xml_val(b, 'DueDate')),
                    'vend_name':    clean(xml_ref(b, 'VendorRef', 'FullName')),
                    'vend_list_id': clean(xml_ref(b, 'VendorRef', 'ListID')),
                    'total_amt':    to_float(xml_val(b, 'AmountDue')),
                    'balance':      to_float(xml_val(b, 'OpenAmount')),
                    'is_paid':      xml_val(b, 'IsPaid') == 'true',
                    'terms':        clean(xml_ref(b, 'TermsRef', 'FullName')),
                    'memo':         clean(xml_val(b, 'Memo')),
                    'lines':        lines,
                })
            if blocks:
                print(f"    {from_date} to {to_date}: {len(blocks)} bills")
        except Exception as e:
            print(f"    {from_date} to {to_date}: ERROR -- {str(e)[:100]}")

    print(f"    {len(all_bills)} vendor bills found total")
    return all_bills


def extract_bill_payments(session, years_back=3, date_range=None):
    """Pull bill payment checks."""
    if date_range is not None:
        print(f"  Extracting bill payments ({date_range[0] or '...'} to {date_range[1] or '...'})...")
    elif years_back == 0:
        print("  Extracting bill payments (all history, chunked by year)...")
    else:
        print(f"  Extracting bill payments (last {years_back} years)...")

    chunks = _build_date_chunks(years_back, date_range)
    all_payments = []
    for from_date, to_date in chunks:
        request = f"""
    <BillPaymentCheckQueryRq requestID="25">
{_txn_date_filter(from_date, to_date)}
      <IncludeLineItems>true</IncludeLineItems>
      <OwnerID>0</OwnerID>
    </BillPaymentCheckQueryRq>
    """
        try:
            xml = session._send(request)
            blocks = xml_blocks(xml, 'BillPaymentCheckRet')
            for b in blocks:
                applied = []
                for ab in xml_blocks(b, 'AppliedToTxnRet'):
                    applied.append({
                        'ref_no': clean(xml_val(ab, 'RefNumber')),
                        'amount': to_float(xml_val(ab, 'Amount')),
                    })
                all_payments.append({
                    'txn_id':       clean(xml_val(b, 'TxnID')),
                    'ref_no':       clean(xml_val(b, 'RefNumber')),
                    'date':         clean(xml_val(b, 'TxnDate')),
                    'vend_name':    clean(xml_ref(b, 'PayeeEntityRef', 'FullName')),
                    'vend_list_id': clean(xml_ref(b, 'PayeeEntityRef', 'ListID')),
                    'total_amt':    to_float(xml_val(b, 'Amount')),
                    'applied':      applied,
                })
            if blocks:
                print(f"    {from_date} to {to_date}: {len(blocks)} bill payments")
        except Exception as e:
            print(f"    {from_date} to {to_date}: ERROR -- {str(e)[:100]}")

    print(f"    {len(all_payments)} bill payments found total")
    return all_payments


def extract_open_ar(session):
    """Pull open AR balances (unpaid invoices) — snapshot as of now."""
    print("  Extracting open AR balances...")
    request = '<InvoiceQueryRq requestID="10"><PaidStatus>NotPaidOnly</PaidStatus><OwnerID>0</OwnerID></InvoiceQueryRq>'
    xml = session._send(request)
    ar = []
    for b in xml_blocks(xml, 'InvoiceRet'):
        bal = to_float(xml_val(b, 'BalanceRemaining'))
        if bal != 0:
            ar.append({
                'inv_no':       clean(xml_val(b, 'RefNumber')),
                'inv_date':     clean(xml_val(b, 'TxnDate')),
                'due_date':     clean(xml_val(b, 'DueDate')),
                'cust_name':    clean(xml_ref(b, 'CustomerRef', 'FullName')),
                'cust_list_id': clean(xml_ref(b, 'CustomerRef', 'ListID')),
                'total_amt':    to_float(xml_val(b, 'TotalAmount')),
                'balance':      bal,
            })
    print(f"    {len(ar)} open AR items found")
    return ar


# ============================================================================
# GENERAL LEDGER — REPORT EXTRACTION (Option A: QBD report engine)
# ============================================================================
#
# Unlike every other extractor in this file, the General Ledger does NOT come
# back as a list of "<Foo>Ret" entity blocks. It is a *report* response:
#
#   <GeneralDetailReportQueryRs statusCode="0" ...>
#     <ReportRet>
#       <ReportTitle>General Ledger</ReportTitle>
#       <ReportBasis>Accrual</ReportBasis>
#       <ColDesc colID="1"><ColTitle titleRow="1" value="Type"/><ColType>TxnType</ColType></ColDesc>
#       <ColDesc colID="2">...</ColDesc>
#       ...
#       <ReportData>
#         <TextRow>...account-name section header...</TextRow>
#         <DataRow>
#           <RowData .../>                <!-- may carry the txn linkage/TxnID -->
#           <ColData colID="1" value="Check"/>
#           <ColData colID="2" value="2024-03-14"/>
#           ...
#         </DataRow>
#         <SubtotalRow>...</SubtotalRow>  <!-- skipped: re-aggregated downstream -->
#       </ReportData>
#     </ReportRet>
#   </GeneralDetailReportQueryRs>
#
# So the existing xml_blocks/xml_val entity helpers don't apply. We:
#   1. read ColDesc to map colID -> canonical field (Date/TxnType/Num/Name/...)
#   2. walk ReportData rows *in document order*, carrying the account from the
#      most recent account section header onto each DataRow beneath it
#   3. skip Subtotal/Total rows (downstream re-aggregates from raw amounts)
#   4. parse money display strings ("1,234.56", "(1,234.56)") to Decimal
#
# The one version-dependent unknown is whether the report exposes each line's
# internal TxnID. Run `--probe-gl` against the real company file FIRST; the
# probe dumps the raw report XML and reports exactly which columns and row-level
# IDs are present, so this parser can be trusted before it feeds drill-down.


def _tag_attrs(tag_text):
    """Parse XML attributes out of a start-tag/self-closing-tag body into a
    dict. Given `colID="1" value="Check"` returns {'colID':'1','value':'Check'}."""
    return dict(re.findall(r'(\w+)\s*=\s*"(.*?)"', tag_text, re.DOTALL))


def _norm_key(s):
    """Collapse a ColType/ColTitle to a lookup key: lowercase, alnum only, so
    'Ref Number', 'RefNumber' and 'refNumber' all map to the same bucket."""
    return re.sub(r'[^a-z0-9]', '', (s or '').lower())


def parse_amount(s):
    """Parse a QB report money/display string to Decimal. Handles thousands
    separators, a leading currency symbol, and parenthesized OR signed
    negatives. Returns None for blank/unparseable so callers can tell "no
    value" apart from 0. NEVER uses float() — the GL must reconcile to the
    penny, and binary float can't represent decimal cents exactly."""
    if s is None:
        return None
    t = s.strip()
    if not t:
        return None
    neg = False
    if t.startswith('(') and t.endswith(')'):
        neg = True
        t = t[1:-1].strip()
    t = t.replace(',', '').replace('$', '').replace(' ', '')
    if t.startswith('-'):
        neg, t = True, t[1:]
    elif t.startswith('+'):
        t = t[1:]
    if not t:
        return None
    try:
        d = Decimal(t)
    except (InvalidOperation, ValueError):
        return None
    return -d if neg else d


# Canonical GL field <- report ColType (preferred, more stable across versions)
# or ColTitle (fallback). The probe prints both so these can be tuned to the
# actual company file if a column comes back unmapped.
_GL_FIELD_BY_COLTYPE = {
    'txntype': 'txn_type',
    'date': 'date',
    'refnumber': 'ref_number',
    'number': 'ref_number',
    'name': 'name',
    'memo': 'memo',
    'splitaccount': 'split',
    'split': 'split',
    'account': 'split',
    'amount': 'amount',
    'runningbalance': 'running_balance',
    'balance': 'running_balance',
    'class': 'class_name',
    'quantity': 'quantity',
    'clr': 'cleared',
}
_GL_FIELD_BY_COLTITLE = {
    'type': 'txn_type',
    'date': 'date',
    'num': 'ref_number',
    'number': 'ref_number',
    'name': 'name',
    'memo': 'memo',
    'split': 'split',
    'amount': 'amount',
    'balance': 'running_balance',
    'class': 'class_name',
    'qty': 'quantity',
    'clr': 'cleared',
}


def _gl_field_for(coltype, coltitle):
    """Resolve a column's canonical field name: ColType first, ColTitle as
    fallback. Returns None for columns we don't carry (kept visible in probe)."""
    return (_GL_FIELD_BY_COLTYPE.get(_norm_key(coltype))
            or _GL_FIELD_BY_COLTITLE.get(_norm_key(coltitle)))


_COLDATA_RE = re.compile(r'<ColData\b((?:"[^"]*"|[^>])*)/?>')
_GL_ROW_RE = re.compile(r'<(TextRow|DataRow|SubtotalRow|TotalRow)\b(.*?)</\1>',
                        re.DOTALL)


def _parse_col_descs(report_xml):
    """Return {colID(int) -> {'coltype','coltitle','field'}} from <ColDesc>
    defs. ColType is an element; ColTitle is a self-closing tag with a `value`
    attr, possibly repeated (one per header row) — the last non-empty wins."""
    cols = {}
    for cd in re.findall(r'<ColDesc\b(.*?)</ColDesc>', report_xml, re.DOTALL):
        head = cd.split('>', 1)[0]
        try:
            col_id = int(_tag_attrs(head).get('colID', '0'))
        except ValueError:
            continue
        coltype = xml_val(cd, 'ColType')
        coltitle = ''
        for t in re.findall(r'<ColTitle\b([^>]*?)/?>', cd):
            v = _tag_attrs(t).get('value', '')
            if v:
                coltitle = v
        cols[col_id] = {
            'coltype':  coltype,
            'coltitle': coltitle,
            'field':    _gl_field_for(coltype, coltitle),
        }
    return cols


def _parse_rowdata(row_body):
    """Extract the row-level <RowData> of a report row: its rowType/value attrs
    and any nested ListID/TxnID (the account or transaction linkage)."""
    m = re.search(r'<RowData\b(.*?)(?:/>|>(.*?)</RowData>)', row_body, re.DOTALL)
    if not m:
        return {'rowType': '', 'value': '', 'list_id': '', 'txn_id': ''}
    attrs = _tag_attrs(m.group(1))
    inner = m.group(2) or ''
    return {
        'rowType': attrs.get('rowType', ''),
        'value':   attrs.get('value', ''),
        'list_id': xml_val(inner, 'ListID'),
        'txn_id':  xml_val(inner, 'TxnID'),
    }


def _index_accounts(accounts):
    """Index the extracted chart of accounts for GL enrichment: match a report's
    account row back to the master by ListID (best), then FullName, then Name."""
    idx = {'by_list_id': {}, 'by_full_name': {}, 'by_name': {}}
    for a in accounts or []:
        if a.get('list_id'):
            idx['by_list_id'][a['list_id']] = a
        if a.get('full_name'):
            idx['by_full_name'].setdefault(a['full_name'], a)
        if a.get('name'):
            idx['by_name'].setdefault(a['name'], a)
    return idx


def parse_general_ledger_report(xml, basis, accounts_index=None):
    """Parse one GeneralDetailReportQueryRs (General Ledger) into posting-line
    dicts, tagged with `basis` ('Accrual'|'Cash').

    Walks ReportData statefully. The account is carried from each row's
    <RowData rowType="account"> tag (present on both the section header and the
    transaction rows) onto the posting line. Header / opening-balance rows (no
    Txn Type, no Amount) and Subtotal/Total rows are skipped — downstream
    re-aggregates from raw amounts. `accounts_index` (from `_index_accounts`)
    enriches account_full_name/account_type/list_id from the extracted master
    instead of trusting the report's display text.

    Returns (rows, status_code, status_message)."""
    status_code, status_msg = '0', ''
    m = re.search(r'<GeneralDetailReportQueryRs\b([^>]*)>', xml)
    if m:
        a = _tag_attrs(m.group(1))
        status_code = a.get('statusCode', '0')
        status_msg = a.get('statusMessage', '')
    if status_code not in ('0', ''):
        return [], status_code, status_msg

    report_blocks = re.findall(r'<ReportRet\b.*?</ReportRet>', xml, re.DOTALL)
    if not report_blocks:
        return [], status_code, status_msg
    report = report_blocks[0]
    cols = _parse_col_descs(report)

    dm = re.search(r'<ReportData\b.*?</ReportData>', report, re.DOTALL)
    if not dm:
        return [], status_code, status_msg
    data = dm.group(0)

    rows_out = []
    cur_account = ''
    cur_account_list_id = ''
    for rm in _GL_ROW_RE.finditer(data):
        row_type, body = rm.group(1), rm.group(2)
        rd = _parse_rowdata(body)

        if row_type in ('SubtotalRow', 'TotalRow'):
            continue

        if row_type == 'TextRow':
            # Account section header. Some TextRows are blank spacers; adopt a
            # new account only when the row actually names one.
            label = clean(rd['value'])
            if not label:
                for c in _COLDATA_RE.findall(body):
                    v = _tag_attrs(c).get('value', '')
                    if v:
                        label = clean(v)
                        break
            if label:
                cur_account = label
                cur_account_list_id = rd['list_id']
            continue

        # DataRow. In the General Ledger the account is not merely a section
        # header — EVERY row (the account header AND each transaction beneath
        # it) carries <RowData rowType="account" value="<account>"> naming the
        # account it belongs to. So use rowType='account' to (re)set the current
        # account, but do NOT skip the row: fall through and parse its ColData.
        # The true header / opening-balance / beginning-balance rows carry no
        # Txn Type and no Amount and are dropped by the posting-line check below.
        if rd['rowType'] in ('account', 'section') and rd['value']:
            cur_account = clean(rd['value'])
            cur_account_list_id = rd['list_id']

        colvals = {}
        for c in _COLDATA_RE.findall(body):
            a = _tag_attrs(c)
            try:
                cid = int(a.get('colID', '0'))
            except ValueError:
                continue
            colvals[cid] = a.get('value', '')

        rec = {
            'txn_type': '', 'ref_number': '', 'date': '', 'name': '',
            'memo': '', 'split': '', 'class_name': '', 'amount': '',
            'running_balance': '',
        }
        for cid, meta in cols.items():
            field = meta.get('field')
            if not field or field not in rec or cid not in colvals:
                continue
            raw = colvals[cid]
            if field in ('amount', 'running_balance'):
                d = parse_amount(raw)
                rec[field] = str(d) if d is not None else ''
            else:
                rec[field] = clean(raw)

        # A real posting line has a transaction type and/or an amount. Account
        # header, opening-balance and beginning-balance rows have neither (only
        # a label plus a running balance) and are dropped. Running balances are
        # recomputed downstream, so nothing is lost.
        if not (rec['txn_type'] or rec['amount']):
            continue

        acct_name = cur_account
        acct_type = ''
        acct_list_id = cur_account_list_id
        if accounts_index:
            hit = None
            if acct_list_id and acct_list_id in accounts_index['by_list_id']:
                hit = accounts_index['by_list_id'][acct_list_id]
            elif cur_account in accounts_index['by_full_name']:
                hit = accounts_index['by_full_name'][cur_account]
            elif cur_account in accounts_index['by_name']:
                hit = accounts_index['by_name'][cur_account]
            if hit:
                acct_name = hit.get('full_name') or acct_name
                acct_type = hit.get('account_type', '')
                acct_list_id = hit.get('list_id') or acct_list_id

        rows_out.append({
            'txn_id':            clean(rd['txn_id']),
            'txn_type':          rec['txn_type'],
            'ref_number':        rec['ref_number'],
            'date':              rec['date'],
            'account_full_name': acct_name,
            'account_list_id':   acct_list_id,
            'account_type':      acct_type,
            'name':              rec['name'],
            'memo':              rec['memo'],
            'split':             rec['split'],
            'class':             rec['class_name'],
            'amount':            rec['amount'],
            'running_balance':   rec['running_balance'],
            'basis':             basis,
        })
    return rows_out, status_code, status_msg


def _add_months(d, n):
    """First-of-month date `n` months after `d` (n may be negative)."""
    m = d.month - 1 + n
    return datetime.date(d.year + m // 12, m % 12 + 1, 1)


def _build_report_period_chunks(years_back, date_range, granularity='month'):
    """Calendar-aligned (from, to) windows for report ReportPeriods.

    Unlike the rolling 365-day windows the entity txn queries use, GL report
    chunks align to month/quarter boundaries: each maps cleanly to a
    ReportPeriod and P&L reconciliation lands on real period boundaries. Chunk
    failures stay isolated to one period."""
    today = datetime.date.today()
    if date_range is not None:
        from_s, to_s = date_range
        start = (datetime.datetime.strptime(from_s, '%Y-%m-%d').date()
                 if from_s else datetime.date(today.year - 30, 1, 1))
        end = (datetime.datetime.strptime(to_s, '%Y-%m-%d').date()
               if to_s else today)
    else:
        yb = 30 if years_back == 0 else years_back
        start = datetime.date(today.year - yb + 1, 1, 1)
        end = today

    step = 3 if granularity == 'quarter' else 1
    if granularity == 'quarter':
        start = datetime.date(start.year, ((start.month - 1) // 3) * 3 + 1, 1)
    else:
        start = datetime.date(start.year, start.month, 1)

    windows = []
    cur = start
    while cur <= end:
        nxt = _add_months(cur, step)
        w_to = min(nxt - datetime.timedelta(days=1), end)
        windows.append((cur.strftime('%Y-%m-%d'), w_to.strftime('%Y-%m-%d')))
        cur = nxt
    return windows


def _gl_report_request(from_date, to_date, basis, request_id='300'):
    """Build a GeneralDetailReportQueryRq for the General Ledger over a period.
    Element order follows the qbXML v13 schema: type, period, basis."""
    return f"""
    <GeneralDetailReportQueryRq requestID="{request_id}">
      <GeneralDetailReportType>GeneralLedger</GeneralDetailReportType>
      <ReportPeriod>
        <FromReportDate>{from_date}</FromReportDate>
        <ToReportDate>{to_date}</ToReportDate>
      </ReportPeriod>
      <ReportBasis>{basis}</ReportBasis>
    </GeneralDetailReportQueryRq>
    """


def extract_general_ledger(session, years_back=3, date_range=None,
                           bases=('Accrual', 'Cash'), granularity='month',
                           accounts=None):
    """Extract the General Ledger via QBD's report engine (Option A).

    Runs GeneralDetailReportQueryRq(GeneralLedger) once per basis per calendar
    chunk and concatenates the posting lines. QBD generates the implicit
    balancing entries and — for basis='Cash' — computes cash-basis itself,
    which is otherwise the hardest part of P&L replication.

    Returns a flat list of posting-line dicts (see parse_general_ledger_report),
    each tagged with its `basis`. Chunk failures are isolated: one bad period is
    logged and skipped rather than losing the whole ledger. Amounts are exact
    decimal strings, not floats."""
    accounts_index = _index_accounts(accounts) if accounts else None
    chunks = _build_report_period_chunks(years_back, date_range, granularity)
    all_rows = []
    for basis in bases:
        print(f"  Extracting general ledger "
              f"[{basis} basis, {len(chunks)} {granularity} chunks]...")
        basis_rows = 0
        for from_date, to_date in chunks:
            request = _gl_report_request(from_date, to_date, basis)
            try:
                xml = session._send(request)
                rows, status_code, status_msg = parse_general_ledger_report(
                    xml, basis, accounts_index)
                if status_code not in ('0', ''):
                    print(f"    {from_date} to {to_date} [{basis}]: "
                          f"status {status_code} {status_msg[:80]}")
                    continue
                if rows:
                    all_rows.extend(rows)
                    basis_rows += len(rows)
                    print(f"    {from_date} to {to_date} [{basis}]: "
                          f"{len(rows)} lines")
            except Exception as e:
                print(f"    {from_date} to {to_date} [{basis}]: "
                      f"ERROR -- {str(e)[:100]}")
        print(f"    {basis_rows} {basis}-basis GL lines")
    print(f"    {len(all_rows)} general ledger lines found total")
    return all_rows


def probe_gl_report(session, date_range=None):
    """Diagnostic (handoff step 2): run ONE General Ledger report over a small
    window, dump the raw response XML, and report its structure — especially
    whether each posting line's internal TxnID is exposed. Run this against the
    real company file BEFORE trusting the GL extractor for drill-down/attachment
    linking. Writes <Company>_gl_probe_<YYYYMMDD>.xml and prints a summary."""
    today = datetime.date.today()
    if date_range is not None:
        from_date, to_date = date_range
        from_date = from_date or today.replace(day=1).strftime('%Y-%m-%d')
        to_date = to_date or today.strftime('%Y-%m-%d')
    else:
        # Last full calendar month.
        last_prev = today.replace(day=1) - datetime.timedelta(days=1)
        from_date = last_prev.replace(day=1).strftime('%Y-%m-%d')
        to_date = last_prev.strftime('%Y-%m-%d')

    print("=" * 60)
    print("GL REPORT PROBE")
    print("=" * 60)
    print(f"  Period: {from_date} to {to_date}  (Accrual basis)")
    request = _gl_report_request(from_date, to_date, 'Accrual', request_id='399')
    xml = session._send(request)

    safe_name = ''.join(c if c.isalnum() or c in (' ', '-', '_') else '_'
                        for c in (session.company_name or 'company')).strip()
    dump = os.path.join(os.getcwd(),
                        f"{safe_name}_gl_probe_{today.strftime('%Y%m%d')}.xml")
    with open(dump, 'w', encoding='utf-8') as f:
        f.write(xml)
    print(f"  Raw response written to: {os.path.basename(dump)}  "
          f"({len(xml)} bytes)")

    m = re.search(r'<GeneralDetailReportQueryRs\b([^>]*)>', xml)
    if m:
        a = _tag_attrs(m.group(1))
        print(f"  statusCode={a.get('statusCode', '?')} "
              f"statusMessage={a.get('statusMessage', '')[:80]}")

    report_blocks = re.findall(r'<ReportRet\b.*?</ReportRet>', xml, re.DOTALL)
    if not report_blocks:
        print("  No <ReportRet> in response — cannot probe columns.")
        print("=" * 60)
        return
    report = report_blocks[0]

    cols = _parse_col_descs(report)
    print(f"\n  Columns ({len(cols)}):")
    for cid in sorted(cols):
        c = cols[cid]
        print(f"    colID {cid}: ColType={c['coltype']!r} "
              f"ColTitle={c['coltitle']!r} -> field={c['field']}")

    dm = re.search(r'<ReportData\b.*?</ReportData>', report, re.DOTALL)
    data = dm.group(0) if dm else ''
    counts = {}
    rowtype_hist = {}
    for rm in _GL_ROW_RE.finditer(data):
        counts[rm.group(1)] = counts.get(rm.group(1), 0) + 1
        rt = _parse_rowdata(rm.group(2))['rowType'] or '(none)'
        rowtype_hist[rt] = rowtype_hist.get(rt, 0) + 1
    print(f"\n  Row types: {counts}")
    print(f"  RowData rowType histogram: {rowtype_hist}")

    # TxnID presence — THE key question for drill-down.
    data_rows = re.findall(r'<DataRow\b.*?</DataRow>', data, re.DOTALL)
    with_txnid = sum(1 for r in data_rows if '<TxnID>' in r)
    with_idlist = sum(1 for r in data_rows if '<IDList>' in r or '<ListID>' in r)
    print(f"\n  DataRows: {len(data_rows)}")
    print(f"    with <TxnID>: {with_txnid}   <-- drill-down/attachments need this")
    print(f"    with <ListID>/<IDList>: {with_idlist}")
    if with_txnid == 0:
        print("    => Report does NOT expose TxnID. Plan for RefNumber-based")
        print("       linking, or an Option-B txn join on Type+RefNumber+Date+Amount.")
    else:
        print("    => Report EXPOSES TxnID. Drill-down can link directly.")

    if data_rows:
        print("\n  Sample DataRow (raw, first 600 chars):")
        print("    " + data_rows[0].strip()[:600])

    parsed, _, _ = parse_general_ledger_report(xml, 'Accrual')
    print(f"\n  Parser extracted {len(parsed)} posting lines. First 3:")
    for r in parsed[:3]:
        print(f"    {r}")
    print("=" * 60)


# ============================================================================
# CLI / MAIN
# ============================================================================

def parse_args():
    """CLI args. All optional — falls back to interactive prompt for years
    when nothing is passed, so the EXE-double-click workflow still works."""
    p = argparse.ArgumentParser(
        description="Extract QuickBooks Desktop data to a single JSON bundle.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""\
Examples:
  QBExtract.py                                  # interactive
  QBExtract.py --years 3                        # last 3 years of transactions
  QBExtract.py --years 0                        # all history
  QBExtract.py --year 2024                      # only 2024
  QBExtract.py --from-date 2024-01-01 --to-date 2024-06-30
  QBExtract.py --corrupt-safe --years 0         # full history, corrupt file
  QBExtract.py --probe-gl --year 2024           # probe GL report structure, exit
  QBExtract.py --gl-basis accrual --years 3     # GL accrual only (default: both)
  QBExtract.py --no-gl --years 3                # skip General Ledger
""")
    p.add_argument('--years', type=int, default=None,
                   help='Years of transaction history (0=all, default=3 if interactive declines).')
    p.add_argument('--year', type=int, default=None,
                   help='Single year (e.g., 2024). Implies date range 2024-01-01..2024-12-31.')
    p.add_argument('--from-date', dest='from_date', default=None,
                   help='Transaction range start (YYYY-MM-DD). Overrides --years.')
    p.add_argument('--to-date', dest='to_date', default=None,
                   help='Transaction range end (YYYY-MM-DD). Overrides --years.')
    p.add_argument('--corrupt-safe', action='store_true',
                   help='Narrow master queries to non-history fields. Use for corrupt source files.')
    p.add_argument('--company-file', dest='company_file', default=None,
                   help='Full path to a .QBW file to open directly (e.g. '
                        r'"C:\QB\Company.QBW"). Default: use the file QuickBooks '
                        'currently has open. A path only works when QB is closed '
                        '(and this app has unattended access) or already open '
                        'with that same file.')
    p.add_argument('--gl-basis', dest='gl_basis',
                   choices=['accrual', 'cash', 'both'], default='accrual',
                   help='General Ledger report basis to extract (default: accrual). '
                        'Use "cash" or "both" to also/instead pull QBD-computed '
                        'cash basis.')
    p.add_argument('--gl-granularity', dest='gl_granularity',
                   choices=['month', 'quarter'], default='month',
                   help='Calendar chunk size for GL report queries (default: month). '
                        'Smaller chunks isolate failures to a shorter period.')
    p.add_argument('--no-gl', dest='no_gl', action='store_true',
                   help='Skip General Ledger extraction.')
    p.add_argument('--probe-gl', dest='probe_gl', action='store_true',
                   help='Run the GL report structure probe (TxnID check) and exit. '
                        'Run this against the real company file FIRST.')
    p.add_argument('--output', default=None,
                   help='Output JSON path. Default: <Company>_export_<YYYYMMDD>.json in cwd.')
    p.add_argument('--no-pause', dest='no_pause', action='store_true',
                   help='Do not wait for ENTER before exiting. Use when scripting '
                        'or scheduling the extractor.')
    return p.parse_args()


def _pause_before_exit(args=None):
    """Keep the console window open so a double-clicked EXE doesn't close before
    the user can read the output. Skipped when `--no-pause` is set, or when the
    tool was launched with any CLI argument (a terminal/scripted run), so it
    only ever blocks on a bare double-click."""
    if args is not None and getattr(args, 'no_pause', False):
        return
    if len(sys.argv) > 1:
        return
    try:
        input("\nPress ENTER to exit.")
    except EOFError:
        pass


def main():
    args = parse_args()

    print("=" * 60)
    print("QBExtract — QuickBooks Data Extractor (v3)")
    print("=" * 60)
    print()

    # Resolve date range vs years
    date_range = None
    if args.from_date or args.to_date:
        date_range = (args.from_date, args.to_date)
        years = 0  # ignored when date_range set
    elif args.year:
        date_range = (f'{args.year}-01-01', f'{args.year}-12-31')
        years = 0
    elif args.years is not None:
        years = args.years
    else:
        print("Make sure QuickBooks is OPEN with your company file loaded.")
        print()
        print("How many years of transaction history to export?")
        print("  1 = last 1 year (fastest)")
        print("  3 = last 3 years (recommended for demo)")
        print("  0 = all history (may be slow for large files)")
        try:
            years = int(input("Years [3]: ").strip() or "3")
        except ValueError:
            years = 3
        if years < 0:
            years = 0

    print()
    print("Connecting to QuickBooks...")

    session = QBSession()
    bundle = {}

    try:
        if args.company_file:
            print(f"Opening company file: {args.company_file}")
        session.connect(args.company_file or '')
        print()

        # GL structure probe: run against the real company file first to answer
        # the TxnID question, then exit without writing the full bundle.
        if args.probe_gl:
            probe_gl_report(session, date_range)
            return

        bundle['meta'] = {
            'company':        session.company_name,
            'company_file':   args.company_file or '',
            'exported_at':    datetime.datetime.now().isoformat(),
            'years_back':     years,
            'date_range':     date_range,
            'corrupt_safe':   args.corrupt_safe,
            'gl':             not args.no_gl,
            'gl_basis':       args.gl_basis,
            'gl_granularity': args.gl_granularity,
            'extractor':      'QBExtract v3.0',
        }

        # Masters first
        bundle['accounts']           = extract_accounts(session)
        bundle['terms']              = extract_terms(session)
        bundle['ship_methods']       = extract_ship_methods(session)
        bundle['payment_methods']    = extract_payment_methods(session)
        bundle['sales_tax_codes']    = extract_sales_tax_codes(session)
        bundle['sales_tax_items']    = extract_sales_tax_items(session)
        bundle['sales_reps']         = extract_sales_reps(session)
        bundle['price_levels']       = extract_price_levels(session)
        bundle['quantity_discounts'] = extract_quantity_discounts(session)
        bundle['vendors']            = extract_vendors(session)
        bundle['items']              = extract_items(session, corrupt_safe=args.corrupt_safe)
        bundle['customers']          = extract_customers(session, corrupt_safe=args.corrupt_safe)

        # Transactions
        bundle['invoices']           = extract_invoices(session, years, date_range)
        bundle['payments']           = extract_payments(session, years, date_range)
        bundle['credit_memos']       = extract_credit_memos(session, years, date_range)
        bundle['sales_orders']       = extract_sales_orders(session, date_range)
        bundle['purchase_orders']    = extract_purchase_orders(session, date_range)
        bundle['bills']              = extract_bills(session, years, date_range)
        bundle['bill_payments']      = extract_bill_payments(session, years, date_range)
        bundle['open_ar']            = extract_open_ar(session)

        # General Ledger (report engine — Option A). Wrapped so a report-engine
        # failure never loses the rest of the bundle.
        if not args.no_gl:
            bases = {'accrual': ('Accrual',), 'cash': ('Cash',),
                     'both': ('Accrual', 'Cash')}[args.gl_basis]
            try:
                bundle['general_ledger'] = extract_general_ledger(
                    session, years, date_range, bases=bases,
                    granularity=args.gl_granularity,
                    accounts=bundle.get('accounts'))
            except Exception as e:
                print(f"  GL extraction failed (continuing): {str(e)[:150]}")
                bundle['general_ledger'] = []
        else:
            bundle['general_ledger'] = []

    except Exception as e:
        print(f"\nERROR: {e}")
        print()
        traceback.print_exc()
        _pause_before_exit(args)
        sys.exit(1)
    finally:
        session.disconnect()

    # Write output file
    if args.output:
        filepath = args.output
        filename = os.path.basename(filepath)
    else:
        safe_name = ''.join(c if c.isalnum() or c in (' ', '-', '_') else '_'
                            for c in session.company_name).strip()
        date_str  = datetime.date.today().strftime('%Y%m%d')
        filename  = f"{safe_name}_export_{date_str}.json"
        filepath  = os.path.join(os.getcwd(), filename)

    print()
    print("Writing output file...")
    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(bundle, f, indent=2, ensure_ascii=False)

    size_mb = os.path.getsize(filepath) / (1024 * 1024)

    print()
    print("=" * 60)
    print("DONE")
    print("=" * 60)
    print(f"  Company:           {session.company_name}")
    print(f"  Accounts:          {len(bundle['accounts'])}")
    print(f"  Customers:         {len(bundle['customers'])}")
    print(f"  Items:             {len(bundle['items'])}")
    print(f"  Vendors:           {len(bundle['vendors'])}")
    print(f"  Invoices:          {len(bundle['invoices'])}")
    print(f"  Payments:          {len(bundle['payments'])}")
    print(f"  Credit Memos:      {len(bundle['credit_memos'])}")
    print(f"  Sales Orders:      {len(bundle['sales_orders'])}")
    print(f"  Purchase Orders:   {len(bundle['purchase_orders'])}")
    print(f"  Vendor Bills:      {len(bundle['bills'])}")
    print(f"  Bill Payments:     {len(bundle['bill_payments'])}")
    print(f"  Open AR items:     {len(bundle['open_ar'])}")
    gl = bundle.get('general_ledger', [])
    print(f"  General Ledger:    {len(gl)} lines")
    if gl:
        by_basis = {}
        for r in gl:
            by_basis[r['basis']] = by_basis.get(r['basis'], 0) + 1
        for b in sorted(by_basis):
            print(f"    {b} basis:       {by_basis[b]} lines")
    print(f"  File:              {filename}  ({size_mb:.1f} MB)")
    print()
    print("Send this file to your ERP consultant for import.")
    print()
    _pause_before_exit(args)


if __name__ == '__main__':
    main()
