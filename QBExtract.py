"""
QBExtract.py — QuickBooks Desktop Data Extractor
=================================================
Reads data directly from a running QuickBooks Desktop/Enterprise company file
via the QB SDK (QBXML) and produces a single JSON bundle for import.

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


# ============================================================================
# QB SESSION — uses VBScript for COM (100% reliable on any Windows)
# ============================================================================

# VBScript template that handles all QB SDK COM calls.
# Python writes the QBXML request to a file, calls this script,
# and reads the response from another file.
VBS_TEMPLATE = r'''
' QBExtract VBScript bridge — called by Python
' Args: action, request_file, response_file
Dim action, reqFile, resFile
action  = WScript.Arguments(0)
reqFile = WScript.Arguments(1)
resFile = WScript.Arguments(2)

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
    ticket = rp.BeginSession("", 2)
    If Err.Number <> 0 Then
        Err.Clear
        ticket = rp.BeginSession("", 0)
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
    ticket = rp.BeginSession("", 2)
    If Err.Number <> 0 Then
        Err.Clear
        ticket = rp.BeginSession("", 0)
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


class QBSession:
    """Manages connection to QuickBooks via VBScript COM bridge."""

    def __init__(self):
        self.company_name = ''
        self._vbs_path = None
        self._req_path = None
        self._res_path = None

    def connect(self):
        # Write VBScript bridge to temp file
        self._vbs_path = os.path.join(tempfile.gettempdir(), 'qbextract_bridge.vbs')
        self._req_path = os.path.join(tempfile.gettempdir(), 'qbextract_req.xml')
        self._res_path = os.path.join(tempfile.gettempdir(), 'qbextract_res.xml')

        with open(self._vbs_path, 'w', encoding='utf-8') as f:
            f.write(VBS_TEMPLATE)

        # Test connection
        result = self._run_vbs('connect')
        if result.startswith('ERROR:'):
            raise RuntimeError(result)

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
                 action, self._req_path, self._res_path],
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
        """Send QBXML request via VBScript bridge, return response XML."""
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
    p.add_argument('--output', default=None,
                   help='Output JSON path. Default: <Company>_export_<YYYYMMDD>.json in cwd.')
    return p.parse_args()


def main():
    args = parse_args()

    print("=" * 60)
    print("QBExtract — QuickBooks Data Extractor (v2)")
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
        session.connect()
        print()

        bundle['meta'] = {
            'company':       session.company_name,
            'exported_at':   datetime.datetime.now().isoformat(),
            'years_back':    years,
            'date_range':    date_range,
            'corrupt_safe':  args.corrupt_safe,
            'extractor':     'QBExtract v2.0',
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

    except Exception as e:
        print(f"\nERROR: {e}")
        print()
        traceback.print_exc()
        try:
            input("\nPress ENTER to exit.")
        except EOFError:
            pass
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
    print(f"  File:              {filename}  ({size_mb:.1f} MB)")
    print()
    print("Send this file to your ERP consultant for import.")
    print()
    try:
        input("Press ENTER to exit.")
    except EOFError:
        pass


if __name__ == '__main__':
    main()
