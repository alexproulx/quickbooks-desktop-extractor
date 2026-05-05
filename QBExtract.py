"""
QBExtract.py — QuickBooks Desktop Data Extractor
=================================================
Reads data directly from a running QuickBooks Desktop/Enterprise company file
via the QB SDK (QBXML) and produces a single JSON bundle for import.

Requirements:
  - QuickBooks Desktop/Enterprise must be OPEN with the company file loaded
  - QuickBooks SDK must be installed (comes with QB or download from Intuit)
  - No Python or pywin32 needed when compiled to EXE (uses VBScript for COM)

Usage:
  python QBExtract.py
  (or compiled to QBExtract.exe via PyInstaller)

Output:
  <CompanyName>_export_<YYYYMMDD>.json  in the same folder as this script

Compile to EXE:
  pip install pyinstaller
  pyinstaller --onefile --console QBExtract.py
"""

import os
import sys
import json
import datetime
import traceback

import subprocess
import tempfile


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
               "<QBXML><QBXMLMsgsRq onError=""stopOnError"">" & _
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
        # Ensure req/res files exist
        if not os.path.exists(self._req_path):
            with open(self._req_path, 'w') as f:
                f.write('')

        # Clear response file
        with open(self._res_path, 'w') as f:
            f.write('')

        try:
            proc = subprocess.run(
                ['cscript', '//Nologo', self._vbs_path,
                 action, self._req_path, self._res_path],
                capture_output=True, text=True, timeout=300
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
        import re
        m = re.search(r'<CompanyName>(.*?)</CompanyName>', xml)
        return m.group(1) if m else 'UnknownCompany'


# ============================================================================
# XML HELPERS
# ============================================================================

def xml_val(xml, tag, default=''):
    """Extract first value of a tag from XML string."""
    import re
    m = re.search(rf'<{tag}>(.*?)</{tag}>', xml, re.DOTALL)
    return m.group(1).strip() if m else default


def xml_all(xml, tag):
    """Extract all occurrences of a tag."""
    import re
    return re.findall(rf'<{tag}>(.*?)</{tag}>', xml, re.DOTALL)


def xml_blocks(xml, tag):
    """Extract all blocks between opening and closing tags."""
    import re
    return re.findall(rf'<{tag}[\s>].*?</{tag}>', xml, re.DOTALL)


def xml_ref(xml, ref_tag, sub_tag='FullName'):
    """Extract a sub-tag from a Ref block, e.g. CustomerRef -> FullName.
    QB returns nested refs like <CustomerRef><FullName>X</FullName></CustomerRef>
    """
    import re
    m = re.search(rf'<{ref_tag}>(.*?)</{ref_tag}>', xml, re.DOTALL)
    if not m:
        return ''
    inner = m.group(1)
    m2 = re.search(rf'<{sub_tag}>(.*?)</{sub_tag}>', inner, re.DOTALL)
    return m2.group(1).strip() if m2 else ''


def clean(s):
    """Strip whitespace and XML entities."""
    if not s:
        return ''
    return (s.strip()
            .replace('&amp;', '&')
            .replace('&lt;', '<')
            .replace('&gt;', '>')
            .replace('&quot;', '"')
            .replace('&#39;', "'"))


def to_float(s):
    try:
        return float(s.strip()) if s and s.strip() else 0.0
    except (ValueError, AttributeError):
        return 0.0


# ============================================================================
# EXTRACTORS
# ============================================================================

def extract_customers(session):
    """Pull all customers with address, terms, price level, balance."""
    print("  Extracting customers...")
    request = """
    <CustomerQueryRq requestID="2">
      <ActiveStatus>All</ActiveStatus>
      <OwnerID>0</OwnerID>
    </CustomerQueryRq>
    """
    xml = session._send(request)
    blocks = xml_blocks(xml, 'CustomerRet')
    customers = []
    for b in blocks:
        # Skip sub-customers (jobs) — they have a ParentRef
        if '<ParentRef>' in b:
            continue
        # Extract address from BillAddress sub-block to avoid mixing with ShipAddress
        bill_blocks = xml_blocks(b, 'BillAddress')
        bill = bill_blocks[0] if bill_blocks else ''
        cust = {
            'list_id':      clean(xml_val(b, 'ListID')),
            'name':         clean(xml_val(b, 'Name')),
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
        }
        customers.append(cust)

    print(f"    {len(customers)} customers found")
    return customers


def extract_items(session):
    """Pull all inventory items and non-inventory items with all price levels."""
    print("  Extracting items...")

    # Pull inventory items
    request_inv = """
    <ItemInventoryQueryRq requestID="3">
      <ActiveStatus>All</ActiveStatus>
      <OwnerID>0</OwnerID>
    </ItemInventoryQueryRq>
    """
    # Pull non-inventory items (services, etc.)
    request_non = """
    <ItemNonInventoryQueryRq requestID="4">
      <ActiveStatus>All</ActiveStatus>
      <OwnerID>0</OwnerID>
    </ItemNonInventoryQueryRq>
    """

    items = []

    for req, item_type in [(request_inv, 'INV'), (request_non, 'NON')]:
        xml = session._send(req)
        tag = 'ItemInventoryRet' if item_type == 'INV' else 'ItemNonInventoryRet'
        blocks = xml_blocks(xml, tag)
        for b in blocks:
            item = {
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
            }
            items.append(item)

    print(f"    {len(items)} items found")
    return items


def extract_price_levels(session):
    """Pull all price levels (Enterprise feature).
    Returns:
      - price_levels: list of {name, type, pct, fixed_prices: [{item_list_id, price}]}
    ERP supports 5 tiers (A-E); extras are exported and the importer
    will store them as customer-specific price_exceptions.
    """
    print("  Extracting price levels...")
    request = """
    <PriceLevelQueryRq requestID="5">
      <ActiveStatus>All</ActiveStatus>
    </PriceLevelQueryRq>
    """
    xml = session._send(request)
    blocks = xml_blocks(xml, 'PriceLevelRet')
    price_levels = []

    for b in blocks:
        pl_name = clean(xml_val(b, 'Name'))
        pl_type = clean(xml_val(b, 'PriceLevelType'))  # 'FixedPercentage' or 'PerItem'
        pct      = to_float(xml_val(b, 'PriceLevelFixedPct'))

        # Per-item price overrides (Enterprise feature)
        fixed_prices = []
        item_blocks = xml_blocks(b, 'PriceLevelPerItemRet')
        for ib in item_blocks:
            fixed_prices.append({
                'item_list_id':  clean(xml_ref(ib, 'ItemRef', 'ListID')),
                'item_name':     clean(xml_ref(ib, 'ItemRef', 'FullName')),
                'custom_price':  to_float(xml_val(ib, 'CustomPrice')),
                'custom_pct':    to_float(xml_val(ib, 'CustomPricePercent')),
            })

        price_levels.append({
            'name':         pl_name,
            'type':         pl_type,
            'pct':          pct,
            'fixed_prices': fixed_prices,
        })

    print(f"    {len(price_levels)} price levels found")
    if len(price_levels) > 5:
        print(f"    NOTE: ERP supports 5 price tiers (A-E). {len(price_levels) - 5} extra levels")
        print(f"          will be imported as customer-specific price exceptions.")
    return price_levels


def extract_quantity_discounts(session):
    """Pull quantity-discount items from QB (ItemDiscount + ItemGroup with qty breaks).
    QB doesn't expose Enterprise Price Rules via QBXML, but we can extract:
      - ItemDiscount entries (flat discount items applied on invoices)
      - ItemGroup entries (bundles that imply qty pricing)
    Returns list of {item_list_id, item_name, discount_rate, discount_type}
    """
    print("  Extracting quantity discounts...")

    # Discount items (used on invoices for line-level discounts)
    request = """
    <ItemDiscountQueryRq requestID="11">
      <ActiveStatus>All</ActiveStatus>
    </ItemDiscountQueryRq>
    """
    xml = session._send(request)
    blocks = xml_blocks(xml, 'ItemDiscountRet')
    discounts = []

    for b in blocks:
        rate = to_float(xml_val(b, 'DiscountRate'))
        pct  = to_float(xml_val(b, 'DiscountRatePercent'))
        discounts.append({
            'list_id':       clean(xml_val(b, 'ListID')),
            'item_name':     clean(xml_val(b, 'Name')),
            'description':   clean(xml_val(b, 'ItemDesc')),
            'discount_rate': rate,
            'discount_pct':  pct,
            'account':       clean(xml_ref(b, 'AccountRef', 'FullName')),
            'active':        xml_val(b, 'IsActive') == 'true',
        })

    print(f"    {len(discounts)} discount items found")
    return discounts


def extract_vendors(session):
    """Pull all vendors."""
    print("  Extracting vendors...")
    request = """
    <VendorQueryRq requestID="6">
      <ActiveStatus>All</ActiveStatus>
    </VendorQueryRq>
    """
    xml = session._send(request)
    blocks = xml_blocks(xml, 'VendorRet')
    vendors = []
    for b in blocks:
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


def _parse_invoice_blocks(blocks):
    """Parse InvoiceRet XML blocks into dicts."""
    invoices = []
    for b in blocks:
        lines = []
        line_num = 1
        line_blocks = xml_blocks(b, 'InvoiceLineRet')
        for lb in line_blocks:
            lines.append({
                'line_no':     line_num,
                'item':        clean(xml_ref(lb, 'ItemRef', 'FullName')),
                'item_list_id':clean(xml_ref(lb, 'ItemRef', 'ListID')),
                'description': clean(xml_val(lb, 'Desc')),
                'qty':         to_float(xml_val(lb, 'Quantity')),
                'unitms':      clean(xml_val(lb, 'UnitOfMeasure')),
                'price':       to_float(xml_val(lb, 'Rate')),
                'ext_price':   to_float(xml_val(lb, 'Amount')),
                'tax_code':    clean(xml_ref(lb, 'SalesTaxCodeRef', 'FullName')),
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
            'total_amt':    to_float(xml_val(b, 'TotalAmount')),
            'balance':      to_float(xml_val(b, 'BalanceRemaining')),
            'memo':         clean(xml_val(b, 'Memo')),
            'is_paid':      xml_val(b, 'IsPaid') == 'true',
            'lines':        lines,
        })
    return invoices


def extract_invoices(session, years_back=3):
    """Pull invoice headers and lines, chunked by year to avoid QB timeouts."""
    if years_back == 0:
        print("  Extracting invoices (all history, chunked by year)...")
    else:
        print(f"  Extracting invoices (last {years_back} years)...")

    # Build list of (from_date, to_date) chunks
    today = datetime.date.today()
    if years_back == 0:
        # Go back 30 years in 1-year chunks — covers any realistic QB history
        max_years = 30
    else:
        max_years = years_back

    chunks = []
    for i in range(max_years):
        chunk_end   = today - datetime.timedelta(days=365 * i)
        chunk_start = today - datetime.timedelta(days=365 * (i + 1))
        chunks.append((chunk_start.strftime('%Y-%m-%d'), chunk_end.strftime('%Y-%m-%d')))

    all_invoices = []
    for from_date, to_date in chunks:
        request = f"""
    <InvoiceQueryRq requestID="7">
      <TxnDateRangeFilter>
        <FromTxnDate>{from_date}</FromTxnDate>
        <ToTxnDate>{to_date}</ToTxnDate>
      </TxnDateRangeFilter>
      <IncludeLineItems>true</IncludeLineItems>
      <OwnerID>0</OwnerID>
    </InvoiceQueryRq>
    """
        try:
            xml = session._send(request)
            blocks = xml_blocks(xml, 'InvoiceRet')
            if blocks:
                batch = _parse_invoice_blocks(blocks)
                all_invoices.extend(batch)
                print(f"    {from_date} to {to_date}: {len(batch)} invoices")
        except Exception as e:
            err = str(e)
            if 'timeout' in err.lower():
                print(f"    {from_date} to {to_date}: TIMEOUT — skipping chunk")
            else:
                print(f"    {from_date} to {to_date}: ERROR — {err[:100]}")

    print(f"    {len(all_invoices)} invoices found total")
    return all_invoices


def extract_sales_reps(session):
    """Pull sales rep list."""
    print("  Extracting sales reps...")
    request = """
    <SalesRepQueryRq requestID="8">
      <ActiveStatus>All</ActiveStatus>
    </SalesRepQueryRq>
    """
    xml = session._send(request)
    blocks = xml_blocks(xml, 'SalesRepRet')
    reps = []
    for b in blocks:
        reps.append({
            'initial':   clean(xml_val(b, 'Initial')),
            'name':      clean(xml_ref(b, 'SalesRepEntityRef', 'FullName')),
            'active':    xml_val(b, 'IsActive') == 'true',
        })
    print(f"    {len(reps)} sales reps found")
    return reps


def extract_terms(session):
    """Pull payment terms."""
    print("  Extracting payment terms...")
    request = """
    <TermsQueryRq requestID="9">
      <ActiveStatus>All</ActiveStatus>
    </TermsQueryRq>
    """
    xml = session._send(request)
    # Standard terms
    std_blocks  = xml_blocks(xml, 'StandardTermsRet')
    date_blocks = xml_blocks(xml, 'DateDrivenTermsRet')
    terms = []
    for b in std_blocks:
        terms.append({
            'name':       clean(xml_val(b, 'Name')),
            'type':       'Standard',
            'net_days':   int(xml_val(b, 'NetDays') or '0'),
            'disc_days':  int(xml_val(b, 'DiscountDays') or '0'),
            'disc_pct':   to_float(xml_val(b, 'DiscountPct')),
        })
    for b in date_blocks:
        terms.append({
            'name':       clean(xml_val(b, 'Name')),
            'type':       'DateDriven',
            'net_days':   int(xml_val(b, 'NetDays') or '0'),
            'disc_days':  0,
            'disc_pct':   to_float(xml_val(b, 'DiscountPct')),
        })
    print(f"    {len(terms)} terms found")
    return terms


def extract_payments(session, years_back=3):
    """Pull received payments, chunked by year like invoices."""
    if years_back == 0:
        print("  Extracting payments (all history, chunked by year)...")
    else:
        print(f"  Extracting payments (last {years_back} years)...")

    today = datetime.date.today()
    max_years = 30 if years_back == 0 else years_back

    chunks = []
    for i in range(max_years):
        chunk_end   = today - datetime.timedelta(days=365 * i)
        chunk_start = today - datetime.timedelta(days=365 * (i + 1))
        chunks.append((chunk_start.strftime('%Y-%m-%d'), chunk_end.strftime('%Y-%m-%d')))

    all_payments = []
    for from_date, to_date in chunks:
        request = f"""
    <ReceivePaymentQueryRq requestID="20">
      <TxnDateRangeFilter>
        <FromTxnDate>{from_date}</FromTxnDate>
        <ToTxnDate>{to_date}</ToTxnDate>
      </TxnDateRangeFilter>
      <IncludeLineItems>true</IncludeLineItems>
      <OwnerID>0</OwnerID>
    </ReceivePaymentQueryRq>
    """
        try:
            xml = session._send(request)
            blocks = xml_blocks(xml, 'ReceivePaymentRet')
            for b in blocks:
                # Each payment can apply to multiple invoices
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
            err = str(e)
            if 'timeout' in err.lower():
                print(f"    {from_date} to {to_date}: TIMEOUT -- skipping")
            else:
                print(f"    {from_date} to {to_date}: ERROR -- {err[:100]}")

    print(f"    {len(all_payments)} payments found total")
    return all_payments


def extract_credit_memos(session, years_back=3):
    """Pull credit memos, chunked by year."""
    if years_back == 0:
        print("  Extracting credit memos (all history, chunked by year)...")
    else:
        print(f"  Extracting credit memos (last {years_back} years)...")

    today = datetime.date.today()
    max_years = 30 if years_back == 0 else years_back

    chunks = []
    for i in range(max_years):
        chunk_end   = today - datetime.timedelta(days=365 * i)
        chunk_start = today - datetime.timedelta(days=365 * (i + 1))
        chunks.append((chunk_start.strftime('%Y-%m-%d'), chunk_end.strftime('%Y-%m-%d')))

    all_memos = []
    for from_date, to_date in chunks:
        request = f"""
    <CreditMemoQueryRq requestID="21">
      <TxnDateRangeFilter>
        <FromTxnDate>{from_date}</FromTxnDate>
        <ToTxnDate>{to_date}</ToTxnDate>
      </TxnDateRangeFilter>
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
                        'line_no':     line_num,
                        'item':        clean(xml_ref(lb, 'ItemRef', 'FullName')),
                        'item_list_id':clean(xml_ref(lb, 'ItemRef', 'ListID')),
                        'description': clean(xml_val(lb, 'Desc')),
                        'qty':         to_float(xml_val(lb, 'Quantity')),
                        'price':       to_float(xml_val(lb, 'Rate')),
                        'ext_price':   to_float(xml_val(lb, 'Amount')),
                    })
                    line_num += 1

                all_memos.append({
                    'txn_id':       clean(xml_val(b, 'TxnID')),
                    'ref_no':       clean(xml_val(b, 'RefNumber')),
                    'date':         clean(xml_val(b, 'TxnDate')),
                    'cust_name':    clean(xml_ref(b, 'CustomerRef', 'FullName')),
                    'cust_list_id': clean(xml_ref(b, 'CustomerRef', 'ListID')),
                    'total_amt':    to_float(xml_val(b, 'TotalAmount')),
                    'balance':      to_float(xml_val(b, 'CreditRemaining')),
                    'memo':         clean(xml_val(b, 'Memo')),
                    'lines':        lines,
                })
            if blocks:
                print(f"    {from_date} to {to_date}: {len(blocks)} credit memos")
        except Exception as e:
            err = str(e)
            if 'timeout' in err.lower():
                print(f"    {from_date} to {to_date}: TIMEOUT -- skipping")
            else:
                print(f"    {from_date} to {to_date}: ERROR -- {err[:100]}")

    print(f"    {len(all_memos)} credit memos found total")
    return all_memos


def extract_sales_orders(session):
    """Pull open/all sales orders."""
    print("  Extracting sales orders...")
    request = """
    <SalesOrderQueryRq requestID="22">
      <IncludeLineItems>true</IncludeLineItems>
      <OwnerID>0</OwnerID>
    </SalesOrderQueryRq>
    """
    try:
        xml = session._send(request)
    except Exception as e:
        if 'not enabled' in str(e).lower() or 'not supported' in str(e).lower():
            print("    Sales Orders not available (QB Pro/Premier feature)")
            return []
        raise
    blocks = xml_blocks(xml, 'SalesOrderRet')
    orders = []
    for b in blocks:
        lines = []
        line_num = 1
        for lb in xml_blocks(b, 'SalesOrderLineRet'):
            lines.append({
                'line_no':     line_num,
                'item':        clean(xml_ref(lb, 'ItemRef', 'FullName')),
                'item_list_id':clean(xml_ref(lb, 'ItemRef', 'ListID')),
                'description': clean(xml_val(lb, 'Desc')),
                'qty':         to_float(xml_val(lb, 'Quantity')),
                'price':       to_float(xml_val(lb, 'Rate')),
                'ext_price':   to_float(xml_val(lb, 'Amount')),
                'qty_invoiced':to_float(xml_val(lb, 'Invoiced')),
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


def extract_purchase_orders(session):
    """Pull purchase orders."""
    print("  Extracting purchase orders...")
    request = """
    <PurchaseOrderQueryRq requestID="23">
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
    blocks = xml_blocks(xml, 'PurchaseOrderRet')
    orders = []
    for b in blocks:
        lines = []
        line_num = 1
        for lb in xml_blocks(b, 'PurchaseOrderLineRet'):
            lines.append({
                'line_no':     line_num,
                'item':        clean(xml_ref(lb, 'ItemRef', 'FullName')),
                'item_list_id':clean(xml_ref(lb, 'ItemRef', 'ListID')),
                'description': clean(xml_val(lb, 'Desc')),
                'qty':         to_float(xml_val(lb, 'Quantity')),
                'price':       to_float(xml_val(lb, 'Rate')),
                'ext_price':   to_float(xml_val(lb, 'Amount')),
                'qty_received':to_float(xml_val(lb, 'ReceivedQuantity')),
            })
            line_num += 1

        orders.append({
            'txn_id':       clean(xml_val(b, 'TxnID')),
            'ref_no':       clean(xml_val(b, 'RefNumber')),
            'date':         clean(xml_val(b, 'TxnDate')),
            'expected_date':clean(xml_val(b, 'ExpectedDate')),
            'vend_name':    clean(xml_ref(b, 'VendorRef', 'FullName')),
            'vend_list_id': clean(xml_ref(b, 'VendorRef', 'ListID')),
            'total_amt':    to_float(xml_val(b, 'TotalAmount')),
            'is_fully_rcvd':xml_val(b, 'IsFullyReceived') == 'true',
            'memo':         clean(xml_val(b, 'Memo')),
            'lines':        lines,
        })
    print(f"    {len(orders)} purchase orders found")
    return orders


def extract_bills(session, years_back=3):
    """Pull vendor bills (AP), chunked by year."""
    if years_back == 0:
        print("  Extracting vendor bills (all history, chunked by year)...")
    else:
        print(f"  Extracting vendor bills (last {years_back} years)...")

    today = datetime.date.today()
    max_years = 30 if years_back == 0 else years_back

    chunks = []
    for i in range(max_years):
        chunk_end   = today - datetime.timedelta(days=365 * i)
        chunk_start = today - datetime.timedelta(days=365 * (i + 1))
        chunks.append((chunk_start.strftime('%Y-%m-%d'), chunk_end.strftime('%Y-%m-%d')))

    all_bills = []
    for from_date, to_date in chunks:
        request = f"""
    <BillQueryRq requestID="24">
      <TxnDateRangeFilter>
        <FromTxnDate>{from_date}</FromTxnDate>
        <ToTxnDate>{to_date}</ToTxnDate>
      </TxnDateRangeFilter>
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
                # Expense lines
                for lb in xml_blocks(b, 'ExpenseLineRet'):
                    lines.append({
                        'line_no':     line_num,
                        'description': clean(xml_val(lb, 'Memo')),
                        'amount':      to_float(xml_val(lb, 'Amount')),
                        'gl_code':     clean(xml_ref(lb, 'AccountRef', 'FullName')),
                    })
                    line_num += 1
                # Item lines
                for lb in xml_blocks(b, 'ItemLineRet'):
                    lines.append({
                        'line_no':     line_num,
                        'description': clean(xml_val(lb, 'Desc')),
                        'amount':      to_float(xml_val(lb, 'Amount')),
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
            err = str(e)
            if 'timeout' in err.lower():
                print(f"    {from_date} to {to_date}: TIMEOUT -- skipping")
            else:
                print(f"    {from_date} to {to_date}: ERROR -- {err[:100]}")

    print(f"    {len(all_bills)} vendor bills found total")
    return all_bills


def extract_bill_payments(session, years_back=3):
    """Pull bill payment checks, chunked by year."""
    if years_back == 0:
        print("  Extracting bill payments (all history, chunked by year)...")
    else:
        print(f"  Extracting bill payments (last {years_back} years)...")

    today = datetime.date.today()
    max_years = 30 if years_back == 0 else years_back

    chunks = []
    for i in range(max_years):
        chunk_end   = today - datetime.timedelta(days=365 * i)
        chunk_start = today - datetime.timedelta(days=365 * (i + 1))
        chunks.append((chunk_start.strftime('%Y-%m-%d'), chunk_end.strftime('%Y-%m-%d')))

    all_payments = []
    for from_date, to_date in chunks:
        request = f"""
    <BillPaymentCheckQueryRq requestID="25">
      <TxnDateRangeFilter>
        <FromTxnDate>{from_date}</FromTxnDate>
        <ToTxnDate>{to_date}</ToTxnDate>
      </TxnDateRangeFilter>
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
                        'ref_no':  clean(xml_val(ab, 'RefNumber')),
                        'amount':  to_float(xml_val(ab, 'Amount')),
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
            err = str(e)
            if 'timeout' in err.lower():
                print(f"    {from_date} to {to_date}: TIMEOUT -- skipping")
            else:
                print(f"    {from_date} to {to_date}: ERROR -- {err[:100]}")

    print(f"    {len(all_payments)} bill payments found total")
    return all_payments


def extract_open_ar(session):
    """Pull open AR balances (unpaid invoices)."""
    print("  Extracting open AR balances...")
    request = """
    <InvoiceQueryRq requestID="10">
      <PaidStatus>NotPaidOnly</PaidStatus>
      <OwnerID>0</OwnerID>
    </InvoiceQueryRq>
    """
    xml = session._send(request)
    blocks = xml_blocks(xml, 'InvoiceRet')
    ar = []
    for b in blocks:
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
# MAIN
# ============================================================================

def main():
    print("=" * 60)
    print("QBExtract — QuickBooks Data Extractor")
    print("=" * 60)
    print()
    print("Make sure QuickBooks is OPEN with your company file loaded.")
    print()

    # How many years of invoice history to pull
    print("How many years of invoice history to export?")
    print("  1 = last 1 year (fastest)")
    print("  3 = last 3 years (recommended for demo)")
    print("  0 = all history (may be slow for large files)")
    try:
        years = int(input("Years [3]: ").strip() or "3")
    except ValueError:
        years = 3
    if years <= 0:
        years = 0  # no date filter — all history

    print()
    print("Connecting to QuickBooks...")

    session = QBSession()
    bundle = {}

    try:
        session.connect()
        print()

        bundle['meta'] = {
            'company':      session.company_name,
            'exported_at':  datetime.datetime.now().isoformat(),
            'years_back':   years,
            'extractor':    'QBExtract v1.0',
        }

        bundle['customers']          = extract_customers(session)
        bundle['items']              = extract_items(session)
        bundle['price_levels']       = extract_price_levels(session)
        bundle['quantity_discounts'] = extract_quantity_discounts(session)
        bundle['vendors']            = extract_vendors(session)
        bundle['sales_reps']         = extract_sales_reps(session)
        bundle['terms']              = extract_terms(session)
        bundle['invoices']           = extract_invoices(session, years)
        bundle['payments']           = extract_payments(session, years)
        bundle['credit_memos']       = extract_credit_memos(session, years)
        bundle['sales_orders']       = extract_sales_orders(session)
        bundle['purchase_orders']    = extract_purchase_orders(session)
        bundle['bills']              = extract_bills(session, years)
        bundle['bill_payments']      = extract_bill_payments(session, years)
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
    safe_name = ''.join(c if c.isalnum() or c in (' ', '-', '_') else '_'
                        for c in session.company_name).strip()
    date_str  = datetime.date.today().strftime('%Y%m%d')
    filename  = f"{safe_name}_export_{date_str}.json"
    filepath  = os.path.join(os.getcwd(), filename)

    print()
    print(f"Writing output file...")
    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(bundle, f, indent=2, ensure_ascii=False)

    size_mb = os.path.getsize(filepath) / (1024 * 1024)

    print()
    print("=" * 60)
    print("DONE")
    print("=" * 60)
    print(f"  Company:        {session.company_name}")
    print(f"  Customers:      {len(bundle['customers'])}")
    print(f"  Items:          {len(bundle['items'])}")
    print(f"  Invoices:       {len(bundle['invoices'])}")
    print(f"  Payments:       {len(bundle.get('payments', []))}")
    print(f"  Credit Memos:   {len(bundle.get('credit_memos', []))}")
    print(f"  Sales Orders:   {len(bundle.get('sales_orders', []))}")
    print(f"  Vendors:        {len(bundle['vendors'])}")
    print(f"  Purchase Orders:{len(bundle.get('purchase_orders', []))}")
    print(f"  Vendor Bills:   {len(bundle.get('bills', []))}")
    print(f"  Bill Payments:  {len(bundle.get('bill_payments', []))}")
    print(f"  File:           {filename}  ({size_mb:.1f} MB)")
    print()
    print("Send this file to your ERP consultant for import.")
    print()
    try:
        input("Press ENTER to exit.")
    except EOFError:
        pass


if __name__ == '__main__':
    main()
