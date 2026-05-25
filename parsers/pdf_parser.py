"""
Parse T-format financial statements from PDF files.
Uses pdfplumber to extract tables and text.
Returns the same structure as xlsx_parser: list of dicts.
"""
import re

try:
    import pdfplumber
except ImportError:
    pdfplumber = None


def _clean_amount(val):
    """Convert a string/number value to a float amount."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s or s in ('-', '—', '–', 'nil', 'Nil', 'NIL', '0'):
        return 0.0
    neg = False
    if s.startswith('(') and s.endswith(')'):
        neg = True
        s = s[1:-1]
    s = s.replace(',', '').replace('₹', '').replace('Rs.', '').replace('Rs', '').strip()
    try:
        v = float(s)
        return -v if neg else v
    except ValueError:
        return 0.0


def _clean_name(val):
    """Clean an account name string."""
    if val is None:
        return ''
    s = str(val).strip()
    s = re.sub(r'^(To|By)\s+', '', s, flags=re.IGNORECASE).strip()
    s = re.sub(r'\s+', ' ', s)
    return s


def _is_skip_name(name):
    """Check if a name should be skipped (totals, headers, etc.)."""
    skip_patterns = [
        r'^total', r'^grand total', r'^net profit', r'^net loss',
        r'^gross profit', r'^gross loss', r'^balance c/d',
        r'^balance b/d', r'^as per', r'^amount',
        r'^particulars', r'^dr\.?$', r'^cr\.?$',
        r'^assets?$', r'^liabilit', r'^equity',
        r'^current assets?$', r'^non.?current', r'^fixed assets?$',
    ]
    lower = name.lower()
    return any(re.match(p, lower) for p in skip_patterns)


def _extract_year_from_text(text):
    """Try to find a financial year pattern in text."""
    m = re.search(r'(\d{4})\s*[-–—]\s*(\d{2,4})', text)
    if m:
        y1 = m.group(1)
        y2 = m.group(2)
        if len(y2) == 2:
            y2 = y1[:2] + y2
        return f'{y1}-{y2[-2:]}'
    m = re.search(r'31\s*(st)?\s*march\s*,?\s*(\d{4})', text, re.IGNORECASE)
    if m:
        y = int(m.group(2))
        return f'{y-1}-{str(y)[-2:]}'
    return ''


def _detect_section(text):
    """Detect which section a text belongs to."""
    t = text.lower().strip()
    trading_kw = ['trading', 'trading account', 'trading and profit', 'trading & profit',
                  'profit and loss', 'profit & loss', 'income and expenditure']
    bs_kw = ['balance sheet', 'balancesheet']
    capital_kw = ['capital account', 'capital a/c']

    for kw in trading_kw:
        if kw in t:
            return 'trading_pl'
    for kw in bs_kw:
        if kw in t:
            return 'balance_sheet'
    for kw in capital_kw:
        if kw in t:
            return 'capital'
    return None


def _is_scanned_pdf(pdf):
    """Check if a PDF appears to be scanned (no extractable text)."""
    total_chars = 0
    pages_checked = min(len(pdf.pages), 3)
    for i in range(pages_checked):
        page = pdf.pages[i]
        text = page.extract_text() or ''
        total_chars += len(text.strip())
    # If very little text found across pages, likely scanned
    return total_chars < 50


def _scan_section_headers(text):
    """
    Scan a block of page text and return a list of (line_index, section)
    tuples for every section header found. Lines are 0-indexed within `text`.
    """
    headers = []
    if not text:
        return headers
    lines = text.split('\n')
    for idx, line in enumerate(lines):
        section = _detect_section(line)
        if section:
            headers.append((idx, section))
    return headers


def _extract_tables_best(page):
    """
    Try multiple table-extraction strategies and return the result with the
    most rows. pdfplumber sometimes finds nothing with default settings but
    succeeds with explicit vertical/horizontal strategies.
    """
    strategies = [
        None,  # default settings
        {"vertical_strategy": "lines", "horizontal_strategy": "lines"},
        {"vertical_strategy": "text", "horizontal_strategy": "text"},
    ]
    best_tables = []
    best_row_count = 0
    for settings in strategies:
        try:
            if settings is None:
                tables = page.extract_tables()
            else:
                tables = page.extract_tables(settings)
        except Exception:
            tables = []
        if not tables:
            continue
        row_count = sum(len(t) for t in tables if t)
        if row_count > best_row_count:
            best_row_count = row_count
            best_tables = tables
    return best_tables


# Indian-format amount: 1,23,456.78 or 12,345.00 or 1234 or (1,234.56)
_AMOUNT_RX = re.compile(
    r'\(?\s*(?:\d{1,3}(?:,\d{2,3})+(?:\.\d+)?|\d+(?:\.\d+)?)\s*\)?'
)


def _split_name_amount(line):
    """
    Given a line like "Sundry Debtors        1,23,456.78", extract
    (name, amount) where amount is the rightmost number on the line.
    Returns (None, 0.0) if no clear amount found.
    """
    if not line:
        return None, 0.0
    matches = list(_AMOUNT_RX.finditer(line))
    if not matches:
        return None, 0.0
    last = matches[-1]
    amt_str = last.group(0).strip()
    # Must be at or near the end of the line
    tail = line[last.end():].strip()
    if tail and not re.match(r'^[\)\.\s]*$', tail):
        return None, 0.0
    name = line[:last.start()].strip()
    if not name:
        return None, 0.0
    amt = _clean_amount(amt_str)
    if amt == 0.0:
        return None, 0.0
    return name, amt


def _parse_t_format_rows(rows):
    """
    Parse rows from a T-format table (side-by-side Dr/Cr or Liabilities/Assets).
    rows is a list of lists (each inner list is a row of cell values).
    Returns (left_items, right_items).
    """
    if not rows or len(rows) < 2:
        return [], []

    num_cols = max(len(r) for r in rows)
    if num_cols < 2:
        return [], []

    mid = num_cols // 2
    left_items = []
    right_items = []

    for row in rows[1:]:  # skip header
        # Pad row to full width
        row = list(row) + [None] * (num_cols - len(row))

        # Left side
        left_name = ''
        left_amt = 0.0
        for c in range(0, mid):
            cell_val = row[c]
            if cell_val is None:
                continue
            text = str(cell_val).strip()
            if not text:
                continue
            amt = _clean_amount(text)
            if amt != 0:
                left_amt = amt
            else:
                cleaned = _clean_name(text)
                if cleaned and not left_name:
                    left_name = cleaned

        if left_name and not _is_skip_name(left_name) and left_amt != 0:
            left_items.append({'name': left_name, 'amount': abs(left_amt)})

        # Right side
        right_name = ''
        right_amt = 0.0
        for c in range(mid, num_cols):
            cell_val = row[c]
            if cell_val is None:
                continue
            text = str(cell_val).strip()
            if not text:
                continue
            amt = _clean_amount(text)
            if amt != 0:
                right_amt = amt
            else:
                cleaned = _clean_name(text)
                if cleaned and not right_name:
                    right_name = cleaned

        if right_name and not _is_skip_name(right_name) and right_amt != 0:
            right_items.append({'name': right_name, 'amount': abs(right_amt)})

    return left_items, right_items


def _parse_capital_rows(rows):
    """Extract capital account details from table rows."""
    cap = {
        'opening': 0, 'capital_introduced': 0, 'net_profit': 0,
        'interest_on_capital': 0, 'drawings': 0, 'closing': 0,
    }
    for row in rows:
        name = ''
        amt = 0.0
        for cell in row:
            if cell is None:
                continue
            text = str(cell).strip()
            if not text:
                continue
            a = _clean_amount(text)
            if a != 0:
                amt = a
            elif text:
                name = text.lower()

        if not name:
            continue
        if 'opening' in name or 'balance b/d' in name:
            cap['opening'] = abs(amt)
        elif 'capital introduced' in name or 'additional capital' in name:
            cap['capital_introduced'] = abs(amt)
        elif 'net profit' in name or 'profit for' in name:
            cap['net_profit'] = abs(amt)
        elif 'interest on capital' in name:
            cap['interest_on_capital'] = abs(amt)
        elif 'drawing' in name or 'withdrawal' in name:
            cap['drawings'] = abs(amt)
        elif 'closing' in name or 'balance c/d' in name:
            cap['closing'] = abs(amt)

    if cap['closing'] == 0 and cap['opening'] > 0:
        cap['closing'] = (cap['opening'] + cap['capital_introduced'] +
                          cap['net_profit'] + cap['interest_on_capital'] - cap['drawings'])
    return cap


def _infer_constitution(data):
    """Infer whether entity is proprietorship or partnership."""
    all_names = []
    for section in ['trading_pl', 'balance_sheet']:
        if section in data and isinstance(data[section], dict):
            for side_key in ['debit', 'credit', 'liabilities', 'assets']:
                if side_key in data[section]:
                    all_names.extend([i['name'].lower() for i in data[section][side_key]])

    entity = data.get('entity_name', '').lower()
    if '& co' in entity or 'and co' in entity:
        return 'partnership'
    for n in all_names:
        if 'partner' in n:
            return 'partnership'
    return 'proprietorship'


def _extract_entity_from_text(text):
    """Try to find entity name from top lines of text."""
    lines = text.split('\n')
    skip = ['trading', 'profit', 'balance', 'capital', 'for the',
            'as at', 'year ended', 'dr', 'cr', 'particulars', 'amount', 'page']
    for line in lines[:10]:
        line = line.strip()
        if not line or len(line) < 4:
            continue
        t_lower = line.lower()
        if any(s in t_lower for s in skip):
            continue
        if re.match(r'^[\d,.\s]+$', line):
            continue
        return line
    return 'Entity Name'


def _assign_table_section(table, page_text, page_headers, used_sections):
    """
    Decide which section (trading_pl / balance_sheet / capital / None)
    a given table belongs to. First checks the table's own header row,
    then falls back to the nearest preceding section header in the page
    text. Returns the detected section string or None.
    """
    # 1. Check the table's own header row
    if table:
        header_text = ' '.join(str(cell or '') for cell in table[0])
        section = _detect_section(header_text)
        if section:
            return section

    # 2. Locate this table's first non-empty cell text in the page text,
    #    then pick the most recent header above it. If we can't find the
    #    table text, just use the LAST header on the page that hasn't been
    #    used yet (typical T-format pages have one header per table).
    if not page_headers:
        return None

    if table and len(table) >= 2:
        # Try to find any cell value of the first data row in the page text
        probe_row = table[1] if len(table) > 1 else table[0]
        probe_texts = [str(c).strip() for c in (probe_row or []) if c and str(c).strip()]
        for probe in probe_texts:
            # Trim long probes to avoid false negatives
            snippet = probe[:30]
            idx = page_text.find(snippet)
            if idx >= 0:
                # Convert char offset to line index
                line_idx = page_text[:idx].count('\n')
                # Pick the most recent header at or above this line
                candidate = None
                for h_line, h_sec in page_headers:
                    if h_line <= line_idx:
                        candidate = h_sec
                if candidate:
                    return candidate
                break

    # Fallback: first unused header on this page
    for _, h_sec in page_headers:
        if h_sec not in used_sections:
            return h_sec

    # Final fallback: first header on the page
    return page_headers[0][1]


def parse_pdf(filepath):
    """
    Parse a PDF file containing T-format financial statements.
    Returns a list of parsed data dicts (same structure as xlsx_parser).
    """
    if pdfplumber is None:
        return []

    try:
        pdf = pdfplumber.open(filepath)
    except Exception:
        return []

    try:
        # Check for scanned PDF
        if _is_scanned_pdf(pdf):
            pdf.close()
            raise ValueError(
                'Scanned PDFs are not supported. Please use a digital/text-based PDF '
                'or convert to Excel/Word format.'
            )

        # Extract full text from first page for entity name and year
        first_page_text = pdf.pages[0].extract_text() or '' if pdf.pages else ''
        entity_name = _extract_entity_from_text(first_page_text)
        year = ''
        for page in pdf.pages[:3]:
            page_text = page.extract_text() or ''
            y = _extract_year_from_text(page_text)
            if y:
                year = y
                break

        data = {
            'entity_name': entity_name,
            'fy': year,
            'source_sheet': 'pdf',
        }

        # Walk pages: detect section headers from text first, then associate
        # tables (extracted with the best of several strategies) to them.
        used_sections = set()
        for page in pdf.pages:
            page_text = page.extract_text() or ''
            page_headers = _scan_section_headers(page_text)

            tables = _extract_tables_best(page)
            valid_tables = [t for t in tables if t and len(t) >= 2]

            for table in valid_tables:
                section = _assign_table_section(
                    table, page_text, page_headers, used_sections
                )

                if section == 'trading_pl' and 'trading_pl' not in data:
                    debit, credit = _parse_t_format_rows(table)
                    if debit or credit:
                        data['trading_pl'] = {'debit': debit, 'credit': credit}
                        used_sections.add('trading_pl')
                        continue

                if section == 'balance_sheet' and 'balance_sheet' not in data:
                    liab, assets = _parse_t_format_rows(table)
                    if liab or assets:
                        data['balance_sheet'] = {'liabilities': liab, 'assets': assets}
                        used_sections.add('balance_sheet')
                        continue

                if section == 'capital' and 'capital_account' not in data:
                    data['capital_account'] = _parse_capital_rows(table)
                    used_sections.add('capital')
                    continue

                # No clear section — guess from content keywords
                left, right = _parse_t_format_rows(table)
                if not (left or right):
                    continue
                all_names = [i['name'].lower() for i in left + right]
                joined = ' '.join(all_names)
                if 'trading_pl' not in data and any(
                    kw in joined for kw in
                    ['sales', 'purchase', 'salary', 'rent', 'stock', 'depreciation']
                ):
                    data['trading_pl'] = {'debit': left, 'credit': right}
                    used_sections.add('trading_pl')
                    continue
                if 'balance_sheet' not in data and any(
                    kw in joined for kw in
                    ['capital', 'bank', 'cash', 'debtor', 'creditor', 'loan', 'sundry']
                ):
                    data['balance_sheet'] = {'liabilities': left, 'assets': right}
                    used_sections.add('balance_sheet')

        # If no tables found, try improved text-line extraction
        if not any(k in data for k in ['trading_pl', 'balance_sheet', 'capital_account']):
            current_section = None
            trading_left, trading_right = [], []
            bs_left, bs_right = [], []
            raw_items = []

            for page in pdf.pages:
                text = page.extract_text() or ''
                for line in text.split('\n'):
                    line = line.rstrip()
                    if not line.strip():
                        continue

                    detected = _detect_section(line)
                    if detected:
                        current_section = detected
                        continue

                    name, amt = _split_name_amount(line)
                    if not name:
                        # Fallback to the older split approach
                        parts = re.split(r'\s{2,}|\t+', line.strip())
                        if len(parts) >= 2:
                            name = _clean_name(parts[0])
                            amt = _clean_amount(parts[-1])
                    else:
                        name = _clean_name(name)

                    if not name or amt == 0 or _is_skip_name(name):
                        continue

                    entry = {'name': name, 'amount': abs(amt)}
                    if current_section == 'trading_pl':
                        trading_left.append(entry)
                    elif current_section == 'balance_sheet':
                        bs_left.append(entry)
                    else:
                        raw_items.append(entry)

            if trading_left or trading_right:
                data['trading_pl'] = {'debit': trading_left, 'credit': trading_right}
            if bs_left or bs_right:
                data['balance_sheet'] = {'liabilities': bs_left, 'assets': bs_right}
            if not any(k in data for k in ['trading_pl', 'balance_sheet']) and raw_items:
                data['raw_items'] = raw_items

        pdf.close()

        # Only return if we found financial data
        if any(k in data for k in ['trading_pl', 'balance_sheet', 'capital_account', 'raw_items']):
            data['constitution'] = _infer_constitution(data)
            data['proprietor_name'] = entity_name if data['constitution'] == 'proprietorship' else ''
            return [data]

        return []

    except ValueError:
        # Re-raise ValueError (scanned PDF message) to caller
        pdf.close()
        raise
    except Exception:
        pdf.close()
        return []
