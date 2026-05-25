"""
Parse T-format financial statements from PDF files.

Uses pdfplumber's word-level extraction (with x/y coordinates) to reconstruct
two-column T-format layouts (Trading & P&L, Balance Sheet, Capital Account)
that Tally and similar Indian accounting software produce as PDF.

Strategy:
  1. Extract all words from each page with x0/top coordinates.
  2. Group words into rows by `top` (allowing small vertical tolerance).
  3. For each row, split words into LEFT and RIGHT columns at the page midpoint.
  4. Reconstruct line text per column. Track the current section (trading_pl /
     balance_sheet / capital) by scanning rows for section header keywords.
  5. Within each section, parse each column's lines: lines with an amount become
     items, lines without an amount are skipped as category headers
     ("TO PURCHASES", "FIXED ASSETS", etc.).

Returns the same structure as xlsx_parser: list of dicts.
"""
import re

try:
    import pdfplumber
except ImportError:
    pdfplumber = None


# ---------- Helpers ----------

# Indian / Western number formats: 1,23,456.78 / 1,234.56 / 38,280 / (1,234.56)
_AMOUNT_RE = re.compile(r'\(?-?[\d,]+(?:\.\d{1,2})?\)?')

# Stricter check: must contain at least one digit and either a comma or decimal,
# OR be a multi-digit number. Used to filter out random small numbers in names.
_AMOUNT_RE_STRICT = re.compile(
    r'\(?-?(?:\d{1,3}(?:,\d{2,3})+(?:\.\d{1,2})?|\d{2,}(?:\.\d{1,2})?|\d+\.\d{2})\)?'
)


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


# Lines that are pure totals / grand totals / column headers — never items.
_SKIP_LINE_RE = re.compile(
    r'^(total|grand\s+total|sub\s*total|particulars|amount|dr\.?|cr\.?|'
    r'liabilities|assets?|equity|liabilities\s+and|liability|amounts?)$',
    re.IGNORECASE,
)

# Category-only headers (no amount on the same line) that should be skipped.
# Items appear on subsequent lines.
_CATEGORY_HEADER_RE = re.compile(
    r'^(fixed\s+assets?|current\s+assets?|non.?current\s+assets?|'
    r'investments?|loans?\s+(and|&)\s+advances?|cash\s+and\s+bank|'
    r'other\s+current\s+assets?|other\s+non.?current\s+assets?|'
    r'trade\s+receivables?|trade\s+payables?|sundry\s+(debtors?|creditors?)|'
    r'capital\s+account|reserves?\s+and\s+surplus|'
    r'secured\s+loans?|unsecured\s+loans?|borrowings?|'
    r'long.?term\s+(borrowings?|liabilities)|'
    r'short.?term\s+(borrowings?|liabilities|provisions?)|'
    r'current\s+liabilities|other\s+current\s+liabilities|'
    r'(by|to)\s+(opening|purchase|sales|direct|indirect|gross\s+profit\s*a/c|'
    r'closing|net\s+profit\s*a/c|trading|expenses?)\s*(a/c|account)?)$',
    re.IGNORECASE,
)


def _extract_year_from_text(text):
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
    """Detect which section a header line belongs to."""
    t = text.lower().strip()
    if not t:
        return None
    if 'trading' in t and ('profit' in t or 'loss' in t):
        return 'trading_pl'
    if 'profit and loss' in t or 'profit & loss' in t:
        return 'trading_pl'
    if 'income and expenditure' in t:
        return 'trading_pl'
    if 'trading account' in t:
        return 'trading_pl'
    if 'balance sheet' in t or 'balancesheet' in t:
        return 'balance_sheet'
    if 'capital account' in t or 'capital a/c' in t:
        return 'capital'
    return None


def _is_scanned_pdf(pdf):
    total_chars = 0
    for i in range(min(len(pdf.pages), 3)):
        text = pdf.pages[i].extract_text() or ''
        total_chars += len(text.strip())
    return total_chars < 50


# ---------- Word-coordinate based row reconstruction ----------

def _group_words_into_rows(words, y_tol=3):
    """Group words by `top` coordinate. Returns list of rows; each row is a
    sorted-by-x list of word dicts."""
    if not words:
        return []
    # Sort by top, then x0
    sorted_words = sorted(words, key=lambda w: (w['top'], w['x0']))
    rows = []
    current_row = [sorted_words[0]]
    current_top = sorted_words[0]['top']
    for w in sorted_words[1:]:
        if abs(w['top'] - current_top) <= y_tol:
            current_row.append(w)
        else:
            rows.append(sorted(current_row, key=lambda x: x['x0']))
            current_row = [w]
            current_top = w['top']
    rows.append(sorted(current_row, key=lambda x: x['x0']))
    return rows


def _find_column_split(rows, page_width):
    """Find the x-coordinate that best splits left/right columns.

    Look at the AMOUNT column for the left side — usually around x = 0.45*width.
    Then anything beyond that until the next text run is the right column.
    Default to page midpoint if heuristic fails.
    """
    # Look for a header row containing two "PARTICULARS" / "AMOUNT" / etc.
    for row in rows[:20]:
        texts_lower = [w['text'].lower() for w in row]
        # Find two occurrences of a header word
        positions = []
        for i, t in enumerate(texts_lower):
            if t in ('particulars', 'liabilities', 'assets'):
                positions.append(row[i]['x0'])
        if len(positions) >= 2:
            # The right column starts at the second occurrence
            return positions[1] - 5
    return page_width / 2


def _row_text_in_column(row, x_min, x_max):
    """Return the joined text of words in [x_min, x_max), preserving order."""
    parts = [w['text'] for w in row if x_min <= w['x0'] < x_max]
    return ' '.join(parts).strip()


def _parse_line_to_item(line):
    """
    Split a single column-line into (name, amount).
    Examples:
      'OPENING STOCK 38,280.00'   -> ('OPENING STOCK', 38280.0)
      'TO GROSS PROFIT 7,19,541.00' -> ('GROSS PROFIT', 719541.0)
      'CAPITAL 10,08,271.00'      -> ('CAPITAL', 1008271.0)
      'TO OPENING STOCK'          -> ('TO OPENING STOCK', None)  # header, skip
      '26,86,336.00'              -> (None, 2686336.0)  # standalone amount
    Returns (name_or_None, amount_or_None).
    """
    s = line.strip()
    if not s:
        return None, None

    # Find all amount-looking tokens
    matches = list(_AMOUNT_RE_STRICT.finditer(s))
    if not matches:
        return s, None

    # The LAST match is the amount; everything before it is the name
    last = matches[-1]
    amt_str = last.group(0)
    amount = _clean_amount(amt_str)
    if amount == 0:
        # Probably wasn't really an amount
        return s, None

    name = s[:last.start()].strip()
    if not name:
        return None, amount
    return name, amount


def _filter_items(lines):
    """Convert a list of column-lines into a list of {name, amount} items.

    Skips: pure category headers (no amount), standalone totals, and column
    header rows. A line with an amount is always treated as an item — even
    if its name happens to match a category-header pattern (e.g. when there
    is exactly one item under that category, like 'OTHER CURRENT ASSET 84,115').
    """
    items = []
    for line in lines:
        name, amount = _parse_line_to_item(line)
        if amount is None:
            # Pure header line — skip
            continue
        if name is None:
            # Standalone amount (sub-total). Skip.
            continue
        cleaned = _clean_name(name)
        if not cleaned:
            continue
        if _SKIP_LINE_RE.match(cleaned):
            continue
        # NOTE: don't skip on _CATEGORY_HEADER_RE here — that pattern is meant
        # for header lines without amounts, which are already filtered above.
        items.append({'name': cleaned, 'amount': abs(amount)})
    return items


def _parse_capital_from_lines(left_lines, right_lines):
    """Capital account has Withdrawals/Closing on left and Opening/Profit/etc on right."""
    cap = {
        'opening': 0, 'capital_introduced': 0, 'net_profit': 0,
        'interest_on_capital': 0, 'drawings': 0, 'closing': 0,
    }
    for lines in (left_lines, right_lines):
        for line in lines:
            name, amt = _parse_line_to_item(line)
            if amt is None or name is None:
                continue
            n = _clean_name(name).lower()
            if 'opening' in n or 'balance b/d' in n:
                cap['opening'] = abs(amt)
            elif 'capital introduced' in n or 'additional capital' in n or 'add capital' in n:
                cap['capital_introduced'] = abs(amt)
            elif 'net profit' in n or 'profit for' in n or 'share of profit' in n:
                cap['net_profit'] = abs(amt)
            elif 'interest on capital' in n:
                cap['interest_on_capital'] = abs(amt)
            elif 'drawing' in n or 'withdrawal' in n:
                cap['drawings'] = abs(amt)
            elif 'closing' in n or 'balance c/d' in n:
                cap['closing'] = abs(amt)
            elif 'bank interest' in n or 'interest received' in n:
                # Bank interest credited to capital — track as extra
                cap.setdefault('bank_interest', 0)
                cap['bank_interest'] = abs(amt)

    if cap['closing'] == 0 and cap['opening'] > 0:
        cap['closing'] = (cap['opening'] + cap['capital_introduced'] +
                          cap['net_profit'] + cap['interest_on_capital']
                          + cap.get('bank_interest', 0) - cap['drawings'])
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
    for line in lines[:6]:
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


# ---------- Main parser ----------

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
        if _is_scanned_pdf(pdf):
            pdf.close()
            raise ValueError(
                'Scanned PDFs are not supported. Please use a digital/text-based PDF '
                'or convert to Excel/Word format.'
            )

        # Header / entity / year metadata
        first_text = pdf.pages[0].extract_text() or '' if pdf.pages else ''
        entity_name = _extract_entity_from_text(first_text)
        year = ''
        for page in pdf.pages[:3]:
            t = page.extract_text() or ''
            y = _extract_year_from_text(t)
            if y:
                year = y
                break

        data = {
            'entity_name': entity_name,
            'fy': year,
            'source_sheet': 'pdf',
        }

        # Accumulators — multiple sections of same kind from different pages get merged
        trading_left_lines = []
        trading_right_lines = []
        bs_left_lines = []
        bs_right_lines = []
        capital_left_lines = []
        capital_right_lines = []

        for page in pdf.pages:
            words = page.extract_words()
            if not words:
                continue
            rows = _group_words_into_rows(words)
            split_x = _find_column_split(rows, page.width)

            current_section = None
            for row in rows:
                row_text = ' '.join(w['text'] for w in row).strip()
                # Section header detection — full row text
                sec = _detect_section(row_text)
                if sec:
                    current_section = sec
                    continue
                if not current_section:
                    continue

                # Skip the column-header row ("PARTICULARS AMOUNT PARTICULARS AMOUNT")
                if re.fullmatch(r'(particulars|liabilities|assets|amount|\s)+',
                                row_text, re.IGNORECASE):
                    continue

                # Get left & right column text
                left_text = _row_text_in_column(row, 0, split_x)
                right_text = _row_text_in_column(row, split_x, page.width + 1)

                # Skip TOTAL lines (e.g. "TOTAL 7,19,541.00 TOTAL 7,19,541.00")
                if re.match(r'^(total|grand\s+total)\b', left_text, re.IGNORECASE):
                    continue

                if current_section == 'trading_pl':
                    if left_text:
                        trading_left_lines.append(left_text)
                    if right_text:
                        trading_right_lines.append(right_text)
                elif current_section == 'balance_sheet':
                    if left_text:
                        bs_left_lines.append(left_text)
                    if right_text:
                        bs_right_lines.append(right_text)
                elif current_section == 'capital':
                    if left_text:
                        capital_left_lines.append(left_text)
                    if right_text:
                        capital_right_lines.append(right_text)

        # Build sections
        if trading_left_lines or trading_right_lines:
            debit = _filter_items(trading_left_lines)
            credit = _filter_items(trading_right_lines)
            if debit or credit:
                data['trading_pl'] = {'debit': debit, 'credit': credit}

        if bs_left_lines or bs_right_lines:
            liab = _filter_items(bs_left_lines)
            assets = _filter_items(bs_right_lines)
            if liab or assets:
                data['balance_sheet'] = {'liabilities': liab, 'assets': assets}

        if capital_left_lines or capital_right_lines:
            cap = _parse_capital_from_lines(capital_left_lines, capital_right_lines)
            if any(v for v in cap.values()):
                data['capital_account'] = cap

        pdf.close()

        if any(k in data for k in ['trading_pl', 'balance_sheet', 'capital_account']):
            data['constitution'] = _infer_constitution(data)
            data['proprietor_name'] = entity_name if data['constitution'] == 'proprietorship' else ''
            return [data]

        return []

    except ValueError:
        try:
            pdf.close()
        except Exception:
            pass
        raise
    except Exception:
        try:
            pdf.close()
        except Exception:
            pass
        return []
