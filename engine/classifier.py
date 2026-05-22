"""
Classify parsed T-format line items into ICAI NCE Schedule III categories.
"""
import re
from config import ACCOUNT_KEYWORDS


def _match_category(name, keywords_map):
    name_lower = name.lower().strip()
    name_lower = re.sub(r'\s+', ' ', name_lower)

    best_match = None
    best_len = 0
    for category, keywords in keywords_map.items():
        for kw in keywords:
            kw_lower = kw.lower()
            if kw_lower == name_lower:
                return category
            if kw_lower in name_lower and len(kw_lower) > best_len:
                best_match = category
                best_len = len(kw_lower)

    return best_match


def _classify_item(name, amount, side, section):
    cat = _match_category(name, ACCOUNT_KEYWORDS)

    if cat:
        return cat, name, abs(amount)

    name_lower = name.lower()

    if section == 'trading_pl':
        if side == 'debit':
            if any(w in name_lower for w in ['stock', 'inventory']):
                if 'opening' in name_lower:
                    return 'opening_stock', name, abs(amount)
                elif 'closing' in name_lower:
                    return 'closing_stock', name, abs(amount)
            if any(w in name_lower for w in ['purchase', 'buy']):
                return 'purchases', name, abs(amount)
            if any(w in name_lower for w in ['wage', 'labour', 'salary', 'staff']):
                return 'employee_benefits', name, abs(amount)
            if any(w in name_lower for w in ['interest', 'bank charge', 'finance']):
                return 'finance_costs', name, abs(amount)
            if 'depreciation' in name_lower or 'dep.' in name_lower:
                return 'depreciation', name, abs(amount)
            return 'other_expenses', name, abs(amount)
        else:
            if any(w in name_lower for w in ['sale', 'revenue', 'turnover']):
                return 'revenue_from_operations', name, abs(amount)
            if 'closing stock' in name_lower:
                return 'closing_stock', name, abs(amount)
            return 'other_income', name, abs(amount)

    if section == 'balance_sheet':
        if side in ('liabilities', 'debit'):
            if any(w in name_lower for w in ['capital', 'proprietor']):
                return 'capital', name, abs(amount)
            if any(w in name_lower for w in ['creditor', 'payable', 'trade payable']):
                return 'trade_payables', name, abs(amount)
            if any(w in name_lower for w in ['loan', 'borrow']):
                return 'bank_loan', name, abs(amount)
            if any(w in name_lower for w in ['gst', 'tds', 'tax', 'provision']):
                return 'other_statutory', name, abs(amount)
            return 'trade_payables', name, abs(amount)
        else:
            if any(w in name_lower for w in ['cash', 'petty cash']):
                return 'cash_hand', name, abs(amount)
            if any(w in name_lower for w in ['bank', 'sbi', 'hdfc', 'icici', 'axis', 'bob', 'pnb']):
                return 'bank_balance', name, abs(amount)
            if any(w in name_lower for w in ['debtor', 'receivable']):
                return 'trade_receivables', name, abs(amount)
            if any(w in name_lower for w in ['stock', 'inventory']):
                return 'closing_stock', name, abs(amount)
            if any(w in name_lower for w in ['gold', 'jewel', 'ornament', 'bullion']):
                return 'gold_investments', name, abs(amount)
            if any(w in name_lower for w in ['deposit', 'security']):
                return 'security_deposit', name, abs(amount)
            if any(w in name_lower for w in ['fixed deposit', ' fd', 'fd ']):
                return 'fixed_deposit', name, abs(amount)
            if any(w in name_lower for w in ['gst input', 'input', 'itc']):
                return 'gst_input', name, abs(amount)
            if any(w in name_lower for w in ['tds', 'advance tax']):
                return 'tds_advance_tax', name, abs(amount)
            for ppe_kw in ACCOUNT_KEYWORDS.get('ppe', []):
                if ppe_kw.lower() in name_lower:
                    return 'ppe', name, abs(amount)
            return 'other_ca', name, abs(amount)

    return 'other_expenses', name, abs(amount)


def classify_data(cy_parsed, py_parsed):
    classified = {
        'entity_name': cy_parsed.get('entity_name', 'Entity Name'),
        'fy_cy': cy_parsed.get('fy', ''),
        'fy_py': py_parsed.get('fy', ''),
        'constitution': cy_parsed.get('constitution', 'proprietorship'),
        'proprietor_name': cy_parsed.get('proprietor_name', ''),

        'capital_cy': cy_parsed.get('capital_account', {}),
        'capital_py': py_parsed.get('capital_account', {}),

        'revenue_from_operations_cy': [],
        'revenue_from_operations_py': [],
        'other_income_cy': [],
        'other_income_py': [],
        'purchases_cy': [],
        'purchases_py': [],
        'opening_stock_cy': 0,
        'opening_stock_py': 0,
        'closing_stock_cy': 0,
        'closing_stock_py': 0,
        'employee_benefits_cy': [],
        'employee_benefits_py': [],
        'finance_costs_cy': [],
        'finance_costs_py': [],
        'depreciation_cy': 0,
        'depreciation_py': 0,
        'other_expenses_cy': [],
        'other_expenses_py': [],

        'ppe_cy': [],
        'ppe_py': [],
        'gold_investments_cy': [],
        'gold_investments_py': [],
        'trade_receivables_cy': [],
        'trade_receivables_py': [],
        'inventories_cy': 0,
        'inventories_py': 0,
        'cash_hand_cy': [],
        'cash_hand_py': [],
        'bank_balance_cy': [],
        'bank_balance_py': [],
        'fixed_deposit_cy': [],
        'fixed_deposit_py': [],
        'security_deposit_cy': [],
        'security_deposit_py': [],
        'gst_input_cy': [],
        'gst_input_py': [],
        'tds_advance_tax_cy': [],
        'tds_advance_tax_py': [],
        'other_ca_cy': [],
        'other_ca_py': [],

        'trade_payables_cy': [],
        'trade_payables_py': [],
        'bank_loan_cy': [],
        'bank_loan_py': [],
        'bank_od_cc_cy': [],
        'bank_od_cc_py': [],
        'unsecured_loan_cy': [],
        'unsecured_loan_py': [],
        'gst_payable_cy': [],
        'gst_payable_py': [],
        'tds_payable_cy': [],
        'tds_payable_py': [],
        'other_statutory_cy': [],
        'other_statutory_py': [],
        'provision_tax_cy': [],
        'provision_tax_py': [],

        'net_profit_cy': 0,
        'net_profit_py': 0,
    }

    for suffix, parsed in [('cy', cy_parsed), ('py', py_parsed)]:
        if 'trading_pl' in parsed:
            tp = parsed['trading_pl']
            for item in tp.get('debit', []):
                if item.get('is_total'):
                    continue
                cat, name, amt = _classify_item(item['name'], item['amount'], 'debit', 'trading_pl')
                _add_to_classified(classified, cat, suffix, name, amt)

            for item in tp.get('credit', []):
                if item.get('is_total'):
                    continue
                cat, name, amt = _classify_item(item['name'], item['amount'], 'credit', 'trading_pl')
                _add_to_classified(classified, cat, suffix, name, amt)

        if 'balance_sheet' in parsed:
            bs = parsed['balance_sheet']
            for item in bs.get('liabilities', []):
                if item.get('is_total'):
                    continue
                cat, name, amt = _classify_item(item['name'], item['amount'], 'liabilities', 'balance_sheet')
                if cat == 'capital':
                    cap = classified.get(f'capital_{suffix}', {})
                    if not cap.get('closing'):
                        cap['closing'] = amt
                    classified[f'capital_{suffix}'] = cap
                else:
                    _add_to_classified(classified, cat, suffix, name, amt)

            for item in bs.get('assets', []):
                if item.get('is_total'):
                    continue
                cat, name, amt = _classify_item(item['name'], item['amount'], 'assets', 'balance_sheet')
                _add_to_classified(classified, cat, suffix, name, amt)

    for suffix in ['cy', 'py']:
        tp = (cy_parsed if suffix == 'cy' else py_parsed).get('trading_pl', {})
        total_debit = sum(i['amount'] for i in tp.get('debit', []) if not i.get('is_total'))
        total_credit = sum(i['amount'] for i in tp.get('credit', []) if not i.get('is_total'))
        gp_items = [i for i in tp.get('debit', []) if i.get('is_total') and 'gross profit' in i['name'].lower()]
        np_items = [i for i in tp.get('debit', []) if i.get('is_total') and 'net profit' in i['name'].lower()]
        if not np_items:
            np_items = [i for i in tp.get('credit', []) if i.get('is_total') and 'net profit' in i['name'].lower()]

        if np_items:
            classified[f'net_profit_{suffix}'] = np_items[0]['amount']
        elif total_credit > total_debit:
            classified[f'net_profit_{suffix}'] = total_credit - total_debit

    classified['inventories_cy'] = classified.get('closing_stock_cy', 0)
    classified['inventories_py'] = classified.get('closing_stock_py', 0)

    _derive_year_info(classified)

    return classified


def _add_to_classified(classified, category, suffix, name, amount):
    key = f'{category}_{suffix}'

    scalar_keys = ['opening_stock', 'closing_stock', 'depreciation']
    for sk in scalar_keys:
        if category == sk:
            classified[f'{sk}_{suffix}'] = classified.get(f'{sk}_{suffix}', 0) + amount
            return

    if key in classified:
        if isinstance(classified[key], list):
            existing = next((i for i in classified[key] if i['name'].lower() == name.lower()), None)
            if existing:
                existing['amount'] += amount
            else:
                classified[key].append({'name': name, 'amount': amount})
        else:
            classified[key] = classified.get(key, 0) + amount
    else:
        classified[key] = [{'name': name, 'amount': amount}]


def _derive_year_info(classified):
    fy_cy = classified.get('fy_cy', '')
    fy_py = classified.get('fy_py', '')

    if fy_cy:
        m = re.match(r'(\d{4})-(\d{2,4})', fy_cy)
        if m:
            y2 = m.group(2)
            if len(y2) == 2:
                y2 = m.group(1)[:2] + y2
            classified['year_ending_cy'] = f'31 March {y2}'
            classified['year_start_cy'] = f'1 April {m.group(1)}'
        else:
            classified['year_ending_cy'] = ''
            classified['year_start_cy'] = ''
    else:
        classified['year_ending_cy'] = ''
        classified['year_start_cy'] = ''

    if fy_py:
        m = re.match(r'(\d{4})-(\d{2,4})', fy_py)
        if m:
            y2 = m.group(2)
            if len(y2) == 2:
                y2 = m.group(1)[:2] + y2
            classified['year_ending_py'] = f'31 March {y2}'
            classified['year_start_py'] = f'1 April {m.group(1)}'
        else:
            classified['year_ending_py'] = ''
            classified['year_start_py'] = ''
    else:
        classified['year_ending_py'] = ''
        classified['year_start_py'] = ''


def _sum_list(items):
    if isinstance(items, list):
        return sum(i.get('amount', 0) for i in items)
    return items if isinstance(items, (int, float)) else 0
