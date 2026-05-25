"""
Note suppression and dynamic numbering for ICAI NCE financial statements.
Only retains notes backed by data; renumbers sequentially.
"""


def _sum_items(items):
    if isinstance(items, list):
        return sum(i.get('amount', 0) for i in items)
    if isinstance(items, (int, float)):
        return items
    return 0


def _has_data(classified, *keys):
    for k in keys:
        val = classified.get(k)
        if val:
            if isinstance(val, list) and len(val) > 0:
                if _sum_items(val) != 0:
                    return True
            elif isinstance(val, (int, float)) and val != 0:
                return True
            elif isinstance(val, dict) and any(v for v in val.values() if v):
                return True
    return False


def _get_note_definitions(constitution='proprietorship'):
    """
    Build note definitions dynamically based on constitution.
    Partnership firms get different capital title and always-retained ST Provisions.
    """
    is_partnership = (constitution == 'partnership')

    defs = [
        {
            'id': 'brief',
            'title': 'Brief about the Entity',
            'always_retain': True,
            'bs_line': None,
            'pl_line': None,
        },
        {
            'id': 'policies',
            'title': 'Significant Accounting Policies',
            'always_retain': True,
            'bs_line': None,
            'pl_line': None,
        },
        {
            'id': 'capital',
            'title': "Partners' Capital Account" if is_partnership else "Owner's Capital Account",
            'always_retain': True,
            'bs_line': 'capital',
            'pl_line': None,
        },
        {
            'id': 'lt_borrowings',
            'title': 'Long-Term Borrowings',
            'always_retain': False,
            'data_keys': ['bank_loan_cy', 'bank_loan_py'],
            'bs_line': 'lt_borrowings',
        },
        {
            'id': 'st_borrowings',
            'title': 'Short-Term Borrowings',
            'always_retain': False,
            'data_keys': ['bank_od_cc_cy', 'bank_od_cc_py', 'unsecured_loan_cy', 'unsecured_loan_py'],
            'bs_line': 'st_borrowings',
        },
        {
            'id': 'trade_payables',
            'title': 'Trade Payables',
            'always_retain': False,
            'data_keys': ['trade_payables_cy', 'trade_payables_py'],
            'bs_line': 'trade_payables',
        },
        {
            'id': 'other_cl',
            'title': 'Other Current Liabilities',
            'always_retain': False,
            'data_keys': ['gst_payable_cy', 'gst_payable_py', 'tds_payable_cy', 'tds_payable_py',
                          'other_statutory_cy', 'other_statutory_py'],
            'bs_line': 'other_cl',
        },
        {
            'id': 'st_provisions',
            'title': 'Short-Term Provisions',
            # Partnership always retains (firm tax provision); Proprietorship suppresses if empty
            'always_retain': True if is_partnership else False,
            'data_keys': ['provision_tax_cy', 'provision_tax_py'],
            'bs_line': 'st_provisions',
        },
        {
            'id': 'ppe',
            'title': 'Property, Plant and Equipment',
            'always_retain': False,
            'data_keys': ['ppe_cy', 'ppe_py'],
            'bs_line': 'ppe',
        },
        {
            'id': 'non_current_investments',
            'title': 'Non-Current Investments',
            'always_retain': False,
            'data_keys': ['gold_investments_cy', 'gold_investments_py'],
            'bs_line': 'non_current_investments',
        },
        {
            'id': 'lt_loans_advances',
            'title': 'Long-Term Loans and Advances',
            'always_retain': False,
            'data_keys': ['security_deposit_cy', 'security_deposit_py'],
            'bs_line': 'lt_loans_advances',
        },
        {
            'id': 'inventories',
            'title': 'Inventories',
            'always_retain': False,
            'data_keys': ['inventories_cy', 'inventories_py', 'closing_stock_cy', 'closing_stock_py'],
            'bs_line': 'inventories',
        },
        {
            'id': 'trade_receivables',
            'title': 'Trade Receivables',
            'always_retain': False,
            'data_keys': ['trade_receivables_cy', 'trade_receivables_py'],
            'bs_line': 'trade_receivables',
        },
        {
            'id': 'cash_bank',
            'title': 'Cash and Bank Balances',
            'always_retain': False,
            'data_keys': ['cash_hand_cy', 'cash_hand_py', 'bank_balance_cy', 'bank_balance_py',
                          'fixed_deposit_cy', 'fixed_deposit_py'],
            'bs_line': 'cash_bank',
        },
        {
            'id': 'st_loans_advances',
            'title': 'Short-Term Loans and Advances',
            'always_retain': False,
            'data_keys': ['gst_input_cy', 'gst_input_py', 'tds_advance_tax_cy', 'tds_advance_tax_py'],
            'bs_line': 'st_loans_advances',
        },
        {
            'id': 'other_ca',
            'title': 'Other Current Assets',
            'always_retain': False,
            'data_keys': ['other_ca_cy', 'other_ca_py'],
            'bs_line': 'other_ca',
        },
        {
            'id': 'revenue',
            'title': 'Revenue from Operations',
            'always_retain': False,
            'data_keys': ['revenue_from_operations_cy', 'revenue_from_operations_py'],
            'pl_line': 'revenue',
        },
        {
            'id': 'other_income',
            'title': 'Other Income',
            'always_retain': False,
            'data_keys': ['other_income_cy', 'other_income_py'],
            'pl_line': 'other_income',
        },
        {
            'id': 'cost_of_materials',
            'title': 'Purchases of Stock-in-Trade',
            'always_retain': False,
            'data_keys': ['purchases_cy', 'purchases_py'],
            'pl_line': 'cost_of_materials',
        },
        {
            'id': 'changes_in_inventories',
            'title': 'Changes in Inventories of Finished Goods and Stock-in-Trade',
            'always_retain': False,
            'data_keys': ['opening_stock_cy', 'opening_stock_py', 'closing_stock_cy', 'closing_stock_py'],
            'pl_line': 'changes_in_inventories',
        },
        {
            'id': 'employee_benefits',
            'title': 'Employee Benefits Expense',
            'always_retain': False,
            'data_keys': ['employee_benefits_cy', 'employee_benefits_py'],
            'pl_line': 'employee_benefits',
        },
        {
            'id': 'finance_costs',
            'title': 'Finance Costs',
            'always_retain': False,
            'data_keys': ['finance_costs_cy', 'finance_costs_py'],
            'pl_line': 'finance_costs',
        },
        {
            'id': 'depreciation_note',
            'title': 'Depreciation and Amortisation Expense',
            'always_retain': False,
            'data_keys': ['depreciation_cy', 'depreciation_py', 'ppe_cy', 'ppe_py'],
            'pl_line': 'depreciation_note',
        },
        {
            'id': 'other_expenses_note',
            'title': 'Other Expenses',
            'always_retain': False,
            'data_keys': ['other_expenses_cy', 'other_expenses_py'],
            'pl_line': 'other_expenses_note',
        },
        {
            'id': 'segment_reporting',
            'title': 'Segment Reporting',
            'always_retain': True,
            'bs_line': None,
            'pl_line': None,
        },
        {
            'id': 'confirmation_balances',
            'title': 'Confirmation of Balances',
            'always_retain': False,
            'data_keys': ['trade_receivables_cy', 'trade_receivables_py',
                          'trade_payables_cy', 'trade_payables_py'],
            'bs_line': None,
            'pl_line': None,
        },
        {
            'id': 'previous_year',
            'title': 'Previous Year Figures',
            'always_retain': True,
            'bs_line': None,
            'pl_line': None,
        },
        {
            'id': 'rounding_off',
            'title': 'Rounding-off',
            'always_retain': True,
            'bs_line': None,
            'pl_line': None,
        },
    ]

    return defs


def build_note_map(classified, constitution=None):
    # Auto-detect constitution from classified data if not explicitly passed
    if constitution is None:
        constitution = classified.get('constitution', 'proprietorship')

    retained = []
    suppressed = []

    note_definitions = _get_note_definitions(constitution)

    for note_def in note_definitions:
        if note_def.get('always_retain'):
            retained.append(note_def.copy())
            continue

        data_keys = note_def.get('data_keys', [])
        if data_keys and _has_data(classified, *data_keys):
            retained.append(note_def.copy())
        else:
            suppressed.append(note_def.copy())

    for idx, note in enumerate(retained):
        note['number'] = idx + 1

    note_map = {}
    for note in retained:
        note_map[note['id']] = note

    bs_line_to_note = {}
    pl_line_to_note = {}
    for note in retained:
        if note.get('bs_line'):
            bs_line_to_note[note['bs_line']] = note['number']
        if note.get('pl_line'):
            pl_line_to_note[note['pl_line']] = note['number']

    return {
        'retained': retained,
        'suppressed': suppressed,
        'note_map': note_map,
        'bs_line_to_note': bs_line_to_note,
        'pl_line_to_note': pl_line_to_note,
        'last_note_number': len(retained),
    }
