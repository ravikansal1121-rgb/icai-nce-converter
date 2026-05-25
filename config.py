import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'docx', 'doc', 'pdf'}

ACCOUNT_KEYWORDS = {
    'revenue_from_operations': [
        'sales', 'sale of goods', 'sale of services', 'revenue from operations',
        'sales account', 'sales a/c', 'sale account', 'sale a/c',
        'sales (gst)', 'local sales', 'interstate sales', 'export sales',
        'job work income', 'job work charges received', 'labour charges received',
    ],
    'other_income': [
        'bank interest', 'interest received', 'interest on fd', 'interest on fixed deposit',
        'interest on bank deposit', 'dividend received', 'dividend income',
        'rent received', 'rent income', 'commission received', 'commission income',
        'discount received', 'cash discount received',
        'profit on sale of asset', 'profit on sale of fixed asset',
        'profit on sale of investment', 'capital gain',
        'miscellaneous income', 'misc income', 'misc. income', 'other income',
        'round off', 'roundoff', 'round off received',
        'insurance claim received', 'bad debts recovered',
        'foreign exchange gain', 'exchange difference',
    ],
    'purchases': [
        'purchases', 'purchase', 'purchase account', 'purchase a/c',
        'purchases account', 'purchases a/c', 'purchase of goods',
        'local purchases', 'interstate purchases', 'import purchases',
        'purchases (gst)', 'purchase of stock-in-trade',
        'purchase of raw materials', 'raw material consumed',
        'cost of materials consumed',
    ],
    'opening_stock': [
        'opening stock', 'opening stock-in-trade', 'opening inventory',
        'stock at commencement', 'opening stock of finished goods',
        'opening stock of raw materials', 'opening stock of wip',
    ],
    'closing_stock': [
        'closing stock', 'closing stock-in-trade', 'closing inventory',
        'stock at close', 'closing stock of finished goods',
        'closing stock of raw materials', 'closing stock of wip',
    ],
    'employee_benefits': [
        'salary', 'salaries', 'salary & wages', 'salaries & wages',
        'wages', 'labour charges', 'labour', 'staff salary',
        'staff welfare', 'staff welfare expenses', 'bonus', 'bonus paid',
        'pf contribution', 'provident fund', 'epf contribution',
        'esi contribution', 'esic contribution', 'gratuity',
        'leave encashment', 'incentive', 'incentives',
    ],
    'finance_costs': [
        'interest paid', 'bank interest paid', 'interest on loan',
        'interest on cc', 'interest on cash credit', 'interest on od',
        'interest on overdraft', 'interest on term loan',
        'loan processing charges', 'loan processing fees',
        'bank charges', 'bank commission', 'lc charges',
        'finance charges', 'finance cost',
    ],
    'depreciation': [
        'depreciation', 'depreciation on fixed assets', 'dep on fa',
        'amortisation', 'amortization',
    ],
    'other_expenses': [
        'rent paid', 'rent', 'office rent', 'godown rent', 'shop rent',
        'electricity', 'electricity charges', 'electricity expenses', 'power & fuel',
        'telephone', 'telephone expenses', 'mobile expenses', 'communication expenses',
        'internet charges', 'internet expenses',
        'printing & stationery', 'printing and stationery', 'stationery',
        'postage & courier', 'postage and courier', 'postage', 'courier charges',
        'travelling expenses', 'travelling', 'travel expenses', 'conveyance',
        'vehicle expenses', 'vehicle running', 'petrol expenses', 'diesel expenses',
        'repair & maintenance', 'repairs and maintenance', 'repairs', 'maintenance',
        'insurance', 'insurance premium', 'insurance expenses',
        'professional fees', 'legal & professional', 'legal and professional',
        'professional charges', 'legal fees', 'consultation charges',
        'audit fees', 'audit fee', 'auditor remuneration', 'accounting charges',
        'advertisement', 'advertisement expenses', 'advertising',
        'office expenses', 'general expenses', 'miscellaneous expenses',
        'misc expenses', 'misc. expenses', 'sundry expenses',
        'donation', 'charity', 'charitable donation',
        'subscription', 'membership fees', 'news paper',
        'loading & unloading', 'loading and unloading', 'hamali',
        'freight', 'freight charges', 'transport charges', 'cartage',
        'packaging', 'packing charges', 'packing material',
        'commission paid', 'brokerage', 'commission & brokerage',
        'bad debts', 'bad debts written off', 'irrecoverable debts',
        'discount allowed', 'cash discount', 'trade discount',
        'gst penalty', 'late fee', 'penalty', 'fine', 'interest on gst',
        'water charges', 'water expenses', 'municipal tax',
        'property tax', 'house tax', 'ground rent',
        'security charges', 'housekeeping', 'cleaning charges',
        'computer expenses', 'software expenses', 'amc charges',
        'festival expenses', 'entertainment', 'business promotion',
        'sample expenses', 'sales promotion',
        'foreign exchange loss', 'exchange loss',
        'rates & taxes', 'rates and taxes',
        'weighing charges', 'octroi', 'entry tax',
        'motor car expenses', 'two wheeler expenses',
        'tea & refreshment', 'food expenses', 'canteen expenses',
        'diwali expenses', 'pooja expenses', 'gift expenses',
    ],
    'ppe': [
        'furniture', 'furniture & fixtures', 'furniture and fixtures', 'f&f',
        'plant & machinery', 'plant and machinery', 'machinery',
        'vehicle', 'vehicles', 'motor vehicle', 'motor car', 'two wheeler', 'bike',
        'computer', 'computers', 'laptop', 'computer & printer',
        'office equipment', 'office equipments', 'air conditioner', 'ac',
        'building', 'factory building', 'office building',
        'land', 'land & building', 'land and building',
        'mobile', 'mobile phone', 'mobile handset',
        'electrical installation', 'electrical fittings',
        'tools & implements', 'dies & moulds',
        'generator', 'inverter', 'ups',
        'cctv', 'camera', 'printer', 'scanner',
    ],
    'gold_investments': [
        'gold', 'gold ornaments', 'gold jewellery', 'jewellery',
        'gold bullion', 'silver', 'silver ornaments', 'bullion',
    ],
    'trade_receivables': [
        'sundry debtors', 'trade receivables', 'debtors',
        'accounts receivable', 'book debts',
    ],
    'trade_payables': [
        'sundry creditors', 'trade payables', 'creditors',
        'accounts payable', 'trade creditors',
    ],
    'cash_hand': [
        'cash', 'cash on hand', 'cash in hand', 'cash balance',
        'cash a/c', 'cash account', 'petty cash',
    ],
    'bank_balance': [
        'bank', 'bank balance', 'bank account', 'bank a/c',
        'bank current account', 'bank savings account',
        'savings account', 'current account',
        'bank of baroda', 'state bank', 'sbi', 'hdfc bank', 'icici bank',
        'axis bank', 'kotak bank', 'pnb', 'bob', 'canara bank',
        'idbi bank', 'union bank', 'bank of india', 'boi',
        'indusind bank', 'yes bank', 'rbl bank', 'bandhan bank',
        'federal bank', 'south indian bank', 'karur vysya',
    ],
    'fixed_deposit': [
        'fixed deposit', 'fd', 'fixed deposit with bank',
        'term deposit', 'fd with bank', 'bank fd',
    ],
    'security_deposit': [
        'security deposit', 'deposit', 'deposits',
        'electricity deposit', 'rent deposit', 'telephone deposit',
        'caution deposit', 'earnest money deposit',
    ],
    'gst_input': [
        'gst input', 'gst input credit', 'input gst', 'igst input',
        'cgst input', 'sgst input', 'input tax credit', 'itc',
        'gst receivable', 'input igst', 'input cgst', 'input sgst',
        'cenvat credit', 'service tax input',
    ],
    'tds_advance_tax': [
        'tds', 'tds receivable', 'tds on income', 'tds deducted',
        'advance tax', 'advance income tax', 'self assessment tax',
        'income tax refund', 'tax deducted at source',
        'advance tax paid', 'tds on bank interest',
    ],
    'capital': [
        'capital', 'capital account', 'capital a/c',
        "proprietor's capital", 'proprietor capital',
        "owner's capital", 'owner capital',
    ],
    'partner_capital': [
        "partners' capital", 'partner capital', 'partners capital',
    ],
    'drawings': [
        'drawings', 'drawing', 'drawings a/c', 'drawing a/c',
        'personal drawings', 'household expenses',
    ],
    'bank_loan': [
        'bank loan', 'term loan', 'bank loan account',
        'secured loan', 'unsecured loan', 'loan from bank',
        'vehicle loan', 'car loan', 'machinery loan',
        'loan account', 'housing loan', 'home loan',
    ],
    'bank_od_cc': [
        'cash credit', 'cc account', 'cc a/c', 'overdraft',
        'bank overdraft', 'od account', 'od a/c',
        'cash credit account', 'cc limit',
    ],
    'unsecured_loan': [
        'unsecured loan', 'loan from relatives', 'loan from friends',
        'loan from directors', 'loan from partners',
        'personal loan', 'other loan',
    ],
    'gst_payable': [
        'gst payable', 'output gst', 'gst output', 'igst payable',
        'cgst payable', 'sgst payable', 'gst liability',
        'output igst', 'output cgst', 'output sgst',
    ],
    'tds_payable': [
        'tds payable', 'tds liability', 'tds on salary',
        'tds on rent', 'tds on professional fees',
        'tds on contractor', 'tds on interest',
    ],
    'other_statutory': [
        'professional tax', 'pt payable', 'professional tax payable',
        'esi payable', 'pf payable', 'epf payable',
        'bonus payable', 'salary payable', 'wages payable',
        'outstanding expenses', 'outstanding salary',
        'outstanding rent', 'outstanding electricity',
        'provision for expenses',
    ],
    'provision_tax': [
        'provision for income tax', 'provision for tax',
        'income tax provision', 'tax provision',
        'provision for it',
    ],
    'interest_on_capital': [
        'interest on capital', 'interest on capital a/c',
    ],
    'partner_remuneration': [
        'partner remuneration', 'partner salary', 'remuneration to partners',
        'salary to partners', 'partners salary', 'partners remuneration',
        'remuneration a/c', 'remuneration account',
    ],
    'partner_interest_on_capital': [
        'interest on capital', 'interest on capital a/c',
        'interest on partners capital', 'int on capital',
    ],
    'partner_drawings': [
        'drawings', 'drawing', 'partner drawings', 'drawings a/c',
        'personal drawings', 'household expenses',
    ],
}

BS_LIABILITY_HEADS = [
    'capital', 'partner_capital', 'reserves_surplus',
    'bank_loan', 'unsecured_loan', 'bank_od_cc',
    'trade_payables', 'gst_payable', 'tds_payable',
    'other_statutory', 'provision_tax',
]

BS_ASSET_HEADS = [
    'ppe', 'gold_investments', 'security_deposit',
    'trade_receivables', 'inventories',
    'cash_hand', 'bank_balance', 'fixed_deposit',
    'gst_input', 'tds_advance_tax',
]
