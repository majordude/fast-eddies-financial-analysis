"""Example personal_config.py for Fast Eddies Financial Spending and Income Analysis.

Copy this file to personal_config.py (same folder as analyze_statements.py) and edit. personal_config.py is ignored
by git, so your names, payees and vendors never reach GitHub. Every setting below is optional: delete what you do
not need. All values here are made-up examples.

After you change rules, delete Analysis\\processed_ledger.json and run the analyzer again so every statement is read
with the new rules (check_payees.json, vendor_overrides.json and brokerage_yearend.json apply without that).
"""

# ---- names shown on the tabs ------------------------------------------------------------------------------------
BANK_NAME = "Example Credit Union"        # also used to recognise the bank's PDFs
BROKERAGE_NAME = "ExampleBrokerage"       # where IRA / brokerage money moves from and to
ADVISOR_NAME = "Advisor"                  # prefix of the two advisor tabs and the stand-alone file
ADVISOR_FIRM = "Example Advisors"         # "Prepared for ..." line on the advisor summary

# ---- income -----------------------------------------------------------------------------------------------------
# Two Social Security deposits arrive each month: the larger belongs to the first name. Remove if not applicable.
SSA_OWNERS = ("Alex", "Sam")

# One entry per pension or employer: (regex on the deposit text, income source label, vendor text). First match wins.
# Include a "Tax refunds" entry if you want refunds shown separately.
INCOME_RULES = [
    (r"^(?=.*REFUND)(?=.*STATE OF EXAMPLE)", "Tax refunds", "STATE TAX REFUND"),
    (r"EXAMPLE PENSION", "Pension - Alex", "EXAMPLE PENSION"),
    (r"EXAMPLE ANNUITY", "Pension - Sam", "EXAMPLE ANNUITY"),
]
CHART_SHORT_EXTRA = {"Pension - Alex": "Pension Alex", "Pension - Sam": "Pension Sam"}   # short chart labels

# ---- brokerage, 529, mortgage ---------------------------------------------------------------------------------
BROKER_PAT = r"EXAMPLEBROKER|BROKERAGE ACH"           # bank text for transfers to and from the brokerage
BROKER_VENDOR_IN = "EXAMPLEBROKER (TRANSFER IN)"
PLAN_529_LABEL = "529 plan contributions"
PLAN_529_VENDOR = "529 PLAN CONTRIBUTION"
PLAN_529_PAT = r"EXAMPLE 529 PLAN"                # leave out if you have none
MORTGAGE_PAT = r"MTG PYMT|EXAMPLE MORTGAGE"
MORTGAGE_VENDOR = "EXAMPLE MORTGAGE"
MORTGAGE_EXTRA_MIN = 5000                         # payments this large or larger count as extra principal (0 = off)

# ---- payees on the bank statement ------------------------------------------------------------------------------
# (regex on the full description, category, vendor or None). Tried before the built-in rules.
BANK_RULES_FIRST = [
    (r"TO LOAN", "Auto Loan", "AUTO LOAN PAYMENT"),
    (r"EXAMPLE POWER CO", "Utilities & Telecom", "EXAMPLE POWER CO"),
    (r"EXAMPLE INSURANCE", "Insurance", "EXAMPLE INSURANCE"),
]
ADVISOR_EXTRA_CATEGORIES = {"Auto Loan": "Transportation"}     # which advisor group a custom category belongs to

# ---- vendors on credit card statements --------------------------------------------------------------------------
# (regex on the statement description, clean vendor name or None, category). Tried before the built-in rules.
CATEGORY_RULES_FIRST = [
    (r"EXAMPLE HARDWARE", "EXAMPLE HARDWARE", "Home Improvement"),
    (r"EXAMPLE CLINIC", "EXAMPLE CLINIC", "Medical & Dental"),
]
MEDICAL_TYPES_EXTRA = [(r"EXAMPLE PHARMACY", "Prescription")]

# Vendors involved in fraud or a breach are flagged: (regex, note).
BREACH_VENDORS = [(r"EXAMPLE SHADY SITE", "Unauthorized charges disputed 2025-04")]

# ---- file and card names ---------------------------------------------------------------------------------------
ACCOUNT_BY_FILE = [(r"medical", "Medical list")]    # spend-list file name pattern -> account name
REFERENCE_ACCOUNTS = {"Medical list"}               # lists that are a cross-check only, not counted
# A statement that closes in one month but should count in another. Key = card tab name | closing date.
STATEMENT_MONTH_OVERRIDES = {}                      # {"Amazon Store Card 1234|2026-01-09": "2025-12"}
# Your own keyword rules for store-card item names (tried before the built-in ones): (regex on the item name, category).
ITEM_RULES_FIRST = []                               # [(r"GUITAR|AMP\b", "Recreation & Entertainment")]
CARD_ALIASES = {}                                   # {"Old card name": "New card name"} when a card is re-numbered
LOCATION_WORDS_EXTRA = "EXAMPLETOWN EXAMPLEVILLE"   # town names that appear on statements (not vendor names)
