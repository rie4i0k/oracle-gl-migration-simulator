"""
generate_legacy_data.py
-----------------------
Step 1 - builds the FICTIONAL legacy (on-premises) general ledger data for
Meridian Card Services (fictional):

  data/legacy/legacy_coa.csv            ~60 rows, 3-part Company-Account-CostCenter
  data/legacy/legacy_trial_balance.csv  balanced trial balance, period Oct-26

The legacy chart of accounts is deliberately messy, the way a 15-year-old
on-prem ledger usually is:
  * duplicate combinations (same code listed twice, different descriptions)
  * near-duplicate accounts that should collapse into one target account
  * inactive / obsolete accounts (some must be retired)
  * inconsistent descriptions (CAPS, abbreviations, trailing spaces)
  * a cost center (110) with no owner in the new structure
  * junk accounts nobody can explain (left unmapped on purpose)

All codes, descriptions and amounts are fictional. Amounts are YTD balances in
USD; debits are positive, credits negative. Each company balances on its own
(retained earnings is the balancing line per company).
"""

from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "data" / "legacy"
PERIOD = "Oct-26"

# Legacy cost center -> legacy reporting division (the old system's own
# management hierarchy; it is NOT derived from the new segment design).
LEGACY_DIVISION = {
    "000": "CORPORATE",
    "100": "CONSUMER",
    "110": "CONSUMER",
    "200": "COMMERCIAL",
    "210": "COMMERCIAL",
    "300": "INTERNATIONAL",
    "400": "MERCHANT & NETWORK",
    "410": "MERCHANT & NETWORK",
    "900": "CORPORATE",
}

# (company, account, cost center, description, type, legacy FS line, status,
#  last activity, YTD balance in USD; None = balancing line)
COA = [
    # ---------------- assets ----------------
    ("01", "10100", "000", "Cash - Operating", "ASSET", "CASH", "Active", "2026-10-31", 1_850_000_000.00),
    ("01", "10110", "000", "CASH-OPERATING ACCT", "ASSET", "CASH", "Inactive", "2019-03-31", 0.00),
    ("02", "10200", "000", "Due From Banks - Intl", "ASSET", "CASH", "Active", "2026-10-31", 420_000_000.00),
    ("01", "12000", "100", "Card Member Receivables - Consumer", "ASSET", "CM RECV", "Active", "2026-10-31", 4_200_118_402.17),
    ("01", "12000", "200", "CM Recv - Corporate Card", "ASSET", "CM RECV", "Active", "2026-10-31", 3_100_552_910.40),
    ("01", "12010", "200", "CM RECV CORP CARD", "ASSET", "CM RECV", "Active", "2026-10-30", 260_004_118.25),
    ("02", "12000", "300", "Card Member Receivables - Intl ", "ASSET", "CM RECV", "Active", "2026-10-31", 1_350_207_655.80),
    ("01", "12900", "000", "Reserve for Credit Losses - Receivables", "ASSET", "CM RECV", "Active", "2026-10-31", -95_000_000.00),
    ("01", "13000", "100", "Card Member Loans - Consumer", "ASSET", "CM LOANS", "Active", "2026-10-31", 9_800_431_220.66),
    ("01", "13000", "210", "Card Member Loans - Small Business", "ASSET", "CM LOANS", "Active", "2026-10-31", 1_900_250_330.10),
    ("02", "13000", "300", "CM Loans Intl", "ASSET", "CM LOANS", "Active", "2026-10-31", 1_100_009_871.45),
    ("01", "13900", "000", "Reserve for Credit Losses - Loans", "ASSET", "CM LOANS", "Active", "2026-10-31", -520_000_000.00),
    ("01", "14000", "400", "Merchant Settlement Receivable", "ASSET", "OTH ASSET", "Active", "2026-10-31", 640_337_120.93),
    ("01", "15000", "900", "I/C Receivable - Meridian Intl", "ASSET", "OTH ASSET", "Active", "2026-10-31", 310_000_000.00),
    ("01", "16000", "900", "Premises & Equipment", "ASSET", "PREMISES", "Active", "2026-10-31", 980_450_000.00),
    ("01", "16900", "900", "Accum Depr - P&E", "ASSET", "PREMISES", "Active", "2026-10-31", -410_225_000.00),
    ("01", "17000", "900", "Prepaid Expenses", "ASSET", "OTH ASSET", "Active", "2026-10-31", 145_018_450.00),
    ("01", "17010", "900", "PREPAID EXP (OLD)", "ASSET", "OTH ASSET", "Inactive", "2018-12-31", 0.00),
    ("01", "19999", "000", "CONVERSION CLEARING", "ASSET", "OTH ASSET", "Active", "2021-06-30", 0.00),
    # ---------------- liabilities ----------------
    ("01", "20000", "900", "Customer Deposits", "LIABILITY", "DEPOSITS", "Active", "2026-10-31", -12_600_880_412.55),
    ("01", "21000", "400", "Accounts Payable - Merchants", "LIABILITY", "AP", "Active", "2026-10-31", -2_450_118_903.42),
    ("01", "21010", "400", "A/P MERCHANT SETTLEMENT", "LIABILITY", "AP", "Active", "2026-10-29", -180_002_115.00),
    ("01", "22000", "100", "Membership Rewards Liability", "LIABILITY", "OTH LIAB", "Active", "2026-10-31", -1_650_440_208.31),
    ("02", "22000", "300", "Rewards Liab - Intl", "LIABILITY", "OTH LIAB", "Active", "2026-10-31", -230_118_660.12),
    ("01", "23000", "900", "Accrued Expenses", "LIABILITY", "OTH LIAB", "Active", "2026-10-31", -390_771_402.08),
    ("02", "25000", "900", "I/C Payable - Meridian US", "LIABILITY", "OTH LIAB", "Active", "2026-10-31", -310_000_000.00),
    ("01", "26000", "900", "Long-Term Debt", "LIABILITY", "LTD", "Active", "2026-10-31", -3_900_000_000.00),
    # ---------------- equity ----------------
    ("01", "30000", "900", "Common Stock", "EQUITY", "EQUITY", "Active", "2026-01-01", -50_000_000.00),
    ("01", "31000", "900", "Retained Earnings", "EQUITY", "RET EARN", "Active", "2026-01-01", None),
    ("02", "31000", "900", "Retained Earnings - Intl", "EQUITY", "RET EARN", "Active", "2026-01-01", None),
    # ---------------- revenue ----------------
    ("01", "40000", "400", "Discount Revenue - US Merchants", "REVENUE", "DISC REV", "Active", "2026-10-31", -2_350_602_118.40),
    ("01", "40010", "400", "DISC REV MERCH", "REVENUE", "DISC REV", "Active", "2026-10-31", -95_110_402.16),
    ("02", "40000", "400", "Discount Revenue - Intl Merchants", "REVENUE", "DISC REV", "Active", "2026-10-31", -410_330_995.71),
    ("01", "41000", "100", "Card Fees - Consumer", "REVENUE", "CARD FEES", "Active", "2026-10-31", -720_118_400.00),
    ("01", "41000", "200", "card fees - corporate", "REVENUE", "CARD FEES", "Active", "2026-10-31", -260_440_150.00),
    ("02", "41000", "300", "NET CARD FEES INTL", "REVENUE", "CARD FEES", "Active", "2026-10-31", -140_221_870.00),
    ("02", "42000", "300", "Service Fees - FX Conversion", "REVENUE", "SVC FEES", "Active", "2026-10-31", -98_440_215.33),
    ("01", "42010", "100", "Delinquency Fees", "REVENUE", "SVC FEES", "Active", "2026-10-31", -64_118_002.19),
    ("01", "42020", "410", "Network Partnership Fees", "REVENUE", "SVC FEES", "Active", "2026-10-31", -150_660_418.00),
    ("01", "43000", "100", "Interest Income - CM Loans Consumer", "REVENUE", "INT INC", "Active", "2026-10-31", -1_480_225_119.84),
    ("01", "43000", "210", "Int Inc - CM Loans SB", "REVENUE", "INT INC", "Active", "2026-10-31", -260_118_440.29),
    ("01", "43100", "900", "Interest on Deposits w/ Banks", "REVENUE", "INT INC", "Active", "2026-10-31", -150_442_009.12),
    # ---------------- interest expense ----------------
    ("01", "44000", "900", "Interest Expense - Customer Deposits", "EXPENSE", "INT EXP", "Active", "2026-10-31", 360_118_402.55),
    ("01", "44100", "900", "Int Exp - LT Debt", "EXPENSE", "INT EXP", "Active", "2026-10-31", 190_004_118.00),
    # ---------------- provisions ----------------
    ("01", "45000", "200", "Provision - CM Receivables", "EXPENSE", "PROVISION", "Active", "2026-10-31", 62_118_440.00),
    ("01", "45100", "100", "Provision - CM Loans", "EXPENSE", "PROVISION", "Active", "2026-10-31", 330_440_118.62),
    ("02", "45100", "300", "Provision CM Loans Intl", "EXPENSE", "PROVISION", "Active", "2026-10-31", 41_002_995.18),
    # ---------------- operating expenses ----------------
    ("01", "50000", "100", "Card Member Rewards - Consumer", "EXPENSE", "REWARDS", "Active", "2026-10-31", 1_280_440_219.47),
    ("02", "50000", "300", "CM Rewards - Intl", "EXPENSE", "REWARDS", "Active", "2026-10-31", 165_118_002.09),
    ("01", "50010", "100", "REWARDS EXP (LEGACY)", "EXPENSE", "REWARDS", "Inactive", "2020-09-30", 0.00),
    ("01", "51000", "100", "Business Development - Cobrand Partners", "EXPENSE", "BUS DEV", "Active", "2026-10-31", 520_660_118.20),
    ("01", "51000", "200", "Bus Dev - Corporate Clients", "EXPENSE", "BUS DEV", "Active", "2026-10-31", 140_118_225.00),
    ("01", "52000", "100", "Card Member Services - Travel Benefits", "EXPENSE", "CM SVCS", "Active", "2026-10-31", 470_225_660.18),
    ("01", "52010", "110", "CM SERVICES - CONCIERGE", "EXPENSE", "CM SVCS", "Active", "2025-12-31", 0.00),
    ("01", "53000", "900", "Marketing - Brand", "EXPENSE", "MARKETING", "Active", "2026-10-31", 330_118_440.00),
    ("01", "53000", "100", "Marketing - Acquisition", "EXPENSE", "MARKETING", "Active", "2026-10-31", 210_660_225.40),
    ("01", "54000", "900", "Salaries & Benefits", "EXPENSE", "SALARIES", "Active", "2026-10-31", 760_225_118.33),
    ("01", "54000", "900", "SALARIES AND BENEFITS ", "EXPENSE", "SALARIES", "Active", "2026-10-31", None),  # duplicate row
    ("02", "54000", "900", "Salaries & benefits - Intl", "EXPENSE", "SALARIES", "Active", "2026-10-31", 120_440_660.70),
    ("01", "55000", "900", "Professional Fees", "EXPENSE", "OTHER EXP", "Active", "2026-10-31", 140_118_002.45),
    ("01", "55100", "900", "Occupancy", "EXPENSE", "OTHER EXP", "Active", "2026-10-31", 95_440_118.00),
    ("01", "55200", "900", "Data Processing", "EXPENSE", "OTHER EXP", "Active", "2026-10-31", 180_225_440.66),
    ("01", "58000", "900", "MISC EXP - DO NOT USE", "EXPENSE", "OTHER EXP", "Active", "2022-02-28", 0.00),
    ("01", "59999", "900", "OLD SYSTEM SUSPENSE", "EXPENSE", "OTHER EXP", "Inactive", "2017-05-31", 0.00),
]

COLS = ["Company", "Account", "Cost Center", "Description", "Account Type",
        "Legacy FS Line", "Status", "Last Activity", "_balance"]


def build() -> tuple[pd.DataFrame, pd.DataFrame]:
    coa = pd.DataFrame(COA, columns=COLS)
    coa["Legacy Combination"] = coa["Company"] + "-" + coa["Account"] + "-" + coa["Cost Center"]
    coa["Legacy Division"] = coa["Cost Center"].map(LEGACY_DIVISION)
    dup = coa.duplicated("Legacy Combination", keep="first")

    # Balancing line per company: retained earnings absorbs the difference.
    bal = coa[~dup].copy()
    for company in ("01", "02"):
        is_co = bal["Company"] == company
        plug_mask = is_co & bal["_balance"].isna()
        assert plug_mask.sum() == 1
        bal.loc[plug_mask, "_balance"] = -round(bal.loc[is_co, "_balance"].sum(skipna=True), 2)

    tb = bal[bal["_balance"].fillna(0) != 0][["Legacy Combination", "Description", "_balance"]].copy()
    tb["_balance"] = tb["_balance"].round(2)
    tb["Debit"] = tb["_balance"].clip(lower=0).round(2)
    tb["Credit"] = (-tb["_balance"]).clip(lower=0).round(2)
    tb.insert(0, "Period", PERIOD)
    tb = tb.drop(columns="_balance")

    coa = coa[["Legacy Combination", "Company", "Account", "Cost Center", "Description",
               "Account Type", "Legacy Division", "Legacy FS Line", "Status", "Last Activity"]]
    return coa, tb


def validate(coa: pd.DataFrame, tb: pd.DataFrame) -> None:
    assert 55 <= len(coa) <= 65, len(coa)
    assert coa["Legacy Combination"].duplicated().sum() >= 1, "need a duplicate row"
    assert (coa["Status"] == "Inactive").sum() >= 3
    cents = (tb["Debit"] - tb["Credit"]).mul(100).round().astype("int64")
    assert cents.sum() == 0, "trial balance does not balance"
    for co in ("01", "02"):
        m = tb["Legacy Combination"].str.startswith(co + "-")
        assert cents[m].sum() == 0, f"company {co} does not balance"
    ic = tb.set_index("Legacy Combination")
    assert ic.loc["01-15000-900", "Debit"] == ic.loc["02-25000-900", "Credit"], "intercompany out of balance"


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    coa, tb = build()
    validate(coa, tb)
    coa.to_csv(OUT / "legacy_coa.csv", index=False)
    tb.to_csv(OUT / "legacy_trial_balance.csv", index=False, float_format="%.2f")
    print(f"Legacy COA: {len(coa)} rows ({coa['Legacy Combination'].nunique()} unique combinations); "
          f"trial balance: {len(tb)} lines, debits = credits = {tb['Debit'].sum():,.2f}")


if __name__ == "__main__":
    main()
