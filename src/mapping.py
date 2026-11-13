"""
mapping.py
----------
Steps 2 and 3 - defines the target Oracle Cloud chart of accounts (5 segments)
and maps every legacy Company-Account-CostCenter combination to it.

Outputs:
  data/cloud/target_coa.csv     segment value sets
  data/cloud/coa_mapping.csv    Legacy Combination -> Target Combination

Real-world structure used (verified, see About & Sources):
  * Line of Business values = the four reportable operating segments named in
    American Express's FY2025 Form 10-K, plus "Corporate & Other", which the
    10-K says is NOT a reportable segment.
  * Account groups = line-item captions on the 10-K's Consolidated Statements
    of Income and Consolidated Balance Sheets.
  * Segment labels follow Oracle's public chart-of-accounts documentation.
Everything else (codes, descriptions, cost centers) is fictional.
"""

from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
LEGACY = ROOT / "data" / "legacy"
CLOUD = ROOT / "data" / "cloud"

SEGMENTS = [
    # (Oracle column, segment name, segment label as documented by Oracle)
    ("SEGMENT1", "Company", "Primary Balancing Segment"),
    ("SEGMENT2", "Line of Business", "Second Balancing Segment"),
    ("SEGMENT3", "Cost Center", ""),  # no label: "Cost Center Segment" not verified in Oracle docs
    ("SEGMENT4", "Natural Account", "Natural Account Segment"),
    ("SEGMENT5", "Intercompany", "Intercompany Segment"),
]

COMPANY = {"101": "Meridian Card Services US (fictional)",
           "201": "Meridian Card Services International (fictional)"}

# Verified names (FY2025 10-K). Codes are fictional.
LOB = {
    "10": "U.S. Consumer Services",
    "20": "Commercial Services",
    "30": "International Card Services",
    "40": "Global Merchant and Network Services",
    "90": "Corporate & Other",
}

COST_CENTER = {  # fictional
    "1100": ("Consumer Card Operations", "10"),
    "2100": ("Corporate Card", "20"),
    "2200": ("Small Business Card", "20"),
    "3100": ("International Card Operations", "30"),
    "4100": ("Merchant Services", "40"),
    "4200": ("Network Partnerships", "40"),
    "9000": ("Balance Sheet - No Cost Center", "90"),
    "9100": ("Corporate Finance & Treasury", "90"),
}

INTERCOMPANY = {"000": "No Intercompany",
                "101": "Meridian Card Services US (fictional)",
                "201": "Meridian Card Services International (fictional)"}

# Natural accounts (codes and descriptions fictional). Account Group = the
# 10-K financial-statement caption the account rolls up to.
NATURAL_ACCOUNT = {
    "110000": ("Cash and Due from Banks", "Asset", "Cash and cash equivalents"),
    "121000": ("Card Member Receivables", "Asset", "Card Member receivables"),
    "129000": ("Reserve for Credit Losses - Receivables", "Asset", "Card Member receivables"),
    "131000": ("Card Member Loans", "Asset", "Card Member loans"),
    "139000": ("Reserve for Credit Losses - Loans", "Asset", "Card Member loans"),
    "141000": ("Merchant Settlement Receivable", "Asset", "Other assets"),
    "151000": ("Intercompany Receivable", "Asset", "Other assets"),
    "161000": ("Premises and Equipment", "Asset", "Premises and equipment"),
    "169000": ("Accumulated Depreciation", "Asset", "Premises and equipment"),
    "171000": ("Prepaid Expenses", "Asset", "Other assets"),
    "211000": ("Customer Deposits", "Liability", "Customer deposits"),
    "221000": ("Accounts Payable - Merchants", "Liability", "Accounts payable"),
    "231000": ("Membership Rewards Liability", "Liability", "Other liabilities"),
    "232000": ("Accrued Expenses", "Liability", "Other liabilities"),
    "251000": ("Intercompany Payable", "Liability", "Other liabilities"),
    "261000": ("Long-Term Debt", "Liability", "Long-term debt"),
    "311000": ("Common Shares", "Owners' Equity", "Common shares"),
    "321000": ("Retained Earnings", "Owners' Equity", "Retained earnings"),
    "411000": ("Discount Revenue", "Revenue", "Discount revenue"),
    "421000": ("Net Card Fees", "Revenue", "Net card fees"),
    "431000": ("Service Fees", "Revenue", "Service fees and other revenue"),
    "432000": ("Network Partnership Revenue", "Revenue", "Service fees and other revenue"),
    "441000": ("Interest on Loans", "Revenue", "Total interest income"),
    "442000": ("Interest on Deposits with Banks", "Revenue", "Total interest income"),
    "451000": ("Interest Expense - Deposits", "Expense", "Total interest expense"),
    "452000": ("Interest Expense - Long-Term Debt", "Expense", "Total interest expense"),
    "461000": ("Provision - Card Member Receivables", "Expense", "Provisions for credit losses"),
    "462000": ("Provision - Card Member Loans", "Expense", "Provisions for credit losses"),
    "511000": ("Card Member Rewards", "Expense", "Card Member rewards"),
    "521000": ("Business Development", "Expense", "Business development"),
    "531000": ("Card Member Services", "Expense", "Card Member services"),
    "541000": ("Marketing", "Expense", "Marketing"),
    "551000": ("Salaries and Employee Benefits", "Expense", "Salaries and employee benefits"),
    "561000": ("Professional Fees", "Expense", "Other, net"),
    "562000": ("Occupancy", "Expense", "Other, net"),
    "563000": ("Data Processing", "Expense", "Other, net"),
}

# Legacy FS line (old reporting hierarchy) -> 10-K caption. Used on the legacy
# side of the reconciliation, independently of the account-level mapping.
LEGACY_FS_TO_GROUP = {
    "CASH": "Cash and cash equivalents", "CM RECV": "Card Member receivables",
    "CM LOANS": "Card Member loans", "OTH ASSET": "Other assets",
    "PREMISES": "Premises and equipment", "DEPOSITS": "Customer deposits",
    "AP": "Accounts payable", "OTH LIAB": "Other liabilities", "LTD": "Long-term debt",
    "EQUITY": "Common shares", "RET EARN": "Retained earnings",
    "DISC REV": "Discount revenue", "CARD FEES": "Net card fees",
    "SVC FEES": "Service fees and other revenue", "INT INC": "Total interest income",
    "INT EXP": "Total interest expense", "PROVISION": "Provisions for credit losses",
    "REWARDS": "Card Member rewards", "BUS DEV": "Business development",
    "CM SVCS": "Card Member services", "MARKETING": "Marketing",
    "SALARIES": "Salaries and employee benefits", "OTHER EXP": "Other, net",
}

LEGACY_DIVISION_TO_LOB = {
    "CONSUMER": "10", "COMMERCIAL": "20", "INTERNATIONAL": "30",
    "MERCHANT & NETWORK": "40", "CORPORATE": "90",
}

# ---------------------------------------------------------------------------
# Mapping rules
# ---------------------------------------------------------------------------
COMPANY_MAP = {"01": "101", "02": "201"}
COST_CENTER_MAP = {  # legacy CC -> target CC (110 has no owner in the new design)
    "000": "9000", "100": "1100", "200": "2100", "210": "2200", "300": "3100",
    "400": "4100", "410": "4200", "900": "9100",
}
ACCOUNT_MAP = {
    "10100": "110000", "10200": "110000", "12000": "121000", "12010": "121000",
    "12900": "129000", "13000": "131000", "13900": "139000", "14000": "141000",
    "15000": "151000", "16000": "161000", "16900": "169000", "17000": "171000",
    "20000": "211000", "21000": "221000", "21010": "221000", "22000": "231000",
    "23000": "232000", "25000": "251000", "26000": "261000", "30000": "311000",
    "31000": "321000", "40000": "411000", "40010": "411000", "41000": "421000",
    "42000": "431000", "42010": "431000", "42020": "432000", "43000": "441000",
    "43100": "442000", "44000": "451000", "44100": "452000", "45000": "461000",
    "45100": "462000", "50000": "511000", "51000": "521000", "52000": "531000",
    "52010": "531000", "53000": "541000", "54000": "551000", "55000": "561000",
    "55100": "562000", "55200": "563000",
}
INTERCOMPANY_MAP = {("01", "15000"): "201", ("02", "25000"): "101"}

RETIRED = {  # obsolete, inactive, zero balance - not carried forward
    "01-10110-000": "Inactive since 2019; duplicate of operating cash. Retire.",
    "01-17010-900": "Inactive since 2018; superseded by 17000. Retire.",
    "01-50010-100": "Inactive legacy rewards account; replaced by 50000. Retire.",
    "01-59999-900": "Old-system suspense, zero balance since 2017. Retire.",
}
UNMAPPED_NOTES = {
    "01-19999-000": "Conversion clearing account from a 2021 project. No owner identified - awaiting decision.",
    "01-58000-900": "'Do not use' misc expense. Owner to confirm retirement or a target account.",
}


def target_coa() -> pd.DataFrame:
    rows = []
    for (col, name, label), values in zip(SEGMENTS, [
            {k: (v, "", "") for k, v in COMPANY.items()},
            {k: (v, "", "") for k, v in LOB.items()},
            {k: (v[0], "", "") for k, v in COST_CENTER.items()},
            {k: v for k, v in NATURAL_ACCOUNT.items()},
            {k: (v, "", "") for k, v in INTERCOMPANY.items()}]):
        for code, (desc, acct_type, group) in values.items():
            parent = LOB[COST_CENTER[code][1]] if name == "Cost Center" else ""
            rows.append([col, name, label, code, desc, acct_type, group, parent])
    return pd.DataFrame(rows, columns=["Oracle Column", "Segment", "Segment Label", "Value",
                                       "Description", "Account Type", "Account Group",
                                       "Parent Line of Business"])


def build_mapping(coa: pd.DataFrame, tb: pd.DataFrame) -> pd.DataFrame:
    bal = (tb.assign(Net=tb["Debit"] - tb["Credit"])
             .set_index("Legacy Combination")["Net"])
    uniq = coa.drop_duplicates("Legacy Combination", keep="first")
    dup_combos = set(coa.loc[coa["Legacy Combination"].duplicated(), "Legacy Combination"])

    rows = []
    for r in uniq.itertuples(index=False):
        combo = r[0]
        co, acct, cc = r.Company, r.Account, r[3]
        segs = ["", "", "", "", ""]
        notes = []
        if combo in RETIRED:
            mtype = "Retired"
            notes.append(RETIRED[combo])
        elif acct not in ACCOUNT_MAP:
            mtype = "Unmapped"
            notes.append(UNMAPPED_NOTES.get(combo, "No target natural account defined."))
        else:
            tcc = COST_CENTER_MAP.get(cc, "")
            segs = [COMPANY_MAP[co], COST_CENTER[tcc][1] if tcc else "", tcc,
                    ACCOUNT_MAP[acct], INTERCOMPANY_MAP.get((co, acct), "000")]
            mtype = "1:1"  # refined to Many:1 below
            if not tcc:
                notes.append(f"Legacy cost center {cc} has no owner in the new structure; "
                             "Cost Center and Line of Business left blank pending decision.")
        if combo in dup_combos:
            notes.append("Combination appears twice in legacy COA with different descriptions; de-duplicated.")
        rows.append([combo, r.Description.strip(), r.Status, *segs, mtype,
                     " ".join(notes), round(float(bal.get(combo, 0.0)), 2)])

    m = pd.DataFrame(rows, columns=["Legacy Combination", "Legacy Description", "Legacy Status",
                                    "Segment1", "Segment2", "Segment3", "Segment4", "Segment5",
                                    "Mapping Type", "Notes", "Legacy Net Balance"])
    complete = (m[["Segment1", "Segment2", "Segment3", "Segment4", "Segment5"]] != "").all(axis=1)
    m["Target Combination"] = ""
    m.loc[complete, "Target Combination"] = m.loc[complete, ["Segment1", "Segment2", "Segment3",
                                                             "Segment4", "Segment5"]].agg("-".join, axis=1)
    live = m["Mapping Type"].eq("1:1") & complete
    counts = m.loc[live, "Target Combination"].value_counts()
    many = live & m["Target Combination"].map(counts).gt(1)
    m.loc[many, "Mapping Type"] = "Many:1"
    for i in m.index[many]:
        others = m.loc[many & (m["Target Combination"] == m.at[i, "Target Combination"]) & (m.index != i),
                       "Legacy Combination"].tolist()
        m.at[i, "Notes"] = (f"Consolidated with {', '.join(others)}. " + m.at[i, "Notes"]).strip()
    for i in m.index[live & ~many]:
        if not m.at[i, "Notes"]:
            m.at[i, "Notes"] = "Direct mapping."

    m["Flag"] = "OK"
    m.loc[m["Mapping Type"].eq("Retired"), "Flag"] = "RETIRED"
    m.loc[m["Mapping Type"].eq("Retired") & m["Legacy Net Balance"].ne(0), "Flag"] = "RETIRED WITH BALANCE"
    m.loc[m["Mapping Type"].eq("Unmapped"), "Flag"] = "UNMAPPED"
    m.loc[m["Mapping Type"].isin(["1:1", "Many:1"]) & ~complete, "Flag"] = "MISSING SEGMENT"
    return m[["Legacy Combination", "Legacy Description", "Legacy Status", "Legacy Net Balance",
              "Target Combination", "Segment1", "Segment2", "Segment3", "Segment4", "Segment5",
              "Mapping Type", "Notes", "Flag"]]


def main() -> None:
    CLOUD.mkdir(parents=True, exist_ok=True)
    coa = pd.read_csv(LEGACY / "legacy_coa.csv", dtype=str)
    tb = pd.read_csv(LEGACY / "legacy_trial_balance.csv", dtype={"Legacy Combination": str})
    tcoa = target_coa()
    mapping = build_mapping(coa, tb)

    # integrity: every target value used must exist in the target COA
    valid = {col: set(tcoa.loc[tcoa["Oracle Column"] == f"SEGMENT{i}", "Value"])
             for i, col in enumerate(["Segment1", "Segment2", "Segment3", "Segment4", "Segment5"], 1)}
    for col, vals in valid.items():
        used = set(mapping[col]) - {""}
        assert used <= vals, f"{col} uses undefined values {used - vals}"
    assert set(LEGACY_FS_TO_GROUP.values()) <= {v[2] for v in NATURAL_ACCOUNT.values()}

    tcoa.to_csv(CLOUD / "target_coa.csv", index=False)
    mapping.to_csv(CLOUD / "coa_mapping.csv", index=False, float_format="%.2f")
    print("Mapping types:", mapping["Mapping Type"].value_counts().to_dict(),
          "| flags:", mapping["Flag"].value_counts().to_dict())


if __name__ == "__main__":
    main()
