"""
transform.py
------------
Step 4 - converts the legacy trial balance into a SIMPLIFIED FBDI-style GL
journal (a simplified version, NOT the official Oracle JournalImportTemplate.xlsm)
and collects validation errors.

Column order and the values STATUS = NEW and ACTUAL_FLAG = A follow Oracle's
public GL_INTERFACE documentation and GlInterface.ctl control file. Only the
first five SEGMENT columns are used. Description maps to REFERENCE10
(journal entry line description).

Validation checks:
  UNMAPPED          legacy combination has no target combination
  RETIRED BALANCE   retired combination still carries a balance
  MISSING SEGMENT   mapped combination with one or more blank segments
  DUPLICATE         combination listed more than once in the legacy COA
  UNBALANCED        journal debits <> credits (overall, or per balancing segment)

A check is BLOCKING when the affected combination carries a non-zero balance
(it would drop value from the load). Zero-balance findings are still logged
as errors to be cleared, but do not block the load.

Outputs:
  data/cloud/fbdi_gl_interface.csv
  data/cloud/validation_errors.csv
"""

from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
LEGACY = ROOT / "data" / "legacy"
CLOUD = ROOT / "data" / "cloud"

LEDGER = "MCS USD Primary Ledger"         # fictional ledger name
ACCOUNTING_DATE = "2026/10/31"            # GlInterface.ctl date format YYYY/MM/DD
SOURCE = "Legacy Conversion"              # fictional journal source (<= 25 chars)
CATEGORY = "Balance Conversion"           # fictional journal category (<= 25 chars)
CURRENCY = "USD"

FBDI_COLS = ["Status", "Ledger", "Accounting Date", "Source", "Category", "Currency",
             "Actual Flag", "Segment1", "Segment2", "Segment3", "Segment4", "Segment5",
             "Entered Dr", "Entered Cr", "Description"]
SEG = ["Segment1", "Segment2", "Segment3", "Segment4", "Segment5"]
ERR_COLS = ["Issue ID", "Check", "Severity", "Blocks Load", "Legacy Combination",
            "Target Combination", "Balance Impact", "Detail", "Resolution"]


def to_cents(x) -> int:
    return int(round(float(x) * 100))


def transform(tb: pd.DataFrame, mapping: pd.DataFrame, coa: pd.DataFrame):
    tb = tb.copy()
    tb["Net"] = tb["Debit"] - tb["Credit"]
    m = mapping.set_index("Legacy Combination")
    errors = []

    def err(check, combo, target, bal, detail, resolution):
        blocking = to_cents(bal) != 0
        errors.append([None, check, "Error", "Yes" if blocking else "No", combo, target,
                       round(bal, 2), detail, resolution])

    # mapping-level findings (cover zero-balance accounts too)
    for combo, row in m.iterrows():
        bal = float(row["Legacy Net Balance"])
        if row["Mapping Type"] == "Unmapped":
            err("UNMAPPED", combo, "", bal, row["Notes"],
                "Account owner to confirm retirement or provide a target before go-live.")
        elif row["Mapping Type"] == "Retired" and to_cents(bal) != 0:
            err("RETIRED BALANCE", combo, "", bal, "Retired combination still carries a balance.",
                "Reclassify balance in legacy before cutover.")
        elif row["Mapping Type"] in ("1:1", "Many:1"):
            blanks = [s for s in SEG if row[s] == ""]
            if blanks:
                err("MISSING SEGMENT", combo, "", bal,
                    f"Blank target segment(s): {', '.join(blanks)}. {row['Notes']}",
                    "Assign cost center owner, or retire if no longer used.")
    dups = coa.loc[coa["Legacy Combination"].duplicated(keep=False)]
    for combo, grp in dups.groupby("Legacy Combination"):
        descs = " | ".join(f"'{d}'" for d in grp["Description"])
        errors.append([None, "DUPLICATE", "Warning", "No", combo, m.at[combo, "Target Combination"]
                       if combo in m.index else "", 0.0,
                       f"Listed {len(grp)} times in legacy COA: {descs}.",
                       "De-duplicated in mapping; first description kept."])

    # trial-balance lines that cannot be loaded are excluded from the journal
    tb = tb.join(m[SEG + ["Mapping Type", "Target Combination"]], on="Legacy Combination")
    tb["Mapping Type"] = tb["Mapping Type"].fillna("Unmapped")
    for c in SEG + ["Target Combination"]:
        tb[c] = tb[c].fillna("")
    loadable = tb["Mapping Type"].isin(["1:1", "Many:1"]) & (tb[SEG] != "").all(axis=1)
    for r in tb[~loadable].itertuples():
        if r[2] not in m.index:  # TB line with no row in the mapping at all
            err("UNMAPPED", r[2], "", r.Net, "Trial balance line not found in mapping table.",
                "Add to mapping table.")

    # aggregate to one journal line per target combination
    agg = (tb[loadable].groupby(SEG, as_index=False)
             .agg(Net=("Net", "sum"),
                  Sources=("Legacy Combination", lambda s: ", ".join(sorted(s)))))
    agg["Net"] = agg["Net"].round(2)
    agg = agg[agg["Net"].apply(to_cents) != 0]
    j = pd.DataFrame({
        "Status": "NEW", "Ledger": LEDGER, "Accounting Date": ACCOUNTING_DATE,
        "Source": SOURCE, "Category": CATEGORY, "Currency": CURRENCY, "Actual Flag": "A",
    }, index=agg.index)
    for s in SEG:
        j[s] = agg[s]
    j["Entered Dr"] = agg["Net"].clip(lower=0).round(2)
    j["Entered Cr"] = (-agg["Net"]).clip(lower=0).round(2)
    j["Description"] = "Oct-26 balance conversion from legacy " + agg["Sources"]
    j = j.sort_values(SEG).reset_index(drop=True)[FBDI_COLS]

    # balance checks: overall and per primary balancing segment (Company)
    def unbalanced(frame, scope):
        diff = sum(map(to_cents, frame["Entered Dr"])) - sum(map(to_cents, frame["Entered Cr"]))
        if diff:
            errors.append([None, "UNBALANCED", "Error", "Yes", "", scope, diff / 100,
                           f"Journal debits exceed credits by {diff / 100:,.2f} for {scope}.",
                           "Investigate mapping; journal must balance before import."])
    unbalanced(j, "all companies")
    for co, grp in j.groupby("Segment1"):
        unbalanced(grp, f"Company {co}")

    errs = pd.DataFrame(errors, columns=ERR_COLS)
    errs["Issue ID"] = [f"V{i:02d}" for i in range(1, len(errs) + 1)]
    return j, errs


def self_test(tb, mapping, coa) -> None:
    """Negative test: break the inputs on purpose and confirm each check fires."""
    bad_map = mapping.copy()
    i = bad_map.index[bad_map["Legacy Combination"] == "01-55100-900"][0]
    bad_map.loc[i, "Segment3"] = ""                                   # missing segment
    bad_map.loc[i, "Target Combination"] = ""
    bad_tb = pd.concat([tb, pd.DataFrame([["Oct-26", "01-77777-900", "Test", 1000.0, 0.0]],
                                         columns=tb.columns)], ignore_index=True)  # unmapped + unbalanced
    _, e = transform(bad_tb, bad_map, coa)
    fired = set(e.loc[e["Blocks Load"] == "Yes", "Check"])
    assert {"MISSING SEGMENT", "UNMAPPED"} <= fired, fired
    unbal_tb = tb.copy()
    unbal_tb.loc[0, "Debit"] += 0.01                                   # 1 cent out of balance
    _, e = transform(unbal_tb, mapping, coa)
    assert "UNBALANCED" in set(e["Check"]), "unbalanced check did not fire"
    print("Self-test passed: UNMAPPED, MISSING SEGMENT and UNBALANCED checks all fire on bad input.")


def main() -> None:
    tb = pd.read_csv(LEGACY / "legacy_trial_balance.csv", dtype={"Legacy Combination": str})
    coa = pd.read_csv(LEGACY / "legacy_coa.csv", dtype=str)
    mapping = pd.read_csv(CLOUD / "coa_mapping.csv", dtype=str, keep_default_na=False)
    mapping["Legacy Net Balance"] = mapping["Legacy Net Balance"].astype(float)

    journal, errors = transform(tb, mapping, coa)
    self_test(tb, mapping, coa)

    journal.to_csv(CLOUD / "fbdi_gl_interface.csv", index=False, float_format="%.2f")
    errors.to_csv(CLOUD / "validation_errors.csv", index=False, float_format="%.2f")
    blocking = (errors["Blocks Load"] == "Yes").sum()
    print(f"FBDI journal: {len(journal)} lines, Dr {journal['Entered Dr'].sum():,.2f} = "
          f"Cr {journal['Entered Cr'].sum():,.2f} | validation issues: {len(errors)} "
          f"({blocking} blocking)")
    if blocking:
        raise SystemExit("Blocking validation errors - fix before import.")


if __name__ == "__main__":
    main()
