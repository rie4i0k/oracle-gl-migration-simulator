"""
reconcile.py
------------
Step 5 - proves the converted Cloud journal equals the legacy trial balance:
  * overall (total debits, total credits, net)
  * by Line of Business  - legacy side uses the OLD system's own division
                           hierarchy, Cloud side uses Segment2
  * by account group     - legacy side uses the OLD system's FS lines, Cloud
                           side uses the target natural account's group
Using the legacy system's own hierarchies (not the new mapping) on the legacy
side makes this an independent check rather than a circular one.

Outputs data/cloud/reconciliation.csv and exits non-zero if anything FAILS.
"""

from pathlib import Path

import pandas as pd

from mapping import LEGACY_DIVISION_TO_LOB, LEGACY_FS_TO_GROUP, LOB, NATURAL_ACCOUNT

ROOT = Path(__file__).resolve().parents[1]
LEGACY = ROOT / "data" / "legacy"
CLOUD = ROOT / "data" / "cloud"
TOLERANCE_CENTS = 0


def cents(s: pd.Series) -> pd.Series:
    return (s.astype(float) * 100).round().astype("int64")


def reconcile() -> pd.DataFrame:
    coa = pd.read_csv(LEGACY / "legacy_coa.csv", dtype=str).drop_duplicates("Legacy Combination")
    tb = pd.read_csv(LEGACY / "legacy_trial_balance.csv", dtype={"Legacy Combination": str})
    j = pd.read_csv(CLOUD / "fbdi_gl_interface.csv", dtype=str)

    tb = tb.merge(coa[["Legacy Combination", "Legacy Division", "Legacy FS Line"]], how="left")
    tb["dr"], tb["cr"] = cents(tb["Debit"]), cents(tb["Credit"])
    tb["net"] = tb["dr"] - tb["cr"]
    tb["LOB"] = tb["Legacy Division"].map(LEGACY_DIVISION_TO_LOB)
    tb["Group"] = tb["Legacy FS Line"].map(LEGACY_FS_TO_GROUP)
    assert tb["LOB"].notna().all() and tb["Group"].notna().all()

    j["dr"], j["cr"] = cents(j["Entered Dr"]), cents(j["Entered Cr"])
    j["net"] = j["dr"] - j["cr"]
    j["LOB"] = j["Segment2"]
    j["Group"] = j["Segment4"].map(lambda a: NATURAL_ACCOUNT[a][2])

    rows = [
        ("Overall", "Total debits", tb["dr"].sum(), j["dr"].sum()),
        ("Overall", "Total credits", tb["cr"].sum(), j["cr"].sum()),
        ("Overall", "Net (debits - credits)", tb["net"].sum(), j["net"].sum()),
    ]
    for code, name in LOB.items():
        rows.append(("Line of Business", name, tb.loc[tb["LOB"] == code, "net"].sum(),
                     j.loc[j["LOB"] == code, "net"].sum()))
    groups = list(dict.fromkeys(v[2] for v in NATURAL_ACCOUNT.values()))
    for g in groups:
        rows.append(("Account group", g, tb.loc[tb["Group"] == g, "net"].sum(),
                     j.loc[j["Group"] == g, "net"].sum()))

    r = pd.DataFrame(rows, columns=["Level", "Item", "legacy", "cloud"])
    r["variance"] = r["cloud"] - r["legacy"]
    r["Result"] = r["variance"].abs().le(TOLERANCE_CENTS).map({True: "PASS", False: "FAIL"})
    for c in ("legacy", "cloud", "variance"):
        r[c] = r[c] / 100
    return r.rename(columns={"legacy": "Legacy Amount", "cloud": "Cloud Amount", "variance": "Variance"})


def main() -> None:
    r = reconcile()
    r.to_csv(CLOUD / "reconciliation.csv", index=False, float_format="%.2f")
    fails = r[r["Result"] == "FAIL"]
    overall = "PASS" if fails.empty else "FAIL"
    print(f"Reconciliation: {overall} ({len(r)} checks, {len(fails)} failed)")
    if not fails.empty:
        print(fails.to_string(index=False))
        raise SystemExit(1)


if __name__ == "__main__":
    main()
