"""
Cleans and validates raw Dynamics 365 exports and produces a single
consolidated, model-ready CSV per data domain. This is the kind of
validation a Controller/FP&A Analyst performs before month-end close sign-off:
  - standardize column names / casing
  - drop duplicate journal lines
  - flag and fix blank cost centers
  - reconcile FX (convert EUR entities to USD using a stated month-end rate)
  - produce a data-quality log for the close checklist
"""
import pandas as pd
import numpy as np

RAW = "/home/claude/fpa-project/data/d365_raw_exports"
CLEAN = "/home/claude/fpa-project/data/clean"

# Month-end EUR/USD rates (would come from Treasury / ECB in real life)
FX = {m: r for m, r in zip(
    pd.date_range("2025-01-01", "2025-12-01", freq="MS").strftime("%Y-%m"),
    [1.09, 1.08, 1.09, 1.07, 1.08, 1.10, 1.09, 1.08, 1.07, 1.09, 1.10, 1.11]
)}

dq_log = []

def log(step, detail, count):
    dq_log.append({"Step": step, "Detail": detail, "RowsAffected": count})

# --- Chart of accounts ---
coa = pd.read_csv(f"{RAW}/chart_of_accounts_export.csv")
coa.to_csv(f"{CLEAN}/chart_of_accounts.csv", index=False)

# --- Cost centers ---
cc = pd.read_csv(f"{RAW}/cost_center_master_export.csv")
before = cc["CostCenter"].tolist()
cc["CostCenter"] = cc["CostCenter"].str.strip().str.upper()
log("Cost center cleanup", "Trimmed whitespace / standardized casing", (pd.Series(before) != cc["CostCenter"]).sum())
cc.to_csv(f"{CLEAN}/cost_centers.csv", index=False)

# --- GL Trial Balance (P&L) ---
tb = pd.read_csv(f"{RAW}/gl_trial_balance_export.csv")
tb.columns = [c.strip() for c in tb.columns]
tb["LegalEntity"] = tb["LegalEntity"].str.strip().str.upper()

dupe_count = tb.duplicated().sum()
tb = tb.drop_duplicates()
log("Duplicate GL lines", "Removed exact-duplicate journal rows from D365 export", int(dupe_count))

blank_cc = tb["CostCenter"].isna().sum() + (tb["CostCenter"] == "").sum()
tb["CostCenter"] = tb["CostCenter"].fillna("UNALLOCATED").replace("", "UNALLOCATED")
log("Blank cost centers", "Mapped blank/non-dimensioned lines to UNALLOCATED for review", int(blank_cc))

tb = tb.merge(coa[["MainAccount", "AccountName", "AccountCategory"]], on="MainAccount", how="left")

def to_usd(row):
    if row["Currency"] == "USD":
        return row["Amount"]
    return round(row["Amount"] * FX[row["Period"]], 2)

tb["AmountUSD"] = tb.apply(to_usd, axis=1)
log("FX translation", "Converted EUR entity amounts to USD at month-end rate", int((tb["Currency"] == "EUR").sum()))

tb.to_csv(f"{CLEAN}/gl_trial_balance_clean.csv", index=False)

# --- Balance Sheet ---
bs = pd.read_csv(f"{RAW}/gl_balance_sheet_export.csv")
bs = bs.merge(coa[["MainAccount", "AccountName", "AccountCategory"]], on="MainAccount", how="left")
bs["AmountUSD"] = bs.apply(lambda r: r["Amount"] if r["Currency"] == "USD" else round(r["Amount"] * FX[r["Period"]], 2), axis=1)
bs.to_csv(f"{CLEAN}/balance_sheet_clean.csv", index=False)

# --- Budget ---
bud = pd.read_csv(f"{RAW}/budget_register_export.csv")
bud = bud.merge(coa[["MainAccount", "AccountName", "AccountCategory"]], on="MainAccount", how="left")
bud.to_csv(f"{CLEAN}/budget_clean.csv", index=False)

# --- Data quality log (feeds the Month-End Close Checklist tab) ---
pd.DataFrame(dq_log).to_csv(f"{CLEAN}/data_quality_log.csv", index=False)

print("Clean files written to", CLEAN)
print(pd.DataFrame(dq_log).to_string(index=False))
