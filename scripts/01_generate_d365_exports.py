"""
Simulates raw data exports from Microsoft Dynamics 365 Finance & Operations.
In a real environment these would come from:
  - D365 F&O 'General ledger > Trial balance' export, or the GeneralJournalAccountEntry
    entity via the D365 OData/Data Management Framework
  - D365 F&O 'Chart of accounts' export (MainAccount entity)
  - D365 F&O 'Cost center' financial dimension export (DimensionFinancialTag entity)
  - D365 F&O 'Budget register entries' export (LedgerBudgetRegisterEntry entity)

The exports intentionally include the kind of mess a real D365 pull has:
inconsistent casing, duplicate rows, a few blank cost centers, stray whitespace,
and one entity with a currency in EUR vs USD.
"""
import pandas as pd
import numpy as np
import random

random.seed(42)
np.random.seed(42)

RAW = "/home/claude/fpa-project/data/d365_raw_exports"

ENTITIES = [
    {"LegalEntity": "DE01", "Name": "Germany GmbH Berlin", "Currency": "EUR"},
    {"LegalEntity": "DE02", "Name": "Germany GmbH Munich", "Currency": "EUR"},
    {"LegalEntity": "US01", "Name": "US Holdings Inc",     "Currency": "USD"},
]

MONTHS = pd.date_range("2025-01-01", "2025-12-01", freq="MS")

# --- Chart of Accounts (MainAccount entity export) ---
coa = [
    ("4000", "Revenue - Product Sales", "Revenue", "P&L"),
    ("4010", "Revenue - Services",       "Revenue", "P&L"),
    ("5000", "COGS - Materials",         "COGS",    "P&L"),
    ("5010", "COGS - Freight",           "COGS",    "P&L"),
    ("6000", "Salaries & Wages",         "OpEx",    "P&L"),
    ("6010", "Marketing Expense",        "OpEx",    "P&L"),
    ("6020", "IT & Software",            "OpEx",    "P&L"),
    ("6030", "Facilities & Rent",        "OpEx",    "P&L"),
    ("6040", "Professional Fees",        "OpEx",    "P&L"),
    ("7000", "Depreciation & Amortization","OpEx",  "P&L"),
    ("8000", "Interest Expense",         "NonOp",   "P&L"),
    ("1000", "Cash & Cash Equivalents",  "Asset",   "BS"),
    ("1200", "Accounts Receivable",      "Asset",   "BS"),
    ("1500", "Fixed Assets, net",        "Asset",   "BS"),
    ("2000", "Accounts Payable",         "Liability","BS"),
    ("2100", "Accrued Liabilities",      "Liability","BS"),
    ("2500", "Long-term Debt",           "Liability","BS"),
    ("3000", "Share Capital",            "Equity",  "BS"),
    ("3900", "Retained Earnings",        "Equity",  "BS"),
]
coa_df = pd.DataFrame(coa, columns=["MainAccount", "AccountName", "AccountCategory", "Statement"])
coa_df.to_csv(f"{RAW}/chart_of_accounts_export.csv", index=False)

# --- Cost Center master (financial dimension export) ---
cost_centers = [
    ("CC100", "Sales & Marketing"),
    ("CC200", "Operations"),
    ("CC300", "G&A"),
    ("CC400", "R&D / Product"),
    ("cc500 ", "Customer Support"),   # messy casing/whitespace on purpose
]
cc_df = pd.DataFrame(cost_centers, columns=["CostCenter", "CostCenterName"])
cc_df.to_csv(f"{RAW}/cost_center_master_export.csv", index=False)

# --- GL Trial Balance export (one row per entity/month/account/cost center) ---
rows = []
base_revenue = {"DE01": 480000, "DE02": 260000, "US01": 610000}
for ent in ENTITIES:
    le = ent["LegalEntity"]
    for i, month in enumerate(MONTHS):
        seasonal = 1 + 0.08 * np.sin(i / 2) + np.random.normal(0, 0.03)
        rev_total = base_revenue[le] * seasonal * (1 + 0.015 * i)  # slight growth trend
        for acc, name, cat, stmt in coa:
            if stmt == "BS":
                continue  # BS balances generated separately below
            if cat == "Revenue":
                cc = ""  # revenue typically not cost-center tagged in this model
                amt = rev_total * (0.7 if acc == "4000" else 0.3)
            elif cat == "COGS":
                cc = "CC200"
                amt = rev_total * (0.32 if acc == "5000" else 0.06)
            elif cat == "OpEx":
                weight = {"6000": 0.16, "6010": 0.05, "6020": 0.025,
                          "6030": 0.02, "6040": 0.015, "7000": 0.02}[acc]
                cc_map = {"6000": "CC300", "6010": "CC100", "6020": "CC400",
                          "6030": "CC300", "6040": "CC300", "7000": "CC200"}
                cc = cc_map[acc]
                amt = rev_total * weight * np.random.normal(1, 0.05)
            elif cat == "NonOp":
                cc = ""
                amt = rev_total * 0.01
            else:
                continue
            rows.append({
                "LegalEntity ": le if random.random() > 0.05 else le.lower(),  # messy col/casing
                "Period": month.strftime("%Y-%m"),
                "MainAccount": acc,
                "CostCenter": cc,
                "Amount": round(amt, 2),
                "Currency": ent["Currency"],
            })

tb_df = pd.DataFrame(rows)
# inject a few duplicate rows and blank cost centers, like a real messy export
dupes = tb_df.sample(15, random_state=1)
tb_df = pd.concat([tb_df, dupes], ignore_index=True)
tb_df.to_csv(f"{RAW}/gl_trial_balance_export.csv", index=False)

# --- Balance Sheet closing balances export ---
# Built to genuinely articulate with the P&L above: cash rolls forward from Net Income
# (in USD, same FX table used downstream) plus working-capital and CapEx/debt assumptions,
# so the Cash Flow statement built later in Excel ties out exactly to the Balance Sheet.
FX = {m: r for m, r in zip(
    [mo.strftime("%Y-%m") for mo in MONTHS],
    [1.09, 1.08, 1.09, 1.07, 1.08, 1.10, 1.09, 1.08, 1.07, 1.09, 1.10, 1.11]
)}
tmp = tb_df.drop_duplicates().rename(columns=lambda c: c.strip()).copy()
tmp["LegalEntity"] = tmp["LegalEntity"].str.upper()
tmp = tmp.merge(coa_df[["MainAccount", "AccountCategory"]], on="MainAccount", how="left")
ccy_map = {e["LegalEntity"]: e["Currency"] for e in ENTITIES}
tmp["AmountUSD"] = tmp.apply(
    lambda row: row["Amount"] if ccy_map[row["LegalEntity"]] == "USD" else round(row["Amount"] * FX[row["Period"]], 2),
    axis=1)
piv = tmp.pivot_table(index=["LegalEntity", "Period"], columns="MainAccount", values="AmountUSD", aggfunc="sum", fill_value=0)

bs_rows = []
for ent in ENTITIES:
    le = ent["LegalEntity"]
    cash = 250000 * (FX["2025-01"] if ent["Currency"] == "EUR" else 1) if le != "US01" else 300000
    ar = 180000 * (FX["2025-01"] if ent["Currency"] == "EUR" else 1) if le != "US01" else 210000
    fa = 900000 * (FX["2025-01"] if ent["Currency"] == "EUR" else 1) if le != "US01" else 1050000
    ap = 140000 * (FX["2025-01"] if ent["Currency"] == "EUR" else 1) if le != "US01" else 160000
    accr = 60000 * (FX["2025-01"] if ent["Currency"] == "EUR" else 1) if le != "US01" else 68000
    debt = 500000 * (FX["2025-01"] if ent["Currency"] == "EUR" else 1) if le != "US01" else 560000
    equity_capital = 400000 * (FX["2025-01"] if ent["Currency"] == "EUR" else 1) if le != "US01" else 450000
    for i, month in enumerate(MONTHS):
        per = month.strftime("%Y-%m")
        row = piv.loc[(le, per)] if (le, per) in piv.index else pd.Series(dtype=float)
        rev = row.get("4000", 0) + row.get("4010", 0)
        cogs = row.get("5000", 0) + row.get("5010", 0)
        opex = row.get("6000", 0) + row.get("6010", 0) + row.get("6020", 0) + row.get("6030", 0) + row.get("6040", 0)
        da = row.get("7000", 0)
        interest = row.get("8000", 0)
        ni = rev - cogs - opex - da - interest

        d_ar = np.random.normal(2000, 3000)
        d_ap = np.random.normal(1500, 2500)
        d_accr = np.random.normal(500, 1000)
        capex = da * 1.1  # capex roughly replaces depreciation, modest net PP&E growth
        debt_chg = -debt * 0.004  # slow scheduled amortization

        ar += d_ar
        ap += d_ap
        accr += d_accr
        fa += capex - da
        debt += debt_chg
        cash += ni + da - d_ar + d_ap + d_accr - capex + debt_chg

        retained_earnings = cash + ar + fa - ap - accr - debt - equity_capital
        for acc, amt in [("1000", cash), ("1200", ar), ("1500", fa),
                         ("2000", ap), ("2100", accr), ("2500", debt),
                         ("3000", equity_capital), ("3900", retained_earnings)]:
            bs_rows.append({
                "LegalEntity": le,
                "Period": per,
                "MainAccount": acc,
                "Amount": round(amt, 2),
                "Currency": "USD",  # already expressed in USD; downstream FX step is a no-op for this file
            })
bs_df = pd.DataFrame(bs_rows)
bs_df.to_csv(f"{RAW}/gl_balance_sheet_export.csv", index=False)

# --- Budget register entries export (annual budget, monthly phased) ---
budget_rows = []
for ent in ENTITIES:
    le = ent["LegalEntity"]
    for acc, name, cat, stmt in coa:
        if stmt != "P&L":
            continue
        annual = {"Revenue": base_revenue[le] * 12 * 1.05,
                   "COGS": base_revenue[le] * 12 * 0.35,
                   "OpEx": base_revenue[le] * 12 * 0.22,
                   "NonOp": base_revenue[le] * 12 * 0.01}[cat] / (2 if cat in ("Revenue",) and acc == "4010" else 1)
        monthly = annual / 12 / (3 if cat == "OpEx" else 1)
        for month in MONTHS:
            budget_rows.append({
                "LegalEntity": le,
                "Period": month.strftime("%Y-%m"),
                "MainAccount": acc,
                "BudgetAmount": round(monthly * np.random.normal(1, 0.02), 2),
            })
budget_df = pd.DataFrame(budget_rows)
budget_df.to_csv(f"{RAW}/budget_register_export.csv", index=False)

print("D365-style raw exports generated:")
for f in ["chart_of_accounts_export.csv", "cost_center_master_export.csv",
          "gl_trial_balance_export.csv", "gl_balance_sheet_export.csv",
          "budget_register_export.csv"]:
    df = pd.read_csv(f"{RAW}/{f}")
    print(f"  {f}: {len(df)} rows")
