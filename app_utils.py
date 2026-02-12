import os
import glob
from typing import Dict, Tuple, Optional, List

import numpy as np
import pandas as pd
import streamlit as st


# ---- Workbook config ----
SRC_DIR = "src"

SHEET_MASTER = "Master_Proforma"
SHEET_CASHFLOW = "Cashflow_View"
SHEET_REV_MIX = "Revenue_Forecast_Model"


# ----------------------------
# File handling
# ----------------------------
def find_newest_excel(src_dir: str) -> str:
    pattern = os.path.join(src_dir, "*.xlsx")
    files = glob.glob(pattern)
    if not files:
        raise FileNotFoundError(f"No .xlsx files found in ./{src_dir}/")
    files.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return files[0]


@st.cache_data(show_spinner=False)
def read_sheet_raw(path: str, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(path, sheet_name=sheet_name, header=None)


# ----------------------------
# Parsing helpers
# ----------------------------
def _find_header_row_index(df_raw: pd.DataFrame) -> int:
    for i in range(min(40, len(df_raw))):
        v = df_raw.iloc[i, 0]
        if isinstance(v, str) and v.strip().lower() == "category":
            return i
    raise ValueError("Could not find header row labeled 'Category'.")


def _parse_month_headers(df_raw: pd.DataFrame, header_idx: int) -> List[pd.Timestamp]:
    header_row = df_raw.iloc[header_idx, :].tolist()
    month_cells = header_row[1:]

    while month_cells and (pd.isna(month_cells[-1]) or str(month_cells[-1]).strip() == ""):
        month_cells.pop()

    if month_cells and isinstance(month_cells[-1], str) and month_cells[-1].strip().lower() == "total":
        month_cells = month_cells[:-1]

    months: List[pd.Timestamp] = []
    for m in month_cells:
        if pd.isna(m):
            continue
        s = str(m).strip()
        dt = pd.to_datetime(s, format="%b-%y", errors="coerce")
        if pd.isna(dt):
            dt = pd.to_datetime(s, errors="coerce")
        if pd.notna(dt):
            months.append(pd.to_datetime(dt))

    return months


def extract_row_timeseries(df_raw: pd.DataFrame, row_label: str) -> pd.DataFrame:
    header_idx = _find_header_row_index(df_raw)
    months = _parse_month_headers(df_raw, header_idx)

    col0 = df_raw.iloc[:, 0].astype(str).str.strip().str.lower()
    label = row_label.strip().lower()

    exact = df_raw[col0 == label]
    if not exact.empty:
        row = exact.iloc[0]
    else:
        contains = df_raw[col0.str.contains(label, na=False)]
        if contains.empty:
            raise ValueError(f"Row label not found: '{row_label}'")
        row = contains.iloc[0]

    values = row.iloc[1 : 1 + len(months)]
    out = pd.DataFrame(
        {"month": pd.to_datetime(months), "value": pd.to_numeric(values, errors="coerce")}
    )
    out = out.dropna(subset=["month"]).sort_values("month").reset_index(drop=True)
    return out


def extract_row_any(df_raw: pd.DataFrame, labels: List[str]) -> pd.DataFrame:
    last_err = None
    for lab in labels:
        try:
            return extract_row_timeseries(df_raw, lab)
        except Exception as e:
            last_err = e
    raise ValueError(f"None of these row labels were found: {labels}. Last error: {last_err}")


def merge_series_on_month(series_map: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    out: Optional[pd.DataFrame] = None
    for key, s in series_map.items():
        s2 = s.rename(columns={"value": key})
        out = s2 if out is None else out.merge(s2, on="month", how="outer")
    out = out.sort_values("month").reset_index(drop=True)
    return out


# ----------------------------
# Canonical model builder
# ----------------------------
@st.cache_data(show_spinner=True)
def build_canonical(path: str) -> Tuple[pd.DataFrame, Dict[str, str]]:
    df_master = read_sheet_raw(path, SHEET_MASTER)
    df_cash = read_sheet_raw(path, SHEET_CASHFLOW)
    df_mix = read_sheet_raw(path, SHEET_REV_MIX)

    # Master_Proforma
    master_series = {
        "revenue_total": extract_row_any(df_master, ["Revenue"]),
        # NEW: Variable Costs renamed in your file
        "variable_costs": extract_row_any(df_master, ["AI / Usage COGS", "Variable Costs", "COGS"]),
        "gross_profit": extract_row_any(df_master, ["Gross Profit"]),
        "gross_margin": extract_row_any(df_master, ["Gross Profit %", "Gross Margin %"]),
        "expenses_total": extract_row_any(df_master, ["Total Expenses"]),
        "people_cost": extract_row_any(df_master, ["People"]),
        "tech_cost": extract_row_any(df_master, ["Tech / AI / Cloud (Fixed)", "Technology"]),
        "marketing_cost": extract_row_any(df_master, ["Marketing & Growth", "Marketing"]),
        "ops_cost": extract_row_any(df_master, ["Legal / Ops / Buffer", "Legal"]),
        "ebitda": extract_row_any(df_master, ["EBITDA"]),
        "ebitda_margin": extract_row_any(df_master, ["EBITDA %"]),
    }
    df = merge_series_on_month(master_series)

    # Cashflow_View
    cash_ending = extract_row_any(df_cash, ["Ending Cash"]).rename(columns={"value": "cash_ending"})
    df = df.merge(cash_ending, on="month", how="left")

    # Revenue_Forecast_Model (mix + implied users)
    mix_series = {
        "consumer_revenue": extract_row_any(df_mix, ["Consumer Revenue"]),
        "business_revenue": extract_row_any(df_mix, ["Business Revenue"]),
        "marketplace_revenue": extract_row_any(df_mix, ["Marketplace Revenue"]),
        "total_monthly_revenue_mix": extract_row_any(df_mix, ["TOTAL MONTHLY REVENUE", "Total Monthly Revenue"]),
        "paying_consumers": extract_row_any(df_mix, ["Implied Paying Consumers"]),
        "business_accounts": extract_row_any(df_mix, ["Implied Business Accounts"]),
        "marketplace_sellers": extract_row_any(df_mix, ["Implied Marketplace Sellers"]),
        "consumer_arpu": extract_row_any(df_mix, ["Ask Lusso - Consumer ARPU"]),
        "business_arpu": extract_row_any(df_mix, ["Ask Lusso - Business ARPU"]),
        "marketplace_arpu": extract_row_any(df_mix, ["Marketplace - Seller ARPU"]),
    }
    df_mix_wide = merge_series_on_month(mix_series)
    df = df.merge(df_mix_wide, on="month", how="left")

    # Derived metrics
    df["burn"] = df["expenses_total"] - df["revenue_total"]
    df["paying_users_total"] = df[["paying_consumers", "business_accounts", "marketplace_sellers"]].sum(axis=1, min_count=1)

    df["arpu_blended"] = np.where(
        df["paying_users_total"] > 0,
        df["revenue_total"] / df["paying_users_total"],
        np.nan
    )

    # Normalize margins to percent if fraction-like
    for col in ["gross_margin", "ebitda_margin"]:
        if col in df.columns:
            s = df[col].dropna()
            if not s.empty:
                frac_like = (s.abs() <= 2).mean()
                if frac_like > 0.8:
                    df[col] = df[col] * 100.0

    # Breakeven (static)
    df["is_breakeven"] = (df["revenue_total"] >= df["expenses_total"]) | (df["ebitda"] >= 0)
    be = df.loc[df["is_breakeven"] == True, "month"].min()

    meta = {"breakeven_month": be.strftime("%b %Y") if pd.notna(be) else "Not reached in horizon"}
    df = df.sort_values("month").reset_index(drop=True)
    return df, meta
