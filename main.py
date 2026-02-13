# main.py
# Streamlit main page (Static V1) — robust to renamed rows / updated Excel
# Reads newest .xlsx from ./src and renders STATIC investor header + charts from Excel.

import os
import glob
from typing import Dict, Tuple, Optional, List

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go


# ----------------------------
# Streamlit config
# ----------------------------
st.set_page_config(page_title="Financial Dashboard", layout="wide")
st.title("Financial Dashboard")
st.caption("Internal financial performance & runway visibility (Static V1)")


# ----------------------------
# Workbook config
# ----------------------------
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
# Parsing helpers (robust)
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

    # Trim trailing empties
    while month_cells and (pd.isna(month_cells[-1]) or str(month_cells[-1]).strip() == ""):
        month_cells.pop()

    # Drop 'Total' if present
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

    # Prefer exact, then contains
    exact = df_raw[col0 == label]
    if not exact.empty:
        row = exact.iloc[0]
    else:
        contains = df_raw[col0.str.contains(label, na=False)]
        if contains.empty:
            raise ValueError(f"Row label not found: '{row_label}'")
        row = contains.iloc[0]

    values = row.iloc[1: 1 + len(months)]
    out = pd.DataFrame(
        {"month": pd.to_datetime(months), "value": pd.to_numeric(values, errors="coerce")}
    )
    out = out.dropna(subset=["month"]).sort_values("month").reset_index(drop=True)
    return out


def extract_row_any(df_raw: pd.DataFrame, labels: List[str]) -> pd.DataFrame:
    last_err: Optional[Exception] = None
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
    assert out is not None
    return out.sort_values("month").reset_index(drop=True)


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
        "variable_costs": extract_row_any(df_master, ["Variable Costs", "AI / Usage COGS", "COGS"]),
        "gross_profit": extract_row_any(df_master, ["Gross Profit"]),
        "gross_margin": extract_row_any(df_master, ["Gross Profit %", "Gross Margin %"]),
        "expenses_total": extract_row_any(df_master, ["Total Expenses"]),
        "people_cost": extract_row_any(df_master, ["People", "Direct Labor"]),  # fallback
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

    # Revenue_Forecast_Model
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

    # Derived metrics (kept for charts/table; NOT used for the investor header KPIs)
    df["burn"] = df["expenses_total"] - df["revenue_total"]
    df["paying_users_total"] = df[["paying_consumers", "business_accounts", "marketplace_sellers"]].sum(
        axis=1, min_count=1
    )
    df["arpu_blended"] = np.where(
        df["paying_users_total"] > 0,
        df["revenue_total"] / df["paying_users_total"],
        np.nan,
    )

    # Normalize margins to percent if fraction-like
    for col in ["gross_margin", "ebitda_margin"]:
        if col in df.columns:
            s = df[col].dropna()
            if not s.empty:
                frac_like = (s.abs() <= 2).mean()
                if frac_like > 0.8:
                    df[col] = df[col] * 100.0

    # Breakeven definition (static)
    df["is_breakeven"] = (df["revenue_total"] >= df["expenses_total"]) | (df["ebitda"] >= 0)
    be = df.loc[df["is_breakeven"] == True, "month"].min()
    meta = {"breakeven_month": be.strftime("%b %Y") if pd.notna(be) else "Not reached in horizon"}

    df = df.sort_values("month").reset_index(drop=True)
    return df, meta


# ----------------------------
# Plot helpers (safe lines)
# ----------------------------
def add_breakeven_vline(fig: go.Figure, df: pd.DataFrame) -> go.Figure:
    be = df.loc[df["is_breakeven"] == True, "month"].min()
    if pd.notna(be):
        be_dt = pd.to_datetime(be).to_pydatetime()
        fig.add_shape(
            type="line",
            x0=be_dt, x1=be_dt,
            y0=0, y1=1,
            xref="x", yref="paper",
            line=dict(width=2, dash="dash"),
        )
        fig.add_annotation(
            x=be_dt, y=1,
            xref="x", yref="paper",
            text=f"Breakeven ({pd.to_datetime(be).strftime('%b %Y')})",
            showarrow=False,
            yanchor="bottom",
            xanchor="left",
        )
    return fig


def add_min_cash_hline(fig: go.Figure, min_cash: float) -> go.Figure:
    fig.add_shape(
        type="line",
        x0=0, x1=1,
        y0=min_cash, y1=min_cash,
        xref="paper", yref="y",
        line=dict(width=1, dash="dot"),
    )
    fig.add_annotation(
        x=1, y=min_cash,
        xref="paper", yref="y",
        text=f"Min Cash (${min_cash:,.0f})",
        showarrow=False,
        xanchor="right",
        yanchor="bottom",
    )
    return fig


# ----------------------------
# Charts
# ----------------------------
def chart_cash_balance(df: pd.DataFrame, min_cash: float) -> go.Figure:
    fig = px.line(
        df, x="month", y="cash_ending",
        markers=True,
        title="1) Cash Balance Over Time",
        hover_data={"month": "|%b %Y", "cash_ending": ":$,.0f"},
    )
    fig.update_layout(hovermode="x unified", legend_title_text="")
    fig.update_yaxes(tickprefix="$", separatethousands=True)
    fig = add_min_cash_hline(fig, min_cash)
    fig = add_breakeven_vline(fig, df)
    return fig


def chart_burn_vs_revenue(df: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df["month"], y=df["expenses_total"], mode="lines+markers", name="Total Expenses"))
    fig.add_trace(go.Scatter(x=df["month"], y=df["revenue_total"], mode="lines+markers", name="Revenue"))
    fig.update_layout(title="2) Monthly Burn vs Revenue", hovermode="x unified", legend_title_text="")
    fig.update_yaxes(tickprefix="$", separatethousands=True)
    fig = add_breakeven_vline(fig, df)
    return fig


def chart_revenue_mix(df: pd.DataFrame) -> go.Figure:
    mix = df[["month", "consumer_revenue", "business_revenue", "marketplace_revenue"]].copy()
    mix = mix.rename(columns={"consumer_revenue": "Consumer", "business_revenue": "Business", "marketplace_revenue": "Marketplace"})
    mix_long = mix.melt(id_vars=["month"], var_name="Stream", value_name="Revenue")
    fig = px.area(
        mix_long, x="month", y="Revenue", color="Stream",
        title="3) Revenue Mix by Product Stream (Stacked)",
        hover_data={"month": "|%b %Y", "Revenue": ":$,.0f"},
    )
    fig.update_layout(hovermode="x unified", legend_title_text="")
    fig.update_yaxes(tickprefix="$", separatethousands=True)
    return fig


def chart_expense_breakdown(df: pd.DataFrame) -> go.Figure:
    exp = df[["month", "people_cost", "tech_cost", "marketing_cost", "ops_cost"]].copy()
    exp = exp.rename(columns={"people_cost": "People", "tech_cost": "Tech / AI / Cloud", "marketing_cost": "Marketing", "ops_cost": "Legal / Ops / Buffer"})
    exp_long = exp.melt(id_vars=["month"], var_name="Category", value_name="Cost")
    fig = px.bar(
        exp_long, x="month", y="Cost", color="Category",
        title="4) Expense Breakdown by Category (Stacked)",
        hover_data={"month": "|%b %Y", "Cost": ":$,.0f"},
    )
    fig.update_layout(barmode="stack", hovermode="x unified", legend_title_text="")
    fig.update_yaxes(tickprefix="$", separatethousands=True)
    return fig


def chart_gross_margin(df: pd.DataFrame) -> go.Figure:
    fig = px.line(
        df, x="month", y="gross_margin",
        markers=True,
        title="5) Gross Margin % Over Time",
        hover_data={"month": "|%b %Y", "gross_margin": ":.1f"},
    )
    fig.update_yaxes(title_text="Gross Margin (%)", ticksuffix="%")
    fig.update_layout(hovermode="x unified", legend_title_text="")
    return fig


def chart_ebitda_combo(df: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    fig.add_trace(go.Bar(x=df["month"], y=df["ebitda"], name="EBITDA ($)"))
    fig.add_trace(go.Scatter(x=df["month"], y=df["ebitda_margin"], name="EBITDA (%)", mode="lines+markers", yaxis="y2"))
    fig.update_layout(
        title="6) EBITDA ($) and EBITDA (%)",
        hovermode="x unified",
        legend_title_text="",
        yaxis=dict(title="EBITDA ($)", tickprefix="$", separatethousands=True),
        yaxis2=dict(title="EBITDA (%)", overlaying="y", side="right", ticksuffix="%"),
    )
    fig = add_breakeven_vline(fig, df)
    return fig


# ----------------------------
# Sidebar controls
# ----------------------------
with st.sidebar:
    st.header("Investor Header (Static)")

    cash_2026 = st.number_input(
        "Cash in hand (2026)",
        min_value=0.0,
        value=4_161_250.0,
        step=50_000.0,
        format="%.0f",
    )
    cash_2027 = st.number_input(
        "Cash in hand (2027)",
        min_value=0.0,
        value=19_811_813.0,
        step=50_000.0,
        format="%.0f",
    )

    # Static “story” tiles you can tweak fast before a meeting
    breakeven_target = st.text_input("Breakeven target (static label)", value="Nov 2026")
    revenue_model_label = st.text_input("Revenue model (static label)", value="Subscriptions + Marketplace")
    primary_kpi_label = st.text_input("Primary KPI (static label)", value="MRR growth")
    margin_focus_label = st.text_input("Margin focus (static label)", value="Gross margin expansion")

    st.caption("These tiles are static info (not calculated).")

    st.header("Data Source")
    st.write("Place your proforma Excel file inside `./src/`.")
    min_cash_threshold = st.number_input(
        "Min cash threshold (reference line)",
        min_value=0,
        value=500_000,
        step=50_000,
    )
    st.divider()
    st.caption("Use the left nav to open the Dynamic Dashboard page.")


# ----------------------------
# Load newest excel + build model
# ----------------------------
try:
    excel_path = find_newest_excel(SRC_DIR)
except Exception as e:
    st.error(str(e))
    st.stop()

st.caption(f"Loaded: `{os.path.basename(excel_path)}`")

try:
    df, meta = build_canonical(excel_path)
except Exception as e:
    st.error("Failed to build canonical model.")
    st.exception(e)
    st.stop()


# ----------------------------
# Static Investor Header KPIs (NOT computed)
# ----------------------------
def fmt_money(x: float) -> str:
    return f"${x:,.0f}"


st.caption("**Investor snapshot (static)** — not calculated from the Excel in this view.")
k1, k2, k3, k4, k5, k6 = st.columns(6)

k1.metric("Cash in hand (2026)", fmt_money(cash_2026))
k2.metric("Cash in hand (2027)", fmt_money(cash_2027))
k3.metric("Revenue model", revenue_model_label)
k4.metric("Primary KPI", primary_kpi_label)
k5.metric("Margin focus", margin_focus_label)
k6.metric("Breakeven target", breakeven_target)

st.divider()


# ----------------------------
# Render charts (still driven by Excel)
# ----------------------------
df_sorted = df.dropna(subset=["month"]).sort_values("month").reset_index(drop=True)

c1, c2 = st.columns(2)
c1.plotly_chart(chart_cash_balance(df_sorted, min_cash=min_cash_threshold), use_container_width=True)
c2.plotly_chart(chart_burn_vs_revenue(df_sorted), use_container_width=True)

c3, c4 = st.columns(2)
c3.plotly_chart(chart_revenue_mix(df_sorted), use_container_width=True)
c4.plotly_chart(chart_expense_breakdown(df_sorted), use_container_width=True)

c5, c6 = st.columns(2)
c5.plotly_chart(chart_gross_margin(df_sorted), use_container_width=True)
c6.plotly_chart(chart_ebitda_combo(df_sorted), use_container_width=True)

st.divider()


# ----------------------------
# Table + export
# ----------------------------
with st.expander("Show canonical monthly table + download"):
    show_cols = [
        "month",
        "cash_ending",
        "revenue_total",
        "expenses_total",
        "burn",
        "people_cost",
        "tech_cost",
        "marketing_cost",
        "ops_cost",
        "variable_costs",
        "gross_margin",
        "gross_profit",
        "ebitda",
        "ebitda_margin",
        "consumer_revenue",
        "business_revenue",
        "marketplace_revenue",
        "paying_users_total",
        "arpu_blended",
    ]
    existing = [c for c in show_cols if c in df_sorted.columns]
    st.dataframe(df_sorted[existing].copy(), use_container_width=True)

    csv = df_sorted[existing].to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download canonical CSV",
        data=csv,
        file_name="financial_dashboard_canonical.csv",
        mime="text/csv",
    )
