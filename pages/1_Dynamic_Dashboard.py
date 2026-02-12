import os
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

from app_utils import find_newest_excel, build_canonical, SRC_DIR


st.set_page_config(page_title="Dynamic Dashboard", layout="wide")
st.title("Dynamic Dashboard")
st.caption("Scenario modeling: customize customer counts by segment and instantly see updated financials.")


# ---------- Load canonical ----------
excel_path = find_newest_excel(SRC_DIR)
df_base, meta = build_canonical(excel_path)
df_base = df_base.dropna(subset=["month"]).sort_values("month").reset_index(drop=True)


# ---------- Sidebar controls ----------
with st.sidebar:
    st.header("Scenario Controls")

    months = df_base["month"].tolist()
    min_m, max_m = months[0], months[-1]
    start_m, end_m = st.slider(
        "Month range",
        min_value=min_m.to_pydatetime(),
        max_value=max_m.to_pydatetime(),
        value=(min_m.to_pydatetime(), max_m.to_pydatetime()),
        format="MMM YYYY",
    )
    start_m = pd.to_datetime(start_m)
    end_m = pd.to_datetime(end_m)

    st.divider()
    st.subheader("Customer Overrides (key drivers)")

    # Pre-filled from Excel implied counts
    # You can override using multipliers (fast) + optional per-month table editor (precise)
    ask_mult = st.number_input("Ask Lusso customers multiplier", value=1.00, min_value=0.0, step=0.05)
    ent_mult = st.number_input("Enterprise customers multiplier", value=1.00, min_value=0.0, step=0.05)
    mkt_mult = st.number_input("Marketplace sellers multiplier", value=1.00, min_value=0.0, step=0.05)

    st.caption("Tip: 1.20 = +20% customers vs Excel projection.")

    st.divider()
    st.subheader("Economics")
    gross_margin_pct = st.number_input("Gross margin (%)", value=float(df_base["gross_margin"].dropna().iloc[-1]), min_value=-100.0, max_value=100.0, step=1.0)
    starting_cash = st.number_input("Starting cash ($)", value=float(df_base["cash_ending"].dropna().iloc[0]) if df_base["cash_ending"].notna().any() else 2_500_000.0, step=50_000.0)

    st.divider()
    show_editor = st.checkbox("Edit monthly customer counts (advanced)", value=False)


# ---------- Filter range ----------
df = df_base[(df_base["month"] >= start_m) & (df_base["month"] <= end_m)].copy().reset_index(drop=True)

# Base implied counts
df["ask_customers_base"] = df["paying_consumers"].fillna(0)
df["enterprise_customers_base"] = df["business_accounts"].fillna(0)
df["marketplace_sellers_base"] = df["marketplace_sellers"].fillna(0)

# Apply multipliers
df["ask_customers"] = (df["ask_customers_base"] * ask_mult).round(0)
df["enterprise_customers"] = (df["enterprise_customers_base"] * ent_mult).round(0)
df["marketplace_sellers"] = (df["marketplace_sellers_base"] * mkt_mult).round(0)

# Optional per-month edits
if show_editor:
    edit_df = df[["month", "ask_customers", "enterprise_customers", "marketplace_sellers"]].copy()
    edit_df["month"] = edit_df["month"].dt.strftime("%b %Y")
    st.info("Edit customer counts below. Changes recalc the scenario instantly.")
    edited = st.data_editor(
        edit_df,
        width="stretch",
        num_rows="fixed",
        column_config={
            "ask_customers": st.column_config.NumberColumn("Ask Lusso customers", min_value=0, step=1),
            "enterprise_customers": st.column_config.NumberColumn("Enterprise customers", min_value=0, step=1),
            "marketplace_sellers": st.column_config.NumberColumn("Marketplace sellers", min_value=0, step=1),
        },
    )
    # write back
    df["ask_customers"] = pd.to_numeric(edited["ask_customers"], errors="coerce").fillna(0)
    df["enterprise_customers"] = pd.to_numeric(edited["enterprise_customers"], errors="coerce").fillna(0)
    df["marketplace_sellers"] = pd.to_numeric(edited["marketplace_sellers"], errors="coerce").fillna(0)

# ARPUs from Excel
df["consumer_arpu"] = df["consumer_arpu"].fillna(method="ffill").fillna(0)
df["business_arpu"] = df["business_arpu"].fillna(method="ffill").fillna(0)
df["marketplace_arpu"] = df["marketplace_arpu"].fillna(method="ffill").fillna(0)

# Revenue recompute based on customers * ARPU
df["rev_ask"] = df["ask_customers"] * df["consumer_arpu"]
df["rev_ent"] = df["enterprise_customers"] * df["business_arpu"]
df["rev_mkt"] = df["marketplace_sellers"] * df["marketplace_arpu"]
df["revenue_scenario"] = df["rev_ask"] + df["rev_ent"] + df["rev_mkt"]

# Costs fixed from Excel (your preference)
df["expenses_scenario"] = df["expenses_total"].fillna(0)

# Gross profit + EBITDA scenario
gm = gross_margin_pct / 100.0
df["gross_profit_scenario"] = df["revenue_scenario"] * gm
df["variable_costs_scenario"] = df["revenue_scenario"] - df["gross_profit_scenario"]
df["ebitda_scenario"] = df["gross_profit_scenario"] - df["expenses_scenario"]

# Cash projection from starting cash
df["net_burn_scenario"] = df["expenses_scenario"] - df["revenue_scenario"]  # positive = burn
cash = []
cur = float(starting_cash)
for b in df["net_burn_scenario"].tolist():
    cur = cur - float(b)
    cash.append(cur)
df["cash_scenario"] = cash

# ARR
df["arr_scenario"] = df["revenue_scenario"] * 12.0


# ---------- KPIs ----------
c1, c2, c3, c4, c5 = st.columns(5)
last = df.iloc[-1]

c1.metric("Scenario Revenue (latest)", f"${last['revenue_scenario']:,.0f}")
c2.metric("Scenario EBITDA (latest)", f"${last['ebitda_scenario']:,.0f}")
c3.metric("Scenario ARR (latest)", f"${last['arr_scenario']:,.0f}")
c4.metric("Scenario Cash (latest)", f"${last['cash_scenario']:,.0f}")
c5.metric("Gross Margin (assumed)", f"{gross_margin_pct:.0f}%")

st.divider()


# ---------- Graphs (investor-friendly) ----------
# 1) Customers by segment
cust_long = df.melt(
    id_vars=["month"],
    value_vars=["ask_customers", "enterprise_customers", "marketplace_sellers"],
    var_name="Segment",
    value_name="Customers",
)
cust_long["Segment"] = cust_long["Segment"].replace({
    "ask_customers": "Ask Lusso",
    "enterprise_customers": "Enterprise",
    "marketplace_sellers": "Marketplace",
})
fig1 = px.bar(cust_long, x="month", y="Customers", color="Segment", barmode="stack",
              title="1) Customer Counts by Segment (Scenario)")
fig1.update_layout(hovermode="x unified", legend_title_text="")
fig1.update_yaxes(separatethousands=True)

# 2) Revenue mix scenario
rev_mix = df[["month", "rev_ask", "rev_ent", "rev_mkt"]].copy().rename(columns={
    "rev_ask": "Ask Lusso", "rev_ent": "Enterprise", "rev_mkt": "Marketplace"
})
rev_long = rev_mix.melt(id_vars=["month"], var_name="Stream", value_name="Revenue")
fig2 = px.area(rev_long, x="month", y="Revenue", color="Stream",
               title="2) Revenue Mix by Stream (Scenario)")
fig2.update_layout(hovermode="x unified", legend_title_text="")
fig2.update_yaxes(tickprefix="$", separatethousands=True)

# 3) Scenario vs Base revenue
fig3 = go.Figure()
fig3.add_trace(go.Scatter(x=df["month"], y=df["revenue_total"], mode="lines+markers", name="Base (Excel)"))
fig3.add_trace(go.Scatter(x=df["month"], y=df["revenue_scenario"], mode="lines+markers", name="Scenario"))
fig3.update_layout(title="3) Revenue: Base vs Scenario", hovermode="x unified", legend_title_text="")
fig3.update_yaxes(tickprefix="$", separatethousands=True)

# 4) EBITDA scenario
fig4 = px.bar(df, x="month", y="ebitda_scenario", title="4) EBITDA ($) (Scenario)")
fig4.update_layout(hovermode="x unified", legend_title_text="")
fig4.update_yaxes(tickprefix="$", separatethousands=True)

# 5) Cash runway scenario
fig5 = px.line(df, x="month", y="cash_scenario", markers=True, title="5) Cash Balance Over Time (Scenario)")
fig5.update_layout(hovermode="x unified", legend_title_text="")
fig5.update_yaxes(tickprefix="$", separatethousands=True)

# Layout
r1c1, r1c2 = st.columns(2)
r1c1.plotly_chart(fig1, width="stretch")
r1c2.plotly_chart(fig2, width="stretch")

r2c1, r2c2 = st.columns(2)
r2c1.plotly_chart(fig3, width="stretch")
r2c2.plotly_chart(fig4, width="stretch")

st.plotly_chart(fig5, width="stretch")

st.divider()

with st.expander("Show scenario table + download"):
    show = df[[
        "month",
        "ask_customers","enterprise_customers","marketplace_sellers",
        "consumer_arpu","business_arpu","marketplace_arpu",
        "revenue_scenario","expenses_scenario","ebitda_scenario","cash_scenario","arr_scenario"
    ]].copy()
    st.dataframe(show, width="stretch")
    csv = show.to_csv(index=False).encode("utf-8")
    st.download_button("Download scenario CSV", data=csv, file_name="dynamic_scenario.csv", mime="text/csv")
