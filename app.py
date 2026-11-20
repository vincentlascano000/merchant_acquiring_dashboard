"""
Merchant Acquiring Dashboard (PF vs DM)
Run: streamlit run app.py
Requirements: streamlit, pandas, plotly
"""
import calendar
from datetime import date

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Merchant Acquiring Dashboard", layout="wide")

DATA_PATH = "MACQ - MIS_2026-11-19-0907 (1).csv"   # put the CSV in the repo root (or change path)
TAG_COLORS = {"Direct Merchant": "#F28C28", "Payment Facilitator": "#1F4E79"}


# ---------------------------------------------------------------- DATA
@st.cache_data
def load_data(src):
    try:
        df = pd.read_csv(src, encoding="utf-8")
    except UnicodeDecodeError:
        df = pd.read_csv(src, encoding="cp1252")
    df["YEAR_MONTH"] = df["YEAR_MONTH"].astype(str)
    # helper columns for weighted rates (credits carry negative volume -> net them)
    df["MDR_X_VOL"] = df["MDR"] * df["TOTAL_VOLUME"]
    df["IC_X_VOL"] = df["INTERCHANGE_RATE"] * df["TOTAL_VOLUME"]
    df["INTERCHANGE_PHP"] = -df["INTERCHANGE_EXPENSE_PHP"]          # show as positive cost
    return df


def summarize(d, by):
    """Aggregate volume/count + weighted & simple-average rates."""
    g = d.groupby(by, dropna=False).agg(
        VOLUME=("TOTAL_VOLUME", "sum"),
        TXN_COUNT=("TOTAL_COUNT", "sum"),
        MDR_X_VOL=("MDR_X_VOL", "sum"),
        IC_X_VOL=("IC_X_VOL", "sum"),
        INTERCHANGE_PHP=("INTERCHANGE_PHP", "sum"),
        BLENDED_MDR=("MDR", "mean"),                 # simple (unweighted) average
        BLENDED_IC=("INTERCHANGE_RATE", "mean"),
    ).reset_index()
    g["WEIGHTED_MDR"] = g["MDR_X_VOL"] / g["VOLUME"].where(g["VOLUME"] != 0)
    g["WEIGHTED_IC"] = g["IC_X_VOL"] / g["VOLUME"].where(g["VOLUME"] != 0)
    g["TICKET_SIZE"] = g["VOLUME"] / g["TXN_COUNT"].where(g["TXN_COUNT"] != 0)
    return g.drop(columns=["MDR_X_VOL", "IC_X_VOL"])


def run_rate(g, factor):
    """Scale flow metrics (volume, count, interchange PHP) to full month. Rates are untouched."""
    g = g.copy()
    for c in ["VOLUME", "TXN_COUNT", "INTERCHANGE_PHP"]:
        g[c] = g[c] * factor
    return g


def prev_month(ym, n=1):
    p = pd.Period(ym, "M") - n
    return str(p)


def pct(a, b):
    return (a / b - 1) if b not in (0, None) and pd.notna(b) else None


def fmt_php(x):
    if pd.isna(x):
        return "-"
    for div, s in [(1e9, "B"), (1e6, "M"), (1e3, "K")]:
        if abs(x) >= div:
            return f"₱{x/div:,.2f}{s}"
    return f"₱{x:,.0f}"


# ---------------------------------------------------------------- SIDEBAR
st.sidebar.title("Settings")
upload = st.sidebar.file_uploader("Upload MIS CSV (optional)", type="csv")
df = load_data(upload if upload is not None else DATA_PATH)

months = sorted(df["YEAR_MONTH"].unique())
cur_m = st.sidebar.selectbox("Current month", months, index=len(months) - 1)
cur_p = pd.Period(cur_m, "M")
days_in_month = calendar.monthrange(cur_p.year, cur_p.month)[1]
as_of = st.sidebar.date_input(
    "Data as of", value=date(cur_p.year, cur_p.month, min(14, days_in_month)),
    min_value=date(cur_p.year, cur_p.month, 1),
    max_value=date(cur_p.year, cur_p.month, days_in_month))
rr_factor = days_in_month / as_of.day
st.sidebar.caption(f"Run-rate factor = {days_in_month} / {as_of.day} = **{rr_factor:.3f}x**")

include_credits = st.sidebar.checkbox("Net off credit (refund) transactions", value=True)
if not include_credits:
    df = df[df["TRXN_CODE"] == "Sale Transaction"]

mom_m, yoy_m = prev_month(cur_m, 1), prev_month(cur_m, 12)


def period_table(d, by):
    """Current (run-rated), prior month, same month LY + MoM/YoY deltas."""
    by = by if isinstance(by, list) else [by]
    cur = run_rate(summarize(d[d.YEAR_MONTH == cur_m], by), rr_factor)
    mom = summarize(d[d.YEAR_MONTH == mom_m], by)
    yoy = summarize(d[d.YEAR_MONTH == yoy_m], by)
    t = cur.merge(mom, on=by, how="outer", suffixes=("", "_PM")) \
           .merge(yoy.add_suffix("_LY").rename(columns={f"{b}_LY": b for b in by}), on=by, how="outer")
    for m in ["VOLUME", "TXN_COUNT", "INTERCHANGE_PHP"]:
        t[f"{m}_MoM%"] = t[m] / t[f"{m}_PM"] - 1
        t[f"{m}_YoY%"] = t[m] / t[f"{m}_LY"] - 1
    for m in ["WEIGHTED_MDR", "BLENDED_MDR", "WEIGHTED_IC", "BLENDED_IC"]:
        t[f"{m}_MoM_bps"] = (t[m] - t[f"{m}_PM"]) * 1e4
        t[f"{m}_YoY_bps"] = (t[m] - t[f"{m}_LY"]) * 1e4
    return t


def kpi_row(t):
    """Big-number cards for a single-row period table."""
    r = t.iloc[0]
    c = st.columns(6)
    c[0].metric("Volume (RR)", fmt_php(r.VOLUME), f"{r['VOLUME_MoM%']:+.1%} MoM | {r['VOLUME_YoY%']:+.1%} YoY")
    c[1].metric("Txn Count (RR)", f"{r.TXN_COUNT:,.0f}", f"{r['TXN_COUNT_MoM%']:+.1%} MoM | {r['TXN_COUNT_YoY%']:+.1%} YoY")
    c[2].metric("Weighted MDR", f"{r.WEIGHTED_MDR:.2%}", f"{r.WEIGHTED_MDR_MoM_bps:+.1f} bps MoM | {r.WEIGHTED_MDR_YoY_bps:+.1f} YoY")
    c[3].metric("Blended MDR (simple avg)", f"{r.BLENDED_MDR:.2%}", f"{r.BLENDED_MDR_MoM_bps:+.1f} bps MoM | {r.BLENDED_MDR_YoY_bps:+.1f} YoY")
    c[4].metric("Weighted Interchange", f"{r.WEIGHTED_IC:.2%}", f"{r.WEIGHTED_IC_MoM_bps:+.1f} bps MoM | {r.WEIGHTED_IC_YoY_bps:+.1f} YoY", delta_color="inverse")
    c[5].metric("Interchange (RR)", fmt_php(r.INTERCHANGE_PHP), f"{r['INTERCHANGE_PHP_MoM%']:+.1%} MoM | {r['INTERCHANGE_PHP_YoY%']:+.1%} YoY", delta_color="inverse")


def trend_df(d, by=None):
    """Monthly trend; current month replaced by run-rate value."""
    keys = ["YEAR_MONTH"] + ([by] if by else [])
    g = summarize(d, keys)
    m = g.YEAR_MONTH == cur_m
    g.loc[m, ["VOLUME", "TXN_COUNT", "INTERCHANGE_PHP"]] *= rr_factor
    g["LABEL"] = g.YEAR_MONTH.where(~m, g.YEAR_MONTH + " (RR)")
    return g[g.YEAR_MONTH <= cur_m]


DISPLAY_COLS = {
    "VOLUME": "Volume (RR)", "VOLUME_MoM%": "Vol MoM", "VOLUME_YoY%": "Vol YoY",
    "TXN_COUNT": "Txn (RR)", "TXN_COUNT_MoM%": "Txn MoM", "TXN_COUNT_YoY%": "Txn YoY",
    "WEIGHTED_MDR": "Wtd MDR", "WEIGHTED_MDR_MoM_bps": "Wtd MDR MoM (bps)", "WEIGHTED_MDR_YoY_bps": "Wtd MDR YoY (bps)",
    "BLENDED_MDR": "Blended MDR",
    "WEIGHTED_IC": "Wtd IC", "WEIGHTED_IC_MoM_bps": "Wtd IC MoM (bps)", "WEIGHTED_IC_YoY_bps": "Wtd IC YoY (bps)",
    "BLENDED_IC": "Blended IC",
    "INTERCHANGE_PHP": "Interchange ₱ (RR)", "INTERCHANGE_PHP_MoM%": "IC ₱ MoM", "INTERCHANGE_PHP_YoY%": "IC ₱ YoY",
}


def show_table(t, keys):
    keys = keys if isinstance(keys, list) else [keys]
    out = t[keys + list(DISPLAY_COLS)].rename(columns=DISPLAY_COLS)
    pct_cols = [v for k, v in DISPLAY_COLS.items() if k.endswith("%")]
    rate_cols = ["Wtd MDR", "Blended MDR", "Wtd IC", "Blended IC"]
    bps_cols = [v for k, v in DISPLAY_COLS.items() if k.endswith("bps")]
    st.dataframe(
        out.style.format({**{c: "{:+.1%}" for c in pct_cols}, **{c: "{:.2%}" for c in rate_cols},
                          **{c: "{:+.1f}" for c in bps_cols},
                          "Volume (RR)": "{:,.0f}", "Txn (RR)": "{:,.0f}", "Interchange ₱ (RR)": "{:,.0f}"},
                         na_rep="-"),
        width="stretch", hide_index=True)


# ---------------------------------------------------------------- HEADER
st.title("Merchant Acquiring Dashboard")
st.caption(f"Current: **{cur_m}** (MTD as of {as_of:%b %d, %Y}, run-rated ×{rr_factor:.2f}) · "
           f"MoM vs **{mom_m}** · YoY vs **{yoy_m}** · Rates in bps deltas")

tab1, tab2, tab3 = st.tabs(["1 · PF vs DM Overview", "2 · Category Mix", "3 · Merchant Level"])

# ================================================================ TAB 1
with tab1:
    st.subheader("Total Portfolio")
    kpi_row(period_table(df.assign(ALL="Total"), "ALL"))

    st.subheader("By Segment")
    seg = period_table(df, "PF_DM_TAG")
    for _, r in seg.iterrows():
        st.markdown(f"**{r.PF_DM_TAG}**")
        kpi_row(seg[seg.PF_DM_TAG == r.PF_DM_TAG].reset_index(drop=True))

    tr = trend_df(df, "PF_DM_TAG")
    c1, c2 = st.columns(2)
    c1.plotly_chart(px.bar(tr, x="LABEL", y="VOLUME", color="PF_DM_TAG", color_discrete_map=TAG_COLORS,
                           title="Monthly Volume (current month run-rated)"), width="stretch")
    c2.plotly_chart(px.bar(tr, x="LABEL", y="TXN_COUNT", color="PF_DM_TAG", color_discrete_map=TAG_COLORS,
                           title="Monthly Transaction Count"), width="stretch")
    c3, c4 = st.columns(2)
    c3.plotly_chart(px.line(tr, x="LABEL", y="WEIGHTED_MDR", color="PF_DM_TAG", markers=True,
                            color_discrete_map=TAG_COLORS, title="Weighted MDR").update_yaxes(tickformat=".2%"),
                    width="stretch")
    c4.plotly_chart(px.line(tr, x="LABEL", y="WEIGHTED_IC", color="PF_DM_TAG", markers=True,
                            color_discrete_map=TAG_COLORS, title="Weighted Interchange Rate").update_yaxes(tickformat=".2%"),
                    width="stretch")
    show_table(seg, "PF_DM_TAG")

# ================================================================ TAB 2
with tab2:
    seg_pick = st.radio("Segment", ["All", "Direct Merchant", "Payment Facilitator"], horizontal=True, key="cat_seg")
    d2 = df if seg_pick == "All" else df[df.PF_DM_TAG == seg_pick]
    cat = period_table(d2, "INDUSTRY").sort_values("VOLUME", ascending=False)
    cat["VOL_SHARE"] = cat.VOLUME / cat.VOLUME.sum()
    cat["TXN_SHARE"] = cat.TXN_COUNT / cat.TXN_COUNT.sum()
    cat["VOL_SHARE_PM"] = cat.VOLUME_PM / cat.VOLUME_PM.sum()
    cat["SHARE_CHG_PP"] = (cat.VOL_SHARE - cat.VOL_SHARE_PM) * 100

    c1, c2 = st.columns(2)
    c1.plotly_chart(px.pie(cat, names="INDUSTRY", values="VOLUME", hole=0.45,
                           title=f"Volume Share – {cur_m} (RR)").update_traces(textinfo="percent"),
                    width="stretch")
    c2.plotly_chart(px.pie(cat, names="INDUSTRY", values="TXN_COUNT", hole=0.45,
                           title=f"Transaction Count Share – {cur_m} (RR)").update_traces(textinfo="percent"),
                    width="stretch")

    # share trend (100% stacked) - top 8 categories, rest = Others
    tr2 = trend_df(d2, "INDUSTRY")
    top = cat.INDUSTRY.head(8).tolist()
    tr2["CAT"] = tr2.INDUSTRY.where(tr2.INDUSTRY.isin(top), "All other categories")
    tr2 = tr2.groupby(["LABEL", "CAT"], as_index=False).VOLUME.sum()
    tr2["SHARE"] = tr2.VOLUME / tr2.groupby("LABEL").VOLUME.transform("sum")
    st.plotly_chart(px.bar(tr2, x="LABEL", y="SHARE", color="CAT", title="Volume Share Trend (Top 8 categories)")
                    .update_yaxes(tickformat=".0%"), width="stretch")

    st.plotly_chart(px.bar(cat.head(15), x="SHARE_CHG_PP", y="INDUSTRY", orientation="h",
                           title=f"Volume Share Change vs {mom_m} (pp)")
                    .update_yaxes(autorange="reversed"), width="stretch")

    out = cat.copy()
    show_table(out.assign(**{"Vol Share": out.VOL_SHARE}), ["INDUSTRY", "Vol Share"])

# ================================================================ TAB 3
with tab3:
    seg3 = st.radio("Segment", ["Payment Facilitator", "Direct Merchant"], horizontal=True, key="m_seg")
    d3 = df[df.PF_DM_TAG == seg3]
    mer = period_table(d3, "PF_DM_NAME").sort_values("VOLUME", ascending=False)
    mer["VOL_SHARE"] = mer.VOLUME / mer.VOLUME.sum()
    top_n = st.slider("Top N merchants", 5, 30, 10)
    topm = mer.head(top_n)

    c1, c2 = st.columns(2)
    c1.plotly_chart(px.bar(topm, x="VOLUME", y="PF_DM_NAME", orientation="h", text=topm.VOL_SHARE.map("{:.1%}".format),
                           title=f"Top {top_n} {seg3}s – Volume (RR) & Share", color_discrete_sequence=[TAG_COLORS[seg3]])
                    .update_yaxes(autorange="reversed"), width="stretch")
    mv = topm.melt(id_vars="PF_DM_NAME", value_vars=["VOLUME_MoM%", "VOLUME_YoY%"], var_name="Comparison", value_name="Growth")
    c2.plotly_chart(px.bar(mv, x="Growth", y="PF_DM_NAME", color="Comparison", barmode="group", orientation="h",
                           title="Volume Growth – MoM vs YoY").update_xaxes(tickformat=".0%")
                    .update_yaxes(autorange="reversed"), width="stretch")

    st.plotly_chart(px.scatter(topm, x="WEIGHTED_IC", y="WEIGHTED_MDR", size="VOLUME", hover_name="PF_DM_NAME",
                               title="Weighted MDR vs Weighted Interchange (bubble = volume)")
                    .update_xaxes(tickformat=".2%").update_yaxes(tickformat=".2%"), width="stretch")

    show_table(mer.assign(**{"Vol Share": mer.VOL_SHARE}), ["PF_DM_NAME", "Vol Share"])

    # ---- drill-down into a single merchant / PF
    st.divider()
    pick = st.selectbox(f"Drill down: {seg3}", mer.PF_DM_NAME.tolist())
    dd = d3[d3.PF_DM_NAME == pick]
    kpi_row(period_table(dd.assign(ALL=pick), "ALL"))
    trd = trend_df(dd)
    c1, c2 = st.columns(2)
    c1.plotly_chart(px.bar(trd, x="LABEL", y="VOLUME", title=f"{pick} – Monthly Volume",
                           color_discrete_sequence=[TAG_COLORS[seg3]]), width="stretch")
    c2.plotly_chart(px.line(trd, x="LABEL", y=["WEIGHTED_MDR", "WEIGHTED_IC"], markers=True,
                            title=f"{pick} – Weighted MDR vs Interchange").update_yaxes(tickformat=".2%"),
                    width="stretch")
    if seg3 == "Payment Facilitator":
        st.markdown(f"**Sub-merchants under {pick} (top 20 by run-rated volume)**")
        sub = period_table(dd, "MERCHANT_NAME_ORIG").sort_values("VOLUME", ascending=False).head(20)
        show_table(sub, "MERCHANT_NAME_ORIG")
