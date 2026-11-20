"""
Merchant Acquiring Dashboard (PF vs DM)
Run: streamlit run app.py
Requirements: streamlit, pandas, plotly

Run-rate rule:
  * ONLY the latest month in the data (the current, incomplete month) is run-rated.
  * Every prior month (incl. MoM / YoY comparison bases) = actual, simple summation.
"""
import calendar
from datetime import date

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Merchant Acquiring Dashboard", layout="wide")

DATA_PATH = "MACQ - MIS_2026-11-19-0907 (1).csv"
TAG_COLORS = {"Direct Merchant": "#F28C28", "Payment Facilitator": "#1F4E79"}
FLOW_COLS = ["VOLUME", "TXN_COUNT", "INTERCHANGE_PHP"]   # only these get run-rated; rates never do
PIE_TOP_N = 5
OTHERS_LABEL = "Others"
REQUIRED_COLS = ["YEAR_MONTH", "PF_DM_TAG", "PF_DM_NAME", "INDUSTRY", "TRXN_CODE", "TOTAL_VOLUME",
                 "TOTAL_COUNT", "MDR", "INTERCHANGE_RATE", "INTERCHANGE_EXPENSE_PHP"]


# ---------------------------------------------------------------- DATA
@st.cache_data
def load_data(src):
    try:
        df = pd.read_csv(src, encoding="utf-8")
    except UnicodeDecodeError:
        if hasattr(src, "seek"):
            src.seek(0)
        df = pd.read_csv(src, encoding="cp1252")
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"Uploaded file is missing columns: {missing}")
    df["YEAR_MONTH"] = df["YEAR_MONTH"].astype(str).str[:7]
    df["MDR_X_VOL"] = df["MDR"] * df["TOTAL_VOLUME"]
    df["IC_X_VOL"] = df["INTERCHANGE_RATE"] * df["TOTAL_VOLUME"]
    df["INTERCHANGE_PHP"] = -df["INTERCHANGE_EXPENSE_PHP"]
    return df


def summarize(d, by):
    """Actual (simple-sum) aggregation + weighted & simple-average rates."""
    g = d.groupby(by, dropna=False).agg(
        VOLUME=("TOTAL_VOLUME", "sum"),
        TXN_COUNT=("TOTAL_COUNT", "sum"),
        MDR_X_VOL=("MDR_X_VOL", "sum"),
        IC_X_VOL=("IC_X_VOL", "sum"),
        INTERCHANGE_PHP=("INTERCHANGE_PHP", "sum"),
        BLENDED_MDR=("MDR", "mean"),
        BLENDED_IC=("INTERCHANGE_RATE", "mean"),
    ).reset_index()
    g["WEIGHTED_MDR"] = g["MDR_X_VOL"] / g["VOLUME"].where(g["VOLUME"] != 0)
    g["WEIGHTED_IC"] = g["IC_X_VOL"] / g["VOLUME"].where(g["VOLUME"] != 0)
    return g.drop(columns=["MDR_X_VOL", "IC_X_VOL"])


def month_factor(ym):
    """Run-rate multiplier: >1 only for the latest (current) month, 1.0 for every other month."""
    return RR_FACTOR if ym == LATEST_M else 1.0


def prev_month(ym, n=1):
    return str(pd.Period(ym, "M") - n)


def top_n_others(d, value_col, n=PIE_TOP_N):
    """Keep top-n categories by value_col; everything else (incl. any native 'Others') -> 'Others'."""
    ranked = d[d["INDUSTRY"] != OTHERS_LABEL].sort_values(value_col, ascending=False)
    keep = ranked["INDUSTRY"].head(n).tolist()
    out = d.assign(INDUSTRY=d["INDUSTRY"].where(d["INDUSTRY"].isin(keep), OTHERS_LABEL))
    out = out.groupby("INDUSTRY", as_index=False)[value_col].sum()
    out["_ORDER"] = out["INDUSTRY"].eq(OTHERS_LABEL)          # Others last
    return out.sort_values(["_ORDER", value_col], ascending=[True, False]).drop(columns="_ORDER")


def pie(d, value_col, title):
    fig = px.pie(d, names="INDUSTRY", values=value_col, hole=0.45, title=title,
                 category_orders={"INDUSTRY": d["INDUSTRY"].tolist()})
    fig.update_traces(textinfo="percent+label", textposition="outside", sort=False)
    fig.update_layout(showlegend=False, margin=dict(t=60, b=20, l=20, r=20))
    return fig


def fmt_php(x):
    if pd.isna(x):
        return "-"
    for div, s in [(1e9, "B"), (1e6, "M"), (1e3, "K")]:
        if abs(x) >= div:
            return f"₱{x/div:,.2f}{s}"
    return f"₱{x:,.0f}"


def fmt_pct(x):
    return "n/a" if pd.isna(x) else f"{x:+.1%}"


def fmt_bps(x):
    return "n/a" if pd.isna(x) else f"{x:+.1f}"


# ---------------------------------------------------------------- SIDEBAR
st.sidebar.title("Settings")
upload = st.sidebar.file_uploader("Upload updated MIS CSV (replaces the default file for this session)", type="csv")
try:
    df = load_data(upload if upload is not None else DATA_PATH)
except Exception as e:
    st.error(f"Could not load data: {e}")
    st.stop()

months = sorted(df["YEAR_MONTH"].unique())
LATEST_M = months[-1]                                   # current, incomplete month
st.sidebar.caption(f"Source: **{upload.name if upload else DATA_PATH}**  \n"
                   f"Months loaded: **{months[0]} → {LATEST_M}** · {len(df):,} rows")
lp = pd.Period(LATEST_M, "M")
dim = calendar.monthrange(lp.year, lp.month)[1]
as_of = st.sidebar.date_input(
    f"Data as of ({LATEST_M})", value=date(lp.year, lp.month, min(14, dim)),
    min_value=date(lp.year, lp.month, 1), max_value=date(lp.year, lp.month, dim))
RR_FACTOR = dim / as_of.day
st.sidebar.caption(f"Run-rate applies to **{LATEST_M} only**: {dim} / {as_of.day} = **{RR_FACTOR:.3f}x**. "
                   "All earlier months are actuals.")

cur_m = st.sidebar.selectbox("Reporting month", months, index=len(months) - 1)
CUR_FACTOR = month_factor(cur_m)

include_credits = st.sidebar.checkbox("Net off credit (refund) transactions", value=True)
if not include_credits:
    df = df[df["TRXN_CODE"] == "Sale Transaction"]

mom_m, yoy_m = prev_month(cur_m, 1), prev_month(cur_m, 12)


def period_table(d, by):
    """Reporting month (run-rated only if it is the latest month) vs prior month & same month LY (actuals)."""
    by = by if isinstance(by, list) else [by]
    cur = summarize(d[d.YEAR_MONTH == cur_m], by)
    cur[FLOW_COLS] = cur[FLOW_COLS] * CUR_FACTOR
    mom = summarize(d[d.YEAR_MONTH == mom_m], by)               # actual
    yoy = summarize(d[d.YEAR_MONTH == yoy_m], by)               # actual
    t = cur.merge(mom, on=by, how="outer", suffixes=("", "_PM")) \
           .merge(yoy.add_suffix("_LY").rename(columns={f"{b}_LY": b for b in by}), on=by, how="outer")
    for m in FLOW_COLS:
        t[f"{m}_MoM%"] = t[m] / t[f"{m}_PM"] - 1
        t[f"{m}_YoY%"] = t[m] / t[f"{m}_LY"] - 1
    for m in ["WEIGHTED_MDR", "BLENDED_MDR", "WEIGHTED_IC", "BLENDED_IC"]:
        t[f"{m}_MoM_bps"] = (t[m] - t[f"{m}_PM"]) * 1e4
        t[f"{m}_YoY_bps"] = (t[m] - t[f"{m}_LY"]) * 1e4
    return t


def trend_df(d, by=None):
    """Monthly trend up to reporting month: actuals, except latest month = run-rate."""
    keys = ["YEAR_MONTH"] + ([by] if by else [])
    g = summarize(d[d.YEAR_MONTH <= cur_m], keys)
    f = g["YEAR_MONTH"].map(month_factor)
    for c in FLOW_COLS:
        g[c] = g[c] * f
    g["LABEL"] = g["YEAR_MONTH"].where(g["YEAR_MONTH"] != LATEST_M, LATEST_M + " (RR)")
    return g


VOL_LBL = "Volume (RR)" if CUR_FACTOR != 1 else "Volume"
TXN_LBL = "Txn Count (RR)" if CUR_FACTOR != 1 else "Txn Count"
IC_LBL = "Interchange ₱ (RR)" if CUR_FACTOR != 1 else "Interchange ₱"


def kpi_row(t):
    r = t.iloc[0]
    c = st.columns(6)
    c[0].metric(VOL_LBL, fmt_php(r.VOLUME), f"{fmt_pct(r['VOLUME_MoM%'])} MoM | {fmt_pct(r['VOLUME_YoY%'])} YoY")
    c[1].metric(TXN_LBL, f"{r.TXN_COUNT:,.0f}", f"{fmt_pct(r['TXN_COUNT_MoM%'])} MoM | {fmt_pct(r['TXN_COUNT_YoY%'])} YoY")
    c[2].metric("Weighted MDR", f"{r.WEIGHTED_MDR:.2%}", f"{fmt_bps(r.WEIGHTED_MDR_MoM_bps)} bps MoM | {fmt_bps(r.WEIGHTED_MDR_YoY_bps)} YoY")
    c[3].metric("Blended MDR (simple avg)", f"{r.BLENDED_MDR:.2%}", f"{fmt_bps(r.BLENDED_MDR_MoM_bps)} bps MoM | {fmt_bps(r.BLENDED_MDR_YoY_bps)} YoY")
    c[4].metric("Weighted Interchange", f"{r.WEIGHTED_IC:.2%}", f"{fmt_bps(r.WEIGHTED_IC_MoM_bps)} bps MoM | {fmt_bps(r.WEIGHTED_IC_YoY_bps)} YoY", delta_color="inverse")
    c[5].metric(IC_LBL, fmt_php(r.INTERCHANGE_PHP), f"{fmt_pct(r['INTERCHANGE_PHP_MoM%'])} MoM | {fmt_pct(r['INTERCHANGE_PHP_YoY%'])} YoY", delta_color="inverse")


DISPLAY_COLS = {
    "VOLUME": VOL_LBL, "VOLUME_PM": f"Volume {mom_m} (Actual)", "VOLUME_MoM%": "Vol MoM", "VOLUME_YoY%": "Vol YoY",
    "TXN_COUNT": TXN_LBL, "TXN_COUNT_MoM%": "Txn MoM", "TXN_COUNT_YoY%": "Txn YoY",
    "WEIGHTED_MDR": "Wtd MDR", "WEIGHTED_MDR_MoM_bps": "Wtd MDR MoM (bps)", "WEIGHTED_MDR_YoY_bps": "Wtd MDR YoY (bps)",
    "BLENDED_MDR": "Blended MDR",
    "WEIGHTED_IC": "Wtd IC", "WEIGHTED_IC_MoM_bps": "Wtd IC MoM (bps)", "WEIGHTED_IC_YoY_bps": "Wtd IC YoY (bps)",
    "BLENDED_IC": "Blended IC",
    "INTERCHANGE_PHP": IC_LBL, "INTERCHANGE_PHP_MoM%": "IC ₱ MoM", "INTERCHANGE_PHP_YoY%": "IC ₱ YoY",
}


def show_table(t, keys):
    keys = keys if isinstance(keys, list) else [keys]
    out = t[keys + list(DISPLAY_COLS)].rename(columns=DISPLAY_COLS)
    pct_cols = [v for k, v in DISPLAY_COLS.items() if k.endswith("%")]
    rate_cols = ["Wtd MDR", "Blended MDR", "Wtd IC", "Blended IC"]
    bps_cols = [v for k, v in DISPLAY_COLS.items() if k.endswith("bps")]
    num_cols = [VOL_LBL, DISPLAY_COLS["VOLUME_PM"], TXN_LBL, IC_LBL]
    fmt = {**{c: "{:+.1%}" for c in pct_cols}, **{c: "{:.2%}" for c in rate_cols},
           **{c: "{:+.1f}" for c in bps_cols}, **{c: "{:,.0f}" for c in num_cols}}
    if "Vol Share" in out.columns:
        fmt["Vol Share"] = "{:.1%}"
    st.dataframe(out.style.format(fmt, na_rep="-"), width="stretch", hide_index=True)


# ---------------------------------------------------------------- HEADER
st.title("Merchant Acquiring Dashboard")
basis = (f"MTD as of {as_of:%b %d, %Y}, run-rated ×{CUR_FACTOR:.2f}" if CUR_FACTOR != 1 else "actuals")
st.caption(f"Reporting month: **{cur_m}** ({basis}) · MoM vs **{mom_m}** (actual) · YoY vs **{yoy_m}** (actual) · "
           f"Trend charts: actuals; only **{LATEST_M}** is run-rated")

tab1, tab2, tab3, tab4 = st.tabs(["1 · PF vs DM Overview", "2 · Category Mix", "3 · Merchant Level", "4 · YoY & YTD"])

# ================================================================ TAB 1
with tab1:
    st.subheader("Total Portfolio")
    kpi_row(period_table(df.assign(ALL="Total"), "ALL"))

    st.subheader("By Segment")
    seg = period_table(df, "PF_DM_TAG")
    for tag in seg.PF_DM_TAG.dropna():
        st.markdown(f"**{tag}**")
        kpi_row(seg[seg.PF_DM_TAG == tag].reset_index(drop=True))

    tr = trend_df(df, "PF_DM_TAG")
    yrs1 = sorted(tr.YEAR_MONTH.str[:4].unique())
    sel1 = st.multiselect("Years shown in monthly trend charts", yrs1, default=yrs1[-2:], key="yr_tab1")
    tr = tr[tr.YEAR_MONTH.str[:4].isin(sel1)]
    if tr.empty:
        st.info("Select at least one year to show the trend charts.")
    c1, c2 = st.columns(2)
    c1.plotly_chart(px.bar(tr, x="LABEL", y="VOLUME", color="PF_DM_TAG", color_discrete_map=TAG_COLORS,
                           title="Monthly Volume (actuals; latest month run-rated)"), width="stretch")
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
    cat["VOL_SHARE_PM"] = cat.VOLUME_PM / cat.VOLUME_PM.sum()
    cat["SHARE_CHG_PP"] = (cat.VOL_SHARE - cat.VOL_SHARE_PM) * 100

    # Pies: top 5 categories of the selected segment (ranked separately for volume and count), rest = Others
    c1, c2 = st.columns(2)
    c1.plotly_chart(pie(top_n_others(cat, "VOLUME"), "VOLUME",
                        f"Volume Share – {seg_pick} – {cur_m} (Top {PIE_TOP_N} + Others)"), width="stretch")
    c2.plotly_chart(pie(top_n_others(cat, "TXN_COUNT"), "TXN_COUNT",
                        f"Txn Count Share – {seg_pick} – {cur_m} (Top {PIE_TOP_N} + Others)"), width="stretch")

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
    show_table(cat.assign(**{"Vol Share": cat.VOL_SHARE}), ["INDUSTRY", "Vol Share"])

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
                           title=f"Top {top_n} {seg3}s – {VOL_LBL} & Share", color_discrete_sequence=[TAG_COLORS[seg3]])
                    .update_yaxes(autorange="reversed"), width="stretch")
    mv = topm.melt(id_vars="PF_DM_NAME", value_vars=["VOLUME_MoM%", "VOLUME_YoY%"], var_name="Comparison", value_name="Growth")
    c2.plotly_chart(px.bar(mv, x="Growth", y="PF_DM_NAME", color="Comparison", barmode="group", orientation="h",
                           title="Volume Growth – MoM vs YoY").update_xaxes(tickformat=".0%")
                    .update_yaxes(autorange="reversed"), width="stretch")
    st.plotly_chart(px.scatter(topm, x="WEIGHTED_IC", y="WEIGHTED_MDR", size=topm.VOLUME.clip(lower=0), hover_name="PF_DM_NAME",
                               title="Weighted MDR vs Weighted Interchange (bubble = volume)")
                    .update_xaxes(tickformat=".2%").update_yaxes(tickformat=".2%"), width="stretch")
    show_table(mer.assign(**{"Vol Share": mer.VOL_SHARE}), ["PF_DM_NAME", "Vol Share"])

    st.divider()
    pick = st.selectbox(f"Drill down: {seg3}", mer.PF_DM_NAME.dropna().tolist())
    dd = d3[d3.PF_DM_NAME == pick]
    kpi_row(period_table(dd.assign(ALL=pick), "ALL"))
    trd = trend_df(dd)
    c1, c2 = st.columns(2)
    c1.plotly_chart(px.bar(trd, x="LABEL", y="VOLUME", title=f"{pick} – Monthly Volume",
                           color_discrete_sequence=[TAG_COLORS[seg3]]), width="stretch")
    c2.plotly_chart(px.line(trd, x="LABEL", y=["WEIGHTED_MDR", "WEIGHTED_IC"], markers=True,
                            title=f"{pick} – Weighted MDR vs Interchange").update_yaxes(tickformat=".2%"), width="stretch")
    if seg3 == "Payment Facilitator":
        st.markdown(f"**Sub-merchants under {pick} (top 20 by volume)**")
        sub = period_table(dd, "MERCHANT_NAME_ORIG").sort_values("VOLUME", ascending=False).head(20)
        show_table(sub, "MERCHANT_NAME_ORIG")


# ================================================================ TAB 4 – YoY & YTD
RR_SCALE_COLS = ["TOTAL_VOLUME", "TOTAL_COUNT", "INTERCHANGE_PHP", "MDR_X_VOL", "IC_X_VOL"]
MIN_LY_BASE = 1_000_000          # merchants with LY YTD volume below this show YoY as n.m.


def scale_rr(d):
    """Run-rate only the latest month's rows (rates stay intact because numerator & denominator both scale)."""
    d = d.copy()
    f = d["YEAR_MONTH"].map(month_factor)
    for c in RR_SCALE_COLS:
        d[c] = d[c] * f
    return d


def ytd_slice(d, year, end_mm):
    ym = d["YEAR_MONTH"]
    return d[(ym.str[:4] == str(year)) & (ym.str[5:7].astype(int) <= end_mm)]


def ytd_compare(d, by):
    """YTD CY (latest month run-rated) vs YTD LY (actual, same months) + FY outlook vs FY LY actual."""
    by = by if isinstance(by, list) else [by]
    d = scale_rr(d)
    cy = summarize(ytd_slice(d, CY, END_MM), by)
    ly = summarize(ytd_slice(d, LY, END_MM), by)
    fy_ly = summarize(d[d.YEAR_MONTH.str[:4] == str(LY)], by)[by + FLOW_COLS]
    t = cy.merge(ly, on=by, how="outer", suffixes=("", "_LY"))
    t = t.merge(fy_ly.rename(columns={c: f"{c}_FYLY" for c in FLOW_COLS}), on=by, how="outer")
    for m in FLOW_COLS:
        t[f"{m}_YoY%"] = t[m] / t[f"{m}_LY"] - 1
        t[f"{m}_FY"] = t[m] / END_MM * 12                               # simple run-rate outlook
        t[f"{m}_FY_YoY%"] = t[f"{m}_FY"] / t[f"{m}_FYLY"] - 1
    for m in ["WEIGHTED_MDR", "BLENDED_MDR", "WEIGHTED_IC", "BLENDED_IC"]:
        t[f"{m}_YoY_bps"] = (t[m] - t[f"{m}_LY"]) * 1e4
    return t


cp = pd.Period(cur_m, "M")
CY, LY, END_MM = cp.year, cp.year - 1, cp.month
MON = calendar.month_abbr
ytd_lbl = f"Jan–{MON[END_MM]}"

with tab4:
    seg4 = st.radio("Segment", ["All", "Direct Merchant", "Payment Facilitator"], horizontal=True, key="yoy_seg")
    d4 = df if seg4 == "All" else df[df.PF_DM_TAG == seg4]
    rr_note = f" ({MON[END_MM]} run-rated ×{CUR_FACTOR:.2f})" if CUR_FACTOR != 1 else ""
    st.caption(f"YTD {CY} = {ytd_lbl} {CY}{rr_note} vs YTD {LY} = {ytd_lbl} {LY} (actual). "
               f"FY {CY} outlook = YTD ÷ {END_MM} × 12 (simple run-rate) vs FY {LY} actual.")

    t = ytd_compare(d4.assign(ALL="Total"), "ALL").iloc[0]

    st.subheader(f"YTD {ytd_lbl}: {CY} vs {LY}")
    c = st.columns(5)
    c[0].metric(f"YTD Volume {CY}", fmt_php(t.VOLUME), f"{fmt_pct(t['VOLUME_YoY%'])} YoY (LY {fmt_php(t.VOLUME_LY)})")
    c[1].metric(f"YTD Txn Count {CY}", f"{t.TXN_COUNT:,.0f}", f"{fmt_pct(t['TXN_COUNT_YoY%'])} YoY (LY {t.TXN_COUNT_LY:,.0f})")
    c[2].metric("YTD Weighted MDR", f"{t.WEIGHTED_MDR:.2%}", f"{fmt_bps(t.WEIGHTED_MDR_YoY_bps)} bps YoY (LY {t.WEIGHTED_MDR_LY:.2%})")
    c[3].metric("YTD Weighted IC", f"{t.WEIGHTED_IC:.2%}", f"{fmt_bps(t.WEIGHTED_IC_YoY_bps)} bps YoY (LY {t.WEIGHTED_IC_LY:.2%})", delta_color="inverse")
    c[4].metric(f"YTD Interchange ₱ {CY}", fmt_php(t.INTERCHANGE_PHP), f"{fmt_pct(t['INTERCHANGE_PHP_YoY%'])} YoY", delta_color="inverse")

    st.subheader(f"Full-Year Outlook: FY {CY} (run-rate) vs FY {LY} (actual)")
    c = st.columns(3)
    c[0].metric(f"FY {CY} Volume (RR)", fmt_php(t.VOLUME_FY), f"{fmt_pct(t['VOLUME_FY_YoY%'])} vs FY {LY} {fmt_php(t.VOLUME_FYLY)}")
    c[1].metric(f"FY {CY} Txn Count (RR)", f"{t.TXN_COUNT_FY:,.0f}", f"{fmt_pct(t['TXN_COUNT_FY_YoY%'])} vs FY {LY} {t.TXN_COUNT_FYLY:,.0f}")
    c[2].metric(f"FY {CY} Interchange ₱ (RR)", fmt_php(t.INTERCHANGE_PHP_FY), f"{fmt_pct(t['INTERCHANGE_PHP_FY_YoY%'])} vs FY {LY}", delta_color="inverse")

    # ---- month-by-month overlay & cumulative YTD, by year
    m = summarize(scale_rr(d4)[lambda x: x.YEAR_MONTH <= cur_m], "YEAR_MONTH")
    m["YEAR"] = m.YEAR_MONTH.str[:4]
    m["MONTH"] = m.YEAR_MONTH.str[5:7].astype(int).map(lambda i: MON[i])
    m = m.sort_values("YEAR_MONTH")
    m["CUM_VOLUME"] = m.groupby("YEAR").VOLUME.cumsum()
    m["YOY%"] = m.VOLUME / m.groupby(m.YEAR_MONTH.str[5:7]).VOLUME.shift(1) - 1
    order = {"MONTH": list(MON)[1:]}
    yrs4 = sorted(m.YEAR.unique())
    sel4 = st.multiselect("Years shown in month-on-month charts", yrs4, default=yrs4, key="yr_tab4")
    m_all = m                                   # YoY % still uses full history (needs LY)
    m = m[m.YEAR.isin(sel4)]

    c1, c2 = st.columns(2)
    c1.plotly_chart(px.line(m, x="MONTH", y="VOLUME", color="YEAR", markers=True, category_orders=order,
                            title=f"Monthly Volume by Year ({LATEST_M[5:]}/{LATEST_M[:4]} run-rated)"), width="stretch")
    c2.plotly_chart(px.line(m, x="MONTH", y="CUM_VOLUME", color="YEAR", markers=True, category_orders=order,
                            title="Cumulative YTD Volume by Year"), width="stretch")
    c3, c4 = st.columns(2)
    c3.plotly_chart(px.bar(m_all[m_all.YEAR == str(CY)], x="MONTH", y="YOY%", category_orders=order,
                           title=f"Monthly Volume YoY % ({CY} vs {LY})").update_yaxes(tickformat=".0%"), width="stretch")
    c4.plotly_chart(px.line(m, x="MONTH", y="WEIGHTED_MDR", color="YEAR", markers=True, category_orders=order,
                            title="Weighted MDR by Year").update_yaxes(tickformat=".2%"), width="stretch")

    # ---- YTD YoY tables
    YTD_COLS = {
        "VOLUME": f"YTD Vol {CY}", "VOLUME_LY": f"YTD Vol {LY}", "VOLUME_YoY%": "Vol YoY",
        "TXN_COUNT": f"YTD Txn {CY}", "TXN_COUNT_LY": f"YTD Txn {LY}", "TXN_COUNT_YoY%": "Txn YoY",
        "WEIGHTED_MDR": "Wtd MDR", "WEIGHTED_MDR_YoY_bps": "Wtd MDR YoY (bps)",
        "BLENDED_MDR": "Blended MDR", "BLENDED_MDR_YoY_bps": "Blended MDR YoY (bps)",
        "WEIGHTED_IC": "Wtd IC", "WEIGHTED_IC_YoY_bps": "Wtd IC YoY (bps)",
        "BLENDED_IC": "Blended IC", "BLENDED_IC_YoY_bps": "Blended IC YoY (bps)",
        "VOLUME_FY": f"FY {CY} Vol (RR)", "VOLUME_FYLY": f"FY {LY} Vol", "VOLUME_FY_YoY%": "FY Vol YoY",
    }

    def show_ytd(tb, key):
        out = tb[[key] + list(YTD_COLS)].rename(columns=YTD_COLS)
        fmt = {v: ("{:+.1%}" if k.endswith("%") else "{:+.1f}" if k.endswith("bps")
                   else "{:.2%}" if ("MDR" in k or "IC" in k) and "INTERCHANGE" not in k else "{:,.0f}")
               for k, v in YTD_COLS.items()}
        st.dataframe(out.style.format(fmt, na_rep="-"), width="stretch", hide_index=True)

    st.subheader(f"YTD YoY by Segment")
    show_ytd(ytd_compare(df, "PF_DM_TAG"), "PF_DM_TAG")

    st.subheader(f"YTD YoY by Category – {seg4}")
    show_ytd(ytd_compare(d4, "INDUSTRY").sort_values("VOLUME", ascending=False), "INDUSTRY")

    st.subheader(f"YTD YoY by Merchant – {seg4}")
    mer4 = ytd_compare(d4, "PF_DM_NAME")
    mer4.loc[mer4.VOLUME_LY.fillna(0) < MIN_LY_BASE, "VOLUME_YoY%"] = float("nan")   # n.m. on tiny LY base
    mer4["VOL_CHG"] = mer4.VOLUME.fillna(0) - mer4.VOLUME_LY.fillna(0)
    movers = pd.concat([mer4.nlargest(10, "VOL_CHG"), mer4.nsmallest(10, "VOL_CHG")])
    st.plotly_chart(px.bar(movers.sort_values("VOL_CHG"), x="VOL_CHG", y="PF_DM_NAME", orientation="h",
                           color=movers.sort_values("VOL_CHG").VOL_CHG > 0,
                           color_discrete_map={True: "#2E8B57", False: "#C0392B"},
                           title=f"Top 10 Gainers / Decliners – YTD Volume change vs {LY} (₱)")
                    .update_layout(showlegend=False, height=600), width="stretch")
    st.caption(f"Vol YoY shown as '-' (n.m.) where YTD {LY} volume < ₱{MIN_LY_BASE:,.0f}.")
    show_ytd(mer4.sort_values("VOLUME", ascending=False), "PF_DM_NAME")
