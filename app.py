"""
Merchant Acquiring Dashboard (Card PF/DM + QRPH)
Run: streamlit run app.py
Requirements: streamlit>=1.50, pandas, plotly, openpyxl

Data sources
  * Card MIS CSV  (DATA_PATH or sidebar upload)  -> tabs 1-4
  * QRPH monthly sheet (QRPH_PATH or sidebar upload; CSV or XLSX) -> tabs 5-6
    Only YEAR_MONTH, VOLUME_POSTED, TXN_COUNT_POSTED are used.

Run-rate rule (applied per source):
  * ONLY the latest month of each source is run-rated (MTD x days-in-month / as-of day).
  * Every prior month (incl. MoM / YoY / YTD LY bases) = actual, simple summation.
"""
import calendar
import os
from datetime import date

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Merchant Acquiring Dashboard", layout="wide")

DATA_PATH = "MACQ - MIS_2026-11-19-0907 (1).csv"
QRPH_PATH = "data/qrph/QRPH_MIS.csv"                    # default QRPH file in repo (csv or xlsx)
TAG_COLORS = {"Direct Merchant": "#F28C28", "Payment Facilitator": "#1F4E79"}
SRC_COLORS = {"Card (PF + DM)": "#1F4E79", "QRPH": "#2E8B57"}
FLOW_COLS = ["VOLUME", "TXN_COUNT", "INTERCHANGE_PHP"]
CAT_TOP_N = 10
OTHERS_LABEL = "Others"
MIN_LY_BASE = 1_000_000
REQUIRED_COLS = ["YEAR_MONTH", "PF_DM_TAG", "PF_DM_NAME", "INDUSTRY", "TRXN_CODE", "TOTAL_VOLUME",
                 "TOTAL_COUNT", "MDR", "INTERCHANGE_RATE", "INTERCHANGE_EXPENSE_PHP"]
QRPH_REQUIRED = ["YEAR_MONTH", "VOLUME_POSTED", "TXN_COUNT_POSTED"]
MON = calendar.month_abbr


# ================================================================ DATA & HELPERS
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
        raise ValueError(f"Card MIS file is missing columns: {missing}")
    df["YEAR_MONTH"] = df["YEAR_MONTH"].astype(str).str[:7]
    df["MDR_X_VOL"] = df["MDR"] * df["TOTAL_VOLUME"]
    df["IC_X_VOL"] = df["INTERCHANGE_RATE"] * df["TOTAL_VOLUME"]
    df["INTERCHANGE_PHP"] = -df["INTERCHANGE_EXPENSE_PHP"]
    return df


def _norm_year_month(s):
    s = s.astype(str).str.strip()
    ok = s.str.match(r"^\d{4}-\d{2}")
    out = s.str[:7].where(ok)
    if (~ok).any():                                     # Excel dates / other formats
        out = out.fillna(pd.to_datetime(s[~ok], errors="coerce").dt.strftime("%Y-%m"))
    return out


@st.cache_data
def load_qrph(src, name):
    if str(name).lower().endswith((".xlsx", ".xls")):
        q = pd.read_excel(src)
    else:
        try:
            q = pd.read_csv(src, encoding="utf-8")
        except UnicodeDecodeError:
            if hasattr(src, "seek"):
                src.seek(0)
            q = pd.read_csv(src, encoding="cp1252")
    q.columns = q.columns.astype(str).str.strip().str.upper()
    missing = [c for c in QRPH_REQUIRED if c not in q.columns]
    if missing:
        raise ValueError(f"QRPH file is missing columns: {missing}")
    q["YEAR_MONTH"] = _norm_year_month(q["YEAR_MONTH"])
    for c in ["VOLUME_POSTED", "TXN_COUNT_POSTED"]:
        q[c] = pd.to_numeric(q[c].astype(str).str.replace(",", ""), errors="coerce")
    q = q.dropna(subset=["YEAR_MONTH"])
    return (q.groupby("YEAR_MONTH", as_index=False)
              .agg(VOLUME=("VOLUME_POSTED", "sum"), TXN_COUNT=("TXN_COUNT_POSTED", "sum"))
              .sort_values("YEAR_MONTH"))


def summarize(d, by):
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


def prev_month(ym, n=1):
    return str(pd.Period(ym, "M") - n)


def top_n_others(d, value_col, n=CAT_TOP_N):
    ranked = d[d["INDUSTRY"] != OTHERS_LABEL].sort_values(value_col, ascending=False)
    keep = ranked["INDUSTRY"].head(n).tolist()
    out = d.assign(INDUSTRY=d["INDUSTRY"].where(d["INDUSTRY"].isin(keep), OTHERS_LABEL))
    out = out.groupby("INDUSTRY", as_index=False)[value_col].sum()
    out["_ORDER"] = out["INDUSTRY"].eq(OTHERS_LABEL)
    return out.sort_values(["_ORDER", value_col], ascending=[True, False]).drop(columns="_ORDER")


def pie(d, value_col, title, names="INDUSTRY"):
    fig = px.pie(d, names=names, values=value_col, hole=0.45, title=title,
                 category_orders={names: d[names].tolist()})
    fig.update_traces(textinfo="percent", textposition="inside", insidetextorientation="horizontal", sort=False)
    fig.update_layout(showlegend=True, legend=dict(font=dict(size=10)), uniformtext_minsize=9,
                      uniformtext_mode="hide", margin=dict(t=60, b=20, l=20, r=20), height=480)
    return fig


def fmt_php(x):
    if pd.isna(x):
        return "-"
    for div, s in [(1e9, "B"), (1e6, "M"), (1e3, "K")]:
        if abs(x) >= div:
            return f"₱{x/div:,.2f}{s}"
    return f"₱{x:,.0f}"


def fmt_num(x):
    return "-" if pd.isna(x) else f"{x:,.0f}"


def fmt_pct(x):
    return "n/a" if pd.isna(x) else f"{x:+.1%}"


def fmt_bps(x):
    return "n/a" if pd.isna(x) else f"{x:+.1f}"


def safe_div(a, b):
    return a / b if pd.notna(a) and pd.notna(b) and b != 0 else float("nan")


# ================================================================ SIDEBAR
st.sidebar.title("Settings")
st.sidebar.subheader("Card MIS (PF / DM)")
upload = st.sidebar.file_uploader("Upload updated Card MIS CSV", type="csv", key="up_card")
try:
    df = load_data(upload if upload is not None else DATA_PATH)
except Exception as e:
    st.error(f"Could not load Card MIS data: {e}")
    st.stop()

months = sorted(df["YEAR_MONTH"].unique())
LATEST_M = months[-1]
st.sidebar.caption(f"Source: **{upload.name if upload else DATA_PATH}**  \n"
                   f"Months: **{months[0]} → {LATEST_M}** · {len(df):,} rows")
lp = pd.Period(LATEST_M, "M")
dim = calendar.monthrange(lp.year, lp.month)[1]
as_of = st.sidebar.date_input(
    f"Card data as of ({LATEST_M})", value=date(lp.year, lp.month, min(14, dim)),
    min_value=date(lp.year, lp.month, 1), max_value=date(lp.year, lp.month, dim), key="asof_card")
RR_FACTOR = dim / as_of.day
st.sidebar.caption(f"Run-rate applies to **{LATEST_M} only**: {dim} / {as_of.day} = **{RR_FACTOR:.3f}x**.")

# ---- QRPH
st.sidebar.divider()
st.sidebar.subheader("QRPH")
up_qr = st.sidebar.file_uploader("Upload QRPH sheet (CSV or XLSX)", type=["csv", "xlsx"], key="up_qrph")
qr, QR_LATEST, QR_RR = None, None, 1.0
try:
    if up_qr is not None:
        qr = load_qrph(up_qr, up_qr.name)
    elif os.path.exists(QRPH_PATH):
        qr = load_qrph(QRPH_PATH, QRPH_PATH)
except Exception as e:
    st.sidebar.error(f"QRPH: {e}")
if qr is not None and not qr.empty:
    QR_LATEST = qr["YEAR_MONTH"].max()
    qp = pd.Period(QR_LATEST, "M")
    qdim = calendar.monthrange(qp.year, qp.month)[1]
    default_day = as_of.day if QR_LATEST == LATEST_M else qdim
    qr_as_of = st.sidebar.date_input(
        f"QRPH data as of ({QR_LATEST})", value=date(qp.year, qp.month, min(default_day, qdim)),
        min_value=date(qp.year, qp.month, 1), max_value=date(qp.year, qp.month, qdim), key="asof_qr")
    QR_RR = qdim / qr_as_of.day
    st.sidebar.caption(f"Source: **{up_qr.name if up_qr else QRPH_PATH}**  \n"
                       f"Months: **{qr.YEAR_MONTH.min()} → {QR_LATEST}** · "
                       f"run-rate {QR_LATEST} only: **{QR_RR:.3f}x**")
else:
    st.sidebar.caption("No QRPH data loaded – tabs 5 and 6 will be empty.")

st.sidebar.divider()
cur_m = st.sidebar.selectbox("Reporting month", months, index=len(months) - 1)
include_credits = st.sidebar.checkbox("Net off credit (refund) transactions (Card)", value=True)
if not include_credits:
    df = df[df["TRXN_CODE"] == "Sale Transaction"]


def month_factor(ym):
    return RR_FACTOR if ym == LATEST_M else 1.0


def qr_factor(ym):
    return QR_RR if ym == QR_LATEST else 1.0


CUR_FACTOR = month_factor(cur_m)
mom_m, yoy_m = prev_month(cur_m, 1), prev_month(cur_m, 12)
cp = pd.Period(cur_m, "M")
CY, LY, END_MM = cp.year, cp.year - 1, cp.month
ytd_lbl = f"Jan–{MON[END_MM]}"

VOL_LBL = "Volume (RR)" if CUR_FACTOR != 1 else "Volume"
TXN_LBL = "Txn Count (RR)" if CUR_FACTOR != 1 else "Txn Count"
IC_LBL = "Interchange ₱ (RR)" if CUR_FACTOR != 1 else "Interchange ₱"


# ================================================================ CARD PERIOD LOGIC
def period_table(d, by):
    by = by if isinstance(by, list) else [by]
    cur = summarize(d[d.YEAR_MONTH == cur_m], by)
    cur[FLOW_COLS] = cur[FLOW_COLS] * CUR_FACTOR
    mom = summarize(d[d.YEAR_MONTH == mom_m], by)
    yoy = summarize(d[d.YEAR_MONTH == yoy_m], by)
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
    keys = ["YEAR_MONTH"] + ([by] if by else [])
    g = summarize(d[d.YEAR_MONTH <= cur_m], keys)
    f = g["YEAR_MONTH"].map(month_factor)
    for c in FLOW_COLS:
        g[c] = g[c] * f
    g["LABEL"] = g["YEAR_MONTH"].where(g["YEAR_MONTH"] != LATEST_M, LATEST_M + " (RR)")
    return g


RR_SCALE_COLS = ["TOTAL_VOLUME", "TOTAL_COUNT", "INTERCHANGE_PHP", "MDR_X_VOL", "IC_X_VOL"]


def scale_rr(d):
    d = d.copy()
    f = d["YEAR_MONTH"].map(month_factor)
    for c in RR_SCALE_COLS:
        d[c] = d[c] * f
    return d


def ytd_slice(d, year, end_mm):
    ym = d["YEAR_MONTH"]
    return d[(ym.str[:4] == str(year)) & (ym.str[5:7].astype(int) <= end_mm)]


def ytd_compare(d, by):
    by = by if isinstance(by, list) else [by]
    d = scale_rr(d)
    cy = summarize(ytd_slice(d, CY, END_MM), by)
    ly = summarize(ytd_slice(d, LY, END_MM), by)
    fy_ly = summarize(d[d.YEAR_MONTH.str[:4] == str(LY)], by)[by + FLOW_COLS]
    t = cy.merge(ly, on=by, how="outer", suffixes=("", "_LY"))
    t = t.merge(fy_ly.rename(columns={c: f"{c}_FYLY" for c in FLOW_COLS}), on=by, how="outer")
    for m in FLOW_COLS:
        t[f"{m}_YoY%"] = t[m] / t[f"{m}_LY"] - 1
        t[f"{m}_FY"] = t[m] / END_MM * 12
        t[f"{m}_FY_YoY%"] = t[f"{m}_FY"] / t[f"{m}_FYLY"] - 1
    for m in ["WEIGHTED_MDR", "BLENDED_MDR", "WEIGHTED_IC", "BLENDED_IC"]:
        t[f"{m}_YoY_bps"] = (t[m] - t[f"{m}_LY"]) * 1e4
    return t


# ================================================================ MONTHLY-SERIES LOGIC (QRPH / Combined)
def rr_series(m, fac):
    """Monthly series (YEAR_MONTH, VOLUME, TXN_COUNT) with only the source's latest month run-rated."""
    m = m[["YEAR_MONTH", "VOLUME", "TXN_COUNT"]].copy()
    f = m["YEAR_MONTH"].map(fac)
    m["VOLUME"] *= f
    m["TXN_COUNT"] *= f
    return m


def series_metrics(s):
    """s: run-rated monthly series. Returns {metric: {...periods & growth...}} for VOLUME, TXN_COUNT, TICKET."""
    s = s.set_index("YEAR_MONTH")
    ym = s.index.to_series()

    def at(m):
        return s.loc[m] if m in s.index else pd.Series({"VOLUME": float("nan"), "TXN_COUNT": float("nan")})

    def span(year, upto):
        sel = s[(ym.str[:4] == str(year)) & (ym.str[5:7].astype(int) <= upto)]
        return sel.sum() if len(sel) == upto else pd.Series({"VOLUME": float("nan"), "TXN_COUNT": float("nan")})

    P = {"CUR": at(cur_m), "PM": at(mom_m), "LYM": at(yoy_m),
         "YTD": span(CY, END_MM), "YTDLY": span(LY, END_MM), "FYLY": span(LY, 12)}
    P["FY"] = P["YTD"] / END_MM * 12
    out = {}
    for k in ["VOLUME", "TXN_COUNT"]:
        out[k] = {p: P[p][k] for p in P}
    out["TICKET"] = {p: safe_div(P[p]["VOLUME"], P[p]["TXN_COUNT"]) for p in P}
    for k, v in out.items():
        v["MoM%"] = safe_div(v["CUR"], v["PM"]) - 1
        v["YoY%"] = safe_div(v["CUR"], v["LYM"]) - 1
        v["YTD_YoY%"] = safe_div(v["YTD"], v["YTDLY"]) - 1
        v["FY_YoY%"] = safe_div(v["FY"], v["FYLY"]) - 1
    return out


def series_table(res_by_src):
    """res_by_src: {source label: series_metrics()} -> one comparison table."""
    names = {"VOLUME": "Volume (₱)", "TXN_COUNT": "Txn Count", "TICKET": "Avg Ticket (₱)"}
    rows = []
    for src, res in res_by_src.items():
        for k, lbl in names.items():
            v = res[k]
            rows.append({"Source": src, "Metric": lbl,
                         f"{cur_m}{' (RR)' if src != 'Card (PF + DM)' or CUR_FACTOR != 1 else ''}": v["CUR"],
                         f"{mom_m} (PM)": v["PM"], "MoM": v["MoM%"],
                         f"{yoy_m} (LY)": v["LYM"], "YoY": v["YoY%"],
                         f"YTD {CY}": v["YTD"], f"YTD {LY}": v["YTDLY"], "YTD YoY": v["YTD_YoY%"],
                         f"FY {CY} (RR)": v["FY"], f"FY {LY}": v["FYLY"], "FY YoY": v["FY_YoY%"]})
    t = pd.DataFrame(rows)
    pct = ["MoM", "YoY", "YTD YoY", "FY YoY"]
    fmt = {c: ("{:+.1%}" if c in pct else "{:,.0f}") for c in t.columns if c not in ("Source", "Metric")}
    st.dataframe(t.style.format(fmt, na_rep="-"), width="stretch", hide_index=True)


def series_kpis(res, title):
    v, n, t = res["VOLUME"], res["TXN_COUNT"], res["TICKET"]
    st.markdown(f"**{title}**")
    c = st.columns(5)
    c[0].metric(f"Volume {cur_m}", fmt_php(v["CUR"]), f"{fmt_pct(v['MoM%'])} MoM | {fmt_pct(v['YoY%'])} YoY")
    c[1].metric(f"Txn Count {cur_m}", fmt_num(n["CUR"]), f"{fmt_pct(n['MoM%'])} MoM | {fmt_pct(n['YoY%'])} YoY")
    c[2].metric(f"YTD Volume {CY}", fmt_php(v["YTD"]), f"{fmt_pct(v['YTD_YoY%'])} YoY (LY {fmt_php(v['YTDLY'])})")
    c[3].metric(f"YTD Txn Count {CY}", fmt_num(n["YTD"]), f"{fmt_pct(n['YTD_YoY%'])} YoY (LY {fmt_num(n['YTDLY'])})")
    c[4].metric(f"FY {CY} Volume (RR)", fmt_php(v["FY"]), f"{fmt_pct(v['FY_YoY%'])} vs FY {LY} {fmt_php(v['FYLY'])}")


def year_overlay(s, key, default_last=None):
    s = s[s.YEAR_MONTH <= cur_m].copy()
    s["YEAR"] = s.YEAR_MONTH.str[:4]
    s["MONTH"] = s.YEAR_MONTH.str[5:7].astype(int).map(lambda i: MON[i])
    s = s.sort_values("YEAR_MONTH")
    s["CUM_VOLUME"] = s.groupby("YEAR").VOLUME.cumsum()
    s["CUM_TXN"] = s.groupby("YEAR").TXN_COUNT.cumsum()
    s["TICKET"] = s.VOLUME / s.TXN_COUNT
    yrs = sorted(s.YEAR.unique())
    sel = st.multiselect("Years shown in month-on-month charts", yrs,
                         default=yrs if default_last is None else yrs[-default_last:], key=key)
    return s[s.YEAR.isin(sel)], {"MONTH": list(MON)[1:]}


# ================================================================ CARD DISPLAY HELPERS
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
    fmt = {}
    for k, v in YTD_COLS.items():
        if k.endswith("%"):
            fmt[v] = "{:+.1%}"
        elif k.endswith("bps"):
            fmt[v] = "{:+.1f}"
        elif k in ("WEIGHTED_MDR", "BLENDED_MDR", "WEIGHTED_IC", "BLENDED_IC"):
            fmt[v] = "{:.2%}"
        else:
            fmt[v] = "{:,.0f}"
    st.dataframe(out.style.format(fmt, na_rep="-"), width="stretch", hide_index=True)


# ================================================================ HEADER & TABS
st.title("Merchant Acquiring Dashboard")
basis = (f"MTD as of {as_of:%b %d, %Y}, run-rated ×{CUR_FACTOR:.2f}" if CUR_FACTOR != 1 else "actuals")
st.caption(f"Reporting month: **{cur_m}** ({basis}) · MoM vs **{mom_m}** (actual) · YoY vs **{yoy_m}** (actual) · "
           f"Only each source's latest month is run-rated")

tab_yoy, tab_ovw, tab_cat, tab_mer, tab_qr, tab_tot = st.tabs(
    ["1 · YoY & YTD", "2 · PF vs DM Overview", "3 · Category Mix", "4 · Merchant Level",
     "5 · QRPH", "6 · Total Acquiring (Card + QRPH)"])

# ================================================================ TAB 1 – YoY & YTD (Card)
with tab_yoy:
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

    m = summarize(scale_rr(d4)[lambda x: x.YEAR_MONTH <= cur_m], "YEAR_MONTH")
    m["YEAR"] = m.YEAR_MONTH.str[:4]
    m["MONTH"] = m.YEAR_MONTH.str[5:7].astype(int).map(lambda i: MON[i])
    m = m.sort_values("YEAR_MONTH")
    m["CUM_VOLUME"] = m.groupby("YEAR").VOLUME.cumsum()
    m["YOY%"] = m.VOLUME / m.groupby(m.YEAR_MONTH.str[5:7]).VOLUME.shift(1) - 1
    order = {"MONTH": list(MON)[1:]}
    yrs4 = sorted(m.YEAR.unique())
    sel4 = st.multiselect("Years shown in month-on-month charts", yrs4, default=yrs4, key="yr_tab4")
    m_all = m
    m = m[m.YEAR.isin(sel4)]

    c1, c2 = st.columns(2)
    c1.plotly_chart(px.line(m, x="MONTH", y="VOLUME", color="YEAR", markers=True, category_orders=order,
                            title=f"Monthly Volume by Year ({LATEST_M} run-rated)"), width="stretch")
    c2.plotly_chart(px.line(m, x="MONTH", y="CUM_VOLUME", color="YEAR", markers=True, category_orders=order,
                            title="Cumulative YTD Volume by Year"), width="stretch")
    c3, c4 = st.columns(2)
    c3.plotly_chart(px.bar(m_all[m_all.YEAR == str(CY)], x="MONTH", y="YOY%", category_orders=order,
                           title=f"Monthly Volume YoY % ({CY} vs {LY})").update_yaxes(tickformat=".0%"), width="stretch")
    c4.plotly_chart(px.line(m, x="MONTH", y="WEIGHTED_MDR", color="YEAR", markers=True, category_orders=order,
                            title="Weighted MDR by Year").update_yaxes(tickformat=".2%"), width="stretch")

    st.subheader("YTD YoY by Segment")
    show_ytd(ytd_compare(df, "PF_DM_TAG"), "PF_DM_TAG")
    st.subheader(f"YTD YoY by Category – {seg4}")
    show_ytd(ytd_compare(d4, "INDUSTRY").sort_values("VOLUME", ascending=False), "INDUSTRY")
    st.subheader(f"YTD YoY by Merchant – {seg4}")
    mer4 = ytd_compare(d4, "PF_DM_NAME")
    mer4.loc[mer4.VOLUME_LY.fillna(0) < MIN_LY_BASE, "VOLUME_YoY%"] = float("nan")
    mer4["VOL_CHG"] = mer4.VOLUME.fillna(0) - mer4.VOLUME_LY.fillna(0)
    movers = pd.concat([mer4.nlargest(10, "VOL_CHG"), mer4.nsmallest(10, "VOL_CHG")]).drop_duplicates("PF_DM_NAME")
    movers = movers.sort_values("VOL_CHG")
    st.plotly_chart(px.bar(movers, x="VOL_CHG", y="PF_DM_NAME", orientation="h",
                           color=(movers.VOL_CHG > 0).map({True: "Gainer", False: "Decliner"}),
                           color_discrete_map={"Gainer": "#2E8B57", "Decliner": "#C0392B"},
                           title=f"Top 10 Gainers / Decliners – YTD Volume change vs {LY} (₱)")
                    .update_layout(showlegend=False, height=600), width="stretch")
    st.caption(f"Vol YoY shown as '-' (n.m.) where YTD {LY} volume < ₱{MIN_LY_BASE:,.0f}.")
    show_ytd(mer4.sort_values("VOLUME", ascending=False), "PF_DM_NAME")

# ================================================================ TAB 2 – PF vs DM Overview
with tab_ovw:
    st.subheader("Total Portfolio (Card)")
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
    else:
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

# ================================================================ TAB 3 – Category Mix
with tab_cat:
    seg_pick = st.radio("Segment", ["All", "Direct Merchant", "Payment Facilitator"], horizontal=True, key="cat_seg")
    d2 = df if seg_pick == "All" else df[df.PF_DM_TAG == seg_pick]
    cat = period_table(d2, "INDUSTRY").sort_values("VOLUME", ascending=False)
    cat["VOL_SHARE"] = cat.VOLUME / cat.VOLUME.sum()
    cat["VOL_SHARE_PM"] = cat.VOLUME_PM / cat.VOLUME_PM.sum()
    cat["SHARE_CHG_PP"] = (cat.VOL_SHARE - cat.VOL_SHARE_PM) * 100

    c1, c2 = st.columns(2)
    c1.plotly_chart(pie(top_n_others(cat, "VOLUME"), "VOLUME",
                        f"Volume Share – {seg_pick} – {cur_m} (Top {CAT_TOP_N} + Others)"), width="stretch")
    c2.plotly_chart(pie(top_n_others(cat, "TXN_COUNT"), "TXN_COUNT",
                        f"Txn Count Share – {seg_pick} – {cur_m} (Top {CAT_TOP_N} + Others)"), width="stretch")

    tr2 = trend_df(d2, "INDUSTRY")
    top = cat.loc[cat.INDUSTRY != OTHERS_LABEL, "INDUSTRY"].head(CAT_TOP_N).tolist()
    tr2["CAT"] = tr2.INDUSTRY.where(tr2.INDUSTRY.isin(top), OTHERS_LABEL)
    tr2 = tr2.groupby(["LABEL", "CAT"], as_index=False).VOLUME.sum()
    tr2["SHARE"] = tr2.VOLUME / tr2.groupby("LABEL").VOLUME.transform("sum")
    st.plotly_chart(px.bar(tr2, x="LABEL", y="SHARE", color="CAT", category_orders={"CAT": top + [OTHERS_LABEL]},
                           title=f"Volume Share Trend (Top {CAT_TOP_N} + Others)")
                    .update_yaxes(tickformat=".0%"), width="stretch")
    st.plotly_chart(px.bar(cat.head(15), x="SHARE_CHG_PP", y="INDUSTRY", orientation="h",
                           title=f"Volume Share Change vs {mom_m} (pp)")
                    .update_yaxes(autorange="reversed"), width="stretch")
    show_table(cat.assign(**{"Vol Share": cat.VOL_SHARE}), ["INDUSTRY", "Vol Share"])

# ================================================================ TAB 4 – Merchant Level
with tab_mer:
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

# ================================================================ SERIES FOR TABS 5–6
card_m = summarize(df, "YEAR_MONTH")[["YEAR_MONTH", "VOLUME", "TXN_COUNT"]]
card_s = rr_series(card_m, month_factor)
qr_s = rr_series(qr, qr_factor) if qr is not None else None

# ================================================================ TAB 5 – QRPH
with tab_qr:
    if qr_s is None:
        st.info("Upload the QRPH sheet in the sidebar (needs YEAR_MONTH, VOLUME_POSTED, TXN_COUNT_POSTED).")
    else:
        qr_basis = (f"{QR_LATEST} run-rated ×{QR_RR:.2f}" if QR_RR != 1 else "actuals")
        st.caption(f"Basis: **posted** volume and transaction count only · {qr_basis} · "
                   f"MoM vs {mom_m} · YoY vs {yoy_m} · YTD {ytd_lbl} {CY} vs {LY} · FY outlook = YTD ÷ {END_MM} × 12")
        if cur_m not in set(qr.YEAR_MONTH):
            st.warning(f"QRPH sheet has no data for {cur_m}.")
        res_qr = series_metrics(qr_s)
        series_kpis(res_qr, f"QRPH – {cur_m}{' (RR)' if cur_m == QR_LATEST and QR_RR != 1 else ''}")
        tk = res_qr["TICKET"]
        st.caption(f"Avg ticket (posted): {fmt_php(tk['CUR'])} · {fmt_pct(tk['MoM%'])} MoM · "
                   f"{fmt_pct(tk['YoY%'])} YoY · YTD {fmt_php(tk['YTD'])} ({fmt_pct(tk['YTD_YoY%'])} YoY)")

        st.subheader("QRPH comparison table")
        series_table({"QRPH": res_qr})

        qo, order = year_overlay(qr_s, "yr_qr")
        c1, c2 = st.columns(2)
        c1.plotly_chart(px.line(qo, x="MONTH", y="VOLUME", color="YEAR", markers=True, category_orders=order,
                                title="QRPH Monthly Posted Volume by Year"), width="stretch")
        c2.plotly_chart(px.line(qo, x="MONTH", y="TXN_COUNT", color="YEAR", markers=True, category_orders=order,
                                title="QRPH Monthly Posted Txn Count by Year"), width="stretch")
        c3, c4 = st.columns(2)
        c3.plotly_chart(px.line(qo, x="MONTH", y="CUM_VOLUME", color="YEAR", markers=True, category_orders=order,
                                title="QRPH Cumulative YTD Posted Volume"), width="stretch")
        c4.plotly_chart(px.line(qo, x="MONTH", y="TICKET", color="YEAR", markers=True, category_orders=order,
                                title="QRPH Avg Ticket Size (₱)"), width="stretch")

# ================================================================ TAB 6 – Total Acquiring (Card + QRPH)
with tab_tot:
    if qr_s is None:
        st.info("Upload the QRPH sheet in the sidebar to see combined Card + QRPH metrics.")
    else:
        st.caption("Total = Card MIS (PF + DM) + QRPH posted. Each source is run-rated only in its own latest month. "
                   "Card MDR / interchange metrics are not applicable to QRPH and are excluded here.")
        both = card_s.merge(qr_s, on="YEAR_MONTH", how="outer", suffixes=("_CARD", "_QR")).sort_values("YEAR_MONTH")
        missing_qr = both.loc[both.VOLUME_QR.isna() & (both.YEAR_MONTH <= cur_m), "YEAR_MONTH"].tolist()
        missing_card = both.loc[both.VOLUME_CARD.isna() & (both.YEAR_MONTH <= cur_m), "YEAR_MONTH"].tolist()
        if missing_qr or missing_card:
            st.warning(f"Month coverage differs – no QRPH for: {missing_qr or 'none'}; no Card for: {missing_card or 'none'}. "
                       "Totals for those months include only the available source.")
        tot_s = pd.DataFrame({"YEAR_MONTH": both.YEAR_MONTH,
                              "VOLUME": both.VOLUME_CARD.fillna(0) + both.VOLUME_QR.fillna(0),
                              "TXN_COUNT": both.TXN_COUNT_CARD.fillna(0) + both.TXN_COUNT_QR.fillna(0)})

        res_tot, res_card, res_qr = series_metrics(tot_s), series_metrics(card_s), series_metrics(qr_s)
        series_kpis(res_tot, "Total Acquiring (Card + QRPH)")
        series_kpis(res_card, "Card (PF + DM)")
        series_kpis(res_qr, "QRPH")

        st.subheader("Comparison table")
        series_table({"Total (Card + QRPH)": res_tot, "Card (PF + DM)": res_card, "QRPH": res_qr})

        long = pd.concat([card_s.assign(SOURCE="Card (PF + DM)"), qr_s.assign(SOURCE="QRPH")])
        long = long[long.YEAR_MONTH <= cur_m]
        yrs = sorted(long.YEAR_MONTH.str[:4].unique())
        sel = st.multiselect("Years shown in monthly charts", yrs, default=yrs[-2:], key="yr_tot")
        long = long[long.YEAR_MONTH.str[:4].isin(sel)].copy()
        long["LABEL"] = long.YEAR_MONTH.where(long.YEAR_MONTH != cur_m,
                                              cur_m + (" (RR)" if cur_m in (LATEST_M, QR_LATEST) else ""))
        c1, c2 = st.columns(2)
        c1.plotly_chart(px.bar(long, x="LABEL", y="VOLUME", color="SOURCE", color_discrete_map=SRC_COLORS,
                               title="Monthly Volume – Card vs QRPH"), width="stretch")
        c2.plotly_chart(px.bar(long, x="LABEL", y="TXN_COUNT", color="SOURCE", color_discrete_map=SRC_COLORS,
                               title="Monthly Txn Count – Card vs QRPH"), width="stretch")

        sh = long.pivot_table(index="LABEL", columns="SOURCE", values=["VOLUME", "TXN_COUNT"], aggfunc="sum")
        share = pd.DataFrame({
            "Volume share": sh[("VOLUME", "QRPH")] / sh["VOLUME"].sum(axis=1),
            "Txn share": sh[("TXN_COUNT", "QRPH")] / sh["TXN_COUNT"].sum(axis=1)}).reset_index() \
            .melt(id_vars="LABEL", var_name="Measure", value_name="QRPH share")
        c3, c4 = st.columns(2)
        c3.plotly_chart(px.line(share, x="LABEL", y="QRPH share", color="Measure", markers=True,
                                title="QRPH Share of Total Acquiring").update_yaxes(tickformat=".0%"), width="stretch")
        mix = pd.DataFrame({"INDUSTRY": ["Card (PF + DM)", "QRPH"],
                            "VOLUME": [res_card["VOLUME"]["YTD"], res_qr["VOLUME"]["YTD"]]})
        c4.plotly_chart(pie(mix, "VOLUME", f"YTD {CY} Volume Mix").update_traces(
            marker=dict(colors=[SRC_COLORS["Card (PF + DM)"], SRC_COLORS["QRPH"]])), width="stretch")
