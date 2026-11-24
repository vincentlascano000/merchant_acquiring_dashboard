"""
Merchant Acquiring – Profitability Map (separate Streamlit app, own resource limit)
Run: streamlit run profit_app.py      Deploy: same repo, main file path = profit_app.py
Requirements: streamlit>=1.50, pandas, plotly, openpyxl

Basis: net MDR revenue = Σ(MDR × volume) − Σ(interchange rate × volume), actuals (not run-rated).
Reads the same Card MIS file as app.py (auto-detected in the repo by column headers, or uploaded).
"""
import calendar
import os
import re
from datetime import date

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from streamlit.delta_generator import DeltaGenerator as _DG

st.set_page_config(page_title="MA Profitability Map", layout="wide")

_orig_plotly_chart = _DG.plotly_chart
_CHART_N = [0]


def _plotly_chart_keyed(self, fig, *args, **kwargs):
    if "key" not in kwargs:
        _CHART_N[0] += 1
        kwargs["key"] = f"chart_{_CHART_N[0]}"
    return _orig_plotly_chart(self, fig, *args, **kwargs)


_DG.plotly_chart = _plotly_chart_keyed

SCAN_DIRS = [".", "data", "data/qrph", "data/card"]
TAG_COLORS = {"Direct Merchant": "#F28C28", "Payment Facilitator": "#1F4E79"}
SRC_COLORS = {"Card (PF + DM)": "#1F4E79", "QRPH": "#2E8B57"}
RR_COLOR = "#E8541E"
BASIS_ORDER = ["Actual", "Actual (MTD)", "Run-rate uplift"]
FLOW_COLS = ["VOLUME", "TXN_COUNT", "INTERCHANGE_PHP"]
CAT_TOP_N = 10
MID_TOP_N = 8
OTHERS_LABEL = "Others"
MIN_LY_BASE = 1_000_000
RATE_MIN_SPAN_BPS = 20
CARD_REQUIRED = ["YEAR_MONTH", "PF_DM_TAG", "PF_DM_NAME", "INDUSTRY", "TRXN_CODE", "TOTAL_VOLUME",
                 "TOTAL_COUNT", "MDR", "INTERCHANGE_RATE", "INTERCHANGE_EXPENSE_PHP"]
QRPH_REQUIRED = ["YEAR_MONTH", "VOLUME_POSTED", "TXN_COUNT_POSTED"]
DATE_COL = "MAX_TRAN_DATE"
SUB_COL = "MERCHANT_NAME_ORIG"
MID_COL = "MID"
MON = calendar.month_abbr
RR_LEGEND = ("Run-rated month: **solid** = MTD actual, **hatched / dotted ★** = run-rate, "
             "**open ○** = MTD actual point. Ratios (MDR, interchange rate, ticket size) are not run-rated.")
SEPS = [",", ";", "\t", "|"]
# Only these columns are read from the files – everything else (expense lines, NET_INCOME, etc.) is skipped
KEEP_COLS = set(CARD_REQUIRED + QRPH_REQUIRED + [DATE_COL, SUB_COL, "CARD_ACCEPTOR_ID"])
AGG_KEYS = ["YEAR_MONTH", "PF_DM_TAG", "PF_DM_NAME", SUB_COL, MID_COL, "INDUSTRY", "TRXN_CODE"]
QUAD_COLORS = {"① Protect – high volume, above hurdle": "#2E8B57",
               "② Fix – high volume, below hurdle": "#C0392B",
               "③ Grow – low volume, above hurdle": "#2F80ED",
               "④ Review – low volume, below hurdle": "#F2994A"}
MAIN_APP_URL = ""          # optional: paste the main dashboard URL to show a link back in the sidebar


# ================================================================ LOADING
def _clean_col(c):
    c = str(c).replace("\ufeff", "").replace("ï»¿", "").strip().strip('"').strip("'").upper()
    return re.sub(r"[^A-Z0-9]+", "_", c).strip("_")


def classify(cols):
    cols = {_clean_col(c) for c in cols}
    if set(CARD_REQUIRED) <= cols:
        return "card"
    if set(QRPH_REQUIRED) <= cols:
        return "qrph"
    return None


def _header_row(lines):
    for i, ln in enumerate(lines[:15]):
        if "YEAR_MONTH" in {_clean_col(x) for x in re.split(r"[,;\t|]", ln)}:
            return i
    return 0


def _read_any(src, name, nrows=None):
    if str(name).lower().endswith((".xlsx", ".xls")):
        raw = pd.read_excel(src, header=None, nrows=15)
        skip = next((i for i in range(len(raw))
                     if "YEAR_MONTH" in {_clean_col(x) for x in raw.iloc[i].astype(str)}), 0)
        if hasattr(src, "seek"):
            src.seek(0)
        return pd.read_excel(src, skiprows=skip, nrows=nrows, usecols=lambda c: _clean_col(c) in KEEP_COLS)
    for enc in ("utf-8-sig", "cp1252"):
        try:
            if hasattr(src, "seek"):
                src.seek(0)
                head = src.read(20000)
                src.seek(0)
                head = head.decode(enc) if isinstance(head, bytes) else head
            else:
                with open(src, encoding=enc) as fh:
                    head = fh.read(20000)
            lines = head.splitlines()
            skip = _header_row(lines)
            sep = max(SEPS, key=(lines[skip] if lines else "").count)
            return pd.read_csv(src, encoding=enc, low_memory=False, sep=sep, skiprows=skip, nrows=nrows,
                               usecols=lambda c: _clean_col(c) in KEEP_COLS,
                               dtype={"CARD_ACCEPTOR_ID": str, "MAX_TRAN_DATE": str})
        except UnicodeDecodeError:
            continue
    raise ValueError(f"{name}: could not decode file")


def _norm_ym(s):
    s = s.astype(str).str.strip()
    ok = s.str.match(r"^\d{4}-\d{2}")
    out = s.str[:7].where(ok)
    if (~ok).any():
        out = out.fillna(pd.to_datetime(s[~ok], errors="coerce").dt.strftime("%Y-%m"))
    return out


def _norm_mid(s):
    s = s.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    sci = s.str.match(r"^\d+(\.\d+)?[eE]\+\d+$")
    if sci.any():
        s = s.where(~sci, pd.to_numeric(s[sci], errors="coerce").map(lambda v: f"{v:.0f}"))
    return s.replace({"nan": "NO MID", "": "NO MID", "None": "NO MID"})


def _parse_dates(s):
    if pd.api.types.is_datetime64_any_dtype(s):
        return s
    raw = s.astype(str).str.strip().replace({"": None, "nan": None, "NaT": None, "None": None})
    out = pd.to_datetime(raw.str[:10], format="%Y-%m-%d", errors="coerce")
    for fmt in ("%m/%d/%Y", "%d/%m/%Y", "%Y/%m/%d", "%d-%b-%Y", "%d-%b-%y", "%m/%d/%y"):
        miss = out.isna() & raw.notna()
        if not miss.any():
            break
        out = out.fillna(pd.to_datetime(raw.where(miss).str.split(" ").str[0], format=fmt, errors="coerce"))
    miss = out.isna() & raw.notna()
    if miss.any():
        num = pd.to_numeric(raw.where(miss), errors="coerce")
        ser = pd.to_datetime(num.where(num.between(20000, 80000)), unit="D", origin="1899-12-30", errors="coerce")
        out = out.fillna(ser)
    return out


@st.cache_resource(show_spinner="Loading files…", max_entries=2)
def load_sources(files):
    card, qr, log = [], [], []
    for name, src in files:
        d = _read_any(src, name)
        d.columns = [_clean_col(c) for c in d.columns]
        kind = classify(d.columns)
        if kind == "card":
            card.append(d)
            log.append(f"✅ **{name}** → Card MIS" + (f" (has {DATE_COL})" if DATE_COL in d.columns
                                                       else f" (⚠️ no {DATE_COL} column)"))
        elif kind == "qrph":
            qr.append(d)
            log.append(f"✅ **{name}** → QRPH")
        else:
            log.append(f"⚠️ **{name}** not recognised (needs Card MIS or QRPH columns) – skipped")

    cdf = None
    if card:
        cdf = pd.concat(card, ignore_index=True)
        cdf["YEAR_MONTH"] = _norm_ym(cdf["YEAR_MONTH"])
        for c in ["TOTAL_VOLUME", "TOTAL_COUNT", "MDR", "INTERCHANGE_RATE", "INTERCHANGE_EXPENSE_PHP"]:
            cdf[c] = pd.to_numeric(cdf[c].astype(str).str.replace(",", ""), errors="coerce").fillna(0)
        if SUB_COL not in cdf.columns:
            cdf[SUB_COL] = cdf["PF_DM_NAME"]
        cdf[SUB_COL] = cdf[SUB_COL].fillna("UNTAGGED").astype(str).str.strip()
        cdf[MID_COL] = _norm_mid(cdf["CARD_ACCEPTOR_ID"]) if "CARD_ACCEPTOR_ID" in cdf.columns else "NO MID"
        if DATE_COL in cdf.columns:
            cdf[DATE_COL] = _parse_dates(cdf[DATE_COL])
        cdf["MDR_X_VOL"] = cdf["MDR"] * cdf["TOTAL_VOLUME"]
        cdf["IC_X_VOL"] = cdf["INTERCHANGE_RATE"] * cdf["TOTAL_VOLUME"]
        cdf["INTERCHANGE_PHP"] = -cdf["INTERCHANGE_EXPENSE_PHP"]
        # pre-aggregate to one row per month × segment × merchant × sub-merchant × MID × industry × txn code
        cdf = cdf.dropna(subset=["YEAR_MONTH"])
        for k in AGG_KEYS[1:]:
            cdf[k] = cdf[k].fillna("UNTAGGED").astype(str).str.strip()
        aggs = dict(TOTAL_VOLUME=("TOTAL_VOLUME", "sum"), TOTAL_COUNT=("TOTAL_COUNT", "sum"),
                    MDR_X_VOL=("MDR_X_VOL", "sum"), IC_X_VOL=("IC_X_VOL", "sum"),
                    INTERCHANGE_PHP=("INTERCHANGE_PHP", "sum"),
                    MDR_SUM=("MDR", "sum"), IC_SUM=("INTERCHANGE_RATE", "sum"), N_ROWS=("MDR", "size"))
        if DATE_COL in cdf.columns:
            aggs[DATE_COL] = (DATE_COL, "max")
        cdf = cdf.groupby(AGG_KEYS, as_index=False, sort=False).agg(**aggs)

    qdf = None
    if qr:
        q = pd.concat(qr, ignore_index=True)
        q["YEAR_MONTH"] = _norm_ym(q["YEAR_MONTH"])
        for c in QRPH_REQUIRED[1:]:
            q[c] = pd.to_numeric(q[c].astype(str).str.replace(",", ""), errors="coerce")
        q = q.dropna(subset=["YEAR_MONTH"]).drop_duplicates(subset=["YEAR_MONTH"] + QRPH_REQUIRED[1:])
        qdf = (q.groupby("YEAR_MONTH", as_index=False)
                .agg(VOLUME=("VOLUME_POSTED", "sum"), TXN_COUNT=("TXN_COUNT_POSTED", "sum"))
                .sort_values("YEAR_MONTH"))
    return cdf, qdf, log


def _sniff_header(path):
    if path.lower().endswith((".xlsx", ".xls")):
        return list(_read_any(path, path).columns)
    for enc in ("utf-8-sig", "cp1252"):
        try:
            with open(path, encoding=enc) as fh:
                lines = [fh.readline() for _ in range(15)]
            ln = lines[_header_row(lines)].strip()
            return ln.split(max(SEPS, key=ln.count))
        except UnicodeDecodeError:
            continue
    return []


@st.cache_data(ttl=300)
def scan_repo():
    found, allf = {"card": [], "qrph": []}, []
    for d in SCAN_DIRS:
        if not os.path.isdir(d):
            continue
        for f in sorted(os.listdir(d)):
            p = os.path.normpath(os.path.join(d, f))
            if not f.lower().endswith((".csv", ".xlsx", ".xls")) or f.startswith("~$") or p in [a[0] for a in allf]:
                continue
            try:
                hdr = [_clean_col(h) for h in _sniff_header(p)]
                kind = classify(hdr)
            except Exception as e:
                hdr, kind = [f"read error: {e}"], None
            allf.append((p, kind, hdr[:12], DATE_COL in hdr))
            if kind:
                found[kind].append((p, DATE_COL in hdr, os.path.getmtime(p)))
    for k in ("card", "qrph"):
        found[k] = [p for p, has_dt, mt in sorted(found[k], key=lambda x: (x[1], x[2]), reverse=True)]
    found["_all"] = allf
    return found


# ================================================================ HELPERS
def prev_month(ym, n=1):
    return str(pd.Period(ym, "M") - n)


def fmt_php(x):
    if pd.isna(x):
        return "-"
    for div, s in [(1e9, "B"), (1e6, "M"), (1e3, "K")]:
        if abs(x) >= div:
            return f"₱{x/div:,.2f}{s}"
    return f"₱{x:,.0f}"


def safe_div(a, b):
    return a / b if pd.notna(a) and pd.notna(b) and b != 0 else float("nan")


def month_picker_date(label, ym, default_day, key):
    p = pd.Period(ym, "M")
    dim = calendar.monthrange(p.year, p.month)[1]
    d = st.sidebar.date_input(label, value=date(p.year, p.month, min(default_day, dim)),
                              min_value=date(p.year, p.month, 1), max_value=date(p.year, p.month, dim), key=key)
    return dim / d.day, d


def rate_line(d, x, y, title, color=None, order=None, cmap=None):
    ys = y if isinstance(y, list) else [y]
    d = d.copy()
    for c in ys:
        d[c] = d[c].round(6)
    fig = px.line(d, x=x, y=y, color=color, markers=True, title=title, category_orders=order or {},
                  color_discrete_map=cmap or {})
    vals = pd.concat([d[c] for c in ys]).dropna()
    vals = vals[vals.between(-1, 1)]
    if not vals.empty:
        lo, hi = vals.min(), vals.max()
        span = max(hi - lo, RATE_MIN_SPAN_BPS / 1e4)
        mid = (lo + hi) / 2
        pad = span * 0.15
        fig.update_yaxes(range=[mid - span / 2 - pad, mid + span / 2 + pad])
    fig.update_yaxes(tickformat=".2%", title=None)
    fig.update_xaxes(type="category")
    fig.update_traces(hovertemplate="%{x}<br>%{y:.3%}<extra>%{fullData.name}</extra>")
    fig.update_layout(legend_title_text="")
    return fig


def ytd_slice(d, year, end_mm):
    ym = d["YEAR_MONTH"]
    return d[(ym.str[:4] == str(year)) & (ym.str[5:7].astype(int) <= end_mm)]


def profit_table(d, keys):
    """Per-entity volume and NET MDR REVENUE = Σ(MDR × volume) − Σ(interchange rate × volume)
    = gross MDR revenue net of interchange. Margin (bps) = net MDR revenue ÷ volume × 10,000
    (= volume-weighted MDR − interchange spread). Actuals – not run-rated."""
    keys = keys if isinstance(keys, list) else [keys]
    g = d.groupby(keys, dropna=False).agg(
        VOLUME=("TOTAL_VOLUME", "sum"), TXN_COUNT=("TOTAL_COUNT", "sum"),
        MDR_REV=("MDR_X_VOL", "sum"), IC_COST=("IC_X_VOL", "sum"),
        MONTHS=("YEAR_MONTH", "nunique")).reset_index()
    g = g[g.VOLUME > 0].copy()
    g["NET_REV"] = g.MDR_REV - g.IC_COST
    g["WTD_MDR"] = g.MDR_REV / g.VOLUME
    g["WTD_IC"] = g.IC_COST / g.VOLUME
    g["MARGIN_BPS"] = g.NET_REV / g.VOLUME * 1e4
    g["TICKET"] = g.VOLUME / g.TXN_COUNT.where(g.TXN_COUNT > 0)
    g["NET_REV_PER_TXN"] = g.NET_REV / g.TXN_COUNT.where(g.TXN_COUNT > 0)
    g["NAME"] = g[keys].astype(str).agg(" › ".join, axis=1)
    return g


def assign_quadrants(g, pareto_pct, hurdle_bps):
    g = g.sort_values("VOLUME", ascending=False).copy()
    share = g.VOLUME / g.VOLUME.sum()
    g["VOL_SHARE"] = share
    g["HIGH_VOL"] = (share.cumsum() - share) < pareto_pct / 100      # merchants making up the top X% of volume
    vol_cut = g.loc[g.HIGH_VOL, "VOLUME"].min() if g.HIGH_VOL.any() else g.VOLUME.max()
    prof = g.MARGIN_BPS >= hurdle_bps
    quads = list(QUAD_COLORS)
    g["QUADRANT"] = np.select([g.HIGH_VOL & prof, g.HIGH_VOL & ~prof, ~g.HIGH_VOL & prof],
                              quads[:3], default=quads[3])
    return g, vol_cut


def need_card():
    return True


# ================================================================ SIDEBAR
st.sidebar.title("Profitability Map")
if MAIN_APP_URL:
    st.sidebar.markdown(f"[← Back to main dashboard]({MAIN_APP_URL})")
upload = st.sidebar.file_uploader("Upload Card MIS CSV (optional)", type=["csv", "xlsx"], key="up_card")
repo_found = scan_repo()
files = []
if upload is not None:
    files = [(upload.name, upload)]
elif repo_found["card"]:
    cands = repo_found["card"]
    card_file = (st.sidebar.selectbox("Card MIS file to use", cands, key="card_file") if len(cands) > 1
                 else cands[0])
    files = [(card_file, card_file)]
if st.sidebar.button("Re-scan repo files"):
    scan_repo.clear()
    load_sources.clear()
    st.rerun()
if not files:
    st.info("No Card MIS file found in the repo – upload one in the sidebar.")
    st.stop()
try:
    card_raw, _, load_log = load_sources(tuple(files))
except Exception as e:
    st.error(f"Could not load data: {e}")
    st.stop()
for line in load_log:
    st.sidebar.caption(line)
if card_raw is None:
    st.error("The selected file is not a Card MIS file.")
    st.stop()


def data_max_tran_date(d, ym):
    if DATE_COL not in d.columns:
        return None, f"the loaded card file has no {DATE_COL} column"
    s_all = d.loc[d["YEAR_MONTH"] == ym, DATE_COL]
    s = s_all.dropna()
    if s.empty:
        return None, f"{DATE_COL} is blank / unreadable for {ym} ({len(s_all):,} rows)"
    p = pd.Period(ym, "M")
    s_in = s[(s >= p.start_time) & (s <= p.end_time)]
    if s_in.empty:
        return None, f"{DATE_COL} values for {ym} fall outside that month (e.g. {s.max():%Y-%m-%d})"
    return s_in.max().date(), None


LATEST_M = card_raw["YEAR_MONTH"].max()
lp_ = pd.Period(LATEST_M, "M")
dim_ = calendar.monthrange(lp_.year, lp_.month)[1]
auto_dt, _why = data_max_tran_date(card_raw, LATEST_M)
card_asof = auto_dt or date(lp_.year, lp_.month, dim_)
st.sidebar.caption(f"Latest month {LATEST_M}: data as of **{card_asof:%b %d, %Y}**"
                   + ("" if auto_dt else f" (no {DATE_COL} – assumed full month)"))
months = sorted(card_raw["YEAR_MONTH"].dropna().unique())
cur_m = st.sidebar.selectbox("Reporting month", months, index=len(months) - 1, key="cur_m")
include_credits = st.sidebar.checkbox("Net off credit (refund) transactions", value=True)
df = card_raw if include_credits else card_raw[card_raw["TRXN_CODE"] == "Sale Transaction"]

CUR_FACTOR = dim_ / card_asof.day if cur_m == LATEST_M else 1.0
cp = pd.Period(cur_m, "M")
CY, LY, END_MM = cp.year, cp.year - 1, cp.month
ytd_lbl = f"Jan–{calendar.month_abbr[END_MM]}"

st.title("Merchant Acquiring – Profitability Map")
with st.container():
    if need_card():
        st.caption(
            "Profitability basis = **net MDR revenue** = Σ(MDR × volume) − Σ(interchange rate × volume), i.e. gross "
            "MDR revenue net of interchange. Other costs/income (txn, MID, license, processing fees) are **not** "
            "included. **Margin (bps)** = net MDR revenue ÷ volume × 10,000 = weighted MDR − weighted interchange. "
            "Actuals, not run-rated. Quadrants: **high volume** = merchants that together make up the top X% of the "
            "segment's volume; **above hurdle** = margin ≥ hurdle.")
        cA, cB, cC = st.columns([1.3, 1.3, 1.4])
        per = cA.radio("Period", ["YTD", "MONTH"], horizontal=True, key="prof_per",
                       format_func=lambda x: f"YTD {CY} ({ytd_lbl})" if x == "YTD"
                       else f"{cur_m}{' (MTD)' if CUR_FACTOR != 1 else ''}")
        pareto = cB.select_slider("High volume = top X% of segment volume", [50, 60, 70, 80, 90, 95],
                                  value=80, key="prof_pareto")
        hurdle = cC.number_input("Margin hurdle (bps of volume)", value=0.0, step=5.0, key="prof_hurdle",
                                 help="0 = MDR at least covers interchange. Raise it to require a minimum spread, "
                                      "e.g. to cover other per-volume costs.")
        dper = ytd_slice(df, CY, END_MM) if per == "YTD" else df[df.YEAR_MONTH == cur_m]
        per_lbl = f"YTD {CY}" if per == "YTD" else cur_m

        # ---- segment comparison
        st.subheader(f"Direct Merchants vs Payment Facilitators – {per_lbl}")
        rows = []
        for s in ["Direct Merchant", "Payment Facilitator"]:
            g = profit_table(dper[dper.PF_DM_TAG == s], "PF_DM_NAME")
            if g.empty:
                continue
            g, _ = assign_quadrants(g, pareto, hurdle)
            below = g.MARGIN_BPS < hurdle
            rows.append({"Segment": s, "Merchants": len(g), "Volume": g.VOLUME.sum(),
                         "Gross MDR revenue": g.MDR_REV.sum(), "Interchange": g.IC_COST.sum(),
                         "Net MDR revenue": g.NET_REV.sum(),
                         "Wtd MDR": g.MDR_REV.sum() / g.VOLUME.sum(), "Wtd IC": g.IC_COST.sum() / g.VOLUME.sum(),
                         "Margin (bps)": g.NET_REV.sum() / g.VOLUME.sum() * 1e4,
                         "# below hurdle": int(below.sum()),
                         "Vol share below hurdle": g.loc[below, "VOLUME"].sum() / g.VOLUME.sum(),
                         "Net MDR rev of below-hurdle": g.loc[below, "NET_REV"].sum()})
        if rows:
            comp = pd.DataFrame(rows)
            st.dataframe(comp.style.format({"Volume": "{:,.0f}", "Gross MDR revenue": "{:,.0f}",
                                            "Interchange": "{:,.0f}", "Net MDR revenue": "{:,.0f}",
                                            "Wtd MDR": "{:.2%}", "Wtd IC": "{:.2%}", "Margin (bps)": "{:+.1f}",
                                            "Vol share below hurdle": "{:.1%}",
                                            "Net MDR rev of below-hurdle": "{:,.0f}"}, na_rep="-"),
                         width="stretch", hide_index=True)

        # ---- one segment
        st.divider()
        cA, cB = st.columns([1.3, 2])
        seg_p = cA.radio("Segment", ["Direct Merchant", "Payment Facilitator"], horizontal=True, key="prof_seg")
        lvl_map = {"PF_DM_NAME": ["PF_DM_NAME"]}
        if seg_p == "Payment Facilitator":
            lvl_map.update({MID_COL: ["PF_DM_NAME", MID_COL], SUB_COL: ["PF_DM_NAME", SUB_COL]})
        lvl = cB.radio("Level", list(lvl_map), horizontal=True, key=f"prof_lvl_{seg_p}",
                       format_func={"PF_DM_NAME": "Merchant / PF", MID_COL: "MID", SUB_COL: "Sub-merchant"}.get)
        g = profit_table(dper[dper.PF_DM_TAG == seg_p], lvl_map[lvl])
        if g.empty:
            st.info("No volume for this segment / period.")
        else:
            g, vol_cut = assign_quadrants(g, pareto, hurdle)
            tot_v, tot_rev = g.VOLUME.sum(), g.NET_REV.sum()
            below = g.MARGIN_BPS < hurdle
            c = st.columns(5)
            c[0].metric(f"Volume ({per_lbl})", fmt_php(tot_v))
            c[1].metric("Gross MDR revenue", fmt_php(g.MDR_REV.sum()), f"less interchange {fmt_php(g.IC_COST.sum())}",
                        delta_color="off")
            c[2].metric("Net MDR revenue", fmt_php(tot_rev), f"{tot_rev / tot_v * 1e4:+.1f} bps of volume",
                        delta_color="off")
            c[3].metric(f"Below hurdle ({hurdle:+.0f} bps)", f"{int(below.sum()):,} of {len(g):,}",
                        f"{g.loc[below, 'VOLUME'].sum() / tot_v:.1%} of volume", delta_color="off")
            c[4].metric("Net MDR rev from below-hurdle", fmt_php(g.loc[below, "NET_REV"].sum()),
                        f"vs {fmt_php(g.loc[~below, 'NET_REV'].sum())} from the rest", delta_color="off")

            qs = (g.groupby("QUADRANT").agg(Entities=("NAME", "count"), Volume=("VOLUME", "sum"),
                                            Net_rev=("NET_REV", "sum")).reindex(list(QUAD_COLORS)).fillna(0)
                  .reset_index())
            qs["Vol share"] = qs.Volume / tot_v
            qs["Rev share"] = qs.Net_rev / tot_rev if tot_rev else float("nan")
            qs["Margin (bps)"] = qs.Net_rev / qs.Volume.where(qs.Volume > 0) * 1e4
            st.dataframe(qs.rename(columns={"QUADRANT": "Quadrant", "Net_rev": "Net MDR revenue"})
                         .style.format({"Entities": "{:,.0f}", "Volume": "{:,.0f}", "Net MDR revenue": "{:,.0f}",
                                        "Vol share": "{:.1%}", "Rev share": "{:.1%}", "Margin (bps)": "{:+.1f}"},
                                       na_rep="-"), width="stretch", hide_index=True)

            # ---- map: volume (log) vs margin bps
            lo, hi = np.nanpercentile(g.MARGIN_BPS, [2, 98])
            lo, hi = min(lo, hurdle - 10), max(hi, hurdle + 10)
            pad = (hi - lo) * 0.08
            g["Y"] = g.MARGIN_BPS.clip(lo - pad, hi + pad)
            g["Shown"] = np.where(g.MARGIN_BPS.between(lo - pad, hi + pad), "in range", "clipped (outlier)")
            g["SIZE"] = g.NET_REV.abs()
            hov = {"VOLUME": ":,.0f", "NET_REV": ":,.0f", "MARGIN_BPS": ":+.1f", "WTD_MDR": ":.2%", "WTD_IC": ":.2%",
                   "TICKET": ":,.0f", "Y": False, "SIZE": False, "Shown": False, "QUADRANT": False}
            lbl = {"VOLUME": f"Volume ({per_lbl}, log scale)", "Y": "Net MDR margin (bps)",
                   "NET_REV": "Net MDR revenue", "MARGIN_BPS": "Margin bps", "WTD_MDR": "Wtd MDR",
                   "WTD_IC": "Wtd IC", "TICKET": "Avg ticket"}
            fig = px.scatter(g, x="VOLUME", y="Y", color="QUADRANT", size="SIZE", size_max=40, log_x=True,
                             symbol="Shown", symbol_map={"in range": "circle", "clipped (outlier)": "triangle-up"},
                             color_discrete_map=QUAD_COLORS, category_orders={"QUADRANT": list(QUAD_COLORS)},
                             hover_name="NAME", hover_data=hov, labels=lbl,
                             title=f"Profitability map – {seg_p} ({per_lbl}) · bubble size = |net MDR revenue|")
            fig.add_hline(y=hurdle, line_dash="dash", line_color="grey",
                          annotation_text=f"hurdle {hurdle:+.0f} bps", annotation_position="top left")
            fig.add_vline(x=vol_cut, line_dash="dash", line_color="grey",
                          annotation_text=f"high-volume cut-off {fmt_php(vol_cut)}", annotation_position="top")
            fig.update_layout(height=620, legend_title_text="")
            st.plotly_chart(fig, width="stretch")

            # ---- MDR vs interchange: distance above the diagonal = margin
            c1, c2 = st.columns(2)
            f2 = px.scatter(g, x="WTD_IC", y="WTD_MDR", color="QUADRANT", size="VOLUME", size_max=30,
                            color_discrete_map=QUAD_COLORS, hover_name="NAME",
                            hover_data={"WTD_IC": ":.2%", "WTD_MDR": ":.2%", "MARGIN_BPS": ":+.1f",
                                        "VOLUME": ":,.0f", "QUADRANT": False},
                            labels={"WTD_IC": "Weighted interchange", "WTD_MDR": "Weighted MDR",
                                    "MARGIN_BPS": "Margin bps"},
                            title="MDR vs interchange – dotted line = MDR equals interchange "
                                  "(below it = selling under cost); dashed = MDR equals interchange + hurdle")
            m = max(g.WTD_MDR.max(), g.WTD_IC.max()) * 1.05
            f2.add_shape(type="line", x0=0, y0=0, x1=m, y1=m, line=dict(dash="dot", color="grey"))
            if hurdle:
                f2.add_shape(type="line", x0=0, y0=hurdle / 1e4, x1=m, y1=m + hurdle / 1e4,
                             line=dict(dash="dash", color="grey"))
            f2.update_xaxes(tickformat=".2%", range=[0, m]).update_yaxes(tickformat=".2%", range=[0, m])
            c1.plotly_chart(f2.update_layout(height=480, showlegend=False), width="stretch")
            movers = pd.concat([g.nsmallest(10, "NET_REV"), g.nlargest(10, "NET_REV")]).drop_duplicates("NAME") \
                .sort_values("NET_REV")
            c2.plotly_chart(px.bar(movers, x="NET_REV", y="NAME", orientation="h", color="QUADRANT",
                                   color_discrete_map=QUAD_COLORS,
                                   title="Top 10 / bottom 10 net MDR revenue contributors (₱)",
                                   hover_data={"MARGIN_BPS": ":+.1f", "VOLUME": ":,.0f"},
                                   labels={"NET_REV": "Net MDR revenue"})
                            .update_layout(height=480, showlegend=False).update_yaxes(type="category", title=None),
                            width="stretch")

            # ---- repricing view: revenue uplift if below-hurdle merchants were brought to the hurdle
            if below.any():
                gap = g.loc[below].assign(UPLIFT=lambda x: (hurdle - x.MARGIN_BPS) / 1e4 * x.VOLUME)
                st.caption(f"Bringing the {int(below.sum())} below-hurdle entities up to a {hurdle:+.0f} bps margin "
                           f"(same volume) would add **{fmt_php(gap.UPLIFT.sum())}** net MDR revenue for {per_lbl}.")

            tbl = g.sort_values("MARGIN_BPS")[["NAME", "QUADRANT", "VOLUME", "VOL_SHARE", "TXN_COUNT", "TICKET",
                                               "WTD_MDR", "WTD_IC", "MARGIN_BPS", "MDR_REV", "IC_COST", "NET_REV",
                                               "NET_REV_PER_TXN", "MONTHS"]]
            tbl.columns = ["Name", "Quadrant", "Volume", "Vol share", "Txn", "Avg ticket", "Wtd MDR", "Wtd IC",
                           "Margin (bps)", "Gross MDR revenue", "Interchange", "Net MDR revenue",
                           "Net MDR rev / txn", "Active months"]
            st.dataframe(tbl.style.format({"Volume": "{:,.0f}", "Vol share": "{:.2%}", "Txn": "{:,.0f}",
                                           "Avg ticket": "{:,.0f}", "Wtd MDR": "{:.2%}", "Wtd IC": "{:.2%}",
                                           "Margin (bps)": "{:+.1f}", "Gross MDR revenue": "{:,.0f}",
                                           "Interchange": "{:,.0f}", "Net MDR revenue": "{:,.0f}",
                                           "Net MDR rev / txn": "{:,.2f}"}, na_rep="-"),
                         width="stretch", hide_index=True, height=420)
            st.download_button("Download profitability table (CSV)", tbl.to_csv(index=False).encode(),
                               file_name=f"net_mdr_revenue_{seg_p.replace(' ', '_')}_{per_lbl}.csv", key="prof_dl")

            # ---- monthly trend for one entity
            pick_n = st.selectbox("Monthly trend for", g.sort_values("VOLUME", ascending=False).NAME.tolist(),
                                  key=f"prof_pick_{seg_p}_{lvl}")
            keys = lvl_map[lvl]
            sel_row = g[g.NAME == pick_n].iloc[0]
            mask = np.logical_and.reduce([df[k].astype(str) == str(sel_row[k]) for k in keys])
            dsel = df[(df.PF_DM_TAG == seg_p) & mask & (df.YEAR_MONTH <= cur_m)
                      & (df.YEAR_MONTH >= prev_month(cur_m, 12))]
            mt = profit_table(dsel, "YEAR_MONTH").sort_values("YEAR_MONTH")
            c1, c2 = st.columns(2)
            c1.plotly_chart(px.bar(mt, x="YEAR_MONTH", y="NET_REV",
                                   title=f"{pick_n} – monthly net MDR revenue (₱, actual)",
                                   color=(mt.NET_REV >= 0).map({True: "Positive", False: "Negative"}),
                                   color_discrete_map={"Positive": "#2E8B57", "Negative": "#C0392B"},
                                   labels={"NET_REV": "Net MDR revenue"})
                            .update_layout(showlegend=False).update_xaxes(type="category"), width="stretch")
            c2.plotly_chart(rate_line(mt, "YEAR_MONTH", ["WTD_MDR", "WTD_IC"],
                                      f"{pick_n} – weighted MDR vs interchange (gap = margin)"), width="stretch")
            if cur_m == LATEST_M and CUR_FACTOR != 1:
                st.caption(f"{cur_m} is month-to-date (as of {card_asof:%b %d}); net MDR revenue is MTD, "
                           "margins are unaffected.")
