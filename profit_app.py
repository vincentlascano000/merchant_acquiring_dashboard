"""
Merchant Acquiring – Profitability (separate Streamlit app, own resource limit)
Run: streamlit run profit_app.py      Deploy: same repo, main file path = profit_app.py
Requirements: streamlit>=1.50, pandas, plotly, openpyxl

Sections (only the selected one is computed):
  ① Profitability Map   – net MDR revenue = Σ(MDR × volume) − Σ(interchange × volume), quadrants
  ② Repricing Simulator – reprice a selection, assume volume loss, see Δ revenue + break-even volume loss
  ③ Minimum MDR Floor   – floor = interchange + required spread (+ per-txn cost ÷ ticket), by industry / MCC / ticket
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

st.set_page_config(page_title="MA Profitability", layout="wide")

_orig_plotly_chart = _DG.plotly_chart
_CHART_N = [0]


def _plotly_chart_keyed(self, fig, *args, **kwargs):
    if "key" not in kwargs:
        _CHART_N[0] += 1
        kwargs["key"] = f"chart_{_CHART_N[0]}"
    return _orig_plotly_chart(self, fig, *args, **kwargs)


_DG.plotly_chart = _plotly_chart_keyed

SCAN_DIRS = [".", "data", "data/card"]
RR_COLOR = "#E8541E"
CARD_REQUIRED = ["YEAR_MONTH", "PF_DM_TAG", "PF_DM_NAME", "INDUSTRY", "TRXN_CODE", "TOTAL_VOLUME",
                 "TOTAL_COUNT", "MDR", "INTERCHANGE_RATE", "INTERCHANGE_EXPENSE_PHP"]
QRPH_REQUIRED = ["YEAR_MONTH", "VOLUME_POSTED", "TXN_COUNT_POSTED"]
DATE_COL = "MAX_TRAN_DATE"
SUB_COL = "MERCHANT_NAME_ORIG"
MID_COL = "MID"
SEPS = [",", ";", "\t", "|"]
KEEP_COLS = set(CARD_REQUIRED + [DATE_COL, SUB_COL, "CARD_ACCEPTOR_ID", "MCC"])
AGG_KEYS = ["YEAR_MONTH", "PF_DM_TAG", "PF_DM_NAME", SUB_COL, MID_COL, "INDUSTRY", "MCC", "TRXN_CODE"]
MAIN_APP_URL = ""          # optional: paste the main dashboard URL to show a link back in the sidebar
QUAD_COLORS = {"① Protect – high volume, above hurdle": "#2E8B57",
               "② Fix – high volume, below hurdle": "#C0392B",
               "③ Grow – low volume, above hurdle": "#2F80ED",
               "④ Review – low volume, below hurdle": "#F2994A"}


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
        cdf[MID_COL] = _norm_mid(cdf["CARD_ACCEPTOR_ID"]) if "CARD_ACCEPTOR_ID" in cdf.columns else "NO MID"
        if DATE_COL in cdf.columns:
            cdf[DATE_COL] = _parse_dates(cdf[DATE_COL])
        cdf["MDR_X_VOL"] = cdf["MDR"] * cdf["TOTAL_VOLUME"]
        cdf["IC_X_VOL"] = cdf["INTERCHANGE_RATE"] * cdf["TOTAL_VOLUME"]
        cdf["INTERCHANGE_PHP"] = -cdf["INTERCHANGE_EXPENSE_PHP"]
        cdf["MCC"] = (cdf["MCC"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
                      if "MCC" in cdf.columns else "N/A")
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
        return list(_read_any(path, path, nrows=5).columns)
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


def data_max_tran_date(d, ym):
    if DATE_COL not in d.columns:
        return None, f"the loaded card file has no {DATE_COL} column"
    s_all = d.loc[d["YEAR_MONTH"] == ym, DATE_COL]
    s = s_all.dropna()
    if s.empty:
        return None, f"{DATE_COL} is blank / unreadable for {ym}"
    p = pd.Period(ym, "M")
    s_in = s[(s >= p.start_time) & (s <= p.end_time)]
    if s_in.empty:
        return None, f"{DATE_COL} values for {ym} fall outside that month (e.g. {s.max():%Y-%m-%d})"
    return s_in.max().date(), None

# ================================================================ PROFIT-APP HELPERS
def fmt_bp(x):
    return "-" if pd.isna(x) else f"{x:+.1f} bps"


def months_eq(month_list):
    """Months in the period; the latest month counts as as-of day ÷ days-in-month when it is partial."""
    n = float(len(month_list))
    if LATEST_M in month_list and PARTIAL:
        n -= 1 - card_asof.day / DIM_LATEST
    return n


def period_rows(per):
    ym = df["YEAR_MONTH"]
    if per == "YTD":
        mlist = [m for m in MONTHS if m[:4] == str(CY) and m <= cur_m]
    elif per == "L12M":
        mlist = [m for m in MONTHS if prev_month(cur_m, 11) <= m <= cur_m]
    else:
        mlist = [cur_m]
    return df[ym.isin(mlist)], mlist


def _main_attr(d, keys, col):
    t = (d.groupby(keys + [col], dropna=False).TOTAL_VOLUME.sum().reset_index()
         .sort_values("TOTAL_VOLUME", ascending=False).drop_duplicates(keys))
    return t[keys + [col]]


def profit_table(d, keys):
    """Per entity: volume, gross MDR revenue Σ(MDR×vol), interchange Σ(IC×vol), net MDR revenue, margin bps,
    weighted MDR / IC, ticket, main industry / MCC. Actuals (not run-rated)."""
    keys = keys if isinstance(keys, list) else [keys]
    g = d.groupby(keys, dropna=False).agg(
        VOLUME=("TOTAL_VOLUME", "sum"), TXN_COUNT=("TOTAL_COUNT", "sum"),
        MDR_REV=("MDR_X_VOL", "sum"), IC_COST=("IC_X_VOL", "sum"),
        MONTHS=("YEAR_MONTH", "nunique")).reset_index()
    g = g[g.VOLUME > 0].copy()
    if g.empty:
        return g
    g["NET_REV"] = g.MDR_REV - g.IC_COST
    g["WTD_MDR"] = g.MDR_REV / g.VOLUME
    g["WTD_IC"] = g.IC_COST / g.VOLUME
    g["MARGIN_BPS"] = g.NET_REV / g.VOLUME * 1e4
    g["TICKET"] = g.VOLUME / g.TXN_COUNT.where(g.TXN_COUNT > 0)
    g["NAME"] = g[keys].astype(str).agg(" › ".join, axis=1)
    if "INDUSTRY" not in keys:
        g = g.merge(_main_attr(d, keys, "INDUSTRY"), on=keys, how="left")
    if "MCC" not in keys:
        g = g.merge(_main_attr(d, keys, "MCC"), on=keys, how="left")
    return g


def assign_quadrants(g, pareto_pct, hurdle_bps):
    g = g.sort_values("VOLUME", ascending=False).copy()
    share = g.VOLUME / g.VOLUME.sum()
    g["VOL_SHARE"] = share
    g["HIGH_VOL"] = (share.cumsum() - share) < pareto_pct / 100
    vol_cut = g.loc[g.HIGH_VOL, "VOLUME"].min() if g.HIGH_VOL.any() else g.VOLUME.max()
    prof = g.MARGIN_BPS >= hurdle_bps
    q = list(QUAD_COLORS)
    g["QUADRANT"] = np.select([g.HIGH_VOL & prof, g.HIGH_VOL & ~prof, ~g.HIGH_VOL & prof], q[:3], default=q[3])
    return g, vol_cut


def seg_level_picker(prefix, allow_all=False):
    cA, cB = st.columns([1.5, 2])
    opts = (["All"] if allow_all else []) + ["Direct Merchant", "Payment Facilitator"]
    seg = cA.radio("Segment", opts, horizontal=True, key=f"{prefix}_seg")
    lv = {"Merchant / PF": ["PF_DM_NAME"]}
    if seg == "Payment Facilitator":
        lv.update({"MID": ["PF_DM_NAME", MID_COL], "Sub-merchant": ["PF_DM_NAME", SUB_COL]})
    lvl = cB.radio("Level", list(lv), horizontal=True, key=f"{prefix}_lvl_{seg}")
    d = dper if seg == "All" else dper[dper.PF_DM_TAG == seg]
    return seg, lv[lvl], d


def money_table(t, fmt, height=None):
    kw = dict(width="stretch", hide_index=True)
    if height:
        kw["height"] = height
    st.dataframe(t.style.format(fmt, na_rep="-"), **kw)


# ================================================================ SIDEBAR
st.sidebar.title("Profitability")
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

MONTHS = sorted(card_raw["YEAR_MONTH"].dropna().unique())
LATEST_M = MONTHS[-1]
_lp = pd.Period(LATEST_M, "M")
DIM_LATEST = calendar.monthrange(_lp.year, _lp.month)[1]
auto_dt, _why = data_max_tran_date(card_raw, LATEST_M)
card_asof = auto_dt or date(_lp.year, _lp.month, DIM_LATEST)
PARTIAL = card_asof.day < DIM_LATEST
st.sidebar.caption(f"Latest month {LATEST_M}: data as of **{card_asof:%b %d, %Y}**"
                   + ("" if auto_dt else f" (no {DATE_COL} – assumed full month)"))
cur_m = st.sidebar.selectbox("Reporting month", MONTHS, index=len(MONTHS) - 1, key="cur_m")
include_credits = st.sidebar.checkbox("Net off credit (refund) transactions", value=True)
df = card_raw if include_credits else card_raw[card_raw["TRXN_CODE"] == "Sale Transaction"]
cp = pd.Period(cur_m, "M")
CY, LY, END_MM = cp.year, cp.year - 1, cp.month
ytd_lbl = f"Jan–{calendar.month_abbr[END_MM]}"

PER = st.sidebar.radio("Period", ["YTD", "L12M", "MONTH"], key="per",
                       format_func={"YTD": f"YTD {CY} ({ytd_lbl})", "L12M": "Last 12 months",
                                    "MONTH": f"{cur_m} only"}.get)
dper, PER_MONTHS = period_rows(PER)
PER_LBL = {"YTD": f"YTD {CY}", "L12M": f"L12M to {cur_m}", "MONTH": cur_m}[PER]
M_EQ = months_eq(PER_MONTHS)
ANN = 12 / M_EQ if M_EQ else float("nan")
st.sidebar.caption(f"Period = {len(PER_MONTHS)} month(s) ≈ {M_EQ:.2f} full months → annualise ×{ANN:.2f}")

SECTIONS = ["① Profitability Map", "② Repricing Simulator", "③ Minimum MDR Floor"]
st.title("Merchant Acquiring – Profitability")
SEC = st.radio("Section", SECTIONS, horizontal=True, key="sec", label_visibility="collapsed")
st.caption(f"Basis: **net MDR revenue = Σ(MDR × volume) − Σ(interchange rate × volume)** · {PER_LBL} · "
           "actual volumes (not run-rated); annualised figures scale by 12 ÷ months in period. "
           "Other costs (txn, MID, processing, license fees) are not included.")

# ================================================================ ① PROFITABILITY MAP
if SEC == SECTIONS[0]:
    cA, cB = st.columns(2)
    pareto = cA.select_slider("High volume = top X% of segment volume", [50, 60, 70, 80, 90, 95], value=80,
                              key="prof_pareto")
    hurdle = cB.number_input("Margin hurdle (bps of volume)", value=0.0, step=5.0, key="prof_hurdle")

    rows = []
    for s in ["Direct Merchant", "Payment Facilitator"]:
        g = profit_table(dper[dper.PF_DM_TAG == s], "PF_DM_NAME")
        if g.empty:
            continue
        below = g.MARGIN_BPS < hurdle
        rows.append({"Segment": s, "Merchants": len(g), "Volume": g.VOLUME.sum(), "Gross MDR revenue": g.MDR_REV.sum(),
                     "Interchange": g.IC_COST.sum(), "Net MDR revenue": g.NET_REV.sum(),
                     "Wtd MDR": g.MDR_REV.sum() / g.VOLUME.sum(), "Wtd IC": g.IC_COST.sum() / g.VOLUME.sum(),
                     "Margin (bps)": g.NET_REV.sum() / g.VOLUME.sum() * 1e4, "# below hurdle": int(below.sum()),
                     "Vol share below hurdle": g.loc[below, "VOLUME"].sum() / g.VOLUME.sum()})
    if rows:
        st.subheader(f"Direct Merchants vs Payment Facilitators – {PER_LBL}")
        money_table(pd.DataFrame(rows), {"Volume": "{:,.0f}", "Gross MDR revenue": "{:,.0f}", "Interchange": "{:,.0f}",
                                         "Net MDR revenue": "{:,.0f}", "Wtd MDR": "{:.2%}", "Wtd IC": "{:.2%}",
                                         "Margin (bps)": "{:+.1f}", "Vol share below hurdle": "{:.1%}"})
    st.divider()
    seg_p, keys, dseg = seg_level_picker("map")
    g = profit_table(dseg, keys)
    if g.empty:
        st.info("No volume for this segment / period.")
    else:
        g, vol_cut = assign_quadrants(g, pareto, hurdle)
        tot_v, tot_r = g.VOLUME.sum(), g.NET_REV.sum()
        below = g.MARGIN_BPS < hurdle
        c = st.columns(4)
        c[0].metric(f"Volume ({PER_LBL})", fmt_php(tot_v))
        c[1].metric("Net MDR revenue", fmt_php(tot_r), f"{tot_r / tot_v * 1e4:+.1f} bps of volume", delta_color="off")
        c[2].metric(f"Below hurdle ({hurdle:+.0f} bps)", f"{int(below.sum()):,} of {len(g):,}",
                    f"{g.loc[below, 'VOLUME'].sum() / tot_v:.1%} of volume", delta_color="off")
        c[3].metric("Net MDR rev from below-hurdle", fmt_php(g.loc[below, "NET_REV"].sum()), delta_color="off")
        lo, hi = np.nanpercentile(g.MARGIN_BPS, [2, 98])
        lo, hi = min(lo, hurdle - 10), max(hi, hurdle + 10)
        pad = (hi - lo) * 0.08
        g["Y"] = g.MARGIN_BPS.clip(lo - pad, hi + pad)
        g["Shown"] = np.where(g.MARGIN_BPS.between(lo - pad, hi + pad), "in range", "clipped (outlier)")
        g["SIZE"] = g.NET_REV.abs()
        fig = px.scatter(g, x="VOLUME", y="Y", color="QUADRANT", size="SIZE", size_max=40, log_x=True,
                         symbol="Shown", symbol_map={"in range": "circle", "clipped (outlier)": "triangle-up"},
                         color_discrete_map=QUAD_COLORS, category_orders={"QUADRANT": list(QUAD_COLORS)},
                         hover_name="NAME",
                         hover_data={"VOLUME": ":,.0f", "NET_REV": ":,.0f", "MARGIN_BPS": ":+.1f", "WTD_MDR": ":.2%",
                                     "WTD_IC": ":.2%", "Y": False, "SIZE": False, "Shown": False, "QUADRANT": False},
                         labels={"VOLUME": f"Volume ({PER_LBL}, log)", "Y": "Net MDR margin (bps)"},
                         title=f"Profitability map – {seg_p} · bubble = |net MDR revenue|")
        fig.add_hline(y=hurdle, line_dash="dash", line_color="grey", annotation_text=f"hurdle {hurdle:+.0f} bps")
        fig.add_vline(x=vol_cut, line_dash="dash", line_color="grey",
                      annotation_text=f"high-volume cut-off {fmt_php(vol_cut)}")
        st.plotly_chart(fig.update_layout(height=600, legend_title_text=""), width="stretch")
        qs = (g.groupby("QUADRANT").agg(Entities=("NAME", "count"), Volume=("VOLUME", "sum"), Net=("NET_REV", "sum"))
              .reindex(list(QUAD_COLORS)).fillna(0).reset_index())
        qs["Vol share"] = qs.Volume / tot_v
        qs["Margin (bps)"] = qs.Net / qs.Volume.where(qs.Volume > 0) * 1e4
        money_table(qs.rename(columns={"QUADRANT": "Quadrant", "Net": "Net MDR revenue"}),
                    {"Entities": "{:,.0f}", "Volume": "{:,.0f}", "Net MDR revenue": "{:,.0f}", "Vol share": "{:.1%}",
                     "Margin (bps)": "{:+.1f}"})
        tbl = g.sort_values("MARGIN_BPS")[["NAME", "QUADRANT", "INDUSTRY", "VOLUME", "TICKET", "WTD_MDR", "WTD_IC",
                                           "MARGIN_BPS", "NET_REV"]]
        tbl.columns = ["Name", "Quadrant", "Main industry", "Volume", "Avg ticket", "Wtd MDR", "Wtd IC",
                       "Margin (bps)", "Net MDR revenue"]
        money_table(tbl, {"Volume": "{:,.0f}", "Avg ticket": "{:,.0f}", "Wtd MDR": "{:.2%}", "Wtd IC": "{:.2%}",
                          "Margin (bps)": "{:+.1f}", "Net MDR revenue": "{:,.0f}"}, height=380)

# ================================================================ ② REPRICING SIMULATOR
if SEC == SECTIONS[1]:
    st.caption("Pick who to reprice and how, then set the share of volume you expect to lose from merchants whose "
               "MDR goes **up**. New net MDR revenue = (new MDR − current weighted interchange) × retained volume. "
               "Interchange rates are assumed unchanged.")
    seg_s, keys, dseg = seg_level_picker("sim", allow_all=True)
    g = profit_table(dseg, keys)
    if g.empty:
        st.info("No volume for this segment / period.")
        st.stop()
    g, _ = assign_quadrants(g, 80, 0.0)

    st.markdown("##### 1 · Who to reprice")
    cA, cB = st.columns([1.2, 2.8])
    mode = cA.radio("Select by", ["Margin below threshold", "Industry", "Quadrant", "Specific names",
                                  "Everyone in segment"], key="sim_mode")
    if mode == "Margin below threshold":
        thr = cB.number_input("Reprice entities with net MDR margin below (bps)", value=0.0, step=5.0, key="sim_thr")
        sel = g.MARGIN_BPS < thr
    elif mode == "Industry":
        inds = cB.multiselect("Industries", sorted(g.INDUSTRY.dropna().unique()), key="sim_ind")
        sel = g.INDUSTRY.isin(inds)
    elif mode == "Quadrant":
        qd = cB.multiselect("Quadrants (80% volume cut-off, 0 bps hurdle)", list(QUAD_COLORS),
                            default=[list(QUAD_COLORS)[1]], key="sim_quad")
        sel = g.QUADRANT.isin(qd)
    elif mode == "Specific names":
        names = cB.multiselect("Entities", g.sort_values("VOLUME", ascending=False).NAME.tolist(), key="sim_names")
        sel = g.NAME.isin(names)
    else:
        sel = pd.Series(True, index=g.index)
    min_vol = cB.number_input("…and with period volume of at least (₱)", value=0.0, step=100000.0, key="sim_minvol")
    sel &= g.VOLUME >= min_vol

    st.markdown("##### 2 · Pricing action")
    cA, cB, cC = st.columns([1.6, 1.2, 1.2])
    action = cA.radio("Action", ["Add bps to current MDR", "Set MDR to a fixed rate",
                                 "Raise to floor (interchange + spread)"], key="sim_act")
    if action == "Add bps to current MDR":
        x = cB.number_input("Increase (bps)", value=10.0, step=5.0, key="sim_add")
        new = g.WTD_MDR + x / 1e4
    elif action == "Set MDR to a fixed rate":
        x = cB.number_input("New MDR (%)", value=2.00, step=0.05, format="%.2f", key="sim_set")
        new = pd.Series(x / 100, index=g.index)
    else:
        x = cB.number_input("Minimum spread over interchange (bps)", value=30.0, step=5.0, key="sim_spread")
        new = np.maximum(g.WTD_MDR, g.WTD_IC + x / 1e4)
    attr = cC.slider("Volume lost from merchants whose MDR rises (%)", 0, 60, 10, key="sim_attr") / 100

    g["NEW_MDR"] = np.where(sel, new, g.WTD_MDR)
    g["CHANGED"] = sel & (g.NEW_MDR.round(8) != g.WTD_MDR.round(8))
    g["UP"] = g.NEW_MDR > g.WTD_MDR + 1e-10
    g["NEW_VOL"] = np.where(g.UP, g.VOLUME * (1 - attr), g.VOLUME)
    g["NEW_REV"] = (g.NEW_MDR - g.WTD_IC) * g.NEW_VOL
    g["DELTA"] = g.NEW_REV - g.NET_REV

    ch = g[g.CHANGED]
    cur_tot, new_tot = g.NET_REV.sum(), g.NEW_REV.sum()
    st.markdown("##### 3 · Result")
    if ch.empty:
        st.info("No entities are repriced with these settings.")
    else:
        full = ((ch.NEW_MDR - ch.WTD_IC) * ch.VOLUME).sum()
        cur_ch = ch.NET_REV.sum()
        if full <= 0:
            be_txt = "n/a – new prices still below interchange"
        elif cur_ch <= 0:
            be_txt = "any – repriced group earns ≤ 0 today"
        else:
            be_txt = f"{max(0.0, 1 - cur_ch / full):.1%}"
        c = st.columns(5)
        c[0].metric("Entities repriced", f"{len(ch):,} of {len(g):,}",
                    f"{ch.VOLUME.sum() / g.VOLUME.sum():.1%} of volume", delta_color="off")
        c[1].metric(f"Net MDR revenue – current ({PER_LBL})", fmt_php(cur_tot),
                    f"{cur_tot / g.VOLUME.sum() * 1e4:+.1f} bps", delta_color="off")
        c[2].metric("Net MDR revenue – simulated", fmt_php(new_tot),
                    f"{fmt_php(new_tot - cur_tot)} ({(new_tot - cur_tot) / abs(cur_tot):+.1%})" if cur_tot else None)
        c[3].metric("Annualised uplift", fmt_php((new_tot - cur_tot) * ANN), f"×{ANN:.2f} of period", delta_color="off")
        c[4].metric("Break-even volume loss", be_txt, "above this, repricing loses money", delta_color="off")

        # sensitivity
        xs = np.linspace(0, 0.6, 61)
        base_unch = g.loc[~g.UP, "NEW_REV"].sum()
        up = g[g.UP]
        ys = [base_unch + ((up.NEW_MDR - up.WTD_IC) * up.VOLUME * (1 - a)).sum() for a in xs]
        sens = pd.DataFrame({"Volume lost": xs, "Simulated net MDR revenue": ys})
        f1 = px.line(sens, x="Volume lost", y="Simulated net MDR revenue",
                     title="Sensitivity – simulated net MDR revenue vs volume lost from repriced merchants")
        f1.add_hline(y=cur_tot, line_dash="dash", line_color="grey", annotation_text="current")
        f1.add_vline(x=attr, line_dash="dot", line_color=RR_COLOR, annotation_text="assumed")
        f1.update_xaxes(tickformat=".0%")
        c1, c2 = st.columns(2)
        c1.plotly_chart(f1.update_layout(height=430), width="stretch")
        top = ch.reindex(ch.DELTA.abs().sort_values(ascending=False).index).head(15).sort_values("DELTA")
        c2.plotly_chart(px.bar(top, x="DELTA", y="NAME", orientation="h", title="Largest revenue changes (₱, period)",
                               color=(top.DELTA >= 0).map({True: "Gain", False: "Loss"}),
                               color_discrete_map={"Gain": "#2E8B57", "Loss": "#C0392B"},
                               hover_data={"WTD_MDR": ":.2%", "NEW_MDR": ":.2%", "WTD_IC": ":.2%"})
                        .update_layout(height=430, showlegend=False).update_yaxes(type="category", title=None),
                        width="stretch")
        tbl = ch.sort_values("DELTA", ascending=False)[
            ["NAME", "INDUSTRY", "VOLUME", "NEW_VOL", "WTD_IC", "WTD_MDR", "NEW_MDR", "MARGIN_BPS", "NET_REV",
             "NEW_REV", "DELTA"]].copy()
        tbl.insert(8, "NEW_BPS", (tbl.NEW_MDR - tbl.WTD_IC) * 1e4)
        tbl["ANN_DELTA"] = tbl.DELTA * ANN
        tbl.columns = ["Name", "Main industry", "Volume", "Retained volume", "Wtd IC", "Current MDR", "New MDR",
                       "Current margin (bps)", "New margin (bps)", "Current net MDR rev", "New net MDR rev",
                       "Δ (period)", "Δ (annualised)"]
        money_table(tbl, {"Volume": "{:,.0f}", "Retained volume": "{:,.0f}", "Wtd IC": "{:.2%}", "Current MDR": "{:.2%}",
                          "New MDR": "{:.2%}", "Current margin (bps)": "{:+.1f}", "New margin (bps)": "{:+.1f}",
                          "Current net MDR rev": "{:,.0f}", "New net MDR rev": "{:,.0f}", "Δ (period)": "{:+,.0f}",
                          "Δ (annualised)": "{:+,.0f}"}, height=400)
        st.download_button("Download repricing scenario (CSV)", tbl.to_csv(index=False).encode(),
                           file_name=f"repricing_{PER_LBL.replace(' ', '_')}.csv", key="sim_dl")
        st.caption("Break-even volume loss = 1 − current net MDR revenue ÷ net MDR revenue at the new price with full "
                   "volume (for the repriced group). Merchants whose MDR goes down (fixed-rate action) keep their "
                   "volume in this model.")

# ================================================================ ③ MINIMUM MDR FLOOR
if SEC == SECTIONS[2]:
    st.caption("Floor MDR for each merchant = **its own weighted interchange + required spread + per-transaction "
               "cost ÷ its average ticket**. Set the per-transaction cost to 0 to use the spread over interchange only.")
    seg_f, keys, dseg = seg_level_picker("floor", allow_all=True)
    cA, cB, cC = st.columns(3)
    spread = cA.number_input("Required spread over interchange (bps)", value=30.0, step=5.0, key="fl_spread")
    ptx = cB.number_input("Per-transaction cost to recover (₱, optional)", value=0.0, step=0.5, key="fl_ptx")
    grp = cC.radio("Summarise floor by", ["INDUSTRY", "MCC", "TICKET_BAND"], horizontal=True, key="fl_grp",
                   format_func={"INDUSTRY": "Industry", "MCC": "MCC", "TICKET_BAND": "Ticket size band"}.get)
    g = profit_table(dseg, keys)
    if g.empty:
        st.info("No volume for this segment / period.")
        st.stop()
    g["FLOOR"] = g.WTD_IC + spread / 1e4 + (ptx / g.TICKET).fillna(0)
    g["GAP_BPS"] = (g.WTD_MDR - g.FLOOR) * 1e4
    g["BELOW"] = g.GAP_BPS < 0
    g["UPLIFT"] = np.where(g.BELOW, (g.FLOOR - g.WTD_MDR) * g.VOLUME, 0.0)
    bands = [0, 500, 2000, 10000, 50000, np.inf]
    blabels = ["< ₱500", "₱500–2K", "₱2K–10K", "₱10K–50K", "≥ ₱50K"]
    g["TICKET_BAND"] = pd.cut(g.TICKET, bands, labels=blabels, right=False).astype(str)

    below = g[g.BELOW]
    c = st.columns(4)
    c[0].metric("Entities below floor", f"{len(below):,} of {len(g):,}",
                f"{below.VOLUME.sum() / g.VOLUME.sum():.1%} of volume", delta_color="off")
    c[1].metric("Vol-weighted floor MDR", f"{(g.FLOOR * g.VOLUME).sum() / g.VOLUME.sum():.2%}",
                f"current {g.MDR_REV.sum() / g.VOLUME.sum():.2%}", delta_color="off")
    c[2].metric(f"Uplift if all raised to floor ({PER_LBL})", fmt_php(below.UPLIFT.sum()), "no volume loss",
                delta_color="off")
    c[3].metric("Annualised uplift", fmt_php(below.UPLIFT.sum() * ANN), f"×{ANN:.2f}", delta_color="off")

    def _grp(t):
        v = t.VOLUME.sum()
        return pd.Series({
            "Entities": len(t), "Volume": v, "Avg ticket": v / t.TXN_COUNT.sum() if t.TXN_COUNT.sum() else np.nan,
            "Wtd IC": t.IC_COST.sum() / v, "Current wtd MDR": t.MDR_REV.sum() / v,
            "Floor (vol-wtd)": (t.FLOOR * t.VOLUME).sum() / v,
            "Floor (90th pct merchant)": np.nanpercentile(t.FLOOR, 90),
            "Min current MDR": t.WTD_MDR.min(),
            "# below floor": int(t.BELOW.sum()), "Vol below floor": t.loc[t.BELOW, "VOLUME"].sum() / v,
            "Uplift to floor": t.UPLIFT.sum()})

    gs = g.groupby(grp, dropna=False).apply(_grp, include_groups=False).reset_index()
    if grp == "TICKET_BAND":
        gs[grp] = pd.Categorical(gs[grp], [b for b in blabels if b in set(gs[grp])] +
                                 [b for b in gs[grp] if b not in blabels], ordered=True)
        gs = gs.sort_values(grp)
    else:
        gs = gs.sort_values("Volume", ascending=False)
    gs["Headroom (bps)"] = (gs["Current wtd MDR"] - gs["Floor (vol-wtd)"]) * 1e4
    gname = {"INDUSTRY": "Industry", "MCC": "MCC", "TICKET_BAND": "Ticket band"}[grp]
    st.subheader(f"Suggested MDR floor by {gname.lower()} – {seg_f}, {PER_LBL}")
    money_table(gs.rename(columns={grp: gname}),
                {"Entities": "{:,.0f}", "Volume": "{:,.0f}", "Avg ticket": "{:,.0f}", "Wtd IC": "{:.2%}",
                 "Current wtd MDR": "{:.2%}", "Floor (vol-wtd)": "{:.2%}", "Floor (90th pct merchant)": "{:.2%}",
                 "Min current MDR": "{:.2%}", "Vol below floor": "{:.1%}", "Uplift to floor": "{:,.0f}",
                 "Headroom (bps)": "{:+.1f}"}, height=360)
    st.caption("**Floor (vol-wtd)** = typical floor for the group. **Floor (90th pct merchant)** = a conservative "
               "single floor that covers 90% of merchants in the group (useful as a published minimum). "
               "**Headroom** = current weighted MDR − vol-weighted floor.")

    top = gs.head(20) if grp != "TICKET_BAND" else gs
    long = pd.concat([pd.DataFrame({gname: top[grp].astype(str), "Rate": top[c_], "Series": c_})
                      for c_ in ["Wtd IC", "Floor (vol-wtd)", "Current wtd MDR"]])
    c1, c2 = st.columns(2)
    c1.plotly_chart(px.bar(long, x=gname, y="Rate", color="Series", barmode="group",
                           color_discrete_map={"Wtd IC": "#A9B8C9", "Floor (vol-wtd)": RR_COLOR,
                                               "Current wtd MDR": "#1F4E79"},
                           title=f"Interchange vs floor vs current MDR by {gname.lower()}")
                    .update_yaxes(tickformat=".2%").update_xaxes(type="category")
                    .update_layout(height=460, legend_title_text=""), width="stretch")
    gg = g[g.TICKET > 0].copy()
    gg["Status"] = gg.BELOW.map({True: "Below floor", False: "At / above floor"})
    gg["GAP_PLOT"] = gg.GAP_BPS.clip(*np.nanpercentile(gg.GAP_BPS, [2, 98]))
    c2.plotly_chart(px.scatter(gg, x="TICKET", y="GAP_PLOT", color="Status", size="VOLUME", size_max=30, log_x=True,
                               color_discrete_map={"Below floor": "#C0392B", "At / above floor": "#2E8B57"},
                               hover_name="NAME", hover_data={"GAP_BPS": ":+.1f", "WTD_MDR": ":.2%", "FLOOR": ":.2%",
                                                              "GAP_PLOT": False, "VOLUME": ":,.0f"},
                               labels={"TICKET": "Average ticket (₱, log)", "GAP_PLOT": "Current MDR − floor (bps)"},
                               title="Gap to floor by ticket size – small tickets need a higher floor "
                                     "when per-txn cost > 0")
                    .add_hline(y=0, line_dash="dash", line_color="grey")
                    .update_layout(height=460, legend_title_text=""), width="stretch")

    st.subheader("Entities priced below their floor")
    bt = below.sort_values("UPLIFT", ascending=False)[["NAME", "INDUSTRY", "MCC", "VOLUME", "TICKET", "WTD_IC",
                                                        "WTD_MDR", "FLOOR", "GAP_BPS", "UPLIFT"]].copy()
    bt["ANN"] = bt.UPLIFT * ANN
    bt.columns = ["Name", "Main industry", "Main MCC", "Volume", "Avg ticket", "Wtd IC", "Current MDR", "Floor MDR",
                  "Gap (bps)", "Uplift to floor (period)", "Uplift (annualised)"]
    money_table(bt, {"Volume": "{:,.0f}", "Avg ticket": "{:,.0f}", "Wtd IC": "{:.2%}", "Current MDR": "{:.2%}",
                     "Floor MDR": "{:.2%}", "Gap (bps)": "{:+.1f}", "Uplift to floor (period)": "{:,.0f}",
                     "Uplift (annualised)": "{:,.0f}"}, height=380)
    cA, cB = st.columns(2)
    cA.download_button("Download floor by group (CSV)", gs.to_csv(index=False).encode(),
                       file_name=f"mdr_floor_by_{grp.lower()}.csv", key="fl_dl1")
    cB.download_button("Download below-floor list (CSV)", bt.to_csv(index=False).encode(),
                       file_name="below_floor_merchants.csv", key="fl_dl2")
    st.caption("Uplift assumes no volume loss – use ② Repricing Simulator → 'Raise to floor' to test attrition.")
