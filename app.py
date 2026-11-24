"""
Merchant Acquiring Dashboard (Total = Card PF/DM + QRPH)
Run: streamlit run app.py
Requirements: streamlit>=1.50, pandas, plotly, openpyxl

Data sources - found automatically in the repo (root, data/, data/qrph/, data/card/) by COLUMN HEADERS,
or uploaded in the sidebar (also auto-detected):
  * Card MIS : YEAR_MONTH, [MAX_TRAN_DATE -> automatic as-of date for the run-rate], PF_DM_TAG, PF_DM_NAME,
               MERCHANT_NAME_ORIG, CARD_ACCEPTOR_ID (MID), INDUSTRY, TRXN_CODE, TOTAL_VOLUME, TOTAL_COUNT, MDR,
               INTERCHANGE_RATE, INTERCHANGE_EXPENSE_PHP, [NET_INCOME -> Profitability Map tab]
  * QRPH     : YEAR_MONTH, VOLUME_POSTED, TXN_COUNT_POSTED (other columns ignored)

Run-rate rule (per source): ONLY each source's latest month is run-rated (MTD x days / as-of day).
  Card as-of day = latest MAX_TRAN_DATE within the latest YEAR_MONTH (read from the file; can be overridden).
Every prior month (incl. MoM / YoY / YTD LY bases) = actual, simple summation.

Chart convention for the run-rated month:
  * Bars : solid = MTD actual, hatched lighter segment on top = run-rate uplift, label "RR ₱x" = full run-rate
  * Lines: solid line = actuals, dotted line + ★ = run-rate, open circle = MTD actual
Rate charts (MDR / interchange): y-axis spans at least RATE_MIN_SPAN_BPS so sub-bp noise doesn't look like movement.
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

st.set_page_config(page_title="Merchant Acquiring Dashboard", layout="wide")

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
NI_COL = "NET_INCOME"
SUB_COL = "MERCHANT_NAME_ORIG"
MID_COL = "MID"
MON = calendar.month_abbr
RR_LEGEND = ("Run-rated month: **solid** = MTD actual, **hatched / dotted ★** = run-rate, "
             "**open ○** = MTD actual point. Ratios (MDR, interchange rate, ticket size) are not run-rated.")
SEPS = [",", ";", "\t", "|"]
QUAD_COLORS = {"① Protect – high volume, profitable": "#2E8B57",
               "② Fix – high volume, below hurdle": "#C0392B",
               "③ Grow – low volume, profitable": "#2F80ED",
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


def _read_any(src, name):
    if str(name).lower().endswith((".xlsx", ".xls")):
        raw = pd.read_excel(src, header=None, nrows=15)
        skip = next((i for i in range(len(raw))
                     if "YEAR_MONTH" in {_clean_col(x) for x in raw.iloc[i].astype(str)}), 0)
        if hasattr(src, "seek"):
            src.seek(0)
        return pd.read_excel(src, skiprows=skip)
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
            return pd.read_csv(src, encoding=enc, low_memory=False, sep=sep, skiprows=skip,
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


@st.cache_data(show_spinner="Loading files…")
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
        num_cols = ["TOTAL_VOLUME", "TOTAL_COUNT", "MDR", "INTERCHANGE_RATE", "INTERCHANGE_EXPENSE_PHP"]
        if NI_COL in cdf.columns:
            num_cols.append(NI_COL)
        for c in num_cols:
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


# ================================================================ GENERIC HELPERS
def summarize(d, by):
    g = d.groupby(by, dropna=False).agg(
        VOLUME=("TOTAL_VOLUME", "sum"), TXN_COUNT=("TOTAL_COUNT", "sum"),
        MDR_X_VOL=("MDR_X_VOL", "sum"), IC_X_VOL=("IC_X_VOL", "sum"),
        INTERCHANGE_PHP=("INTERCHANGE_PHP", "sum"),
        BLENDED_MDR=("MDR", "mean"), BLENDED_IC=("INTERCHANGE_RATE", "mean"),
    ).reset_index()
    g["WEIGHTED_MDR"] = g["MDR_X_VOL"] / g["VOLUME"].where(g["VOLUME"] != 0)
    g["WEIGHTED_IC"] = g["IC_X_VOL"] / g["VOLUME"].where(g["VOLUME"] != 0)
    return g.drop(columns=["MDR_X_VOL", "IC_X_VOL"])


def prev_month(ym, n=1):
    return str(pd.Period(ym, "M") - n)


def top_n_others(d, value_col, n=CAT_TOP_N, key="INDUSTRY"):
    ranked = d[d[key] != OTHERS_LABEL].sort_values(value_col, ascending=False)
    keep = ranked[key].head(n).tolist()
    out = d.assign(**{key: d[key].where(d[key].isin(keep), OTHERS_LABEL)})
    out = out.groupby(key, as_index=False)[value_col].sum()
    out["_o"] = out[key].eq(OTHERS_LABEL)
    return out.sort_values(["_o", value_col], ascending=[True, False]).drop(columns="_o")


def pie(d, value_col, title, names="INDUSTRY"):
    fig = px.pie(d, names=names, values=value_col, hole=0.45, title=title,
                 category_orders={names: d[names].tolist()})
    fig.update_traces(textinfo="percent", textposition="inside", insidetextorientation="horizontal", sort=False)
    fig.update_layout(legend=dict(font=dict(size=10)), uniformtext_minsize=9, uniformtext_mode="hide",
                      margin=dict(t=60, b=20, l=20, r=20), height=480)
    return fig


def fmt_php(x):
    if pd.isna(x):
        return "-"
    for div, s in [(1e9, "B"), (1e6, "M"), (1e3, "K")]:
        if abs(x) >= div:
            return f"₱{x/div:,.2f}{s}"
    return f"₱{x:,.0f}"


def fmt_num(x):
    if pd.isna(x):
        return "-"
    for div, s in [(1e6, "M"), (1e3, "K")]:
        if abs(x) >= div:
            return f"{x/div:,.1f}{s}"
    return f"{x:,.0f}"


def fmt_int(x):
    return "-" if pd.isna(x) else f"{x:,.0f}"


def fmt_pct(x):
    return "n/a" if pd.isna(x) else f"{x:+.1%}"


def fmt_bps(x):
    return "n/a" if pd.isna(x) else f"{x:+.1f}"


def fmt_rate(x):
    return "-" if pd.isna(x) else f"{x:.2%}"


def safe_div(a, b):
    return a / b if pd.notna(a) and pd.notna(b) and b != 0 else float("nan")


def month_picker_date(label, ym, default_day, key):
    p = pd.Period(ym, "M")
    dim = calendar.monthrange(p.year, p.month)[1]
    d = st.sidebar.date_input(label, value=date(p.year, p.month, min(default_day, dim)),
                              min_value=date(p.year, p.month, 1), max_value=date(p.year, p.month, dim), key=key)
    return dim / d.day, d


# ================================================================ RATE CHARTS
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


# ================================================================ RUN-RATE CHART HELPERS
def expand_rr(g, val):
    a = g[g["_F"] == 1].assign(BASIS="Actual")
    r = g[g["_F"] != 1]
    mtd = r.assign(BASIS="Actual (MTD)")
    up = r.assign(BASIS="Run-rate uplift")
    up[val] = r[val] * (r["_F"] - 1)
    return pd.concat([a, mtd, up], ignore_index=True)


def bar_rr(g, cat, val, title, color=None, cmap=None, horizontal=False, fmt=fmt_php, height=None):
    long = expand_rr(g, val)
    xy = dict(x=val, y=cat, orientation="h") if horizontal else dict(x=cat, y=val)
    pat = dict(pattern_shape="BASIS",
               pattern_shape_map={"Actual": "", "Actual (MTD)": "", "Run-rate uplift": "/"})
    if color is None:
        fig = px.bar(long, **xy, color="BASIS", title=title, category_orders={"BASIS": BASIS_ORDER},
                     color_discrete_map={"Actual": "#1F4E79", "Actual (MTD)": "#1F4E79", "Run-rate uplift": RR_COLOR},
                     **pat)
    else:
        fig = px.bar(long, **xy, color=color, title=title, color_discrete_map=cmap or {},
                     category_orders={"BASIS": BASIS_ORDER}, **pat)
    for tr in fig.data:
        if "Run-rate uplift" in (tr.name or ""):
            tr.marker.opacity = 0.55
    fig.update_layout(barmode="relative", legend_title_text="")
    rr = g[g["_F"] != 1]
    if len(rr):
        tot = (rr[val] * rr["_F"]).groupby(rr[cat]).sum()
        for k, v in tot.items():
            if horizontal:
                fig.add_annotation(x=v, y=k, text=f"RR {fmt(v)}", showarrow=False, xanchor="left", xshift=4,
                                   font=dict(color=RR_COLOR, size=10))
            else:
                fig.add_annotation(x=k, y=v, text=f"RR {fmt(v)}", showarrow=False, yshift=10,
                                   font=dict(color=RR_COLOR, size=10))
    if horizontal:
        fig.update_yaxes(autorange="reversed", type="category")
    else:
        fig.update_xaxes(type="category")
    if height:
        fig.update_layout(height=height)
    return fig


def line_rr(g, x, y, title, color=None, order=None, cum=False, y_rr=None, tickformat=None, cmap=None):
    g = g.copy().sort_values("YEAR_MONTH")
    if color is None:
        g["Series"] = "Total"
        color = "Series"
    g["_YRR"] = g[y] * g["_F"] if y_rr is None else g[y_rr]
    if cum:
        g["_CUMA"] = g.groupby(color)[y].cumsum()
        g["_YRR"] = g["_CUMA"] + (g["_YRR"] - g[y])
        y = "_CUMA"
    is_rr = g["_F"] != 1
    base = g[~is_rr]
    fig = px.line(base, x=x, y=y, color=color, markers=True, title=title,
                  category_orders=order or {}, color_discrete_map=cmap or {},
                  color_discrete_sequence=px.colors.qualitative.Plotly)
    cols = {tr.name: tr.line.color for tr in fig.data}
    pal = px.colors.qualitative.Plotly
    for i, (_, r) in enumerate(g[is_rr].iterrows()):
        grp = str(r[color])
        c = cols.get(grp) or (cmap or {}).get(grp) or pal[i % len(pal)]
        prev = base[base[color].astype(str) == grp].tail(1)
        xs = ([prev[x].iloc[0]] if len(prev) else []) + [r[x]]
        ys = ([prev[y].iloc[0]] if len(prev) else []) + [r["_YRR"]]
        fig.add_scatter(x=xs, y=ys, mode="lines+markers", name=f"{grp} – run-rate", legendgroup=grp,
                        line=dict(color=c, dash="dot", width=2),
                        marker=dict(symbol="star", size=[0] * (len(xs) - 1) + [15], color=c,
                                    line=dict(color=RR_COLOR, width=1.5)))
        fig.add_scatter(x=[r[x]], y=[r[y]], mode="markers", name=f"{grp} – MTD actual", legendgroup=grp,
                        marker=dict(symbol="circle-open", size=11, color=c, line=dict(width=2)))
    fig.update_layout(legend_title_text="", yaxis_title=None)
    fig.update_xaxes(type="category")
    if tickformat:
        fig.update_yaxes(tickformat=tickformat)
    return fig


def add_period_cols(g):
    g = g.copy()
    g["YEAR"] = g.YEAR_MONTH.str[:4]
    g["MONTH"] = g.YEAR_MONTH.str[5:7].astype(int).map(lambda i: MON[i])
    g["LABEL"] = g.YEAR_MONTH.where(g["_F"] == 1, g.YEAR_MONTH + " (RR)")
    return g


def yoy_bar(g, title):
    g = g.copy().sort_values("YEAR_MONTH")
    g["_VRR"] = g.VOLUME * g["_F"]
    ly = g.set_index("YEAR_MONTH")["_VRR"]
    g["YOY%"] = [safe_div(v, ly.get(prev_month(m, 12))) - 1 for m, v in zip(g.YEAR_MONTH, g._VRR)]
    cy = g[g.YEAR == str(CY)].copy()
    cy["BASIS"] = cy["_F"].ne(1).map({True: "Run-rate", False: "Actual"})
    fig = px.bar(cy, x="MONTH", y="YOY%", color="BASIS", title=title, text=cy["YOY%"].map(fmt_pct),
                 category_orders={"MONTH": list(MON)[1:]},
                 color_discrete_map={"Actual": "#1F4E79", "Run-rate": RR_COLOR})
    fig.update_traces(textposition="outside")
    fig.update_layout(legend_title_text="").update_yaxes(tickformat=".0%")
    return fig


# ================================================================ SIDEBAR – DATA
st.sidebar.title("Settings")
uploads = st.sidebar.file_uploader(
    "Upload data files – Card MIS CSV and/or QRPH sheet (auto-detected)",
    type=["csv", "xlsx"], accept_multiple_files=True, key="uploads")
files = [(u.name, u) for u in uploads] if uploads else []
up_kinds = set()
for u in uploads or []:
    try:
        up_kinds.add(classify(_read_any(u, u.name).columns))
    except Exception:
        pass
    u.seek(0)
repo_found = scan_repo()

if "card" not in up_kinds and repo_found["card"]:
    cands = repo_found["card"]
    if len(cands) > 1:
        card_file = st.sidebar.selectbox(
            "Card MIS file to use", cands, index=0, key="card_file",
            help=f"Several card files found in the repo. Files with a {DATE_COL} column are listed first.")
    else:
        card_file = cands[0]
    files.append((card_file, card_file))
if "qrph" not in up_kinds and repo_found["qrph"]:
    files.append((repo_found["qrph"][0], repo_found["qrph"][0]))

with st.sidebar.expander("Repo files detected", expanded=not repo_found["card"] or not repo_found["qrph"]):
    for p, k, hdr, has_dt in repo_found.get("_all", []):
        if k:
            st.caption(f"✅ `{p}` → {k}" + (f" · {DATE_COL} ✔" if k == "card" and has_dt else
                                            f" · no {DATE_COL}" if k == "card" else ""))
        else:
            st.caption(f"⚠️ `{p}` → not recognised. Header read: `{', '.join(hdr) or '(empty)'}`")
    if not repo_found.get("_all"):
        st.caption("No .csv / .xlsx found in " + ", ".join(SCAN_DIRS))
    if st.button("Re-scan repo files"):
        scan_repo.clear()
        load_sources.clear()
        st.rerun()
if not files:
    st.info("Upload the Card MIS CSV and/or the QRPH sheet in the sidebar.")
    st.stop()
try:
    card_raw, qr, load_log = load_sources(tuple(files))
except Exception as e:
    st.error(f"Could not load data: {e}")
    st.stop()
for line in load_log:
    st.sidebar.caption(line)
HAS_CARD, HAS_QR = card_raw is not None, qr is not None and not qr.empty
if not HAS_CARD:
    st.sidebar.warning("No Card MIS loaded – card tabs will be empty.")
if not HAS_QR:
    st.sidebar.warning("No QRPH sheet loaded – QRPH figures will be empty.")


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


LATEST_M, RR_FACTOR, card_asof = None, 1.0, None
if HAS_CARD:
    LATEST_M = card_raw["YEAR_MONTH"].max()
    lp_ = pd.Period(LATEST_M, "M")
    dim_ = calendar.monthrange(lp_.year, lp_.month)[1]
    auto_dt, why = data_max_tran_date(card_raw, LATEST_M)
    if auto_dt is not None:
        override = st.sidebar.checkbox(f"Override card as-of date (file says {auto_dt:%b %d, %Y})",
                                       value=False, key="asof_override")
        if override:
            RR_FACTOR, card_asof = month_picker_date(f"Card data as of ({LATEST_M})", LATEST_M, auto_dt.day,
                                                     "asof_card")
        else:
            card_asof = auto_dt
            RR_FACTOR = dim_ / auto_dt.day
            st.sidebar.caption(f"Card data as of **{auto_dt:%b %d, %Y}** – latest {DATE_COL} in {LATEST_M}")
    else:
        st.sidebar.warning(f"Couldn't read the as-of date automatically: {why}. Set it manually.")
        RR_FACTOR, card_asof = month_picker_date(f"Card data as of ({LATEST_M})", LATEST_M, 14, "asof_card")
    st.sidebar.caption(f"Card run-rate ({LATEST_M} only): {dim_} / {card_asof.day} = **{RR_FACTOR:.3f}x**"
                       + (" (full month – no run-rate)" if RR_FACTOR == 1 else ""))
    if DATE_COL in card_raw.columns:
        with st.sidebar.expander(f"{DATE_COL} by month"):
            mx = (card_raw.groupby("YEAR_MONTH")[DATE_COL].max().dropna().sort_index(ascending=False)
                  .dt.strftime("%Y-%m-%d").rename("Max tran date").to_frame())
            st.dataframe(mx, width="stretch", height=240)

QR_LATEST, QR_RR = None, 1.0
if HAS_QR:
    QR_LATEST = qr["YEAR_MONTH"].max()
    _t = date.today()
    if HAS_CARD and QR_LATEST == LATEST_M:
        dflt = card_asof.day
    elif _t.strftime("%Y-%m") == QR_LATEST:
        dflt = max(_t.day - 1, 1)
    else:
        dflt = 31
    QR_RR, _ = month_picker_date(f"QRPH data as of ({QR_LATEST})", QR_LATEST, dflt, f"asof_qr_{dflt}")
    st.sidebar.caption(f"QRPH run-rate ({QR_LATEST} only): **{QR_RR:.3f}x**")

st.sidebar.divider()
months = sorted((set(card_raw["YEAR_MONTH"].dropna()) if HAS_CARD else set())
                | (set(qr["YEAR_MONTH"]) if HAS_QR else set()))
cur_m = st.sidebar.selectbox("Reporting month", months, index=len(months) - 1)
include_credits = st.sidebar.checkbox("Net off credit (refund) transactions (Card)", value=True)
df = card_raw
if HAS_CARD and not include_credits:
    df = df[df["TRXN_CODE"] == "Sale Transaction"]


def month_factor(ym):
    return RR_FACTOR if HAS_CARD and ym == LATEST_M else 1.0


def qr_factor(ym):
    return QR_RR if ym == QR_LATEST else 1.0


CUR_FACTOR = month_factor(cur_m)
mom_m, yoy_m = prev_month(cur_m, 1), prev_month(cur_m, 12)
cp = pd.Period(cur_m, "M")
CY, LY, END_MM = cp.year, cp.year - 1, cp.month
ytd_lbl = f"Jan–{MON[END_MM]}"
RR_TAG = " (RR)" if CUR_FACTOR != 1 else ""
VOL_LBL = f"Volume{RR_TAG}"
TXN_LBL = f"Txn Count{RR_TAG}"
IC_LBL = f"Interchange ₱{RR_TAG}"


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


def trend_actual(d, by=None):
    keys = ["YEAR_MONTH"] + ([by] if by else [])
    g = summarize(d[d.YEAR_MONTH <= cur_m], keys)
    g["_F"] = g["YEAR_MONTH"].map(month_factor)
    return add_period_cols(g)


def scale_rr(d):
    d = d.copy()
    f = d["YEAR_MONTH"].map(month_factor)
    for c in ["TOTAL_VOLUME", "TOTAL_COUNT", "INTERCHANGE_PHP", "MDR_X_VOL", "IC_X_VOL"]:
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


def full_compare(d, by):
    by = by if isinstance(by, list) else [by]
    p = period_table(d, by)
    y = ytd_compare(d, by)
    keep = by + [c for c in y.columns if c not in by]
    y = y[keep].rename(columns={c: f"YTD__{c}" for c in keep if c not in by})
    t = p.merge(y, on=by, how="outer")
    cur = d[d.YEAR_MONTH == cur_m]
    cnt = cur.groupby(by).agg(N_SUB=(SUB_COL, "nunique"), N_MID=(MID_COL, "nunique")).reset_index()
    t = t.merge(cnt, on=by, how="left")
    t.loc[t["YTD__VOLUME_LY"].fillna(0) < MIN_LY_BASE, "YTD__VOLUME_YoY%"] = float("nan")
    t.loc[t["VOLUME_LY"].fillna(0) < MIN_LY_BASE / 10, "VOLUME_YoY%"] = float("nan")
    return t


# ================================================================ MONTHLY-SERIES LOGIC (QRPH / Total)
def rr_series(m, fac):
    m = m[["YEAR_MONTH", "VOLUME", "TXN_COUNT"]].copy()
    f = m["YEAR_MONTH"].map(fac)
    m["VOLUME"] *= f
    m["TXN_COUNT"] *= f
    return m


def series_metrics(s):
    s = s.set_index("YEAR_MONTH")[["VOLUME", "TXN_COUNT"]]
    ym = s.index.to_series()
    nan = pd.Series({"VOLUME": float("nan"), "TXN_COUNT": float("nan")})

    def at(m):
        return s.loc[m] if m in s.index else nan

    def span(year, upto):
        sel = s[(ym.str[:4] == str(year)) & (ym.str[5:7].astype(int) <= upto)]
        return sel.sum() if len(sel) == upto else nan

    P = {"CUR": at(cur_m), "PM": at(mom_m), "LYM": at(yoy_m),
         "YTD": span(CY, END_MM), "YTDLY": span(LY, END_MM), "FYLY": span(LY, 12)}
    P["FY"] = P["YTD"] / END_MM * 12
    out = {k: {p: P[p][k] for p in P} for k in ["VOLUME", "TXN_COUNT"]}
    out["TICKET"] = {p: safe_div(P[p]["VOLUME"], P[p]["TXN_COUNT"]) for p in P}
    for v in out.values():
        v["MoM%"] = safe_div(v["CUR"], v["PM"]) - 1
        v["YoY%"] = safe_div(v["CUR"], v["LYM"]) - 1
        v["YTD_YoY%"] = safe_div(v["YTD"], v["YTDLY"]) - 1
        v["FY_YoY%"] = safe_div(v["FY"], v["FYLY"]) - 1
    return out


def series_table(res_by_src):
    names = {"VOLUME": "Volume (₱)", "TXN_COUNT": "Txn Count", "TICKET": "Avg Ticket (₱)"}
    rows = []
    for src, res in res_by_src.items():
        for k, lbl in names.items():
            v = res[k]
            rows.append({"Source": src, "Metric": lbl, f"{cur_m}": v["CUR"],
                         f"{mom_m} (PM)": v["PM"], "MoM": v["MoM%"], f"{yoy_m} (LY)": v["LYM"], "YoY": v["YoY%"],
                         f"YTD {CY}": v["YTD"], f"YTD {LY}": v["YTDLY"], "YTD YoY": v["YTD_YoY%"],
                         f"FY {CY} (RR)": v["FY"], f"FY {LY}": v["FYLY"], "FY YoY": v["FY_YoY%"]})
    t = pd.DataFrame(rows)
    pct = ["MoM", "YoY", "YTD YoY", "FY YoY"]
    fmt = {c: ("{:+.1%}" if c in pct else "{:,.0f}") for c in t.columns if c not in ("Source", "Metric")}
    st.dataframe(t.style.format(fmt, na_rep="-"), width="stretch", hide_index=True)


def series_kpis(res, title, rr=False):
    v, n = res["VOLUME"], res["TXN_COUNT"]
    tag = " (RR)" if rr else ""
    st.markdown(f"**{title}**")
    c = st.columns(5)
    c[0].metric(f"Volume {cur_m}{tag}", fmt_php(v["CUR"]), f"{fmt_pct(v['MoM%'])} MoM | {fmt_pct(v['YoY%'])} YoY")
    c[1].metric(f"Txn Count {cur_m}{tag}", fmt_int(n["CUR"]), f"{fmt_pct(n['MoM%'])} MoM | {fmt_pct(n['YoY%'])} YoY")
    c[2].metric(f"YTD Volume {CY}", fmt_php(v["YTD"]), f"{fmt_pct(v['YTD_YoY%'])} YoY (LY {fmt_php(v['YTDLY'])})")
    c[3].metric(f"YTD Txn Count {CY}", fmt_int(n["YTD"]), f"{fmt_pct(n['YTD_YoY%'])} YoY (LY {fmt_int(n['YTDLY'])})")
    c[4].metric(f"FY {CY} Volume (RR)", fmt_php(v["FY"]), f"{fmt_pct(v['FY_YoY%'])} vs FY {LY} {fmt_php(v['FYLY'])}")


def year_filter(g, key, default_last=None, label="Years shown in monthly charts"):
    g = g[g.YEAR_MONTH <= cur_m]
    yrs = sorted(g.YEAR.unique())
    sel = st.multiselect(label, yrs, default=yrs if default_last is None else yrs[-default_last:], key=key)
    return g[g.YEAR.isin(sel)]


ORDER_M = {"MONTH": list(MON)[1:]}


# ================================================================ CARD DISPLAY HELPERS
def kpi_row(t):
    r = t.iloc[0]
    c = st.columns(6)
    c[0].metric(VOL_LBL, fmt_php(r.VOLUME), f"{fmt_pct(r['VOLUME_MoM%'])} MoM | {fmt_pct(r['VOLUME_YoY%'])} YoY")
    c[1].metric(TXN_LBL, fmt_int(r.TXN_COUNT), f"{fmt_pct(r['TXN_COUNT_MoM%'])} MoM | {fmt_pct(r['TXN_COUNT_YoY%'])} YoY")
    c[2].metric("Weighted MDR", fmt_rate(r.WEIGHTED_MDR),
                f"{fmt_bps(r.WEIGHTED_MDR_MoM_bps)} bps MoM | {fmt_bps(r.WEIGHTED_MDR_YoY_bps)} YoY")
    c[3].metric("Blended MDR (simple avg)", fmt_rate(r.BLENDED_MDR),
                f"{fmt_bps(r.BLENDED_MDR_MoM_bps)} bps MoM | {fmt_bps(r.BLENDED_MDR_YoY_bps)} YoY")
    c[4].metric("Weighted Interchange", fmt_rate(r.WEIGHTED_IC),
                f"{fmt_bps(r.WEIGHTED_IC_MoM_bps)} bps MoM | {fmt_bps(r.WEIGHTED_IC_YoY_bps)} YoY", delta_color="inverse")
    c[5].metric(IC_LBL, fmt_php(r.INTERCHANGE_PHP),
                f"{fmt_pct(r['INTERCHANGE_PHP_MoM%'])} MoM | {fmt_pct(r['INTERCHANGE_PHP_YoY%'])} YoY", delta_color="inverse")


def ytd_kpi_block(d, label):
    t = ytd_compare(d.assign(ALL=label), "ALL").iloc[0]
    small_ly = pd.isna(t.VOLUME_LY) or t.VOLUME_LY < MIN_LY_BASE
    vol_yoy = "n.m. (LY base too small)" if small_ly else f"{fmt_pct(t['VOLUME_YoY%'])} YoY"
    c = st.columns(6)
    c[0].metric(f"YTD Volume {CY}", fmt_php(t.VOLUME), f"{vol_yoy} (LY {fmt_php(t.VOLUME_LY)})",
                delta_color="off" if small_ly else "normal")
    c[1].metric(f"YTD Txn Count {CY}", fmt_int(t.TXN_COUNT),
                f"{fmt_pct(t['TXN_COUNT_YoY%'])} YoY (LY {fmt_int(t.TXN_COUNT_LY)})")
    c[2].metric("YTD Weighted MDR", fmt_rate(t.WEIGHTED_MDR),
                f"{fmt_bps(t.WEIGHTED_MDR_YoY_bps)} bps YoY (LY {fmt_rate(t.WEIGHTED_MDR_LY)})")
    c[3].metric("YTD Weighted IC", fmt_rate(t.WEIGHTED_IC),
                f"{fmt_bps(t.WEIGHTED_IC_YoY_bps)} bps YoY (LY {fmt_rate(t.WEIGHTED_IC_LY)})", delta_color="inverse")
    c[4].metric(f"YTD Interchange ₱ {CY}", fmt_php(t.INTERCHANGE_PHP),
                f"{fmt_pct(t['INTERCHANGE_PHP_YoY%'])} YoY", delta_color="inverse")
    c[5].metric(f"FY {CY} Volume (RR)", fmt_php(t.VOLUME_FY),
                f"{fmt_pct(t['VOLUME_FY_YoY%'])} vs FY {LY} {fmt_php(t.VOLUME_FYLY)}")


DISPLAY_COLS = {
    "VOLUME": VOL_LBL, "VOLUME_PM": f"Volume {mom_m} (Actual)", "VOLUME_MoM%": "Vol MoM", "VOLUME_YoY%": "Vol YoY",
    "TXN_COUNT": TXN_LBL, "TXN_COUNT_MoM%": "Txn MoM", "TXN_COUNT_YoY%": "Txn YoY",
    "WEIGHTED_MDR": "Wtd MDR", "WEIGHTED_MDR_MoM_bps": "Wtd MDR MoM (bps)", "WEIGHTED_MDR_YoY_bps": "Wtd MDR YoY (bps)",
    "BLENDED_MDR": "Blended MDR",
    "WEIGHTED_IC": "Wtd IC", "WEIGHTED_IC_MoM_bps": "Wtd IC MoM (bps)", "WEIGHTED_IC_YoY_bps": "Wtd IC YoY (bps)",
    "BLENDED_IC": "Blended IC",
    "INTERCHANGE_PHP": IC_LBL, "INTERCHANGE_PHP_MoM%": "IC ₱ MoM", "INTERCHANGE_PHP_YoY%": "IC ₱ YoY",
}
RATE_KEYS = ("WEIGHTED_MDR", "BLENDED_MDR", "WEIGHTED_IC", "BLENDED_IC")


def show_table(t, keys):
    keys = keys if isinstance(keys, list) else [keys]
    out = t[keys + list(DISPLAY_COLS)].rename(columns=DISPLAY_COLS)
    fmt = {v: ("{:+.1%}" if k.endswith("%") else "{:+.1f}" if k.endswith("bps")
               else "{:.2%}" if k in RATE_KEYS else "{:,.0f}") for k, v in DISPLAY_COLS.items()}
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
    fmt = {v: ("{:+.1%}" if k.endswith("%") else "{:+.1f}" if k.endswith("bps")
               else "{:.2%}" if k in RATE_KEYS else "{:,.0f}") for k, v in YTD_COLS.items()}
    st.dataframe(out.style.format(fmt, na_rep="-"), width="stretch", hide_index=True)


FULL_COLS = {
    "YTD__VOLUME": f"YTD Vol {CY}", "YTD__VOLUME_LY": f"YTD Vol {LY}", "YTD__VOLUME_YoY%": "YTD Vol YoY",
    "YTD__TXN_COUNT": f"YTD Txn {CY}", "YTD__TXN_COUNT_YoY%": "YTD Txn YoY",
    "YTD__WEIGHTED_MDR": "YTD Wtd MDR", "YTD__WEIGHTED_MDR_YoY_bps": "YTD MDR YoY (bps)",
    "YTD__WEIGHTED_IC": "YTD Wtd IC", "YTD__WEIGHTED_IC_YoY_bps": "YTD IC YoY (bps)",
    "YTD__VOLUME_FY": f"FY {CY} Vol (RR)", "YTD__VOLUME_FY_YoY%": "FY Vol YoY",
    "VOLUME": f"{cur_m} Vol{RR_TAG}", "VOLUME_PM": f"{mom_m} Vol", "VOLUME_MoM%": "Vol MoM",
    "VOLUME_LY": f"{yoy_m} Vol", "VOLUME_YoY%": "Vol YoY (month)",
    "TXN_COUNT": f"{cur_m} Txn{RR_TAG}", "TXN_COUNT_MoM%": "Txn MoM", "TXN_COUNT_YoY%": "Txn YoY (month)",
    "WEIGHTED_MDR": f"{cur_m} Wtd MDR", "WEIGHTED_MDR_MoM_bps": "MDR MoM (bps)",
    "WEIGHTED_IC": f"{cur_m} Wtd IC", "WEIGHTED_IC_MoM_bps": "IC MoM (bps)",
    "INTERCHANGE_PHP": f"{cur_m} IC ₱{RR_TAG}",
    "N_MID": f"# MIDs ({cur_m})", "N_SUB": f"# Sub-merchants ({cur_m})",
}


def show_full(t, keys, extra=None, height=None):
    keys = keys if isinstance(keys, list) else [keys]
    cols = [c for c in FULL_COLS if c in t.columns]
    out = t[keys + (extra or []) + cols].rename(columns=FULL_COLS)
    fmt = {}
    for k, v in FULL_COLS.items():
        if k.endswith("%"):
            fmt[v] = "{:+.1%}"
        elif k.endswith("bps"):
            fmt[v] = "{:+.1f}"
        elif k.replace("YTD__", "") in RATE_KEYS:
            fmt[v] = "{:.2%}"
        else:
            fmt[v] = "{:,.0f}"
    for e in extra or []:
        fmt[e] = "{:.1%}"
    kw = dict(width="stretch", hide_index=True)
    if height:
        kw["height"] = height
    st.dataframe(out.style.format(fmt, na_rep="-"), **kw)


def need_card():
    if not HAS_CARD:
        st.info("Upload the Card MIS CSV in the sidebar to see this tab.")
        return False
    return True


def rr_snapshot(g, cat, val_rr):
    out = g[[cat] + ([c for c in ["PF_DM_NAME"] if c in g.columns and c != cat])].copy()
    out[val_rr] = g[val_rr] / CUR_FACTOR
    out["_F"] = CUR_FACTOR
    return out


def ytd_vs_ly_bar(t, key, title, top=15):
    t = t.sort_values("YTD__VOLUME", ascending=False).head(top)
    long = pd.concat([
        pd.DataFrame({key: t[key], "Period": f"YTD {CY}" + (" (incl. RR)" if CUR_FACTOR != 1 else ""),
                      "Volume": t["YTD__VOLUME"]}),
        pd.DataFrame({key: t[key], "Period": f"YTD {LY}", "Volume": t["YTD__VOLUME_LY"]})])
    fig = px.bar(long, x="Volume", y=key, color="Period", barmode="group", orientation="h", title=title,
                 color_discrete_sequence=["#1F4E79", "#A9B8C9"])
    return fig.update_yaxes(autorange="reversed", type="category").update_layout(
        height=max(380, 32 * len(t)), legend_title_text="")


# ================================================================ PROFITABILITY HELPERS
def profit_table(d, keys):
    """Per-entity volume, NET_INCOME, NI margin (bps of volume), MDR–IC spread (bps), ticket size.
    Uses ACTUAL figures (no run-rate) – margins are ratios so they are period-consistent."""
    keys = keys if isinstance(keys, list) else [keys]
    g = d.groupby(keys, dropna=False).agg(
        VOLUME=("TOTAL_VOLUME", "sum"), TXN_COUNT=("TOTAL_COUNT", "sum"), NI=(NI_COL, "sum"),
        MDR_X_VOL=("MDR_X_VOL", "sum"), IC_X_VOL=("IC_X_VOL", "sum"),
        MONTHS=("YEAR_MONTH", "nunique")).reset_index()
    g = g[g.VOLUME > 0].copy()
    g["WTD_MDR"] = g.MDR_X_VOL / g.VOLUME
    g["WTD_IC"] = g.IC_X_VOL / g.VOLUME
    g["SPREAD_BPS"] = (g.WTD_MDR - g.WTD_IC) * 1e4
    g["NI_BPS"] = g.NI / g.VOLUME * 1e4
    g["TICKET"] = g.VOLUME / g.TXN_COUNT.where(g.TXN_COUNT > 0)
    g["NAME"] = g[keys].astype(str).agg(" › ".join, axis=1)
    return g.drop(columns=["MDR_X_VOL", "IC_X_VOL"])


def assign_quadrants(g, pareto_pct, hurdle_bps):
    g = g.sort_values("VOLUME", ascending=False).copy()
    share = g.VOLUME / g.VOLUME.sum()
    g["VOL_SHARE"] = share
    g["HIGH_VOL"] = (share.cumsum() - share) < pareto_pct / 100      # merchants making up the top X% of volume
    vol_cut = g.loc[g.HIGH_VOL, "VOLUME"].min() if g.HIGH_VOL.any() else g.VOLUME.max()
    prof = g.NI_BPS >= hurdle_bps
    quads = list(QUAD_COLORS)
    g["QUADRANT"] = np.select([g.HIGH_VOL & prof, g.HIGH_VOL & ~prof, ~g.HIGH_VOL & prof],
                              quads[:3], default=quads[3])
    return g, vol_cut


# ================================================================ HEADER & TABS
st.title("Merchant Acquiring Dashboard")
notes = []
if HAS_CARD:
    notes.append(f"Card: {LATEST_M} as of {card_asof:%b %d}, run-rated ×{RR_FACTOR:.2f}")
if HAS_QR:
    notes.append(f"QRPH: {QR_LATEST} run-rated ×{QR_RR:.2f}")
st.caption(f"Reporting month: **{cur_m}** · MoM vs **{mom_m}** · YoY vs **{yoy_m}** · YTD {ytd_lbl} · "
           + " · ".join(notes) + " · all earlier months are actuals")
st.caption(RR_LEGEND)

tab_tot, tab_yoy, tab_ovw, tab_cat, tab_mer, tab_pf, tab_prof, tab_qr = st.tabs(
    ["1 · Total Acquiring (Card + QRPH)", "2 · Card YoY & YTD", "3 · PF vs DM Overview",
     "4 · Category Mix", "5 · Merchant Level", "6 · PF Level (MID → Sub-merchant)", "7 · Profitability Map",
     "8 · QRPH"])

card_a = None
if HAS_CARD:
    card_a = summarize(df, "YEAR_MONTH")[["YEAR_MONTH", "VOLUME", "TXN_COUNT"]]
    card_a["_F"] = card_a.YEAR_MONTH.map(month_factor)
qr_a = None
if HAS_QR:
    qr_a = qr[["YEAR_MONTH", "VOLUME", "TXN_COUNT"]].copy()
    qr_a["_F"] = qr_a.YEAR_MONTH.map(qr_factor)

# ================================================================ TAB 1 – TOTAL ACQUIRING
with tab_tot:
    st.caption("Total = Card MIS (PF + DM) + QRPH posted (separate data sources, simple sum). "
               "MDR / interchange are card-only and shown in the card tabs.")
    parts_a = {k: v for k, v in {"Card (PF + DM)": card_a, "QRPH": qr_a}.items() if v is not None}
    parts_rr = {k: rr_series(v, (month_factor if k.startswith("Card") else qr_factor)) for k, v in parts_a.items()}

    both = None
    for name, s in parts_a.items():
        s2 = s.rename(columns={"VOLUME": f"V_{name}", "TXN_COUNT": f"N_{name}", "_F": f"F_{name}"})
        both = s2 if both is None else both.merge(s2, on="YEAR_MONTH", how="outer")
    both = both.sort_values("YEAR_MONTH")
    gaps = {n: both.loc[both[f"V_{n}"].isna() & (both.YEAR_MONTH <= cur_m), "YEAR_MONTH"].tolist() for n in parts_a}
    if len(parts_a) < 2:
        st.warning(f"Only {list(parts_a)[0]} is loaded – total equals that source.")
    elif any(gaps.values()):
        st.warning("Month coverage differs – " + "; ".join(f"no {n} for {g}" for n, g in gaps.items() if g)
                   + ". YTD totals needing those months show '-'.")

    tot = pd.DataFrame({"YEAR_MONTH": both.YEAR_MONTH})
    tot["VOLUME"] = sum(both[f"V_{n}"] for n in parts_a)
    tot["TXN_COUNT"] = sum(both[f"N_{n}"] for n in parts_a)
    tot["VOLUME_RR"] = sum(both[f"V_{n}"] * both[f"F_{n}"] for n in parts_a)
    tot["TXN_COUNT_RR"] = sum(both[f"N_{n}"] * both[f"F_{n}"] for n in parts_a)
    tot = tot.dropna(subset=["VOLUME"])
    tot["_F"] = (tot.VOLUME_RR / tot.VOLUME).where(lambda x: (x - 1).abs() > 1e-9, 1.0)
    tot_rr = tot[["YEAR_MONTH"]].assign(VOLUME=tot.VOLUME_RR, TXN_COUNT=tot.TXN_COUNT_RR)

    any_rr = cur_m in {m for m, f in zip(tot.YEAR_MONTH, tot._F) if f != 1}
    res = {"Total (Card + QRPH)": series_metrics(tot_rr)}
    res.update({n: series_metrics(s) for n, s in parts_rr.items()})
    for n, r in res.items():
        is_rr = any_rr if n.startswith("Total") else (
            (cur_m == LATEST_M and RR_FACTOR != 1) if n.startswith("Card") else (cur_m == QR_LATEST and QR_RR != 1))
        series_kpis(r, n, rr=is_rr)
    st.subheader("Comparison table – current vs prev month vs LY, YTD, FY outlook")
    series_table(res)

    tot = add_period_cols(tot)
    tf = year_filter(tot, "yr_tot_line")
    c1, c2 = st.columns(2)
    c1.plotly_chart(line_rr(tf, "MONTH", "VOLUME", "Total Acquiring – Monthly Volume by Year", color="YEAR",
                            order=ORDER_M, y_rr="VOLUME_RR"), width="stretch")
    c2.plotly_chart(line_rr(tf, "MONTH", "VOLUME", "Total Acquiring – Cumulative YTD Volume", color="YEAR",
                            order=ORDER_M, y_rr="VOLUME_RR", cum=True), width="stretch")

    long = pd.concat([add_period_cols(s.assign(SOURCE=n)) for n, s in parts_a.items()])
    long = long[long.YEAR_MONTH <= cur_m]
    rr_months = set(long.loc[long._F != 1, "YEAR_MONTH"])
    long["LABEL"] = long.YEAR_MONTH.where(~long.YEAR_MONTH.isin(rr_months), long.YEAR_MONTH + " (RR)")
    yrs = sorted(long.YEAR.unique())
    sel = st.multiselect("Years shown in source-split charts", yrs, default=yrs[-2:], key="yr_tot")
    long = long[long.YEAR.isin(sel)].sort_values("YEAR_MONTH")
    c1, c2 = st.columns(2)
    c1.plotly_chart(bar_rr(long, "LABEL", "VOLUME", "Monthly Volume – Card vs QRPH", color="SOURCE",
                           cmap=SRC_COLORS), width="stretch")
    c2.plotly_chart(bar_rr(long, "LABEL", "TXN_COUNT", "Monthly Txn Count – Card vs QRPH", color="SOURCE",
                           cmap=SRC_COLORS, fmt=fmt_num), width="stretch")
    if len(parts_a) == 2 and not long.empty:
        long["V_RR"] = long.VOLUME * long._F
        long["N_RR"] = long.TXN_COUNT * long._F
        p = long.pivot_table(index=["YEAR_MONTH"], columns="SOURCE", values=["VOLUME", "TXN_COUNT", "V_RR", "N_RR"],
                             aggfunc="sum")
        sh = []
        for meas, a, r in [("Volume share", "VOLUME", "V_RR"), ("Txn share", "TXN_COUNT", "N_RR")]:
            sh.append(pd.DataFrame({
                "YEAR_MONTH": p.index, "Measure": meas,
                "SHARE": (p[(a, "QRPH")] / p[a].sum(axis=1)).values,
                "SHARE_RR": (p[(r, "QRPH")] / p[r].sum(axis=1)).values}))
        sh = pd.concat(sh)
        sh["_F"] = sh.YEAR_MONTH.isin(rr_months).map({True: 2.0, False: 1.0})
        sh["LABEL"] = sh.YEAR_MONTH.where(sh._F == 1, sh.YEAR_MONTH + " (RR)")
        c3, c4 = st.columns(2)
        c3.plotly_chart(line_rr(sh, "LABEL", "SHARE", "QRPH Share of Total Acquiring", color="Measure",
                                y_rr="SHARE_RR", tickformat=".0%"), width="stretch")
        mix = pd.DataFrame({"SRC": list(parts_rr), "VOLUME": [res[n]["VOLUME"]["YTD"] for n in parts_rr]})
        c4.plotly_chart(pie(mix, "VOLUME", f"YTD {CY} Volume Mix (incl. run-rated {MON[END_MM]})", names="SRC")
                        .update_traces(marker=dict(colors=[SRC_COLORS[n] for n in parts_rr])), width="stretch")

# ================================================================ TAB 2 – CARD YoY & YTD
with tab_yoy:
    if need_card():
        seg4 = st.radio("Segment", ["All", "Direct Merchant", "Payment Facilitator"], horizontal=True, key="yoy_seg")
        d4 = df if seg4 == "All" else df[df.PF_DM_TAG == seg4]
        st.caption(f"YTD {CY} = {ytd_lbl} {CY} (latest month run-rated) vs {ytd_lbl} {LY} actual. "
                   f"FY outlook = YTD ÷ {END_MM} × 12.")
        st.subheader(f"YTD {ytd_lbl}: {CY} vs {LY}")
        ytd_kpi_block(d4, "Total")
        m_all = trend_actual(d4)
        m = year_filter(m_all, "yr_tab4")
        c1, c2 = st.columns(2)
        c1.plotly_chart(line_rr(m, "MONTH", "VOLUME", "Monthly Volume by Year", color="YEAR", order=ORDER_M), width="stretch")
        c2.plotly_chart(line_rr(m, "MONTH", "VOLUME", "Cumulative YTD Volume by Year", color="YEAR", order=ORDER_M, cum=True),
                        width="stretch")
        c3, c4 = st.columns(2)
        c3.plotly_chart(yoy_bar(m_all, f"Monthly Volume YoY % ({CY} vs {LY})"), width="stretch")
        c4.plotly_chart(rate_line(m, "MONTH", "WEIGHTED_MDR", "Weighted MDR by Year (latest month = MTD; not run-rated)",
                                  color="YEAR", order=ORDER_M), width="stretch")
        st.subheader("YTD YoY by Segment")
        show_ytd(ytd_compare(df, "PF_DM_TAG"), "PF_DM_TAG")
        st.subheader(f"YTD YoY by Category – {seg4}")
        show_ytd(ytd_compare(d4, "INDUSTRY").sort_values("VOLUME", ascending=False), "INDUSTRY")
        st.subheader(f"YTD YoY by Merchant – {seg4}")
        mer4 = ytd_compare(d4, "PF_DM_NAME")
        mer4.loc[mer4.VOLUME_LY.fillna(0) < MIN_LY_BASE, "VOLUME_YoY%"] = float("nan")
        mer4["VOL_CHG"] = mer4.VOLUME.fillna(0) - mer4.VOLUME_LY.fillna(0)
        movers = pd.concat([mer4.nlargest(10, "VOL_CHG"), mer4.nsmallest(10, "VOL_CHG")]) \
            .drop_duplicates("PF_DM_NAME").sort_values("VOL_CHG")
        st.plotly_chart(px.bar(movers, x="VOL_CHG", y="PF_DM_NAME", orientation="h",
                               color=(movers.VOL_CHG > 0).map({True: "Gainer", False: "Decliner"}),
                               color_discrete_map={"Gainer": "#2E8B57", "Decliner": "#C0392B"},
                               title=f"Top 10 Gainers / Decliners – YTD Volume change vs {LY} (₱, YTD incl. run-rated {MON[END_MM]})")
                        .update_layout(showlegend=False, height=600), width="stretch")
        st.caption(f"Vol YoY shown as '-' (n.m.) where YTD {LY} volume < ₱{MIN_LY_BASE:,.0f}.")
        show_ytd(mer4.sort_values("VOLUME", ascending=False), "PF_DM_NAME")

# ================================================================ TAB 3 – PF vs DM
with tab_ovw:
    if need_card():
        st.subheader("Total Portfolio (Card)")
        kpi_row(period_table(df.assign(ALL="Total"), "ALL"))
        st.subheader("By Segment")
        seg = period_table(df, "PF_DM_TAG")
        for tag in seg.PF_DM_TAG.dropna():
            st.markdown(f"**{tag}**")
            kpi_row(seg[seg.PF_DM_TAG == tag].reset_index(drop=True))
        tr = year_filter(trend_actual(df, "PF_DM_TAG"), "yr_tab1", default_last=2).sort_values("YEAR_MONTH")
        if not tr.empty:
            c1, c2 = st.columns(2)
            c1.plotly_chart(bar_rr(tr, "LABEL", "VOLUME", "Monthly Volume", color="PF_DM_TAG", cmap=TAG_COLORS), width="stretch")
            c2.plotly_chart(bar_rr(tr, "LABEL", "TXN_COUNT", "Monthly Transaction Count", color="PF_DM_TAG",
                                   cmap=TAG_COLORS, fmt=fmt_num), width="stretch")
            c3, c4 = st.columns(2)
            c3.plotly_chart(rate_line(tr, "LABEL", "WEIGHTED_MDR", "Weighted MDR (not run-rated)",
                                      color="PF_DM_TAG", cmap=TAG_COLORS), width="stretch")
            c4.plotly_chart(rate_line(tr, "LABEL", "WEIGHTED_IC", "Weighted Interchange (not run-rated)",
                                      color="PF_DM_TAG", cmap=TAG_COLORS), width="stretch")
        show_table(seg, "PF_DM_TAG")

# ================================================================ TAB 4 – CATEGORY MIX
with tab_cat:
    if need_card():
        seg_pick = st.radio("Segment", ["All", "Direct Merchant", "Payment Facilitator"], horizontal=True, key="cat_seg")
        d2 = df if seg_pick == "All" else df[df.PF_DM_TAG == seg_pick]
        cat = period_table(d2, "INDUSTRY").sort_values("VOLUME", ascending=False)
        cat["VOL_SHARE"] = cat.VOLUME / cat.VOLUME.sum()
        cat["VOL_SHARE_PM"] = cat.VOLUME_PM / cat.VOLUME_PM.sum()
        cat["SHARE_CHG_PP"] = (cat.VOL_SHARE - cat.VOL_SHARE_PM) * 100
        rr_t = " – run-rate basis (share = MTD share)" if CUR_FACTOR != 1 else ""
        c1, c2 = st.columns(2)
        c1.plotly_chart(pie(top_n_others(cat, "VOLUME"), "VOLUME", f"Volume Share – {seg_pick} – {cur_m}{rr_t}"), width="stretch")
        c2.plotly_chart(pie(top_n_others(cat, "TXN_COUNT"), "TXN_COUNT", f"Txn Count Share – {seg_pick} – {cur_m}{rr_t}"), width="stretch")
        top = cat.loc[cat.INDUSTRY != OTHERS_LABEL, "INDUSTRY"].head(CAT_TOP_N).tolist()
        tr2 = year_filter(trend_actual(d2, "INDUSTRY"), "yr_cat", default_last=2)
        tr2["CAT"] = tr2.INDUSTRY.where(tr2.INDUSTRY.isin(top), OTHERS_LABEL)
        tr2 = tr2.groupby(["YEAR_MONTH", "LABEL", "CAT", "_F"], as_index=False).VOLUME.sum().sort_values("YEAR_MONTH")
        c1, c2 = st.columns(2)
        c1.plotly_chart(bar_rr(tr2, "LABEL", "VOLUME", f"Monthly Volume by Category (Top {CAT_TOP_N} + Others)",
                               color="CAT"), width="stretch")
        tr2["SHARE"] = tr2.VOLUME / tr2.groupby("LABEL").VOLUME.transform("sum")
        c2.plotly_chart(px.bar(tr2, x="LABEL", y="SHARE", color="CAT", category_orders={"CAT": top + [OTHERS_LABEL]},
                               title="Volume Share Trend ((RR) month share = MTD share)")
                        .update_yaxes(tickformat=".0%").update_xaxes(type="category"), width="stretch")
        st.plotly_chart(px.bar(cat.head(15), x="SHARE_CHG_PP", y="INDUSTRY", orientation="h",
                               title=f"Volume Share Change vs {mom_m} (pp)").update_yaxes(autorange="reversed"), width="stretch")
        show_table(cat.assign(**{"Vol Share": cat.VOL_SHARE}), ["INDUSTRY", "Vol Share"])

# ================================================================ TAB 5 – MERCHANT LEVEL
with tab_mer:
    if need_card():
        seg3 = st.radio("Segment", ["Payment Facilitator", "Direct Merchant"], horizontal=True, key="m_seg")
        d3 = df[df.PF_DM_TAG == seg3]
        mer = period_table(d3, "PF_DM_NAME").sort_values("VOLUME", ascending=False)
        mer["VOL_SHARE"] = mer.VOLUME / mer.VOLUME.sum()
        top_n = st.slider("Top N merchants", 5, 30, 10, key="topn_mer")
        topm = mer.head(top_n)
        c1, c2 = st.columns(2)
        c1.plotly_chart(bar_rr(rr_snapshot(topm, "PF_DM_NAME", "VOLUME"), "PF_DM_NAME", "VOLUME",
                               f"Top {top_n} {seg3}s – {cur_m} volume", horizontal=True,
                               height=max(400, 35 * top_n)), width="stretch")
        mv = topm.melt(id_vars="PF_DM_NAME", value_vars=["VOLUME_MoM%", "VOLUME_YoY%"], var_name="Comparison", value_name="Growth")
        c2.plotly_chart(px.bar(mv, x="Growth", y="PF_DM_NAME", color="Comparison", barmode="group", orientation="h",
                               title=f"Volume Growth – MoM vs YoY{' (on run-rate)' if CUR_FACTOR != 1 else ''}")
                        .update_xaxes(tickformat=".0%").update_yaxes(autorange="reversed")
                        .update_layout(height=max(400, 35 * top_n)), width="stretch")
        show_table(mer.assign(**{"Vol Share": mer.VOL_SHARE}), ["PF_DM_NAME", "Vol Share"])

        if seg3 == "Payment Facilitator":
            st.divider()
            show_sub = st.toggle("Show sub-merchants under each Payment Facilitator (MERCHANT_NAME_ORIG)",
                                 value=False, key="sub_toggle")
            if show_sub:
                pf_list = mer.PF_DM_NAME.dropna().tolist()
                pfs = st.multiselect("Payment Facilitators to expand", pf_list, default=pf_list[:3], key="sub_pfs")
                ds = d3[d3.PF_DM_NAME.isin(pfs)]
                if ds.empty:
                    st.info("Select at least one Payment Facilitator.")
                else:
                    sub = period_table(ds, ["PF_DM_NAME", SUB_COL]).sort_values("VOLUME", ascending=False)
                    sub["Share of PF"] = sub.VOLUME / sub.groupby("PF_DM_NAME").VOLUME.transform("sum")
                    sub_n = st.slider("Top N sub-merchants per PF", 3, 30, 10, key="sub_n")
                    subt = sub.groupby("PF_DM_NAME", group_keys=False).head(sub_n)
                    tm = subt[subt.VOLUME > 0]
                    st.plotly_chart(px.treemap(tm, path=[px.Constant("Selected PFs"), "PF_DM_NAME", SUB_COL],
                                               values="VOLUME",
                                               title=f"PF → Sub-merchant volume ({cur_m}{' run-rate' if CUR_FACTOR != 1 else ''}) – click a PF to zoom in")
                                    .update_traces(texttemplate="%{label}<br>%{percentParent:.1%}")
                                    .update_layout(height=550, margin=dict(t=50, l=5, r=5, b=5)), width="stretch")
                    sb = subt.head(30)
                    st.plotly_chart(bar_rr(rr_snapshot(sb, SUB_COL, "VOLUME"), SUB_COL, "VOLUME",
                                           f"Top sub-merchants – {cur_m} volume (colour = PF)", color="PF_DM_NAME",
                                           horizontal=True, height=max(450, 28 * len(sb))), width="stretch")
                    show_table(subt.assign(**{"Vol Share": subt["Share of PF"]}), ["PF_DM_NAME", SUB_COL, "Vol Share"])
                    st.caption("'Vol Share' here = share of the sub-merchant within its PF.")

        st.divider()
        st.subheader("Merchant drill-down")
        cA, cB = st.columns(2)
        pick = cA.selectbox(f"Drill down: {seg3}", mer.PF_DM_NAME.dropna().tolist(), key="drill_pf")
        dd = d3[d3.PF_DM_NAME == pick]
        sub_pick = "All"
        if seg3 == "Payment Facilitator":
            subs = summarize(dd[dd.YEAR_MONTH == cur_m], SUB_COL).sort_values("VOLUME", ascending=False)[SUB_COL].tolist()
            subs += sorted(set(dd[SUB_COL]) - set(subs))
            sub_pick = cB.selectbox(f"Sub-merchant under {pick}", ["All"] + subs, key="drill_sub")
            if sub_pick != "All":
                dd = dd[dd[SUB_COL] == sub_pick]
        label = pick if sub_pick == "All" else f"{pick} › {sub_pick}"
        trd_all = trend_actual(dd)
        st.markdown(f"#### ① YoY & YTD – {label}")
        st.caption(f"YTD {ytd_lbl} {CY}{' (incl. run-rated ' + MON[END_MM] + ')' if CUR_FACTOR != 1 else ''} "
                   f"vs YTD {ytd_lbl} {LY} actual · FY outlook = YTD ÷ {END_MM} × 12 vs FY {LY} actual")
        ytd_kpi_block(dd, label)
        trd_y = year_filter(trd_all, "yr_drill")
        c1, c2 = st.columns(2)
        c1.plotly_chart(line_rr(trd_y, "MONTH", "VOLUME", f"{label} – Monthly Volume by Year", color="YEAR",
                                order=ORDER_M), width="stretch")
        c2.plotly_chart(line_rr(trd_y, "MONTH", "VOLUME", f"{label} – Cumulative YTD Volume by Year", color="YEAR",
                                order=ORDER_M, cum=True), width="stretch")
        c3, c4 = st.columns(2)
        c3.plotly_chart(yoy_bar(trd_all, f"{label} – Monthly Volume YoY % ({CY} vs {LY})"), width="stretch")
        c4.plotly_chart(rate_line(trd_y, "MONTH", "WEIGHTED_MDR", f"{label} – Weighted MDR by Year (not run-rated)",
                                  color="YEAR", order=ORDER_M), width="stretch")
        if seg3 == "Payment Facilitator" and sub_pick == "All":
            st.markdown(f"**YTD YoY – sub-merchants under {pick}**")
            sy = ytd_compare(dd, SUB_COL)
            sy.loc[sy.VOLUME_LY.fillna(0) < MIN_LY_BASE, "VOLUME_YoY%"] = float("nan")
            show_ytd(sy.sort_values("VOLUME", ascending=False), SUB_COL)
        st.markdown(f"#### ② Month on month – {label}: {cur_m}{RR_TAG} vs {mom_m}")
        kpi_row(period_table(dd.assign(ALL=label), "ALL"))
        trd = trd_all.sort_values("YEAR_MONTH")
        trd = trd[trd.YEAR_MONTH >= prev_month(cur_m, 12)]
        c1, c2 = st.columns(2)
        c1.plotly_chart(bar_rr(trd, "LABEL", "VOLUME", f"{label} – Monthly Volume (last 13 months)"), width="stretch")
        c2.plotly_chart(rate_line(trd, "LABEL", ["WEIGHTED_MDR", "WEIGHTED_IC"],
                                  f"{label} – Weighted MDR vs Interchange (not run-rated)"), width="stretch")
        if seg3 == "Payment Facilitator" and sub_pick == "All":
            st.markdown(f"**MoM – sub-merchants under {pick}**")
            show_table(period_table(dd, SUB_COL).sort_values("VOLUME", ascending=False), SUB_COL)

# ================================================================ TAB 6 – PF LEVEL (MID → SUB-MERCHANT)
with tab_pf:
    if need_card():
        dpf = df[df.PF_DM_TAG == "Payment Facilitator"]
        if dpf.empty:
            st.info("No Payment Facilitator rows in the Card MIS.")
        else:
            st.caption(f"Hierarchy: **PF** (PF_DM_NAME) → **MID** (CARD_ACCEPTOR_ID) → **Sub-merchant** "
                       f"(MERCHANT_NAME_ORIG). YTD {ytd_lbl} {CY}{' incl. run-rated ' + MON[END_MM] if CUR_FACTOR != 1 else ''} "
                       f"vs {ytd_lbl} {LY} · month {cur_m}{RR_TAG} vs {mom_m} (MoM) and {yoy_m} (YoY). "
                       f"YoY shown as '-' where the LY base is too small.")
            st.subheader("All Payment Facilitators – YTD, YoY, MoM")
            allpf = full_compare(dpf, "PF_DM_NAME").sort_values("YTD__VOLUME", ascending=False)
            allpf["YTD Share"] = allpf.YTD__VOLUME / allpf.YTD__VOLUME.sum()
            show_full(allpf, "PF_DM_NAME", extra=["YTD Share"], height=380)

            st.divider()
            pf = st.selectbox("Select Payment Facilitator", allpf.PF_DM_NAME.dropna().tolist(), key="pf_pick")
            dp = dpf[dpf.PF_DM_NAME == pf]
            st.markdown(f"#### ① {pf} – YoY & YTD")
            ytd_kpi_block(dp, pf)
            cur_rows, pm_rows, ly_rows = dp[dp.YEAR_MONTH == cur_m], dp[dp.YEAR_MONTH == mom_m], dp[dp.YEAR_MONTH == yoy_m]
            c = st.columns(4)
            n_mid, n_sub = cur_rows[MID_COL].nunique(), cur_rows[SUB_COL].nunique()
            c[0].metric(f"Active MIDs ({cur_m})", f"{n_mid:,}",
                        f"{n_mid - pm_rows[MID_COL].nunique():+,} MoM | {n_mid - ly_rows[MID_COL].nunique():+,} YoY")
            c[1].metric(f"Active sub-merchants ({cur_m})", f"{n_sub:,}",
                        f"{n_sub - pm_rows[SUB_COL].nunique():+,} MoM | {n_sub - ly_rows[SUB_COL].nunique():+,} YoY")
            new_subs = set(cur_rows[SUB_COL]) - set(dp[dp.YEAR_MONTH < cur_m][SUB_COL])
            c[2].metric(f"New sub-merchants in {cur_m}", f"{len(new_subs):,}", "first-ever transaction month",
                        delta_color="off")
            c[3].metric(f"Sub-merchants transacting YTD {CY}", f"{ytd_slice(dp, CY, END_MM)[SUB_COL].nunique():,}",
                        f"LY YTD {ytd_slice(dp, LY, END_MM)[SUB_COL].nunique():,}", delta_color="off")
            tp_all = trend_actual(dp)
            tp = year_filter(tp_all, "yr_pf")
            c1, c2 = st.columns(2)
            c1.plotly_chart(line_rr(tp, "MONTH", "VOLUME", f"{pf} – Monthly Volume by Year", color="YEAR",
                                    order=ORDER_M), width="stretch")
            c2.plotly_chart(line_rr(tp, "MONTH", "VOLUME", f"{pf} – Cumulative YTD Volume", color="YEAR",
                                    order=ORDER_M, cum=True), width="stretch")
            c3, c4 = st.columns(2)
            c3.plotly_chart(yoy_bar(tp_all, f"{pf} – Monthly Volume YoY % ({CY} vs {LY})"), width="stretch")
            c4.plotly_chart(rate_line(tp, "MONTH", "WEIGHTED_MDR", f"{pf} – Weighted MDR by Year (not run-rated)",
                                      color="YEAR", order=ORDER_M), width="stretch")
            st.markdown(f"#### ② {pf} – Month on month: {cur_m}{RR_TAG} vs {mom_m}")
            kpi_row(period_table(dp.assign(ALL=pf), "ALL"))

            st.markdown(f"#### ③ MIDs under {pf}")
            mids = full_compare(dp, MID_COL).sort_values("YTD__VOLUME", ascending=False)
            mids["YTD Share of PF"] = mids.YTD__VOLUME / mids.YTD__VOLUME.sum()
            ind = (dp.groupby([MID_COL, "INDUSTRY"]).TOTAL_VOLUME.sum().reset_index()
                     .sort_values("TOTAL_VOLUME", ascending=False).drop_duplicates(MID_COL)
                     .rename(columns={"INDUSTRY": "Main industry"})[[MID_COL, "Main industry"]])
            mids = mids.merge(ind, on=MID_COL, how="left")
            show_full(mids, [MID_COL, "Main industry"], extra=["YTD Share of PF"], height=320)
            tm = trend_actual(dp, MID_COL)
            top_mids = mids[MID_COL].head(MID_TOP_N).tolist()
            tm["MID_GRP"] = tm[MID_COL].where(tm[MID_COL].isin(top_mids), "Other MIDs")
            tm = year_filter(tm, "yr_pf_mid", default_last=2, label="Years shown in MID chart")
            tm = tm.groupby(["YEAR_MONTH", "LABEL", "MID_GRP", "_F"], as_index=False).VOLUME.sum().sort_values("YEAR_MONTH")
            c1, c2 = st.columns(2)
            c1.plotly_chart(bar_rr(tm, "LABEL", "VOLUME", f"{pf} – Monthly Volume by MID (top {MID_TOP_N})",
                                   color="MID_GRP"), width="stretch")
            c2.plotly_chart(ytd_vs_ly_bar(mids, MID_COL, f"{pf} – YTD {CY} vs YTD {LY} by MID"), width="stretch")

            st.markdown("#### ④ Sub-merchants under each MID")
            cA, cB = st.columns([2, 1])
            mid_idx = mids.set_index(MID_COL)
            mid_opts = ["All MIDs"] + mids[MID_COL].tolist()

            def _mid_lbl(m):
                if m == "All MIDs":
                    return m
                r = mid_idx.loc[m]
                n = int(r["N_SUB"]) if pd.notna(r["N_SUB"]) else 0
                return f"{m} · {r['Main industry']} · {n} subs"

            mid_pick = cA.selectbox("MID", mid_opts, index=1 if len(mid_opts) > 1 else 0, key="pf_mid_pick",
                                    format_func=_mid_lbl)
            rank_by = cB.radio("Rank sub-merchants by", ["YTD__VOLUME", "VOLUME"], horizontal=True, key="pf_rank",
                               format_func=lambda x: f"YTD {CY}" if x.startswith("YTD") else f"{cur_m}{RR_TAG}")
            dm = dp if mid_pick == "All MIDs" else dp[dp[MID_COL] == mid_pick]
            subs = full_compare(dm, [MID_COL, SUB_COL]).sort_values(rank_by, ascending=False)
            subs["Share of MID (YTD)"] = subs.YTD__VOLUME / subs.groupby(MID_COL).YTD__VOLUME.transform("sum")
            subs = subs.drop(columns=["N_MID", "N_SUB"])
            tmap = subs[subs.YTD__VOLUME > 0]
            if not tmap.empty:
                st.plotly_chart(px.treemap(tmap, path=[px.Constant(pf), MID_COL, SUB_COL], values="YTD__VOLUME",
                                           color="YTD__VOLUME_YoY%", color_continuous_scale="RdYlGn",
                                           color_continuous_midpoint=0, range_color=[-1, 1],
                                           title=f"{pf} → MID → Sub-merchant · size = YTD {CY} volume, "
                                                 f"colour = YTD YoY % (grey = new / n.m.)")
                                .update_traces(texttemplate="%{label}<br>%{percentParent:.1%}")
                                .update_layout(height=560, margin=dict(t=50, l=5, r=5, b=5)), width="stretch")
            top_s = st.slider("Top N sub-merchants", 5, 40, 15, key="pf_sub_n")
            ts = subs.head(top_s)
            c1, c2 = st.columns(2)
            c1.plotly_chart(bar_rr(rr_snapshot(ts, SUB_COL, "VOLUME"), SUB_COL, "VOLUME",
                                   f"Top {top_s} sub-merchants – {cur_m} volume", horizontal=True,
                                   height=max(400, 30 * len(ts))), width="stretch")
            c2.plotly_chart(ytd_vs_ly_bar(ts.drop_duplicates(SUB_COL), SUB_COL,
                                          f"Top {top_s} sub-merchants – YTD {CY} vs YTD {LY}", top=top_s), width="stretch")
            show_full(subs, [MID_COL, SUB_COL], extra=["Share of MID (YTD)"], height=420)
            sub_pick2 = st.selectbox("Sub-merchant trend", subs[SUB_COL].drop_duplicates().tolist(), key="pf_sub_pick")
            ds = dm[dm[SUB_COL] == sub_pick2]
            ytd_kpi_block(ds, sub_pick2)
            ts_tr = trend_actual(ds).sort_values("YEAR_MONTH")
            ts_tr = ts_tr[ts_tr.YEAR_MONTH >= prev_month(cur_m, 12)]
            c1, c2 = st.columns(2)
            c1.plotly_chart(bar_rr(ts_tr, "LABEL", "VOLUME", f"{sub_pick2} – Monthly Volume (last 13 months)"),
                            width="stretch")
            c2.plotly_chart(rate_line(ts_tr, "LABEL", ["WEIGHTED_MDR", "WEIGHTED_IC"],
                                      f"{sub_pick2} – Weighted MDR vs Interchange (not run-rated)"), width="stretch")

# ================================================================ TAB 7 – PROFITABILITY MAP
with tab_prof:
    if need_card():
        if NI_COL not in df.columns:
            st.info(f"The loaded card file has no {NI_COL} column – the profitability map needs it.")
        else:
            st.caption(
                f"Profitability = **{NI_COL}** as reported in the MIS (actuals – not run-rated). "
                "**NI margin (bps)** = NET_INCOME ÷ volume × 10,000; **Spread (bps)** = weighted MDR − weighted "
                "interchange. Quadrants: **high volume** = merchants that together make up the top X% of the "
                "segment's volume; **profitable** = NI margin ≥ hurdle.")
            cA, cB, cC = st.columns([1.3, 1.3, 1.4])
            per = cA.radio("Period", ["YTD", "MONTH"], horizontal=True, key="prof_per",
                           format_func=lambda x: f"YTD {CY} ({ytd_lbl})" if x == "YTD"
                           else f"{cur_m}{' (MTD)' if CUR_FACTOR != 1 else ''}")
            pareto = cB.select_slider("High volume = top X% of segment volume", [50, 60, 70, 80, 90, 95],
                                      value=80, key="prof_pareto")
            hurdle = cC.number_input("NI margin hurdle (bps)", value=0.0, step=5.0, key="prof_hurdle")
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
                loss = g.NI_BPS < hurdle
                rows.append({"Segment": s, "Merchants": len(g), "Volume": g.VOLUME.sum(), "Net income": g.NI.sum(),
                             "NI margin (bps)": g.NI.sum() / g.VOLUME.sum() * 1e4,
                             "Spread (bps)": ((g.WTD_MDR - g.WTD_IC) * g.VOLUME).sum() / g.VOLUME.sum() * 1e4,
                             "# below hurdle": int(loss.sum()),
                             "Vol share below hurdle": g.loc[loss, "VOLUME"].sum() / g.VOLUME.sum(),
                             "NI of below-hurdle": g.loc[loss, "NI"].sum()})
            if rows:
                comp = pd.DataFrame(rows)
                st.dataframe(comp.style.format({"Volume": "{:,.0f}", "Net income": "{:,.0f}",
                                                "NI margin (bps)": "{:+.1f}", "Spread (bps)": "{:+.1f}",
                                                "Vol share below hurdle": "{:.1%}", "NI of below-hurdle": "{:,.0f}"},
                                               na_rep="-"), width="stretch", hide_index=True)

            # ---- one segment
            st.divider()
            cA, cB = st.columns([1.3, 2])
            seg_p = cA.radio("Segment", ["Direct Merchant", "Payment Facilitator"], horizontal=True, key="prof_seg")
            lvl_map = {"PF_DM_NAME": ["PF_DM_NAME"]}
            if seg_p == "Payment Facilitator":
                lvl_map.update({MID_COL: ["PF_DM_NAME", MID_COL], SUB_COL: ["PF_DM_NAME", SUB_COL]})
            lvl = cB.radio("Level", list(lvl_map), horizontal=True, key=f"prof_lvl_{seg_p}",
                           format_func={"PF_DM_NAME": "Merchant / PF", MID_COL: "MID",
                                        SUB_COL: "Sub-merchant"}.get)
            dseg = dper[dper.PF_DM_TAG == seg_p]
            g = profit_table(dseg, lvl_map[lvl])
            if g.empty:
                st.info("No volume for this segment / period.")
            else:
                g, vol_cut = assign_quadrants(g, pareto, hurdle)
                tot_v, tot_ni = g.VOLUME.sum(), g.NI.sum()
                below = g.NI_BPS < hurdle
                c = st.columns(5)
                c[0].metric(f"Volume ({per_lbl})", fmt_php(tot_v))
                c[1].metric("Net income", fmt_php(tot_ni))
                c[2].metric("NI margin", f"{tot_ni / tot_v * 1e4:+.1f} bps")
                c[3].metric(f"Below hurdle ({hurdle:+.0f} bps)", f"{int(below.sum()):,} of {len(g):,}",
                            f"{g.loc[below, 'VOLUME'].sum() / tot_v:.1%} of volume", delta_color="off")
                c[4].metric("NI drag from below-hurdle", fmt_php(g.loc[below, "NI"].sum()),
                            f"vs {fmt_php(g.loc[~below, 'NI'].sum())} from the rest", delta_color="off")

                # quadrant summary
                qs = (g.groupby("QUADRANT").agg(Entities=("NAME", "count"), Volume=("VOLUME", "sum"),
                                                Net_income=("NI", "sum")).reindex(list(QUAD_COLORS)).fillna(0)
                      .reset_index())
                qs["Vol share"] = qs.Volume / tot_v
                qs["NI margin (bps)"] = qs.Net_income / qs.Volume.where(qs.Volume > 0) * 1e4
                st.dataframe(qs.rename(columns={"QUADRANT": "Quadrant", "Net_income": "Net income"})
                             .style.format({"Entities": "{:,.0f}", "Volume": "{:,.0f}", "Net income": "{:,.0f}",
                                            "Vol share": "{:.1%}", "NI margin (bps)": "{:+.1f}"}, na_rep="-"),
                             width="stretch", hide_index=True)

                # map: volume (log) vs NI margin; y clipped to 2nd–98th pct so a few tiny accounts don't squash it
                lo, hi = np.nanpercentile(g.NI_BPS, [2, 98])
                lo, hi = min(lo, hurdle - 10), max(hi, hurdle + 10)
                pad = (hi - lo) * 0.08
                g["Y"] = g.NI_BPS.clip(lo - pad, hi + pad)
                g["Shown"] = np.where(g.NI_BPS.between(lo - pad, hi + pad), "in range", "clipped (outlier)")
                g["SIZE"] = g.NI.abs()
                fig = px.scatter(g, x="VOLUME", y="Y", color="QUADRANT", size="SIZE", size_max=40, log_x=True,
                                 symbol="Shown", symbol_map={"in range": "circle", "clipped (outlier)": "triangle-up"},
                                 color_discrete_map=QUAD_COLORS, category_orders={"QUADRANT": list(QUAD_COLORS)},
                                 hover_name="NAME",
                                 hover_data={"VOLUME": ":,.0f", "NI": ":,.0f", "NI_BPS": ":+.1f",
                                             "SPREAD_BPS": ":+.1f", "TICKET": ":,.0f", "Y": False, "SIZE": False,
                                             "Shown": False, "QUADRANT": False},
                                 labels={"VOLUME": f"Volume ({per_lbl}, log scale)", "Y": "NI margin (bps)",
                                         "NI": "Net income", "NI_BPS": "NI margin bps", "SPREAD_BPS": "Spread bps",
                                         "TICKET": "Avg ticket"},
                                 title=f"Profitability map – {seg_p} ({per_lbl}) · bubble size = |net income|")
                fig.add_hline(y=hurdle, line_dash="dash", line_color="grey",
                              annotation_text=f"hurdle {hurdle:+.0f} bps", annotation_position="top left")
                fig.add_vline(x=vol_cut, line_dash="dash", line_color="grey",
                              annotation_text=f"high-volume cut-off {fmt_php(vol_cut)}", annotation_position="top")
                fig.update_layout(height=620, legend_title_text="")
                st.plotly_chart(fig, width="stretch")

                # spread vs NI margin – explains WHY a merchant is unprofitable
                c1, c2 = st.columns(2)
                f2 = px.scatter(g, x="SPREAD_BPS", y="Y", color="QUADRANT", size="VOLUME", size_max=30,
                                color_discrete_map=QUAD_COLORS, hover_name="NAME",
                                hover_data={"SPREAD_BPS": ":+.1f", "NI_BPS": ":+.1f", "TICKET": ":,.0f",
                                            "Y": False, "VOLUME": ":,.0f", "QUADRANT": False},
                                labels={"SPREAD_BPS": "MDR − interchange spread (bps)", "Y": "NI margin (bps)"},
                                title="Spread vs NI margin – positive spread but negative NI ⇒ cost-driven "
                                      "(fixed / per-txn costs, small tickets)")
                f2.add_hline(y=hurdle, line_dash="dash", line_color="grey")
                f2.add_vline(x=0, line_dash="dash", line_color="grey")
                c1.plotly_chart(f2.update_layout(height=480, showlegend=False), width="stretch")
                movers = pd.concat([g.nsmallest(10, "NI"), g.nlargest(10, "NI")]).drop_duplicates("NAME") \
                    .sort_values("NI")
                c2.plotly_chart(px.bar(movers, x="NI", y="NAME", orientation="h", color="QUADRANT",
                                       color_discrete_map=QUAD_COLORS, title="Top 10 profit and loss contributors (₱)",
                                       hover_data={"NI_BPS": ":+.1f", "VOLUME": ":,.0f"})
                                .update_layout(height=480, showlegend=False).update_yaxes(type="category", title=None),
                                width="stretch")

                tbl = g.sort_values("NI")[["NAME", "QUADRANT", "VOLUME", "VOL_SHARE", "TXN_COUNT", "TICKET", "NI",
                                           "NI_BPS", "WTD_MDR", "WTD_IC", "SPREAD_BPS", "MONTHS"]]
                tbl.columns = ["Name", "Quadrant", "Volume", "Vol share", "Txn", "Avg ticket", "Net income",
                               "NI margin (bps)", "Wtd MDR", "Wtd IC", "Spread (bps)", "Active months"]
                st.dataframe(tbl.style.format({"Volume": "{:,.0f}", "Vol share": "{:.2%}", "Txn": "{:,.0f}",
                                               "Avg ticket": "{:,.0f}", "Net income": "{:,.0f}",
                                               "NI margin (bps)": "{:+.1f}", "Wtd MDR": "{:.2%}", "Wtd IC": "{:.2%}",
                                               "Spread (bps)": "{:+.1f}"}, na_rep="-"),
                             width="stretch", hide_index=True, height=420)
                st.download_button("Download profitability table (CSV)", tbl.to_csv(index=False).encode(),
                                   file_name=f"profitability_{seg_p.replace(' ', '_')}_{per_lbl}.csv",
                                   key="prof_dl")

                # trend for one entity
                pick_n = st.selectbox("Monthly NI trend for", g.sort_values("VOLUME", ascending=False).NAME.tolist(),
                                      key=f"prof_pick_{seg_p}_{lvl}")
                keys = lvl_map[lvl]
                sel_row = g[g.NAME == pick_n].iloc[0]
                mask = np.logical_and.reduce([df[k].astype(str) == str(sel_row[k]) for k in keys])
                dsel = df[(df.PF_DM_TAG == seg_p) & mask & (df.YEAR_MONTH <= cur_m)
                          & (df.YEAR_MONTH >= prev_month(cur_m, 12))]
                mt = profit_table(dsel, "YEAR_MONTH").sort_values("YEAR_MONTH")
                c1, c2 = st.columns(2)
                c1.plotly_chart(px.bar(mt, x="YEAR_MONTH", y="NI", title=f"{pick_n} – monthly net income (₱, actual)",
                                       color=(mt.NI >= 0).map({True: "Profit", False: "Loss"}),
                                       color_discrete_map={"Profit": "#2E8B57", "Loss": "#C0392B"})
                                .update_layout(showlegend=False).update_xaxes(type="category"), width="stretch")
                c2.plotly_chart(px.line(mt, x="YEAR_MONTH", y=["NI_BPS", "SPREAD_BPS"], markers=True,
                                        title=f"{pick_n} – NI margin vs spread (bps)")
                                .add_hline(y=hurdle, line_dash="dash", line_color="grey")
                                .update_layout(legend_title_text="").update_xaxes(type="category"), width="stretch")
                if cur_m == LATEST_M and CUR_FACTOR != 1:
                    st.caption(f"{cur_m} is month-to-date (as of {card_asof:%b %d}); fixed monthly costs may make "
                               "its NI margin look worse than a full month.")

# ================================================================ TAB 8 – QRPH
with tab_qr:
    if not HAS_QR:
        st.info("Upload the QRPH sheet in the sidebar (needs YEAR_MONTH, VOLUME_POSTED, TXN_COUNT_POSTED).")
    else:
        st.caption(f"Basis: **posted** volume and txn count only · {QR_LATEST} run-rated ×{QR_RR:.2f} · "
                   f"FY outlook = YTD ÷ {END_MM} × 12")
        if cur_m not in set(qr.YEAR_MONTH):
            st.warning(f"QRPH sheet has no data for {cur_m}.")
        res_qr = series_metrics(rr_series(qr_a, qr_factor))
        series_kpis(res_qr, "QRPH", rr=(cur_m == QR_LATEST and QR_RR != 1))
        tk = res_qr["TICKET"]
        st.caption(f"Avg ticket (posted): {fmt_php(tk['CUR'])} · {fmt_pct(tk['MoM%'])} MoM · {fmt_pct(tk['YoY%'])} YoY · "
                   f"YTD {fmt_php(tk['YTD'])} ({fmt_pct(tk['YTD_YoY%'])} YoY)")
        st.subheader("QRPH comparison table")
        series_table({"QRPH": res_qr})
        qa = add_period_cols(qr_a)
        qa["TICKET"] = qa.VOLUME / qa.TXN_COUNT
        qo = year_filter(qa, "yr_qr")
        c1, c2 = st.columns(2)
        c1.plotly_chart(line_rr(qo, "MONTH", "VOLUME", "QRPH Monthly Posted Volume by Year", color="YEAR",
                                order=ORDER_M), width="stretch")
        c2.plotly_chart(line_rr(qo, "MONTH", "TXN_COUNT", "QRPH Monthly Posted Txn Count by Year", color="YEAR",
                                order=ORDER_M), width="stretch")
        c3, c4 = st.columns(2)
        c3.plotly_chart(line_rr(qo, "MONTH", "VOLUME", "QRPH Cumulative YTD Posted Volume", color="YEAR",
                                order=ORDER_M, cum=True), width="stretch")
        c4.plotly_chart(yoy_bar(qa, f"QRPH Monthly Volume YoY % ({CY} vs {LY})"), width="stretch")
        st.plotly_chart(px.line(qo, x="MONTH", y="TICKET", color="YEAR", markers=True, category_orders=ORDER_M,
                                title="QRPH Avg Ticket Size (₱, not run-rated)"), width="stretch")
