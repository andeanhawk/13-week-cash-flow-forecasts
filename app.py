# CODE REFINED , MINOR CHANGES DONE
import io
import pathlib
import datetime as dt
import logging
import sys

import numpy as np
import pandas as pd

# UI backend
import streamlit as st

# matplotlib optional import (we'll handle failures)
try:
    import matplotlib.pyplot as plt
    MATPLOTLIB_AVAILABLE = True
except Exception:
    MATPLOTLIB_AVAILABLE = False

# Configure simple logging to Streamlit output
logger = logging.getLogger("cashflow_app")
if not logger.handlers:
    handler = logging.StreamHandler(stream=sys.stdout)
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    logger.addHandler(handler)
logger.setLevel(logging.INFO)

st.set_page_config(page_title="13-Week Cash Flow — Excel App", layout="wide")
st.title("13-Week Cash Flow Forecaster — Excel App")
st.write("Upload an Excel with sheets `Inputs`, `Receipts`, `Disbursements` (H&M format).")

# --- CONFIG ---
MAX_UPLOAD_MB = 12  # reject files larger than this
ALLOWED_EXT = {".xlsx", ".xls", ".xlsm"}
WEEKS = 13

# --- HELPERS ---
def human_mb(nbytes):
    return f"{nbytes/1024/1024:.1f} MB"

def normalize_probability(v):
    """Accepts numeric like 0.8, 80, or strings like '80%' and returns 0..1"""
    if pd.isna(v):
        return 1.0
    try:
        if isinstance(v, str):
            s = v.strip().replace("%", "")
            val = float(s)
        else:
            val = float(v)
        if val > 1:
            return val / 100.0
        return max(0.0, min(1.0, val))
    except Exception:
        return 1.0

def safe_to_datetime(series):
    return pd.to_datetime(series, errors="coerce")

# --- FILE LOADER ---
@st.cache_data(show_spinner=False)
def load_from_excel(file_name: str, file_bytes: bytes):
    """
    Returns (inputs_df, receipts_df, disb_df) or raises ValueError with user-friendly msg.
    """
    suffix = pathlib.Path(file_name.lower()).suffix
    if suffix not in ALLOWED_EXT:
        raise ValueError(f"Unsupported file type: {suffix}. Please upload {', '.join(ALLOWED_EXT)}.")

    bio = io.BytesIO(file_bytes)
    # Try engines with informative errors
    try:
        if suffix in (".xlsx", ".xlsm"):
            xl = pd.ExcelFile(bio, engine="openpyxl")
        elif suffix == ".xls":
            xl = pd.ExcelFile(bio, engine="xlrd")
    except Exception as e:
        logger.exception("Excel file read error")
        raise ValueError(f"Failed to read Excel file: {e}")

    def safe_parse(name: str):
        try:
            df = xl.parse(name)
            return df
        except Exception:
            return None

    inputs = safe_parse("Inputs")
    receipts = safe_parse("Receipts")
    disb = safe_parse("Disbursements")
    return inputs, receipts, disb

# --- COERCION / CLEANING ---
def coerce_receipts_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.shape[0] == 0:
        return pd.DataFrame(columns=["Date", "Source", "Category", "Amount", "Probability"])
    out = pd.DataFrame()
    # Flexible column names
    out["Date"] = safe_to_datetime(df.get("Date", df.iloc[:, 0]))
    out["Source"] = df.get("Customer/Source", df.get("Source", df.columns[1] if len(df.columns) > 1 else df.columns[0])).astype(str)
    out["Category"] = df.get("Category", "Sales Receipts").astype(str)
    # find amount column
    amt_candidates = [c for c in df.columns if "amount" in str(c).lower() and "expected" not in str(c).lower()]
    amt_col = amt_candidates[0] if amt_candidates else df.columns[min(3, len(df.columns)-1)]
    out["Amount"] = pd.to_numeric(df[amt_col], errors="coerce").fillna(0.0)
    # probability column
    prob_candidates = [c for c in df.columns if "prob" in str(c).lower()]
    prob_col = prob_candidates[0] if prob_candidates else None
    out["Probability"] = out.apply(lambda r: normalize_probability(df.at[r.name, prob_col]) if prob_col in df.columns else 1.0, axis=1) if prob_col else 1.0
    # If prob_col not detected, try to read a column named like 'Probability' directly
    if prob_col is None and "Probability" in df.columns:
        out["Probability"] = df["Probability"].apply(normalize_probability)
    # Ensure Date exists and drop invalid rows
    bad_dates = out["Date"].isna().sum()
    if bad_dates:
        st.warning(f"Dropped {int(bad_dates)} receipt rows with invalid dates.")
    out = out.dropna(subset=["Date"])
    return out[["Date", "Source", "Category", "Amount", "Probability"]]

def coerce_disb_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.shape[0] == 0:
        return pd.DataFrame(columns=["Date", "Payee", "Category", "Amount", "Probability"])
    out = pd.DataFrame()
    out["Date"] = safe_to_datetime(df.get("Date", df.iloc[:, 0]))
    out["Payee"] = df.get("Vendor/Payee", df.get("Payee", df.columns[1] if len(df.columns) > 1 else df.columns[0])).astype(str)
    out["Category"] = df.get("Category", "Vendors").astype(str)
    amt_candidates = [c for c in df.columns if "amount" in str(c).lower() and "expected" not in str(c).lower()]
    amt_col = amt_candidates[0] if amt_candidates else df.columns[min(3, len(df.columns)-1)]
    out["Amount"] = pd.to_numeric(df[amt_col], errors="coerce").fillna(0.0)
    prob_candidates = [c for c in df.columns if "prob" in str(c).lower()]
    prob_col = prob_candidates[0] if prob_candidates else None
    out["Probability"] = out.apply(lambda r: normalize_probability(df.at[r.name, prob_col]) if prob_col in df.columns else 1.0, axis=1) if prob_col else 1.0
    if prob_col is None and "Probability" in df.columns:
        out["Probability"] = df["Probability"].apply(normalize_probability)
    bad_dates = out["Date"].isna().sum()
    if bad_dates:
        st.warning(f"Dropped {int(bad_dates)} disbursement rows with invalid dates.")
    out = out.dropna(subset=["Date"])
    return out[["Date", "Payee", "Category", "Amount", "Probability"]]

# --- INPUT PARSING ---
def parse_inputs(df: pd.DataFrame):
    oc = 0.0; min_cash = 0.0; rf = 1.0; dfac = 1.0
    start_date = dt.date.today() - dt.timedelta(days=dt.date.today().weekday())
    if df is None or df.shape[0] == 0:
        return oc, start_date, min_cash, rf, dfac
    try:
        mapping = {}
        # tolerate layout with label/value per row
        for i in range(min(len(df), 200)):
            key = str(df.iloc[i,0]).strip().lower()
            val = df.iloc[i,1] if df.shape[1] > 1 else None
            mapping[key] = val
        def find_num(key, default):
            for k,v in mapping.items():
                if key in k:
                    try: return float(v)
                    except: 
                        try: return float(pd.to_numeric(v))
                        except: return default
            return default
        def find_date(key, default):
            for k,v in mapping.items():
                if key in k:
                    try: return pd.to_datetime(v).date()
                    except: return default
            return default
        oc = find_num("opening cash", oc)
        min_cash = find_num("min cash", min_cash)
        rf = find_num("receipts factor", rf)
        dfac = find_num("disbursements factor", dfac)
        start_date = find_date("start date", start_date)
    except Exception as e:
        logger.exception("Failed parsing Inputs sheet")
    return oc, start_date, min_cash, rf, dfac

# --- FORECAST LOGIC ---
def weekly_buckets(start_date_):
    starts = [pd.to_datetime(start_date_) + pd.Timedelta(days=7*i) for i in range(WEEKS)]
    ends = [d + pd.Timedelta(days=6) for d in starts]
    labels = [f"Week {i+1}\n{starts[i].date()}" for i in range(WEEKS)]
    return starts, ends, labels

def build_forecast(receipts, disb, opening_cash, min_cash, start_date, rcpt_factor, dsb_factor):
    # ensure columns
    if receipts is None: receipts = pd.DataFrame(columns=["Date","Category","Amount","Probability"])
    if disb is None: disb = pd.DataFrame(columns=["Date","Category","Amount","Probability"])
    r = receipts.copy(); d = disb.copy()
    # Ensure necessary columns exist
    for col in ["Amount","Probability","Date","Category"]:
        if col not in r.columns:
            r[col] = 0.0 if col=="Amount" else 1.0 if col=="Probability" else pd.NaT if col=="Date" else ""
        if col not in d.columns:
            d[col] = 0.0 if col=="Amount" else 1.0 if col=="Probability" else pd.NaT if col=="Date" else ""
    r["Expected"] = r["Amount"].fillna(0.0) * r["Probability"].fillna(1.0) * float(rcpt_factor)
    d["Expected"] = d["Amount"].fillna(0.0) * d["Probability"].fillna(1.0) * float(dsb_factor)

    ws, we, labels = weekly_buckets(start_date)

    def sum_week(df, cat_col=None, category=None):
        vals = []
        for i in range(WEEKS):
            sub = df[(df["Date"] >= ws[i]) & (df["Date"] <= we[i])]
            if cat_col and category is not None:
                sub = sub[sub[cat_col] == category]
            vals.append(float(sub["Expected"].sum()))
        return np.array(vals)

    rcpt_cats = sorted(r["Category"].dropna().unique().tolist())
    dsb_cats = sorted(d["Category"].dropna().unique().tolist())

    idx = (["Opening Cash","Receipts"] + [f"  {c}" for c in rcpt_cats] +
           ["Disbursements"] + [f"  {c}" for c in dsb_cats] +
           ["Net Cash Flow","Ending Cash"])
    fc = pd.DataFrame(index=idx, columns=labels, dtype=float)

    rcpt_tot = np.zeros(WEEKS)
    for c in rcpt_cats:
        vals = sum_week(r, "Category", c); fc.loc[f"  {c}"] = vals; rcpt_tot += vals
    fc.loc["Receipts"] = rcpt_tot

    dsb_tot = np.zeros(WEEKS)
    for c in dsb_cats:
        vals = sum_week(d, "Category", c); fc.loc[f"  {c}"] = vals; dsb_tot += vals
    fc.loc["Disbursements"] = dsb_tot

    net = rcpt_tot - dsb_tot
    fc.loc["Net Cash Flow"] = net

    ending = np.zeros(WEEKS)
    for i in range(WEEKS):
        open_i = opening_cash if i == 0 else ending[i-1]
        fc.loc["Opening Cash", labels[i]] = open_i
        ending[i] = open_i + net[i]
    fc.loc["Ending Cash"] = ending
    return fc, ws, we

# --- EXPORT ---
def export_excel(fc, opening_cash, start_date, min_cash, rcpt_factor, dsb_factor, currency_unit):
    try:
        import xlsxwriter
    except Exception as e:
        raise RuntimeError("xlsxwriter is required to export Excel.") from e

    bio = io.BytesIO()
    wb = xlsxwriter.Workbook(bio, {"in_memory": True})
    fmt_title = wb.add_format({"bold": True, "font_size": 16})
    fmt_h2 = wb.add_format({"bold": True, "font_size": 12})
    fmt_hdr = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1})
    fmt_num = wb.add_format({"num_format": "#,##0.0"})
    fmt_date = wb.add_format({"num_format": "yyyy-mm-dd"})
    fmt_red = wb.add_format({"bg_color": "#FDE9E9"})

    wsI = wb.add_worksheet("Inputs")
    wsI.write("A1", "Inputs", fmt_h2)
    wsI.write("A3", "Currency / Unit"); wsI.write("B3", currency_unit)
    wsI.write("A4", "Opening Cash"); wsI.write_number("B4", float(opening_cash), fmt_num)
    wsI.write("A5", "Start Date"); wsI.write_datetime("B5", pd.to_datetime(start_date).to_pydatetime(), fmt_date)
    wsI.write("A6", "Min Cash Threshold"); wsI.write_number("B6", float(min_cash), fmt_num)
    wsI.write("A8", "Receipts Factor"); wsI.write_number("B8", float(rcpt_factor), fmt_num)
    wsI.write("A9", "Disbursements Factor"); wsI.write_number("B9", float(dsb_factor), fmt_num)

    wsF = wb.add_worksheet("Forecast")
    wsF.write("A1", "13-Week Cash Flow Forecast (Direct Method)", fmt_title)
    wsF.write("A3", "Metric", fmt_hdr)
    for j, col in enumerate(fc.columns, start=2):
        wsF.write(2, j-1, col, fmt_hdr)
    for r, idx in enumerate(fc.index, start=4):
        wsF.write(r-1, 0, idx)
        for j, col in enumerate(fc.columns, start=2):
            val = float(fc.loc[idx, col]) if pd.notna(fc.loc[idx, col]) else 0.0
            wsF.write_number(r-1, j-1, val, fmt_num)

    end_row = 4 + list(fc.index).index("Ending Cash")
    wsF.conditional_format(end_row-1, 1, end_row-1, 1+WEEKS, {
        "type": "cell", "criteria": "<", "value": float(min_cash), "format": fmt_red
    })

    wsD = wb.add_worksheet("Dashboard")
    wsD.write("A1", "Dashboard", fmt_title)
    wsD.write("A3", "Key KPIs", fmt_h2)
    wsD.write("A5", "Opening Cash"); wsD.write_number("B5", float(fc.iloc[0,0]), fmt_num)
    wsD.write("A6", "Min Ending Cash"); wsD.write_number("B6", float(fc.loc["Ending Cash"].min()), fmt_num)
    wsD.write("A7", "Total Receipts"); wsD.write_number("B7", float(fc.loc["Receipts"].sum()), fmt_num)
    wsD.write("A8", "Total Disbursements"); wsD.write_number("B8", float(fc.loc["Disbursements"].sum()), fmt_num)

    # charts
    chart1 = wb.add_chart({"type": "line"})
    chart1.add_series({
        "name": "Ending Cash",
        "categories": ["Forecast", 2, 1, 2, 1+WEEKS-1],
        "values":     ["Forecast", end_row-1, 1, end_row-1, 1+WEEKS-1],
        "marker": {"type": "circle", "size": 5},
    })
    chart1.set_title({"name": "Ending Cash (13 Weeks)"})
    chart1.set_x_axis({"name": "Week"})
    chart1.set_y_axis({"name": currency_unit})
    wsD.insert_chart("A12", chart1, {"x_scale": 1.5, "y_scale": 1.2})

    chart2 = wb.add_chart({"type": "column"})
    rc_row = 4 + list(fc.index).index("Receipts")
    ds_row = 4 + list(fc.index).index("Disbursements")
    chart2.add_series({
        "name": "Receipts",
        "categories": ["Forecast", 2, 1, 2, 1+WEEKS-1],
        "values":     ["Forecast", rc_row-1, 1, rc_row-1, 1+WEEKS-1],
    })
    chart2.add_series({
        "name": "Disbursements",
        "categories": ["Forecast", 2, 1, 2, 1+WEEKS-1],
        "values":     ["Forecast", ds_row-1, 1, ds_row-1, 1+WEEKS-1],
    })
    chart2.set_title({"name": "Receipts vs Disbursements (Weekly)"})
    chart2.set_x_axis({"name": "Week"})
    chart2.set_y_axis({"name": currency_unit})
    wsD.insert_chart("H12", chart2, {"x_scale": 1.5, "y_scale": 1.2})

    wb.close()
    bio.seek(0)
    return bio.getvalue()

# --- UI & Main Flow ---
uploaded = st.file_uploader("Upload Excel (.xlsx or .xls)", type=["xlsx", "xls"])
if uploaded is None:
    st.info("Upload an Excel file (H&M format). You can use the HM_Format_Template_Blank.xlsx sample.")
    st.stop()

# file size guard
uploaded.seek(0, io.SEEK_END)
size = uploaded.tell()
uploaded.seek(0)
if size > MAX_UPLOAD_MB * 1024 * 1024:
    st.error(f"File too large: {human_mb(size)}. Max allowed: {MAX_UPLOAD_MB} MB.")
    st.stop()

# load
try:
    inputs_df, receipts_raw, disb_raw = load_from_excel(uploaded.name, uploaded.read())
except ValueError as e:
    st.error(str(e))
    st.stop()
except Exception as e:
    logger.exception("Unexpected loader error")
    st.error(f"Unexpected error reading file: {e}")
    st.stop()

# sheet presence
missing = [n for n, df in {"Inputs": inputs_df, "Receipts": receipts_raw, "Disbursements": disb_raw}.items() if df is None]
if missing:
    st.error(f"Missing sheets: {', '.join(missing)}. Ensure your workbook has these sheets.")
    st.stop()

# parse inputs & data coercion
opening_cash, start_date, min_cash, rcpt_factor, dsb_factor = parse_inputs(inputs_df)

# sidebar overrides
currency_unit = st.sidebar.text_input("Currency / Unit label", "INR lakh")
opening_cash = st.sidebar.number_input("Opening Cash", value=float(opening_cash or 0.0), step=10.0, format="%.1f")
start_date = st.sidebar.date_input("Week 1 start date (Monday)", value=start_date)
min_cash = st.sidebar.number_input("Min Cash Threshold", value=float(min_cash or 0.0), step=10.0, format="%.1f")
st.sidebar.header("Scenario")
rcpt_factor = st.sidebar.slider("Receipts Factor", 0.5, 1.5, float(rcpt_factor or 1.0), 0.05)
dsb_factor = st.sidebar.slider("Disbursements Factor", 0.5, 1.5, float(dsb_factor or 1.0), 0.05)

receipts_df = coerce_receipts_df(receipts_raw)
disb_df = coerce_disb_df(disb_raw)

# Build forecast
try:
    fc, ws, we = build_forecast(receipts_df, disb_df, opening_cash, min_cash, start_date, rcpt_factor, dsb_factor)
except Exception as e:
    logger.exception("Forecast build failed")
    st.error(f"Failed to build forecast: {e}")
    st.stop()

# UI: previews and KPIs
st.subheader("Receipts (preview)")
st.dataframe(receipts_df.head(20))
st.subheader("Disbursements (preview)")
st.dataframe(disb_df.head(20))

c1, c2, c3, c4 = st.columns(4)
try:
    c1.metric("Opening Cash", f"{fc.iloc[0,0]:,.1f} {currency_unit}")
    c2.metric("Min Ending Cash", f"{fc.loc['Ending Cash'].min():,.1f} {currency_unit}")
    c3.metric("Total Receipts", f"{fc.loc['Receipts'].sum():,.1f} {currency_unit}")
    c4.metric("Total Disbursements", f"{fc.loc['Disbursements'].sum():,.1f} {currency_unit}")
except Exception:
    # defensive fallback
    c1.metric("Opening Cash", f"{opening_cash:,.1f} {currency_unit}")

# Charts: try matplotlib, fallback to streamlit charts
if MATPLOTLIB_AVAILABLE:
    try:
        fig1, ax1 = plt.subplots()
        ax1.plot(range(1, WEEKS+1), fc.loc["Ending Cash"].values, marker="o")
        ax1.set_title("Ending Cash (13 Weeks)"); ax1.set_xlabel("Week"); ax1.set_ylabel(currency_unit)
        st.pyplot(fig1)
    except Exception:
        st.warning("Matplotlib plotting failed; using Streamlit charts instead.")
        st.line_chart(fc.loc["Ending Cash"].rename("Ending Cash"))
else:
    st.line_chart(fc.loc["Ending Cash"].rename("Ending Cash"))

# Receipts vs Disbursements
try:
    if MATPLOTLIB_AVAILABLE:
        fig2, ax2 = plt.subplots()
        idx = np.arange(1, WEEKS+1)
        ax2.bar(idx-0.15, fc.loc["Receipts"].values, width=0.3, label="Receipts")
        ax2.bar(idx+0.15, fc.loc["Disbursements"].values, width=0.3, label="Disbursements")
        ax2.set_title("Receipts vs Disbursements (Weekly)")
        ax2.set_xlabel("Week"); ax2.set_ylabel(currency_unit); ax2.legend()
        st.pyplot(fig2)
    else:
        st.bar_chart(pd.DataFrame({"Receipts": fc.loc["Receipts"].values, "Disbursements": fc.loc["Disbursements"].values},
                                 index=[f"W{i}" for i in range(1, WEEKS+1)]))
except Exception:
    st.warning("Plotting receipts vs disbursements failed.")

st.subheader("Forecast Table")
st.dataframe(fc.style.format("{:,.1f}"))

# Download/export
try:
    xlsx_bytes = export_excel(fc, opening_cash, start_date, min_cash, rcpt_factor, dsb_factor, currency_unit)
    st.download_button("⬇️ Download Final Excel", data=xlsx_bytes, file_name="Final_13Week_CashFlow.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
except Exception as e:
    logger.exception("Excel export failed")
    st.error(f"Export failed: {e}")
