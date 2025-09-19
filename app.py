import io
import datetime as dt
import numpy as np
import pandas as pd
import streamlit as st

# -------------------------
# Config
# -------------------------
st.set_page_config(page_title="13-Week Cash Flow Forecaster", layout="wide")
WEEKS = 13

# -------------------------
# Helpers
# -------------------------
def normalize_prob(v):
    """Normalize probability values (0–1, %, 100)."""
    if pd.isna(v):
        return 1.0
    try:
        s = str(v).replace("%", "")
        val = float(s)
        if val > 1:
            return val / 100.0
        return val
    except:
        return 1.0

# -------------------------
# Data Coercion
# -------------------------
def coerce_receipts_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=["Date", "Source", "Category", "Amount", "Probability"])

    out = pd.DataFrame()
    out["Date"] = pd.to_datetime(df.get("Date", df.iloc[:, 0]), errors="coerce")

    if "Customer/Source" in df.columns:
        out["Source"] = df["Customer/Source"].astype(str)
    elif "Source" in df.columns:
        out["Source"] = df["Source"].astype(str)
    else:
        out["Source"] = df.iloc[:, 1].astype(str) if df.shape[1] > 1 else "Unknown"

    if "Category" in df.columns:
        out["Category"] = df["Category"].astype(str)
    else:
        out["Category"] = "Sales Receipts"

    amt_col = next((c for c in df.columns if "amount" in str(c).lower()), df.columns[min(3, df.shape[1]-1)])
    out["Amount"] = pd.to_numeric(df[amt_col], errors="coerce").fillna(0.0)

    if "Probability" in df.columns:
        raw_prob = df["Probability"]
        out["Probability"] = raw_prob.apply(normalize_prob)
    else:
        out["Probability"] = 1.0

    return out.dropna(subset=["Date"]).reset_index(drop=True)


def coerce_disb_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=["Date", "Payee", "Category", "Amount", "Probability"])

    out = pd.DataFrame()
    out["Date"] = pd.to_datetime(df.get("Date", df.iloc[:, 0]), errors="coerce")

    if "Vendor/Payee" in df.columns:
        out["Payee"] = df["Vendor/Payee"].astype(str)
    elif "Payee" in df.columns:
        out["Payee"] = df["Payee"].astype(str)
    else:
        out["Payee"] = df.iloc[:, 1].astype(str) if df.shape[1] > 1 else "Unknown"

    if "Category" in df.columns:
        out["Category"] = df["Category"].astype(str)
    else:
        out["Category"] = "Vendors"

    amt_col = next((c for c in df.columns if "amount" in str(c).lower()), df.columns[min(3, df.shape[1]-1)])
    out["Amount"] = pd.to_numeric(df[amt_col], errors="coerce").fillna(0.0)

    if "Probability" in df.columns:
        raw_prob = df["Probability"]
        out["Probability"] = raw_prob.apply(normalize_prob)
    else:
        out["Probability"] = 1.0

    return out.dropna(subset=["Date"]).reset_index(drop=True)

# -------------------------
# Inputs Parser
# -------------------------
def parse_inputs(df: pd.DataFrame):
    oc, min_cash, rf, dfac = 0.0, 0.0, 1.0, 1.0
    start_date = dt.date.today() - dt.timedelta(days=dt.date.today().weekday())

    if df is None or df.empty:
        return oc, start_date, min_cash, rf, dfac

    mapping = {}
    for i in range(min(len(df), 50)):
        key = str(df.iloc[i, 0]).strip().lower()
        val = df.iloc[i, 1] if df.shape[1] > 1 else None
        mapping[key] = val

    def find_num(key, default):
        for k, v in mapping.items():
            if key in k:
                try:
                    return float(v)
                except:
                    return default
        return default

    def find_date(key, default):
        for k, v in mapping.items():
            if key in k:
                try:
                    return pd.to_datetime(v).date()
                except:
                    return default
        return default

    oc = find_num("opening cash", oc)
    min_cash = find_num("min cash", min_cash)
    rf = find_num("receipts factor", rf)
    dfac = find_num("disbursements factor", dfac)
    start_date = find_date("start date", start_date)
    return oc, start_date, min_cash, rf, dfac

# -------------------------
# Forecast Logic
# -------------------------
def weekly_buckets(start_date):
    starts = [pd.to_datetime(start_date) + pd.Timedelta(days=7 * i) for i in range(WEEKS)]
    ends = [d + pd.Timedelta(days=6) for d in starts]
    labels = [f"Week {i+1}\n{starts[i].date()}" for i in range(WEEKS)]
    return starts, ends, labels


def build_forecast(receipts, disb, opening_cash, start_date, rcpt_factor, dsb_factor):
    ws, we, labels = weekly_buckets(start_date)
    r = receipts.copy()
    d = disb.copy()

    r["Expected"] = r["Amount"] * r["Probability"] * rcpt_factor
    d["Expected"] = d["Amount"] * d["Probability"] * dsb_factor

    def sum_week(df):
        vals = []
        for i in range(WEEKS):
            sub = df[(df["Date"] >= ws[i]) & (df["Date"] <= we[i])]
            vals.append(float(sub["Expected"].sum()))
        return np.array(vals)

    rcpt_tot = sum_week(r)
    dsb_tot = sum_week(d)
    net = rcpt_tot - dsb_tot

    ending = np.zeros(WEEKS)
    for i in range(WEEKS):
        opening = opening_cash if i == 0 else ending[i - 1]
        ending[i] = opening + net[i]

    fc = pd.DataFrame({
        "Receipts": rcpt_tot,
        "Disbursements": dsb_tot,
        "Net Cash Flow": net,
        "Ending Cash": ending
    }, index=labels)

    return fc

# -------------------------
# Streamlit UI
# -------------------------
st.title("📊 13-Week Rolling Cash Flow Forecaster")

uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
if not uploaded:
    st.info("Upload an Excel file with sheets: Inputs, Receipts, Disbursements")
    st.stop()

try:
    xl = pd.ExcelFile(uploaded, engine="openpyxl")
    inputs_df = xl.parse("Inputs")
    receipts_raw = xl.parse("Receipts")
    disb_raw = xl.parse("Disbursements")
except Exception as e:
    st.error(f"Error reading Excel: {e}")
    st.stop()

opening_cash, start_date, min_cash, rcpt_factor, dsb_factor = parse_inputs(inputs_df)

currency_unit = st.sidebar.text_input("Currency/Unit", "INR Lakh")
opening_cash = st.sidebar.number_input("Opening Cash", value=float(opening_cash))
start_date = st.sidebar.date_input("Week 1 Start Date", value=start_date)
min_cash = st.sidebar.number_input("Min Cash Threshold", value=float(min_cash))
rcpt_factor = st.sidebar.slider("Receipts Factor", 0.5, 1.5, float(rcpt_factor), 0.05)
dsb_factor = st.sidebar.slider("Disbursements Factor", 0.5, 1.5, float(dsb_factor), 0.05)

receipts_df = coerce_receipts_df(receipts_raw)
disb_df = coerce_disb_df(disb_raw)

fc = build_forecast(receipts_df, disb_df, opening_cash, start_date, rcpt_factor, dsb_factor)

st.subheader("Forecast Table")
st.dataframe(fc.style.format("{:,.1f}"))

st.subheader("Charts")
st.line_chart(fc["Ending Cash"])
st.bar_chart(fc[["Receipts", "Disbursements"]])

st.download_button(
    "⬇️ Download Forecast (CSV)",
    data=fc.to_csv().encode("utf-8"),
    file_name="13Week_CashFlow.csv",
    mime="text/csv"
)
