
import io
import datetime as dt
import pandas as pd
import numpy as np
import streamlit as st

st.set_page_config(page_title="13-Week Cash Flow — Excel App", layout="wide")
st.title("13-Week Cash Flow Forecaster — Excel App")

st.write("Upload a workbook with sheets: **Inputs**, **Receipts**, **Disbursements**. "
         "Then review/override inputs, see KPIs & charts, and export a final Excel.")

uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

# @st.cache_data
# def load_from_excel(file_bytes: bytes):
#     xl = pd.ExcelFile(io.BytesIO(file_bytes))
#     def safe_read(sheet_name):
#         try:
#             return xl.parse(sheet_name)
#         except Exception:
#             return None
#     return safe_read("Inputs"), safe_read("Receipts"), safe_read("Disbursements")
@st.cache_data
def load_from_excel(file_bytes: bytes):
    # Force the openpyxl engine for .xlsx files
    xl = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    def safe_read(sheet_name: str):
        try:
            return xl.parse(sheet_name)
        except Exception:
            return None
    return safe_read("Inputs"), safe_read("Receipts"), safe_read("Disbursements")

def coerce_receipts_df(df):
    out = pd.DataFrame()
    out["Date"] = pd.to_datetime(df.get("Date", df.iloc[:,0]), errors="coerce")
    out["Source"] = df.get("Customer/Source", df.get("Source", df.iloc[:,1])).astype(str)
    out["Category"] = df.get("Category", "Sales Receipts").astype(str)
    amt_col = next((c for c in df.columns if "Amount" in c and "Expected" not in c), df.columns[3])
    out["Amount"] = pd.to_numeric(df[amt_col], errors="coerce").fillna(0.0)
    prob_col = next((c for c in df.columns if "Probability" in c), None)
    out["Probability"] = pd.to_numeric(df.get(prob_col, 1.0), errors="coerce").fillna(1.0)
    return out.dropna(subset=["Date"])

def coerce_disb_df(df):
    out = pd.DataFrame()
    out["Date"] = pd.to_datetime(df.get("Date", df.iloc[:,0]), errors="coerce")
    out["Payee"] = df.get("Vendor/Payee", df.get("Payee", df.iloc[:,1])).astype(str)
    out["Category"] = df.get("Category", "Vendors").astype(str)
    amt_col = next((c for c in df.columns if "Amount" in c and "Expected" not in c), df.columns[3])
    out["Amount"] = pd.to_numeric(df[amt_col], errors="coerce").fillna(0.0)
    prob_col = next((c for c in df.columns if "Probability" in c), None)
    out["Probability"] = pd.to_numeric(df.get(prob_col, 1.0), errors="coerce").fillna(1.0)
    return out.dropna(subset=["Date"])

def parse_inputs(df):
    open_cash = 0.0; min_cash = 0.0; rcpt_factor = 1.0; dsb_factor = 1.0
    start_date = dt.date.today() - dt.timedelta(days=dt.date.today().weekday())
    try:
        mapping = {str(df.iloc[i,0]).strip().lower(): df.iloc[i,1] for i in range(min(len(df), 60))}
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
        open_cash = find_num("opening cash", open_cash)
        min_cash = find_num("min cash", min_cash)
        rcpt_factor = find_num("receipts factor", rcpt_factor)
        dsb_factor = find_num("disbursements factor", dsb_factor)
        start_date = find_date("start date", start_date)
    except Exception:
        pass
    return open_cash, start_date, min_cash, rcpt_factor, dsb_factor

def weekly_buckets(start_date):
    starts = [pd.to_datetime(start_date) + pd.Timedelta(days=7*i) for i in range(13)]
    ends = [d + pd.Timedelta(days=6) for d in starts]
    labels = [f"Week {i+1}\n{starts[i].date()}" for i in range(13)]
    return starts, ends, labels

def build_forecast(receipts, disb, opening_cash, min_cash, start_date, rcpt_factor, dsb_factor, currency_unit):
    receipts = receipts.copy(); disb = disb.copy()
    receipts["Expected"] = receipts["Amount"] * receipts["Probability"] * rcpt_factor
    disb["Expected"] = disb["Amount"] * disb["Probability"] * dsb_factor

    ws, we, labels = weekly_buckets(start_date)

    def sum_week(df, cat_col=None, category=None):
        vals = []
        for i in range(13):
            sub = df[(df["Date"] >= ws[i]) & (df["Date"] <= we[i])]
            if cat_col and category is not None:
                sub = sub[sub[cat_col] == category]
            vals.append(float(sub["Expected"].sum()))
        return np.array(vals)

    rcpt_cats = sorted(receipts["Category"].dropna().unique().tolist())
    dsb_cats = sorted(disb["Category"].dropna().unique().tolist())

    fc = pd.DataFrame(index=[
        "Opening Cash",
        "Receipts",
    ] + [f"  {c}" for c in rcpt_cats] + [
        "Disbursements",
    ] + [f"  {c}" for c in dsb_cats] + [
        "Net Cash Flow",
        "Ending Cash",
    ], columns=labels, dtype=float)

    rcpt_tot = np.zeros(13)
    for c in rcpt_cats:
        vals = sum_week(receipts, "Category", c)
        fc.loc[f"  {c}"] = vals; rcpt_tot += vals
    fc.loc["Receipts"] = rcpt_tot

    dsb_tot = np.zeros(13)
    for c in dsb_cats:
        vals = sum_week(disb, "Category", c)
        fc.loc[f"  {c}"] = vals; dsb_tot += vals
    fc.loc["Disbursements"] = dsb_tot

    net = rcpt_tot - dsb_tot
    fc.loc["Net Cash Flow"] = net

    ending = np.zeros(13)
    for i in range(13):
        open_i = opening_cash if i == 0 else ending[i-1]
        fc.loc["Opening Cash", labels[i]] = open_i
        ending[i] = open_i + net[i]
    fc.loc["Ending Cash"] = ending

    return fc, ws, we

def export_excel(fc, opening_cash, start_date, min_cash, rcpt_factor, dsb_factor, currency_unit):
    import xlsxwriter
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
    wsI.write("A4", "Opening Cash"); wsI.write_number("B4", float(fc.iloc[0,0]), fmt_num)
    wsI.write("A5", "Start Date"); wsI.write_datetime("B5", pd.to_datetime(start_date).to_pydatetime(), fmt_date)
    wsI.write("A6", "Min Cash Threshold"); wsI.write_number("B6", float(min_cash), fmt_num)
    wsI.write("A8", "Receipts Factor"); wsI.write_number("B8", float(rcpt_factor), fmt_num)
    wsI.write("A9", "Disbursements Factor"); wsI.write_number("B9", float(dsb_factor), fmt_num)

    wsF = wb.add_worksheet("Forecast")
    wsF.write("A1", "13-Week Cash Flow Forecast (Direct Method)", fmt_title)
    wsF.write("A3", "Metric", fmt_hdr)
    for i, col in enumerate(fc.columns, start=2):
        wsF.write(2, i-1, col, fmt_hdr)
    for r, idx in enumerate(fc.index, start=4):
        wsF.write(r-1, 0, idx)
        for i, col in enumerate(fc.columns, start=2):
            wsF.write_number(r-1, i-1, float(fc.loc[idx, col]), fmt_num)

    end_row = 4 + list(fc.index).index("Ending Cash")
    wsF.conditional_format(end_row-1, 1, end_row-1, 13, {
        "type": "cell", "criteria": "<", "value": float(min_cash), "format": fmt_red
    })

    wsD = wb.add_worksheet("Dashboard")
    wsD.write("A1", "Dashboard", fmt_title)
    wsD.write("A3", "Key KPIs", fmt_h2)
    wsD.write("A5", "Opening Cash"); wsD.write_number("B5", float(fc.iloc[0,0]), fmt_num)
    wsD.write("A6", "Min Ending Cash"); wsD.write_number("B6", float(fc.loc["Ending Cash"].min()), fmt_num)
    wsD.write("A7", "Total Receipts"); wsD.write_number("B7", float(fc.loc["Receipts"].sum()), fmt_num)
    wsD.write("A8", "Total Disbursements"); wsD.write_number("B8", float(fc.loc["Disbursements"].sum()), fmt_num)

    chart1 = wb.add_chart({"type": "line"})
    chart1.add_series({
        "name": "Ending Cash",
        "categories": ["Forecast", 2, 1, 2, 13],
        "values":     ["Forecast", end_row-1, 1, end_row-1, 13],
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
        "categories": ["Forecast", 2, 1, 2, 13],
        "values":     ["Forecast", rc_row-1, 1, rc_row-1, 13],
    })
    chart2.add_series({
        "name": "Disbursements",
        "categories": ["Forecast", 2, 1, 2, 13],
        "values":     ["Forecast", ds_row-1, 1, ds_row-1, 13],
    })
    chart2.set_title({"name": "Receipts vs Disbursements (Weekly)"})
    chart2.set_x_axis({"name": "Week"})
    chart2.set_y_axis({"name": currency_unit})
    wsD.insert_chart("H12", chart2, {"x_scale": 1.5, "y_scale": 1.2})

    wb.close()
    bio.seek(0)
    return bio.getvalue()

# Sidebar inputs
st.sidebar.header("Inputs")
currency_unit = st.sidebar.text_input("Currency / Unit label", "INR lakh")

# Defaults
opening_cash = 0.0
start_date = dt.date.today() - dt.timedelta(days=dt.date.today().weekday())
min_cash = 0.0
rcpt_factor = 1.0
dsb_factor = 1.0

receipts_df = None; disb_df = None

if uploaded:
    inputs_df, receipts_raw, disb_raw = load_from_excel(uploaded.read())
    if inputs_df is not None and len(inputs_df) > 0:
        oc, sd, mc, rf, dfac = parse_inputs(inputs_df)
        opening_cash = oc; start_date = sd; min_cash = mc; rcpt_factor = rf; dsb_factor = dfac
    if receipts_raw is not None and len(receipts_raw) > 0:
        receipts_df = coerce_receipts_df(receipts_raw)
    if disb_raw is not None and len(disb_raw) > 0:
        disb_df = coerce_disb_df(disb_raw)

opening_cash = st.sidebar.number_input("Opening Cash", value=float(opening_cash or 0.0), step=10.0, format="%.1f")
start_date = st.sidebar.date_input("Week 1 start date (Monday)", value=start_date)
min_cash = st.sidebar.number_input("Min Cash Threshold", value=float(min_cash or 0.0), step=10.0, format="%.1f")

st.sidebar.header("Scenario")
rcpt_factor = st.sidebar.slider("Receipts Factor", 0.5, 1.5, float(rcpt_factor or 1.0), 0.05)
dsb_factor = st.sidebar.slider("Disbursements Factor", 0.5, 1.5, float(dsb_factor or 1.0), 0.05)

if receipts_df is None or disb_df is None:
    st.info("Upload an Excel to continue. Expect sheets named Inputs, Receipts, Disbursements.")
    st.stop()

st.subheader("Receipts (preview)")
st.dataframe(receipts_df.head(20))

st.subheader("Disbursements (preview)")
st.dataframe(disb_df.head(20))

fc, _, _ = build_forecast(receipts_df, disb_df, opening_cash, min_cash, start_date, rcpt_factor, dsb_factor, currency_unit)

c1, c2, c3, c4 = st.columns(4)
c1.metric("Opening Cash", f"{fc.iloc[0,0]:,.1f} {currency_unit}")
c2.metric("Min Ending Cash", f"{fc.loc['Ending Cash'].min():,.1f} {currency_unit}")
c3.metric("Total Receipts", f"{fc.loc['Receipts'].sum():,.1f} {currency_unit}")
c4.metric("Total Disbursements", f"{fc.loc['Disbursements'].sum():,.1f} {currency_unit}")

import matplotlib.pyplot as plt

fig1, ax1 = plt.subplots()
ax1.plot(range(1,14), fc.loc["Ending Cash"].values, marker="o")
ax1.set_title("Ending Cash (13 Weeks)")
ax1.set_xlabel("Week")
ax1.set_ylabel(currency_unit)
st.pyplot(fig1)

fig2, ax2 = plt.subplots()
idx = np.arange(1,14)
ax2.bar(idx-0.15, fc.loc["Receipts"].values, width=0.3, label="Receipts")
ax2.bar(idx+0.15, fc.loc["Disbursements"].values, width=0.3, label="Disbursements")
ax2.set_title("Receipts vs Disbursements (Weekly)")
ax2.set_xlabel("Week")
ax2.set_ylabel(currency_unit)
ax2.legend()
st.pyplot(fig2)

st.subheader("Forecast Table")
st.dataframe(fc.style.format("{:,.1f}"))

xlsx_bytes = export_excel(fc, opening_cash, start_date, min_cash, rcpt_factor, dsb_factor, currency_unit)
st.download_button("⬇️ Download Final Excel", data=xlsx_bytes, file_name="Final_13Week_CashFlow.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
