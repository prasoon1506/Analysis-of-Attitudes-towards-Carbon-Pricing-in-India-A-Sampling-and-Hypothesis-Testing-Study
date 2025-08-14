import re
from datetime import datetime
from collections import defaultdict
import io

import pdfplumber
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import streamlit as st

# ---------------------------
# PDF Extraction
# ---------------------------
def extract_lines_from_pdf(file):
    """Extract all lines of text from a PDF file object."""
    lines = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for ln in text.split("\n"):
                lines.append(re.sub(r"\s+", " ", ln).strip())
    return lines

# ---------------------------
# Ledger Parsing
# ---------------------------
LINE_PATTERN = re.compile(
    r"(?P<date>\d{2}\.\d{2}\.\d{4}).*?Sales of-(?P<ctype>[A-Z0-9\s\+\-/&]+?)\s+"
    r"(?P<qty>\d+(?:\.\d+)?)\s+"
    r"(?P<rate>\d{1,3}(?:,\d{3})*(?:\.\d+)?)\s+"
    r"(?P<debit>\d{1,3}(?:,\d{3})*(?:\.\d{2}))"
    r"(?:\s+(?P<credit>\d{1,3}(?:,\d{3})*(?:\.\d{2})))?"
)

def parse_sales(lines):
    records = []
    for line in lines:
        m = LINE_PATTERN.search(line)
        if m:
            date_obj = datetime.strptime(m.group("date"), "%d.%m.%Y")
            cement_type = m.group("ctype").strip()
            qty = float(m.group("qty"))
            rate = float(m.group("rate").replace(",", ""))
            debit = float(m.group("debit").replace(",", ""))
            credit = float(m.group("credit").replace(",", "")) if m.group("credit") else 0.0
            records.append({
                "date": date_obj,
                "cement_type": cement_type,
                "qty": qty,
                "rate": rate,
                "debit": debit,
                "credit": credit
            })
    df = pd.DataFrame(records)
    if not df.empty:
        df.sort_values(["cement_type", "date"], inplace=True)
    return df

# ---------------------------
# Computations
# ---------------------------
def compute_summary(df):
    grp = df.groupby("cement_type", as_index=False).agg(
        total_qty=("qty", "sum"),
        total_debit=("debit", "sum"),
        total_credit=("credit", "sum"),
    )
    grp["avg_invoice_price"] = (grp["total_debit"] / grp["total_qty"]) / 20.0

    change_records = []
    change_counts = defaultdict(int)
    for ctype, sub in df.groupby("cement_type"):
        sub = sub.sort_values("date")
        last_rate = None
        for _, row in sub.iterrows():
            if last_rate is None or row["rate"] != last_rate:
                change_records.append({"cement_type": ctype, "change_date": row["date"].date(), "price": row["rate"]})
                change_counts[ctype] += 1
                last_rate = row["rate"]
    change_log_df = pd.DataFrame(change_records)

    summary_df = grp.merge(
        pd.DataFrame([{"cement_type": k, "price_changes": v} for k, v in change_counts.items()]),
        on="cement_type", how="left"
    )
    return summary_df, change_log_df

def qty_per_price(df):
    return df.groupby(["cement_type", "rate"], as_index=False).agg(qty_sold=("qty", "sum"))

def build_price_timeline(df):
    timelines = {}
    if df.empty:
        return timelines
    overall_start = df["date"].min().date()
    overall_end = df["date"].max().date()
    calendar = pd.date_range(overall_start, overall_end, freq="D")
    for ctype, sub in df.groupby("cement_type"):
        last_of_day = sub.sort_values(["date"]).groupby(sub["date"].dt.date).tail(1)
        s = last_of_day.set_index(last_of_day["date"].dt.date)["rate"]
        s = s.reindex(calendar.date, method=None).ffill()
        timelines[ctype] = pd.DataFrame({"rate": s.values}, index=calendar)
    return timelines

def get_prices_on_date(timelines, selected_date):
    rows = []
    for ctype, tl in timelines.items():
        price = float(tl.loc[pd.to_datetime(selected_date), "rate"])
        rows.append({"cement_type": ctype, "rate_on_date": price})
    return pd.DataFrame(rows)

# ---------------------------
# Streamlit UI
# ---------------------------
st.set_page_config(page_title="Cement Ledger Dashboard", layout="wide")
st.title("🏗 Cement Ledger Dashboard")

uploaded_file = st.file_uploader("Upload Ledger PDF", type=["pdf"])
if uploaded_file:
    lines = extract_lines_from_pdf(uploaded_file)
    df = parse_sales(lines)

    if df.empty:
        st.error("No sales data found in the uploaded PDF.")
    else:
        summary_df, change_log_df = compute_summary(df)
        qty_price_df = qty_per_price(df)
        timelines = build_price_timeline(df)

        # Tabs for navigation
        tab1, tab2, tab3, tab4 = st.tabs(["📊 Summary", "📈 Price Analysis", "📅 Price Lookup", "📜 Change Log"])

        with tab1:
            st.subheader("Overall Summary")
            for _, row in summary_df.iterrows():
                st.metric(label=f"{row['cement_type']} - Avg Price", value=f"{row['avg_invoice_price']:.2f}")
            st.dataframe(summary_df)

        with tab2:
            st.subheader("Quantity Sold per Price")
            cement_types = summary_df["cement_type"].unique()
            for c in cement_types:
                sub = qty_price_df[qty_price_df["cement_type"] == c]
                fig, ax = plt.subplots()
                ax.bar(sub["rate"].astype(str), sub["qty_sold"])
                ax.set_title(f"{c} - Qty Sold per Price")
                ax.set_ylabel("Qty (MT)")
                ax.set_xlabel("Price (per bag)")
                st.pyplot(fig)

            st.subheader("Daily Price Timeline")
            for c, tl in timelines.items():
                fig, ax = plt.subplots()
                ax.plot(tl.index, tl["rate"])
                ax.set_title(f"{c} - Price Timeline")
                ax.set_ylabel("Price (per bag)")
                ax.set_xlabel("Date")
                st.pyplot(fig)

        with tab3:
            st.subheader("Price Lookup by Date")
            all_dates = sorted(df["date"].dt.date.unique())
            selected_date = st.date_input("Select date", min_value=min(all_dates), max_value=max(all_dates))
            prices_df = get_prices_on_date(timelines, selected_date)
            st.dataframe(prices_df)

        with tab4:
            st.subheader("Price Change Log")
            st.dataframe(change_log_df)

        # Download buttons
        st.download_button(
            "Download Full Ledger CSV",
            df.to_csv(index=False).encode("utf-8"),
            "ledger_full.csv",
            "text/csv"
        )
        st.download_button(
            "Download Summary CSV",
            summary_df.to_csv(index=False).encode("utf-8"),
            "ledger_summary.csv",
            "text/csv"
        )
