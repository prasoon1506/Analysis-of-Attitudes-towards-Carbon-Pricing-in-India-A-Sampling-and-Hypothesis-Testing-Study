import re
from datetime import datetime
import pdfplumber
import pandas as pd
import numpy as np
import streamlit as st

# ---------------------------
# PDF Extraction
# ---------------------------
def extract_lines_from_pdf(file):
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
        df.sort_values("date", inplace=True)
    return df

# ---------------------------
# Report Computation
# ---------------------------
def generate_monthly_report(df, cement_type_filter=None):
    if cement_type_filter:
        df = df[df["cement_type"] == cement_type_filter]

    if df.empty:
        return pd.DataFrame(), {}

    df["Year"] = df["date"].dt.year
    df["Month"] = df["date"].dt.strftime("%b")
    df["YearMonth"] = df["date"].dt.to_period("M")

    monthly = df.groupby(["Year", "Month"], as_index=False).agg(
        Qty_MT=("qty", "sum"),
        PriceBag=("rate", "mean"),
        Discount=("credit", "sum")
    )

    # Add year subtotals
    subtotals = monthly.groupby("Year").agg(
        Qty_MT=("Qty_MT", "sum"),
        PriceBag=("PriceBag", "mean"),
        Discount=("Discount", "sum")
    ).reset_index()

    subtotals["Month"] = "TOTAL"
    report = pd.concat([monthly, subtotals], ignore_index=True)

    # Add Grand Total
    grand_total = pd.DataFrame([{
        "Year": "Grand",
        "Month": "Total",
        "Qty_MT": monthly["Qty_MT"].sum(),
        "PriceBag": monthly["PriceBag"].mean(),
        "Discount": monthly["Discount"].sum()
    }])
    report = pd.concat([report, grand_total], ignore_index=True)

    # Footer calculations
    total_qty = monthly["Qty_MT"].sum()
    total_discount = monthly["Discount"].sum()
    discount_per_bag = total_discount / (total_qty * 20) if total_qty else 0
    gst = 8.75
    nod = discount_per_bag + gst

    footer = {
        "TotalQty": total_qty,
        "DiscountPerBag": discount_per_bag,
        "GST": gst,
        "TotalDiscount": discount_per_bag + gst,
        "NOD": nod
    }

    return report, footer

# ---------------------------
# Streamlit UI
# ---------------------------
st.set_page_config(page_title="Cement Ledger Monthly Report", layout="wide")
st.title("🏗 Cement Ledger Monthly Report")

uploaded_file = st.file_uploader("Upload Ledger PDF", type=["pdf"])
if uploaded_file:
    lines = extract_lines_from_pdf(uploaded_file)
    df = parse_sales(lines)

    if df.empty:
        st.error("No sales data found in the uploaded PDF.")
    else:
        cement_types = df["cement_type"].unique().tolist()
        cement_type_filter = st.selectbox("Select Cement Type", ["All"] + cement_types)

        if cement_type_filter != "All":
            report, footer = generate_monthly_report(df, cement_type_filter)
            st.subheader(f"Report for {cement_type_filter}")
        else:
            report, footer = generate_monthly_report(df)
            st.subheader("Report for All Cement Types")

        st.dataframe(report.style.format({
            "Qty_MT": "{:,.2f}",
            "PriceBag": "{:,.2f}",
            "Discount": "{:,.0f}"
        }))

        st.markdown("---")
        st.subheader("📊 Footer Summary")
        st.write(f"**Total Qty (All products combined):** {footer['TotalQty']:.2f}")
        st.write(f"**Discount/Bag:** ₹ {footer['DiscountPerBag']:.2f}")
        st.write(f"**GST:** ₹ {footer['GST']:.2f}")
        st.write(f"**Total Discount:** ₹ {footer['TotalDiscount']:.2f}")
        st.write(f"**NOD:** ₹ {footer['NOD']:.2f}")

        # Download option
        st.download_button(
            "Download Report as CSV",
            report.to_csv(index=False).encode("utf-8"),
            "monthly_report.csv",
            "text/csv"
        )
