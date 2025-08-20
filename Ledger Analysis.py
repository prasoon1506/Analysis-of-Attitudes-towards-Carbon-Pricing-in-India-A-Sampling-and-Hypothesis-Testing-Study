# app.py — Minimal, reliable Streamlit ledger report (PPC-first)
# pip install streamlit pdfplumber pandas numpy

import io
import re
from datetime import datetime
from typing import List

import numpy as np
import pandas as pd
import pdfplumber
import streamlit as st

# ------------------------ Config ------------------------
st.set_page_config(page_title="Cement Ledger — Simple PPC Report", layout="wide")

# ------------------------ Helpers ------------------------
def _to_float(s: str):
    if s is None:
        return None
    s = s.replace(",", "").strip()
    try:
        return float(s)
    except:
        return None

def _dt(s: str):
    try:
        return datetime.strptime(s, "%d.%m.%Y")
    except:
        return None

def _norm_product(raw: str) -> str:
    """Normalize cement name just enough for grouping. PPC by default focus."""
    t = (raw or "").upper()
    if "PPC" in t:
        return "PPC"
    if "43" in t and "GRADE" in t:
        return "43 GRADE"
    if "SUPERSTRONG" in t or "ADSTAR" in t:
        return "SUPERSTRONG ADSTAR"
    # fallback: first token
    t = re.sub(r"\s+", " ", t).strip()
    return t.split(" ")[0] if t else "OTHER"

# Strict, simple regex for "Sales of-..." lines.
# We do NOT try to infer columns heuristically — just capture what we need.
SALES_RE = re.compile(
    r"(?P<date>\d{2}\.\d{2}\.\d{4}).*?Sales of-(?P<ctype>[A-Za-z0-9 /+\-&]+?)\s+"
    r"(?P<qty>\d+(?:\.\d+)?)\s+"
    r"(?P<rate>\d{1,3}(?:,\d{3})*(?:\.\d+)?)\s+"
    r"(?P<debit>\d{1,3}(?:,\d{3})*(?:\.\d{2}))"
    r"(?:\s+(?P<credit>\d{1,3}(?:,\d{3})*(?:\.\d{2})))?"
)

# Credit note lines: include /DG/ or RQDBN or "CREDIT NOTE"
# Exclude bank/deposits/adjustments: PIF, COLL, BANK, NEFT, RTGS, UPI, CASH, DZ, ADJUST
CREDIT_INCLUDE = re.compile(r"(?:/DG/|RQDBN|CREDIT NOTE)", re.IGNORECASE)
CREDIT_EXCLUDE = re.compile(r"(PIF|COLL|BANK|NEFT|RTGS|UPI|CASH|/DZ/|ADJUST)", re.IGNORECASE)
DATE_RE = re.compile(r"^\d{2}\.\d{2}\.\d{4}")

AMOUNT_AT_END_RE = re.compile(r"(\d{1,3}(?:,\d{3})*(?:\.\d{2}))\s*$")

def extract_lines_from_pdf(file_bytes: bytes) -> List[str]:
    lines: List[str] = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text(x_tolerance=1) or ""
            for ln in txt.splitlines():
                lines.append(" ".join(ln.split()))
    return lines

def parse_sales(lines: List[str]) -> pd.DataFrame:
    rows = []
    for ln in lines:
        m = SALES_RE.search(ln)
        if not m:
            continue
        d = _dt(m.group("date"))
        if not d:
            continue
        ctype = _norm_product(m.group("ctype"))
        qty = _to_float(m.group("qty"))
        rate = _to_float(m.group("rate"))  # not used for avg, but kept for reference
        debit = _to_float(m.group("debit"))
        credit = _to_float(m.group("credit")) if m.group("credit") else None
        rows.append(dict(date=d, product=ctype, qty_mt=qty, rate_bag=rate, debit=debit, credit=credit))
    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values("date").reset_index(drop=True)
    return df

def parse_credit_notes(lines: List[str]) -> pd.DataFrame:
    """Return date + credit amount for credit notes only (overall, not per product)."""
    rows = []
    for ln in lines:
        # must start with a date
        if not DATE_RE.match(ln):
            continue
        if CREDIT_EXCLUDE.search(ln):
            continue
        if not CREDIT_INCLUDE.search(ln):
            continue
        # grab last currency number on the line = credit column
        m_amt = AMOUNT_AT_END_RE.search(ln)
        if not m_amt:
            continue
        amt = _to_float(m_amt.group(1))
        d = _dt(ln[:10])
        if d and amt is not None:
            rows.append(dict(date=d, credit=amt))
    return pd.DataFrame(rows)

def weighted_price_per_bag(total_debit: float, total_qty_mt: float) -> float:
    """Price per bag = total_debit / (total_qty_mt * 20)."""
    if not total_qty_mt or total_qty_mt == 0:
        return np.nan
    return total_debit / (total_qty_mt * 20.0)

def monthly_ppc_table(df_sales: pd.DataFrame, df_credit: pd.DataFrame, include_gst: bool, gst_pct: float) -> pd.DataFrame:
    # Focus PPC
    ppc = df_sales[df_sales["product"] == "PPC"].copy()
    if ppc.empty:
        return pd.DataFrame(columns=["Year/Month", "QTY(MT) [PPC]", "Price/Bag", "Credit Note (Overall)"])

    ppc["ym"] = ppc["date"].dt.to_period("M").astype(str)
    # total debit, total qty per month => weighted price
    m = ppc.groupby("ym", as_index=False).agg(
        qty_mt=("qty_mt", "sum"),
        debit=("debit", "sum"),
    )
    m["price_bag"] = m.apply(lambda r: weighted_price_per_bag(r["debit"], r["qty_mt"]), axis=1)

    # Credit notes overall (independent of product)
    if df_credit.empty:
        cn_m = pd.DataFrame({"ym": [], "credit_note_overall": []})
    else:
        df_credit["ym"] = df_credit["date"].dt.to_period("M").astype(str)
        cn_m = df_credit.groupby("ym", as_index=False).agg(credit_note_overall=("credit", "sum"))

    out = m.merge(cn_m, on="ym", how="left").fillna({"credit_note_overall": 0.0})

    # GST option on price/bag only
    if include_gst and gst_pct > 0:
        out["price_bag"] = out["price_bag"] * (1.0 + gst_pct / 100.0)

    out = out.rename(columns={
        "ym": "Year/Month",
        "qty_mt": "QTY(MT) [PPC]",
        "price_bag": "Price/Bag",
        "credit_note_overall": "Credit Note (Overall)",
    }).sort_values("Year/Month")

    # Grand total row (weighted price across all months)
    total_qty = out["QTY(MT) [PPC]"].sum()
    total_debit = m["debit"].sum()  # use pre-GST debit for weighting
    overall_price = weighted_price_per_bag(total_debit, total_qty)
    if include_gst and gst_pct > 0:
        overall_price = overall_price * (1.0 + gst_pct / 100.0)

    grand = pd.DataFrame([{
        "Year/Month": "Grand Total",
        "QTY(MT) [PPC]": total_qty,
        "Price/Bag": overall_price,
        "Credit Note (Overall)": out["Credit Note (Overall)"].sum()
    }])

    return pd.concat([out, grand], ignore_index=True)

# ------------------------ UI ------------------------
st.title("🧱 Simple PPC Monthly Report")

files = st.file_uploader("Upload ledger PDFs", type=["pdf"], accept_multiple_files=True)

col = st.columns(3)
with col[0]:
    cement_choice = st.selectbox("Product", options=["PPC (primary)", "All sales table"], index=0)
with col[1]:
    include_gst = st.checkbox("Include GST in Price/Bag?", value=False)
with col[2]:
    gst_pct = st.number_input("GST %", value=28.0, min_value=0.0, max_value=50.0, step=0.25, disabled=not include_gst)

if not files:
    st.info("Upload one or more ledgers to begin.")
    st.stop()

# Parse all PDFs (simple, robust)
all_sales = []
all_credit = []
for f in files:
    lines = extract_lines_from_pdf(f.read())
    all_sales.append(parse_sales(lines))
    all_credit.append(parse_credit_notes(lines))

df_sales = pd.concat(all_sales, ignore_index=True) if all_sales else pd.DataFrame()
df_credit = pd.concat(all_credit, ignore_index=True) if all_credit else pd.DataFrame()

if df_sales.empty and df_credit.empty:
    st.error("No recognizable rows found. If your ledger layout differs, please share a sample line.")
    st.stop()

# --- PPC report
if cement_choice.startswith("PPC"):
    report = monthly_ppc_table(df_sales, df_credit, include_gst, gst_pct)
    st.subheader("Monthly Report — PPC")
    if report.empty:
        st.warning("No PPC sales found.")
    else:
        st.dataframe(
            report.style.format({
                "QTY(MT) [PPC]": "{:,.2f}",
                "Price/Bag": "₹ {:,.2f}",
                "Credit Note (Overall)": "₹ {:,.0f}",
            }),
            use_container_width=True
        )
        st.download_button(
            "Download CSV",
            report.to_csv(index=False).encode("utf-8"),
            "ppc_monthly_report.csv",
            "text/csv",
        )

# --- Optional: simple raw table to verify numbers if needed
if cement_choice.endswith("table"):
    st.subheader("All Sales (raw)")
    if df_sales.empty:
        st.info("No sales parsed.")
    else:
        # Show only essential columns so you can quickly validate qty & debit used for price
        raw = df_sales[["date", "product", "qty_mt", "debit", "rate_bag"]].copy()
        raw = raw.sort_values(["date", "product"])
        st.dataframe(raw, use_container_width=True, height=420)
        st.caption("Price/Bag for reports is computed as total_debit / (total_qty * 20), not the simple mean of rate.")
