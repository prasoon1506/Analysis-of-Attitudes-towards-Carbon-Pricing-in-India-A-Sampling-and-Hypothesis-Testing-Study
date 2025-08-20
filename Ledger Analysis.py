# app.py — PPC monthly report (supports single/multiple months per ledger file)
# pip install streamlit pdfplumber pandas numpy

import io, re
from datetime import datetime
from typing import List

import numpy as np
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(page_title="Cement Ledger — PPC Monthly Report", layout="wide")

# ---------------- patterns ----------------
NUM_RE = re.compile(r'(?<![A-Za-z])[+-]?\d{1,3}(?:,\d{3})*(?:\.\d+)?')
DATE_AT_START = re.compile(r'^\d{2}\.\d{2}\.\d{4}')
CREDIT_INCLUDE = re.compile(r'(?:/DG/|RQDBN|CREDIT\s*NOTE)', re.IGNORECASE)
CREDIT_EXCLUDE = re.compile(r'(?:PIF|COLL|BANK|NEFT|RTGS|UPI|CASH|/DZ/|ADJUST)', re.IGNORECASE)
SALES_RE = re.compile(
    r'(?P<date>\d{2}\.\d{2}\.\d{4}).*?Sales of-(?P<ctype>[A-Za-z0-9 /+\-&]+?)\s+'
    r'(?P<qty>\d+(?:\.\d+)?)\s+'
    r'(?P<rate>\d{1,3}(?:,\d{3})*(?:\.\d+)?)\s+'
    r'(?P<debit>\d{1,3}(?:,\d{3})*(?:\.\d{2}))'
    r'(?:\s+(?P<credit>\d{1,3}(?:,\d{3})*(?:\.\d{2})))?'
)

def _to_float(s: str):
    if not s: return None
    try: return float(s.replace(",", "").strip())
    except: return None

def _dt(s: str):
    try: return datetime.strptime(s, "%d.%m.%Y")
    except: return None

def _norm_product(raw: str) -> str:
    t = (raw or "").upper()
    if "PPC" in t: return "PPC"
    if "43" in t and "GRADE" in t: return "43 GRADE"
    if "SUPERSTRONG" in t or "ADSTAR" in t: return "SUPERSTRONG ADSTAR"
    return "OTHER"

def extract_lines(file_bytes: bytes) -> List[str]:
    lines=[]
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for p in pdf.pages:
            txt=p.extract_text(x_tolerance=1) or ""
            for ln in txt.splitlines():
                lines.append(" ".join(ln.split()))
    return lines

def parse_sales(lines: List[str]) -> pd.DataFrame:
    rows=[]
    for ln in lines:
        m=SALES_RE.search(ln)
        if not m: continue
        dt=_dt(m.group("date")); 
        if not dt: continue
        rows.append(dict(
            date=dt,
            product=_norm_product(m.group("ctype")),
            qty_mt=_to_float(m.group("qty")),
            rate_bag=_to_float(m.group("rate")),
            debit=_to_float(m.group("debit")),
            credit=_to_float(m.group("credit")) if m.group("credit") else None
        ))
    df=pd.DataFrame(rows)
    if not df.empty: df=df.sort_values("date").reset_index(drop=True)
    return df

def parse_credit_notes(lines: List[str]) -> pd.DataFrame:
    rows=[]
    for ln in lines:
        if not DATE_AT_START.match(ln): continue
        if CREDIT_EXCLUDE.search(ln): continue
        if not CREDIT_INCLUDE.search(ln): continue
        nums=NUM_RE.findall(ln)
        if len(nums)<2: continue
        floats=[_to_float(x) for x in nums if _to_float(x) is not None]
        if len(floats)<2: continue
        amt=floats[-2]   # credit column
        dt=_dt(ln[:10])
        if dt and amt is not None:
            rows.append(dict(date=dt, credit=amt))
    return pd.DataFrame(rows)

def weighted_price(debit, qty_mt): 
    return debit/(qty_mt*20.0) if qty_mt else np.nan

def monthly_ppc(df_sales, df_credit, include_gst, gst_pct):
    ppc=df_sales[df_sales["product"]=="PPC"].copy()
    if ppc.empty: return pd.DataFrame()
    ppc["ym"]=ppc["date"].dt.to_period("M").astype(str)
    agg=ppc.groupby("ym",as_index=False).agg(qty=("qty_mt","sum"), debit=("debit","sum"))
    agg["price_bag"]=agg.apply(lambda r: weighted_price(r["debit"],r["qty"]), axis=1)
    if include_gst: agg["price_bag"]*= (1+gst_pct/100.0)
    if not df_credit.empty:
        df_credit=df_credit.copy()
        df_credit["ym"]=df_credit["date"].dt.to_period("M").astype(str)
        cn=df_credit.groupby("ym",as_index=False).agg(credit_note=("credit","sum"))
    else:
        cn=pd.DataFrame({"ym":[],"credit_note":[]})
    out=agg.merge(cn,on="ym",how="left").fillna({"credit_note":0.0})
    out=out.rename(columns={"ym":"Year/Month","qty":"QTY(MT) [PPC]","price_bag":"Price/Bag","credit_note":"Credit Note (Overall)"})
    # Grand total
    total_qty=out["QTY(MT) [PPC]"].sum(); total_debit=agg["debit"].sum()
    price=weighted_price(total_debit,total_qty)
    if include_gst: price*=(1+gst_pct/100.0)
    grand=pd.DataFrame([{"Year/Month":"Grand Total","QTY(MT) [PPC]":total_qty,"Price/Bag":price,"Credit Note (Overall)":out["Credit Note (Overall)"].sum()}])
    return pd.concat([out,grand],ignore_index=True)

# ---------------- UI ----------------
st.title("🧱 PPC Monthly Report (Multi-month aware)")

files=st.file_uploader("Upload one or more ledger PDFs (single or multi-month)", type="pdf", accept_multiple_files=True)
c1,c2=st.columns(2)
with c1: include_gst=st.checkbox("Include GST?", value=False)
with c2: gst_pct=st.number_input("GST %", value=28.0, step=0.25, disabled=not include_gst)

if files:
    all_sales, all_cn=[],[]
    for f in files:
        lines=extract_lines(f.read())
        all_sales.append(parse_sales(lines))
        all_cn.append(parse_credit_notes(lines))
    df_sales=pd.concat(all_sales,ignore_index=True) if all_sales else pd.DataFrame()
    df_credit=pd.concat(all_cn,ignore_index=True) if all_cn else pd.DataFrame()
    if df_sales.empty: st.error("No 'Sales of-' rows found."); st.stop()
    report=monthly_ppc(df_sales, df_credit, include_gst, gst_pct)
    st.subheader("Monthly Report — PPC (auto-splits multi-month ledgers)")
    st.dataframe(report.style.format({"QTY(MT) [PPC]":"{:,.2f}","Price/Bag":"₹{:,.2f}","Credit Note (Overall)":"₹{:,.0f}"}),use_container_width=True)
    st.download_button("Download CSV", report.to_csv(index=False).encode("utf-8"), "ppc_monthly_report.csv","text/csv")
