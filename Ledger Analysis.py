# app.py
# Streamlit Ledger Analyzer for JK Cement PDFs (PPC-first)
# --------------------------------------------------------
# pip install streamlit pdfplumber pandas numpy plotly python-dateutil

import io
import re
import math
import pdfplumber
import pandas as pd
import numpy as np
from datetime import datetime, date, timedelta
import plotly.express as px
import streamlit as st

# ----------------------------- UI THEME TWEAKS -----------------------------
st.set_page_config(page_title="Cement Ledger Analyzer", page_icon="🧱", layout="wide")
st.markdown("""
<style>
/* Clean, modern look */
.block-container {padding-top: 1rem; padding-bottom: 2rem; }
.reportview-container .main .block-container{max-width: 1400px;}
div[data-testid="stMetricValue"] { font-size: 2rem; }
.small-note { color:#7a7a7a; font-size:0.85rem; }
thead tr th { background:#0f172a !important; color:#fff !important; }
tbody tr:hover { background: #f6f8ff !important; }
.kpi-card { border-radius: 16px; padding:16px; background: #0f172a; color:#fff; }
h1,h2,h3 { margin-top: .2rem; }
</style>
""", unsafe_allow_html=True)

# ----------------------------- HELPERS -----------------------------
NUM_RE = re.compile(r'(?<![A-Za-z])[-]?\d{1,3}(?:,\d{3})*(?:\.\d+)?')

def to_float(x):
    if x is None or x == "":
        return None
    x = str(x).replace(",", "")
    try:
        return float(x)
    except:
        return None

def ddmmyyyy_to_date(s):
    try:
        return datetime.strptime(s, "%d.%m.%Y").date()
    except:
        return None

def classify_row(doc_col, particulars):
    """Return (entry_type, product) based on text clues."""
    doc_col = doc_col or ""
    particulars = particulars or ""
    text = (doc_col + " " + particulars).upper()

    # Identify product only for Sales
    product = None
    if "SALES OF-" in text:
        # Extract the token after 'SALES OF-'
        product = text.split("SALES OF-")[-1].strip()
        # Normalize categories
        if "PPC" in product:
            product = "PPC"
        elif "43 GRADE" in product:
            product = "43 GRADE"
        elif "SUPERSTRONG" in product or "ADSTAR" in product:
            product = "SUPERSTRONG ADSTAR"
        else:
            # keep first 22 chars as generic label
            product = product.split()[0]

    # Credit Note (overall) — detect /DG/ & "RQDBN" etc.
    if "/DG/" in text or "RQDBN" in text or "CREDIT NOTE" in text:
        return "CreditNote", None

    # Bank/Fund deposit/adjustments that must not count as credit notes
    bank_words = ["PIF", "COLL", "BANK", "NEFT", "RTGS", "UPI", "CASH", "DEPOSIT", "FUND", "ADJUST"]
    if any(w in text for w in bank_words) or "/DZ/" in text:
        return "Bank", None

    # Sales entries
    if "SALES OF-" in text:
        return "Sale", product

    if "OPENING BALANCE" in text:
        return "Opening", None

    return "Other", None

def parse_pdf(file_bytes, filename=""):
    """Extract ledger rows from a JK Cement ledger PDF into a clean DataFrame."""
    records = []
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        for page in pdf.pages:
            # read page text lines
            text = page.extract_text(x_tolerance=1) or ""
            for raw in text.splitlines():
                line = " ".join(raw.split())
                # only process lines that begin with a date like 01.06.2025 OR 'Opening Balance'
                if not (re.match(r'^\d{2}\.\d{2}\.\d{4}', line) or "Opening Balance" in line):
                    continue

                # Opening balance row
                if "Opening Balance" in line and not re.match(r'^\d{2}\.\d{2}\.\d{4}', line):
                    nums = NUM_RE.findall(line)
                    credit_amt = to_float(nums[-1]) if nums else None
                    records.append(dict(
                        source=filename, doc_date=None, doc_col="", inv_no="",
                        particulars="Opening Balance", qty_mt=None, rate_per_bag=None,
                        debit_amount=None, credit_amount=credit_amt, cumulative=None,
                        entry_type="Opening", product=None
                    ))
                    continue

                # Split columns roughly: date (first 10), then rest
                doc_date = ddmmyyyy_to_date(line[:10])
                rest = line[10:].strip()

                # Try to peel out the descriptive columns
                # There are many spaces; use two anchored chunks:
                #   [doc/plant col] [inv no] [particulars] [numbers...]
                # We will try by capturing numbers from the right.
                nums = NUM_RE.findall(line)
                nums_f = [to_float(n) for n in nums]
                cumulative = nums_f[-1] if nums_f else None

                # infer debit/credit/qty/rate looking from the end
                credit_amount = None
                debit_amount = None
                rate_per_bag = None
                qty_mt = None

                # how many numeric tokens belong to tail columns?
                # tail layout: [ ... qty, rate, debit?, credit?, cumulative ]
                tail = nums_f[-5:] if len(nums_f) >= 5 else nums_f[:]
                # align tail to the right
                tlen = len(tail)

                # map by position from the right
                # pos -1: cumulative
                # pos -2: credit or debit single
                # pos -3: debit (if present)
                # pos -4: rate
                # pos -5: qty
                if tlen >= 2:
                    maybe_credit = tail[-2]
                    credit_amount = maybe_credit
                if tlen >= 3:
                    debit_amount = tail[-3]
                if tlen >= 4:
                    rate_per_bag = tail[-4]
                if tlen >= 5:
                    qty_mt = tail[-5]

                # Extract non-numeric text (to classify)
                # Remove the trailing numeric block we just consumed
                # and pull the last two "text columns" (doc col + particulars + inv)
                # A pragmatic approach: split by two+ spaces and pick middle chunks
                chunks = re.split(r'\s{2,}', line)
                # chunks often look like [date, doccol, invno, particulars, qty, rate, debit, credit, cumulative]
                doc_col = ""
                inv_no = ""
                particulars = ""
                if len(chunks) >= 4:
                    doc_col = chunks[1]
                    inv_no = chunks[2]
                    particulars = chunks[3]
                elif len(chunks) >= 3:
                    doc_col = chunks[1]
                    particulars = chunks[2]
                else:
                    particulars = rest

                entry_type, product = classify_row(doc_col, particulars)

                # Disambiguate debit/credit based on entry type
                if entry_type == "Sale":
                    credit_amount = None
                elif entry_type in ("CreditNote", "Bank"):
                    debit_amount = None

                records.append(dict(
                    source=filename, doc_date=doc_date, doc_col=doc_col, inv_no=inv_no,
                    particulars=particulars, qty_mt=qty_mt, rate_per_bag=rate_per_bag,
                    debit_amount=debit_amount, credit_amount=credit_amount, cumulative=cumulative,
                    entry_type=entry_type, product=product
                ))

    df = pd.DataFrame.from_records(records)
    if df.empty:
        return df

    # Clean
    df["doc_date"] = pd.to_datetime(df["doc_date"])
    # treat negatives with trailing '-' if any crept in
    for c in ["qty_mt", "rate_per_bag", "debit_amount", "credit_amount", "cumulative"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    # Remove rows that somehow have both debit and credit populated (rare parsing artifact)
    both = df["debit_amount"].notna() & df["credit_amount"].notna()
    df.loc[both, ["credit_amount"]] = np.nan

    return df

def weighted_avg_price_per_bag(df_sales):
    """Average invoice price per bag = total debit / (qty*20)."""
    qty_bags = (df_sales["qty_mt"].fillna(0) * 20.0)
    total_debit = df_sales["debit_amount"].fillna(0).sum()
    total_bags = qty_bags.sum()
    return total_debit / total_bags if total_bags else np.nan

def price_by_day(df_sales):
    group = df_sales.groupby(df_sales["doc_date"].dt.date).apply(
        lambda g: g["debit_amount"].sum() / (g["qty_mt"].sum() * 20.0) if g["qty_mt"].sum() else np.nan
    ).rename("price_per_bag").reset_index().rename(columns={"doc_date":"date"})
    return group

def ffill_price(daily_df):
    if daily_df.empty:
        return daily_df
    # build continuous date range
    s = daily_df["date"].min()
    e = daily_df["date"].max()
    rng = pd.DataFrame({"date": pd.date_range(s, e, freq="D").date})
    out = rng.merge(daily_df, on="date", how="left").sort_values("date")
    out["price_per_bag"] = out["price_per_bag"].ffill()
    return out

def month_key(d):
    return d.strftime("%Y-%m")

def monthly_summary_ppc(df_all, include_gst=False, gst_pct=0.0):
    # Sales limited to PPC
    pp = df_all[(df_all["entry_type"]=="Sale") & (df_all["product"]=="PPC")].copy()
    if pp.empty:
        return pd.DataFrame()

    pp["month"] = pp["doc_date"].dt.to_period("M").astype(str)
    # Weighted average price per month
    # avg = total debit / (qty*20)
    month_grp = pp.groupby("month").agg(
        qty_mt = ("qty_mt","sum"),
        debit = ("debit_amount","sum")
    ).reset_index()
    month_grp["avg_price_bag"] = month_grp["debit"] / (month_grp["qty_mt"]*20.0)

    # GST option
    if include_gst:
        month_grp["avg_price_bag"] = month_grp["avg_price_bag"] * (1.0 + gst_pct/100.0)

    # CREDIT NOTE (overall, not per product), per month:
    cn = df_all[df_all["entry_type"]=="CreditNote"].copy()
    cn["month"] = cn["doc_date"].dt.to_period("M").astype(str)
    cn_month = cn.groupby("month").agg(credit_note=("credit_amount","sum")).reset_index()

    out = month_grp.merge(cn_month, on="month", how="left")
    out["credit_note"] = out["credit_note"].fillna(0.0)

    # Friendly columns like the screenshot
    # (We keep PPC focus; a "Grand Total" can be added separately)
    out.rename(columns={
        "month":"Year/Month",
        "qty_mt":"QTY(MT) [PPC]",
        "avg_price_bag":"Price/Bag",
        "credit_note":"Credit Note (Overall)"
    }, inplace=True)
    return out

def qty_by_price(df_sales, include_gst=False, gst_pct=0.0):
    # per-line price per bag, then group by exact price
    tmp = df_sales.copy()
    tmp["bags"] = tmp["qty_mt"]*20.0
    # per line price per bag (avoid division by zero)
    tmp["price_per_bag"] = np.where(tmp["bags"]>0, tmp["debit_amount"]/tmp["bags"], np.nan)
    if include_gst:
        tmp["price_per_bag"] = tmp["price_per_bag"] * (1.0 + gst_pct/100.0)
    tb = tmp.groupby(tmp["price_per_bag"].round(2)).agg(qty_mt=("qty_mt","sum"), bags=("bags","sum")).reset_index()
    tb = tb.sort_values("price_per_bag")
    return tb

# ----------------------------- SIDEBAR CONTROLS -----------------------------
st.sidebar.title("⚙️ Controls")
uploaded = st.sidebar.file_uploader("Upload one or more JK Cement ledgers (PDF)", type=["pdf"], accept_multiple_files=True)
include_gst = st.sidebar.toggle("Include GST in price/bag?", value=False)
gst_pct = st.sidebar.number_input("GST % (applied to price/bag if included)", min_value=0.0, max_value=50.0, value=0.0 if not include_gst else 28.0, step=0.25)
default_type = st.sidebar.selectbox("Primary product", options=["PPC","43 GRADE","SUPERSTRONG ADSTAR","(All)"], index=0)

st.sidebar.caption("Tip: Credit Note totals are computed overall from /DG/ entries (RQDBN), while deposits (/DZ/, bank/NEFT/UPI etc.) are excluded.")

# ----------------------------- LOAD + PARSE -----------------------------
if not uploaded:
    st.title("🧱 Cement Ledger Analyzer")
    st.write("Upload your ledgers on the left to get started.")
    st.stop()

dfs = []
for up in uploaded:
    dfs.append(parse_pdf(up.read(), filename=up.name))
df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

if df.empty:
    st.error("Couldn’t read any ledger rows from the PDFs. If the layout differs, ping me and we’ll add a parser tweak.")
    st.stop()

# ----------------------------- PPC-FIRST SUMMARIES -----------------------------
st.title("🧾 Ledger Insights (PPC-first)")

# KPI row
sales_ppc = df[(df["entry_type"]=="Sale") & (df["product"]=="PPC")]
total_qty_ppc = sales_ppc["qty_mt"].sum()
avg_price_ppc = weighted_avg_price_per_bag(sales_ppc)
if include_gst: avg_price_ppc = avg_price_ppc * (1.0 + gst_pct/100.0)

credit_notes_overall = df[df["entry_type"]=="CreditNote"]["credit_amount"].sum()
bank_credits = df[df["entry_type"]=="Bank"]["credit_amount"].sum()

c1,c2,c3,c4 = st.columns(4)
with c1:
    st.metric("Total Qty (PPC) MT", f"{total_qty_ppc:,.2f}")
with c2:
    st.metric(("Avg Invoice Price/Bag (incl GST)" if include_gst else "Avg Invoice Price/Bag"), f"₹ {avg_price_ppc:,.2f}")
with c3:
    st.metric("Credit Notes (overall)", f"₹ {credit_notes_overall:,.0f}")
with c4:
    st.metric("Deposits/Bank credits (excluded)", f"₹ {bank_credits:,.0f}")

# ----------------------------- TABS -----------------------------
tab1, tab2, tab3, tab4 = st.tabs(["📊 PPC Monthly Report", "📈 Price Timeline & Changes", "🏷️ Credit Notes (Monthly)", "🧩 Raw & Other Products"])

with tab1:
    st.subheader("PPC Monthly Report")
    monthly = monthly_summary_ppc(df, include_gst=include_gst, gst_pct=gst_pct)
    if monthly.empty:
        st.info("No PPC sales found.")
    else:
        # totals/footer
        gt_qty = monthly["QTY(MT) [PPC]"].sum()
        # weighted avg overall
        overall_avg = weighted_avg_price_per_bag(sales_ppc)
        if include_gst: overall_avg *= (1.0 + gst_pct/100.0)
        gt_credit = monthly["Credit Note (Overall)"].sum()

        # Pretty table
        st.dataframe(monthly.style.format({
            "QTY(MT) [PPC]":"{:.2f}",
            "Price/Bag":"₹ {:.2f}",
            "Credit Note (Overall)":"₹ {:.0f}"
        }), use_container_width=True)

        # Grand total bar
        st.markdown(f"""
        <div class='kpi-card'>
            <b>Grand Total Qty:</b> {gt_qty:,.2f} MT &nbsp; | &nbsp;
            <b>Overall Avg Price/Bag{ ' (incl GST)' if include_gst else '' }:</b> ₹ {overall_avg:,.2f} &nbsp; | &nbsp;
            <b>Total Credit Notes (Overall):</b> ₹ {gt_credit:,.0f}
        </div>
        """, unsafe_allow_html=True)

        # Visuals
        fig = px.bar(monthly, x="Year/Month", y=["QTY(MT) [PPC]"], title="Monthly Qty (PPC)")
        st.plotly_chart(fig, use_container_width=True)

        fig2 = px.line(monthly, x="Year/Month", y="Price/Bag", markers=True, title=f"Monthly Avg Price/Bag{' (incl GST)' if include_gst else ''} - PPC")
        st.plotly_chart(fig2, use_container_width=True)

with tab2:
    st.subheader("Daily Price Timeline (PPC)")
    pp = sales_ppc.copy()
    if pp.empty:
        st.info("No PPC sales found.")
    else:
        daily = price_by_day(pp)
        daily_ff = ffill_price(daily)

        if include_gst:
            daily["price_per_bag"] = daily["price_per_bag"] * (1.0 + gst_pct/100.0)
            daily_ff["price_per_bag"] = daily_ff["price_per_bag"] * (1.0 + gst_pct/100.0)

        # price change points
        daily_sorted = daily.sort_values("date")
        daily_sorted["prev"] = daily_sorted["price_per_bag"].shift()
        changes = daily_sorted[daily_sorted["price_per_bag"].round(2) != daily_sorted["prev"].round(2)]
        st.write("**Detected price-change dates (PPC):**")
        st.dataframe(changes[["date","price_per_bag"]].rename(columns={"price_per_bag":"Price/Bag"}))

        # Date picker for “price on this date” with carry-forward
        pick_date = st.date_input(
            "Pick a date to fetch the active price (carry-forward if no sale that day)",
            value=daily_ff["date"].max() if not daily_ff.empty else date.today(),
            min_value=daily_ff["date"].min() if not daily_ff.empty else date.today(),
            max_value=daily_ff["date"].max() if not daily_ff.empty else date.today()
        )
        sel = daily_ff[daily_ff["date"]==pick_date]
        if not sel.empty:
            st.success(f"Invoice price on **{pick_date}**: ₹ {sel['price_per_bag'].iloc[0]:.2f} per bag")
        else:
            st.warning("No price available in the selected range.")

        # Qty by exact invoice price
        dist = qty_by_price(pp, include_gst=include_gst, gst_pct=gst_pct)
        st.subheader("Qty sold at each invoice price (exact)")
        st.dataframe(dist.rename(columns={"price_per_bag":"Price/Bag"}).style.format({
            "Price/Bag":"₹ {:.2f}", "qty_mt":"{:.2f}", "bags":"{:.0f}"
        }), use_container_width=True)

        figp = px.bar(dist, x="Price/Bag", y="qty_mt", title="Qty (MT) by Invoice Price (PPC)")
        st.plotly_chart(figp, use_container_width=True)

with tab3:
    st.subheader("Credit Notes by Month (Overall)")
    cn = df[df["entry_type"]=="CreditNote"].copy()
    if cn.empty:
        st.info("No credit notes found (/DG/ RQDBN).")
    else:
        cn["month"] = cn["doc_date"].dt.to_period("M").astype(str)
        cnm = cn.groupby("month").agg(**{"Credit Notes (₹)":("credit_amount","sum")}).reset_index()
        st.dataframe(cnm, use_container_width=True)
        st.plotly_chart(px.bar(cnm, x="month", y="Credit Notes (₹)", title="Monthly Credit Notes (Overall)"), use_container_width=True)

    st.caption("Note: Deposits/fund adjustments (e.g., PIF AXIS-COLL, NEFT/RTGS/UPI) are excluded from the above.")

with tab4:
    st.subheader("All Parsed Rows")
    st.dataframe(df, use_container_width=True, height=500)

    st.subheader("Other Product Summaries")
    if default_type == "(All)":
        filt = df[(df["entry_type"]=="Sale") & df["product"].notna()]
    else:
        filt = df[(df["entry_type"]=="Sale") & (df["product"]==default_type)]
    if filt.empty:
        st.info("No matching sales.")
    else:
        by_prod = (filt
                   .assign(bags=lambda x: x["qty_mt"]*20.0,
                           price_bag=lambda x: x["debit_amount"]/np.where(x["qty_mt"]>0, x["qty_mt"]*20.0, np.nan))
                   .groupby("product")
                   .agg(qty_mt=("qty_mt","sum"), debit=("debit_amount","sum"), avg_price_bag=("price_bag","mean"))
                   .reset_index())
        if include_gst:
            by_prod["avg_price_bag"] = by_prod["avg_price_bag"] * (1.0 + gst_pct/100.0)
        st.dataframe(by_prod.style.format({
            "qty_mt":"{:.2f}", "debit":"₹ {:.0f}", "avg_price_bag":"₹ {:.2f}"
        }), use_container_width=True)

# ----------------------------- DOWNLOADS -----------------------------
with st.expander("⬇️ Export"):
    # Monthly PPC
    monthly = monthly_summary_ppc(df, include_gst=include_gst, gst_pct=gst_pct)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as xw:
        monthly.to_excel(xw, index=False, sheet_name="PPC Monthly")
        # Credit notes
        if not df[df["entry_type"]=="CreditNote"].empty:
            cn = df[df["entry_type"]=="CreditNote"].copy()
            cn["month"] = cn["doc_date"].dt.to_period("M").astype(str)
            cnm = cn.groupby("month").agg(Credit_Notes=("credit_amount","sum")).reset_index()
            cnm.to_excel(xw, index=False, sheet_name="CreditNotes")
        # Raw
        df.to_excel(xw, index=False, sheet_name="Raw")
    st.download_button("Download Excel", data=out.getvalue(), file_name="ledger_analysis.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
