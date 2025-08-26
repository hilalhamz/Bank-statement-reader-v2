import io, re
from datetime import datetime
import pandas as pd
import streamlit as st
import pdfplumber
import xlsxwriter

st.set_page_config(page_title="Bank Statement Reader V2", layout="wide")
st.title("Bank Statement Reader V2 — PDF & CSV → Excel")

st.write(
    "Upload a bank statement (PDF or CSV). Get a clean Excel with categorized transactions, KPIs, "
    "monthly chart, and category breakdown. If your PDF is password-protected, enter it below. "
    "Scanned PDFs (images) aren’t supported—export CSV from your bank or OCR first."
)

uploaded_file   = st.file_uploader("Upload statement (PDF or CSV)", type=["pdf", "csv"])
pdf_password    = st.text_input("PDF password (if protected)", type="password")
kw_file         = st.file_uploader("Optional keyword map (CSV: Keyword,Category)", type=["csv"])

DEFAULT_KEYWORDS = {
    "uber":"Transport","taxi":"Transport","fuel":"Transport","shell":"Transport",
    "carrefour":"Groceries","grocery":"Groceries",
    "netflix":"Subscriptions","spotify":"Subscriptions",
    "amazon":"Shopping","noon":"Shopping",
    "gym":"Health & Fitness","pharmacy":"Health & Fitness",
    "rent":"Housing","etisalat":"Utilities","du ":"Utilities",
    "salary":"Income","transfer in":"Income","transfer out":"Transfers",
}

# --- patterns (added 13JUL25 style) ---
DATE_PATS = [
    r"\b\d{2}[A-Za-z]{3}\d{2,4}\b",       # 13JUL25 / 13JUL2025
    r"\b\d{4}-\d{2}-\d{2}\b",             # 2025-08-05
    r"\b\d{2}/\d{2}/\d{4}\b",             # 05/08/2025
    r"\b\d{2}-\d{2}-\d{4}\b",             # 05-08-2025
    r"\b\d{2}\s+[A-Za-z]{3}\s+\d{4}\b",   # 05 Aug 2025
    r"\b[A-Za-z]{3}\s+\d{2},\s+\d{4}\b"   # Aug 05, 2025
]
AMOUNT_PAT = r"[+-]?\(?\s*[$AEDQRSAR£€₹]?\s*\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?\s*(?:Cr|DR|CR|Dr)?\s*\)?"

def load_keywords(uploaded):
    if not uploaded: return DEFAULT_KEYWORDS
    try:
        df = pd.read_csv(uploaded)
        df.columns = [c.strip() for c in df.columns]
        if not {"Keyword","Category"}.issubset(df.columns):
            st.warning("Keyword CSV must have columns: Keyword, Category. Using defaults.")
            return DEFAULT_KEYWORDS
        mp = {}
        for k,c in zip(df["Keyword"], df["Category"]):
            k=str(k).strip().lower(); c=str(c).strip()
            if k: mp[k]=c
        return mp or DEFAULT_KEYWORDS
    except Exception as e:
        st.warning(f"Could not read keyword CSV ({e}). Using defaults.")
        return DEFAULT_KEYWORDS

KEYWORDS = load_keywords(kw_file)

def categorize(desc:str)->str:
    if pd.isna(desc): return "Uncategorized"
    t=str(desc).lower()
    for kw,cat in KEYWORDS.items():
        if kw in t: return cat
    return "Uncategorized"

# ---------------- CSV ----------------
def parse_csv(file)->pd.DataFrame:
    try: df = pd.read_csv(file)
    except UnicodeDecodeError: df = pd.read_csv(file, encoding="latin1")
    df.columns = [c.lower().strip() for c in df.columns]

    date_col = next((c for c in df.columns if any(k in c for k in ["date","txn date","posting date"])), None) or df.columns[0]
    desc_col = next((c for c in df.columns if any(k in c for k in ["description","details","memo","narration","merchant"])), None) or df.columns[min(1,len(df.columns)-1)]
    amt_col  = next((c for c in df.columns if any(k in c for k in ["amount","amt","value"])), None)

    if amt_col is None:
        debit = next((c for c in df.columns if "debit" in c), None)
        credit= next((c for c in df.columns if "credit" in c), None)
        if debit or credit:
            df["__amount__"]=0.0
            if debit:
                df["__amount__"] -= pd.to_numeric(pd.Series(df[debit]).astype(str).str.replace(",","").str.extract(r"([-+]?\d*\.?\d+)")[0], errors="coerce").fillna(0)
            if credit:
                df["__amount__"] += pd.to_numeric(pd.Series(df[credit]).astype(str).str.replace(",","").str.extract(r"([-+]?\d*\.?\d+)")[0], errors="coerce").fillna(0)
            amt_col="__amount__"
        else:
            amt_col=df.columns[-1]

    out = pd.DataFrame({
        "Date": pd.to_datetime(df[date_col], errors="coerce"),
        "Description": df[desc_col].astype(str),
        "Amount": pd.to_numeric(pd.Series(df[amt_col]).astype(str).str.replace(",","").str.replace("(","-").str.replace(")",""), errors="coerce")
    }).dropna(subset=["Date","Amount"])
    out["Category"]=out["Description"].apply(categorize)
    out["Type"]=out["Amount"].apply(lambda x:"Income" if x>0 else "Expense" if x<0 else "")
    return out

# ------------- PDF (tables + text) -------------
def _clean_amount(s:str):
    s = s.replace(",","").strip()
    neg = "(" in s and ")" in s
    s = re.sub(r"[^\d\.\-]", "", s)  # strip currency, Cr/DR
    if not s: return None
    try:
        v = float(s)
        return -abs(v) if neg else v
    except: return None

def parse_pdf(file, password:str="")->pd.DataFrame:
    data=[]
    with pdfplumber.open(file, password=(password or "")) as pdf:
        # 1) TABLES: map columns by header (Date / Description / Debits / Credits)
        for page in pdf.pages:
            tables = page.extract_tables() or []
            for tbl in tables:
                if not tbl or len(tbl)<2: continue
                header = [str(h or "").strip().lower() for h in tbl[0]]
                # try match headers
                def _idx(names):
                    for i,h in enumerate(header):
                        if any(n in h for n in names): return i
                    return None
                i_date = _idx(["date"])
                i_desc = _idx(["description","details","narration"])
                i_deb  = _idx(["debit"])
                i_cred = _idx(["credit"])
                i_bal  = _idx(["balance"])
                # if we have at least date + one of debit/credit, treat as structured
                if i_date is not None and (i_deb is not None or i_cred is not None):
                    for row in tbl[1:]:
                        cells = ["" if c is None else str(c) for c in row]
                        raw_date = cells[i_date] if i_date < len(cells) else ""
                        # skip headers/footers like "CARRIED FORWARD" rows
                        if re.search(r"carried\s+forward|brought\s+forward", " ".join(cells), re.I): 
                            continue
                        # some banks print empty desc on amount lines; join middle cells
                        desc_parts=[]
                        for j,c in enumerate(cells):
                            if j in [i_date,i_deb,i_cred,i_bal]: continue
                            desc_parts.append(str(c).strip())
                        desc = " ".join([p for p in desc_parts if p]).strip()
                        debit = _clean_amount(cells[i_deb]) if (i_deb is not None and i_deb < len(cells)) else None
                        credit= _clean_amount(cells[i_cred]) if (i_cred is not None and i_cred < len(cells)) else None
                        amount = (credit or 0) - (debit or 0)
                        # parse date (handles 13JUL25)
                        date_txt = str(raw_date).strip()
                        # quick normalize 13JUL25 -> 13-JUL-2025
                        m = re.fullmatch(r"(\d{2})([A-Za-z]{3})(\d{2,4})", date_txt)
                        if m:
                            dd,mon,yy = m.groups()
                            yy = "20"+yy if len(yy)==2 else yy
                            date_parsed = pd.to_datetime(f"{dd}-{mon}-{yy}", errors="coerce", dayfirst=True)
                        else:
                            # try generic parse
                            date_parsed = pd.to_datetime(date_txt, errors="coerce", dayfirst=True)
                        if pd.notna(date_parsed) and (debit or credit):
                            data.append([date_parsed, desc or "", amount])
        # 2) TEXT fallback (date + last number on line, avoids Balance by skipping 'balance' lines)
        if not data:
            for page in pdf.pages:
                text = page.extract_text() or ""
                for line in text.splitlines():
                    ln = line.strip()
                    if not ln or re.search(r"balance", ln, re.I): 
                        continue
                    # date find
                    dm=None; dtxt=""
                    for pat in DATE_PATS:
                        m=re.search(pat, ln)
                        if m: dm=m; dtxt=m.group(0); break
                    if not dm: continue
                    # amount = last numeric token
                    last=None
                    for m in re.finditer(AMOUNT_PAT, ln.replace(",","")):
                        last=m
                    if not last: continue
                    amt = _clean_amount(last.group(0))
                    if amt is None: continue
                    before = ln[:dm.start()].strip()
                    after  = ln[dm.end():].strip()
                    # remove the amount token from the end region
                    after_wo = re.sub(AMOUNT_PAT, "", after).strip()
                    desc = (before + " " + after_wo).strip()
                    # normalize date token
                    m = re.fullmatch(r"(\d{2})([A-Za-z]{3})(\d{2,4})", dtxt)
                    if m:
                        dd,mon,yy = m.groups()
                        yy = "20"+yy if len(yy)==2 else yy
                        d = pd.to_datetime(f"{dd}-{mon}-{yy}", errors="coerce", dayfirst=True)
                    else:
                        d = pd.to_datetime(dtxt, errors="coerce", dayfirst=True)
                    if pd.notna(d):
                        data.append([d, desc, amt])

    if not data:
        return pd.DataFrame(columns=["Date","Description","Amount","Category","Type"])

    df = pd.DataFrame(data, columns=["Date","Description","Amount"])
    df = df.dropna(subset=["Date","Amount"])
    df["Category"] = df["Description"].apply(categorize)
    df["Type"] = df["Amount"].apply(lambda x: "Income" if x>0 else "Expense" if x<0 else "")
    return df

# ------------- Excel export -------------
def to_excel(transactions: pd.DataFrame, keywords_map: dict) -> bytes:
    bio = io.BytesIO()
    wb = xlsxwriter.Workbook(bio, {"in_memory": True})
    fmt_h = wb.add_format({"bold":True,"bg_color":"#F2F2F2","border":1})
    fmt_cur = wb.add_format({"num_format":"#,##0.00"})
    fmt_date = wb.add_format({"num_format":"yyyy-mm-dd"})
    fmt_b = wb.add_format({"bold":True})

    sh = wb.add_worksheet("Transactions")
    cols = ["Date","Description","Amount","Category","Type"]
    sh.write_row(0,0,cols,fmt_h)
    for r,row in enumerate(transactions.itertuples(index=False), start=1):
        d = row.Date if isinstance(row.Date, pd.Timestamp) else pd.to_datetime(row.Date)
        sh.write_datetime(r,0, datetime.combine(d.date(), datetime.min.time()), fmt_date)
        sh.write(r,1, row.Description)
        sh.write_number(r,2, float(row.Amount), fmt_cur)
        sh.write(r,3, row.Category)
        sh.write(r,4, row.Type)
    sh.autofilter(0,0,max(1,len(transactions)),len(cols)-1)
    sh.set_column(0,0,12); sh.set_column(1,1,48); sh.set_column(2,4,16)

    kw = wb.add_worksheet("Keywords")
    kw.write_row(0,0,["Keyword","Category"],fmt_h)
    for i,(k,c) in enumerate(KEYWORDS.items(), start=1):
        kw.write(i,0,k); kw.write(i,1,c)
    kw.set_column(0,1,24)

    sm = wb.add_worksheet("Summary")
    sm.write("A1","KPIs",fmt_b)
    sm.write_row("A3",["Total Income","Total Expenses","Net"],fmt_h)
    sm.write_formula("A4",'=IFERROR(SUMIF(Transactions!E:E,"Income",Transactions!C:C),0)',fmt_cur)
    sm.write_formula("B4",'=IFERROR(-SUMIF(Transactions!E:E,"Expense",Transactions!C:C),0)',fmt_cur)
    sm.write_formula("C4","=A4-B4",fmt_cur)

    sh.write(0,5,"Month",fmt_h)
    for r in range(1,len(transactions)+1):
        sh.write_formula(r,5,f'=TEXT(A{r+1},"yyyy-mm")')

    sm.write("A7","Monthly Summary",fmt_b)
    sm.write_row("A9",["Month","Income","Expenses"],fmt_h)
    sm.write_formula("A10",'=IFERROR(UNIQUE(FILTER(Transactions!F:F,Transactions!F:F<>"")), "")')
    for i in range(0,24):
        rr=10+i
        sm.write_formula(rr-1,1,f'=IF(A{rr}="", "", IFERROR(SUMIFS(Transactions!C:C,Transactions!E:E,"Income",Transactions!F:F,A{rr}),0))',fmt_cur)
        sm.write_formula(rr-1,2,f'=IF(A{rr}="", "", IFERROR(-SUMIFS(Transactions!C:C,Transactions!E:E,"Expense",Transactions!F:F,A{rr}),0))',fmt_cur)

    chart1 = wb.add_chart({"type":"column"})
    chart1.add_series({"name":"Income","categories":"=Summary!$A$10:$A$33","values":"=Summary!$B$10:$B$33"})
    chart1.add_series({"name":"Expenses","categories":"=Summary!$A$10:$A$33","values":"=Summary!$C$10:$C$33"})
    chart1.set_title({"name":"Monthly Income vs Expenses"})
    sm.insert_chart("E9",chart1,{"x_scale":1.2,"y_scale":1.2})

    chart2 = wb.add_chart({"type":"doughnut"})
    sm.write("A36","Expenses by Category",fmt_b)
    sm.write_row("A38",["Category","Total Spent"],fmt_h)
    sm.write_formula("A39",'=IFERROR(UNIQUE(FILTER(Transactions!D:D,Transactions!D:D<>"")), "")')
    for i in range(0,50):
        rr=39+i
        sm.write_formula(rr-1,1,f'=IF(A{rr}="", "", IFERROR(-SUMIFS(Transactions!C:C,Transactions!D:D,A{rr},Transactions!E:E,"Expense"),0))',fmt_cur)
    chart2.add_series({"name":"Expenses by Category","categories":"=Summary!$A$39:$A$88","values":"=Summary!$B$39:$B$88"})
    chart2.set_title({"name":"Expenses by Category"})
    sm.insert_chart("E38",chart2,{"x_scale":1.2,"y_scale":1.2})

    wb.close(); bio.seek(0)
    return bio.getvalue()

# ---------------- UI ----------------
btn = st.button("Parse & Generate")
if btn:
    if not uploaded_file:
        st.warning("Upload a PDF or CSV to continue."); st.stop()
    try:
        if uploaded_file.name.lower().endswith(".csv"):
            tx = parse_csv(uploaded_file)
        else:
            tx = parse_pdf(uploaded_file, pdf_password)

        if tx.empty:
            st.info("No transactions found. Try another file or use CSV export from your bank."); st.stop()

        st.success(f"Parsed {len(tx):,} transactions.")
        st.dataframe(tx.head(50), use_container_width=True)

        income  = float(tx.loc[tx["Amount"]>0, "Amount"].sum())
        expense = float(-tx.loc[tx["Amount"]<0, "Amount"].sum())
        net     = income - expense
        c1,c2,c3 = st.columns(3)
        c1.metric("Total Income",   f"{income:,.2f}")
        c2.metric("Total Expenses", f"{expense:,.2f}")
        c3.metric("Net",            f"{net:,.2f}")

        tmp = tx.copy(); tmp["Month"] = pd.to_datetime(tmp["Date"]).dt.to_period("M").astype(str)
        st.subheader("Monthly Net Flow"); st.bar_chart(tmp.groupby("Month")["Amount"].sum())
        st.subheader("Category Breakdown"); st.bar_chart(tx.groupby("Category")["Amount"].sum())

        st.download_button(
            "Download Excel (transactions + summary + charts)",
            data=to_excel(tx, KEYWORDS),
            file_name="Bank_Statement_Reader_V2.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.caption("Tip: If your PDF is scanned (no selectable text), use CSV export or OCR first.")
    except Exception:
        st.error("Parsing failed. Check PDF password or use CSV export.")


