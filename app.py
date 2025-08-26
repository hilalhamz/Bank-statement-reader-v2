import io, re
from datetime import datetime
import pandas as pd
import streamlit as st
import pdfplumber
import xlsxwriter

# -------------------- App --------------------
st.set_page_config(page_title="Bank Statement Reader V2", layout="wide")
st.title("Bank Statement Reader V2 — PDF & CSV → Excel")
st.write(
    "Upload a bank statement (PDF or CSV). You’ll get a clean, categorized Excel with KPIs, "
    "monthly chart, and category breakdown. If your PDF is password-protected, enter it below. "
    "Scanned PDFs (images) aren’t supported—export CSV or OCR first."
)

uploaded_file = st.file_uploader("Upload statement (PDF or CSV)", type=["pdf", "csv"])
pdf_password  = st.text_input("PDF password (if protected)", type="password")
kw_file       = st.file_uploader("Optional keyword map (CSV: Keyword,Category)", type=["csv"])

# -------------------- Keyword mapping --------------------
DEFAULT_KEYWORDS = {
    "uber":"Transport","taxi":"Transport","fuel":"Transport","shell":"Transport",
    "carrefour":"Groceries","grocery":"Groceries",
    "netflix":"Subscriptions","spotify":"Subscriptions",
    "amazon":"Shopping","noon":"Shopping",
    "gym":"Health & Fitness","pharmacy":"Health & Fitness",
    "rent":"Housing","etisalat":"Utilities","du ":"Utilities",
    "salary":"Income","transfer in":"Income","refund":"Income","reversal":"Income",
}

def load_keywords(uploaded):
    if not uploaded: return DEFAULT_KEYWORDS
    try:
        df = pd.read_csv(uploaded)
        df.columns = [c.strip() for c in df.columns]
        if not {"Keyword","Category"}.issubset(df.columns):
            st.warning("Keyword CSV must have columns: Keyword, Category. Using defaults.")
            return DEFAULT_KEYWORDS
        m={}
        for k,c in zip(df["Keyword"], df["Category"]):
            k=str(k).strip().lower(); c=str(c).strip()
            if k: m[k]=c
        return m or DEFAULT_KEYWORDS
    except Exception as e:
        st.warning(f"Could not read keyword CSV ({e}). Using defaults.")
        return DEFAULT_KEYWORDS

KEYWORDS = load_keywords(kw_file)

def categorize(desc: str) -> str:
    if pd.isna(desc): return "Uncategorized"
    t = str(desc).lower()
    for kw, cat in KEYWORDS.items():
        if kw in t: return cat
    return "Uncategorized"

# -------------------- CSV parser --------------------
def parse_csv(file) -> pd.DataFrame:
    try: df = pd.read_csv(file)
    except UnicodeDecodeError: df = pd.read_csv(file, encoding="latin1")
    df.columns = [c.lower().strip() for c in df.columns]

    date_col = next((c for c in df.columns if any(k in c for k in ["date","txn date","posting date"])), None) or df.columns[0]
    desc_col = next((c for c in df.columns if any(k in c for k in ["description","details","memo","narration","merchant"])), None) or df.columns[min(1,len(df.columns)-1)]
    amt_col  = next((c for c in df.columns if any(k in c for k in ["amount","amt","value"])), None)

    if amt_col is None:
        debit  = next((c for c in df.columns if "debit"  in c), None)
        credit = next((c for c in df.columns if "credit" in c), None)
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
    }).dropna(subset=["Date","Amount"], how="any")

    out["Category"] = out["Description"].apply(categorize)
    out["Type"] = out["Amount"].apply(lambda x: "Income" if x > 0 else "Expense" if x < 0 else "")
    return out

# -------------------- PDF helpers --------------------
DATE_PATS = [
    r"\b\d{2}[A-Za-z]{3}\d{2,4}\b",       # 13JUL25 / 13JUL2025
    r"\b\d{4}-\d{2}-\d{2}\b",             # 2025-08-05
    r"\b\d{2}/\d{2}/\d{4}\b",             # 05/08/2025
    r"\b\d{2}-\d{2}-\d{4}\b",             # 05-08-2025
    r"\b\d{2}\s+[A-Za-z]{3}\s+\d{4}\b",   # 05 Aug 2025
    r"\b[A-Za-z]{3}\s+\d{2},\s+\d{4}\b",  # Aug 05, 2025
]
AMOUNT_TOKEN = r"[+-]?\(?\s*[$AEDQRSAR£€₹]?\s*\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?\s*\)?"

EXPENSE_HINTS = ["pos-purchase","purchase","debit card","card no.","food","restaurant","grocer","super market","noon","amazon","uber","taxi","fuel"]
INCOME_HINTS  = ["salary","refund","credit","reversal","transfer in","deposit"]

NOISE_LINES = re.compile(r"^(CARD\s+NO\.|VALUE\s+DATE|AUTH\s+CODE|REF\s+NO\.|TERMINAL|ACCOUNT\s+NO\.)", re.I)

def looks_expense(text: str) -> bool:
    t = text.lower()
    return any(k in t for k in EXPENSE_HINTS)

def looks_income(text: str) -> bool:
    t = text.lower()
    return any(k in t for k in INCOME_HINTS)

def clean_amount_strict(s: str):
    """float or None. Reject tokens that look like balance (contain Cr/DR)."""
    if re.search(r"\b(cr|dr)\b", s, re.I): return None
    s = s.replace(",", "").strip()
    neg = "(" in s and ")" in s
    s = re.sub(r"[^\d\.\-]", "", s)
    if not s: return None
    try:
        v = float(s)
        return -abs(v) if neg else v
    except: return None

def parse_bank_date(tok: str):
    tok = tok.strip()
    m = re.fullmatch(r"(\d{2})([A-Za-z]{3})(\d{2,4})", tok)  # 13JUL25
    if m:
        dd, mon, yy = m.groups()
        yy = "20"+yy if len(yy)==2 else yy
        return pd.to_datetime(f"{dd}-{mon}-{yy}", errors="coerce", dayfirst=True)
    return pd.to_datetime(tok, errors="coerce", dayfirst=True)

def clean_description(text: str) -> str:
    if not text: return ""
    # merge lines, drop known noise rows
    parts = [ln.strip() for ln in re.split(r"[\r\n]+", str(text)) if ln.strip() and not NOISE_LINES.match(ln)]
    if not parts: return ""
    # If "POS-PURCHASE" present, prefer merchant lines after it
    if any("pos" in p.lower() for p in parts):
        # keep next 1-2 informative lines (letters-heavy)
        keep = []
        for p in parts:
            if "pos" in p.lower(): 
                continue
            # skip lines that are just amounts or dates
            if re.search(r"\b\d{2}\b.*\b[A-Za-z]{3}\b|\bAED\b\s*\d", p): 
                # may include city+amount; still keep merchant name tokens
                pass
            # prefer lines with letters
            if sum(ch.isalpha() for ch in p) >= 4:
                keep.append(p)
        if keep:
            return " ".join(keep[:3])
    # fallback: join first two informative lines
    return " ".join(parts[:3])

# -------------------- PDF parser (tables first, text fallback) --------------------
def parse_pdf(file, password: str = "") -> pd.DataFrame:
    data = []
    with pdfplumber.open(file, password=(password or "")) as pdf:
        # ---------- 1) TABLES: map by exact headers ----------
        for page in pdf.pages:
            tables = page.extract_tables() or []
            for tbl in tables:
                if not tbl or len(tbl) < 2: 
                    continue
                header_raw = [str(h or "") for h in tbl[0]]
                header = [h.strip().lower() for h in header_raw]

                # exact/contains match for bilingual headers
                def find_idx(names):
                    for i, h in enumerate(header):
                        h_clean = re.sub(r"[^a-z]", "", h)  # drop non-latin
                        if any(n in h for n in names) or any(n in h_clean for n in names):
                            return i
                    return None

                i_date = find_idx(["date"])
                i_desc = find_idx(["description","details","narration"])
                i_deb  = find_idx(["debit","debits"])
                i_cred = find_idx(["credit","credits"])
                i_bal  = find_idx(["balance"])  # we never read amount from balance

                if i_date is not None and (i_deb is not None or i_cred is not None):
                    for row in tbl[1:]:
                        cells = ["" if c is None else str(c) for c in row]
                        if re.search(r"\b(brought|carried)\s+forward\b", " ".join(cells), re.I):
                            continue

                        # date
                        raw_date = cells[i_date] if i_date < len(cells) else ""
                        d = parse_bank_date(str(raw_date))

                        # description (merge multi-line cell; strip noise)
                        raw_desc = cells[i_desc] if (i_desc is not None and i_desc < len(cells)) else ""
                        desc = clean_description(raw_desc)

                        # amounts
                        debit  = clean_amount_strict(cells[i_deb])  if (i_deb  is not None and i_deb  < len(cells)) else None
                        credit = clean_amount_strict(cells[i_cred]) if (i_cred is not None and i_cred < len(cells)) else None

                        amount = None
                        if debit is not None and abs(debit) > 0:
                            amount = -abs(debit)   # Expense
                        elif credit is not None and abs(credit) > 0:
                            amount =  abs(credit)  # Income

                        if pd.notna(d) and amount is not None:
                            data.append([d, desc, amount])

        # ---------- 2) TEXT fallback (skip Balance; prefer AED <amount>) ----------
        if not data:
            for page in pdf.pages:
                txt = page.extract_text() or ""
                for line in txt.splitlines():
                    ln = line.strip()
                    if not ln or re.search(r"\bbalance\b", ln, re.I):
                        continue
                    # date
                    dm=None; dtxt=""
                    for pat in DATE_PATS:
                        m = re.search(pat, ln)
                        if m: dm=m; dtxt=m.group(0); break
                    if not dm: 
                        continue
                    # amount
                    amt = None
                    aed = re.search(r"\bAED\s*([0-9,]+(?:\.\d{1,2})?)\b", ln, re.I)
                    if aed:
                        amt = clean_amount_strict(aed.group(1))
                    else:
                        last=None
                        for m in re.finditer(AMOUNT_TOKEN, ln.replace(",","")):
                            last=m
                        if last and not re.search(r"\b(cr|dr)\b", last.group(0), re.I):
                            amt = clean_amount_strict(last.group(0))
                    if amt is None:
                        continue
                    d = parse_bank_date(dtxt)
                    if pd.isna(d):
                        continue
                    before = ln[:dm.start()].strip()
                    after  = ln[dm.end():].strip()
                    if aed:
                        after = after.replace(aed.group(0), "").strip()
                    else:
                        after = re.sub(AMOUNT_TOKEN, "", after).strip()
                    desc = clean_description(before + " " + after)

                    # infer sign from description if needed
                    if looks_expense(desc) and amt > 0: amt = -abs(amt)
                    if looks_income(desc)  and amt < 0: amt =  abs(amt)

                    data.append([d, desc, amt])

    if not data:
        return pd.DataFrame(columns=["Date","Description","Amount","Category","Type"])

    df = pd.DataFrame(data, columns=["Date","Description","Amount"]).dropna(subset=["Date","Amount"])
    df["Category"] = df["Description"].apply(categorize)
    df["Type"] = df["Amount"].apply(lambda x: "Income" if x > 0 else "Expense" if x < 0 else "")
    return df

# -------------------- Excel export --------------------
def to_excel(transactions: pd.DataFrame, keywords_map: dict) -> bytes:
    bio = io.BytesIO()
    wb = xlsxwriter.Workbook(bio, {"in_memory": True})
    fmt_h  = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1})
    fmt_cur= wb.add_format({"num_format": "#,##0.00"})
    fmt_dt = wb.add_format({"num_format": "yyyy-mm-dd"})
    fmt_b  = wb.add_format({"bold": True})

    sh = wb.add_worksheet("Transactions")
    cols = ["Date","Description","Amount","Category","Type"]
    sh.write_row(0,0,cols,fmt_h)
    for r,row in enumerate(transactions.itertuples(index=False), start=1):
        d = row.Date if isinstance(row.Date, pd.Timestamp) else pd.to_datetime(row.Date)
        sh.write_datetime(r,0, datetime.combine(d.date(), datetime.min.time()), fmt_dt)
        sh.write(r,1,row.Description)
        sh.write_number(r,2,float(row.Amount),fmt_cur)
        sh.write(r,3,row.Category)
        sh.write(r,4,row.Type)
    sh.autofilter(0,0,max(1,len(transactions)),len(cols)-1)
    sh.set_column(0,0,12); sh.set_column(1,1,48); sh.set_column(2,4,16)

    kw = wb.add_worksheet("Keywords")
    kw.write_row(0,0,["Keyword","Category"],fmt_h)
    for i,(k,c) in enumerate(keywords_map.items(), start=1):
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
        rr = 10 + i
        sm.write_formula(rr-1,1,f'=IF(A{rr}="", "", IFERROR(SUMIFS(Transactions!C:C,Transactions!E:E,"Income",Transactions!F:F,A{rr}),0))',fmt_cur)
        sm.write_formula(rr-1,2,f'=IF(A{rr}="", "", IFERROR(-SUMIFS(Transactions!C:C,Transactions!E:E,"Expense",Transactions!F:F,A{rr}),0))',fmt_cur)

    chart1 = wb.add_chart({"type":"column"})
    chart1.add_series({"name":"Income","categories":"=Summary!$A$10:$A$33","values":"=Summary!$B$10:$B$33"})
    chart1.add_series({"name":"Expenses","categories":"=Summary!$A$10:$A$33","values":"=Summary!$C$10:$C$33"})
    chart1.set_title({"name":"Monthly Income vs Expenses"})
    sm.insert_chart("E9", chart1, {"x_scale":1.2, "y_scale":1.2})

    chart2 = wb.add_chart({"type":"doughnut"})
    sm.write("A36","Expenses by Category",fmt_b)
    sm.write_row("A38",["Category","Total Spent"],fmt_h)
    sm.write_formula("A39",'=IFERROR(UNIQUE(FILTER(Transactions!D:D,Transactions!D:D<>"")), "")')
    for i in range(0,50):
        rr = 39 + i
        sm.write_formula(rr-1,1,f'=IF(A{rr}="", "", IFERROR(-SUMIFS(Transactions!C:C,Transactions!D:D,A{rr},Transactions!E:E,"Expense"),0))',fmt_cur)
    chart2.add_series({"name":"Expenses by Category","categories":"=Summary!$A$39:$A$88","values":"=Summary!$B$39:$B$88"})
    chart2.set_title({"name":"Expenses by Category"})
    sm.insert_chart("E38", chart2, {"x_scale":1.2, "y_scale":1.2})

    wb.close(); bio.seek(0)
    return bio.getvalue()

# -------------------- Run --------------------
if st.button("Parse & Generate"):
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
        st.subheader("Monthly Net Flow");   st.bar_chart(tmp.groupby("Month")["Amount"].sum())
        st.subheader("Category Breakdown"); st.bar_chart(tx.groupby("Category")["Amount"].sum())

        st.download_button(
            "Download Excel (transactions + summary + charts)",
            data=to_excel(tx, KEYWORDS),
            file_name="Bank_Statement_Reader_V2.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.caption("Debits → negative (Expense). Credits → positive (Income). "
                   "If the PDF is scanned (no selectable text), use CSV or OCR first.")
    except Exception:
        st.error("Parsing failed. Check PDF password or try CSV export.")

