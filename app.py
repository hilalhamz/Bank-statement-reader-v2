import io, re
from datetime import datetime
import pandas as pd
import streamlit as st
import pdfplumber
import xlsxwriter

# ================== UI ==================
st.set_page_config(page_title="Bank Statement Reader V2", layout="wide")
st.title("Bank Statement Reader V2 — PDF & CSV → Excel")
st.write(
    "Uploads: PDF (with optional password) or CSV. "
    "Outputs: tidy transactions, KPIs, charts, and an Excel download. "
    "Note: scanned PDFs (images) must be OCR’d or exported as CSV."
)

uploaded_file = st.file_uploader("Upload statement (PDF or CSV)", type=["pdf", "csv"])
pdf_password  = st.text_input("PDF password (if protected)", type="password")
kw_file       = st.file_uploader("Optional keyword map (CSV: Keyword,Category)", type=["csv"])

# ================== Keyword mapping ==================
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
        m = {}
        for k,c in zip(df["Keyword"], df["Category"]):
            k = str(k).strip().lower(); c = str(c).strip()
            if k: m[k] = c
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

# ================== CSV parser ==================
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
            df["__amount__"] = 0.0
            if debit:
                df["__amount__"] -= pd.to_numeric(
                    pd.Series(df[debit]).astype(str).str.replace(",","").str.extract(r"([-+]?\d*\.?\d+)")[0],
                    errors="coerce").fillna(0)
            if credit:
                df["__amount__"] += pd.to_numeric(
                    pd.Series(df[credit]).astype(str).str.replace(",","").str.extract(r"([-+]?\d*\.?\d+)")[0],
                    errors="coerce").fillna(0)
            amt_col = "__amount__"
        else:
            amt_col = df.columns[-1]

    out = pd.DataFrame({
        "Date": pd.to_datetime(df[date_col], errors="coerce"),
        "Description": df[desc_col].astype(str),
        "Amount": pd.to_numeric(
            pd.Series(df[amt_col]).astype(str).str.replace(",","").str.replace("(","-").str.replace(")",""),
            errors="coerce"
        )
    }).dropna(subset=["Date","Amount"], how="any")

    out["Category"] = out["Description"].apply(categorize)
    out["Type"] = out["Amount"].apply(lambda x: "Income" if x > 0 else "Expense" if x < 0 else "")
    return out

# ================== PDF helpers ==================
# Dates like 13JUL25, 13JUL2025, 05/08/2025, etc.
DATE_TOKEN = re.compile(r"(\d{2}[A-Za-z]{3}\d{2,4}|\d{4}-\d{2}-\d{2}|\d{2}/\d{2}/\d{4}|\d{2}-\d{2}-\d{4}|\d{2}\s+[A-Za-z]{3}\s+\d{4}|[A-Za-z]{3}\s+\d{2},\s+\d{4})")
AMOUNT_NUM = re.compile(r"[+-]?\(?\s*\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?\s*\)?$")
AED_IN_DESC = re.compile(r"\bAED\s*([0-9,]+(?:\.\d{1,2})?)\b", re.I)
BALANCE_HINT = re.compile(r"\b(cr|dr)\b", re.I)
NOISE_LINE = re.compile(r"^(CARD\s+NO\.|VALUE\s+DATE|AUTH\s+CODE|REF\s+NO\.|TERMINAL|ACCOUNT\s+NO\.)", re.I)

def parse_bank_date(tok: str):
    tok = tok.strip()
    m = re.fullmatch(r"(\d{2})([A-Za-z]{3})(\d{2,4})", tok)  # 13JUL25
    if m:
        dd, mon, yy = m.groups()
        yy = "20"+yy if len(yy)==2 else yy
        return pd.to_datetime(f"{dd}-{mon}-{yy}", errors="coerce", dayfirst=True)
    return pd.to_datetime(tok, errors="coerce", dayfirst=True)

def clean_amount(s: str):
    s = s.replace(",", "").strip()
    neg = "(" in s and ")" in s
    s = re.sub(r"[^\d\.\-]", "", s)
    if not s: return None
    try:
        v = float(s)
        return -abs(v) if neg else v
    except: return None

def clean_description(block: str) -> str:
    lines = [ln.strip() for ln in re.split(r"[\r\n]+", str(block)) if ln.strip()]
    lines = [ln for ln in lines if not NOISE_LINE.match(ln)]
    if not lines: return ""
    # prefer merchant lines (letters-heavy)
    keep = [ln for ln in lines if sum(ch.isalpha() for ch in ln) >= 4]
    return " ".join(keep[:3]) if keep else " ".join(lines[:3])

# ================== PDF parser (column-locked) ==================
def detect_column_boundaries(first_page):
    """
    Find x-positions of headers: Date | Description | Debits | Credits | Balance.
    Returns list of x midpoints that split the page into 5 column zones.
    """
    words = first_page.extract_words(use_text_flow=True, keep_blank_chars=False)
    headers = {"date": None, "description": None, "debits": None, "credits": None, "balance": None}
    for w in words:
        t = w["text"].strip().lower()
        if t in headers and headers[t] is None:
            headers[t] = (w["x0"], w["x1"])
    # If any not found, try fuzzy contains
    if None in [headers["date"], headers["description"], headers["debits"], headers["credits"], headers["balance"]]:
        for w in words:
            t = re.sub(r"[^a-z]", "", w["text"].lower())
            if "date" in t and headers["date"] is None: headers["date"] = (w["x0"], w["x1"])
            if "description" in t and headers["description"] is None: headers["description"] = (w["x0"], w["x1"])
            if "debit" in t and headers["debits"] is None: headers["debits"] = (w["x0"], w["x1"])
            if "credit" in t and headers["credits"] is None: headers["credits"] = (w["x0"], w["x1"])
            if "balance" in t and headers["balance"] is None: headers["balance"] = (w["x0"], w["x1"])
    if any(v is None for v in headers.values()):
        return None  # fall back to pdfplumber tables later

    # Sort by x and compute split boundaries midway between headers
    centers = sorted([(k, (x0+x1)/2) for k,(x0,x1) in headers.items()], key=lambda x: x[1])
    xs = [c for _, c in centers]
    splits = [(xs[i] + xs[i+1]) / 2 for i in range(len(xs)-1)]
    # Create column windows: (-inf, s0], (s0, s1], ... , (s3, +inf)
    return splits  # 4 split x's produce 5 zones

def assign_column(x_center, splits):
    # zones: 0=Date, 1=Description, 2=Debits, 3=Credits, 4=Balance
    if x_center <= splits[0]: return 0
    if x_center <= splits[1]: return 1
    if x_center <= splits[2]: return 2
    if x_center <= splits[3]: return 3
    return 4

def parse_pdf(file, password: str = "") -> pd.DataFrame:
    rows = []
    with pdfplumber.open(file, password=(password or "")) as pdf:
        if not pdf.pages: return pd.DataFrame(columns=["Date","Description","Amount","Category","Type"])

        # Detect column splits on the first page
        splits = detect_column_boundaries(pdf.pages[0])

        if splits:
            # -------- Column-locked extraction (robust for your layout) --------
            for page in pdf.pages:
                words = page.extract_words(use_text_flow=False, keep_blank_chars=False)
                # Group words into lines using y-center tolerance
                lines = {}
                for w in words:
                    y = round((w["top"] + w["bottom"]) / 2, 1)
                    # ignore header band (big bold yellow band tends to have larger y)
                    lines.setdefault(y, []).append(w)
                # Process lines top-to-bottom
                for y in sorted(lines.keys()):
                    ws = sorted(lines[y], key=lambda z: z["x0"])
                    # Build buckets per column
                    cols = {0:[],1:[],2:[],3:[],4:[]}
                    for w in ws:
                        col = assign_column((w["x0"]+w["x1"])/2, splits)
                        cols[col].append(w["text"])

                    date_txt = " ".join(cols[0]).strip()
                    desc_txt = " ".join(cols[1]).strip()
                    debit_txt = " ".join(cols[2]).strip()
                    credit_txt= " ".join(cols[3]).strip()
                    bal_txt   = " ".join(cols[4]).strip()

                    # Skip header rows & carried/brought forward
                    if any(h in desc_txt.lower() for h in ["brought forward","carried forward"]):
                        continue
                    if date_txt.lower() in ("date","التاريخ") or desc_txt.lower() in ("description","التفاصيل"):
                        continue

                    # a valid row must have a date token or an amount token in debit/credit columns
                    has_date = bool(DATE_TOKEN.search(date_txt))
                    has_debit = AMOUNT_NUM.search(debit_txt) and not BALANCE_HINT.search(debit_txt)
                    has_credit= AMOUNT_NUM.search(credit_txt) and not BALANCE_HINT.search(credit_txt)

                    # If only description (multi-line), attach to last open row
                    if not has_date and not has_debit and not has_credit and desc_txt:
                        if rows and rows[-1].get("_open"):
                            rows[-1]["Description"] += " " + clean_description(desc_txt)
                        continue

                    # Start a new row if date present
                    if has_date:
                        d = parse_bank_date(DATE_TOKEN.search(date_txt).group(0))
                        rows.append({"Date": d, "Description": clean_description(desc_txt), "Amount": None, "_open": True})
                        # keep going to quantities below

                    # If we have amounts, close the current open row or create one
                    if has_debit or has_credit:
                        amt = None
                        if has_debit:
                            amt = -abs(clean_amount(debit_txt) or 0.0)  # debit = expense
                        if has_credit:
                            amt = abs(clean_amount(credit_txt) or 0.0)  # credit = income

                        if rows and rows[-1].get("_open") and rows[-1]["Amount"] is None:
                            rows[-1]["Amount"] = amt
                            rows[-1]["_open"] = False
                        else:
                            # no open row (some banks repeat date above) — synthesize with date from this line if any
                            dd = None
                            if has_date:
                                dd = parse_bank_date(DATE_TOKEN.search(date_txt).group(0))
                            rows.append({"Date": dd, "Description": clean_description(desc_txt), "Amount": amt, "_open": False})
        else:
            # -------- Fallback to pdfplumber tables, then text --------
            data = []
            for page in pdf.pages:
                tables = page.extract_tables() or []
                for tbl in tables:
                    if not tbl or len(tbl) < 2: continue
                    header = [str(h or "").strip().lower() for h in tbl[0]]
                    def idx(names):
                        for i,h in enumerate(header):
                            if any(n in h for n in names): return i
                        return None
                    i_date = idx(["date"]); i_desc=idx(["description","details","narration"])
                    i_deb  = idx(["debit"]); i_cred=idx(["credit"]); i_bal=idx(["balance"])
                    if i_date is not None and (i_deb is not None or i_cred is not None):
                        for row in tbl[1:]:
                            if not any(row): continue
                            if re.search(r"\b(brought|carried)\s+forward\b"," ".join([str(x) for x in row if x]), re.I): 
                                continue
                            d = parse_bank_date(str(row[i_date])) if i_date < len(row) else pd.NaT
                            desc = clean_description(row[i_desc]) if (i_desc is not None and i_desc < len(row)) else ""
                            debit  = clean_amount(str(row[i_deb]))  if (i_deb  is not None and i_deb  < len(row)) else None
                            credit = clean_amount(str(row[i_cred])) if (i_cred is not None and i_cred < len(row)) else None
                            amount = None
                            if debit not in (None, 0): amount = -abs(debit)
                            elif credit not in (None, 0): amount = abs(credit)
                            if pd.notna(d) and amount is not None:
                                data.append([d, desc, amount])

            if not data:
                # line-text parsing
                for page in pdf.pages:
                    text = page.extract_text() or ""
                    for line in text.splitlines():
                        ln = line.strip()
                        if not ln or re.search(r"\bbalance\b", ln, re.I): 
                            continue
                        dm = DATE_TOKEN.search(ln)
                        if not dm: continue
                        d = parse_bank_date(dm.group(0))
                        # prefer AED amount inside description
                        aed = AED_IN_DESC.search(ln)
                        amt = clean_amount(aed.group(1)) if aed else None
                        if amt is None:
                            # last numeric token that is NOT balance-like
                            nums = [m.group(0) for m in re.finditer(AMOUNT_NUM, ln)]
                            if nums:
                                cand = nums[-1]
                                if not BALANCE_HINT.search(cand): amt = clean_amount(cand)
                        if amt is None: continue
                        before = ln[:dm.start()].strip()
                        after  = ln[dm.end():].strip()
                        if aed: after = after.replace(aed.group(0), "").strip()
                        desc = clean_description(before + " " + after)
                        if re.search(r"pos|purchase|restaurant|grocery|uber|noon|amazon", desc, re.I) and amt > 0:
                            amt = -abs(amt)  # infer purchases as expense
                        rows.append({"Date": d, "Description": desc, "Amount": amt, "_open": False})

    # Normalize assembled rows
    tx = pd.DataFrame(rows)
    if tx.empty:
        return pd.DataFrame(columns=["Date","Description","Amount","Category","Type"])
    tx = tx.dropna(subset=["Amount"])
    # if any Date missing on lines with only amounts, forward-fill from above
    tx["Date"] = tx["Date"].ffill()
    tx["Description"] = tx["Description"].fillna("").str.strip().replace({"^POS\\s*$":"POS-PURCHASE"}, regex=True)
    tx = tx[["Date","Description","Amount"]]
    tx["Category"] = tx["Description"].apply(categorize)
    tx["Type"] = tx["Amount"].apply(lambda x: "Income" if x > 0 else "Expense" if x < 0 else "")
    return tx

# ================== Excel export ==================
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

# ================== Run ==================
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
    except Exception as e:
        st.error("Parsing failed. Check PDF password or try CSV export.")

