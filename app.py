import io
from datetime import datetime

import pandas as pd
import pdfplumber
import streamlit as st
import xlsxwriter

# ---------------------------
# App setup
# ---------------------------
st.set_page_config(page_title="Bank Statement Reader V2", layout="wide")
st.title("🏦 Bank Statement Reader V2 — PDF & CSV → Excel")

st.write(
    "Upload a **bank statement** (PDF or CSV). Get a clean, categorized Excel with KPIs, "
    "monthly chart, and category breakdown. For **password-protected PDFs**, enter the password below. "
    "Scanned PDFs (images) aren’t supported—export CSV from your bank or OCR first."
)

# ---------------------------
# Inputs
# ---------------------------
uploaded_file = st.file_uploader("Upload statement (PDF or CSV)", type=["pdf", "csv"])
pdf_password = st.text_input("PDF password (if your PDF is protected)", type="password", value="")

kw_file = st.file_uploader("Optional: Keyword mapping (CSV with columns: Keyword,Category)", type=["csv"])

# Default keywords (you can change these)
DEFAULT_KEYWORDS = {
    "uber": "Transport",
    "taxi": "Transport",
    "fuel": "Transport",
    "shell": "Transport",
    "carrefour": "Groceries",
    "grocery": "Groceries",
    "netflix": "Subscriptions",
    "spotify": "Subscriptions",
    "amazon": "Shopping",
    "noon": "Shopping",
    "gym": "Health & Fitness",
    "pharmacy": "Health & Fitness",
    "rent": "Housing",
    "etisalat": "Utilities",
    "du ": "Utilities",
    "salary": "Income",
    "transfer in": "Income",
    "transfer out": "Transfers",
}

# ---------------------------
# Keyword loader
# ---------------------------
def load_keywords(uploaded):
    if not uploaded:
        return DEFAULT_KEYWORDS
    try:
        df = pd.read_csv(uploaded)
        df.columns = [c.strip() for c in df.columns]
        if not {"Keyword", "Category"}.issubset(df.columns):
            st.warning("Keyword CSV must have columns: Keyword, Category. Using default mapping.")
            return DEFAULT_KEYWORDS
        # first-match wins; preserve order
        mapping = {}
        for k, c in zip(df["Keyword"], df["Category"]):
            k = str(k).strip().lower()
            c = str(c).strip()
            if k:
                mapping[k] = c
        return mapping or DEFAULT_KEYWORDS
    except Exception as e:
        st.warning(f"Could not read keyword CSV ({e}). Using default mapping.")
        return DEFAULT_KEYWORDS

KEYWORDS = load_keywords(kw_file)

def categorize(desc: str) -> str:
    if pd.isna(desc):
        return "Uncategorized"
    txt = str(desc).lower()
    for kw, cat in KEYWORDS.items():
        if kw in txt:
            return cat
    return "Uncategorized"

# ---------------------------
# CSV parser
# ---------------------------
def parse_csv(file) -> pd.DataFrame:
    try:
        df = pd.read_csv(file)
    except UnicodeDecodeError:
        df = pd.read_csv(file, encoding="latin1")

    # Normalize headers
    cols_lower = [c.lower().strip() for c in df.columns]
    df.columns = cols_lower

    # Heuristics for common schemas
    date_col = None
    descr_col = None
    amount_col = None

    # date
    for c in df.columns:
        if any(k in c for k in ["date", "txn date", "posting date"]):
            date_col = c
            break
    if date_col is None:
        date_col = df.columns[0]  # fallback

    # description
    for c in df.columns:
        if any(k in c for k in ["description", "details", "memo", "narration", "merchant"]):
            descr_col = c
            break
    if descr_col is None:
        descr_col = df.columns[min(1, len(df.columns) - 1)]

    # amount
    for c in df.columns:
        if any(k in c for k in ["amount", "amt", "value"]):
            amount_col = c
            break
    if amount_col is None:
        # If debit/credit exist, combine
        debit = next((c for c in df.columns if "debit" in c), None)
        credit = next((c for c in df.columns if "credit" in c), None)
        if debit or credit:
            df["__amount__"] = 0.0
            if debit:
                df["__amount__"] -= pd.to_numeric(
                    pd.Series(df[debit]).astype(str).str.replace(",", "").str.extract(r"([-+]?\d*\.?\d+)")[0],
                    errors="coerce",
                ).fillna(0)
            if credit:
                df["__amount__"] += pd.to_numeric(
                    pd.Series(df[credit]).astype(str).str.replace(",", "").str.extract(r"([-+]?\d*\.?\d+)")[0],
                    errors="coerce",
                ).fillna(0)
            amount_col = "__amount__"
        else:
            amount_col = df.columns[-1]  # last column fallback

    out = pd.DataFrame(
        {
            "Date": pd.to_datetime(df[date_col], errors="coerce"),
            "Description": df[descr_col].astype(str),
            "Amount": pd.to_numeric(
                pd.Series(df[amount_col])
                .astype(str)
                .str.replace(",", "")
                .str.replace("(", "-")
                .str.replace(")", ""),
                errors="coerce",
            ),
        }
    ).dropna(subset=["Date", "Amount"], how="any")

    out["Category"] = out["Description"].apply(categorize)
    out["Type"] = out["Amount"].apply(lambda x: "Income" if x > 0 else "Expense" if x < 0 else "")
    return out

# ---------------------------
# PDF parser (password-aware)
# ---------------------------
def parse_pdf(file, password: str = "") -> pd.DataFrame:
    data = []

    try:
        with pdfplumber.open(file, password=(password or "")) as pdf:
            for page in pdf.pages:
                # Try all tables on the page
                tables = page.extract_tables() or []
                for tbl in tables:
                    if not tbl or len(tbl) < 1:
                        continue

                    # Header guess
                    header = [str(h) if h is not None else "" for h in tbl[0]]
                    # If header looks junky, treat all rows as data
                    start_idx = 1
                    if sum(1 for h in header if any(ch.isalpha() for ch in h)) < 2:
                        header = [f"col{i+1}" for i in range(len(header))]
                        start_idx = 0

                    for row in tbl[start_idx:]:
                        if not row:
                            continue
                        cells = [(c if c is not None else "") for c in row]

                        # Try first cell as date
                        date_val = pd.to_datetime(str(cells[0]).strip(), errors="coerce")

                        # Amount: search numeric-looking from right
                        amt_val = None
                        for cell in reversed(cells):
                            s = str(cell).replace(",", "").replace("AED", "").strip()
                            try:
                                amt_val = float(s)
                                break
                            except:
                                continue

                        # Description: use second cell or join middle cells
                        descr_val = str(cells[1]).strip() if len(cells) > 1 else " ".join(cells)

                        if pd.notna(date_val) and amt_val is not None:
                            data.append([date_val, descr_val, amt_val])

    except Exception as e:
        st.error(
            "Couldn't open this PDF. If it's password-protected, enter the correct password above. "
            "If it's a scanned image, export CSV from your bank or OCR it first."
        )
        raise e

    if not data:
        st.warning("No table-like data found in this PDF. Try CSV export from your bank.")
        return pd.DataFrame(columns=["Date", "Description", "Amount", "Category", "Type"])

    out = pd.DataFrame(data, columns=["Date", "Description", "Amount"])
    out["Category"] = out["Description"].apply(categorize)
    out["Type"] = out["Amount"].apply(lambda x: "Income" if x > 0 else "Expense" if x < 0 else "")
    return out

# ---------------------------
# Excel export
# ---------------------------
def to_excel(transactions: pd.DataFrame, keywords_map: dict) -> bytes:
    bio = io.BytesIO()
    wb = xlsxwriter.Workbook(bio, {"in_memory": True})

    fmt_header = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1})
    fmt_currency = wb.add_format({"num_format": "#,##0.00"})
    fmt_date = wb.add_format({"num_format": "yyyy-mm-dd"})
    fmt_bold = wb.add_format({"bold": True})

    # Transactions
    sh_tx = wb.add_worksheet("Transactions")
    cols = ["Date", "Description", "Amount", "Category", "Type"]
    sh_tx.write_row(0, 0, cols, fmt_header)
    for r, row in enumerate(transactions.itertuples(index=False), start=1):
        sh_tx.write_datetime(r, 0, datetime.combine(row.Date.to_pydatetime().date(), datetime.min.time()), fmt_date)
        sh_tx.write(r, 1, row.Description)
        sh_tx.write_number(r, 2, float(row.Amount), fmt_currency)
        sh_tx.write(r, 3, row.Category)
        sh_tx.write(r, 4, row.Type)
    sh_tx.autofilter(0, 0, max(1, len(transactions)), len(cols) - 1)
    sh_tx.set_column(0, 0, 12)
    sh_tx.set_column(1, 1, 48)
    sh_tx.set_column(2, 4, 16)

    # Keywords
    sh_kw = wb.add_worksheet("Keywords")
    sh_kw.write_row(0, 0, ["Keyword", "Category"], fmt_header)
    for i, (k, c) in enumerate(keywords_map.items(), start=1):
        sh_kw.write(i, 0, k)
        sh_kw.write(i, 1, c)
    sh_kw.set_column(0, 1, 24)

    # Summary (with formulas)
    sh_sum = wb.add_worksheet("Summary")
    sh_sum.write("A1", "KPIs", fmt_bold)
    sh_sum.write_row("A3", ["Total Income", "Total Expenses", "Net"], fmt_header)
    sh_sum.write_formula("A4", '=IFERROR(SUMIF(Transactions!E:E,"Income",Transactions!C:C),0)', fmt_currency)
    sh_sum.write_formula("B4", '=IFERROR(-SUMIF(Transactions!E:E,"Expense",Transactions!C:C),0)', fmt_currency)
    sh_sum.write_formula("C4", "=A4-B4", fmt_currency)

    # Month derivation on Transactions
    sh_tx.write(0, 5, "Month", fmt_header)
    for r in range(1, len(transactions) + 1):
        sh_tx.write_formula(r, 5, f'=TEXT(A{r+1},"yyyy-mm")')

    # Monthly summary (spill)
    sh_sum.write("A7", "Monthly Summary", fmt_bold)
    sh_sum.write_row("A9", ["Month", "Income", "Expenses"], fmt_header)
    sh_sum.write_formula("A10", '=IFERROR(UNIQUE(FILTER(Transactions!F:F,Transactions!F:F<>"")), "")')
    for i in range(0, 24):
        rr = 10 + i
        sh_sum.write_formula(rr - 1, 1, f'=IF(A{rr}="", "", IFERROR(SUMIFS(Transactions!C:C,Transactions!E:E,"Income",Transactions!F:F,A{rr}),0))', fmt_currency)
        sh_sum.write_formula(rr - 1, 2, f'=IF(A{rr}="", "", IFERROR(-SUMIFS(Transactions!C:C,Transactions!E:E,"Expense",Transactions!F:F,A{rr}),0))', fmt_currency)

    # Charts
    chart1 = wb.add_chart({"type": "column"})
    chart1.add_series({"name": "Income", "categories": "=Summary!$A$10:$A$33", "values": "=Summary!$B$10:$B$33"})
    chart1.add_series({"name": "Expenses", "categories": "=Summary!$A$10:$A$33", "values": "=Summary!$C$10:$C$33"})
    chart1.set_title({"name": "Monthly Income vs Expenses"})
    sh_sum.insert_chart("E9", chart1, {"x_scale": 1.2, "y_scale": 1.2})

    chart2 = wb.add_chart({"type": "doughnut"})
    sh_sum.write("A36", "Expenses by Category", fmt_bold)
    sh_sum.write_row("A38", ["Category", "Total Spent"], fmt_header)
    sh_sum.write_formula("A39", '=IFERROR(UNIQUE(FILTER(Transactions!D:D,Transactions!D:D<>"")), "")')
    for i in range(0, 50):
        rr = 39 + i
        sh_sum.write_formula(rr - 1, 1, f'=IF(A{rr}="", "", IFERROR(-SUMIFS(Transactions!C:C,Transactions!D:D,A{rr},Transactions!E:E,"Expense"),0))', fmt_currency)
    chart2.add_series({"name": "Expenses by Category", "categories": "=Summary!$A$39:$A$88", "values": "=Summary!$B$39:$B$88"})
    chart2.set_title({"name": "Expenses by Category"})
    sh_sum.insert_chart("E38", chart2, {"x_scale": 1.2, "y_scale": 1.2})

    wb.close()
    bio.seek(0)
    return bio.getvalue()

# ---------------------------
# Run
# ---------------------------
process = st.button("Parse & Generate")

if process:
    if not uploaded_file:
        st.warning("Upload a PDF or CSV to continue.")
        st.stop()

    try:
        if uploaded_file.name.lower().endswith(".csv"):
            tx = parse_csv(uploaded_file)
        else:
            tx = parse_pdf(uploaded_file, pdf_password)

        if tx.empty:
            st.info("No transactions found. Try another file or use CSV export from your bank.")
            st.stop()

        st.success(f"Parsed {len(tx):,} transactions.")
        st.dataframe(tx.head(50), use_container_width=True)

        income = float(tx.loc[tx["Amount"] > 0, "Amount"].sum())
        expense = float(-tx.loc[tx["Amount"] < 0, "Amount"].sum())
        net = income - expense

        c1, c2, c3 = st.columns(3)
        c1.metric("Total Income", f"{income:,.2f}")
        c2.metric("Total Expenses", f"{expense:,.2f}")
        c3.metric("Net", f"{net:,.2f}")

        # Monthly chart
        tmp = tx.copy()
        tmp["Month"] = tx["Date"].dt.to_period("M").astype(str)
        st.subheader("📈 Monthly Net Flow")
        st.bar_chart(tmp.groupby("Month")["Amount"].sum())

        st.subheader("📊 Category Breakdown")
        st.bar_chart(tx.groupby("Category")["Amount"].sum())

        excel_bytes = to_excel(tx, KEYWORDS)
        st.download_button(
            "📥 Download Excel (transactions + summary + charts)",
            data=excel_bytes,
            file_name="Bank_Statement_Reader_V2.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.caption("Note: Works with digital PDFs (text-selectable). Scanned PDFs need OCR or CSV export.")
    except Exception as e:
        st.error("Parsing failed. Check password (for PDFs) or try CSV export from your bank.")

