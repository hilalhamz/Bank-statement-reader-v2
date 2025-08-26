import streamlit as st
import pandas as pd
import pdfplumber
import io
from datetime import datetime
import xlsxwriter

st.set_page_config(page_title="Bank Statement Reader V2", layout="wide")

st.title("🏦 Bank Statement Reader V2 — PDF & CSV → Excel")

st.write("""
Upload a **bank statement (CSV or PDF)** and get a clean Excel report with:
- Categorized transactions  
- Monthly income vs expenses  
- Category breakdown charts  
""")

# --- File Uploaders ---
uploaded_file = st.file_uploader("Upload a Bank Statement (PDF or CSV)", type=["pdf", "csv"])
uploaded_keywords = st.file_uploader("Upload Keyword Mapping (CSV with Keyword,Category)", type=["csv"])

# Default keywords
keywords = {"uber": "Transport", "netflix": "Entertainment", "carrefour": "Groceries"}

if uploaded_keywords:
    try:
        kw_df = pd.read_csv(uploaded_keywords)
        keywords = dict(zip(kw_df["Keyword"].str.lower(), kw_df["Category"]))
        st.success("Custom keyword mapping loaded.")
    except Exception as e:
        st.error(f"Keyword file error: {e}")

# --- Helper: Categorize transactions ---
def categorize(description):
    if pd.isna(description):
        return "Uncategorized"
    desc = str(description).lower()
    for key, cat in keywords.items():
        if key in desc:
            return cat
    return "Uncategorized"

# --- Parse CSV ---
def parse_csv(file):
    try:
        df = pd.read_csv(file)
    except UnicodeDecodeError:
        df = pd.read_csv(file, encoding="latin1")

    df.columns = [c.lower().strip() for c in df.columns]
    if "date" not in df.columns or "amount" not in df.columns:
        # Try alternative column names
        if "debit" in df.columns and "credit" in df.columns:
            df["amount"] = df["credit"].fillna(0) - df["debit"].fillna(0)
        if "description" not in df.columns:
            df["description"] = df.iloc[:, 1]

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    df["category"] = df["description"].apply(categorize)
    df["type"] = df["amount"].apply(lambda x: "Income" if x > 0 else "Expense")
    return df[["date", "description", "amount", "category", "type"]]

# --- Parse PDF ---
def parse_pdf(file):
    data = []
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if not table:
                continue
            for row in table:
                if len(row) < 2:
                    continue
                try:
                    date = pd.to_datetime(row[0], errors="coerce")
                    desc = row[1]
                    amt = None
                    for cell in row[2:]:
                        try:
                            amt = float(str(cell).replace(",", "").replace("AED", "").strip())
                            break
                        except:
                            continue
                    if date and amt is not None:
                        data.append([date, desc, amt])
                except:
                    continue
    df = pd.DataFrame(data, columns=["date", "description", "amount"])
    df["category"] = df["description"].apply(categorize)
    df["type"] = df["amount"].apply(lambda x: "Income" if x > 0 else "Expense")
    return df

# --- Main processing ---
if uploaded_file:
    if uploaded_file.name.endswith(".csv"):
        df = parse_csv(uploaded_file)
    else:
        df = parse_pdf(uploaded_file)

    st.subheader("📊 Preview")
    st.dataframe(df.head(20))

    # KPIs
    income = df[df["amount"] > 0]["amount"].sum()
    expenses = df[df["amount"] < 0]["amount"].sum()
    net = income + expenses

    col1, col2, col3 = st.columns(3)
    col1.metric("Total Income", f"{income:,.2f}")
    col2.metric("Total Expenses", f"{expenses:,.2f}")
    col3.metric("Net", f"{net:,.2f}")

    # Charts
    monthly = df.groupby(df["date"].dt.to_period("M"))["amount"].sum().reset_index()
    monthly["date"] = monthly["date"].astype(str)

    st.subheader("📈 Monthly Net Flow")
    st.bar_chart(monthly, x="date", y="amount")

    cats = df.groupby("category")["amount"].sum().reset_index()
    st.subheader("📊 Category Breakdown")
    st.bar_chart(cats, x="category", y="amount")

    # --- Export to Excel ---
    def to_excel(df):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name="Transactions", index=False)
            # Summary
            summary = pd.DataFrame({
                "Metric": ["Income", "Expenses", "Net"],
                "Amount": [income, expenses, net]
            })
            summary.to_excel(writer, sheet_name="Summary", index=False)
        return output.getvalue()

    st.download_button(
        label="📥 Download Excel",
        data=to_excel(df),
        file_name="bank_statement_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Please upload a PDF or CSV to begin.")
# Streamlit app entrypoint placeholder
