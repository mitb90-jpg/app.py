import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
import pdfplumber

# ---------------- PAGE CONFIG ----------------
st.set_page_config(
    page_title="Smart Transaction Categorizer",
    page_icon="📊",
    layout="wide"
)

# ---------------- HEADER ----------------
col1, col2 = st.columns([1, 6])

with col1:
    st.image("WhatsApp Image 2026-06-12 at 5.22.49 PM.jpeg", width=100)

with col2:
    st.markdown(
        """
        <h1 style="color:#1f4e79; font-size:42px; font-weight:bold; margin-bottom:0;">
        Prime Accounting and Tax
        </h1>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <p style="font-size:30px; color:gray; margin-top:0;">
        LAURIE ANACTA NURSING PROFESSIONAL CORPORATION
        </p>
        """,
        unsafe_allow_html=True
    )

# ---------------- FILE UPLOAD ----------------
uploaded_excel = st.sidebar.file_uploader(
    "Upload Excel File",
    type=["xlsx"],
    key="excel_uploader"
)

uploaded_pdf = st.sidebar.file_uploader(
    "Upload PDF File",
    type=["pdf"],
    key="pdf_uploader"
)

df = None

# ---------------- READ EXCEL ----------------
if uploaded_excel is not None:
    df = pd.read_excel(uploaded_excel)
    st.success("Excel File Uploaded Successfully")

# ---------------- READ PDF ----------------
elif uploaded_pdf is not None:
    st.success("PDF Uploaded Successfully")
    st.write("File Name:", uploaded_pdf.name)

    tables = []

    with pdfplumber.open(uploaded_pdf) as pdf:
        for page in pdf.pages:
            extracted_tables = page.extract_tables()

            for table in extracted_tables:
                if table and len(table) > 1:
                    temp_df = pd.DataFrame(table[1:], columns=table[0])
                    tables.append(temp_df)

    if tables:
        df = pd.concat(tables, ignore_index=True)
    else:
        st.error("No tables found in the PDF.")

# ---------------- PROCESS DATA ----------------
if df is not None:

    # ---------------- CLEAN DATA ----------------
    df.columns = df.columns.astype(str).str.strip()

    df = df.loc[:, ~df.columns.str.contains("^Unnamed", na=False)]
    df = df.dropna(axis=1, how="all")
    df = df.dropna(how="all")

    # ---------------- DATE FIX ----------------
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date

    # ---------------- CATEGORY COLUMN ----------------
    df["Category"] = ""

    # ---------------- CREDIT RULE ----------------
    if "Credit" in df.columns and "Description" in df.columns:
        credit_mask = (
            df["Credit"].notna() &
            df["Description"].astype(str).str.contains(
                "Direct Deposit, HELPING HANDS N PAY/PAY",
                case=False,
                na=False
            )
        )
        df.loc[credit_mask, "Category"] = "Revenue"

    # ---------------- DEBIT RULES ----------------
    rules = {
        "Car rental": "Car rental",
        "INTERAC e-Transfer Sent": "Drawings",
        "CNO": "Dues and Subscriptions",
        "Online Bill Payment, CRA-AMT-OWING": "Government taxes",
        "RNAO renewal|NS RN Lisence": "Lisence Fee",
        "Amazon|Blundston|Walmart|Staples": "Office Supplies",
        "ACLS": "Trainings",
        "Debit Card Purchase, TFC CANADA INC": "Transportation Charges",
        "Air Canada|INT'L POS PUR": "Lisence Fee",
        "Shell": "Vehicle Expenses",
        "Plan Fee|Statement Fee|e-Transfer Fee": "Interest and Bank Charges"
    }

    if "Debit" in df.columns and "Description" in df.columns:
        for keyword, category in rules.items():
            mask = (
                df["Debit"].notna() &
                df["Description"].astype(str).str.contains(
                    keyword,
                    case=False,
                    na=False
                )
            )
            df.loc[mask, "Category"] = category

    # ---------------- ADD Sr. No ----------------
    df = df.reset_index(drop=True)
    df.insert(0, "Sr. No", range(1, len(df) + 1))

    # ---------------- FINAL TABLE ----------------
    st.subheader("📊 Categorized Transactions")
    st.dataframe(df, use_container_width=True, hide_index=True)

    # ---------------- DOWNLOAD ----------------
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Transactions")

    output.seek(0)

    st.download_button(
        "⬇️ Download Excel File",
        data=output,
        file_name="Categorized_Transactions.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Please upload an Excel or PDF file to begin.")
