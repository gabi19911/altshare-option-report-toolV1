import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Altshare Outstanding Report Tool", layout="wide")

# --- HEADER ---
st.markdown("<h1 style='text-align:center; color:#4A4A4A;'>Altshare - Outstanding Report Tool</h1>", unsafe_allow_html=True)

# Load logo safely
try:
    st.image("altshare_logo.png", width=150)
except:
    st.warning("⚠️ Logo not found. Upload 'altshare_logo.png' to the same folder as app.py")

st.write("---")

# --- FILE UPLOAD ---
uploaded_file = st.file_uploader("📂 העלאת קובץ אקסל", type=["xlsx"])

report_date = st.date_input("📅 תאריך הדוח")
closing_price = st.number_input("💲 מחיר סגירה", min_value=0.0, step=0.01)

# FX rates
eur_rate = st.number_input("EUR → USD שער", min_value=0.0, step=0.0001)
gbp_rate = st.number_input("GBP → USD שער", min_value=0.0, step=0.0001)
ils_rate = st.number_input("ILS → USD שער", min_value=0.0, step=0.0001)

st.write("---")

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)

        # --- FIX DATE COLUMNS ---
        date_cols = [
            "Employment Termination Date",
            "Original Expiry Date",
            "Updated Expiry Date",
            "Original Grant Date",
            "Grant Date",
            "Vesting Start Date"
        ]

        for col in date_cols:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], dayfirst=True, errors="coerce")

        st.subheader("📄 הקובץ נטען בהצלחה")
        st.dataframe(df.head(), height=250)

        # --- FX Conversion ---
        def convert_currency(row):
            currency = str(row["Exercise Price Currency"]).upper().strip()
            fx = 1.0  # default = USD

            if currency == "EUR":
                fx = eur_rate
            elif currency == "GBP":
                fx = gbp_rate
            elif currency in ["ILS", "NIS", "₪"]:
                fx = ils_rate

            try:
                return float(row["Exercise Price"]) * fx
            except:
                return None

        df["Exercise Price USD"] = df.apply(convert_currency, axis=1)

        # --- CALCULATE COLUMN O ---
        def calc_O(row):
            rep = pd.to_datetime(report_date)

            emp_term = row["Employment Termination Date"]
            orig_exp = row["Original Expiry Date"]
            upd_exp = row["Updated Expiry Date"]

            # Case 1: terminated but expiry still valid
            if pd.notnull(emp_term) and pd.notnull(orig_exp):
                if emp_term > rep and orig_exp > rep:
                    return (orig_exp - rep).days / 365

            # Case 2: updated expiry still valid
            if pd.notnull(upd_exp) and upd_exp > rep:
                return (upd_exp - rep).days / 365

            return 0

        df["O"] = df.apply(calc_O, axis=1)

        # --- COLUMN X: intrinsic ---
        df["X"] = df.apply(
            lambda row: max(closing_price - row["Exercise Price USD"], 0)
            if pd.notnull(row["Exercise Price USD"]) else 0,
            axis=1
        )

        # --- BASE COLUMNS ---
        AE = df["Outstanding"].fillna(0)
        AH = df["Exercisable"].fillna(0)
        O = df["O"].fillna(0)
        X = df["X"].fillna(0)
        EP = df["Exercise Price USD"].fillna(0)

        # --- RESULTS ---
        results = {
            "Weighted Average Exercise Price - Outstanding": (EP * (AE / AE.sum())).sum(),
            "Weighted Average Exercise Price - Exercisable": (EP * (AH / AH.sum())).sum(),
            "Weighted Average Remaining Contractual Life - Outstanding": (O * (AE / AE.sum())).sum(),
            "Weighted Average Remaining Contractual Life - Exercisable": (O * (AH / AH.sum())).sum(),
            "Aggregate Intrinsic Value - Outstanding": (AE * X).sum(),
            "Aggregate Intrinsic Value - Exercisable": (AH * X).sum(),
        }

        st.subheader("📊 תוצאות החישוב")
        st.write("---")

        # Styled results
        for k, v in results.items():
            st.markdown(
                f"""
                <div style='background-color:#F4F6F9; padding:12px; 
                border-radius:8px; margin-bottom:8px; font-size:18px;'>
                <b>{k}:</b> <span style='color:#2A6ACF;'>{v:,.4f}</span>
                </div>
                """,
                unsafe_allow_html=True
            )

    except Exception as e:
        st.error(f"❌ שגיאה בעיבוד הקובץ: {e}")
