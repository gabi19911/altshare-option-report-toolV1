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

        st.subheader("📄 הקובץ נטען בהצלחה")
        st.dataframe(df.head(), height=250)

        # --- Fix currency field ---
        def convert_currency(row):
            currency = str(row["Exercise Price Currency"]).upper().strip()

            # Default = USD
            fx = 1.0

            if currency in ["EUR"]:
                fx = eur_rate
            elif currency in ["GBP"]:
                fx = gbp_rate
            elif currency in ["ILS", "NIS", "₪"]:
                fx = ils_rate
            elif currency in ["", "0", "NONE", "N/A", "USD"]:
                fx = 1.0
            else:
                fx = 1.0  # unknown → USD

            try:
                return float(row["Exercise Price"]) * fx
            except:
                return None

        df["Exercise Price USD"] = df.apply(convert_currency, axis=1)

        # Column O calculation (Remaining Contractual Life)
        def calc_O(row):
            try:
                rep = report_date
                emp_term = row["Employment Termination Date"]
                orig_exp = row["Original Expiry Date"]
                upd_exp = row["Updated Expiry Date"]

                if pd.notnull(emp_term) and pd.notnull(orig_exp):
                    if emp_term > rep and orig_exp > rep:
                        return (orig_exp - rep).days / 365

                if pd.notnull(upd_exp) and upd_exp > rep:
                    return (upd_exp - rep).days / 365

                return 0
            except:
                return 0

        df["O"] = df.apply(calc_O, axis=1)

        # Column X = intrinsic (closing – exercise)
        df["X"] = df.apply(
            lambda row: max(closing_price - row["Exercise Price USD"], 0)
            if pd.notnull(row["Exercise Price USD"]) else 0,
            axis=1
        )

        # Weight bases
        AE = df["Outstanding"]
        AH = df["Exercisable"]
        O = df["O"]
        X = df["X"]
        EP = df["Exercise Price USD"]

        # Weighted calculations
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

        # Pretty display
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
