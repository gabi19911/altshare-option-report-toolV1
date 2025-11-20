import streamlit as st
import pandas as pd
from datetime import datetime
import io

# ==========================
#  עיצוב כללי ואזור עליון
# ==========================

st.set_page_config(page_title="Altshare Option Report", layout="wide")

# לוגו במרכז
col1, col2, col3 = st.columns([1,2,1])
with col2:
    st.image("images.png", width=160)
    st.markdown("<h2 style='text-align:center;'>Altshare - Option Valuation Automated Tool</h2>", unsafe_allow_html=True)

st.markdown("---")

# ==========================
#   קלט משתמש
# ==========================

st.sidebar.header("הגדרות חישוב")

report_date_str = st.sidebar.text_input("תאריך הדוח (yyyy-mm-dd):", "")
closing_price = st.sidebar.number_input("מחיר סגירה:", min_value=0.0, step=0.01)

eur_rate = st.sidebar.number_input("EUR → USD (אם לא צריך, השאר 0):", value=0.0, step=0.0001)
gbp_rate = st.sidebar.number_input("GBP → USD (אם לא צריך, השאר 0):", value=0.0, step=0.0001)
ils_rate = st.sidebar.number_input("ILS → USD (אם לא צריך, השאר 0):", value=0.0, step=0.0001)

uploaded_file = st.file_uploader("העלה את קובץ ה-Excel", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        st.success("הקובץ נטען בהצלחה!")
        st.write("### תצוגה מקדימה של הנתונים:")
        st.dataframe(df.head())

        # המרת תאריך דוח
        report_date = datetime.strptime(report_date_str, "%Y-%m-%d")

        # ==========================
        #      המרת מטבעות
        # ==========================

        def convert_currency(row):
            price = row["Exercise Price"]
            currency = str(row.get("Exercise Price Currency", "")).strip().upper()

            if currency == "EUR" and eur_rate > 0:
                return price * eur_rate

            if currency == "GBP" and gbp_rate > 0:
                return price * gbp_rate

            # ILS מעל 0.1 → המרה
            if currency == "ILS" and price > 0.1 and ils_rate > 0:
                return price * ils_rate

            return price  # אין המרה

        df["Exercise Price (Converted)"] = df.apply(convert_currency, axis=1)

        # ==========================
        #      חישוב עמודה O
        # ==========================

        def calc_O(row):
            emp_term = row.get("Employment Termination Date")
            orig_exp = row.get("Original Expiry Date")
            upd_exp = row.get("Updated Expiry Date")

            if pd.notnull(emp_term) and pd.notnull(orig_exp):
                if emp_term > report_date and orig_exp > report_date:
                    return (orig_exp - report_date).days / 365

            if pd.notnull(upd_exp) and upd_exp > report_date:
                return (upd_exp - report_date).days / 365

            return 0

        df["O"] = df.apply(calc_O, axis=1)

        # ==========================
        #      חישוב עמודה X
        # ==========================

        df["X"] = df.apply(
            lambda row: max(closing_price - row["Exercise Price (Converted)"], 0),
            axis=1,
        )

        # ==========================
        #   חישוב מדדים סופיים
        # ==========================

        AE = df["Outstanding"]
        AH = df["Exercisable"]
        O = df["O"]
        X = df["X"]

        results = {
            "Weighted Average Exercise Price - Outstanding": (AE * (df["Exercise Price (Converted)"] / AE.sum())).sum(),
            "Weighted Average Exercise Price - Exercisable": (AH * (df["Exercise Price (Converted)"] / AH.sum())).sum(),
            "Weighted Average Remaining Contractual Life - Outstanding": (AE * (O / AE.sum())).sum(),
            "Weighted Average Remaining Contractual Life - Exercisable": (AH * (O / AH.sum())).sum(),
            "Aggregate Intrinsic Value - Outstanding": (AE * X).sum(),
            "Aggregate Intrinsic Value - Exercisable": (AH * X).sum(),
        }

        st.markdown("---")
        st.markdown("## 📊 תוצאות החישוב")

        result_df = pd.DataFrame(list(results.items()), columns=["Metric", "Value"])

        # צביעה
        st.dataframe(
            result_df.style.format({"Value": "{:,.4f}"}).highlight_max(subset=["Value"], color="lightgreen")
        )

        # הורדת קובץ
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Data")
            result_df.to_excel(writer, index=False, sheet_name="Summary")

        st.download_button(
            label="📥 הורד קובץ תוצאות",
            data=output.getvalue(),
            file_name="Altshare_Option_Report_Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"שגיאה בעיבוד: {e}")
