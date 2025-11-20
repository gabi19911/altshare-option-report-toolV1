import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Altshare Option Report", layout="wide")

# ---------- לוגו וכותרת ----------
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    st.image("images.png", width=160)
    st.markdown(
        "<h2 style='text-align:center;'>Altshare - Option Valuation Automated Tool</h2>",
        unsafe_allow_html=True,
    )

st.markdown("---")

# ---------- סרגל צד – קלטים ----------
st.sidebar.header("הגדרות חישוב")

report_date_str = st.sidebar.text_input("תאריך הדוח (yyyy-mm-dd):", "")
closing_price = st.sidebar.number_input("מחיר סגירה (AA1):", min_value=0.0, step=0.01)

eur_rate = st.sidebar.number_input("EUR → USD (אם לא צריך, השאר 0):", value=0.0, step=0.0001)
gbp_rate = st.sidebar.number_input("GBP → USD (אם לא צריך, השאר 0):", value=0.0, step=0.0001)
ils_rate = st.sidebar.number_input("ILS → USD (אם לא צריך, השאר 0):", value=0.0, step=0.0001)

uploaded_file = st.file_uploader("העלה את קובץ ה-Excel (כמו Grants Status As Of Date ...)", type=["xlsx"])


# ---------- פונקציות עזר ----------

def safe_weighted_avg(weights: pd.Series, values: pd.Series) -> float:
    """ממוצע משוקלל בסגנון SUMPRODUCT/SUM באקסל."""
    w = weights.fillna(0).astype(float)
    v = values.fillna(0).astype(float)
    total_w = w.sum()
    if total_w == 0:
        return 0.0
    return float((w * v).sum() / total_w)


def yearfrac(start_date, end_date):
    """קירוב ל-YEARFRAC באקסל (Actual/365)."""
    if pd.isna(start_date) or pd.isna(end_date):
        return 0.0
    return (end_date - start_date).days / 365.0


# ---------- לוגיקה ראשית ----------
if uploaded_file:
    if not report_date_str:
        st.error("נא להזין תאריך דוח (yyyy-mm-dd) בסרגל הצד.")
    elif closing_price == 0:
        st.warning("מחיר סגירה 0 – ודא שהזנת את ה-AA1 הנכון.")
    else:
        try:
            # קריאת הקובץ
            df = pd.read_excel(uploaded_file)

            st.success("הקובץ נטען בהצלחה!")
            st.write("### תצוגה מקדימה של הנתונים (20 שורות ראשונות):")
            st.dataframe(df.head(20))

            # תאריך דוח
            report_date = datetime.strptime(report_date_str, "%Y-%m-%d")

            # -------- המרת טיפוסי נתונים --------
            date_cols = [
                "Employment Termination Date",
                "Original Expiry Date",
                "Updated Expiry Date",
            ]
            for c in date_cols:
                if c in df.columns:
                    df[c] = pd.to_datetime(df[c], errors="coerce")

            numeric_cols = [
                "Exercise Price",
                "Outstanding",
                "Exercisable",
            ]
            for c in numeric_cols:
                if c in df.columns:
                    df[c] = pd.to_numeric(df[c], errors="coerce")

            # -------- המרת מטבעות לעמודת Exercise Price (Converted) --------
            def convert_currency(row):
                price = row.get("Exercise Price", 0)
                currency = str(row.get("Exercise Price Currency", "")).strip().upper()

                # אם אין מטבע (או USD) – לא ממירים
                if currency in ("", "USD"):
                    return price

                # EUR
                if currency == "EUR" and eur_rate > 0:
                    return price * eur_rate

                # GBP
                if currency == "GBP" and gbp_rate > 0:
                    return price * gbp_rate

                # ILS מעל 0.1 – צריך להמיר, אם הוזן שער
                if currency == "ILS" and price is not None and price > 0.1 and ils_rate > 0:
                    return price * ils_rate

                # כל השאר – בלי המרה (כולל ILS אם אין שער או <=0.1)
                return price

            df["Exercise Price (Converted)"] = df.apply(convert_currency, axis=1)

            # -------- חישוב עמודה O (Remaining Contractual Life) --------
            def calc_O(row):
                emp_term = row.get("Employment Termination Date")
                orig_exp = row.get("Original Expiry Date")
                upd_exp = row.get("Updated Expiry Date")

                # =IF(AND(Employment Termination Date>$B$1,Original Expiry Date>$B$1),
                #      YEARFRAC($B$1,Original Expiry Date),
                #      IF(Updated Expiry Date>$B$1,YEARFRAC($B$1,Updated Expiry Date),0))

                if (
                    pd.notna(emp_term)
                    and pd.notna(orig_exp)
                    and emp_term > report_date
                    and orig_exp > report_date
                ):
                    return yearfrac(report_date, orig_exp)

                if pd.notna(upd_exp) and upd_exp > report_date:
                    return yearfrac(report_date, upd_exp)

                return 0.0

            df["O"] = df.apply(calc_O, axis=1)

            # -------- חישוב עמודה X (Intrinsic per option) --------
            def calc_X(row):
                ex_price = row.get("Exercise Price (Converted)", 0)
                if pd.isna(ex_price):
                    return 0.0
                return max(closing_price - ex_price, 0.0)

            df["X"] = df.apply(calc_X, axis=1)

            # -------- מדדים סופיים (כמו AN באקסל) --------
            AE = df["Outstanding"].fillna(0).astype(float)      # AE באקסל
            AH = df["Exercisable"].fillna(0).astype(float)      # AH באקסל
            price_conv = df["Exercise Price (Converted)"].fillna(0).astype(float)
            O = df["O"].fillna(0).astype(float)
            X = df["X"].fillna(0).astype(float)

            results = {
                # =SUMPRODUCT(AE, V)/SUM(AE)
                "Weighted Average Exercise Price - Outstanding": safe_weighted_avg(AE, price_conv),
                # =SUMPRODUCT(AH, V)/SUM(AH)
                "Weighted Average Exercise Price - Exercisable": safe_weighted_avg(AH, price_conv),
                # =SUMPRODUCT(AE, O)/SUM(AE)
                "Weighted Average Remaining Contractual Life - Outstanding": safe_weighted_avg(AE, O),
                # =SUMPRODUCT(AH, O)/SUM(AH)
                "Weighted Average Remaining Contractual Life - Exercisable": safe_weighted_avg(AH, O),
                # =SUMPRODUCT(AE, X)
                "Aggregate Intrinsic Value - Outstanding": float((AE * X).sum()),
                # =SUMPRODUCT(AH, X)
                "Aggregate Intrinsic Value - Exercisable": float((AH * X).sum()),
            }

            st.markdown("---")
            st.markdown("## 📊 תוצאות החישוב")

            result_df = pd.DataFrame(list(results.items()), columns=["Metric", "Value"])

            st.dataframe(
                result_df.style.format({"Value": "{:,.4f}"})
                .apply(
                    lambda s: [
                        "background-color: #e8f5e9; font-weight: bold;" for _ in s
                    ]
                    if "Aggregate Intrinsic" in s.name
                    else ["" for _ in s],
                    axis=1,
                )
            )

            # אקספנדר קטן לבדיקה / דיבאג
            with st.expander("Debug info (לבדיקת סתירות מול אקסל)"):
                st.write("Sum Outstanding (AE):", AE.sum())
                st.write("Sum Exercisable (AH):", AH.sum())
                st.write("Min/Max Exercise Price (Converted):", price_conv.min(), price_conv.max())
                st.write("Min/Max O:", O.min(), O.max())
                st.write("Min/Max X:", X.min(), X.max())

            # -------- יצוא לאקסל להורדה --------
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df.to_excel(writer, index=False, sheet_name="Data & Calculations")
                result_df.to_excel(writer, index=False, sheet_name="Summary")

            st.download_button(
                label="📥 הורד קובץ תוצאות (Excel)",
                data=output.getvalue(),
                file_name="Altshare_Option_Report_Results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"שגיאה בעיבוד הקובץ: {e}")
