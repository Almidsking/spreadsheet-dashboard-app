import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from datetime import date

# ----------------------------------------------------
# PAGE CONFIG
# ----------------------------------------------------
st.set_page_config(page_title="Excel Dashboard", layout="wide")
st.title("📊 Excel Dashboard Tool")

# ----------------------------------------------------
# EXCEL DATE FIXER (ROBUST & SAFE)
# ----------------------------------------------------
def fix_excel_dates(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:

        # ---- Excel serial dates (e.g. 45992) ----
        if pd.api.types.is_numeric_dtype(df[col]):
            series = df[col].dropna()
            if not series.empty and series.between(30000, 60000).all():
                df[col] = pd.to_datetime(
                    df[col],
                    origin="1899-12-30",
                    unit="D",
                    errors="coerce"
                )

        # ---- Text dates ----
        elif df[col].dtype == object and "date" in col.lower():
            df[col] = pd.to_datetime(df[col], errors="coerce")

    return df

# ----------------------------------------------------
# FILE UPLOAD
# ----------------------------------------------------
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file, engine="openpyxl")

        sheet = st.selectbox("Select sheet", xls.sheet_names)
        header_row = st.number_input(
            "Header row (first row = 1)",
            min_value=1,
            value=1
        )

        df = pd.read_excel(
            xls,
            sheet_name=sheet,
            header=header_row - 1,
            engine="openpyxl"
        )

        df = fix_excel_dates(df)

        st.success("File loaded successfully")
        st.dataframe(df, use_container_width=True)

        # ------------------------------------------------
        # FILTERS
        # ------------------------------------------------
        st.subheader("Filters")
        filtered_df = df.copy()

        filter_cols = st.multiselect(
            "Select columns to filter",
            filtered_df.columns.tolist()
        )

        for col in filter_cols:

            # Numeric
            if pd.api.types.is_numeric_dtype(filtered_df[col]):
                min_v, max_v = filtered_df[col].min(), filtered_df[col].max()
                selected = st.slider(
                    col,
                    float(min_v),
                    float(max_v),
                    (float(min_v), float(max_v))
                )
                filtered_df = filtered_df[
                    filtered_df[col].between(selected[0], selected[1])
                ]

            # Date
            elif pd.api.types.is_datetime64_any_dtype(filtered_df[col]):
                min_d = filtered_df[col].min().date()
                max_d = filtered_df[col].max().date()
                selected = st.date_input(col, [min_d, max_d])

                filtered_df = filtered_df[
                    filtered_df[col].dt.date.between(selected[0], selected[1])
                ]

            # Text
            else:
                values = filtered_df[col].dropna().unique().tolist()
                selected = st.multiselect(col, values, default=values)
                filtered_df = filtered_df[filtered_df[col].isin(selected)]

        st.subheader(f"Filtered Data ({len(filtered_df)} rows)")
        st.dataframe(filtered_df, use_container_width=True)

        # ------------------------------------------------
        # KPIs
        # ------------------------------------------------
        st.subheader("KPIs")

        numeric_cols = filtered_df.select_dtypes(include="number").columns.tolist()

        if numeric_cols:
            kpi_cols = st.multiselect(
                "Select KPI columns",
                numeric_cols,
                default=numeric_cols[:3]
            )

            kpi_df = pd.DataFrame({
                col: {
                    "Sum": filtered_df[col].sum(),
                    "Average": filtered_df[col].mean(),
                    "Count": filtered_df[col].count()
                }
                for col in kpi_cols
            }).T

            st.dataframe(kpi_df)

        # ------------------------------------------------
        # CHARTS
        # ------------------------------------------------
        st.subheader("Charts")

        x_col = st.selectbox("X-axis", filtered_df.columns)
        y_cols = st.multiselect("Y-axis", numeric_cols)
        chart_type = st.selectbox(
            "Chart type",
            ["Line", "Bar", "Pie", "Combo (Bar + Line)"]
        )

        if st.button("Generate Chart") and y_cols:

            # ---------- PIE ----------
            if chart_type == "Pie":
                for col in y_cols:
                    fig, ax = plt.subplots()
                    filtered_df[col].value_counts().plot.pie(
                        autopct="%1.1f%%",
                        startangle=90,
                        ax=ax
                    )
                    ax.set_ylabel("")
                    ax.set_title(col)
                    st.pyplot(fig)

            # ---------- LINE / BAR / COMBO ----------
            else:
                plot_df = filtered_df.copy()

                # Sort by X-axis if it is date
                if pd.api.types.is_datetime64_any_dtype(plot_df[x_col]):
                    plot_df = plot_df.sort_values(by=x_col)

                    # Aggregate duplicates
                    plot_df = plot_df.groupby(
                        x_col, as_index=False
                    )[y_cols].sum()

                df_chart = plot_df.set_index(x_col)[y_cols]

                fig, ax = plt.subplots(figsize=(10, 4))

                if chart_type == "Line":
                    df_chart.plot.line(ax=ax)

                elif chart_type == "Bar":
                    df_chart.plot.bar(ax=ax)

                else:  # Combo
                    df_chart.plot.bar(ax=ax, alpha=0.6)
                    df_chart.plot.line(ax=ax, secondary_y=True)

                ax.set_xlabel(x_col)
                ax.set_ylabel(", ".join(y_cols))
                plt.tight_layout()

                st.pyplot(fig)

                buffer = BytesIO()
                fig.savefig(buffer, format="png")
                st.download_button(
                    "⬇️ Download chart",
                    buffer,
                    "chart.png"
                )

    except Exception as e:
        st.error(f"Error: {e}")
