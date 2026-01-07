import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

st.set_page_config(page_title="Excel Dashboard", layout="wide")
st.title("📊 Excel Dashboard Tool")

# ----------------------------------------------------
# EXCEL DATE FIXER
# ----------------------------------------------------
def fix_excel_dates(df):
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            s = df[col].dropna()
            if not s.empty and s.between(30000, 60000).all():
                df[col] = pd.to_datetime(
                    df[col],
                    origin="1899-12-30",
                    unit="D",
                    errors="coerce"
                )
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
        header_row = st.number_input("Header row (first row = 1)", min_value=1, value=1)

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
            if pd.api.types.is_numeric_dtype(filtered_df[col]):
                min_v, max_v = filtered_df[col].min(), filtered_df[col].max()
                selected = st.slider(col, float(min_v), float(max_v), (float(min_v), float(max_v)))
                filtered_df = filtered_df[filtered_df[col].between(*selected)]

            elif pd.api.types.is_datetime64_any_dtype(filtered_df[col]):
                min_d = filtered_df[col].min().date()
                max_d = filtered_df[col].max().date()
                selected = st.date_input(col, [min_d, max_d])
                filtered_df = filtered_df[
                    filtered_df[col].dt.date.between(selected[0], selected[1])
                ]

            else:
                values = filtered_df[col].dropna().unique().tolist()
                selected = st.multiselect(col, values, default=values)
                filtered_df = filtered_df[filtered_df[col].isin(selected)]

        st.subheader(f"Filtered Data ({len(filtered_df)} rows)")
        st.dataframe(filtered_df, use_container_width=True)

        # ------------------------------------------------
        # CHARTS
        # ------------------------------------------------
        st.subheader("Charts")

        numeric_cols = filtered_df.select_dtypes(include="number").columns.tolist()
        x_col = st.selectbox("X-axis", filtered_df.columns)
        y_cols = st.multiselect("Y-axis", numeric_cols)
        chart_type = st.selectbox("Chart type", ["Line", "Bar", "Pie", "Combo (Bar + Line)"])

        if st.button("Generate Chart") and y_cols:

            # ---------- PIE (FIXED) ----------
            if chart_type == "Pie":
                for col in y_cols:
                    pie_data = filtered_df.groupby(x_col)[col].sum()

                    fig, ax = plt.subplots()
                    ax.pie(
                        pie_data,
                        labels=pie_data.index.astype(str),
                        autopct="%1.1f%%",
                        startangle=90
                    )
                    ax.set_title(f"{col} Distribution")
                    st.pyplot(fig)

            # ---------- LINE / BAR / COMBO ----------
            else:
                plot_df = filtered_df.copy()

                if pd.api.types.is_datetime64_any_dtype(plot_df[x_col]):
                    plot_df = plot_df.sort_values(by=x_col)
                    plot_df = plot_df.groupby(x_col, as_index=False)[y_cols].sum()

                df_chart = plot_df.set_index(x_col)[y_cols]

                fig, ax = plt.subplots(figsize=(10, 4))

                if chart_type == "Line":
                    df_chart.plot.line(ax=ax)
                elif chart_type == "Bar":
                    df_chart.plot.bar(ax=ax)
                else:
                    df_chart.plot.bar(ax=ax, alpha=0.6)
                    df_chart.plot.line(ax=ax, secondary_y=True)

                ax.set_xlabel(x_col)
                ax.set_ylabel(", ".join(y_cols))
                plt.tight_layout()

                st.pyplot(fig)

                buffer = BytesIO()
                fig.savefig(buffer, format="png")
                buffer.seek(0)

                st.download_button("⬇️ Download chart", buffer, "chart.png")

    except Exception as e:
        st.error(f"Error: {e}")
