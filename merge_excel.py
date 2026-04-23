import streamlit as st
import pandas as pd
from io import BytesIO

st.title("📊 Excel Merger App (Clean & Safe)")

uploaded_files = st.file_uploader(
    "Upload Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)

# ---- ADD RUN BUTTON ----
run = st.button("🚀 Merge Files")

if uploaded_files and run:

    df_list = []

    for i, file in enumerate(uploaded_files):

        # FORCE ALL DATA AS TEXT (IMPORTANT FIX)
        df = pd.read_excel(file, dtype=str)

        # remove duplicate header rows from later files
        if i != 0:
            df = df.iloc[1:]

        df_list.append(df)

    merged_df = pd.concat(df_list, ignore_index=True)

    st.write("### Preview")
    st.dataframe(merged_df)

    # ---- EXPORT AS REAL EXCEL (NOT CSV) ----
    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        merged_df.to_excel(writer, index=False, sheet_name="MergedData")

    output.seek(0)

    st.download_button(
        "📥 Download Excel File",
        data=output,
        file_name="merged_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
