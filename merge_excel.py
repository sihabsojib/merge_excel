import streamlit as st
import pandas as pd
from io import BytesIO

st.title("📊 Excel Merger App (Clean & Safe)")

uploaded_files = st.file_uploader(
    "Upload Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)

run = st.button("🚀 Merge Files")

if uploaded_files and run:

    df_list = []
    original_columns = None  # store header of first file

    for i, file in enumerate(uploaded_files):

        df = pd.read_excel(file, dtype=str)

        if i == 0:
            original_columns = list(df.columns)
        else:
            # Remove only rows that EXACTLY match header
            df = df[~(df.apply(lambda row: list(row) == original_columns, axis=1))]

        df_list.append(df)

    merged_df = pd.concat(df_list, ignore_index=True)

    st.write("### Preview")
    st.dataframe(merged_df)

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
