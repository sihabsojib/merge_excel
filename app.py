import streamlit as st
import pandas as pd
from io import BytesIO

st.title("📊 Excel Merger App")

uploaded_files = st.file_uploader(
    "Upload Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    df_list = []

    for file in uploaded_files:
        df = pd.read_excel(file)
        df_list.append(df)

    merged_df = pd.concat(df_list, ignore_index=True)

    st.write("### Preview")
    st.dataframe(merged_df)

    # Convert to Excel in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        merged_df.to_excel(writer, index=False, sheet_name='MergedData')

    output.seek(0)

    st.download_button(
        label="📥 Download Excel File",
        data=output,
        file_name="merged_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
