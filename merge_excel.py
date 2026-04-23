import streamlit as st
import pandas as pd
from io import BytesIO

st.title("📊 Excel Merger App (Clean Version)")

uploaded_files = st.file_uploader(
    "Upload Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    df_list = []

    for i, file in enumerate(uploaded_files):

        # Read Excel as TEXT to avoid scientific notation issues
        df = pd.read_excel(file, dtype=str)

        if i == 0:
            # First file → keep header
            df_list.append(df)
        else:
            # Later files → remove first row (duplicate header)
            df = df.iloc[1:]
            df_list.append(df)

    merged_df = pd.concat(df_list, ignore_index=True)

    st.write("### Preview")
    st.dataframe(merged_df)

    # Export to Excel WITHOUT scientific formatting
    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        merged_df.to_excel(writer, index=False, sheet_name="MergedData")

        # Force all columns to text format (important fix)
        worksheet = writer.sheets["MergedData"]
        for col in worksheet.columns:
            for cell in col:
                cell.number_format = '@'

    output.seek(0)

    st.download_button(
        "📥 Download Clean Excel File",
        output,
        file_name="merged_clean.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
