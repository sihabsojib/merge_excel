import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Excel Merger", page_icon="📊")

st.title("📊 Excel Merger App )")

# -------------------------------
# Initialize uploader key
# -------------------------------
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

# -------------------------------
# File uploader
# -------------------------------
uploaded_files = st.file_uploader(
    "Upload Excel files",
    type=["xlsx"],
    accept_multiple_files=True,
    key=f"uploader_{st.session_state.uploader_key}"
)

# -------------------------------
# Show uploaded file info
# -------------------------------
if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} file(s) uploaded")

    with st.expander("📂 Uploaded Files"):
        for i, file in enumerate(uploaded_files, start=1):
            st.write(f"{i}. {file.name}")

# -------------------------------
# Buttons
# -------------------------------
col1, col2 = st.columns(2)

with col1:
    run = st.button(
        "🚀 Merge Files",
        disabled=not uploaded_files
    )

with col2:
    clear = st.button(
        "🗑 Clear Files",
        disabled=not uploaded_files
    )

# -------------------------------
# Clear uploader
# -------------------------------
if clear:
    st.session_state.uploader_key += 1
    st.rerun()

# -------------------------------
# Merge Files
# -------------------------------
if uploaded_files and run:

    with st.spinner("Merging files..."):

        df_list = []

        for file in uploaded_files:

            # Read everything as text
            df = pd.read_excel(file, dtype=str)

            df_list.append(df)

        merged_df = pd.concat(df_list, ignore_index=True)

    st.success("✅ Merge completed!")

    st.write("### Preview")
    st.dataframe(merged_df, use_container_width=True)

    # -------------------------------
    # Export to Excel
    # -------------------------------
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        merged_df.to_excel(
            writer,
            index=False,
            sheet_name="MergedData"
        )

    output.seek(0)

    st.download_button(
        "📥 Download Merged Excel",
        data=output,
        file_name="merged_file.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
