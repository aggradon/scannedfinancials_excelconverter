import streamlit as st
import os
import tempfile
from main import process_image, update_pnl_data, export_to_excel, format_current_state, verify_and_complete_state, parse_verified_state

st.set_page_config(page_title="PnL Image Processor", layout="wide")

st.title("PnL Image Processor")

uploaded_files = st.file_uploader("Choose PnL image files", accept_multiple_files=True, type=['png', 'jpg', 'jpeg'])

if uploaded_files:
    pnl_data = {}
    all_years = set()

    progress_bar = st.progress(0)
    status_text = st.empty()

    for i, uploaded_file in enumerate(uploaded_files):
        # Save the uploaded file temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded_file.name)[1]) as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            tmp_file_path = tmp_file.name

        status_text.text(f"Processing image {i+1}/{len(uploaded_files)}: {uploaded_file.name}")

        transcription = process_image(tmp_file_path)
        if transcription == "UNCLEAR IMAGE":
            st.warning(f"Warning: Unclear image detected for {uploaded_file.name}")
            continue
        
        new_pnl_data = update_pnl_data({}, transcription)
        
        for line_item, year_data in new_pnl_data.items():
            if line_item not in pnl_data:
                pnl_data[line_item] = {}
            pnl_data[line_item].update(year_data)
            all_years.update(year_data.keys())
        
        current_state = format_current_state(pnl_data, all_years)
        verified_state = verify_and_complete_state(current_state, uploaded_file.name)
        pnl_data = parse_verified_state(verified_state)

        # Remove the temporary file
        os.unlink(tmp_file_path)

        progress_bar.progress((i + 1) / len(uploaded_files))

    status_text.text("Processing complete. Generating Excel file...")

    # Generate a temporary file for the Excel output
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_excel:
        excel_path = tmp_excel.name

    export_to_excel(pnl_data, excel_path)

    status_text.text("Excel file generated. Ready for download.")

    with open(excel_path, "rb") as file:
        btn = st.download_button(
            label="Download Excel file",
            data=file,
            file_name="pnl_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Remove the temporary Excel file
    os.unlink(excel_path)

st.write("Upload PnL image files to process and generate an Excel file.")