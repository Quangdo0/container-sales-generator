from pathlib import Path
from datetime import date
import tempfile

import streamlit as st

from excel_generator import generate_excel


st.set_page_config(
    page_title="Container Sales Sheet Generator",
    page_icon="📦",
    layout="centered",
)

st.title("📦 Container Sales Sheet Generator")

st.write(
    "Upload the 3 Excel files, choose an export date, then generate the sales sheets automatically."
)

st.warning(
    "Uploaded files are processed temporarily and are not stored permanently. "
    "Do not upload confidential files unless you trust this deployment."
)


# Password is optional.
# Locally, if .streamlit/secrets.toml does not exist, app still works.
try:
    APP_PASSWORD = st.secrets.get("APP_PASSWORD", "")
except Exception:
    APP_PASSWORD = ""

if APP_PASSWORD:
    password = st.text_input("Enter password", type="password")

    if password != APP_PASSWORD:
        st.stop()


st.subheader("1. Upload Excel files")

target_file = st.file_uploader(
    "Target sales Excel file",
    type=["xlsx"],
    help="Upload the sales form/template Excel.",
)

source_file = st.file_uploader(
    "Source report Excel file",
    type=["xlsx"],
    help="Upload the report file that contains Xuất kho and Báo cáo kho tổng hợp sheets.",
)

tracking_file = st.file_uploader(
    "Tracking container Excel file",
    type=["xlsx"],
    help="Upload the tracking file that contains KK / truck / container / trailer data.",
)


st.subheader("2. Choose export date")

selected_date = st.date_input(
    "Export date",
    value=date.today(),
    format="DD/MM/YYYY",
)

run_date_text = selected_date.strftime("%d/%m/%Y")

st.info(f"Selected export date: {run_date_text}")


st.subheader("3. Generate Excel")

generate_button = st.button("Generate Excel", type="primary")


if generate_button:
    if not target_file or not source_file or not tracking_file:
        st.error("Please upload all 3 Excel files.")
    else:
        try:
            with st.spinner("Generating Excel sheets..."):
                with tempfile.TemporaryDirectory() as tmpdir:
                    tmpdir = Path(tmpdir)

                    target_path = tmpdir / target_file.name
                    source_path = tmpdir / source_file.name
                    tracking_path = tmpdir / tracking_file.name
                    output_path = tmpdir / "generated_container_sales.xlsx"

                    target_path.write_bytes(target_file.getbuffer())
                    source_path.write_bytes(source_file.getbuffer())
                    tracking_path.write_bytes(tracking_file.getbuffer())

                    result = generate_excel(
                        target_file_path=target_path,
                        source_file_path=source_path,
                        tracking_file_path=tracking_path,
                        run_date_text=run_date_text,
                        output_file_path=output_path,
                    )

                    output_bytes = output_path.read_bytes()

                    st.success(f"Done. Created sheets: {result['created_count']}")

                    st.download_button(
                        label="Download generated Excel",
                        data=output_bytes,
                        file_name=f"generated_container_sales_{run_date_text.replace('/', '-')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

        except Exception as e:
            st.error("Generation failed.")
            st.exception(e)