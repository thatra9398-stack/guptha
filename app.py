import tempfile
from pathlib import Path
import streamlit as st
from openpyxl import Workbook

from ais_to_excel import (
    read_ais,
    AisParserError,
    sanitize_sheet_name,
    sheet_base_name,
    write_summary_sheet,
    write_category_sheet,
)

# Set page config
st.set_page_config(
    page_title="AIS to Excel Converter",
    page_icon="📊",
    layout="centered",
)

# Custom CSS to fix input text visibility, alignment, card containers, and clean up labels
st.markdown(
    """
    <style>
    /* Dark container cards for inputs */
    .stAppViewContainer {
        background-color: #0e1117;
        color: #fafafa;
    }
    
    .block-container {
        max-width: 650px !important;
        padding-top: 3rem !important;
    }

    /* Privacy banner */
    .privacy-card {
        background-color: #1e2638;
        border-left: 4px solid #4f8bf9;
        padding: 1rem 1.25rem;
        border-radius: 6px;
        margin-bottom: 2rem;
        font-size: 0.95rem;
        color: #e0e6ed;
    }

    /* Make input text high-contrast white and clearly readable */
    .stTextInput input {
        color: #ffffff !important;
        background-color: #262730 !important;
        font-family: monospace;
        font-size: 1rem !important;
    }
    
    .stTextInput label {
        color: #b0b8c4 !important;
        font-weight: 600;
        font-size: 0.9rem;
    }

    /* Primary button styling */
    .stButton > button {
        background-color: #ff4b4b !important;
        color: white !important;
        font-size: 1rem !important;
        font-weight: 600 !important;
        height: 3rem !important;
        border-radius: 6px !important;
        border: none !important;
        margin-top: 1rem;
    }
    
    .stButton > button:hover {
        background-color: #ff3333 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Header Title
st.title("📊 AIS JSON to Excel Converter")

# Privacy Card
st.markdown(
    """
    <div class="privacy-card">
        🔒 <strong>100% Private & Local:</strong> All decryption and Excel file generation occur entirely on your local machine. No credentials or files leave your computer.
    </div>
    """,
    unsafe_allow_html=True,
)

# Main Inputs
uploaded_file = st.file_uploader(
    "1. Upload AIS JSON File",
    type=["json"],
    help="Select the AIS JSON file downloaded from the Income Tax portal.",
)

st.markdown("### 2. Enter Password Credentials")

col1, col2 = st.columns(2)
with col1:
    pan = st.text_input(
        "PAN",
        placeholder="ABCDE1234F",
        max_chars=10,
        help="10-character Permanent Account Number",
    ).strip().upper()

with col2:
    dob = st.text_input(
        "Date of Birth (DDMMYYYY)",
        type="password",
        placeholder="01011990",
        max_chars=8,
        help="Date of birth in DDMMYYYY format",
    ).strip()

st.caption("🔑 Password format required by portal: **PAN** + **DOB** (e.g., `ABCDE1234F01011990`).")

convert_clicked = st.button("Convert & Generate Excel", use_container_width=True)

# Processing Logic
if convert_clicked:
    if not uploaded_file:
        st.error("⚠️ Please upload an AIS JSON file first.")
    elif not pan or len(pan) != 10 or not pan.isalnum():
        st.error("⚠️ Please enter a valid 10-character PAN (e.g. ABCDE1234F).")
    elif not dob or len(dob) != 8 or not dob.isdigit():
        st.error("⚠️ Please enter Date of Birth as 8 digits (DDMMYYYY).")
    else:
        with st.spinner("Decrypting JSON and creating Excel workbook..."):
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".json") as tmp:
                    tmp.write(uploaded_file.getvalue())
                    tmp_path = Path(tmp.name)

                password = f"{pan}{dob}"
                statement = read_ais(path=tmp_path, password=password)
                frames = statement.to_df()

                used_names = {"Summary"}
                sheet_name_by_key = {
                    key: sanitize_sheet_name(sheet_base_name(key, df), used_names) if len(df) > 0 else None
                    for key, df in sorted(frames.items())
                }

                wb = Workbook()
                write_summary_sheet(wb, statement, frames, sheet_name_by_key)

                for key, df in sorted(frames.items()):
                    if len(df) == 0:
                        continue
                    rows = df.to_dict(orient="records")
                    write_category_sheet(wb, sheet_name_by_key[key], list(df.columns), rows)

                out_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
                wb.save(out_file.name)

                with open(out_file.name, "rb") as f:
                    excel_bytes = f.read()

                output_filename = uploaded_file.name.replace(".json", ".xlsx")

                st.success(f"✅ Successful! Processed {len(frames)} categories.")
                st.download_button(
                    label="📥 Download Structured Excel Workbook (.xlsx)",
                    data=excel_bytes,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            except AisParserError as exc:
                st.error(f"❌ Decryption failed: {exc}. Please verify your PAN and Date of Birth.")
            except Exception as exc:
                st.error(f"❌ Error during conversion: {exc}")
            finally:
                if 'tmp_path' in locals() and tmp_path.exists():
                    tmp_path.unlink()
