from pathlib import Path
import tempfile
import streamlit as st
from excel_writer import ExcelWriter
from parser_engine import ResultParsingEngine
from result_analyzer import ResultAnalyzer
from result_validator import review_summary
from utils import render_structure

st.set_page_config(page_title="SPPU Result Analyzer", page_icon="📊", layout="wide")
st.title("SPPU Result Analyzer")
st.caption("Upload SPPU result PDFs. The app detects the layout, validates the structure, and creates an Excel workbook.")
st.caption("SGPA is exported once per detected semester; failed or withheld SGPA is shown as ----.")
uploads = st.file_uploader("Upload SPPU result PDFs or Excel files", type=["pdf", "xlsx"], accept_multiple_files=True)
if uploads:
    engine, writer = ResultParsingEngine(), ExcelWriter()
    analyzer = ResultAnalyzer()
    with tempfile.TemporaryDirectory() as folder:
        temp = Path(folder)
        for upload in uploads:
            source = temp / upload.name
            source.write_bytes(upload.getvalue())
            try:
                st.subheader(upload.name)
                if upload.name.endswith(".xlsx"):
                    output = source
                    st.info(f"Analyzing uploaded Excel workbook: {upload.name}")
                    abc_files = analyzer.generate_abc(excel_path=output, output_dir=temp)
                else:
                    result = engine.parse(source)
                    st.code(f"{render_structure(result)}\n{review_summary(result)}", language="text")
                    suffix = " - REVIEW REQUIRED" if result.requires_review else ""
                    output = temp / f"{source.stem}{suffix}.xlsx"
                    dataframe = writer.write(result, output)
                    if result.requires_review:
                        st.warning("This file needs review. Its workbook includes a Review Required sheet and preserved Source Text.")
                    else:
                        st.success(f"Verified: {len(result.students)} students, {dataframe.shape[1]} columns")
                    abc_files = analyzer.generate_abc(excel_path=output, output_dir=temp, pdf_path=source)

                st.write("##### 📥 Downloads")
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.download_button(
                        "Download Excel",
                        output.read_bytes(),
                        output.name,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"excel_{upload.name}",
                        use_container_width=True,
                    )
                with col2:
                    st.download_button(
                        "Download A (PDF)",
                        abc_files["A"].read_bytes(),
                        "A.pdf",
                        "application/pdf",
                        key=f"a_{upload.name}",
                        use_container_width=True,
                    )
                with col3:
                    st.download_button(
                        "Download B (PDF)",
                        abc_files["B"].read_bytes(),
                        "B.pdf",
                        "application/pdf",
                        key=f"b_{upload.name}",
                        use_container_width=True,
                    )
                with col4:
                    st.download_button(
                        "Download C (PDF)",
                        abc_files["C"].read_bytes(),
                        "C.pdf",
                        "application/pdf",
                        key=f"c_{upload.name}",
                        use_container_width=True,
                    )
            except Exception as error:
                st.error(f"{upload.name}: {error}")
