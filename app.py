from pathlib import Path
import tempfile
import io
import zipfile
import streamlit as st

from excel_writer import ExcelWriter
from parser_engine import ResultParsingEngine
from result_analyzer import ResultAnalyzer
from result_validator import review_summary
from utils import render_structure

# ---------------------------------------------------------
# Page Configuration
# ---------------------------------------------------------
st.set_page_config(
    page_title="SPPU Result Analyzer & Report Generator",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------------------------------------------------
# Custom CSS Styling (Deployable & Premium UI Design)
# ---------------------------------------------------------
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    /* Hero Banner */
    .hero-container {
        background: linear-gradient(135deg, #1e1b4b 0%, #312e81 40%, #1e293b 100%);
        padding: 2.2rem 2.5rem;
        border-radius: 18px;
        color: #ffffff;
        margin-bottom: 2rem;
        box-shadow: 0 12px 28px -6px rgba(15, 23, 42, 0.35);
        border: 1px solid rgba(255, 255, 255, 0.12);
    }
    
    .hero-title {
        font-size: 2.3rem;
        font-weight: 800;
        letter-spacing: -0.025em;
        margin: 0 0 0.4rem 0;
        background: linear-gradient(90deg, #ffffff, #c7d2fe);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    
    .hero-subtitle {
        font-size: 1.05rem;
        color: #a5b4fc;
        margin-bottom: 1.2rem;
        font-weight: 400;
        max-width: 850px;
        line-height: 1.5;
    }
    
    .badge-pill {
        display: inline-flex;
        align-items: center;
        padding: 0.35rem 0.85rem;
        border-radius: 9999px;
        font-size: 0.8rem;
        font-weight: 600;
        margin-right: 0.5rem;
        margin-bottom: 0.5rem;
        background: rgba(255, 255, 255, 0.12);
        color: #e0e7ff;
        border: 1px solid rgba(255, 255, 255, 0.2);
        backdrop-filter: blur(4px);
    }
    
    /* Metric Cards */
    .metric-card {
        background: #ffffff;
        border-radius: 14px;
        padding: 1.25rem 1.5rem;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.03);
        text-align: center;
    }
    
    .metric-value {
        font-size: 2.2rem;
        font-weight: 800;
        color: #4f46e5;
        line-height: 1.1;
    }
    
    .metric-label {
        font-size: 0.8rem;
        font-weight: 600;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.06em;
        margin-top: 0.35rem;
    }
    
    /* Document Cards */
    .download-card {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.2rem;
        margin-bottom: 1rem;
    }
    
    /* Footer */
    .footer-text {
        text-align: center;
        color: #94a3b8;
        font-size: 0.85rem;
        margin-top: 3.5rem;
        padding-top: 1.5rem;
        border-top: 1px solid #e2e8f0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------
# Sidebar Documentation & Info
# ---------------------------------------------------------
with st.sidebar:
    st.image("https://img.icons8.com/isometric-folders/100/academic-position.png", width=70)
    st.title("SPPU Result Analyzer")
    st.caption("Automated Academic Analysis & Reporting")
    st.markdown("---")
    
    st.markdown("### 📋 Supported Inputs")
    st.markdown(
        """
        - **SPPU Ledger PDFs**: Standard university result ledger documents (`.pdf`).
        - **Master Excel Files**: Processed student result spreadsheets (`.xlsx`).
        """
    )
    
    st.markdown("---")
    st.markdown("### 📄 Generated Reports")
    st.markdown(
        """
        1. **Master Excel Workbook**: Clean tabulated data with student marks, SGPA per semester, and ATKT status.
        2. **Document A (PDF)**: Summary Analysis, Branch/Class performance & ATKT list.
        3. **Document B (PDF)**: Subject-wise Passing & Marks Statistics.
        4. **Document C (PDF)**: Toppers & High Performers Honor Roll.
        """
    )
    st.markdown("---")
    st.info("💡 **Tip**: Upload multiple PDFs to process an entire batch at once.")

# ---------------------------------------------------------
# Hero Banner
# ---------------------------------------------------------
st.markdown(
    """
    <div class="hero-container">
        <div class="hero-title">🎓 SPPU Academic Result Analyzer</div>
        <div class="hero-subtitle">
            Upload Pune University (SPPU) result ledgers or workbooks to instantly generate structured Master Excel reports, ATKT summaries, Subject-wise statistics, and Topper certificates.
        </div>
        <div>
            <span class="badge-pill">⚡ Layout Auto-Detection</span>
            <span class="badge-pill">📊 SGPA & ATKT Tracking</span>
            <span class="badge-pill">📑 Documents A, B & C (PDF)</span>
            <span class="badge-pill">📦 Batch Processing & ZIP Export</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------
# Helper Backend Processing Wrapper (Untouched Backend Logic)
# ---------------------------------------------------------
def process_upload(upload_name: str, upload_bytes: bytes, is_xlsx: bool):
    engine, writer = ResultParsingEngine(), ExcelWriter()
    analyzer = ResultAnalyzer()
    with tempfile.TemporaryDirectory() as folder:
        temp = Path(folder)
        source = temp / upload_name
        source.write_bytes(upload_bytes)

        if is_xlsx:
            output = source
            abc_files = analyzer.generate_abc(excel_path=output, output_dir=temp)
            return {
                "type": "xlsx",
                "excel": output.read_bytes(),
                "a": abc_files["A"].read_bytes(),
                "b": abc_files["B"].read_bytes(),
                "c": abc_files["C"].read_bytes(),
                "excel_name": output.name,
            }
        else:
            result = engine.parse(source)
            suffix = " - REVIEW REQUIRED" if result.requires_review else ""
            output = temp / f"{source.stem}{suffix}.xlsx"
            dataframe = writer.write(result, output)
            abc_files = analyzer.generate_abc(excel_path=output, output_dir=temp, pdf_path=source)

            return {
                "type": "pdf",
                "excel": output.read_bytes(),
                "a": abc_files["A"].read_bytes(),
                "b": abc_files["B"].read_bytes(),
                "c": abc_files["C"].read_bytes(),
                "structure": render_structure(result),
                "review": review_summary(result),
                "requires_review": result.requires_review,
                "num_students": len(result.students),
                "num_columns": dataframe.shape[1],
                "excel_name": output.name,
            }


def create_zip_package(data: dict) -> bytes:
    buffer = io.BytesIO()
    stem = Path(data["excel_name"]).stem
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(data["excel_name"], data["excel"])
        zf.writestr(f"{stem}_Doc_A_ATKT_Summary.pdf", data["a"])
        zf.writestr(f"{stem}_Doc_B_Subject_Passing.pdf", data["b"])
        zf.writestr(f"{stem}_Doc_C_Toppers.pdf", data["c"])
    return buffer.getvalue()


# ---------------------------------------------------------
# Upload Section
# ---------------------------------------------------------
uploads = st.file_uploader(
    "📤 Drop SPPU Result PDFs or Master Excel workbooks below",
    type=["pdf", "xlsx"],
    accept_multiple_files=True,
    help="Supports standard SPPU engineering ledger PDFs and pre-processed Excel files.",
)

# ---------------------------------------------------------
# Processing & Display Results
# ---------------------------------------------------------
if uploads:
    processed_results = []
    
    with st.spinner("Processing uploads & generating reports..."):
        for upload in uploads:
            try:
                is_xlsx = upload.name.endswith(".xlsx")
                data = process_upload(upload.name, upload.getvalue(), is_xlsx)
                data["original_name"] = upload.name
                processed_results.append(data)
            except Exception as error:
                st.error(f"❌ Error processing `{upload.name}`: {error}")

    if processed_results:
        st.markdown("---")
        
        # Summary Overview Header
        st.subheader("📊 Batch Summary Dashboard")
        
        total_files = len(processed_results)
        total_students = sum(d.get("num_students", 0) for d in processed_results if d["type"] == "pdf")
        total_review = sum(1 for d in processed_results if d.get("requires_review", False))
        total_verified = total_files - total_review

        m1, m2, m3, m4 = st.columns(4)
        with m1:
            st.markdown(f'<div class="metric-card"><div class="metric-value">{total_files}</div><div class="metric-label">Files Processed</div></div>', unsafe_allow_html=True)
        with m2:
            st.markdown(f'<div class="metric-card"><div class="metric-value">{total_students if total_students > 0 else "N/A"}</div><div class="metric-label">Total Students</div></div>', unsafe_allow_html=True)
        with m3:
            st.markdown(f'<div class="metric-card"><div class="metric-value" style="color: #10b981;">{total_verified}</div><div class="metric-label">Verified & Ready</div></div>', unsafe_allow_html=True)
        with m4:
            st.markdown(f'<div class="metric-card"><div class="metric-value" style="color: {"#f59e0b" if total_review > 0 else "#64748b"};">{total_review}</div><div class="metric-label">Review Flagged</div></div>', unsafe_allow_html=True)

        st.markdown("### 📁 Processed Files")
        
        # Display each processed file in a tab or expander
        file_tabs = st.tabs([f"📄 {d['original_name']}" for d in processed_results])

        for index, (tab, data) in enumerate(zip(file_tabs, processed_results)):
            with tab:
                st.markdown(f"#### Results for `{data['original_name']}`")

                if data["type"] == "xlsx":
                    st.info("ℹ️ Excel file parsed successfully for Document A, B, & C generation.")
                else:
                    if data["requires_review"]:
                        st.warning(
                            "⚠️ **Review Recommended**: Some entries had formatting anomalies. "
                            "A 'Review Required' sheet has been included in the generated Excel workbook."
                        )
                    else:
                        st.success(
                            f"✅ **Verified Clean**: Parsed **{data['num_students']}** students across **{data['num_columns']}** columns."
                        )

                # Download Options Hub
                st.markdown("##### 📥 Available Downloads")
                
                # Zip Package Download Banner
                zip_data = create_zip_package(data)
                zip_name = f"{Path(data['excel_name']).stem}_All_Reports.zip"
                
                st.download_button(
                    label=f"📦 Download Complete Package (.ZIP) - All 4 Reports for {data['original_name']}",
                    data=zip_data,
                    file_name=zip_name,
                    mime="application/zip",
                    key=f"zip_{index}_{data['original_name']}",
                    use_container_width=True,
                    type="primary",
                )
                
                st.markdown("<br>", unsafe_allow_html=True)
                d1, d2, d3, d4 = st.columns(4)

                with d1:
                    st.download_button(
                        "📊 Master Excel Workbook",
                        data["excel"],
                        data["excel_name"],
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"excel_{index}_{data['original_name']}",
                        use_container_width=True,
                    )
                with d2:
                    st.download_button(
                        "📑 Document A (ATKT & Summary)",
                        data["a"],
                        f"{Path(data['excel_name']).stem}_Doc_A.pdf",
                        "application/pdf",
                        key=f"a_{index}_{data['original_name']}",
                        use_container_width=True,
                    )
                with d3:
                    st.download_button(
                        "📈 Document B (Subject Passing)",
                        data["b"],
                        f"{Path(data['excel_name']).stem}_Doc_B.pdf",
                        "application/pdf",
                        key=f"b_{index}_{data['original_name']}",
                        use_container_width=True,
                    )
                with d4:
                    st.download_button(
                        "🏆 Document C (Topper List)",
                        data["c"],
                        f"{Path(data['excel_name']).stem}_Doc_C.pdf",
                        "application/pdf",
                        key=f"c_{index}_{data['original_name']}",
                        use_container_width=True,
                    )

                # Structure & Audit Log Details
                if data["type"] == "pdf":
                    with st.expander("🔍 Inspect Structure & Parsing Audit Log"):
                        st.markdown("**Detected Structure Breakdown:**")
                        st.code(data["structure"], language="text")
                        st.markdown("**Validation Review Notes:**")
                        st.code(data["review"], language="text")

else:
    # Empty State Guide
    st.markdown(
        """
        <div style="text-align: center; padding: 3rem 1rem; color: #94a3b8;">
            <img src="https://img.icons8.com/cloud-upload/100/upload-to-cloud.png" width="80" style="opacity: 0.6; margin-bottom: 1rem;" />
            <h3>No files uploaded yet</h3>
            <p style="font-size: 0.95rem; max-width: 500px; margin: 0 auto 1.5rem auto;">
                Drag & drop your SPPU Result PDF ledgers or pre-processed Excel workbooks above to analyze results and generate all documents.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

# ---------------------------------------------------------
# Footer
# ---------------------------------------------------------
st.markdown(
    """
    <div class="footer-text">
        SPPU Result Analyzer & Report Generator • Built for Academic Result Processing & Department Analytics
    </div>
    """,
    unsafe_allow_html=True,
)
