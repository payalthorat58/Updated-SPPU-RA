"""Process every PDF in pdfs/ with automatic layout detection and Proforma A, B, C generation."""

from pathlib import Path

from excel_writer import ExcelWriter
from parser_engine import ResultParsingEngine
from result_analyzer import ResultAnalyzer
from result_validator import review_summary
from utils import render_structure


ROOT = Path(__file__).resolve().parent
PDF_DIRECTORY = ROOT / "pdfs"
OUTPUT_DIRECTORY = ROOT / "output"


def main() -> None:
    pdfs = sorted(PDF_DIRECTORY.glob("*.pdf"))
    if not pdfs:
        raise RuntimeError(f"No PDFs found in {PDF_DIRECTORY}")
    engine = ResultParsingEngine()
    writer = ExcelWriter()
    analyzer = ResultAnalyzer()
    for pdf_path in pdfs:
        result = engine.parse(pdf_path)
        # Validation is intentionally printed before the Excel write.
        print(render_structure(result))
        print(review_summary(result))
        suffix = " - REVIEW REQUIRED" if result.requires_review else ""
        output = OUTPUT_DIRECTORY / f"{pdf_path.stem}{suffix}.xlsx"
        dataframe = writer.write(result, output)
        print(f"Created: {output.name} ({dataframe.shape[0]} students x {dataframe.shape[1]} columns)")

        # Perform analysis on the generated Excel and generate A, B, C documents
        abc_files = analyzer.generate_abc(excel_path=output, output_dir=OUTPUT_DIRECTORY, pdf_path=pdf_path)
        print(
            f"Generated Analysis Documents in {OUTPUT_DIRECTORY.name}:\n"
            f"  - Proforma A: {abc_files['A'].name} ({abc_files['A'].stat().st_size} bytes)\n"
            f"  - Proforma B: {abc_files['B'].name} ({abc_files['B'].stat().st_size} bytes)\n"
            f"  - Proforma C: {abc_files['C'].name} ({abc_files['C'].stat().st_size} bytes)\n"
        )


if __name__ == "__main__":
    main()
