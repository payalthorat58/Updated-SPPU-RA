from pathlib import Path
from parser_engine import ResultParsingEngine
from excel_writer import ExcelWriter
from result_analyzer import ResultAnalyzer
import tempfile

pdf_path = Path(r'C:\Users\Payal\OneDrive\Documents\Result Analyzer Project\result analyzer testing pdfs\BE IT CEGP011520 (28).pdf')
engine = ResultParsingEngine()
writer = ExcelWriter()
analyzer = ResultAnalyzer()

result = engine.parse(pdf_path)
with tempfile.TemporaryDirectory() as folder:
    temp = Path(folder)
    output = temp / f"{pdf_path.stem}.xlsx"
    writer.write(result, output)
    
    data = analyzer.analyze_excel(output, pdf_path)
    print("Proforma A data:")
    for k, v in data["proforma_a"].items():
        print(f"  {k}: {v}")
    
    print("\nProforma C data:")
    for t in data["proforma_c"]:
        print(f"  {t}")
