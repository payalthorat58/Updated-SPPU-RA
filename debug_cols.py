from pathlib import Path
from parser_engine import ResultParsingEngine
from excel_writer import ExcelWriter
import tempfile, openpyxl

pdf_path = Path(r'C:\Users\Payal\OneDrive\Documents\Result Analyzer Project\result analyzer testing pdfs\BE IT CEGP011520 (28).pdf')
engine = ResultParsingEngine()
writer = ExcelWriter()

result = engine.parse(pdf_path)
with tempfile.TemporaryDirectory() as folder:
    temp = Path(folder)
    output = temp / f"{pdf_path.stem}.xlsx"
    writer.write(result, output)
    
    wb = openpyxl.load_workbook(output)
    ws = wb.active
    # Print all row 3 headers (field names)
    for c in range(1, ws.max_column + 1):
        r1 = ws.cell(1, c).value
        r2 = ws.cell(2, c).value
        r3 = ws.cell(3, c).value
        if r3 and 'SGPA' in str(r3):
            print(f"Col {c}: Row1='{r1}' Row2='{r2}' Row3='{r3}'")
