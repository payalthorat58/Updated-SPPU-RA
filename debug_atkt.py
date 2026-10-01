from pathlib import Path
from parser_engine import ResultParsingEngine
from excel_writer import ExcelWriter
from result_analyzer import ResultAnalyzer
import tempfile, openpyxl

pdf_path = Path(r'C:\Users\Payal\OneDrive\Documents\Result Analyzer Project\result analyzer testing pdfs\BE IT CEGP011520 (28).pdf')
engine = ResultParsingEngine()
writer = ExcelWriter()

result = engine.parse(pdf_path)
with tempfile.TemporaryDirectory() as folder:
    temp = Path(folder)
    output = temp / f"{pdf_path.stem}.xlsx"
    writer.write(result, output)
    
    # Check what columns exist in the Excel
    wb = openpyxl.load_workbook(output)
    ws = wb.active
    col_count = ws.max_column
    row1 = [ws.cell(1, c).value for c in range(1, col_count + 1)]
    row2 = [ws.cell(2, c).value for c in range(1, col_count + 1)]
    row3 = [ws.cell(3, c).value for c in range(1, col_count + 1)]
    
    info_cols = {}
    curr_sem = None
    curr_sub = None
    for c in range(1, col_count + 1):
        s1 = row1[c - 1]
        s2 = row2[c - 1]
        f = row3[c - 1]
        if s1:
            curr_sem = str(s1).strip()
        if s2:
            curr_sub = str(s2).strip()
        if curr_sem == "Student Information" or curr_sub == "Student Information":
            if f:
                info_cols[str(f).strip()] = c
    
    print("Info columns with SGPA:")
    for name, col in info_cols.items():
        if "SGPA" in name or "sgpa" in name.lower():
            print(f"  '{name}' -> col {col}")
            # Print first 5 student values
            for r in range(4, min(9, ws.max_row + 1)):
                print(f"    Row {r}: {ws.cell(r, col).value}")
    
    # Check what _sgpa_column picks
    analyzer = ResultAnalyzer()
    picked = analyzer._sgpa_column(info_cols)
    print(f"\n_sgpa_column picked column: {picked}")
    if picked:
        print(f"Column name: {[k for k,v in info_cols.items() if v == picked]}")
    
    # Check ATKT count
    data = analyzer.analyze_excel(output, pdf_path)
    students = data["_students"] if "_students" in data else None
    print(f"\nProforma A atkt info:")
    print(f"  registered: {data['proforma_a']['registered']}")
    print(f"  total_passed_wo: {data['proforma_a']['total_passed_wo']}")
    print(f"  with_atkt_str: {data['proforma_a']['with_atkt_str']}")
