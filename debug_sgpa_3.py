from pathlib import Path
from parser_engine import ResultParsingEngine
from excel_writer import ExcelWriter

pdf_path = r'C:\Users\Payal\OneDrive\Documents\Result Analyzer Project\result analyzer testing pdfs\BE IT CEGP011520 (28).pdf'
result = ResultParsingEngine().parse(Path(pdf_path))
all_year_fields = {
    k for student in result.students for k in student.summary.keys()
    if k.endswith(' SGPA') and not k.startswith('Semester')
}

failed = []
missing = []
for student in result.students:
    v = student.summary.get('Fourth Year SGPA')
    if v is None or str(v).strip() in {'', '-', '--', '---', '----'}:
        calculated = ExcelWriter._missing_semester_sgpas(student, all_year_fields)
        calc_val = calculated.get('Fourth Year SGPA')
        if calc_val is None:
            failed.append(student.name)
        else:
            missing.append((student.name, calc_val))

print('Failed (Calculated as None):', failed)
print('Missing but calculated:', missing)
