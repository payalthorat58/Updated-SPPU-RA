from pathlib import Path
from parser_engine import ResultParsingEngine
from excel_writer import ExcelWriter

pdf_path = r'C:\Users\Payal\OneDrive\Documents\Result Analyzer Project\result analyzer testing pdfs\BE IT CEGP011520 (28).pdf'
result = ResultParsingEngine().parse(Path(pdf_path))
all_year_fields = {
    k for student in result.students for k in student.summary.keys()
    if k.endswith(' SGPA') and not k.startswith('Semester')
}

for student in result.students[:5]:
    print('---', student.name, '---')
    print('Original Summary:', student.summary)
    calculated = ExcelWriter._missing_semester_sgpas(student, all_year_fields)
    print('Calculated SGPA:', calculated)
    for sem, subjects in student.semesters.items():
        for label, subject in subjects.items():
            print(f"{label}: Grd={subject.fields.get('Grd')}, Crd={subject.fields.get('Crd')}, GP={subject.fields.get('GP')}")
