from pathlib import Path
from parser_engine import ResultParsingEngine
from excel_writer import ExcelWriter

result = ResultParsingEngine().parse(Path(r'pdfs\BE Computer CEGPresult.pdf'))
for student in result.students:
    if 'JADHAV' in student.name:
        print('---', repr(student.name), '---')
        print('Calculated SGPA:', ExcelWriter._missing_semester_sgpas(student))
        for sem, subjects in student.semesters.items():
            for label, subject in subjects.items():
                print(f"{label}: Grd={subject.fields.get('Grd')}, Crd={subject.fields.get('Crd')}, GP={subject.fields.get('GP')}")
