"""Result Analyzer module to generate Proforma A, B, and C documents.

Processes generated Excel workbooks and generates Proforma A, B, and C as downloadable PDFs.
"""

from __future__ import annotations

import os
from pathlib import Path
import re
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle


class ResultAnalyzer:
    def __init__(self, abc_docs_dir: Path | str | None = None):
        if abc_docs_dir is None:
            self.abc_docs_dir = Path(__file__).resolve().parent / "ABC docs"
        else:
            self.abc_docs_dir = Path(abc_docs_dir)

    def analyze_excel(self, excel_path: Path, pdf_path: Path | None = None) -> Dict[str, Any]:
        """Read generated Excel file and calculate Proforma A, B, and C data."""
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        sheet_name = "Results" if "Results" in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]

        # Forward fill headers across columns
        col_count = ws.max_column
        row1 = [ws.cell(1, c).value for c in range(1, col_count + 1)]
        row2 = [ws.cell(2, c).value for c in range(1, col_count + 1)]
        row3 = [ws.cell(3, c).value for c in range(1, col_count + 1)]

        curr_sem = None
        curr_sub = None
        subjects: Dict[str, Dict[str, Any]] = {}
        info_cols: Dict[str, int] = {}

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
            else:
                if curr_sub:
                    if curr_sub not in subjects:
                        subjects[curr_sub] = {"sem": curr_sem, "fields": {}}
                    if f:
                        subjects[curr_sub]["fields"][str(f).strip()] = c

        # Find Seat, PRN, Name, SGPA, CGPA column indices
        seat_col = info_cols.get("Seat Number", 1)
        prn_col = info_cols.get("PRN", 2)
        name_col = info_cols.get("Student Name", 3)
        mother_col = info_cols.get("Mother Name", 4)

        # Detect SGPA and CGPA columns
        sgpa_col = None
        cgpa_col = None
        for col_name, c_idx in info_cols.items():
            if "SGPA" in col_name:
                sgpa_col = c_idx
            elif "CGPA" in col_name:
                cgpa_col = c_idx

        if sgpa_col is None:
            sgpa_col = 5
        if cgpa_col is None:
            cgpa_col = 6

        # Extract student records
        students = []
        max_row = ws.max_row
        for r in range(4, max_row + 1):
            seat = ws.cell(r, seat_col).value
            if not seat:
                continue
            seat = str(seat).strip()
            prn = str(ws.cell(r, prn_col).value or "").strip()
            name = str(ws.cell(r, name_col).value or "").strip()
            mother = str(ws.cell(r, mother_col).value or "").strip()
            sgpa_val = ws.cell(r, sgpa_col).value
            cgpa_val = ws.cell(r, cgpa_col).value

            sgpa_str = str(sgpa_val).strip() if sgpa_val is not None else "--"
            cgpa_str = str(cgpa_val).strip() if cgpa_val is not None else ""

            # Check marks obtained and max
            tot_obt = 0
            tot_max = 0
            tot_cp = 0
            student_subject_grades: Dict[str, str] = {}

            for sub_name, s_data in subjects.items():
                f_map = s_data["fields"]
                # Credit Points
                if "CP" in f_map:
                    cp_val = ws.cell(r, f_map["CP"]).value
                    try:
                        tot_cp += int(cp_val)
                    except (ValueError, TypeError):
                        pass

                # Grade
                if "Grd" in f_map:
                    g_val = ws.cell(r, f_map["Grd"]).value
                    if g_val:
                        student_subject_grades[sub_name] = str(g_val).strip()

                # Marks
                if "TOTAL" in f_map:
                    v = str(ws.cell(r, f_map["TOTAL"]).value or "")
                    if "/" in v:
                        p = v.split("/")
                        if p[0].isdigit() and p[1].isdigit():
                            tot_obt += int(p[0])
                            tot_max += int(p[1])
                else:
                    for comp in ["TW", "PR", "OR", "TUT"]:
                        if comp in f_map:
                            v = str(ws.cell(r, f_map[comp]).value or "")
                            if "/" in v:
                                p = v.split("/")
                                if p[0].isdigit() and p[1].isdigit():
                                    tot_obt += int(p[0])
                                    tot_max += int(p[1])

            students.append({
                "seat": seat,
                "prn": prn,
                "name": name,
                "mother": mother,
                "sgpa": sgpa_str,
                "cgpa": cgpa_str,
                "obt": tot_obt,
                "max": tot_max,
                "cp": tot_cp,
                "grades": student_subject_grades,
            })

        # Metadata extraction
        metadata = self._extract_metadata(excel_path, pdf_path, ws, students)

        # Proforma A: Class Distribution Calculation
        proforma_a_data = self._calculate_proforma_a(students, metadata)

        # Proforma B: Subject-wise Calculation
        proforma_b_data = self._calculate_proforma_b(ws, subjects, len(students))

        # Proforma C: Topper List Calculation
        proforma_c_data = self._calculate_proforma_c(students)

        return {
            "metadata": metadata,
            "proforma_a": proforma_a_data,
            "proforma_b": proforma_b_data,
            "proforma_c": proforma_c_data,
        }

    def _extract_metadata(
        self,
        excel_path: Path,
        pdf_path: Path | None,
        ws: openpyxl.worksheet.worksheet.Worksheet,
        students: List[Dict[str, Any]],
    ) -> Dict[str, str]:
        stem = excel_path.stem.upper()
        inst = "B.V.C.O.E.W., Pune-43"
        inst_full = "B.V.C.O.E.FOR WOMEN, PUNE.-43."
        dept = "COMP"
        class_year = "BE (Computer)"
        class_year_full = "BE (Computer Engineering)"
        class_code = "BE COMP"
        exam = "May 2026"
        decl_date = "16/07/2026"
        subm_date = "31/08/2026"
        acad_year = "2025-2026"

        # Detect class from stem or students
        if "BE" in stem or (students and students[0]["seat"].startswith("B")):
            class_year = "BE (Computer)" if "COMP" in stem else "BE"
            class_year_full = "BE (Computer Engineering)" if "COMP" in stem else "BE"
            class_code = "BE COMP" if "COMP" in stem else "BE"
            dept = "COMP" if "COMP" in stem else "ENGG"
        elif "TE" in stem or (students and students[0]["seat"].startswith("T")):
            class_year = "TE (IT)" if "IT" in stem else "TE (Computer)"
            class_year_full = "TE (Information Technology)" if "IT" in stem else "TE"
            class_code = "TE IT" if "IT" in stem else "TE COMP"
            dept = "IT" if "IT" in stem else "COMP"
        elif "SE" in stem or (students and students[0]["seat"].startswith("S")):
            class_year = "SE (Computer)"
            class_year_full = "SE (Computer Engineering)"
            class_code = "SE COMP"
            dept = "COMP"

        if "IT" in stem:
            dept = "IT"
            class_code = class_code.replace("COMP", "IT")
            class_year = class_year.replace("Computer", "IT")

        # Extract from PDF if available
        if pdf_path and pdf_path.exists():
            try:
                import pdfplumber
                with pdfplumber.open(pdf_path) as pdf:
                    if len(pdf.pages) > 0:
                        first_text = pdf.pages[0].extract_text() or ""
                        # Date match
                        d_match = re.search(r"DATE\s*:\s*(\d{1,2}\s+[A-Z]{3}\s+\d{4})", first_text)
                        if d_match:
                            raw_d = d_match.group(1)
                            # Convert e.g. 16 JUL 2026 to 16/07/2026
                            months = {
                                "JAN": "01", "FEB": "02", "MAR": "03", "APR": "04",
                                "MAY": "05", "JUN": "06", "JUL": "07", "AUG": "08",
                                "SEP": "09", "OCT": "10", "NOV": "11", "DEC": "12",
                            }
                            parts = raw_d.split()
                            if len(parts) == 3 and parts[1] in months:
                                decl_date = f"{int(parts[0]):02d}/{months[parts[1]]}/{parts[2]}"

                        # Exam session
                        s_match = re.search(r"(SUMMER|WINTER)\s+SESSION[- ]*(\d{4})", first_text)
                        if s_match:
                            season = s_match.group(1)
                            year = s_match.group(2)
                            month_str = "May" if season == "SUMMER" else "Dec"
                            exam = f"{month_str} {year}"

                        # Branch match
                        b_match = re.search(r"BRANCH CODE:\s*\d+-[A-Z\.\s\(\)]+\(([A-Z\s]+)\)", first_text)
                        if b_match:
                            b_str = b_match.group(1).strip()
                            if "COMPUTER" in b_str:
                                dept = "COMP"
                                class_year = f"{class_code.split()[0]} (Computer)"
                                class_year_full = f"{class_code.split()[0]} (Computer Engineering)"
                            elif "INFORMATION" in b_str or "IT" in b_str:
                                dept = "IT"
                                class_year = f"{class_code.split()[0]} (IT)"
                                class_year_full = f"{class_code.split()[0]} (Information Technology)"
            except Exception:
                pass

        return {
            "institution": inst,
            "institution_full": inst_full,
            "dept": dept,
            "class_year": class_year,
            "class_year_full": class_year_full,
            "class_code": class_code,
            "exam": exam,
            "decl_date": decl_date,
            "subm_date": subm_date,
            "academic_year": acad_year,
            "ref_code": f"BV / COEW/ {dept} /              / {acad_year}",
        }

    def _calculate_proforma_a(self, students: List[Dict[str, Any]], metadata: Dict[str, str]) -> Dict[str, Any]:
        registered = len(students)
        appeared = len(students)

        dist = 0
        fc = 0
        hsc = 0
        sc = 0
        pass_class = 0
        atkt_count = sum(1 for student in students if not str(student.get("cgpa") or "").strip())

        for s in students:
            cgpa_val = None
            if s["cgpa"]:
                try:
                    cgpa_val = float(s["cgpa"])
                except ValueError:
                    pass

            if cgpa_val is not None:
                if cgpa_val >= 7.75:
                    dist += 1
                elif cgpa_val >= 6.75:
                    fc += 1
                elif cgpa_val >= 6.25:
                    hsc += 1
                elif cgpa_val >= 5.50:
                    sc += 1
                elif cgpa_val >= 5.00:
                    pass_class += 1

        total_passed_wo = registered - atkt_count
        total_with_atkt = total_passed_wo + atkt_count

        pct_wo = (total_passed_wo / appeared * 100) if appeared > 0 else 0.0
        pct_with = (total_with_atkt / appeared * 100) if appeared > 0 else 0.0

        hsc_sc_val = (hsc + sc) if (hsc + sc) > 0 else "--"
        pass_val = pass_class if pass_class > 0 else "--"
        with_atkt_str = f"{total_passed_wo}+{atkt_count}={total_with_atkt}"

        return {
            "class_code": metadata["class_code"],
            "registered": registered,
            "appeared": appeared,
            "dist": dist,
            "fc": fc,
            "hsc_sc": hsc_sc_val,
            "pass_class": pass_val,
            "total_passed_wo": total_passed_wo,
            "with_atkt_str": with_atkt_str,
            "total_with_atkt": total_with_atkt,
            "pct_wo": pct_wo,
            "pct_with": pct_with,
        }

    def _calculate_proforma_b(
        self,
        ws: openpyxl.worksheet.worksheet.Worksheet,
        subjects: Dict[str, Dict[str, Any]],
        student_count: int,
    ) -> List[Dict[str, Any]]:
        subject_rows = []
        sr_no = 1

        for sub_name, data in subjects.items():
            f_map = data["fields"]
            grd_col = f_map.get("Grd")
            crd_col = f_map.get("Crd")
            if not grd_col:
                continue

            # Skip audit courses (0 credits or PP/NP)
            crd_val = str(ws.cell(4, crd_col).value or "").strip() if crd_col else ""
            if crd_val in ["00", "0", ""]:
                continue

            counts = {"O": 0, "A+": 0, "A": 0, "B+": 0, "B": 0, "C": 0, "P": 0, "F": 0, "FF": 0}
            appeared_students = 0

            for r in range(4, ws.max_row + 1):
                g = ws.cell(r, grd_col).value
                if g is not None:
                    g_str = str(g).strip().upper()
                    if g_str in counts:
                        counts[g_str] += 1
                        appeared_students += 1

            if appeared_students == 0:
                appeared_students = student_count

            dist = counts["O"] + counts["A+"]
            fc = counts["A"]
            hsc = counts["B+"]
            sc = counts["B"]
            pc = counts["C"] + counts["P"]
            total_passed = dist + fc + hsc + sc + pc

            pct = (total_passed / appeared_students * 100) if appeared_students > 0 else 0.0
            pct_str = f"{pct:.2f}%" if round(pct, 2) != 100.0 else "100%"
            hsc_sc_str = f"{hsc:02d}+{sc:02d}={hsc+sc:02d}"

            # Subject title clean-up: strip leading subject code numbers (e.g. 410241)
            clean_name = re.sub(r"^\d+[A-Z]?\s*[-–]?\s*", "", sub_name).strip()
            # If subject has TW/PR/OR mention
            if "PR" in f_map:
                clean_name += " (PR)"
            elif "OR" in f_map:
                clean_name += " (OR)"
            elif "TW" in f_map and "TOTAL" not in f_map:
                clean_name += " (TW)"

            subject_rows.append({
                "sr_no": f"{sr_no:02d}",
                "subject": clean_name,
                "appeared": appeared_students,
                "dist": dist,
                "fc": fc,
                "hsc_sc_str": hsc_sc_str,
                "pass_class": f"{pc:02d}",
                "total_passed": total_passed,
                "pct_str": pct_str,
                "univ_pct": "",
                "staff_name": "",
                "remarks": "",
            })
            sr_no += 1

        return subject_rows

    def _calculate_proforma_c(self, students: List[Dict[str, Any]], top_n: int = 3) -> List[Dict[str, Any]]:
        valid = [s for s in students if s["sgpa"] and s["sgpa"] != "--"]

        def sort_key(s):
            try:
                sgpa_f = float(s["sgpa"])
            except ValueError:
                sgpa_f = 0.0
            return (sgpa_f, s["cp"], s["obt"])

        valid.sort(key=sort_key, reverse=True)

        toppers = []
        for idx, s in enumerate(valid[:top_n], 1):
            toppers.append({
                "rank": str(idx),
                "seat": s["seat"],
                "name": s["name"],
                "marks": f"{s['obt']}/{s['max']}",
                "cp": str(s["cp"]),
                "sgpa": str(s["sgpa"]),
            })
        return toppers

    def generate_abc(
        self,
        excel_path: Path | str,
        output_dir: Path | str | None = None,
        pdf_path: Path | str | None = None,
    ) -> Dict[str, Path]:
        """Generate A.pdf, B.pdf, C.pdf (and xlsx/docx) from an Excel file."""
        excel_path = Path(excel_path)
        if pdf_path is not None:
            pdf_path = Path(pdf_path)

        if output_dir is None:
            output_dir = excel_path.parent
        else:
            output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        data = self.analyze_excel(excel_path, pdf_path)

        # 1. Generate Proforma A
        path_a_pdf = self.generate_proforma_a(data, output_dir)

        # 2. Generate Proforma B
        path_b_pdf = self.generate_proforma_b(data, output_dir)

        # 3. Generate Proforma C
        path_c_pdf = self.generate_proforma_c(data, output_dir)

        return {
            "A": path_a_pdf,
            "B": path_b_pdf,
            "C": path_c_pdf,
        }

    # =========================================================================
    # PROFORMA A GENERATION
    # =========================================================================
    def generate_proforma_a(self, data: Dict[str, Any], output_dir: Path) -> Path:
        xlsx_template = self.abc_docs_dir / "TE IT New A.xlsx"
        target_xlsx = output_dir / "A.xlsx"
        target_pdf = output_dir / "A.pdf"

        meta = data["metadata"]
        pa = data["proforma_a"]

        if xlsx_template.exists():
            wb = openpyxl.load_workbook(xlsx_template)
            ws = wb.active

            # Fix title encoding
            ws["A3"] = "PROFORMA - A"

            # Update headers
            ws["J1"] = meta["ref_code"]
            ws["A6"] = f"Name of the Institution: {meta['institution_full']}"
            ws["A7"] = f"Class / Year: {meta['class_year_full']}"
            ws["A8"] = f"Exam (Month & Year): {meta['exam']}"
            ws["A9"] = f"Date of Declaration:  {meta['decl_date']}"
            ws["A10"] = f"Date of submission of result Analysis: - {meta['subm_date']}"

            # Update Row 15
            ws["A15"] = pa["class_code"]
            ws["B15"] = pa["registered"]
            ws["C15"] = pa["appeared"]
            ws["D15"] = pa["dist"]
            ws["E15"] = pa["fc"]
            ws["F15"] = pa["hsc_sc"]
            ws["G15"] = pa["pass_class"]
            ws["H15"] = pa["total_passed_wo"]
            ws["I15"] = pa["with_atkt_str"]
            ws["J15"] = round(pa["pct_wo"] / 100, 4)
            ws["K15"] = round(pa["pct_with"] / 100, 4)

            wb.save(target_xlsx)

            # Try Excel COM export to PDF
            exported = self._export_excel_to_pdf_com(target_xlsx, target_pdf)
            if exported and target_pdf.exists():
                return target_pdf

        # Fallback to ReportLab PDF generation
        self._generate_reportlab_proforma_a(data, target_pdf)
        return target_pdf

    # =========================================================================
    # PROFORMA B GENERATION
    # =========================================================================
    def generate_proforma_b(self, data: Dict[str, Any], output_dir: Path) -> Path:
        docx_template = self.abc_docs_dir / "B.docx"
        target_docx = output_dir / "B.docx"
        target_pdf = output_dir / "B.pdf"

        meta = data["metadata"]
        pb = data["proforma_b"]

        if docx_template.exists():
            doc = docx.Document(docx_template)

            # Update paragraphs
            for p in doc.paragraphs:
                txt = p.text
                if "Class/Year" in txt:
                    p.text = f"Class/Year\t\t\t \t\t:  {meta['class_year']}"
                elif "Exam (Month & Year)" in txt:
                    p.text = f"Exam (Month & Year)\t\t\t:  {meta['exam']}"
                elif "Date of result declaration" in txt:
                    p.text = f"Date of result declaration\t\t\t:  {meta['decl_date']}\t"
                elif "Date of submission of result analysis" in txt:
                    p.text = f"Date of submission of result analysis \t:  {meta['subm_date']}"
                elif "HOD" in txt:
                    p.text = f"HOD {meta['dept']}  \t\t\t\t\t\t\t\t\t\t\t\tSignature of Principal/Director"

            # Update table
            if doc.tables:
                t = doc.tables[0]
                # Remove old data rows (keep header rows 0, 1, 2)
                while len(t.rows) > 3:
                    t._tbl.remove(t.rows[3]._tr)

                # Add new rows
                for r_data in pb:
                    row = t.add_row()
                    vals = [
                        r_data["sr_no"],
                        r_data["subject"],
                        str(r_data["appeared"]),
                        str(r_data["dist"]),
                        str(r_data["fc"]),
                        r_data["hsc_sc_str"],
                        r_data["pass_class"],
                        str(r_data["total_passed"]),
                        r_data["pct_str"],
                        r_data["univ_pct"],
                        r_data["staff_name"],
                        r_data["remarks"],
                    ]
                    for idx, val in enumerate(vals):
                        cell = row.cells[idx]
                        cell.text = val
                        p = cell.paragraphs[0]
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        if idx == 1:
                            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        if p.runs:
                            p.runs[0].font.size = Pt(9.5)
                            p.runs[0].font.name = "Times New Roman"

            doc.save(target_docx)

            # Try Word COM export to PDF
            exported = self._export_word_to_pdf_com(target_docx, target_pdf)
            if exported and target_pdf.exists():
                return target_pdf

        # Fallback to ReportLab PDF generation
        self._generate_reportlab_proforma_b(data, target_pdf)
        return target_pdf

    # =========================================================================
    # PROFORMA C GENERATION
    # =========================================================================
    def generate_proforma_c(self, data: Dict[str, Any], output_dir: Path) -> Path:
        docx_template = self.abc_docs_dir / "C.docx"
        target_docx = output_dir / "C.docx"
        target_pdf = output_dir / "C.pdf"

        meta = data["metadata"]
        pc = data["proforma_c"]

        if docx_template.exists():
            doc = docx.Document(docx_template)

            # Update paragraphs
            for p in doc.paragraphs:
                txt = p.text
                if "Class/Year" in txt:
                    p.text = f"Class/Year\t\t\t \t\t:  {meta['class_year']}"
                elif "Exam (Month & Year)" in txt:
                    p.text = f"Exam (Month & Year)\t\t\t:  {meta['exam']}"
                elif "Date of result declaration" in txt:
                    p.text = f"Date of result declaration\t\t\t:  {meta['decl_date']}\t"
                elif "Date of submission of result analysis" in txt:
                    p.text = f"Date of submission of result analysis \t:  {meta['subm_date']}"
                elif "HOD" in txt:
                    p.text = f"HOD {meta['dept']}  \t\t\t\t\t\t\t\t\t\t\t\tSignature of Principal/Director"

            # Update table
            if doc.tables:
                t = doc.tables[0]
                # Remove old data rows (keep header row 0)
                while len(t.rows) > 1:
                    t._tbl.remove(t.rows[1]._tr)

                # Add topper rows
                for r_data in pc:
                    row = t.add_row()
                    vals = [
                        r_data["rank"],
                        r_data["seat"],
                        r_data["name"],
                        r_data["marks"],
                        r_data["cp"],
                        r_data["sgpa"],
                    ]
                    for idx, val in enumerate(vals):
                        cell = row.cells[idx]
                        cell.text = val
                        p = cell.paragraphs[0]
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        if idx == 2:
                            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        if p.runs:
                            p.runs[0].font.size = Pt(10)
                            p.runs[0].font.name = "Times New Roman"

            doc.save(target_docx)

            # Try Word COM export to PDF
            exported = self._export_word_to_pdf_com(target_docx, target_pdf)
            if exported and target_pdf.exists():
                return target_pdf

        # Fallback to ReportLab PDF generation
        self._generate_reportlab_proforma_c(data, target_pdf)
        return target_pdf

    # =========================================================================
    # COM EXPORT HELPERS (WINDOWS OFFICE)
    # =========================================================================
    @staticmethod
    def _export_word_to_pdf_com(docx_path: Path, pdf_path: Path) -> bool:
        if os.name != "nt":
            return False
        import gc

        word = None
        doc = None
        co_init = False
        try:
            import pythoncom
            pythoncom.CoInitialize()
            co_init = True
            import win32com.client
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(str(docx_path.resolve()))
            # 17 = wdExportFormatPDF
            doc.ExportAsFixedFormat(str(pdf_path.resolve()), 17)
            return True
        except Exception:
            return False
        finally:
            try:
                if doc is not None:
                    doc.Close(False)
            except Exception:
                pass
            doc = None
            try:
                if word is not None:
                    word.Quit()
            except Exception:
                pass
            word = None
            gc.collect()
            if co_init:
                try:
                    import pythoncom
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    @staticmethod
    def _export_excel_to_pdf_com(xlsx_path: Path, pdf_path: Path) -> bool:
        if os.name != "nt":
            return False
        excel = None
        wb = None
        co_init = False
        try:
            import pythoncom
            pythoncom.CoInitialize()
            co_init = True
            import win32com.client
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            wb = excel.Workbooks.Open(str(xlsx_path.resolve()))
            # 0 = xlTypePDF
            wb.ExportAsFixedFormat(0, str(pdf_path.resolve()))
            return True
        except Exception:
            return False
        finally:
            try:
                if wb is not None:
                    wb.Close(False)
            except Exception:
                pass
            try:
                if excel is not None:
                    excel.Quit()
            except Exception:
                pass
            if co_init:
                try:
                    import pythoncom
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

    # =========================================================================
    # REPORTLAB STANDALONE PDF GENERATORS (CROSS-PLATFORM & HIGH FIDELITY)
    # =========================================================================
    def _generate_reportlab_proforma_a(self, data: Dict[str, Any], pdf_path: Path) -> None:
        doc = SimpleDocTemplate(
            str(pdf_path),
            pagesize=landscape(A4),
            leftMargin=30,
            rightMargin=30,
            topMargin=25,
            bottomMargin=25,
        )
        meta = data["metadata"]
        pa = data["proforma_a"]

        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            "TitleStyle",
            parent=styles["Heading1"],
            fontName="Helvetica-Bold",
            fontSize=13,
            alignment=1,
            spaceAfter=4,
        )
        subtitle_style = ParagraphStyle(
            "SubTitleStyle",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=11,
            alignment=1,
            spaceAfter=10,
        )
        meta_style = ParagraphStyle(
            "MetaStyle",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=9.5,
            leading=13,
        )
        ref_style = ParagraphStyle(
            "RefStyle",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=9,
            alignment=2,
            spaceAfter=5,
        )

        elements = []
        elements.append(Paragraph(meta["ref_code"], ref_style))
        elements.append(Paragraph("RESULT ANALYSIS", title_style))
        elements.append(Paragraph("PROFORMA - A", subtitle_style))

        # College metadata
        meta_table_data = [
            [Paragraph(f"<b>Name of the Institution:</b> {meta['institution_full']}", meta_style)],
            [Paragraph(f"<b>Class / Year:</b> {meta['class_year_full']}", meta_style)],
            [Paragraph(f"<b>Exam (Month & Year):</b> {meta['exam']}", meta_style)],
            [Paragraph(f"<b>Date of Declaration:</b> {meta['decl_date']}", meta_style)],
            [Paragraph(f"<b>Date of submission of result Analysis:</b> {meta['subm_date']}", meta_style)],
        ]
        meta_table = Table(meta_table_data, colWidths=[780])
        meta_table.setStyle(TableStyle([("BOTTOMPADDING", (0, 0), (-1, -1), 1)]))
        elements.append(meta_table)
        elements.append(Spacer(1, 12))

        # Proforma A Table
        table_style_cell = ParagraphStyle("Cell", fontName="Helvetica", fontSize=8, alignment=1, leading=10)
        table_style_bold = ParagraphStyle("CellB", fontName="Helvetica-Bold", fontSize=8, alignment=1, leading=10)

        t_data = [
            # Row 0
            [
                Paragraph("Class / Year", table_style_bold),
                Paragraph("No. of Candidate Registerd For Exam.", table_style_bold),
                Paragraph("No. of Candidate Actuallay Appeared for Exam", table_style_bold),
                Paragraph("No. of candidates passed", table_style_bold),
                "", "", "",
                Paragraph("Total Passed", table_style_bold),
                "",
                Paragraph("Percentage of Passing", table_style_bold),
                "",
                Paragraph("Board/ University Percentage of Passing", table_style_bold),
                "",
            ],
            # Row 1
            [
                "", "", "",
                Paragraph("I class with Dist.", table_style_bold),
                Paragraph("I class", table_style_bold),
                Paragraph("II class Higher II + II", table_style_bold),
                Paragraph("Pass Class", table_style_bold),
                Paragraph("W/O ATKT", table_style_bold),
                Paragraph("With ATKT (Clear passed plus ATKT students)", table_style_bold),
                Paragraph("W/O ATKT", table_style_bold),
                Paragraph("With ATKT", table_style_bold),
                Paragraph("W/O ATKT", table_style_bold),
                Paragraph("With ATKT", table_style_bold),
            ],
            # Row 2 (numbers)
            [
                Paragraph("1", table_style_bold),
                Paragraph("2", table_style_bold),
                Paragraph("3", table_style_bold),
                Paragraph("4", table_style_bold),
                Paragraph("5", table_style_bold),
                Paragraph("6", table_style_bold),
                Paragraph("7", table_style_bold),
                Paragraph("8", table_style_bold),
                Paragraph("8A", table_style_bold),
                Paragraph("9", table_style_bold),
                Paragraph("9A", table_style_bold),
                Paragraph("10", table_style_bold),
                "",
            ],
            # Row 3 (Values)
            [
                Paragraph(pa["class_code"], table_style_cell),
                Paragraph(str(pa["registered"]), table_style_cell),
                Paragraph(str(pa["appeared"]), table_style_cell),
                Paragraph(str(pa["dist"]), table_style_cell),
                Paragraph(str(pa["fc"]), table_style_cell),
                Paragraph(str(pa["hsc_sc"]), table_style_cell),
                Paragraph(str(pa["pass_class"]), table_style_cell),
                Paragraph(str(pa["total_passed_wo"]), table_style_cell),
                Paragraph(str(pa["with_atkt_str"]), table_style_cell),
                Paragraph(f"{pa['pct_wo']:.2f}%", table_style_cell),
                Paragraph(f"{pa['pct_with']:.2f}%", table_style_cell),
                Paragraph("--", table_style_cell),
                Paragraph("--", table_style_cell),
            ],
        ]

        col_widths = [60, 60, 60, 55, 50, 60, 50, 55, 95, 60, 60, 55, 55]
        t = Table(t_data, colWidths=col_widths)
        t.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
            ("SPAN", (0, 0), (0, 1)),
            ("SPAN", (1, 0), (1, 1)),
            ("SPAN", (2, 0), (2, 1)),
            ("SPAN", (3, 0), (6, 0)),
            ("SPAN", (7, 0), (8, 0)),
            ("SPAN", (9, 0), (10, 0)),
            ("SPAN", (11, 0), (12, 0)),
            ("SPAN", (11, 2), (12, 2)),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BACKGROUND", (0, 0), (-1, 2), colors.HexColor("#F2F4F7")),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        elements.append(t)
        elements.append(Spacer(1, 35))

        # Signatures
        sig_data = [
            [
                Paragraph("<b>Result Analysis Coordinator</b>", meta_style),
                Paragraph(f"<b>HOD {meta['dept']}</b>", meta_style),
                Paragraph("<b>Principal</b>", meta_style),
            ]
        ]
        sig_table = Table(sig_data, colWidths=[280, 280, 220])
        elements.append(sig_table)

        doc.build(elements)

    def _generate_reportlab_proforma_b(self, data: Dict[str, Any], pdf_path: Path) -> None:
        doc = SimpleDocTemplate(
            str(pdf_path),
            pagesize=landscape(A4),
            leftMargin=25,
            rightMargin=25,
            topMargin=20,
            bottomMargin=20,
        )
        meta = data["metadata"]
        pb = data["proforma_b"]

        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            "TitleStyle",
            parent=styles["Heading1"],
            fontName="Helvetica-Bold",
            fontSize=12,
            alignment=0,
        )
        sub_style = ParagraphStyle(
            "SubStyle",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=12,
            alignment=2,
        )
        meta_style = ParagraphStyle(
            "MetaStyle",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=9,
            leading=12,
        )

        header_table = Table([
            [Paragraph("RESULT ANALYSIS", title_style), Paragraph("PROFORMA B", sub_style)]
        ], colWidths=[400, 390])
        header_table.setStyle(TableStyle([("BOTTOMPADDING", (0, 0), (-1, -1), 2)]))

        elements = [header_table, Spacer(1, 4)]

        # Metadata
        meta_data = [
            [Paragraph(f"<b>Name of Institution</b> : {meta['institution']}", meta_style)],
            [Paragraph(f"<b>Class/Year</b> : {meta['class_year']}", meta_style)],
            [Paragraph(f"<b>Exam (Month & Year)</b> : {meta['exam']}", meta_style)],
            [Paragraph(f"<b>Date of result declaration</b> : {meta['decl_date']}", meta_style)],
            [Paragraph(f"<b>Date of submission of result analysis</b> : {meta['subm_date']}", meta_style)],
        ]
        meta_t = Table(meta_data, colWidths=[790])
        meta_t.setStyle(TableStyle([("BOTTOMPADDING", (0, 0), (-1, -1), 1)]))
        elements.append(meta_t)
        elements.append(Spacer(1, 8))

        # Table
        c_style = ParagraphStyle("C", fontName="Helvetica", fontSize=7.5, alignment=1, leading=9)
        c_bold = ParagraphStyle("CB", fontName="Helvetica-Bold", fontSize=7.5, alignment=1, leading=9)
        c_left = ParagraphStyle("CL", fontName="Helvetica", fontSize=7.5, alignment=0, leading=9)

        table_rows = [
            [
                Paragraph("Sr. No.", c_bold),
                Paragraph("Subject", c_bold),
                Paragraph("Actual no. of students appeared", c_bold),
                Paragraph("Total no. of students passed", c_bold),
                "", "", "", "",
                Paragraph("% of passing", c_bold),
                Paragraph("University % of passing", c_bold),
                Paragraph("Name of the staff", c_bold),
                Paragraph("Remarks of the Principal/ Director", c_bold),
            ],
            [
                "", "", "",
                Paragraph("I class with Dist.", c_bold),
                Paragraph("I class", c_bold),
                Paragraph("Higher II class + II class", c_bold),
                Paragraph("Pass class", c_bold),
                Paragraph("Total passed", c_bold),
                "", "", "", "",
            ],
            [
                Paragraph("01", c_bold),
                Paragraph("02", c_bold),
                Paragraph("03", c_bold),
                Paragraph("04", c_bold),
                Paragraph("05", c_bold),
                Paragraph("06", c_bold),
                Paragraph("07", c_bold),
                Paragraph("08", c_bold),
                Paragraph("09", c_bold),
                Paragraph("10", c_bold),
                Paragraph("11", c_bold),
                Paragraph("12", c_bold),
            ],
        ]

        for r in pb:
            table_rows.append([
                Paragraph(r["sr_no"], c_style),
                Paragraph(r["subject"], c_left),
                Paragraph(str(r["appeared"]), c_style),
                Paragraph(str(r["dist"]), c_style),
                Paragraph(str(r["fc"]), c_style),
                Paragraph(r["hsc_sc_str"], c_style),
                Paragraph(r["pass_class"], c_style),
                Paragraph(str(r["total_passed"]), c_style),
                Paragraph(r["pct_str"], c_style),
                Paragraph(r["univ_pct"], c_style),
                Paragraph(r["staff_name"], c_left),
                Paragraph(r["remarks"], c_style),
            ])

        col_w = [30, 180, 55, 45, 40, 65, 40, 45, 45, 55, 100, 90]
        t = Table(table_rows, colWidths=col_w, repeatRows=3)
        t.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
            ("SPAN", (0, 0), (0, 1)),
            ("SPAN", (1, 0), (1, 1)),
            ("SPAN", (2, 0), (2, 1)),
            ("SPAN", (3, 0), (7, 0)),
            ("SPAN", (8, 0), (8, 1)),
            ("SPAN", (9, 0), (9, 1)),
            ("SPAN", (10, 0), (10, 1)),
            ("SPAN", (11, 0), (11, 1)),
            ("BACKGROUND", (0, 0), (-1, 2), colors.HexColor("#F2F4F7")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 2.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
        ]))
        elements.append(t)
        elements.append(Spacer(1, 18))

        # Signatures
        sig_data = [
            [
                Paragraph(f"<b>HOD {meta['dept']}</b>", meta_style),
                Paragraph("<b>Signature of Principal/Director</b>", ParagraphStyle("R", parent=meta_style, alignment=2)),
            ]
        ]
        sig_t = Table(sig_data, colWidths=[400, 390])
        elements.append(sig_t)

        doc.build(elements)

    def _generate_reportlab_proforma_c(self, data: Dict[str, Any], pdf_path: Path) -> None:
        doc = SimpleDocTemplate(
            str(pdf_path),
            pagesize=A4,
            leftMargin=35,
            rightMargin=35,
            topMargin=30,
            bottomMargin=30,
        )
        meta = data["metadata"]
        pc = data["proforma_c"]

        styles = getSampleStyleSheet()
        title_style = ParagraphStyle(
            "TitleStyle",
            parent=styles["Heading1"],
            fontName="Helvetica-Bold",
            fontSize=13,
            alignment=0,
        )
        sub_style = ParagraphStyle(
            "SubStyle",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=13,
            alignment=2,
        )
        meta_style = ParagraphStyle(
            "MetaStyle",
            parent=styles["Normal"],
            fontName="Helvetica",
            fontSize=10,
            leading=14,
        )
        topper_head_style = ParagraphStyle(
            "TopperHead",
            parent=styles["Heading3"],
            fontName="Helvetica-Bold",
            fontSize=12,
            spaceBefore=12,
            spaceAfter=8,
        )

        header_table = Table([
            [Paragraph("RESULT ANALYSIS", title_style), Paragraph("PROFORMA C", sub_style)]
        ], colWidths=[260, 260])
        elements = [header_table, Spacer(1, 8)]

        # Metadata
        meta_data = [
            [Paragraph(f"<b>Name of Institution</b> : {meta['institution']}", meta_style)],
            [Paragraph(f"<b>Class/Year</b> : {meta['class_year']}", meta_style)],
            [Paragraph(f"<b>Exam (Month & Year)</b> : {meta['exam']}", meta_style)],
            [Paragraph(f"<b>Date of result declaration</b> : {meta['decl_date']}", meta_style)],
            [Paragraph(f"<b>Date of submission of result analysis</b> : {meta['subm_date']}", meta_style)],
        ]
        meta_t = Table(meta_data, colWidths=[520])
        meta_t.setStyle(TableStyle([("BOTTOMPADDING", (0, 0), (-1, -1), 2)]))
        elements.append(meta_t)
        elements.append(Paragraph("<b>List of Toppers</b>", topper_head_style))

        # Topper Table
        c_style = ParagraphStyle("C", fontName="Helvetica", fontSize=9, alignment=1, leading=12)
        c_bold = ParagraphStyle("CB", fontName="Helvetica-Bold", fontSize=9, alignment=1, leading=12)
        c_left = ParagraphStyle("CL", fontName="Helvetica", fontSize=9, alignment=0, leading=12)

        table_rows = [
            [
                Paragraph("Rank", c_bold),
                Paragraph("Exam No.", c_bold),
                Paragraph("Name of the student", c_bold),
                Paragraph("Marks obtained", c_bold),
                Paragraph("Credit Points obtained", c_bold),
                Paragraph("Average SGPA", c_bold),
            ]
        ]

        for r in pc:
            table_rows.append([
                Paragraph(r["rank"], c_style),
                Paragraph(r["seat"], c_style),
                Paragraph(r["name"], c_left),
                Paragraph(r["marks"], c_style),
                Paragraph(r["cp"], c_style),
                Paragraph(r["sgpa"], c_style),
            ])

        col_w = [45, 85, 190, 75, 65, 60]
        t = Table(table_rows, colWidths=col_w)
        t.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F2F4F7")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        elements.append(t)
        elements.append(Spacer(1, 40))

        # Signatures
        sig_data = [
            [
                Paragraph(f"<b>HOD {meta['dept']}</b>", meta_style),
                Paragraph("<b>Signature of Principal/Director</b>", ParagraphStyle("R", parent=meta_style, alignment=2)),
            ]
        ]
        sig_t = Table(sig_data, colWidths=[260, 260])
        elements.append(sig_t)

        doc.build(elements)
