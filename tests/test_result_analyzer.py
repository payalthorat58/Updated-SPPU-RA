"""Tests for ResultAnalyzer module."""

from pathlib import Path
import pytest
from result_analyzer import ResultAnalyzer


ROOT = Path(__file__).resolve().parent.parent


def test_result_analyzer_generates_abc():
    excel_path = ROOT / "output" / "BE Computer CEGPresult.xlsx"
    if not excel_path.exists():
        pytest.skip("BE Computer CEGPresult.xlsx does not exist in output/")

    analyzer = ResultAnalyzer()
    out_dir = ROOT / "output"
    results = analyzer.generate_abc(
        excel_path=excel_path,
        output_dir=out_dir,
        pdf_path=ROOT / "pdfs" / "BE Computer CEGPresult.pdf",
    )

    assert "A" in results and results["A"].exists() and results["A"].stat().st_size > 0
    assert "B" in results and results["B"].exists() and results["B"].stat().st_size > 0
    assert "C" in results and results["C"].exists() and results["C"].stat().st_size > 0


def test_result_analyzer_calculations():
    excel_path = ROOT / "output" / "BE Computer CEGPresult.xlsx"
    if not excel_path.exists():
        pytest.skip("BE Computer CEGPresult.xlsx does not exist in output/")

    analyzer = ResultAnalyzer()
    data = analyzer.analyze_excel(excel_path)

    # Proforma A assertions for 70 students BE Computer
    pa = data["proforma_a"]
    assert pa["registered"] == 70
    assert pa["appeared"] == 70
    assert pa["dist"] == 52
    assert pa["fc"] == 9
    assert pa["total_passed_wo"] == 61
    assert pa["total_with_atkt"] == 70

    # Proforma B assertions
    pb = data["proforma_b"]
    assert len(pb) == 15  # 15 credit subjects (8 Sem 7, 7 Sem 8)
    for subj in pb:
        assert subj["appeared"] == 70
        assert subj["total_passed"] == 70
        assert subj["pct_str"] == "100%"

    # Proforma C assertions
    pc = data["proforma_c"]
    assert len(pc) == 3
    assert pc[0]["rank"] == "1"
    assert pc[0]["name"] == "HAMBIRE VAISHNAVI DHANANJAY"
    assert pc[0]["sgpa"] == "10.00"
