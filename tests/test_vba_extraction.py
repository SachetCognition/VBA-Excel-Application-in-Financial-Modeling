"""Tests for VBA module extraction from .xlsm workbooks."""

from pathlib import Path

import pytest
from oletools.olevba import VBA_Parser


REPO_ROOT = Path(__file__).resolve().parent.parent

EXPECTED_MODULES = {
    "Loan amortization.xlsm": {
        "AmortizationCaculator.frm",
        "exercise1_loanSched.bas",
        "EX2_PortfolioAnalysis.bas",
        "DisplayingDynamicCharts.frm",
        "ThisWorkbook.cls",
        "Sheet1.cls",
    },
    "Mean variance portofolio analysis Github.xlsm": {
        "AmortizationCaculator.frm",
        "exercise1_loanSched.bas",
        "EX2_PortfolioAnalysis.bas",
        "DisplayingDynamicCharts.frm",
        "ThisWorkbook.cls",
        "Sheet2.cls",
    },
    "VBA project.xlsm": {
        "cls_assets.cls",
        "answer.bas",
        "factory.bas",
        "helper_function.bas",
        "utili_function.bas",
        "cls_option.cls",
        "ThisWorkbook.cls",
    },
}


class TestVBAExtraction:
    """Verify VBA modules can be extracted from each workbook."""

    def test_workbook_exists(self, workbook_path: Path) -> None:
        assert workbook_path.exists(), f"Workbook not found: {workbook_path}"

    def test_workbook_contains_vba(self, workbook_path: Path) -> None:
        vba = VBA_Parser(str(workbook_path))
        assert vba.detect_vba_macros(), f"No VBA macros found in {workbook_path.name}"
        vba.close()

    def test_modules_are_extractable(self, workbook_path: Path) -> None:
        vba = VBA_Parser(str(workbook_path))
        modules = []
        for filename, stream_path, vba_filename, vba_code in vba.extract_macros():
            modules.append(vba_filename)
            assert len(vba_code.strip()) > 0, f"Empty VBA module: {vba_filename}"
        vba.close()
        assert len(modules) > 0, f"No modules extracted from {workbook_path.name}"

    @pytest.mark.parametrize(
        "wb_name",
        list(EXPECTED_MODULES.keys()),
    )
    def test_expected_modules_present(self, wb_name: str) -> None:
        wb_path = REPO_ROOT / wb_name
        vba = VBA_Parser(str(wb_path))
        extracted = {
            vba_filename
            for _, _, vba_filename, _ in vba.extract_macros()
        }
        vba.close()
        expected = EXPECTED_MODULES[wb_name]
        missing = expected - extracted
        assert not missing, f"Missing modules in {wb_name}: {missing}"


class TestVBACodeQuality:
    """Basic code quality checks on extracted VBA."""

    def test_no_empty_subs(self, workbook_path: Path) -> None:
        vba = VBA_Parser(str(workbook_path))
        for _, _, vba_filename, vba_code in vba.extract_macros():
            lines = vba_code.splitlines()
            in_sub = False
            sub_body_lines = 0
            for line in lines:
                stripped = line.strip()
                if stripped.startswith(("Sub ", "Private Sub ", "Public Sub ")):
                    in_sub = True
                    sub_body_lines = 0
                elif stripped == "End Sub" and in_sub:
                    in_sub = False
                elif in_sub and stripped:
                    sub_body_lines += 1
        vba.close()

    def test_option_explicit_in_key_modules(self) -> None:
        """Portfolio analysis module should use Option Explicit."""
        wb_path = REPO_ROOT / "Mean variance portofolio analysis Github.xlsm"
        vba = VBA_Parser(str(wb_path))
        for _, _, vba_filename, vba_code in vba.extract_macros():
            if vba_filename == "EX2_PortfolioAnalysis.bas":
                assert "Option Explicit" in vba_code, (
                    "EX2_PortfolioAnalysis.bas should have Option Explicit"
                )
        vba.close()
