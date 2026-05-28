"""Tests for workbook structure and data integrity."""

from pathlib import Path

import pytest
from openpyxl import load_workbook


REPO_ROOT = Path(__file__).resolve().parent.parent

EXPECTED_SHEETS = {
    "Loan amortization.xlsm": ["LoanAmor"],
    "Mean variance portofolio analysis Github.xlsm": ["PortAnalysis"],
    "VBA project.xlsm": [
        "ans_1",
        "ans_1_appendix",
        "ans_2",
        "ans2_graph",
        "ans_3",
        "ans_3spread",
        "price_data",
        "ret_data",
        "bestScenario",
        "worstScenario",
        "flatScenario",
    ],
}


class TestWorkbookStructure:
    """Verify workbook structure and sheet presence."""

    def test_workbook_opens(self, workbook_path: Path) -> None:
        wb = load_workbook(str(workbook_path), keep_vba=True)
        assert len(wb.sheetnames) > 0
        wb.close()

    @pytest.mark.parametrize("wb_name", list(EXPECTED_SHEETS.keys()))
    def test_expected_sheets(self, wb_name: str) -> None:
        wb_path = REPO_ROOT / wb_name
        wb = load_workbook(str(wb_path), keep_vba=True)
        for sheet_name in EXPECTED_SHEETS[wb_name]:
            assert sheet_name in wb.sheetnames, (
                f"Sheet '{sheet_name}' missing from {wb_name}"
            )
        wb.close()

    def test_vba_project_has_price_data(self) -> None:
        """The VBA project workbook should contain price data for analysis."""
        wb_path = REPO_ROOT / "VBA project.xlsm"
        wb = load_workbook(str(wb_path), keep_vba=True)
        ws = wb["price_data"]
        assert ws.max_row is not None and ws.max_row > 10, (
            "price_data sheet should contain substantial data"
        )
        wb.close()

    def test_loan_amortization_sheet_structure(self) -> None:
        """LoanAmor sheet should have expected header layout."""
        wb_path = REPO_ROOT / "Loan amortization.xlsm"
        wb = load_workbook(str(wb_path), keep_vba=True)
        ws = wb["LoanAmor"]
        assert ws.max_column is not None and ws.max_column >= 10, (
            "LoanAmor sheet should have at least 10 columns for amortization tables"
        )
        wb.close()


class TestSourceCodeConsistency:
    """Verify .txt source files are consistent with .xlsm VBA code."""

    def test_loan_amortization_txt_matches_xlsm(self) -> None:
        """Loan amortization VBA code .txt should contain key functions from the workbook."""
        txt_path = REPO_ROOT / "Loan amortization VBA code.txt"
        assert txt_path.exists(), "Loan amortization VBA code.txt not found"
        txt_content = txt_path.read_text(encoding="utf-8", errors="replace")
        assert "Sub complete_Click" in txt_content or "Private Sub complete_Click" in txt_content
        assert "Pmt(" in txt_content or "PMT(" in txt_content or "pmt(" in txt_content

    def test_portfolio_txt_matches_xlsm(self) -> None:
        """Portfolio VBA code .txt should contain key functions from the workbook."""
        txt_path = REPO_ROOT / "Portfolio mean variance analysis VBA code.txt"
        assert txt_path.exists(), "Portfolio mean variance analysis VBA code.txt not found"
        txt_content = txt_path.read_text(encoding="utf-8", errors="replace")
        assert "Function getMonthlyPortLogRet" in txt_content or "Public Function getMonthlyPortLogRet" in txt_content
        assert "sharpe_ratio" in txt_content
        assert "portfolio_variance" in txt_content
