"""Shared fixtures for VBA financial modeling tests."""

from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent

WORKBOOKS = {
    "loan": REPO_ROOT / "Loan amortization.xlsm",
    "portfolio": REPO_ROOT / "Mean variance portofolio analysis Github.xlsm",
    "vba_project": REPO_ROOT / "VBA project.xlsm",
}


@pytest.fixture
def repo_root() -> Path:
    return REPO_ROOT


@pytest.fixture(params=list(WORKBOOKS.keys()))
def workbook_path(request: pytest.FixtureRequest) -> Path:
    return WORKBOOKS[request.param]
