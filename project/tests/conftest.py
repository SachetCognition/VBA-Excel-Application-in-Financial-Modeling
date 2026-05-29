"""Shared fixtures for the financial modeling test suite."""

import sys
import os
import pytest
import numpy as np

# Add project directories to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "before"))
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "after"))


@pytest.fixture
def sample_portfolio_data():
    """Generate 60-month synthetic portfolio data (deterministic)."""
    np.random.seed(42)
    n = 61
    asset1 = 100 * np.exp(np.cumsum(np.random.normal(0.008, 0.05, n)))
    asset2 = 100 * np.exp(np.cumsum(np.random.normal(0.005, 0.07, n)))
    rf = 100 * np.exp(np.cumsum(np.random.normal(0.002, 0.005, n)))
    return asset1, asset2, rf
