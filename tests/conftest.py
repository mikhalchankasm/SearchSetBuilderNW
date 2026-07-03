# -*- coding: utf-8 -*-
"""Test setup: make the project root importable and force a headless Qt platform."""
import os
import sys

# Force offscreen Qt BEFORE any PySide6 import (GUI tests run without a display).
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import pytest  # noqa: E402


@pytest.fixture(scope="session")
def qapp():
    """A single QApplication for the whole GUI test session."""
    pytest.importorskip("PySide6")
    from PySide6 import QtWidgets
    return QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
