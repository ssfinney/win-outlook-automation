"""
Pytest configuration: mock Windows-only COM modules and redirect all
file I/O to a temporary directory before outlook_triage is imported.
"""
import os
import sys
import tempfile
from unittest.mock import MagicMock

# --- Mock Windows-only modules before any test module imports outlook_triage ---
for _name in ("win32com", "win32com.client", "pythoncom"):
    if _name not in sys.modules:
        sys.modules[_name] = MagicMock()

# --- Redirect OneDrive-based paths to a tmp directory ---
_tmpdir = tempfile.mkdtemp(prefix="ot_test_")
os.environ["OneDrive"] = _tmpdir
