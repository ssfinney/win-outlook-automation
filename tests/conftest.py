"""
Pytest configuration: mock Windows-only COM modules and redirect all
file I/O to a temporary directory before outlook_triage is imported.
"""
import os
import sys
import tempfile
from unittest.mock import MagicMock

# Ensure imports work when pytest is invoked from outside repo root.
_repo_root = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
if _repo_root not in sys.path:
    sys.path.insert(0, _repo_root)

# --- Mock Windows-only modules before any test module imports outlook_triage ---
for _name in ("win32com", "win32com.client", "pythoncom"):
    if _name not in sys.modules:
        sys.modules[_name] = MagicMock()

# --- Redirect OneDrive-based paths to a tmp directory ---
_tmpdir = tempfile.mkdtemp(prefix="ot_test_")
os.environ["OneDrive"] = _tmpdir
