# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Project Does

Local-only Windows email triage automation for Microsoft Outlook (Classic/COM). Scans the Inbox, scores emails with a rule-based engine, assigns Outlook categories (`Urgent`/`Action`/`Waiting`/`FYI`/`Noise`), flags items, and exports Excel reports for review and ML training.

## Commands

```bash
# Lint (max line length 120)
flake8 outlook_triage.py train_model.py --max-line-length=120

# Security scan (medium+ severity)
bandit outlook_triage.py train_model.py -ll

# Run all tests
python -m pytest tests/ -v --tb=short

# Run a single test
python -m pytest tests/test_outlook_triage.py::test_score_keywords -v
```

**Windows only (requires Classic Outlook):**
```powershell
python outlook_triage.py     # triage run (DRY_RUN=True by default)
python train_model.py        # train ML model from labeled Excel reports
```

## Architecture

Two main scripts with a shared output directory (default `%OneDrive%\AI_Outlook\`).

**`outlook_triage.py`** — triage engine
1. Opens Outlook via COM, filters Inbox to `DAYS_BACK` days (max `MAX_ITEMS` items)
2. Scores each email via rule-based engine (keyword weights, VIP sender boost, age penalty, noise patterns)
3. Optionally loads `model/triage_model.joblib` for a secondary ML bucket prediction
4. Merges rule and model predictions — rule `Urgent` and `Noise` always win
5. Applies Outlook categories/flags and optionally moves noise (all gated by `DRY_RUN`)
6. Writes `data/inbox_scored_*.csv` and `outputs/triage_report_*.xlsx`

**`train_model.py`** — model trainer
1. Globs all `outputs/triage_report_*.xlsx`, loads rows with a `Label` column filled in
2. Deduplicates by `entry_id` (last-seen-wins), normalizes labels
3. Builds sklearn pipeline: TF-IDF on text fields + StandardScaler on numeric features → LogisticRegression
4. Requires ≥ 50 labeled rows; saves to `model/triage_model.joblib` (backs up previous)

**Scoring logic (bucket thresholds):**
- `Urgent` ≥ 80, `Action` ≥ 45, `Waiting` ≥ 20, `FYI` < 20, `Noise` if noise_pattern && score < 0

**Key configuration constants (top of `outlook_triage.py`):**
- `DAYS_BACK`, `MAX_ITEMS`, `DRY_RUN`, `MOVE_NOISE_TO_READ_LATER`, `PROTECT_NON_TRIAGE_CATEGORIES`
- `KEYWORD_WEIGHTS` dict, `NOISE_PATTERNS` regex list, `VIP_SENDERS_FILE` path

## Tests

Tests run cross-platform (Linux/CI). `tests/conftest.py` mocks `win32com`/`pythoncom` and redirects `ONEDRIVE` to a temp dir before any imports.

- `tests/test_outlook_triage.py` — unit tests for all scoring/helper functions
- `tests/test_train_model.py` — unit tests for label normalization and data loading

## CI

`.github/workflows/ci.yml` runs on Ubuntu (Python 3.12): flake8 → bandit → pytest.
