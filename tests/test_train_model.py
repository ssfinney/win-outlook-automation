import pandas as pd

import train_model as tm


def test_strip_excel_formula_escape():
    assert tm.strip_excel_formula_escape("'=SUM(A1:A2)") == "=SUM(A1:A2)"
    assert tm.strip_excel_formula_escape("'+hello") == "+hello"
    assert tm.strip_excel_formula_escape("normal") == "normal"


def test_normalize_label_handles_fyi_and_escaped_values():
    assert tm.normalize_label("fyi") == "FYI"
    assert tm.normalize_label(" FYI ") == "FYI"
    assert tm.normalize_label("'urgent") == ""
    assert tm.normalize_label("'=noise") == ""
    assert tm.normalize_label("noise") == "Noise"


def test_normalize_label_handles_non_string_values():
    assert tm.normalize_label(1) == ""
    assert tm.normalize_label(True) == ""
    assert tm.normalize_label(None) == ""


def test_normalize_text_columns_unescapes_report_export_values():
    df = pd.DataFrame(
        {
            "subject": ["'=urgent task", None],
            "body_snippet": ["'@mention", "plain"],
            "sender_email": ["'+alerts@firm.com", "advisor@firm.com"],
            "to_line": ["'-team@firm.com", "ops@firm.com"],
            "cc_line": ["'|dist@firm.com", ""],
        }
    )

    out = tm.normalize_text_columns(df.copy(), tm.TEXT_COLS)

    assert out.loc[0, "subject"] == "=urgent task"
    assert out.loc[0, "body_snippet"] == "@mention"
    assert out.loc[0, "sender_email"] == "+alerts@firm.com"
    assert out.loc[0, "to_line"] == "-team@firm.com"
    assert out.loc[0, "cc_line"] == "|dist@firm.com"
    assert out.loc[1, "subject"] == ""


def test_load_labeled_rows_normalizes_label_and_text(monkeypatch):
    sample = pd.DataFrame(
        {
            "label": ["fyi", " Action ", "unknown"],
            "subject": ["'=hello", "plain", "drop me"],
            "body_snippet": ["'@body", "ok", "n/a"],
            "sender_email": ["'+sender@x.com", "sender@y.com", "sender@z.com"],
            "to_line": ["'-to@x.com", "to@y.com", "to@z.com"],
            "cc_line": ["'|cc@x.com", "cc@y.com", "cc@z.com"],
            "entry_id": ["1", "2", "3"],
        }
    )

    class _FakePath:
        def __init__(self, name):
            self.name = name

    class _FakeOutputDir:
        def glob(self, _pattern):
            return [_FakePath("triage_report_1.xlsx")]

    monkeypatch.setattr(tm, "OUTPUT_DIR", _FakeOutputDir())
    monkeypatch.setattr(tm.pd, "read_excel", lambda *_args, **_kwargs: sample.copy())

    out = tm.load_labeled_rows()

    assert list(out["label"]) == ["FYI", "Action"]
    assert out.iloc[0]["subject"] == "=hello"
    assert out.iloc[0]["body_snippet"] == "@body"
