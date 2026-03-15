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

        def __lt__(self, other):
            return self.name < other.name

    class _FakeOutputDir:
        def glob(self, _pattern):
            return [_FakePath("triage_report_1.xlsx")]

    monkeypatch.setattr(tm, "OUTPUT_DIR", _FakeOutputDir())
    monkeypatch.setattr(tm.pd, "read_excel", lambda *_args, **_kwargs: sample.copy())

    out = tm.load_labeled_rows()

    assert list(out["label"]) == ["FYI", "Action"]
    assert out.iloc[0]["subject"] == "=hello"
    assert out.iloc[0]["body_snippet"] == "@body"


def test_normalize_label_all_valid_labels():
    """Every canonical label string round-trips through normalize_label."""
    for raw in ["urgent", "action", "waiting", "fyi", "noise"]:
        result = tm.normalize_label(raw)
        assert result in tm.LABELS, f"Expected a valid label for '{raw}', got '{result}'"


def test_build_pipeline_returns_pipeline():
    from sklearn.pipeline import Pipeline

    pipe = tm.build_pipeline()
    assert isinstance(pipe, Pipeline)
    assert "pre" in pipe.named_steps
    assert "clf" in pipe.named_steps


def test_load_labeled_rows_deduplicates_by_entry_id(monkeypatch):
    """When two reports contain the same entry_id, only the last occurrence is kept."""
    row_a = pd.DataFrame(
        {
            "label": ["action"],
            "subject": ["first version"],
            "body_snippet": [""],
            "sender_email": ["s@x.com"],
            "to_line": ["t@x.com"],
            "cc_line": [""],
            "entry_id": ["dup-id-1"],
        }
    )
    row_b = pd.DataFrame(
        {
            "label": ["urgent"],
            "subject": ["second version"],
            "body_snippet": [""],
            "sender_email": ["s@x.com"],
            "to_line": ["t@x.com"],
            "cc_line": [""],
            "entry_id": ["dup-id-1"],
        }
    )

    class _FakePath:
        def __init__(self, name):
            self.name = name

        def __lt__(self, other):
            return self.name < other.name

    class _FakeOutputDir:
        def glob(self, _pattern):
            return [_FakePath("report_a.xlsx"), _FakePath("report_b.xlsx")]

    call_count = 0

    def _fake_read_excel(*_args, **_kwargs):
        nonlocal call_count
        call_count += 1
        return row_a.copy() if call_count == 1 else row_b.copy()

    monkeypatch.setattr(tm, "OUTPUT_DIR", _FakeOutputDir())
    monkeypatch.setattr(tm.pd, "read_excel", _fake_read_excel)

    out = tm.load_labeled_rows()

    # Only one row should remain after deduplication
    assert len(out) == 1
    # last-seen-wins — row_b's label should win
    assert out.iloc[0]["label"] == "Urgent"
    assert out.iloc[0]["subject"] == "second version"
