# train_model.py

import os
import shutil
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime
from pathlib import Path

import pandas as pd
import joblib

from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.linear_model import LogisticRegression
from sklearn.pipeline import Pipeline
from sklearn.compose import ColumnTransformer
from sklearn.preprocessing import StandardScaler
from sklearn.impute import SimpleImputer
from sklearn.metrics import classification_report
from sklearn.model_selection import train_test_split

BASE_DIR = (
    Path(os.environ.get("ONEDRIVE", str(Path.home() / "OneDrive"))) / "AI_Outlook"
)
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "outputs"
MODEL_DIR = BASE_DIR / "model"
MODEL_PATH = MODEL_DIR / "triage_model.joblib"
LOG_FILE = DATA_DIR / "train_model.log"

# Ensure DATA_DIR exists before the log handler tries to open LOG_FILE.
# (outlook_triage.py creates these dirs when it runs, but train_model.py
# may be invoked independently.)
DATA_DIR.mkdir(parents=True, exist_ok=True)

logger = logging.getLogger("train_model")
logger.setLevel(logging.INFO)
if not logger.handlers:
    _handler = RotatingFileHandler(
        LOG_FILE, maxBytes=2_000_000, backupCount=3, encoding="utf-8"
    )
    _handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
    logger.addHandler(_handler)

LABELS = {"Urgent", "Action", "Waiting", "FYI", "Noise"}

TEXT_COLS = ["subject", "body_snippet", "sender_email", "to_line", "cc_line"]
NUMERIC_COLS = [
    "age_hours",
    "has_attachment",
    "rule_score",
    "is_noise_hint",
    "thread_depth",
    "recipient_count",
    "is_reply_or_fwd",
]
ALL_FEATURE_COLS = TEXT_COLS + NUMERIC_COLS
FORMULA_PREFIX_CHARS = ("=", "+", "-", "@", "|", "%")
_LABEL_MAP = {
    "urgent": "Urgent",
    "action": "Action",
    "waiting": "Waiting",
    "fyi": "FYI",
    "noise": "Noise",
}


def strip_excel_formula_escape(value):
    """Undo report-time Excel escaping for formula-like text cells."""
    if not isinstance(value, str):
        return value
    if len(value) >= 2 and value[0] == "'" and value[1] in FORMULA_PREFIX_CHARS:
        return value[1:]
    return value


def normalize_label(value) -> str:
    cleaned = strip_excel_formula_escape(value)
    return _LABEL_MAP.get(str(cleaned).strip().lower(), "")


def normalize_text_columns(df: pd.DataFrame, text_cols) -> pd.DataFrame:
    for col in text_cols:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].fillna("").astype(str).map(strip_excel_formula_escape)
    return df


def load_labeled_rows() -> pd.DataFrame:
    files = sorted(OUTPUT_DIR.glob("triage_report_*.xlsx"))
    if not files:
        raise RuntimeError(f"No triage reports found in {OUTPUT_DIR}")

    rows = []
    for f in files:
        try:
            df = pd.read_excel(f, sheet_name="All Scored")
        except Exception as e:
            logger.error(f"Error reading {f.name}: {e}")
            continue

        if "label" not in df.columns:
            continue

        df = df.dropna(subset=["label"]).copy()
        df["label"] = df["label"].map(normalize_label)
        df = normalize_text_columns(df, TEXT_COLS)
        df = df[df["label"].isin(LABELS)].copy()
        if df.empty:
            continue

        df["source_file"] = f.name
        rows.append(df)

    if not rows:
        raise RuntimeError(
            "No labeled rows found. Fill 'label' with: Urgent/Action/Waiting/FYI/Noise"
        )

    df_combined = pd.concat(rows, ignore_index=True)

    if "entry_id" in df_combined.columns:
        initial_len = len(df_combined)
        df_combined = df_combined.drop_duplicates(subset=["entry_id"], keep="last")
        logger.info(f"Deduplicated data from {initial_len} to {len(df_combined)} rows.")

    return df_combined


def build_pipeline() -> Pipeline:
    pre = ColumnTransformer(
        transformers=[
            # min_df=2 on subject/body filters single-occurrence noise.
            # These fields have repeated tokens across rows.
            ("subject_tfidf", TfidfVectorizer(ngram_range=(1, 2), min_df=2), "subject"),
            (
                "body_tfidf",
                TfidfVectorizer(ngram_range=(1, 2), min_df=2, max_features=500),
                "body_snippet",
            ),
            # sender_email, to_line, cc_line are often mostly unique in
            # small datasets; min_df=1 avoids empty-vocabulary errors.
            (
                "sender_tfidf",
                TfidfVectorizer(
                    token_pattern=r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+",  # nosec B106
                    min_df=1,
                ),
                "sender_email",
            ),
            ("to_tfidf", TfidfVectorizer(ngram_range=(1, 1), min_df=1), "to_line"),
            ("cc_tfidf", TfidfVectorizer(ngram_range=(1, 1), min_df=1), "cc_line"),
            (
                "num",
                Pipeline(
                    steps=[
                        ("imputer", SimpleImputer(strategy="median")),
                        ("scaler", StandardScaler(with_mean=False)),
                    ]
                ),
                NUMERIC_COLS,
            ),
        ],
        remainder="drop",
    )

    clf = LogisticRegression(max_iter=2000, class_weight="balanced")
    return Pipeline(steps=[("pre", pre), ("clf", clf)])


def main():
    MODEL_DIR.mkdir(parents=True, exist_ok=True)

    try:
        df = load_labeled_rows()
    except Exception as e:
        logger.error(str(e))
        print(e)
        return

    # Ensure required numeric columns exist and are numeric
    for col, default in [
        ("age_hours", 0),
        ("has_attachment", 0),
        ("rule_score", 0),
        ("is_noise_hint", 0),
        ("thread_depth", 0),
        ("recipient_count", 0),
        ("is_reply_or_fwd", 0),
    ]:
        if col not in df.columns:
            df[col] = default
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(default)

    # Ensure required text columns exist
    df = normalize_text_columns(df, TEXT_COLS)

    X = df[ALL_FEATURE_COLS]
    y = df["label"].astype(str)

    do_split = len(df) >= 50 and y.nunique() >= 2
    if do_split:
        X_train, X_test, y_train, y_test = train_test_split(
            X, y, test_size=0.2, random_state=42, stratify=y
        )
    else:
        X_train, X_test, y_train, y_test = X, X, y, y
        logger.info(
            "Dataset too small for split "
            f"({len(df)} rows, {y.nunique()} classes). "
            "Training on all data."
        )

    pipe = build_pipeline()
    pipe.fit(X_train, y_train)

    if do_split:
        preds = pipe.predict(X_test)
        report = classification_report(y_test, preds)
        print(report)
        logger.info(f"Classification Report:\n{report}")

    if MODEL_PATH.exists():
        ts = datetime.now().strftime("%Y%m%d_%H%M")
        backup = MODEL_DIR / f"triage_model_{ts}.joblib"
        shutil.copy2(MODEL_PATH, backup)
        logger.info(f"Backed up previous model to {backup}")

    joblib.dump(pipe, MODEL_PATH)

    msg = f"Saved model to: {MODEL_PATH} | Labeled rows used: {len(df)}"
    logger.info(msg)
    print(msg)


if __name__ == "__main__":
    main()
