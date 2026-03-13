# outlook_triage.py
# - Offline, local-only triage for Classic Outlook (Windows)
# - COM-safe iteration (bounded)
# - Guardrails for manual categories
# - Dry-run safe default
# - Health summary: errors counter + Summary sheet + single end-of-run log line

import os
import re
import logging
from logging.handlers import RotatingFileHandler
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Any, Dict, Tuple, Set

import pandas as pd

try:
    import win32com.client
    import pythoncom
except ImportError:
    print("ERROR: pywin32 not installed or incompatible. Use Python 3.12 for this project.")
    raise

try:
    import joblib
except Exception:
    joblib = None

# =====================
# Configuration (edit here)
# =====================
# Storage base. Defaults to OneDrive\\AI_Outlook if OneDrive is set, else ~/OneDrive/AI_Outlook.
BASE_DIR = Path(os.environ.get("OneDrive", str(Path.home() / "OneDrive"))) / "AI_Outlook"

# Lookback window for scanning the inbox.
DAYS_BACK = 7

# Hard cap on enumerated items (newest-first) to keep COM iteration bounded.
MAX_ITEMS = 500

# If True, move "Noise" emails to a "Read Later" folder.
MOVE_NOISE_TO_READ_LATER = False

# Safety: first production runs should not mutate Outlook items.
# When DRY_RUN is True, the script will score + report, but will not:
# - set Categories
# - set flags
# - move messages
DRY_RUN = True

# Guardrail: if True, never modify messages that have non-triage categories.
PROTECT_NON_TRIAGE_CATEGORIES = True

# Outputs and artifacts
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "outputs"
CONFIG_DIR = BASE_DIR / "config"
MODEL_DIR = BASE_DIR / "model"
MODEL_PATH = MODEL_DIR / "triage_model.joblib"
LOG_FILE = DATA_DIR / "triage.log"

# Max rows written per bucket tab in the Excel report.
EXCEL_BUCKET_ROW_LIMIT = 75


def validate_config() -> None:
    if not isinstance(DAYS_BACK, int) or DAYS_BACK <= 0 or DAYS_BACK > 90:
        raise ValueError(f"DAYS_BACK must be an int between 1 and 90 (got {DAYS_BACK})")

    if not isinstance(MAX_ITEMS, int) or MAX_ITEMS <= 0 or MAX_ITEMS > 5000:
        raise ValueError(f"MAX_ITEMS must be an int between 1 and 5000 (got {MAX_ITEMS})")

    if not isinstance(MOVE_NOISE_TO_READ_LATER, bool):
        raise ValueError("MOVE_NOISE_TO_READ_LATER must be a bool")

    if not isinstance(DRY_RUN, bool):
        raise ValueError("DRY_RUN must be a bool")

    if not isinstance(PROTECT_NON_TRIAGE_CATEGORIES, bool):
        raise ValueError("PROTECT_NON_TRIAGE_CATEGORIES must be a bool")


CAT_URGENT = "Urgent"
CAT_ACTION = "Action"
CAT_WAITING = "Waiting"
CAT_FYI = "FYI"
CAT_NOISE = "Noise"
TRIAGE_CATEGORIES = {CAT_URGENT, CAT_ACTION, CAT_WAITING, CAT_FYI, CAT_NOISE}
FOLDER_READ_LATER = "Read Later"

KEYWORD_WEIGHTS = {
    "rollover": 30,
    "esignature": 25,
    "e-signature": 25,
    "underwriting": 35,
    "acat": 35,
    "beneficiary": 30,
    "distribution": 25,
    "rmd": 35,
    "urgent": 40,
    "asap": 30,
    "today": 25,
    "deadline": 30,
}

NOISE_PATTERNS = [
    r"\bunsubscribe\b",
    r"\bnewsletter\b",
    r"\bwebinar\b",
    r"\bdigest\b",
    r"\bpromo\b",
    r"\bmarketing\b",
    r"\bno[- ]reply\b",
]

VIP_SENDERS_CSV = Path(os.environ.get("VIP_SENDERS_CSV_PATH", str(CONFIG_DIR / "vip_senders.csv")))


def ensure_dirs() -> None:
    for p in [BASE_DIR, DATA_DIR, OUTPUT_DIR, CONFIG_DIR, MODEL_DIR]:
        p.mkdir(parents=True, exist_ok=True)


logger = logging.getLogger("outlook_triage")


def _setup_logging() -> None:
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        _handler = RotatingFileHandler(LOG_FILE, maxBytes=2_000_000, backupCount=3, encoding="utf-8")
        _handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
        logger.addHandler(_handler)


@dataclass
class ScoredMail:
    entry_id: str
    received: datetime
    sender_email: str
    sender_name: str
    subject: str
    to_line: str
    cc_line: str
    conversation_id: str
    body_snippet: str
    age_hours: float
    has_attachment: int
    thread_depth: int
    recipient_count: int
    is_reply_or_fwd: int
    rule_score: int
    rule_bucket: str
    model_bucket: str
    final_bucket: str
    reasons: str
    is_noise_hint: int
    action_status: str


def safe_str(x: Any) -> str:
    try:
        return str(x) if x is not None else ""
    except Exception:
        return ""


def naive_dt(dt_val) -> datetime:
    try:
        return dt_val.replace(tzinfo=None)
    except (AttributeError, TypeError):
        # dt_val is not a datetime (e.g. malformed COM return). Return
        # datetime.min so callers can safely compare against a cutoff;
        # the item will be treated as out-of-range and skipped.
        return datetime.min


def load_vips() -> Set[str]:
    vips: Set[str] = set()

    try:
        VIP_SENDERS_CSV.parent.mkdir(parents=True, exist_ok=True)
        if not VIP_SENDERS_CSV.exists():
            VIP_SENDERS_CSV.write_text("", encoding="utf-8")
    except Exception as e:
        logger.warning(f"Could not initialize VIP senders file at {VIP_SENDERS_CSV}: {e}")
        return vips

    for line in VIP_SENDERS_CSV.read_text(encoding="utf-8").splitlines():
        email = line.strip().lower()
        if not email or email.startswith("#"):
            continue
        if re.match(r"^[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}$", email):
            vips.add(email)
        else:
            logger.warning(f"Ignoring invalid VIP sender entry: '{line.strip()}'")

    return vips


def compile_patterns(patterns: List[str]) -> List[re.Pattern]:
    return [re.compile(pat, flags=re.IGNORECASE) for pat in patterns]


def get_sender_email(mail_item) -> str:
    # Best effort for Exchange quirks
    try:
        addr = safe_str(mail_item.SenderEmailAddress)
        if addr and "@" in addr:
            return addr.lower()
    except Exception:
        pass

    try:
        sender = mail_item.Sender
        if sender is not None:
            ex_user = sender.GetExchangeUser()
            if ex_user is not None:
                smtp = safe_str(ex_user.PrimarySmtpAddress)
                if smtp and "@" in smtp:
                    return smtp.lower()
    except Exception:
        pass

    return ""


def is_noise(subject: str, sender_email: str, patterns: List[re.Pattern]) -> bool:
    blob = f"{subject} {sender_email}"
    return any(p.search(blob) for p in patterns)


def keyword_score(text: str) -> Tuple[int, List[str]]:
    s = 0
    hits = []
    t = text.lower()
    for k, w in KEYWORD_WEIGHTS.items():
        if k in t:
            s += w
            hits.append(k)
    return s, hits


def thread_depth(mail_item) -> int:
    try:
        idx = safe_str(getattr(mail_item, "ConversationIndex", ""))
        if len(idx) > 44:
            return (len(idx) - 44) // 10
    except Exception:
        pass
    return 0


def recipient_count(to_line: str) -> int:
    if not to_line.strip():
        return 0
    return len([r for r in to_line.split(";") if r.strip()])


def is_reply_or_forward(subject: str) -> int:
    return 1 if re.match(r"^(RE|FW|FWD)\s*:", subject, re.IGNORECASE) else 0


def already_triaged(mail_item) -> bool:
    """Returns True if any triage category is already present."""
    try:
        cats = safe_str(mail_item.Categories)
        if cats:
            for cat in cats.split(","):
                if cat.strip() in TRIAGE_CATEGORIES:
                    return True
    except Exception:
        pass
    return False


def has_non_triage_categories(mail_item) -> bool:
    """Guardrail: if user set any non-triage categories, do not modify."""
    try:
        cats = safe_str(mail_item.Categories)
        if not cats:
            return False
        existing = [c.strip() for c in cats.split(",") if c.strip()]
        if not existing:
            return False
        return any(c not in TRIAGE_CATEGORIES for c in existing)
    except Exception:
        return False


def merge_categories(existing_cats: str, add_cat: str) -> str:
    existing = [c.strip() for c in safe_str(existing_cats).split(",") if c.strip()]
    if add_cat and add_cat not in existing:
        existing.append(add_cat)
    return ", ".join(existing)


def rule_score_and_bucket(
    mail_item, vips: set, noise_pats: List[re.Pattern], received: datetime = None
) -> Tuple[int, str, str, Dict[str, Any]]:
    subject = safe_str(mail_item.Subject)
    sender_email = get_sender_email(mail_item)
    to_line = safe_str(mail_item.To)
    cc_line = safe_str(mail_item.CC)
    if received is None:
        received = naive_dt(mail_item.ReceivedTime)

    reasons: List[str] = []
    score = 0

    if sender_email and sender_email in vips:
        score += 50
        reasons.append("VIP_sender")

    if to_line.strip():
        score += 10
        reasons.append("To_line_present")
    if cc_line.strip():
        score -= 5
        reasons.append("CC_present")

    body_snippet = ""
    try:
        body_snippet = safe_str(mail_item.Body)[:500]
    except Exception:
        pass

    kscore, hits = keyword_score(f"{subject} {body_snippet}")
    if kscore:
        score += kscore
        reasons.append(f"keywords:{','.join(hits)}")

    attach_count = 0
    try:
        attach_count = mail_item.Attachments.Count
        if attach_count > 0:
            score += 8
            reasons.append("has_attachment")
    except Exception:
        pass

    age_hours = 0.0
    try:
        age_hours = max(0.0, (datetime.now() - received).total_seconds() / 3600.0)
        if age_hours > 24:
            score -= int(min(25, age_hours // 24 * 5))
            reasons.append("age_penalty")
    except Exception:
        pass

    noise = is_noise(subject, sender_email, noise_pats)
    if noise:
        score -= 40
        reasons.append("noise_pattern")

    if noise and score < 0:
        bucket = CAT_NOISE
    elif score >= 80:
        bucket = CAT_URGENT
    elif score >= 45:
        bucket = CAT_ACTION
    elif score >= 20:
        bucket = CAT_WAITING
    else:
        bucket = CAT_FYI

    features = {
        "sender_email": sender_email,
        "sender_name": safe_str(mail_item.SenderName),
        "subject": subject,
        "body_snippet": body_snippet,
        "to_line": to_line,
        "cc_line": cc_line,
        "age_hours": age_hours,
        "has_attachment": int(attach_count > 0),
        "thread_depth": thread_depth(mail_item),
        "recipient_count": recipient_count(to_line),
        "is_reply_or_fwd": is_reply_or_forward(subject),
        "rule_score": score,
        "is_noise_hint": int(noise),
    }

    return score, bucket, ";".join(reasons), features


def load_model():
    if joblib is None:
        return None
    if MODEL_PATH.exists():
        try:
            return joblib.load(MODEL_PATH)
        except Exception as e:
            logger.error(f"Failed to load model: {e}")
            return None
    return None


def choose_final_bucket(rule_bucket: str, model_bucket: str, rule_score: int) -> str:
    if rule_bucket == CAT_URGENT:
        return CAT_URGENT
    if rule_bucket == CAT_NOISE:
        return CAT_NOISE
    if model_bucket:
        return model_bucket
    return rule_bucket


def ensure_outlook_folder(namespace, folder_name: str):
    inbox = namespace.GetDefaultFolder(6)
    try:
        for f in inbox.Folders:
            if safe_str(f.Name).strip().lower() == folder_name.lower():
                return f
        return inbox.Folders.Add(folder_name)
    except Exception as e:
        logger.warning(f"Could not find/create folder '{folder_name}': {e}")
        return None


def apply_actions(mail_item, final_bucket: str, read_later) -> str:
    if DRY_RUN:
        return "dry_run"

    if PROTECT_NON_TRIAGE_CATEGORIES and has_non_triage_categories(mail_item):
        logger.info(f"Skip actions (manual/non-triage categories present): '{safe_str(mail_item.Subject)}'")
        return "skipped_manual_categories"

    try:
        existing_cats = safe_str(mail_item.Categories)
        mail_item.Categories = merge_categories(existing_cats, final_bucket)

        if final_bucket in (CAT_URGENT, CAT_ACTION):
            mail_item.FlagStatus = 2
            mail_item.FlagRequest = "Follow up"

        mail_item.Save()
    except Exception as e:
        logger.warning(f"Failed to apply actions to '{safe_str(mail_item.Subject)}': {e}")
        return "failed_apply"

    if final_bucket == CAT_NOISE and MOVE_NOISE_TO_READ_LATER and read_later is not None:
        try:
            mail_item.Move(read_later)
        except Exception as e:
            logger.warning(f"Failed to move noise email: {e}")
            return "failed_move_noise"

    return "applied"


def collect_items(inbox) -> List:
    """COM-safe enumeration of recent mail items without holding COM objects longer than needed."""
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)

    # Restrict by time first to reduce mailbox traversal work on large inboxes.
    # Outlook Restrict expects US-style date format for ReceivedTime in many locales.
    cutoff = datetime.now() - timedelta(days=DAYS_BACK)
    restrict_succeeded = False
    restrict_date = cutoff.strftime("%m/%d/%Y %I:%M %p")
    try:
        items = items.Restrict(f"[ReceivedTime] >= '{restrict_date}'")
        restrict_succeeded = True
    except Exception as e:
        logger.warning(
            f"Could not apply ReceivedTime restriction; falling back to full scan "
            f"with Python-side cutoff enforcement: {e}"
        )

    result = []
    item = items.GetFirst()
    while item is not None and len(result) < MAX_ITEMS:
        try:
            if getattr(item, "Class", None) == 43:
                # When Restrict failed, enforce the cutoff in Python and stop early
                # (items are sorted newest-first, so once we pass the window we are done).
                if not restrict_succeeded:
                    received = naive_dt(getattr(item, "ReceivedTime", None))
                    if received < cutoff:
                        break
                # Store stable ID, then re-bind during processing. Holding many COM item objects can
                # make Outlook sluggish or unstable on large mailboxes.
                entry_id = safe_str(getattr(item, "EntryID", ""))
                if entry_id:
                    result.append(entry_id)
        except Exception:
            pass
        try:
            item = items.GetNext()
        except Exception:
            break
    return result


def main():
    ensure_dirs()
    validate_config()
    _setup_logging()
    pythoncom.CoInitialize()
    logger.info("Starting Outlook Triage scan...")
    vips = load_vips()
    noise_pats = compile_patterns(NOISE_PATTERNS)
    model = load_model()

    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    except Exception as e:
        logger.error(f"Failed to dispatch Outlook COM object: {e}")
        print(f"Cannot connect to Outlook: {e}")
        return

    inbox = outlook.GetDefaultFolder(6)
    read_later = ensure_outlook_folder(outlook, FOLDER_READ_LATER) if MOVE_NOISE_TO_READ_LATER else None

    items_list = collect_items(inbox)

    scored: List[ScoredMail] = []
    processed = 0
    skipped = 0
    errors = 0

    # Cutoff retained here as a safety net: if collect_items fell back to a full
    # scan (Restrict failed), this prevents processing emails older than DAYS_BACK.
    cutoff = datetime.now() - timedelta(days=DAYS_BACK)

    if DRY_RUN:
        logger.warning("DRY_RUN=True: categories/flags/moves are disabled for this run")
        print("NOTE: DRY_RUN=True, so no categories/flags/moves will be applied.")

    for entry_id in items_list:
        try:
            item = outlook.GetItemFromID(entry_id)
            received = naive_dt(item.ReceivedTime)
        except Exception:
            errors += 1
            continue

        if received < cutoff:
            continue

        if already_triaged(item):
            skipped += 1
            continue

        conv_id = safe_str(getattr(item, "ConversationID", ""))
        sender_name = safe_str(item.SenderName)
        sender_email = get_sender_email(item)
        subject = safe_str(item.Subject)
        to_line = safe_str(item.To)
        cc_line = safe_str(item.CC)

        try:
            rule_score, rule_bucket, reasons, features = rule_score_and_bucket(item, vips, noise_pats, received)

            model_bucket = ""
            if model is not None:
                try:
                    df_features = pd.DataFrame([features])
                    model_bucket = str(model.predict(df_features)[0])
                except Exception as e:
                    errors += 1
                    logger.error(f"Model prediction error for {entry_id}: {e}")

            final_bucket = choose_final_bucket(rule_bucket, model_bucket, rule_score)
            action_status = apply_actions(item, final_bucket, read_later)

            scored.append(
                ScoredMail(
                    entry_id=entry_id,
                    received=received,
                    sender_email=sender_email,
                    sender_name=sender_name,
                    subject=subject,
                    to_line=to_line,
                    cc_line=cc_line,
                    conversation_id=conv_id,
                    body_snippet=features["body_snippet"],
                    age_hours=float(features["age_hours"]),
                    has_attachment=int(features["has_attachment"]),
                    thread_depth=int(features["thread_depth"]),
                    recipient_count=int(features["recipient_count"]),
                    is_reply_or_fwd=int(features["is_reply_or_fwd"]),
                    rule_score=int(rule_score),
                    rule_bucket=str(rule_bucket),
                    model_bucket=str(model_bucket),
                    final_bucket=str(final_bucket),
                    reasons=str(reasons),
                    is_noise_hint=int(features["is_noise_hint"]),
                    action_status=action_status,
                )
            )
            processed += 1
        except Exception as e:
            errors += 1
            logger.error(f"Unhandled processing error for {entry_id}: {e}")
            continue
        finally:
            item = None

    now_tag = datetime.now().strftime("%Y-%m-%d_%H%M")
    report_xlsx = OUTPUT_DIR / f"triage_report_{now_tag}.xlsx"
    log_csv = DATA_DIR / f"inbox_scored_{now_tag}.csv"

    df = pd.DataFrame([s.__dict__ for s in scored])
    if not df.empty:
        df.sort_values(by=["rule_score", "received"], ascending=[False, False], inplace=True)

    df.to_csv(log_csv, index=False, encoding="utf-8")

    with pd.ExcelWriter(report_xlsx, engine="openpyxl") as writer:
        df_summary = pd.DataFrame(
            [
                {
                    "run_timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "dry_run": int(DRY_RUN),
                    "days_back": DAYS_BACK,
                    "max_items": MAX_ITEMS,
                    "processed": processed,
                    "skipped": skipped,
                    "errors": errors,
                    "move_noise_to_read_later": int(MOVE_NOISE_TO_READ_LATER),
                }
            ]
        )
        df_summary.to_excel(writer, sheet_name="Summary", index=False)

        df_out = df.copy()
        df_out.insert(0, "label", "")
        # Sanitize string columns: prepend ' to cells starting with Excel
        # formula trigger characters so they are not evaluated when opened.
        _formula_chars = ("=", "+", "-", "@", "|", "%")
        for col in df_out.select_dtypes(include=["object"]).columns:
            df_out[col] = df_out[col].apply(
                lambda v: ("'" + v) if isinstance(v, str) and v and v[0] in _formula_chars else v
            )
        df_out.to_excel(writer, sheet_name="All Scored", index=False)

        for name, bucket in [
            ("Urgent", CAT_URGENT),
            ("Action", CAT_ACTION),
            ("Waiting", CAT_WAITING),
            ("FYI", CAT_FYI),
            ("Noise", CAT_NOISE),
        ]:
            sub = df_out[df_out["final_bucket"] == bucket].head(EXCEL_BUCKET_ROW_LIMIT) if not df_out.empty else df_out
            sub.to_excel(writer, sheet_name=name, index=False)

    logger.info(
        f"Processed={processed} skipped={skipped} errors={errors} "
        f"dry_run={DRY_RUN} days_back={DAYS_BACK} max_items={MAX_ITEMS} report={report_xlsx}"
    )
    print(
        f"Processed {processed} emails ({skipped} skipped, {errors} errors) "
        f"from the last {DAYS_BACK} days. Dry-run={DRY_RUN}."
    )

    outlook = None
    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass


if __name__ == "__main__":
    main()
