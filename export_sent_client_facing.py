#!/usr/bin/env python3
"""Export representative client-facing sent Outlook emails to JSONL."""

from __future__ import annotations

import argparse
import datetime as dt
import json
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Set

try:
    import win32com.client  # type: ignore
except ImportError as exc:  # pragma: no cover
    raise SystemExit(
        "pywin32 is required. Install with: pip install pywin32"
    ) from exc

PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
REPLY_HEADER_LINE_RE = re.compile(
    r"^(from:|sent:|to:|cc:|subject:)\\s*", re.IGNORECASE
)
ON_WROTE_RE = re.compile(
    r"^on\s+.+\s+wrote:\s*$", re.IGNORECASE
)
FW_SUBJECT_RE = re.compile(r"^\s*fwd?\s*:", re.IGNORECASE)
SIGNOFF_RE = re.compile(
    r"^(thanks,?|thank you,?|best,?|sincerely,?|v/r|[-—]\s*stephen)\s*$",
    re.IGNORECASE,
)
MOBILE_FOOTER_RE = re.compile(
    r"^sent from (my )?(iphone|ipad|android|yahoo mail|gmail|outlook|mobile)\\b",
    re.IGNORECASE,
)
FORWARDED_CONTENT_HINT_RE = re.compile(
    r"(-----original message-----|^from:|^sent:|^to:|^subject:)",
    re.IGNORECASE | re.MULTILINE,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--max", type=int, default=500, help="Max kept messages")
    parser.add_argument(
        "--since",
        type=str,
        default=None,
        help="Only include emails sent on/after YYYY-MM-DD",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=Path("sent_client_facing.jsonl"),
        help="Output JSONL path",
    )
    parser.add_argument("--dry-run", action="store_true", help="Only print counts")
    parser.add_argument(
        "--exclude-email",
        action="append",
        default=[],
        help="Recipient email to exclude (repeat flag for multiple).",
    )
    parser.add_argument(
        "--exclude-file",
        type=Path,
        default=None,
        help="Optional file with one email per line to exclude.",
    )
    return parser.parse_args()


def resolve_recipient_smtp(recipient: Any) -> Optional[str]:
    try:
        ae = recipient.AddressEntry
        if ae is not None:
            try:
                ex_user = ae.GetExchangeUser()
                if ex_user and ex_user.PrimarySmtpAddress:
                    return ex_user.PrimarySmtpAddress.strip().lower()
            except Exception:
                pass
            try:
                ex_dl = ae.GetExchangeDistributionList()
                if ex_dl and ex_dl.PrimarySmtpAddress:
                    return ex_dl.PrimarySmtpAddress.strip().lower()
            except Exception:
                pass
        try:
            pa = recipient.PropertyAccessor
            smtp = pa.GetProperty(PR_SMTP_ADDRESS)
            if smtp:
                return str(smtp).strip().lower()
        except Exception:
            pass
        addr = getattr(recipient, "Address", None)
        if addr:
            return str(addr).strip().lower()
    except Exception:
        return None
    return None


def collect_recipients(mail_item: Any) -> tuple[List[str], List[str]]:
    to_list: List[str] = []
    cc_list: List[str] = []
    for recipient in getattr(mail_item, "Recipients", []):
        smtp = resolve_recipient_smtp(recipient)
        if not smtp:
            continue
        rtype = getattr(recipient, "Type", 0)
        if rtype == 2:
            cc_list.append(smtp)
        else:
            to_list.append(smtp)
    return sorted(set(to_list)), sorted(set(cc_list))


def strip_quoted_text(body: str) -> str:
    lines = body.split("\n")
    cut_idx = len(lines)
    for idx, line in enumerate(lines):
        stripped = line.strip()
        if (
            REPLY_HEADER_LINE_RE.match(stripped)
            or stripped == "-----Original Message-----"
            or stripped == "________________________________"
            or ON_WROTE_RE.match(stripped)
        ):
            cut_idx = min(cut_idx, idx)
            break
    kept = [ln for ln in lines[:cut_idx] if not MOBILE_FOOTER_RE.match(ln.strip())]
    return "\n".join(kept)


def strip_signature_and_fluff(body: str) -> str:
    lines = body.split("\n")
    cut_idx = len(lines)
    for idx, line in enumerate(lines):
        stripped = line.strip()
        if stripped in {"--", "—", "___"}:
            cut_idx = min(cut_idx, idx)
            break
        if SIGNOFF_RE.match(stripped):
            cut_idx = min(cut_idx, idx)
            break
    body = "\n".join(lines[:cut_idx])

    marketing_phrases = [
        "schedule a meeting with me",
        "securely upload documents",
        "write us a google review",
    ]
    kept_lines: List[str] = []
    for line in body.split("\n"):
        lowered = line.lower()
        if any(phrase in lowered for phrase in marketing_phrases):
            continue
        kept_lines.append(line)
    body = "\n".join(kept_lines)

    nm_phrase = "northwestern mutual is the marketing name"
    lower_body = body.lower()
    pos = lower_body.find(nm_phrase)
    if pos != -1:
        body = body[:pos].rstrip()

    return body



def load_exclude_recipients(args: argparse.Namespace) -> Set[str]:
    recipients = {
        "ops@example.com",
        "internal@example.com",
    }
    recipients.update((e or "").strip().lower() for e in args.exclude_email)
    recipients.discard("")

    if args.exclude_file:
        try:
            for line in args.exclude_file.read_text(encoding="utf-8").splitlines():
                cleaned = line.strip().lower()
                if cleaned and not cleaned.startswith("#"):
                    recipients.add(cleaned)
        except Exception as exc:
            print(f"[WARN] Unable to read exclude file {args.exclude_file}: {exc}")

    return recipients


def is_underwriting_or_ops(subject: str, body: str) -> bool:
    corpus = f"{subject}\n{body}".lower()
    patterns = [
        "underwriting",
        "new business",
        "requirements received",
        "aps ordered",
        "case status",
        "policy service request",
        "internal use only",
    ]
    return any(p in corpus for p in patterns)

def normalize_body(body: str) -> str:
    body = body.replace("\r\n", "\n").replace("\r", "\n")
    body = strip_quoted_text(body)
    body = strip_signature_and_fluff(body)
    body = re.sub(r"\n{3,}", "\n\n", body)
    return body.strip()


def word_count(text: str) -> int:
    return len(re.findall(r"\b\w+\b", text))


def to_iso_utc(sent_on: Any) -> str:
    if not isinstance(sent_on, dt.datetime):
        return ""
    if sent_on.tzinfo is None:
        local = sent_on.astimezone()
    else:
        local = sent_on
    return local.astimezone(dt.timezone.utc).isoformat()


def is_signature_only(text: str) -> bool:
    lowered = text.lower()
    low_words = word_count(text)
    if low_words < 20:
        return True
    if any(k in lowered for k in ["confidential", "disclosure", "do not reply"]):
        return low_words < 50
    short_lines = [ln for ln in text.split("\n") if ln.strip()]
    if short_lines and sum(len(ln.split()) for ln in short_lines) / max(len(short_lines), 1) < 4:
        return True
    return False


def mostly_forwarded(subject: str, raw_body: str, clean_body: str) -> bool:
    if not FW_SUBJECT_RE.match(subject or ""):
        return False
    if not raw_body.strip():
        return True
    raw_words = max(word_count(raw_body), 1)
    clean_words = word_count(clean_body)
    forwarded_markers = len(FORWARDED_CONTENT_HINT_RE.findall(raw_body))
    return clean_words / raw_words < 0.35 or forwarded_markers >= 3


def sanity_check(records: List[Dict[str, Any]]) -> None:
    sample = records[:20]
    required = {"id", "sent_utc", "to", "cc", "subject", "body_plain", "category", "thread_hint", "word_count"}
    for idx, rec in enumerate(sample, start=1):
        missing = required - rec.keys()
        if missing:
            raise ValueError(f"Sanity check failed on item {idx}: missing keys {sorted(missing)}")
        dt.datetime.fromisoformat(rec["sent_utc"].replace("Z", "+00:00"))
        if not isinstance(rec["to"], list) or not all(isinstance(x, str) for x in rec["to"]):
            raise ValueError(f"Sanity check failed on item {idx}: invalid 'to'")
        if not isinstance(rec["cc"], list) or not all(isinstance(x, str) for x in rec["cc"]):
            raise ValueError(f"Sanity check failed on item {idx}: invalid 'cc'")
        if not isinstance(rec["body_plain"], str) or not rec["body_plain"].strip():
            raise ValueError(f"Sanity check failed on item {idx}: empty body_plain")


def main() -> None:
    args = parse_args()

    exclude_recipients = load_exclude_recipients(args)

    since_date = None
    if args.since:
        since_date = dt.datetime.strptime(args.since, "%Y-%m-%d").date()

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    sent_folder = namespace.GetDefaultFolder(5)  # olFolderSentMail
    items = sent_folder.Items
    items.Sort("[SentOn]", True)

    if since_date:
        start = dt.datetime.combine(since_date, dt.time.min)
        filter_str = "[SentOn] >= '{}'".format(start.strftime("%m/%d/%Y %I:%M %p"))
        items = items.Restrict(filter_str)

    kept: List[Dict[str, Any]] = []
    counts = {
        "total_scanned": 0,
        "excluded_internal_recipients": 0,
        "excluded_underwriting_patterns": 0,
        "excluded_noise_rules": 0,
        "kept": 0,
        "errors": 0,
    }

    for item in items:
        if counts["kept"] >= args.max:
            break
        counts["total_scanned"] += 1
        try:
            if getattr(item, "Class", None) != 43:
                continue
            to_list, cc_list = collect_recipients(item)
            recipient_set = set(to_list) | set(cc_list)
            if recipient_set & exclude_recipients:
                counts["excluded_internal_recipients"] += 1
                continue

            subject = str(getattr(item, "Subject", "") or "")
            raw_body = str(getattr(item, "Body", "") or "")
            clean_body = normalize_body(raw_body)
            wc = word_count(clean_body)

            if is_underwriting_or_ops(subject, clean_body):
                counts["excluded_underwriting_patterns"] += 1
                continue

            if is_signature_only(clean_body) or mostly_forwarded(subject, raw_body, clean_body):
                counts["excluded_noise_rules"] += 1
                continue

            sent_iso = to_iso_utc(getattr(item, "SentOn", None))
            if not sent_iso:
                counts["errors"] += 1
                print(f"[WARN] Missing/invalid SentOn at scan #{counts['total_scanned']}")
                continue

            rec = {
                "id": str(getattr(item, "EntryID", "") or ""),
                "sent_utc": sent_iso,
                "to": to_list,
                "cc": cc_list,
                "subject": subject,
                "body_plain": clean_body,
                "category": "unlabeled",
                "thread_hint": str(getattr(item, "ConversationID", "") or "") or None,
                "word_count": wc,
            }
            kept.append(rec)
            counts["kept"] += 1

            if counts["total_scanned"] % 100 == 0:
                print(f"Scanned {counts['total_scanned']} | Kept {counts['kept']}")
        except Exception as exc:
            counts["errors"] += 1
            print(f"[WARN] Failed item at scan #{counts['total_scanned']}: {exc}")
            continue

    sanity_check(kept)

    if not args.dry_run:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        with args.out.open("w", encoding="utf-8") as fh:
            for rec in kept:
                fh.write(json.dumps(rec, ensure_ascii=False) + "\n")

    print("Summary:")
    print(f"  total scanned: {counts['total_scanned']}")
    print(f"  excluded by internal recipients: {counts['excluded_internal_recipients']}")
    print(f"  excluded by underwriting patterns: {counts['excluded_underwriting_patterns']}")
    print(f"  excluded by noise rules: {counts['excluded_noise_rules']}")
    print(f"  kept: {counts['kept']}")
    print(f"  errors: {counts['errors']}")
    if args.dry_run:
        print("Dry run mode: no output file written.")
    else:
        print(f"Wrote: {args.out}")


if __name__ == "__main__":
    main()
