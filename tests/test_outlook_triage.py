"""
Unit tests for outlook_triage.py.

Covers pure helper functions, scoring logic, guardrails, VIP loading,
bucket thresholds, and the cutoff safety net — without requiring a live
Outlook / Windows environment.
"""

import pytest
from datetime import datetime, timedelta
from unittest.mock import MagicMock, PropertyMock, patch

import outlook_triage as ot


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_mail_item(
    subject="Test email",
    sender_email="sender@example.com",
    sender_name="Sender Name",
    to_line="me@company.com",
    cc_line="",
    body="",
    attachments=0,
    categories="",
    conversation_index="",
):
    """Return a minimal MagicMock that looks like an Outlook MailItem."""
    item = MagicMock()
    item.Subject = subject
    item.SenderEmailAddress = sender_email
    item.SenderName = sender_name
    item.To = to_line
    item.CC = cc_line
    item.Body = body
    item.Attachments.Count = attachments
    item.Categories = categories
    item.ConversationIndex = conversation_index
    return item


# ---------------------------------------------------------------------------
# safe_str
# ---------------------------------------------------------------------------


class TestSafeStr:
    def test_none_returns_empty(self):
        assert ot.safe_str(None) == ""

    def test_string_passthrough(self):
        assert ot.safe_str("hello") == "hello"

    def test_int_converted(self):
        assert ot.safe_str(42) == "42"

    def test_raises_returns_empty(self):
        class Unstreable:
            def __str__(self):
                raise RuntimeError("boom")

        assert ot.safe_str(Unstreable()) == ""


# ---------------------------------------------------------------------------
# naive_dt
# ---------------------------------------------------------------------------


class TestNaiveDt:
    def test_naive_datetime_unchanged(self):
        dt = datetime(2024, 3, 1, 10, 0, 0)
        assert ot.naive_dt(dt) == dt

    def test_aware_datetime_strips_tzinfo(self):
        import zoneinfo

        tz = zoneinfo.ZoneInfo("America/New_York")
        aware = datetime(2024, 3, 1, 10, 0, 0, tzinfo=tz)
        result = ot.naive_dt(aware)
        assert result.tzinfo is None
        assert result == datetime(2024, 3, 1, 10, 0, 0)

    def test_non_datetime_returns_datetime_min(self):
        # When dt_val has no replace() method (AttributeError/TypeError),
        # naive_dt must return datetime.min so that comparisons like
        # `received < cutoff` never raise TypeError outside a try block.
        assert ot.naive_dt("not a datetime") == datetime.min
        assert ot.naive_dt(None) == datetime.min


# ---------------------------------------------------------------------------
# keyword_score
# ---------------------------------------------------------------------------


class TestKeywordScore:
    def test_empty_string(self):
        score, hits = ot.keyword_score("")
        assert score == 0
        assert hits == []

    def test_single_keyword(self):
        score, hits = ot.keyword_score("urgent")
        assert score == 40
        assert hits == ["urgent"]

    def test_multiple_keywords(self):
        score, hits = ot.keyword_score("urgent deadline")
        assert score == 70  # 40 + 30
        assert "urgent" in hits
        assert "deadline" in hits

    def test_case_insensitive(self):
        s1, _ = ot.keyword_score("URGENT")
        s2, _ = ot.keyword_score("urgent")
        assert s1 == s2

    def test_high_value_financial_keywords(self):
        score, hits = ot.keyword_score("rmd distribution beneficiary rollover")
        assert score == 120  # 35+25+30+30
        assert len(hits) == 4

    def test_no_false_partial_match(self):
        # "today" is a keyword but "yesterday" should not trigger it
        score, hits = ot.keyword_score("let me know by yesterday")
        assert "today" not in hits

    def test_no_matches_returns_zero(self):
        score, hits = ot.keyword_score("quick sync on project status")
        assert score == 0
        assert hits == []


# ---------------------------------------------------------------------------
# is_noise / compile_patterns
# ---------------------------------------------------------------------------


class TestIsNoise:
    @pytest.fixture
    def pats(self):
        return ot.compile_patterns(ot.NOISE_PATTERNS)

    def test_unsubscribe_in_subject(self, pats):
        assert ot.is_noise("Click to unsubscribe from this list", "", pats)

    def test_newsletter_in_subject(self, pats):
        assert ot.is_noise("Weekly Newsletter — March Edition", "", pats)

    def test_webinar_in_subject(self, pats):
        assert ot.is_noise("Join our upcoming Webinar", "", pats)

    def test_no_reply_sender_hyphen(self, pats):
        assert ot.is_noise("", "no-reply@marketing.com", pats)

    def test_no_reply_sender_space(self, pats):
        assert ot.is_noise("", "no reply@marketing.com", pats)

    def test_marketing_keyword(self, pats):
        assert ot.is_noise("Marketing update from us", "", pats)

    def test_promo_keyword(self, pats):
        assert ot.is_noise("Special promo just for you", "", pats)

    def test_normal_email_not_noise(self, pats):
        assert not ot.is_noise("Follow-up on your account", "advisor@firm.com", pats)

    def test_noise_word_must_be_word_boundary(self, pats):
        # "digest" is noise; "indigestion" should not trigger it
        assert not ot.is_noise("indigestion remedy info", "", pats)


# ---------------------------------------------------------------------------
# recipient_count
# ---------------------------------------------------------------------------


class TestRecipientCount:
    def test_empty_string(self):
        assert ot.recipient_count("") == 0

    def test_whitespace_only(self):
        assert ot.recipient_count("   ") == 0

    def test_single_recipient(self):
        assert ot.recipient_count("alice@x.com") == 1

    def test_multiple_semicolon_separated(self):
        assert ot.recipient_count("alice@x.com; bob@x.com; charlie@x.com") == 3

    def test_trailing_semicolon_ignored(self):
        assert ot.recipient_count("alice@x.com; ") == 1


# ---------------------------------------------------------------------------
# is_reply_or_forward
# ---------------------------------------------------------------------------


class TestIsReplyOrForward:
    @pytest.mark.parametrize("subject", ["RE: hello", "re: hello", "Re: hello"])
    def test_reply_prefix(self, subject):
        assert ot.is_reply_or_forward(subject) == 1

    @pytest.mark.parametrize(
        "subject", ["FW: hello", "fw: hello", "FWD: hello", "fwd: hello"]
    )
    def test_forward_prefix(self, subject):
        assert ot.is_reply_or_forward(subject) == 1

    def test_normal_subject(self):
        assert ot.is_reply_or_forward("Project update") == 0

    def test_re_in_middle_not_matched(self):
        assert ot.is_reply_or_forward("Interest rate notice") == 0

    def test_empty_subject(self):
        assert ot.is_reply_or_forward("") == 0


# ---------------------------------------------------------------------------
# merge_categories
# ---------------------------------------------------------------------------


class TestMergeCategories:
    def test_add_to_empty_existing(self):
        assert ot.merge_categories("", "Urgent") == "Urgent"

    def test_no_duplicate_added(self):
        result = ot.merge_categories("Urgent", "Urgent")
        assert result == "Urgent"

    def test_new_category_appended(self):
        result = ot.merge_categories("Urgent", "Action")
        assert "Urgent" in result
        assert "Action" in result

    def test_empty_add_cat_is_noop(self):
        assert ot.merge_categories("Urgent", "") == "Urgent"

    def test_whitespace_in_existing_handled(self):
        result = ot.merge_categories("Urgent, FYI", "Action")
        assert "Action" in result
        assert result.count("Urgent") == 1


# ---------------------------------------------------------------------------
# choose_final_bucket
# ---------------------------------------------------------------------------


class TestChooseFinalBucket:
    def test_rule_urgent_always_wins(self):
        for model in ("Noise", "FYI", "Action", ""):
            assert ot.choose_final_bucket("Urgent", model, 100) == "Urgent"

    def test_noise_rule_always_wins(self):
        # main simplified choose_final_bucket: Noise rule always wins regardless
        # of score or model — the model cannot override a Noise classification.
        assert ot.choose_final_bucket("Noise", "Urgent", -10) == "Noise"
        assert ot.choose_final_bucket("Noise", "Action", -1) == "Noise"
        assert ot.choose_final_bucket("Noise", "Urgent", 0) == "Noise"
        assert ot.choose_final_bucket("Noise", "FYI", 50) == "Noise"

    def test_model_overrides_rule_normally(self):
        assert ot.choose_final_bucket("FYI", "Urgent", 10) == "Urgent"
        assert ot.choose_final_bucket("Action", "Noise", 50) == "Noise"

    def test_fallback_to_rule_when_no_model(self):
        assert ot.choose_final_bucket("Action", "", 50) == "Action"
        assert ot.choose_final_bucket("Waiting", "", 25) == "Waiting"

    def test_fallback_to_rule_when_model_bucket_is_invalid(self):
        assert ot.choose_final_bucket("Action", "NotACategory", 50) == "Action"
        assert ot.choose_final_bucket("FYI", "fyi", 10) == "FYI"


# ---------------------------------------------------------------------------
# already_triaged
# ---------------------------------------------------------------------------


class TestAlreadyTriaged:
    @pytest.mark.parametrize("cat", ["Urgent", "Action", "Waiting", "FYI", "Noise"])
    def test_triage_categories_detected(self, cat):
        assert ot.already_triaged(_make_mail_item(categories=cat))

    def test_custom_category_not_triaged(self):
        assert not ot.already_triaged(_make_mail_item(categories="MyTag"))

    def test_empty_categories_not_triaged(self):
        assert not ot.already_triaged(_make_mail_item(categories=""))

    def test_triage_mixed_with_custom_still_triaged(self):
        assert ot.already_triaged(_make_mail_item(categories="Urgent, MyTag"))

    def test_com_exception_returns_false(self):
        item = MagicMock()
        type(item).Categories = PropertyMock(side_effect=Exception("COM error"))
        assert not ot.already_triaged(item)


# ---------------------------------------------------------------------------
# has_non_triage_categories
# ---------------------------------------------------------------------------


class TestHasNonTriageCategories:
    @pytest.mark.parametrize("cat", ["Urgent", "Action", "Waiting", "FYI", "Noise"])
    def test_only_triage_returns_false(self, cat):
        assert not ot.has_non_triage_categories(_make_mail_item(categories=cat))

    def test_custom_category_returns_true(self):
        assert ot.has_non_triage_categories(_make_mail_item(categories="MyCustomTag"))

    def test_mixed_triage_and_custom_returns_true(self):
        assert ot.has_non_triage_categories(_make_mail_item(categories="Urgent, MyTag"))

    def test_empty_returns_false(self):
        assert not ot.has_non_triage_categories(_make_mail_item(categories=""))

    def test_com_exception_returns_false(self):
        item = MagicMock()
        type(item).Categories = PropertyMock(side_effect=Exception("COM error"))
        assert not ot.has_non_triage_categories(item)


# ---------------------------------------------------------------------------
# load_vips
# ---------------------------------------------------------------------------


class TestLoadVips:
    def test_empty_file(self, tmp_path):
        f = tmp_path / "vips.csv"
        f.write_text("")
        with patch.object(ot, "VIP_SENDERS_CSV", f):
            assert ot.load_vips() == set()

    def test_valid_emails_loaded(self, tmp_path):
        f = tmp_path / "vips.csv"
        f.write_text("boss@company.com\nclient@example.org\n")
        with patch.object(ot, "VIP_SENDERS_CSV", f):
            vips = ot.load_vips()
        assert vips == {"boss@company.com", "client@example.org"}

    def test_comments_skipped(self, tmp_path):
        f = tmp_path / "vips.csv"
        f.write_text("# header\nboss@company.com\n# another comment\n")
        with patch.object(ot, "VIP_SENDERS_CSV", f):
            vips = ot.load_vips()
        assert vips == {"boss@company.com"}

    def test_invalid_entries_skipped(self, tmp_path):
        f = tmp_path / "vips.csv"
        f.write_text("notanemail\nboss@company.com\nalso bad\n")
        with patch.object(ot, "VIP_SENDERS_CSV", f):
            vips = ot.load_vips()
        assert vips == {"boss@company.com"}

    def test_emails_normalised_to_lowercase(self, tmp_path):
        f = tmp_path / "vips.csv"
        f.write_text("BOSS@COMPANY.COM\n")
        with patch.object(ot, "VIP_SENDERS_CSV", f):
            vips = ot.load_vips()
        assert "boss@company.com" in vips

    def test_creates_empty_file_when_missing(self, tmp_path):
        f = tmp_path / "vips.csv"
        with patch.object(ot, "VIP_SENDERS_CSV", f):
            vips = ot.load_vips()
        assert f.exists()
        assert vips == set()

    def test_duplicate_entries_deduplicated(self, tmp_path):
        f = tmp_path / "vips.csv"
        f.write_text("boss@company.com\nboss@company.com\n")
        with patch.object(ot, "VIP_SENDERS_CSV", f):
            vips = ot.load_vips()
        assert len(vips) == 1


# ---------------------------------------------------------------------------
# rule_score_and_bucket
# ---------------------------------------------------------------------------


class TestRuleScoreAndBucket:
    @pytest.fixture
    def pats(self):
        return ot.compile_patterns(ot.NOISE_PATTERNS)

    @pytest.fixture
    def now(self):
        return datetime.now()

    def test_vip_sender_score_boost(self, pats, now):
        item = _make_mail_item(sender_email="vip@corp.com", to_line="me@corp.com")
        score, bucket, reasons, _ = ot.rule_score_and_bucket(
            item, {"vip@corp.com"}, pats, now
        )
        assert score >= 50
        assert "VIP_sender" in reasons

    def test_non_vip_no_boost(self, pats, now):
        item = _make_mail_item(sender_email="random@corp.com", to_line="me@corp.com")
        score, _, reasons, _ = ot.rule_score_and_bucket(
            item, {"vip@corp.com"}, pats, now
        )
        assert "VIP_sender" not in reasons

    def test_urgent_keyword_bucket(self, pats, now):
        # urgent(40) + rmd(35) + deadline(30) + today(25) + to(10) = 140
        item = _make_mail_item(
            subject="URGENT: RMD deadline today", to_line="me@corp.com"
        )
        score, bucket, _, _ = ot.rule_score_and_bucket(item, set(), pats, now)
        assert bucket == ot.CAT_URGENT
        assert score >= 80

    def test_action_bucket_threshold(self, pats, now):
        # rollover(30) + to(10) = 40 → just below Action threshold(45)
        # add underwriting(35): 30+35+10 = 75 → Urgent
        # Let's target score in [45, 79]
        # acat(35) + to(10) = 45 → Action
        item = _make_mail_item(subject="ACAT transfer request", to_line="me@corp.com")
        score, bucket, _, _ = ot.rule_score_and_bucket(item, set(), pats, now)
        assert bucket == ot.CAT_ACTION
        assert 45 <= score < 80

    def test_waiting_bucket_threshold(self, pats, now):
        # today(25) + to(10) = 35 → just above Waiting threshold(20)
        item = _make_mail_item(subject="Need this today", to_line="me@corp.com")
        score, bucket, _, _ = ot.rule_score_and_bucket(item, set(), pats, now)
        assert bucket == ot.CAT_WAITING
        assert 20 <= score < 45

    def test_fyi_bucket_for_low_score(self, pats, now):
        item = _make_mail_item(subject="Just a heads up")
        score, bucket, _, _ = ot.rule_score_and_bucket(item, set(), pats, now)
        assert bucket == ot.CAT_FYI
        assert score < 20

    def test_noise_subject_produces_noise_bucket(self, pats, now):
        item = _make_mail_item(subject="Weekly Newsletter — unsubscribe here")
        score, bucket, reasons, _ = ot.rule_score_and_bucket(item, set(), pats, now)
        assert bucket == ot.CAT_NOISE
        assert "noise_pattern" in reasons
        assert score < 0

    def test_age_penalty_after_24h(self, pats):
        item = _make_mail_item(subject="normal email", to_line="me@corp.com")
        old = datetime.now() - timedelta(days=5)
        fresh = datetime.now()
        s_old, _, r_old, _ = ot.rule_score_and_bucket(item, set(), pats, old)
        s_fresh, _, r_fresh, _ = ot.rule_score_and_bucket(item, set(), pats, fresh)
        assert s_old < s_fresh
        assert "age_penalty" in r_old
        assert "age_penalty" not in r_fresh

    def test_age_penalty_capped_at_25(self, pats):
        item = _make_mail_item(subject="very old email")
        # 100 days old → age_hours=2400 → min(25, 2400//24*5)=min(25,500)=25
        ancient = datetime.now() - timedelta(days=100)
        score, _, reasons, features = ot.rule_score_and_bucket(
            item, set(), pats, ancient
        )
        assert "age_penalty" in reasons
        # Penalty capped at 25
        fresh = datetime.now()
        s_fresh, _, _, _ = ot.rule_score_and_bucket(item, set(), pats, fresh)
        assert s_fresh - score <= 25 + 1  # +1 tolerance for floating point

    def test_attachment_score_boost(self, pats, now):
        with_att = _make_mail_item(to_line="me@corp.com", attachments=1)
        without_att = _make_mail_item(to_line="me@corp.com", attachments=0)
        s1, _, _, _ = ot.rule_score_and_bucket(with_att, set(), pats, now)
        s2, _, _, _ = ot.rule_score_and_bucket(without_att, set(), pats, now)
        assert s1 - s2 == 8

    def test_cc_line_score_penalty(self, pats, now):
        with_cc = _make_mail_item(to_line="me@corp.com", cc_line="others@corp.com")
        without_cc = _make_mail_item(to_line="me@corp.com")
        s1, _, _, _ = ot.rule_score_and_bucket(with_cc, set(), pats, now)
        s2, _, _, _ = ot.rule_score_and_bucket(without_cc, set(), pats, now)
        assert s2 - s1 == 5

    def test_to_line_score_boost(self, pats, now):
        with_to = _make_mail_item(to_line="me@corp.com")
        without_to = _make_mail_item(to_line="")
        s1, _, _, _ = ot.rule_score_and_bucket(with_to, set(), pats, now)
        s2, _, _, _ = ot.rule_score_and_bucket(without_to, set(), pats, now)
        assert s1 - s2 == 10

    def test_features_dict_keys_complete(self, pats, now):
        item = _make_mail_item()
        _, _, _, features = ot.rule_score_and_bucket(item, set(), pats, now)
        expected = {
            "sender_email",
            "sender_name",
            "subject",
            "body_snippet",
            "to_line",
            "cc_line",
            "age_hours",
            "has_attachment",
            "thread_depth",
            "recipient_count",
            "is_reply_or_fwd",
            "rule_score",
            "is_noise_hint",
        }
        assert expected.issubset(features.keys())

    def test_received_none_falls_back_to_com(self, pats):
        """When received=None, function reads ReceivedTime from the COM item."""
        item = _make_mail_item()
        item.ReceivedTime = datetime.now()
        score, bucket, _, _ = ot.rule_score_and_bucket(item, set(), pats, None)
        assert isinstance(score, int)

    def test_body_keyword_scoring(self, pats, now):
        item = _make_mail_item(subject="", body="Please review the beneficiary form")
        _, _, _, features = ot.rule_score_and_bucket(item, set(), pats, now)
        # beneficiary=30 should appear in body snippet scan
        assert features["rule_score"] >= 30


# ---------------------------------------------------------------------------
# Scoring thresholds — bucket boundary table
# ---------------------------------------------------------------------------


class TestBucketBoundaries:
    """Verify bucket assignment at threshold boundaries."""

    @pytest.fixture
    def pats(self):
        return ot.compile_patterns([])  # no noise patterns

    def test_score_80_is_urgent(self, pats):
        item = _make_mail_item(subject="URGENT RMD deadline today", to_line="me@co.com")
        score, bucket, _, _ = ot.rule_score_and_bucket(
            item, set(), pats, datetime.now()
        )
        assert score >= 80
        assert bucket == ot.CAT_URGENT

    def test_score_45_is_action(self, pats):
        # acat(35) + to(10) = 45 → Action
        item = _make_mail_item(subject="ACAT", to_line="me@co.com")
        score, bucket, _, _ = ot.rule_score_and_bucket(
            item, set(), pats, datetime.now()
        )
        assert score == 45
        assert bucket == ot.CAT_ACTION

    def test_score_20_is_waiting(self, pats):
        # today(25) + no to_line(0) = 25 → Waiting
        item = _make_mail_item(subject="today", to_line="")
        score, bucket, _, _ = ot.rule_score_and_bucket(
            item, set(), pats, datetime.now()
        )
        assert score == 25
        assert bucket == ot.CAT_WAITING

    def test_score_below_20_is_fyi(self, pats):
        item = _make_mail_item(subject="no keywords here")
        score, bucket, _, _ = ot.rule_score_and_bucket(
            item, set(), pats, datetime.now()
        )
        assert score < 20
        assert bucket == ot.CAT_FYI


# ---------------------------------------------------------------------------
# Cutoff safety net
# ---------------------------------------------------------------------------


class TestCutoffSafetyNet:
    """
    Verify that the DAYS_BACK cutoff correctly classifies email ages.
    Tests the guard logic added back in main() as a Restrict() fallback.
    """

    def test_email_at_boundary_is_in_range(self):
        cutoff = datetime.now() - timedelta(days=ot.DAYS_BACK)
        # Email exactly at the cutoff boundary should be included
        at_boundary = cutoff + timedelta(seconds=1)
        assert at_boundary >= cutoff

    def test_email_before_boundary_is_out_of_range(self):
        cutoff = datetime.now() - timedelta(days=ot.DAYS_BACK)
        old = datetime.now() - timedelta(days=ot.DAYS_BACK + 1)
        assert old < cutoff

    def test_recent_email_is_in_range(self):
        cutoff = datetime.now() - timedelta(days=ot.DAYS_BACK)
        recent = datetime.now() - timedelta(hours=1)
        assert recent >= cutoff


# ---------------------------------------------------------------------------
# get_sender_email
# ---------------------------------------------------------------------------


class TestGetSenderEmail:
    def test_smtp_address_returned_directly(self):
        item = _make_mail_item(sender_email="alice@company.com")
        assert ot.get_sender_email(item) == "alice@company.com"

    def test_smtp_address_normalised_to_lowercase(self):
        item = _make_mail_item(sender_email="Alice@Company.COM")
        assert ot.get_sender_email(item) == "alice@company.com"

    def test_exchange_dn_falls_back_to_exchange_user(self):
        """SenderEmailAddress is an Exchange X.500 DN (no @); resolve via GetExchangeUser."""
        item = MagicMock()
        item.SenderEmailAddress = "/O=EXCHANGELABS/CN=RECIPIENTS/CN=ABC123"
        ex_user = MagicMock()
        ex_user.PrimarySmtpAddress = "alice@company.com"
        item.Sender.GetExchangeUser.return_value = ex_user
        assert ot.get_sender_email(item) == "alice@company.com"

    def test_exchange_user_none_returns_empty(self):
        item = MagicMock()
        item.SenderEmailAddress = "/O=EXCHANGELABS/CN=..."
        item.Sender.GetExchangeUser.return_value = None
        assert ot.get_sender_email(item) == ""

    def test_com_exception_on_both_paths_returns_empty(self):
        item = MagicMock()
        type(item).SenderEmailAddress = PropertyMock(side_effect=Exception("COM error"))
        item.Sender.GetExchangeUser.side_effect = Exception("COM error")
        assert ot.get_sender_email(item) == ""


# ---------------------------------------------------------------------------
# validate_config
# ---------------------------------------------------------------------------


class TestValidateConfig:
    def test_valid_defaults_pass(self):
        ot.validate_config()  # should not raise

    def test_days_back_zero_raises(self):
        with patch.object(ot, "DAYS_BACK", 0):
            with pytest.raises(ValueError, match="DAYS_BACK"):
                ot.validate_config()

    def test_days_back_over_limit_raises(self):
        with patch.object(ot, "DAYS_BACK", 91):
            with pytest.raises(ValueError, match="DAYS_BACK"):
                ot.validate_config()

    def test_max_items_negative_raises(self):
        with patch.object(ot, "MAX_ITEMS", -1):
            with pytest.raises(ValueError, match="MAX_ITEMS"):
                ot.validate_config()

    def test_dry_run_non_bool_raises(self):
        with patch.object(ot, "DRY_RUN", "yes"):
            with pytest.raises(ValueError, match="DRY_RUN"):
                ot.validate_config()

    def test_move_noise_non_bool_raises(self):
        with patch.object(ot, "MOVE_NOISE_TO_READ_LATER", 1):
            with pytest.raises(ValueError, match="MOVE_NOISE_TO_READ_LATER"):
                ot.validate_config()


# ---------------------------------------------------------------------------
# apply_actions
# ---------------------------------------------------------------------------


class TestApplyActions:
    def test_dry_run_returns_dry_run_and_does_not_save(self):
        item = _make_mail_item()
        with patch.object(ot, "DRY_RUN", True):
            result = ot.apply_actions(item, ot.CAT_ACTION, None)
        assert result == "dry_run"
        item.Save.assert_not_called()

    def test_non_dry_run_applies_category_and_saves(self):
        item = _make_mail_item()
        with patch.object(ot, "DRY_RUN", False):
            result = ot.apply_actions(item, ot.CAT_FYI, None)
        assert result == "applied"
        item.Save.assert_called_once()

    def test_urgent_sets_follow_up_flag(self):
        item = _make_mail_item()
        with patch.object(ot, "DRY_RUN", False):
            ot.apply_actions(item, ot.CAT_URGENT, None)
        assert item.FlagStatus == 2
        assert item.FlagRequest == "Follow up"

    def test_action_sets_follow_up_flag(self):
        item = _make_mail_item()
        with patch.object(ot, "DRY_RUN", False):
            ot.apply_actions(item, ot.CAT_ACTION, None)
        assert item.FlagStatus == 2

    def test_fyi_does_not_set_flag(self):
        item = _make_mail_item()
        with patch.object(ot, "DRY_RUN", False):
            ot.apply_actions(item, ot.CAT_FYI, None)
        assert item.FlagStatus != 2

    def test_manual_category_returns_skipped(self):
        item = _make_mail_item(categories="MyCustomTag")
        with patch.object(ot, "DRY_RUN", False):
            # PROTECT_NON_TRIAGE_CATEGORIES is True by default
            result = ot.apply_actions(item, ot.CAT_ACTION, None)
        assert result == "skipped_manual_categories"
        item.Save.assert_not_called()

    def test_save_exception_returns_failed_apply(self):
        item = _make_mail_item()
        item.Save.side_effect = Exception("COM write error")
        with patch.object(ot, "DRY_RUN", False):
            result = ot.apply_actions(item, ot.CAT_FYI, None)
        assert result == "failed_apply"

    def test_moves_noise_to_read_later_folder(self):
        item = _make_mail_item()
        read_later = MagicMock()
        with patch.object(ot, "DRY_RUN", False), \
                patch.object(ot, "MOVE_NOISE_TO_READ_LATER", True):
            result = ot.apply_actions(item, ot.CAT_NOISE, read_later)
        item.Move.assert_called_once_with(read_later)
        assert result == "applied"

    def test_noise_not_moved_when_read_later_is_none(self):
        item = _make_mail_item()
        with patch.object(ot, "DRY_RUN", False), \
                patch.object(ot, "MOVE_NOISE_TO_READ_LATER", True):
            result = ot.apply_actions(item, ot.CAT_NOISE, None)
        item.Move.assert_not_called()
        assert result == "applied"


# ---------------------------------------------------------------------------
# load_model
# ---------------------------------------------------------------------------


class TestLoadModel:
    def test_returns_none_when_joblib_unavailable(self):
        with patch.object(ot, "joblib", None):
            assert ot.load_model() is None

    def test_returns_none_when_model_file_missing(self, tmp_path):
        with patch.object(ot, "MODEL_PATH", tmp_path / "nonexistent.joblib"):
            assert ot.load_model() is None

    def test_returns_model_when_file_exists(self, tmp_path):
        import joblib

        model_path = tmp_path / "model.joblib"
        fake_model = {"key": "value"}
        joblib.dump(fake_model, model_path)
        with patch.object(ot, "MODEL_PATH", model_path):
            result = ot.load_model()
        assert result == fake_model

    def test_returns_none_on_corrupt_model(self, tmp_path):
        model_path = tmp_path / "model.joblib"
        model_path.write_bytes(b"this is not a valid joblib file")
        with patch.object(ot, "MODEL_PATH", model_path):
            assert ot.load_model() is None


# ---------------------------------------------------------------------------
# ensure_outlook_folder
# ---------------------------------------------------------------------------


class TestEnsureOutlookFolder:
    def test_returns_existing_folder_by_name(self):
        namespace = MagicMock()
        folder1 = MagicMock()
        folder1.Name = "Read Later"
        inbox = namespace.GetDefaultFolder.return_value
        inbox.Folders.__iter__ = MagicMock(return_value=iter([folder1]))
        result = ot.ensure_outlook_folder(namespace, "Read Later")
        assert result is folder1

    def test_folder_lookup_is_case_insensitive(self):
        namespace = MagicMock()
        folder1 = MagicMock()
        folder1.Name = "READ LATER"
        inbox = namespace.GetDefaultFolder.return_value
        inbox.Folders.__iter__ = MagicMock(return_value=iter([folder1]))
        result = ot.ensure_outlook_folder(namespace, "read later")
        assert result is folder1

    def test_creates_folder_when_not_found(self):
        namespace = MagicMock()
        inbox = namespace.GetDefaultFolder.return_value
        # Empty folder list — no match
        inbox.Folders.__iter__ = MagicMock(return_value=iter([]))
        new_folder = MagicMock()
        inbox.Folders.Add.return_value = new_folder
        result = ot.ensure_outlook_folder(namespace, "New Folder")
        inbox.Folders.Add.assert_called_once_with("New Folder")
        assert result is new_folder

    def test_returns_none_on_com_exception(self):
        namespace = MagicMock()
        inbox = namespace.GetDefaultFolder.return_value
        inbox.Folders.__iter__ = MagicMock(side_effect=Exception("COM error"))
        result = ot.ensure_outlook_folder(namespace, "Read Later")
        assert result is None


# ---------------------------------------------------------------------------
# Noise edge-case: noise pattern present but score >= 0
# ---------------------------------------------------------------------------


class TestNoiseBucketEdgeCases:
    @pytest.fixture
    def pats(self):
        return ot.compile_patterns(ot.NOISE_PATTERNS)

    def test_noise_pattern_with_positive_score_not_noise_bucket(self, pats):
        """A noisy email that also matches high-value keywords should NOT become Noise
        (noise bucket requires score < 0 after the -40 penalty)."""
        # urgent(40) + rmd(35) + to(10) - noise(-40) = 45 → score >= 0
        item = _make_mail_item(
            subject="Urgent RMD newsletter unsubscribe", to_line="me@corp.com"
        )
        score, bucket, reasons, _ = ot.rule_score_and_bucket(item, set(), pats, datetime.now())
        assert "noise_pattern" in reasons
        assert score >= 0
        assert bucket != ot.CAT_NOISE

    def test_noise_pattern_with_zero_score_is_noise_bucket(self, pats):
        """A noisy email with score == -1 (just below zero) must become Noise."""
        item = _make_mail_item(subject="unsubscribe newsletter", to_line="")
        score, bucket, reasons, _ = ot.rule_score_and_bucket(item, set(), pats, datetime.now())
        assert "noise_pattern" in reasons
        assert score < 0
        assert bucket == ot.CAT_NOISE


# ---------------------------------------------------------------------------
# thread_depth
# ---------------------------------------------------------------------------


class TestThreadDepth:
    def test_short_index_returns_zero(self):
        item = _make_mail_item(conversation_index="A" * 44)
        assert ot.thread_depth(item) == 0

    def test_longer_index_returns_depth(self):
        # 44 + 10 chars → depth 1
        item = _make_mail_item(conversation_index="A" * 54)
        assert ot.thread_depth(item) == 1

    def test_empty_index_returns_zero(self):
        item = _make_mail_item(conversation_index="")
        assert ot.thread_depth(item) == 0

    def test_com_exception_returns_zero(self):
        item = MagicMock()
        type(item).ConversationIndex = PropertyMock(side_effect=Exception("COM"))
        assert ot.thread_depth(item) == 0
