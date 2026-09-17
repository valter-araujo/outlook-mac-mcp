from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

from outlook_mac_mcp.application.draft_store import DraftStore
from outlook_mac_mcp.application.event_draft import EventDraft
from outlook_mac_mcp.application.event_summary import MAX_BODY_IN_SUMMARY
from outlook_mac_mcp.application.preview_event import PreviewEvent
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.new_event import NewEvent

SAO_PAULO = ZoneInfo("America/Sao_Paulo")
NINE = datetime(2026, 9, 15, 9, tzinfo=SAO_PAULO)
MIDNIGHT = datetime(2026, 9, 15, tzinfo=SAO_PAULO)


def all_day(days: int) -> NewEvent:
    return NewEvent(
        subject="Offsite", start=MIDNIGHT, end=MIDNIGHT + timedelta(days=days), is_all_day=True
    )


def test_parks_the_event_in_the_store_under_the_returned_token() -> None:
    store: DraftStore[EventDraft] = DraftStore()
    event = NewEvent(subject="Planning", start=NINE, end=NINE + timedelta(hours=1))

    draft = PreviewEvent(store).execute(event)

    assert store.take(draft.token).new_event == event


def test_describes_a_timed_event_with_its_offset_place_and_people() -> None:
    event = NewEvent(
        subject="Planning",
        start=NINE,
        end=NINE + timedelta(hours=1),
        location="Room 1",
        attendees=(EmailAddress(address="ana@example.com"), EmailAddress(address="bo@example.com")),
    )

    summary = PreviewEvent(DraftStore()).execute(event).summary

    assert summary == (
        "Planning; Tue 15 Sep 2026 09:00 to 10:00 (UTC-03:00); at Room 1; "
        "with ana@example.com, bo@example.com"
    )


def test_describes_an_event_that_crosses_midnight_with_both_days() -> None:
    event = NewEvent(
        subject="Flight", start=NINE + timedelta(hours=14), end=NINE + timedelta(hours=17)
    )

    summary = PreviewEvent(DraftStore()).execute(event).summary

    assert "Tue 15 Sep 2026 23:00 to Wed 16 Sep 2026 02:00" in summary


def test_describes_a_single_all_day_event_by_its_day() -> None:
    event = NewEvent(
        subject="Offsite", start=MIDNIGHT, end=MIDNIGHT + timedelta(days=1), is_all_day=True
    )

    assert (
        PreviewEvent(DraftStore()).execute(event).summary == "Offsite; all day on Tue 15 Sep 2026"
    )


def test_describes_a_multi_day_all_day_event_with_its_last_day_inclusive() -> None:
    event = NewEvent(
        subject="Offsite", start=MIDNIGHT, end=MIDNIGHT + timedelta(days=3), is_all_day=True
    )

    summary = PreviewEvent(DraftStore()).execute(event).summary

    assert summary == "Offsite; all day from Tue 15 Sep 2026 to Thu 17 Sep 2026"


def test_omits_the_body_clause_when_there_is_no_body() -> None:
    event = NewEvent(subject="Planning", start=NINE, end=NINE + timedelta(hours=1))

    summary = PreviewEvent(DraftStore()).execute(event).summary

    assert "body" not in summary


def test_includes_a_short_body_in_full() -> None:
    event = NewEvent(
        subject="Planning", start=NINE, end=NINE + timedelta(hours=1), body="Bring the deck."
    )

    summary = PreviewEvent(DraftStore()).execute(event).summary

    assert summary.endswith('body: "Bring the deck."')


def test_includes_a_body_at_exactly_the_summary_limit_in_full() -> None:
    body = "a" * MAX_BODY_IN_SUMMARY
    event = NewEvent(subject="Planning", start=NINE, end=NINE + timedelta(hours=1), body=body)

    summary = PreviewEvent(DraftStore()).execute(event).summary

    assert summary.endswith(f'body: "{body}"')
    assert "truncated" not in summary


def test_truncates_a_long_body_with_an_explicit_character_count() -> None:
    body = "a" * (MAX_BODY_IN_SUMMARY + 1)
    event = NewEvent(subject="Planning", start=NINE, end=NINE + timedelta(hours=1), body=body)

    summary = PreviewEvent(DraftStore()).execute(event).summary

    assert f'body: "{"a" * MAX_BODY_IN_SUMMARY}…"' in summary
    assert f"{len(body)} characters total" in summary
    assert "truncated" in summary
    # The preview -> token -> create contract only holds if nothing beyond the
    # truncation point silently vanishes from what the user is shown as a count.
    assert body not in summary
