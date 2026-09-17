from zoneinfo import ZoneInfo

import pytest

from outlook_mac_mcp.bootstrap import build_use_cases
from outlook_mac_mcp.infrastructure.graph.authentication import DeviceCodeAuthenticator
from outlook_mac_mcp.infrastructure.settings import Settings


class NeverAskedAuthenticator:
    def get_access_token(self) -> str:
        raise AssertionError("wiring must not touch the identity platform")


@pytest.fixture(autouse=True)
def keep_the_keychain_out(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(
        DeviceCodeAuthenticator,
        "from_settings",
        classmethod(lambda cls, settings: NeverAskedAuthenticator()),
    )


def settings_with_flags(*, write_enabled: bool, delete_enabled: bool = False) -> Settings:
    return Settings(
        client_id="a-client-id",
        timezone=ZoneInfo("Europe/Lisbon"),
        calendar_write_enabled=write_enabled,
        calendar_delete_enabled=delete_enabled,
    )


def test_wires_no_write_use_cases_when_the_flag_is_off() -> None:
    assert build_use_cases(settings_with_flags(write_enabled=False)).calendar_write is None


def test_wires_the_write_use_cases_when_the_flag_is_on() -> None:
    assert build_use_cases(settings_with_flags(write_enabled=True)).calendar_write is not None


def test_wires_no_deletion_use_cases_when_only_write_is_on() -> None:
    write = build_use_cases(
        settings_with_flags(write_enabled=True, delete_enabled=False)
    ).calendar_write

    assert write is not None
    assert write.create_event is not None
    assert write.update_event is not None
    assert write.preview_event_deletion is None
    assert write.delete_event is None


def test_wires_the_deletion_use_cases_when_both_flags_are_on() -> None:
    write = build_use_cases(
        settings_with_flags(write_enabled=True, delete_enabled=True)
    ).calendar_write

    assert write is not None
    assert write.preview_event_deletion is not None
    assert write.delete_event is not None


def test_wires_nothing_when_write_is_off_even_if_delete_is_on() -> None:
    """Deletion is built inside `_calendar_write`, so it inherits write's own gate: the
    delete flag alone, without write, resolves no scope and wires no use case at all.
    """
    assert (
        build_use_cases(
            settings_with_flags(write_enabled=False, delete_enabled=True)
        ).calendar_write
        is None
    )
