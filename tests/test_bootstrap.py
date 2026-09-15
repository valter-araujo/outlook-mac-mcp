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


def settings_with_calendar_write(enabled: bool) -> Settings:
    return Settings(
        client_id="a-client-id", timezone=ZoneInfo("Europe/Lisbon"), calendar_write_enabled=enabled
    )


def test_wires_no_write_use_cases_when_the_flag_is_off() -> None:
    assert build_use_cases(settings_with_calendar_write(False)).calendar_write is None


def test_wires_the_write_use_cases_when_the_flag_is_on() -> None:
    assert build_use_cases(settings_with_calendar_write(True)).calendar_write is not None
