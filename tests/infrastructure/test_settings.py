from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import pytest

from outlook_mac_mcp.infrastructure import settings
from outlook_mac_mcp.infrastructure.errors import ConfigurationError
from outlook_mac_mcp.infrastructure.settings import (
    CALENDAR_WRITE_ENV_VAR,
    CLIENT_ID_ENV_VAR,
    TIMEZONE_ENV_VAR,
    load_settings,
    read_calendar_write_flag,
    resolve_timezone,
)


def machine_zone_is(monkeypatch: pytest.MonkeyPatch, name: str) -> None:
    """tzlocal is the operating-system boundary, the one place a stand-in is warranted."""
    monkeypatch.setattr(settings, "get_localzone_name", lambda: name)


def test_uses_the_override_when_set(monkeypatch: pytest.MonkeyPatch) -> None:
    machine_zone_is(monkeypatch, "America/Sao_Paulo")

    zone = resolve_timezone({TIMEZONE_ENV_VAR: "Europe/Lisbon"})

    assert zone == ZoneInfo("Europe/Lisbon")


def test_falls_back_to_the_machine_zone_without_the_override(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    machine_zone_is(monkeypatch, "America/Sao_Paulo")

    assert resolve_timezone({}) == ZoneInfo("America/Sao_Paulo")


def test_treats_an_empty_override_as_absent(monkeypatch: pytest.MonkeyPatch) -> None:
    machine_zone_is(monkeypatch, "America/Sao_Paulo")

    assert resolve_timezone({TIMEZONE_ENV_VAR: ""}) == ZoneInfo("America/Sao_Paulo")


@pytest.mark.parametrize(
    "name", ["Mars/Olympus_Mons", "GMT+3", "../../etc/passwd", "/etc/localtime"]
)
def test_rejects_an_override_that_is_not_an_iana_name(name: str) -> None:
    with pytest.raises(ConfigurationError, match=TIMEZONE_ENV_VAR):
        resolve_timezone({TIMEZONE_ENV_VAR: name})


def test_rejects_a_machine_zone_that_zoneinfo_does_not_know(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    machine_zone_is(monkeypatch, "Pacific Standard Time")

    with pytest.raises(ConfigurationError, match=TIMEZONE_ENV_VAR):
        resolve_timezone({})


def test_reports_a_machine_without_a_time_zone(monkeypatch: pytest.MonkeyPatch) -> None:
    def cannot_tell() -> str:
        raise ZoneInfoNotFoundError("no zone")

    monkeypatch.setattr(settings, "get_localzone_name", cannot_tell)

    with pytest.raises(ConfigurationError, match=TIMEZONE_ENV_VAR):
        resolve_timezone({})


def test_loads_the_client_id_and_the_zone_together() -> None:
    loaded = load_settings({CLIENT_ID_ENV_VAR: "a-client-id", TIMEZONE_ENV_VAR: "Europe/Lisbon"})

    assert loaded.client_id == "a-client-id"
    assert loaded.timezone == ZoneInfo("Europe/Lisbon")
    assert loaded.calendar_write_enabled is False


def test_requires_the_client_id() -> None:
    with pytest.raises(ConfigurationError, match=CLIENT_ID_ENV_VAR):
        load_settings({TIMEZONE_ENV_VAR: "Europe/Lisbon"})


def test_calendar_write_is_off_when_the_flag_is_absent() -> None:
    assert read_calendar_write_flag({}) is False


@pytest.mark.parametrize("value", ["true", "TRUE", "1", " true "])
def test_calendar_write_is_on_for_an_explicit_true(value: str) -> None:
    assert read_calendar_write_flag({CALENDAR_WRITE_ENV_VAR: value}) is True


@pytest.mark.parametrize("value", ["false", "False", "0", ""])
def test_calendar_write_is_off_for_an_explicit_false(value: str) -> None:
    assert read_calendar_write_flag({CALENDAR_WRITE_ENV_VAR: value}) is False


@pytest.mark.parametrize("value", ["yes", "on", "enabled", "2"])
def test_rejects_a_flag_value_that_is_neither_true_nor_false(value: str) -> None:
    with pytest.raises(ConfigurationError, match=CALENDAR_WRITE_ENV_VAR):
        read_calendar_write_flag({CALENDAR_WRITE_ENV_VAR: value})


def test_loads_the_calendar_write_flag_with_the_rest() -> None:
    loaded = load_settings(
        {
            CLIENT_ID_ENV_VAR: "a-client-id",
            TIMEZONE_ENV_VAR: "Europe/Lisbon",
            CALENDAR_WRITE_ENV_VAR: "true",
        }
    )

    assert loaded.calendar_write_enabled is True
