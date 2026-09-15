"""What the process needs from its environment, read once and validated at startup.

The MCP server and the sign-in command load the same settings, so a bad value fails the
cheap command too, before a browser or the Keychain is involved.
"""

from collections.abc import Mapping
from dataclasses import dataclass
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from tzlocal import get_localzone_name

from outlook_mac_mcp.infrastructure.errors import ConfigurationError

CLIENT_ID_ENV_VAR = "OUTLOOK_MCP_CLIENT_ID"
TIMEZONE_ENV_VAR = "OUTLOOK_MCP_TIMEZONE"
CALENDAR_WRITE_ENV_VAR = "OUTLOOK_MCP_ENABLE_CALENDAR_WRITE"

# Only these spellings, so a typo such as "yes" or "on" cannot silently leave writes off
# when the user meant them on, nor be read as on when they meant off.
ENABLED_VALUES = frozenset({"true", "1"})
DISABLED_VALUES = frozenset({"false", "0", ""})


@dataclass(frozen=True, slots=True)
class Settings:
    client_id: str
    timezone: ZoneInfo
    calendar_write_enabled: bool


def load_settings(environment: Mapping[str, str]) -> Settings:
    return Settings(
        client_id=_read_client_id(environment),
        timezone=resolve_timezone(environment),
        calendar_write_enabled=read_calendar_write_flag(environment),
    )


def read_calendar_write_flag(environment: Mapping[str, str]) -> bool:
    """Absent means off: a write capability must be asked for, never assumed."""
    value = environment.get(CALENDAR_WRITE_ENV_VAR, "").strip().lower()
    if value in ENABLED_VALUES:
        return True
    if value in DISABLED_VALUES:
        return False
    raise ConfigurationError(f"{CALENDAR_WRITE_ENV_VAR} must be true or false, not {value!r}")


def resolve_timezone(environment: Mapping[str, str]) -> ZoneInfo:
    """The override wins when set; otherwise the machine's zone, as the user sees it.

    Either name goes through zoneinfo, which is what the calendar code will use, so a
    name that resolves here resolves everywhere and a name that does not fails now.
    """
    override = environment.get(TIMEZONE_ENV_VAR, "")
    if override:
        return _zone_named(override, f"{TIMEZONE_ENV_VAR} is not an IANA time zone: {override!r}")
    machine_zone = _machine_zone_name()
    return _zone_named(
        machine_zone,
        f"the machine's time zone {machine_zone!r} is unknown to zoneinfo; set {TIMEZONE_ENV_VAR}",
    )


def _read_client_id(environment: Mapping[str, str]) -> str:
    client_id = environment.get(CLIENT_ID_ENV_VAR, "")
    if not client_id:
        raise ConfigurationError(f"{CLIENT_ID_ENV_VAR} is not set")
    return client_id


def _machine_zone_name() -> str:
    try:
        return get_localzone_name()
    except ZoneInfoNotFoundError as error:
        raise ConfigurationError(
            f"could not determine the machine's time zone; set {TIMEZONE_ENV_VAR}"
        ) from error


def _zone_named(name: str, failure_message: str) -> ZoneInfo:
    """ValueError is zoneinfo's answer to a path-like key, so it is a bad name too."""
    try:
        return ZoneInfo(name)
    except (ZoneInfoNotFoundError, ValueError) as error:
        raise ConfigurationError(failure_message) from error
