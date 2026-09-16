from collections.abc import Mapping
from zoneinfo import ZoneInfo

TEXT_BODY_DIRECTIVE = 'outlook.body-content-type="text"'


def timezone_preference(timezone: ZoneInfo) -> dict[str, str]:
    """Ask Graph to express every time in the response in `timezone`."""
    return {"Prefer": f'outlook.timezone="{timezone.key}"'}


def text_body_preference() -> dict[str, str]:
    """Ask Graph to return a body property as plain text rather than HTML."""
    return {"Prefer": TEXT_BODY_DIRECTIVE}


def combined_preference(*preferences: Mapping[str, str]) -> dict[str, str]:
    """Merge several single-directive Prefer headers into one, comma-separated.

    httpx sends one value per header name, and a plain dict of headers can carry only
    one "Prefer" entry, so two preferences that both matter on the same request — a time
    zone and a body content type, say — have to be combined into one value rather than
    passed as two headers. Graph accepts multiple directives in one Prefer header this way.
    """
    directives = [value for preference in preferences for value in preference.values()]
    return {"Prefer": ", ".join(directives)}
