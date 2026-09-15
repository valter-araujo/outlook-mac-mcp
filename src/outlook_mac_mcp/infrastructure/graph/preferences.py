from zoneinfo import ZoneInfo


def timezone_preference(timezone: ZoneInfo) -> dict[str, str]:
    """Ask Graph to express every time in the response in `timezone`."""
    return {"Prefer": f'outlook.timezone="{timezone.key}"'}
