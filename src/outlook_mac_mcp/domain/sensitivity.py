from enum import StrEnum


class Sensitivity(StrEnum):
    """Graph's event.sensitivity values, used verbatim so the adapter needs no mapping."""

    NORMAL = "normal"
    PERSONAL = "personal"
    PRIVATE = "private"
    CONFIDENTIAL = "confidential"
