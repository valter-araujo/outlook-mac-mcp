"""Typed readers for Graph's loosely typed JSON.

The mappers are the only place `Any` is allowed, and these are the only functions that
turn it into something else, so every value the domain sees was checked exactly here.
"""

from collections.abc import Mapping
from datetime import datetime
from enum import StrEnum
from typing import Any

from outlook_mac_mcp.infrastructure.graph.errors import GraphResponseError


def required_text(resource: Mapping[str, Any], field: str) -> str:
    value = resource.get(field)
    if not isinstance(value, str) or not value:
        raise GraphResponseError(f"the resource carried no {field}")
    return value


def optional_text(resource: Mapping[str, Any], field: str) -> str:
    value = resource.get(field)
    return value if isinstance(value, str) else ""


def required_flag(resource: Mapping[str, Any], field: str) -> bool:
    value = resource.get(field)
    if not isinstance(value, bool):
        raise GraphResponseError(f"the resource carried no {field}")
    return value


def required_datetime(resource: Mapping[str, Any], field: str) -> datetime:
    """Graph's own timestamp fields: ISO 8601, time zone-aware, e.g. receivedDateTime."""
    raw = required_text(resource, field)
    try:
        value = datetime.fromisoformat(raw)
    except ValueError as error:
        raise GraphResponseError(f"{field} was not an ISO 8601 timestamp") from error
    if value.tzinfo is None:
        raise GraphResponseError(f"{field} carried no time zone")
    return value


def optional_int(resource: Mapping[str, Any], field: str) -> int | None:
    """A field a listing never selects, so its absence is normal and reads as None
    rather than an error; see optional_text_body.
    """
    value = resource.get(field)
    if value is None:
        return None
    if not isinstance(value, int) or isinstance(value, bool):
        raise GraphResponseError(f"{field} was not an integer")
    return value


def optional_enum[E: StrEnum](
    resource: Mapping[str, Any], field: str, enum_cls: type[E]
) -> E | None:
    """Same absence rule as optional_int, for a Graph enum-as-string field."""
    value = resource.get(field)
    if value is None:
        return None
    if not isinstance(value, str):
        raise GraphResponseError(f"{field} was not a string")
    try:
        return enum_cls(value)
    except ValueError as error:
        raise GraphResponseError(
            f"{field} was {value!r}, not a known {enum_cls.__name__}"
        ) from error


TEXT_CONTENT_TYPE = "text"


def optional_text_body(resource: Mapping[str, Any], field: str = "body") -> str:
    """Read a Graph `{contentType, content}` body object as plain text.

    Refuses a body that is not plain text: returning HTML as if it were text would hand
    markup to a reader that was told it is reading text, so the mismatch is surfaced
    instead of hidden. A resource with no body at all, or a resource this field was
    never requested for, is normal and reads as empty rather than an error.
    """
    body = resource.get(field)
    if not isinstance(body, dict):
        return ""
    content_type = optional_text(body, "contentType")
    if content_type and content_type != TEXT_CONTENT_TYPE:
        raise GraphResponseError(f"Graph returned a {content_type} body, not plain text")
    return optional_text(body, "content")
