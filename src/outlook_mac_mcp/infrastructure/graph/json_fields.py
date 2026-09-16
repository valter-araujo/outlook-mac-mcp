"""Typed readers for Graph's loosely typed JSON.

The mappers are the only place `Any` is allowed, and these are the only functions that
turn it into something else, so every value the domain sees was checked exactly here.
"""

from collections.abc import Mapping
from datetime import datetime
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
