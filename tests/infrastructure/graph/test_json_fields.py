from enum import StrEnum

import pytest

from outlook_mac_mcp.infrastructure.graph.errors import GraphResponseError
from outlook_mac_mcp.infrastructure.graph.json_fields import optional_enum


class Color(StrEnum):
    RED = "red"
    BLUE = "blue"


def test_reads_a_known_value() -> None:
    assert optional_enum({"color": "red"}, "color", Color) is Color.RED


def test_returns_none_when_the_field_is_absent() -> None:
    assert optional_enum({}, "color", Color) is None


def test_raises_without_echoing_the_unrecognized_value() -> None:
    """Every other GraphResponseError in this module names only the field, never the
    value carried in it; this one must not be the exception.
    """
    with pytest.raises(GraphResponseError) as excinfo:
        optional_enum({"color": "chartreuse"}, "color", Color)

    assert "color" in str(excinfo.value)
    assert "chartreuse" not in str(excinfo.value)
