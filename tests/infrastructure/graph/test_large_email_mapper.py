from datetime import UTC, datetime
from typing import Any

import pytest

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.infrastructure.graph.errors import GraphResponseError
from outlook_mac_mcp.infrastructure.graph.large_email_mapper import (
    MESSAGE_SIZE_PROPERTY_ID_INTEGER,
    MESSAGE_SIZE_PROPERTY_ID_LONG,
    to_email_size,
)

A_MESSAGE: dict[str, Any] = {
    "id": "AAMkAGI2",
    "subject": "Quarterly review",
    "from": {"emailAddress": {"name": "Ana Lima", "address": "ana@example.com"}},
    "receivedDateTime": "2026-09-14T12:30:00Z",
    "singleValueExtendedProperties": [{"id": MESSAGE_SIZE_PROPERTY_ID_INTEGER, "value": "90210"}],
}


def message_without(field: str) -> dict[str, Any]:
    return {key: value for key, value in A_MESSAGE.items() if key != field}


def message_with(field: str, value: Any) -> dict[str, Any]:
    return A_MESSAGE | {field: value}


def test_maps_every_field_of_a_complete_message() -> None:
    email_size = to_email_size(A_MESSAGE)

    assert email_size is not None
    assert email_size.id == "AAMkAGI2"
    assert email_size.subject == "Quarterly review"
    assert email_size.sender == EmailAddress(address="ana@example.com", display_name="Ana Lima")
    assert email_size.received_at == datetime(2026, 9, 14, 12, 30, tzinfo=UTC)
    assert email_size.size_bytes == 90210


def test_parses_the_size_as_an_integer_even_though_graph_sends_it_as_a_string() -> None:
    email_size = to_email_size(A_MESSAGE)

    assert email_size is not None
    assert isinstance(email_size.size_bytes, int)


def test_reads_the_integer_typed_property() -> None:
    message = message_with(
        "singleValueExtendedProperties",
        [{"id": MESSAGE_SIZE_PROPERTY_ID_INTEGER, "value": "512"}],
    )

    email_size = to_email_size(message)

    assert email_size is not None
    assert email_size.size_bytes == 512


def test_reads_the_long_typed_property() -> None:
    """Either type keyword is accepted; see
    test_matches_the_property_exactly_as_graph_echoed_it_back_live for what Graph
    actually returns for either one.
    """
    message = message_with(
        "singleValueExtendedProperties",
        [{"id": MESSAGE_SIZE_PROPERTY_ID_LONG, "value": "4294967296"}],
    )

    email_size = to_email_size(message)

    assert email_size is not None
    assert email_size.size_bytes == 4294967296


@pytest.mark.parametrize("property_id", ["Long 0xe08", "Integer 0xe08"])
def test_matches_the_property_exactly_as_graph_echoed_it_back_live(property_id: str) -> None:
    """Live check, 2026-09-16, personal Outlook.com mailbox, one real message: requesting
    the property as 'Long 0x0E08' or 'Integer 0x0E08' both returned it with its id
    echoed back lowercase and without the leading zero, as asserted here with the exact
    strings observed — never the zero-padded, capitalized form that was requested. See
    the note on _is_message_size_property in large_email_mapper.py.
    """
    message = message_with(
        "singleValueExtendedProperties", [{"id": property_id, "value": "375436"}]
    )

    email_size = to_email_size(message)

    assert email_size is not None
    assert email_size.size_bytes == 375436


def test_does_not_match_a_different_property_sharing_the_same_tag_number() -> None:
    """The type keyword still has to be one this property is known to use; matching by
    tag number with no type check at all would be too loose.
    """
    message = message_with(
        "singleValueExtendedProperties", [{"id": "String 0xe08", "value": "not-a-size"}]
    )

    assert to_email_size(message) is None


def test_finds_the_size_property_among_others() -> None:
    message = message_with(
        "singleValueExtendedProperties",
        [
            {"id": "String {66f5a359-4659-4830-9070-00047ec6ac6e} Name Color", "value": "Green"},
            {"id": MESSAGE_SIZE_PROPERTY_ID_INTEGER, "value": "512"},
        ],
    )

    email_size = to_email_size(message)

    assert email_size is not None
    assert email_size.size_bytes == 512


def test_returns_none_when_the_message_carries_no_extended_properties_field() -> None:
    assert to_email_size(message_without("singleValueExtendedProperties")) is None


def test_returns_none_for_an_empty_extended_properties_list() -> None:
    assert to_email_size(message_with("singleValueExtendedProperties", [])) is None


def test_returns_none_when_extended_properties_do_not_include_the_size() -> None:
    message = message_with(
        "singleValueExtendedProperties",
        [{"id": "String {66f5a359-4659-4830-9070-00047ec6ac6e} Name Color", "value": "Green"}],
    )

    assert to_email_size(message) is None


def test_rejects_a_size_value_that_is_not_an_integer() -> None:
    """Presence with a garbled value is a different failure than absence: it still raises."""
    message = message_with(
        "singleValueExtendedProperties",
        [{"id": MESSAGE_SIZE_PROPERTY_ID_INTEGER, "value": "not-a-number"}],
    )

    with pytest.raises(GraphResponseError):
        to_email_size(message)


def test_rejects_a_size_value_that_is_not_a_string() -> None:
    message = message_with(
        "singleValueExtendedProperties", [{"id": MESSAGE_SIZE_PROPERTY_ID_INTEGER, "value": 512}]
    )

    with pytest.raises(GraphResponseError):
        to_email_size(message)


@pytest.mark.parametrize("field", ["id", "receivedDateTime"])
def test_rejects_a_sized_message_missing_a_required_field(field: str) -> None:
    with pytest.raises(GraphResponseError):
        to_email_size(message_without(field))


def test_maps_an_absent_subject_to_empty() -> None:
    email_size = to_email_size(message_without("subject"))

    assert email_size is not None
    assert email_size.subject == ""
