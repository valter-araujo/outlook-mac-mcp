from datetime import UTC, datetime
from typing import Any

import pytest

from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.infrastructure.graph.errors import GraphResponseError
from outlook_mac_mcp.infrastructure.graph.large_email_mapper import (
    MESSAGE_SIZE_PROPERTY_ID,
    to_email_size,
)

A_MESSAGE: dict[str, Any] = {
    "id": "AAMkAGI2",
    "subject": "Quarterly review",
    "from": {"emailAddress": {"name": "Ana Lima", "address": "ana@example.com"}},
    "receivedDateTime": "2026-09-14T12:30:00Z",
    "singleValueExtendedProperties": [{"id": MESSAGE_SIZE_PROPERTY_ID, "value": "90210"}],
}


def message_without(field: str) -> dict[str, Any]:
    return {key: value for key, value in A_MESSAGE.items() if key != field}


def message_with(field: str, value: Any) -> dict[str, Any]:
    return A_MESSAGE | {field: value}


def test_maps_every_field_of_a_complete_message() -> None:
    email_size = to_email_size(A_MESSAGE)

    assert email_size.id == "AAMkAGI2"
    assert email_size.subject == "Quarterly review"
    assert email_size.sender == EmailAddress(address="ana@example.com", display_name="Ana Lima")
    assert email_size.received_at == datetime(2026, 9, 14, 12, 30, tzinfo=UTC)
    assert email_size.size_bytes == 90210


def test_parses_the_size_as_an_integer_even_though_graph_sends_it_as_a_string() -> None:
    assert isinstance(to_email_size(A_MESSAGE).size_bytes, int)


def test_finds_the_size_property_among_others() -> None:
    message = message_with(
        "singleValueExtendedProperties",
        [
            {"id": "String {66f5a359-4659-4830-9070-00047ec6ac6e} Name Color", "value": "Green"},
            {"id": MESSAGE_SIZE_PROPERTY_ID, "value": "512"},
        ],
    )

    assert to_email_size(message).size_bytes == 512


@pytest.mark.parametrize("field", ["id", "receivedDateTime"])
def test_rejects_a_message_missing_a_required_field(field: str) -> None:
    with pytest.raises(GraphResponseError):
        to_email_size(message_without(field))


def test_rejects_a_message_with_no_extended_properties_field() -> None:
    with pytest.raises(GraphResponseError):
        to_email_size(message_without("singleValueExtendedProperties"))


def test_rejects_an_empty_extended_properties_list() -> None:
    with pytest.raises(GraphResponseError):
        to_email_size(message_with("singleValueExtendedProperties", []))


def test_rejects_extended_properties_that_do_not_include_the_size() -> None:
    message = message_with(
        "singleValueExtendedProperties",
        [{"id": "String {66f5a359-4659-4830-9070-00047ec6ac6e} Name Color", "value": "Green"}],
    )

    with pytest.raises(GraphResponseError):
        to_email_size(message)


def test_rejects_a_size_value_that_is_not_an_integer() -> None:
    message = message_with(
        "singleValueExtendedProperties", [{"id": MESSAGE_SIZE_PROPERTY_ID, "value": "not-a-number"}]
    )

    with pytest.raises(GraphResponseError):
        to_email_size(message)


def test_maps_an_absent_subject_to_empty() -> None:
    assert to_email_size(message_without("subject")).subject == ""
