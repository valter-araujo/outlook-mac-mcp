r"""Live positive-control check that Graph parses our $search as a literal phrase.

Unit tests can only assert the bytes we send. They all passed while a term's operators
were being executed by Graph, so this is the check that would have caught it.

The control is self-provisioning: a common term is searched for, the sender of the first
hit becomes the control, and `<term> AND from:<sender>` is then sent both ways. The naive
form is the one we used to send. If Graph executes its operators the result is filtered
and the count drops; if the phrase is literal, nothing matches it at all.

Measured 2026-09-14 against a personal Outlook.com mailbox:

    "\"http\""                                                   100 hits
    "http AND from:notification@smartrecruiters.com"               1 hit
    "\"http AND from:notification@smartrecruiters.com\""           0 hits

Run with: uv run pytest -m integration
"""

from collections.abc import Mapping
from typing import Any

import pytest

from outlook_mac_mcp.application.search_emails_request import SearchEmailsRequest
from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.domain.search_scope import SearchScope
from outlook_mac_mcp.infrastructure.graph.authentication import DeviceCodeAuthenticator
from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.mail_repository import GraphMailRepository
from outlook_mac_mcp.infrastructure.graph.search_query import to_search_query

pytestmark = pytest.mark.integration

INBOX_MESSAGES = "/me/mailFolders/inbox/messages"
PAGE_SIZE = 100

# Tried in order until one matches something, so the check does not depend on any
# particular mailbox. `http` appears in almost any message carrying a link.
CANDIDATE_TERMS = ("http", "the", "email", "de")


def graph_client() -> GraphClient:
    return GraphClient(DeviceCodeAuthenticator.from_environment())


def messages_matching(client: GraphClient, search_value: str) -> list[Mapping[str, Any]]:
    payload = client.get(
        INBOX_MESSAGES, {"$search": search_value, "$top": PAGE_SIZE, "$select": "id,from"}
    )
    found = payload.get("value")
    assert isinstance(found, list)
    return found


def sender_of(message: Mapping[str, Any]) -> str:
    address = message.get("from", {}).get("emailAddress", {}).get("address", "")
    return address if isinstance(address, str) else ""


def seed_from_mailbox(client: GraphClient) -> tuple[str, str]:
    """Find a term that matches something, and a sender to use as the control.

    Taking the sender from a message that already matched the term guarantees the control
    matches at least one message, so a naive search cannot come back empty by accident.
    """
    for term in CANDIDATE_TERMS:
        for message in messages_matching(client, to_search_query(term, SearchScope.ANY)):
            sender = sender_of(message)
            if sender:
                return term, sender
    pytest.skip("no candidate term matched a message with a sender in this mailbox")


def test_the_nested_phrase_is_not_parsed_as_a_query() -> None:
    client = graph_client()
    term, control_sender = seed_from_mailbox(client)
    injected = f"{term} AND from:{control_sender}"

    literal_term = len(messages_matching(client, to_search_query(term, SearchScope.ANY)))
    naive = len(messages_matching(client, f'"{injected}"'))
    nested = len(messages_matching(client, to_search_query(injected, SearchScope.ANY)))

    assert naive >= 1, "the control matched nothing, so the comparison would prove nothing"
    assert naive < literal_term, "the control did not fire: the naive form was not filtered"
    assert nested < naive, "our form was parsed as a query, not as literal text"


def test_a_scoped_search_is_also_literal() -> None:
    client = graph_client()
    term, control_sender = seed_from_mailbox(client)
    injected = f"{term} AND from:{control_sender}"

    naive = len(messages_matching(client, f'"{injected}"'))
    nested = len(messages_matching(client, to_search_query(injected, SearchScope.SUBJECT)))

    assert nested < naive


def test_a_term_ending_in_a_backslash_is_refused_before_any_request() -> None:
    """InvalidRequestError, not GraphRequestError: our own validation stopped it.

    A term that reached Graph and was rejected there would arrive as GraphRequestError,
    so the type of the failure is what proves nothing was sent.
    """
    repository = GraphMailRepository(graph_client())

    with pytest.raises(InvalidRequestError):
        repository.search(SearchEmailsRequest(term="path\\"))
