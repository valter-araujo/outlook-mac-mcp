r"""Live positive-control check that a quote in a contact search term cannot widen the
$filter Graph actually evaluates.

Unit tests can only assert the bytes we send. contact_query.py refuses a term holding a
single quote before it ever reaches the wire (see test_search_contacts.py), so the
production path never sends what this module sends. This is the check for what a
refusal-free version would have been exposed to: whether Graph's
`emailAddresses/any(a:a/address eq '<term>')` clause — the one place Graph's own docs
restrict filtering to `eq`, see contact_query.py — treats an embedded quote as inert
text or as the start of a new OData expression.

The probe term embeds `' or '1' eq '1`. Spliced into the template unescaped, it turns
`a/address eq '<term>'` into `(a/address eq '<term>') or ('1' eq '1')`. The second
disjunct never mentions `a`, so if Graph evaluates it as OData rather than as a literal
string it is true for every item in the collection, and `any()` then matches every
contact that has any email address at all, not only ones actually equal to the term.

Self-provisioning: it only needs the contacts folder to hold at least one contact with
at least one email address, and skips when the folder is empty. It does not depend on
any contact's actual name or address.

Results have not yet been recorded here. Run with `uv run pytest -m integration` against
a mailbox holding at least one contact with an email address, and record the date and
the measured counts in this docstring once confirmed, per the project's rule that a
property depending on how an external API parses input needs a live check before it can
be trusted.
"""

import os

import pytest

from outlook_mac_mcp.application.search_contacts_request import SearchContactsRequest
from outlook_mac_mcp.domain.errors import InvalidRequestError
from outlook_mac_mcp.infrastructure.graph.authentication import DeviceCodeAuthenticator
from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.contact_repository import GraphContactRepository
from outlook_mac_mcp.infrastructure.settings import load_settings

pytestmark = pytest.mark.integration

CONTACTS_PATH = "/me/contacts"
# No real address is expected to equal this, so a clean hit count of zero is the baseline.
HARMLESS_PROBE_TERM = "outlook-mac-mcp-live-probe-term-that-matches-nobody"
# Widens `a/address eq '<term>'` into an unconditional true if Graph parses the quote as
# OData rather than as inert text inside the string literal.
INJECTED_PROBE_TERM = f"{HARMLESS_PROBE_TERM}' or '1' eq '1"


def graph_client() -> GraphClient:
    return GraphClient(DeviceCodeAuthenticator.from_settings(load_settings(os.environ)))


def email_eq_hit_count(client: GraphClient, term: str) -> int:
    """The email clause alone, built the same way contact_query.py builds it, with the
    term spliced in unescaped exactly as it would be without the refusal.
    """
    filter_value = f"emailAddresses/any(a:a/address eq '{term}')"
    payload = client.get(
        CONTACTS_PATH, {"$filter": filter_value, "$count": "true", "$top": 1, "$select": "id"}
    )
    count = payload.get("@odata.count")
    assert isinstance(count, int)
    return count


def ensure_a_contact_has_an_email_address(client: GraphClient) -> None:
    payload = client.get(CONTACTS_PATH, {"$count": "true", "$top": 1, "$select": "id"})
    total = payload.get("@odata.count")
    assert isinstance(total, int)
    if total == 0:
        pytest.skip("the contacts folder is empty; nothing to run the probe against")


def test_a_quote_in_the_term_cannot_widen_the_email_clause_to_match_every_contact() -> None:
    client = graph_client()
    ensure_a_contact_has_an_email_address(client)

    clean = email_eq_hit_count(client, HARMLESS_PROBE_TERM)
    injected = email_eq_hit_count(client, INJECTED_PROBE_TERM)

    assert clean == 0, "the harmless probe term matched a real address; pick a rarer one"
    assert injected == clean, (
        "a quote in the term widened the eq clause into an unconditional match: "
        f"{injected} contacts matched a term that should have matched none"
    )


def test_our_own_search_refuses_the_injected_term_before_any_request() -> None:
    """InvalidRequestError, not a Graph response: proves nothing was ever sent.

    contact_query.py's refusal is what keeps the production path off the query the other
    test in this module sends directly, to measure whether it would have been safe.
    """
    repository = GraphContactRepository(graph_client())

    with pytest.raises(InvalidRequestError):
        repository.search(SearchContactsRequest(term=INJECTED_PROBE_TERM))
