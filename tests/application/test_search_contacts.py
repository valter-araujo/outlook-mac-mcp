import pytest

from outlook_mac_mcp.application.limits import MAX_LIMIT, MIN_LIMIT
from outlook_mac_mcp.application.search_contacts import SearchContacts
from outlook_mac_mcp.application.search_contacts_request import (
    MAX_TERM_LENGTH,
    MIN_TERM_LENGTH,
    SearchContactsRequest,
)
from outlook_mac_mcp.domain.contact import Contact
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.errors import InvalidRequestError
from tests.fakes.in_memory_contact_repository import InMemoryContactRepository


def make_contact(contact_id: str, display_name: str, *addresses: str) -> Contact:
    return Contact(
        id=contact_id,
        display_name=display_name,
        email_addresses=tuple(EmailAddress(address=address) for address in addresses),
    )


def repository_with(*contacts: Contact) -> InMemoryContactRepository:
    repository = InMemoryContactRepository()
    for contact in contacts:
        repository.add(contact)
    return repository


def test_matches_a_display_name_prefix() -> None:
    repository = repository_with(
        make_contact("1", "Ana Lima", "ana@x.io"), make_contact("2", "Bo", "bo@x.io")
    )

    page = SearchContacts(repository).execute(SearchContactsRequest(term="Ana"))

    assert [contact.id for contact in page.items] == ["1"]


def test_matches_a_display_name_prefix_case_insensitively() -> None:
    repository = repository_with(make_contact("1", "Ana Lima", "ana@x.io"))

    page = SearchContacts(repository).execute(SearchContactsRequest(term="ana l"))

    assert [contact.id for contact in page.items] == ["1"]


def test_matches_an_exact_email_address_but_not_a_prefix_of_one() -> None:
    repository = repository_with(make_contact("1", "Ana Lima", "ana@x.io"))

    exact = SearchContacts(repository).execute(SearchContactsRequest(term="ana@x.io"))
    prefix = SearchContacts(repository).execute(SearchContactsRequest(term="ana@"))

    assert [contact.id for contact in exact.items] == ["1"]
    assert prefix.items == ()


def test_matches_any_of_a_contacts_several_addresses() -> None:
    repository = repository_with(make_contact("1", "Ana Lima", "ana@x.io", "ana.lima@y.io"))

    page = SearchContacts(repository).execute(SearchContactsRequest(term="ana.lima@y.io"))

    assert [contact.id for contact in page.items] == ["1"]


def test_returns_empty_when_nothing_matches() -> None:
    repository = repository_with(make_contact("1", "Ana Lima", "ana@x.io"))

    page = SearchContacts(repository).execute(SearchContactsRequest(term="zed"))

    assert page.items == ()
    assert page.total == 0
    assert page.total_is_exact is True


def test_reports_the_exact_total_beyond_the_page() -> None:
    repository = repository_with(*(make_contact(str(n), "Ana", f"ana{n}@x.io") for n in range(5)))

    page = SearchContacts(repository).execute(SearchContactsRequest(term="Ana", limit=2))

    assert len(page.items) == 2
    assert page.total == 5
    assert page.total_is_exact is True


@pytest.mark.parametrize("length", [MIN_TERM_LENGTH, MAX_TERM_LENGTH])
def test_accepts_a_term_at_the_edge_of_the_allowed_length(length: int) -> None:
    assert SearchContactsRequest(term="a" * length).term


@pytest.mark.parametrize("length", [MIN_TERM_LENGTH - 1, MAX_TERM_LENGTH + 1])
def test_rejects_a_term_outside_the_allowed_length(length: int) -> None:
    with pytest.raises(InvalidRequestError):
        SearchContactsRequest(term="a" * length)


def test_rejects_a_term_holding_a_single_quote() -> None:
    with pytest.raises(InvalidRequestError):
        SearchContactsRequest(term="o'neil")


@pytest.mark.parametrize("limit", [MIN_LIMIT - 1, MAX_LIMIT + 1])
def test_rejects_a_limit_outside_the_shared_bounds(limit: int) -> None:
    with pytest.raises(InvalidRequestError):
        SearchContactsRequest(term="ana", limit=limit)
