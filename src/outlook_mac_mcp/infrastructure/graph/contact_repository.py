from collections.abc import Mapping
from typing import Any

from outlook_mac_mcp.application.search_contacts_request import SearchContactsRequest
from outlook_mac_mcp.domain.contact import Contact
from outlook_mac_mcp.domain.page import Page
from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.contact_mapper import CONTACT_FIELDS, to_contact
from outlook_mac_mcp.infrastructure.graph.contact_query import contact_search_query
from outlook_mac_mcp.infrastructure.graph.errors import GraphResponseError
from outlook_mac_mcp.infrastructure.graph.pagination import read_items

CONTACTS_PATH = "/me/contacts"
COUNT_FIELD = "@odata.count"


class GraphContactRepository:
    """Reads contacts from Microsoft Graph for the signed-in account.

    Satisfies the ContactRepository port structurally.
    """

    def __init__(self, client: GraphClient) -> None:
        self._client = client

    def search(self, request: SearchContactsRequest) -> Page[Contact]:
        payload = self._client.get(
            CONTACTS_PATH,
            {
                **contact_search_query(request.term),
                "$top": request.limit,
                "$select": ",".join(CONTACT_FIELDS),
            },
        )
        contacts = tuple(to_contact(item) for item in read_items(payload))
        return Page(items=contacts, total=_read_count(payload), total_is_exact=True)


def _read_count(payload: Mapping[str, Any]) -> int:
    count = payload.get(COUNT_FIELD)
    if isinstance(count, bool) or not isinstance(count, int) or count < 0:
        raise GraphResponseError(f"the contact collection carried no {COUNT_FIELD}")
    return count
