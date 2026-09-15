from collections.abc import Mapping
from http import HTTPStatus
from typing import Any
from urllib.parse import quote

from outlook_mac_mcp.application.list_emails_request import ListEmailsRequest
from outlook_mac_mcp.application.search_emails_request import SearchEmailsRequest
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_detail import EmailDetail
from outlook_mac_mcp.domain.email_filters import EmailFilters
from outlook_mac_mcp.domain.errors import EmailNotFoundError, InvalidRequestError
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.domain.page import Page
from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.email_mapper import (
    MESSAGE_FIELDS,
    to_email,
    to_email_detail,
)
from outlook_mac_mcp.infrastructure.graph.errors import GraphRequestError, GraphResponseError
from outlook_mac_mcp.infrastructure.graph.mail_query import count_query, list_query
from outlook_mac_mcp.infrastructure.graph.pagination import read_items, read_next_link
from outlook_mac_mcp.infrastructure.graph.search_match_count import count_search_matches
from outlook_mac_mcp.infrastructure.graph.search_query import to_search_query

UNREAD_FILTER = "isRead eq false"
NEWEST_FIRST_ORDER = "receivedDateTime desc"
COUNT_FIELD = "@odata.count"
ID_ONLY_SELECT = "id"
# The smallest page Graph accepts; a count wants the number, not the messages.
COUNT_PAGE_SIZE = 1
DETAIL_FIELDS = (*MESSAGE_FIELDS, "body")
BODY_AS_TEXT_HEADER = {"Prefer": 'outlook.body-content-type="text"'}

# Graph answers 400 with this code for an id that is not a well-formed message id, and
# 404 only for a well-formed id that no longer resolves. Matching the code rather than
# the message keeps this from breaking when Microsoft rewords the text.
MALFORMED_ID_ERROR_CODE = "ErrorInvalidIdMalformed"


class GraphMailRepository:
    """Reads mail from Microsoft Graph for the signed-in account.

    Satisfies the MailRepository port structurally; the use cases never import this module.

    `$top` is a page size, not a cap: with a `$filter` Graph may return fewer items than
    asked while more mail exists behind `@odata.nextLink`. Only the first page is read,
    so `limit` bounds what comes back; the page's total is what says how much there is.
    """

    def __init__(self, client: GraphClient) -> None:
        self._client = client

    def list_unread(self, folder: FolderName, limit: int) -> Page[Email]:
        """Nothing user-supplied is spliced into the query: the folder segment comes from a
        closed enum of well-known names, the filter and order are constants, and `limit` is an
        int validated by the use case, so there is no string for a caller to break out of.

        `$count=true` makes Graph report how many messages match the filter in the same
        response, so the total costs no extra request.
        """
        payload = self._client.get(
            f"/me/mailFolders/{folder.value}/messages",
            {
                "$filter": UNREAD_FILTER,
                "$orderby": NEWEST_FIRST_ORDER,
                "$top": limit,
                "$select": ",".join(MESSAGE_FIELDS),
                "$count": "true",
            },
        )
        emails = tuple(to_email(message) for message in read_items(payload))
        return Page(items=emails, total=_read_count(payload), total_is_exact=True)

    def get_by_id(self, email_id: str) -> EmailDetail:
        """The id is percent-encoded before it becomes a path segment.

        Unlike the folder, it is caller-supplied: Graph ids can contain characters that
        are significant in a URL, and an unencoded one could otherwise change the path.
        """
        try:
            payload = self._client.get(
                f"/me/messages/{quote(email_id, safe='')}",
                {"$select": ",".join(DETAIL_FIELDS)},
                BODY_AS_TEXT_HEADER,
            )
        except GraphRequestError as error:
            if error.error_code == MALFORMED_ID_ERROR_CODE:
                raise InvalidRequestError(f"malformed email id: {email_id}") from error
            if error.status_code == HTTPStatus.NOT_FOUND:
                raise EmailNotFoundError(f"no email with id {email_id}") from error
            raise
        return to_email_detail(payload)

    def list_matching(self, request: ListEmailsRequest) -> Page[Email]:
        """The filter, order and count come from one builder that keeps Graph's rule about
        $orderby properties leading the $filter; see mail_query for the rule and the sentinel.
        """
        payload = self._client.get(
            _messages_path(request.filters.folder),
            {
                **list_query(request.filters, request.sort),
                "$top": request.limit,
                "$select": ",".join(MESSAGE_FIELDS),
            },
        )
        emails = tuple(to_email(message) for message in read_items(payload))
        return Page(items=emails, total=_read_count(payload), total_is_exact=True)

    def count_matching(self, filters: EmailFilters) -> int:
        payload = self._client.get(
            _messages_path(filters.folder),
            {**count_query(filters), "$top": COUNT_PAGE_SIZE, "$select": ID_ONLY_SELECT},
        )
        return _read_count(payload)

    def search(self, request: SearchEmailsRequest) -> Page[Email]:
        """No $orderby: Graph rejects it alongside $search, so results are relevance-ranked.

        The term is sent as a quoted KQL phrase, so text that reads like a query — `from:`,
        `AND`, a stray colon — is searched for rather than executed. Any property
        restriction is built from the scope enum, never from the term.

        Graph cannot count a search, so the matches are walked separately, and only when
        the page is not the last one: a page with no next link already holds every match.
        """
        path = f"/me/mailFolders/{request.folder.value}/messages"
        search = to_search_query(request.term, request.scope)
        payload = self._client.get(
            path,
            {"$search": search, "$top": request.limit, "$select": ",".join(MESSAGE_FIELDS)},
        )
        emails = tuple(to_email(message) for message in read_items(payload))
        if read_next_link(payload) is None:
            return Page(items=emails, total=len(emails), total_is_exact=True)
        count = count_search_matches(self._client, path, search)
        return Page(items=emails, total=count.total, total_is_exact=count.is_exact)


def _messages_path(folder: FolderName) -> str:
    return f"/me/mailFolders/{folder.value}/messages"


def _read_count(payload: Mapping[str, Any]) -> int:
    """A missing count is a malformed answer, not zero: it was asked for explicitly."""
    count = payload.get(COUNT_FIELD)
    if isinstance(count, bool) or not isinstance(count, int) or count < 0:
        raise GraphResponseError(f"the message collection carried no {COUNT_FIELD}")
    return count
