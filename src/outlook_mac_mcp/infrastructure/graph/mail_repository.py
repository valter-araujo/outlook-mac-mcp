from collections.abc import Mapping
from http import HTTPStatus
from typing import Any
from urllib.parse import quote

from outlook_mac_mcp.application.list_emails_request import ListEmailsRequest
from outlook_mac_mcp.application.search_emails_request import SearchEmailsRequest
from outlook_mac_mcp.domain.email import Email
from outlook_mac_mcp.domain.email_detail import EmailDetail
from outlook_mac_mcp.domain.email_filters import EmailFilters
from outlook_mac_mcp.domain.email_search_page import EmailSearchPage
from outlook_mac_mcp.domain.email_size_scan import EmailSizeScan
from outlook_mac_mcp.domain.errors import EmailNotFoundError, InvalidRequestError
from outlook_mac_mcp.domain.page import Page
from outlook_mac_mcp.domain.sender_scan import SenderScan
from outlook_mac_mcp.domain.sort_order import SortOrder
from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.email_mapper import (
    MESSAGE_FIELDS,
    to_email,
    to_email_detail,
)
from outlook_mac_mcp.infrastructure.graph.email_size_scan import scan_email_sizes
from outlook_mac_mcp.infrastructure.graph.errors import GraphRequestError, GraphResponseError
from outlook_mac_mcp.infrastructure.graph.mail_query import count_query, list_query
from outlook_mac_mcp.infrastructure.graph.pagination import read_items, read_next_link
from outlook_mac_mcp.infrastructure.graph.preferences import text_body_preference
from outlook_mac_mcp.infrastructure.graph.search_match_count import MatchCount, count_search_matches
from outlook_mac_mcp.infrastructure.graph.search_query import to_search_query
from outlook_mac_mcp.infrastructure.graph.sender_scan import scan_senders

UNREAD_FILTER = "isRead eq false"
NEWEST_FIRST_ORDER = "receivedDateTime desc"
COUNT_FIELD = "@odata.count"
ID_ONLY_SELECT = "id"
# The smallest page Graph accepts; a count wants the number, not the messages.
COUNT_PAGE_SIZE = 1
DETAIL_FIELDS = (*MESSAGE_FIELDS, "body")

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

    def list_unread(self, folders: tuple[str, ...], limit: int) -> Page[Email]:
        """Nothing user-supplied is spliced into the query: each folder segment is either
        a well-known name or a folder id Graph itself produced while resolving a custom
        path, never raw caller text; the filter and order are constants, and `limit` is
        an int validated by the use case, so there is no string for a caller to break
        out of.

        `$count=true` makes Graph report how many messages match the filter in the same
        response, so the total costs no extra request. More than one folder means one such
        request per folder, merged newest-first and truncated to `limit` afterward, with
        the per-folder counts summed into one exact total.
        """
        emails: list[Email] = []
        total = 0
        for folder in folders:
            payload = self._client.get(
                _messages_path(folder),
                {
                    "$filter": UNREAD_FILTER,
                    "$orderby": NEWEST_FIRST_ORDER,
                    "$top": limit,
                    "$select": ",".join(MESSAGE_FIELDS),
                    "$count": "true",
                },
            )
            emails.extend(to_email(message) for message in read_items(payload))
            total += _read_count(payload)
        merged = _sorted_by_received_at(emails, newest_first=True)[:limit]
        return Page(items=tuple(merged), total=total, total_is_exact=True)

    def get_by_id(self, email_id: str) -> EmailDetail:
        """The id is percent-encoded before it becomes a path segment.

        Unlike the folder, it is caller-supplied: Graph ids can contain characters that
        are significant in a URL, and an unencoded one could otherwise change the path.
        """
        try:
            payload = self._client.get(
                f"/me/messages/{quote(email_id, safe='')}",
                {"$select": ",".join(DETAIL_FIELDS)},
                text_body_preference(),
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

        More than one folder in `request.filters.folders` means one such request per
        folder, merged and re-sorted afterward (each folder's own page already arrives in
        that order, so this only matters once folders are combined) and truncated to
        `request.limit`, with the per-folder counts summed into one exact total.
        """
        emails: list[Email] = []
        total = 0
        for folder in request.filters.folders:
            payload = self._client.get(
                _messages_path(folder),
                {
                    **list_query(request.filters, request.sort),
                    "$top": request.limit,
                    "$select": ",".join(MESSAGE_FIELDS),
                },
            )
            emails.extend(to_email(message) for message in read_items(payload))
            total += _read_count(payload)
        newest_first = request.sort is SortOrder.NEWEST
        merged = _sorted_by_received_at(emails, newest_first=newest_first)[: request.limit]
        return Page(items=tuple(merged), total=total, total_is_exact=True)

    def count_matching(self, filters: EmailFilters) -> int:
        total = 0
        for folder in filters.folders:
            payload = self._client.get(
                _messages_path(folder),
                {**count_query(filters), "$top": COUNT_PAGE_SIZE, "$select": ID_ONLY_SELECT},
            )
            total += _read_count(payload)
        return total

    def scan_senders(self, filters: EmailFilters, ceiling: int) -> SenderScan:
        paths = [_messages_path(folder) for folder in filters.folders]
        return scan_senders(self._client, paths, filters, ceiling)

    def scan_email_sizes(self, filters: EmailFilters, ceiling: int) -> EmailSizeScan:
        paths = [_messages_path(folder) for folder in filters.folders]
        return scan_email_sizes(self._client, paths, filters, ceiling)

    def search(self, request: SearchEmailsRequest) -> EmailSearchPage:
        """No $orderby: Graph rejects it alongside $search, so results are relevance-ranked.

        The term is sent as a quoted KQL phrase, so text that reads like a query — `from:`,
        `AND`, a stray colon — is searched for rather than executed. Any property
        restriction is built from the scope enum, never from the term.

        `page_token`, when given, is followed directly (`GraphClient.follow`, which
        validates its host the same way a fresh request is pinned to graph.microsoft.com)
        instead of the query being rebuilt: `term`/`folders`/`scope` still validate and are
        still available, so the match count below stays consistent across every page of
        the same search, however many folders it spans.

        More than one folder for a fresh search means each is searched up to `limit` and
        the results concatenated folder by folder, then truncated to `limit` overall --
        not a single merged relevance order, since Graph exposes no cross-folder score to
        merge by. A token from one folder's page is never combined with another's, so
        `next_page_token` is only ever produced for a single-folder search; the caller
        gets a fully merged (if truncated) view of `all` in one call instead.

        Graph cannot count a search, so the matches are walked separately. The cheap
        shortcut -- trust this page's own size as the total -- only holds for a single
        folder's first page that is also its last one: anything else, a later page, or
        more than one folder, gets the real count, summed across every folder searched.
        """
        search = to_search_query(request.term, request.scope)
        if request.page_token is not None:
            payload = self._client.follow(request.page_token)
            emails = tuple(to_email(message) for message in read_items(payload))
            count = _count_search_across(self._client, request.folders, search)
            return EmailSearchPage(
                items=emails,
                total=count.total,
                total_is_exact=count.is_exact,
                next_page_token=read_next_link(payload),
            )

        if len(request.folders) == 1:
            path = _messages_path(request.folders[0])
            payload = self._client.get(
                path,
                {"$search": search, "$top": request.limit, "$select": ",".join(MESSAGE_FIELDS)},
            )
            emails = tuple(to_email(message) for message in read_items(payload))
            next_page_token = read_next_link(payload)
            if next_page_token is None:
                return EmailSearchPage(
                    items=emails, total=len(emails), total_is_exact=True, next_page_token=None
                )
            count = count_search_matches(self._client, path, search)
            return EmailSearchPage(
                items=emails,
                total=count.total,
                total_is_exact=count.is_exact,
                next_page_token=next_page_token,
            )

        merged: list[Email] = []
        for folder in request.folders:
            payload = self._client.get(
                _messages_path(folder),
                {"$search": search, "$top": request.limit, "$select": ",".join(MESSAGE_FIELDS)},
            )
            merged.extend(to_email(message) for message in read_items(payload))
        count = _count_search_across(self._client, request.folders, search)
        return EmailSearchPage(
            items=tuple(merged[: request.limit]),
            total=count.total,
            total_is_exact=count.is_exact,
            next_page_token=None,
        )


def _messages_path(folder: str) -> str:
    return f"/me/mailFolders/{folder}/messages"


def _count_search_across(client: GraphClient, folders: tuple[str, ...], search: str) -> MatchCount:
    counts = [count_search_matches(client, _messages_path(folder), search) for folder in folders]
    return MatchCount(
        total=sum(count.total for count in counts),
        is_exact=all(count.is_exact for count in counts),
    )


def _sorted_by_received_at(emails: list[Email], *, newest_first: bool) -> list[Email]:
    return sorted(emails, key=lambda email: email.received_at, reverse=newest_first)


def _read_count(payload: Mapping[str, Any]) -> int:
    """A missing count is a malformed answer, not zero: it was asked for explicitly."""
    count = payload.get(COUNT_FIELD)
    if isinstance(count, bool) or not isinstance(count, int) or count < 0:
        raise GraphResponseError(f"the message collection carried no {COUNT_FIELD}")
    return count
