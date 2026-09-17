"""Walking the mailbox's custom folder tree.

Graph exposes children via `/mailFolders/{id}/childFolders`, one level at a time and
paginated like any other collection; nothing aggregates the whole tree in one call. The
walk starts from each well-known folder and from the mailbox root (Graph's own
"msgfolderroot" alias for it), since a user-created folder can sit directly under
either. A folder found while walking the root that turns out to be one of the five
well-known folders is skipped there -- its own subtree is covered by its own walk, not
duplicated under the root's. A folder that is neither well-known nor user-created (a
Microsoft-managed folder such as Deleted Items) carries no such flag in Graph's
response, so it is discovered and reported the same as a genuine custom folder.

Depth and folder count are both capped, sharing one budget across every root, the same
way sender_scan.py shares one message ceiling across every folder path: once the budget
is spent, every further root and every further page is left unwalked, and that is
reported rather than distinguished from the coincidence of the budget landing exactly
on the tree's last folder.
"""

from collections.abc import Mapping
from dataclasses import dataclass, field
from time import perf_counter
from typing import Any
from urllib.parse import quote

from outlook_mac_mcp.domain.custom_folder_scan import CustomFolderScan
from outlook_mac_mcp.domain.custom_mail_folder import CustomMailFolder
from outlook_mac_mcp.domain.folder_name import FolderName
from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.errors import GraphResponseError
from outlook_mac_mcp.infrastructure.graph.json_fields import required_text
from outlook_mac_mcp.infrastructure.graph.pagination import read_items, read_next_link
from outlook_mac_mcp.logger import project_logger

MAILBOX_ROOT_ID = "msgfolderroot"
WELL_KNOWN_SUMMARY_FIELDS = ("id", "displayName")
CHILD_FIELDS = ("id", "displayName", "unreadItemCount", "totalItemCount", "childFolderCount")
# Comfortably above the folder-count cap most callers will use, so a mailbox within the
# cap is walked in one page per folder.
CHILD_PAGE_SIZE = 200
CUSTOM_FOLDER_SCAN_EVENT = "custom_folder_scan"


@dataclass
class _WellKnownFolder:
    folder_id: str
    display_name: str


@dataclass
class _Budget:
    remaining: int
    limit_reached: bool = False


@dataclass
class _ScanState:
    budget: _Budget
    folders: list[CustomMailFolder] = field(default_factory=list)
    depth_limit_reached: bool = False


def scan_custom_folders(client: GraphClient, max_depth: int, max_folders: int) -> CustomFolderScan:
    """Only counts and timings are logged: a folder's display name is mailbox content."""
    started_at = perf_counter()
    well_known = _read_well_known_folders(client)
    state = _ScanState(budget=_Budget(remaining=max_folders))

    well_known_ids = frozenset(folder.folder_id for folder in well_known.values())
    _walk_children(client, MAILBOX_ROOT_ID, None, 1, max_depth, state, skip_ids=well_known_ids)
    for folder in well_known.values():
        if state.budget.remaining <= 0:
            state.budget.limit_reached = True
            break
        _walk_children(client, folder.folder_id, folder.display_name, 1, max_depth, state)

    _log(len(state.folders), started_at)
    return CustomFolderScan(
        folders=tuple(state.folders),
        depth_limit_reached=state.depth_limit_reached,
        folder_limit_reached=state.budget.limit_reached,
    )


def _read_well_known_folders(client: GraphClient) -> dict[FolderName, _WellKnownFolder]:
    summaries: dict[FolderName, _WellKnownFolder] = {}
    for name in FolderName:
        payload = client.get(
            f"/me/mailFolders/{name.value}", {"$select": ",".join(WELL_KNOWN_SUMMARY_FIELDS)}
        )
        summaries[name] = _WellKnownFolder(
            folder_id=required_text(payload, "id"),
            display_name=required_text(payload, "displayName"),
        )
    return summaries


def _walk_children(
    client: GraphClient,
    parent_id: str,
    path_prefix: str | None,
    depth: int,
    max_depth: int,
    state: _ScanState,
    skip_ids: frozenset[str] = frozenset(),
) -> None:
    payload = client.get(
        f"/me/mailFolders/{quote(parent_id, safe='')}/childFolders",
        {"$top": CHILD_PAGE_SIZE, "$select": ",".join(CHILD_FIELDS)},
    )
    while True:
        if not _collect_page(client, payload, path_prefix, depth, max_depth, state, skip_ids):
            return
        next_link = read_next_link(payload)
        if next_link is None:
            return
        if state.budget.remaining <= 0:
            state.budget.limit_reached = True
            return
        payload = client.follow(next_link)


def _collect_page(
    client: GraphClient,
    payload: Mapping[str, Any],
    path_prefix: str | None,
    depth: int,
    max_depth: int,
    state: _ScanState,
    skip_ids: frozenset[str],
) -> bool:
    """Add this page's folders up to the budget; return whether the caller may still
    look for a next page, which is false once the budget ran out mid-page.
    """
    for item in read_items(payload):
        if state.budget.remaining <= 0:
            state.budget.limit_reached = True
            return False
        folder_id = required_text(item, "id")
        if folder_id in skip_ids:
            continue
        _record_folder(client, item, folder_id, path_prefix, depth, max_depth, state)
    return True


def _record_folder(
    client: GraphClient,
    item: Mapping[str, Any],
    folder_id: str,
    path_prefix: str | None,
    depth: int,
    max_depth: int,
    state: _ScanState,
) -> None:
    display_name = required_text(item, "displayName")
    path = display_name if path_prefix is None else f"{path_prefix}/{display_name}"
    state.folders.append(
        CustomMailFolder(
            folder_id=folder_id,
            display_name=display_name,
            path=path,
            unread_count=_required_count(item, "unreadItemCount"),
            total_count=_required_count(item, "totalItemCount"),
        )
    )
    state.budget.remaining -= 1
    if _required_count(item, "childFolderCount") == 0:
        return
    if depth >= max_depth:
        state.depth_limit_reached = True
        return
    if state.budget.remaining <= 0:
        state.budget.limit_reached = True
        return
    _walk_children(client, folder_id, path, depth + 1, max_depth, state)


def _required_count(resource: Mapping[str, Any], field_name: str) -> int:
    value = resource.get(field_name)
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise GraphResponseError(f"the folder carried no {field_name}")
    return value


def _log(folders_found: int, started_at: float) -> None:
    project_logger().info(
        CUSTOM_FOLDER_SCAN_EVENT,
        extra={
            "fields": {
                "folders_found": folders_found,
                "duration_ms": round((perf_counter() - started_at) * 1000, 3),
            }
        },
    )
