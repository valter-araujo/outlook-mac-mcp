from outlook_mac_mcp.application.list_custom_folders import MAX_FOLDER_DEPTH, MAX_FOLDERS_SCANNED
from outlook_mac_mcp.application.ports.mail_folder_repository import MailFolderRepository
from outlook_mac_mcp.domain.custom_folder_scan import CustomFolderScan
from outlook_mac_mcp.domain.custom_mail_folder import CustomMailFolder
from outlook_mac_mcp.domain.errors import CustomFolderNotFoundError, InvalidRequestError
from outlook_mac_mcp.domain.folder_selection import FolderSelection
from outlook_mac_mcp.domain.resolved_folder import ResolvedFolder

# "all" is short enough that almost any three-letter custom folder name would be
# within an edit or two of it; a longer well-known name can afford a looser match
# without starting to catch unrelated custom paths.
SHORT_WELL_KNOWN_VALUE_LENGTH = 4
SHORT_VALUE_MAX_DISTANCE = 1
LONG_VALUE_MAX_DISTANCE = 2


class ResolveFolders:
    """Turns a mail tool's raw `folder` argument -- a well-known name, `all`, or a
    custom folder's full path -- into the folders to actually query.

    A close-but-not-exact miss of a well-known name (a likely typo) is rejected before
    any Graph call, naming the well-known value it was probably meant to be; this is a
    coarse edit-distance heuristic, not perfect -- a genuine custom folder whose name
    happens to be one edit away from a well-known one (e.g. a folder literally called
    "Archives") is indistinguishable from a typo of "archive" by this check alone, and
    will be rejected the same way. Anything else falls through to a custom path lookup,
    resolved fresh on every call by walking the same tree list_folders walks
    (MailFolderRepository.list_custom, sharing its depth/folder-count caps by default):
    nothing here is cached, so a custom folder argument costs one full tree walk each
    time it's used.
    """

    def __init__(
        self,
        mail_folder_repository: MailFolderRepository,
        max_depth: int = MAX_FOLDER_DEPTH,
        max_folders: int = MAX_FOLDERS_SCANNED,
    ) -> None:
        self._mail_folder_repository = mail_folder_repository
        self._max_depth = max_depth
        self._max_folders = max_folders

    def execute(self, raw: str) -> ResolvedFolder:
        try:
            return self._resolved(FolderSelection(raw))
        except ValueError:
            pass

        match = _closest_well_known_match(raw)
        if match is not None:
            candidate, distance = match
            if distance == 0:
                return self._resolved(candidate)
            raise InvalidRequestError(
                f"{raw!r} is not a recognized folder -- did you mean {candidate.value!r}?"
            )

        return self._resolve_custom_path(raw)

    def _resolved(self, selection: FolderSelection) -> ResolvedFolder:
        return ResolvedFolder(folder_ids=selection.to_folders(), echo=selection.describe_folders())

    def _resolve_custom_path(self, path: str) -> ResolvedFolder:
        scan = self._mail_folder_repository.list_custom(self._max_depth, self._max_folders)
        match = _find_by_path(scan, path)
        if match is None:
            raise CustomFolderNotFoundError(
                f"no custom folder with path {path!r} found within {self._max_depth} "
                f"levels deep and {self._max_folders} folders scanned"
            )
        return ResolvedFolder(folder_ids=(match.folder_id,), echo=match.path)


def _find_by_path(scan: CustomFolderScan, path: str) -> CustomMailFolder | None:
    return next((folder for folder in scan.folders if folder.path == path), None)


def _closest_well_known_match(raw: str) -> tuple[FolderSelection, int] | None:
    """The well-known value or `all` closest to `raw` by case-insensitive edit
    distance, and how far off it is -- or None when `raw` holds a "/" (no well-known
    value ever does, so a path segment can't be a typo of one) or nothing is close
    enough to say anything with confidence, which is the common case: almost every raw
    string reaching this point is a genuine custom folder path, not a typo.
    """
    if "/" in raw:
        return None
    needle = raw.casefold()
    candidate = min(FolderSelection, key=lambda option: _levenshtein_distance(needle, option.value))
    distance = _levenshtein_distance(needle, candidate.value)
    threshold = (
        SHORT_VALUE_MAX_DISTANCE
        if len(candidate.value) <= SHORT_WELL_KNOWN_VALUE_LENGTH
        else LONG_VALUE_MAX_DISTANCE
    )
    return (candidate, distance) if distance <= threshold else None


def _levenshtein_distance(a: str, b: str) -> int:
    """Classic DP edit distance -- the inputs here are at most a handful of short
    strings, so there is no need for anything faster.
    """
    if a == b:
        return 0
    previous_row = list(range(len(b) + 1))
    for i, char_a in enumerate(a, start=1):
        current_row = [i]
        for j, char_b in enumerate(b, start=1):
            insert_cost = current_row[j - 1] + 1
            delete_cost = previous_row[j] + 1
            substitute_cost = previous_row[j - 1] + (char_a != char_b)
            current_row.append(min(insert_cost, delete_cost, substitute_cost))
        previous_row = current_row
    return previous_row[-1]
