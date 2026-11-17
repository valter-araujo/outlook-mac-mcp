from outlook_mac_mcp.application.event_summary import describe_moment
from outlook_mac_mcp.domain.saved_draft import SavedDraft

NO_SUBJECT = "(no subject)"
NO_RECIPIENT = "(nobody yet)"


def describe_draft_for_deletion(draft: SavedDraft) -> str:
    """Enough to tell one draft from another before it is removed: subject, To line and
    when it was last edited, which is what separates two drafts of the same reply.
    """
    to = ", ".join(recipient.address for recipient in draft.to) or NO_RECIPIENT
    return "; ".join(
        (
            draft.subject or NO_SUBJECT,
            f"to {to}",
            f"last edited {describe_moment(draft.last_modified_at)}",
            "it moves to Deleted Items, where it can still be recovered",
        )
    )
