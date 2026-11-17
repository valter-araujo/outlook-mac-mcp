from outlook_mac_mcp.application.event_summary import summarize_body
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.new_mail_draft import NewMailDraft


def describe_mail_draft(draft: NewMailDraft) -> str:
    """One line a person can check against what they meant before anything is saved.

    Every recipient is listed and the body goes through the same truncate-and-count rule
    as an event's, so nothing that will be written is hidden from the confirmation.
    """
    parts = [draft.subject, "to " + _addresses(draft.to)]
    if draft.cc:
        parts.append("cc " + _addresses(draft.cc))
    parts.append(f"body: {summarize_body(draft.body)}" if draft.body else "body: (none)")
    return "; ".join(parts)


def _addresses(recipients: tuple[EmailAddress, ...]) -> str:
    return ", ".join(recipient.address for recipient in recipients)
