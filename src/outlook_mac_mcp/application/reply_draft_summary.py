from outlook_mac_mcp.application.event_summary import summarize_body
from outlook_mac_mcp.application.link_warnings import link_warnings
from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.new_reply_draft import NewReplyDraft
from outlook_mac_mcp.domain.reply_mode import ReplyMode
from outlook_mac_mcp.domain.reply_recipients import address_key
from outlook_mac_mcp.domain.reply_source import ReplySource
from outlook_mac_mcp.domain.unmonitored_address import looks_unmonitored

REPLY_PREFIX = "Re: "
MODE_LABELS = {ReplyMode.REPLY: "reply", ReplyMode.REPLY_ALL: "reply all"}


def reply_subject(subject: str) -> str:
    """The subject the saved reply will carry: one prefix, never "Re: Re:"."""
    if subject.casefold().startswith(REPLY_PREFIX.strip().casefold()):
        return subject
    return REPLY_PREFIX + subject


def describe_reply_draft(source: ReplySource, draft: NewReplyDraft) -> str:
    """One line a person can check before anything is saved.

    Every recipient is listed in full, since on a reply all the recipient list is the
    part the user did not type and most needs to see.
    """
    parts = [
        reply_subject(source.subject),
        MODE_LABELS[draft.mode],
        "to " + _addresses(draft.recipients.to),
    ]
    if draft.recipients.cc:
        parts.append("cc " + _addresses(draft.recipients.cc))
    parts.append(f"body: {summarize_body(draft.body)}" if draft.body else "body: (none)")
    parts.append("the original message is quoted below the body")
    return "; ".join(parts)


def reply_draft_warnings(source: ReplySource, draft: NewReplyDraft) -> tuple[str, ...]:
    warnings = [_reply_to_warning(source)] if _redirects_reply(source) else []
    recipients = (*draft.recipients.to, *draft.recipients.cc)
    warnings.extend(
        f"{recipient.address} looks like an automated or no-reply address; "
        "a reply there may never be read"
        for recipient in recipients
        if looks_unmonitored(recipient.address)
    )
    warnings.extend(link_warnings(draft.body))
    return tuple(warnings)


def _redirects_reply(source: ReplySource) -> bool:
    sender = address_key(source.sender.address)
    return any(address_key(recipient.address) != sender for recipient in source.reply_to)


def _reply_to_warning(source: ReplySource) -> str:
    return (
        f"the message asks for replies to go to {_addresses(source.reply_to)} "
        f"instead of its sender {source.sender.address or '(unknown)'}; the reply follows that"
    )


def _addresses(recipients: tuple[EmailAddress, ...]) -> str:
    return ", ".join(recipient.address for recipient in recipients)
