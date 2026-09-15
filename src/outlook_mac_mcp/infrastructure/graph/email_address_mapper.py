from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.infrastructure.graph.json_fields import optional_text


def to_email_address(recipient: object) -> EmailAddress:
    """Read a Graph recipient, the `{emailAddress: {name, address}}` shape.

    Graph omits the sender on drafts and some system messages and the organizer on some
    events, so an absent or malformed recipient is an empty address, never an error.
    """
    if not isinstance(recipient, dict):
        return EmailAddress(address="")
    email_address = recipient.get("emailAddress")
    if not isinstance(email_address, dict):
        return EmailAddress(address="")
    return EmailAddress(
        address=optional_text(email_address, "address"),
        display_name=optional_text(email_address, "name"),
    )
