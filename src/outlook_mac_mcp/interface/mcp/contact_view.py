from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.domain.contact import Contact


class ContactView(BaseModel):
    """JSON-serializable projection of a Contact.

    Email addresses are flattened to a plain list: nothing downstream needs the
    per-address display name Graph also carries.
    """

    model_config = ConfigDict(frozen=True)

    id: str
    display_name: str
    email_addresses: list[str] = Field(description="The contact's email addresses.")

    @classmethod
    def from_contact(cls, contact: Contact) -> "ContactView":
        return cls(
            id=contact.id,
            display_name=contact.display_name,
            email_addresses=[address.address for address in contact.email_addresses],
        )
