from outlook_mac_mcp.application.list_unread_emails import ListUnreadEmails
from outlook_mac_mcp.infrastructure.graph.authentication import DeviceCodeAuthenticator
from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.mail_repository import GraphMailRepository


def build_list_unread_emails() -> ListUnreadEmails:
    """Wire the use case to the Graph adapter.

    The composition root sits outside the layers because it is the one place allowed to
    know every one of them; nothing here decides behaviour, it only connects.
    """
    authenticator = DeviceCodeAuthenticator.from_environment()
    repository = GraphMailRepository(GraphClient(authenticator))
    return ListUnreadEmails(repository)
