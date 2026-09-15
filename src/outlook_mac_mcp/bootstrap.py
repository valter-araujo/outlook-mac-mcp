from outlook_mac_mcp.application.get_email import GetEmail
from outlook_mac_mcp.application.list_unread_emails import ListUnreadEmails
from outlook_mac_mcp.application.search_emails import SearchEmails
from outlook_mac_mcp.infrastructure.graph.authentication import DeviceCodeAuthenticator
from outlook_mac_mcp.infrastructure.graph.client import GraphClient
from outlook_mac_mcp.infrastructure.graph.mail_repository import GraphMailRepository
from outlook_mac_mcp.infrastructure.settings import Settings
from outlook_mac_mcp.interface.mcp.use_cases import UseCases


def build_use_cases(settings: Settings) -> UseCases:
    """Wire the use cases to the Graph adapter.

    The composition root sits outside the layers because it is the one place allowed to
    know every one of them; nothing here decides behaviour, it only connects. One
    repository serves both use cases, so one token and one HTTP client serve the process.
    """
    authenticator = DeviceCodeAuthenticator.from_settings(settings)
    repository = GraphMailRepository(GraphClient(authenticator))
    return UseCases(
        list_unread_emails=ListUnreadEmails(repository),
        search_emails=SearchEmails(repository),
        get_email=GetEmail(repository),
    )
