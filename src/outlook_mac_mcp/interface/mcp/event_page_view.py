from outlook_mac_mcp.domain.event import Event
from outlook_mac_mcp.domain.page import Page
from outlook_mac_mcp.interface.mcp.event_view import EventView
from outlook_mac_mcp.interface.mcp.page_view import PageView


class EventPageView(PageView[EventView]):
    @classmethod
    def from_page(cls, page: Page[Event]) -> "EventPageView":
        return cls.projected(page, EventView.from_event)
