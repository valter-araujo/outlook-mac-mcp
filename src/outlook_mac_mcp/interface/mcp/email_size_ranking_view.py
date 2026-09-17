from datetime import datetime

from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.domain.email_size import EmailSize
from outlook_mac_mcp.domain.email_size_ranking import EmailSizeRanking
from outlook_mac_mcp.interface.mcp.email_page_view import FOLDER_DESCRIPTION


class EmailSizeView(BaseModel):
    model_config = ConfigDict(frozen=True)

    id: str
    subject: str
    sender_address: str
    sender_name: str
    received_at: datetime
    size_bytes: int = Field(
        description="The email's size in bytes, from a Graph extended property."
    )

    @classmethod
    def from_email_size(cls, item: EmailSize) -> "EmailSizeView":
        return cls(
            id=item.id,
            subject=item.subject,
            sender_address=item.sender.address,
            sender_name=item.sender.display_name,
            received_at=item.received_at,
            size_bytes=item.size_bytes,
        )


class EmailSizeRankingView(BaseModel):
    model_config = ConfigDict(frozen=True)

    items: list[EmailSizeView] = Field(description="Emails, largest first.")
    scanned: int = Field(description="How many emails were checked to build the ranking.")
    skipped: int = Field(
        description=(
            "Of the scanned emails, how many carried no determinable size and were "
            "excluded from the ranking."
        )
    )
    total: int = Field(description="Exactly how many emails match the filters.")
    coverage_is_complete: bool = Field(
        description=(
            "Whether every matching email was scanned. When false, scanned is less than "
            "total and the ranking covers only the scanned emails."
        )
    )
    folder: str = Field(description=FOLDER_DESCRIPTION)

    @classmethod
    def from_ranking(cls, ranking: EmailSizeRanking, folder: str) -> "EmailSizeRankingView":
        return cls(
            items=[EmailSizeView.from_email_size(item) for item in ranking.items],
            scanned=ranking.scanned,
            skipped=ranking.skipped,
            total=ranking.total,
            coverage_is_complete=ranking.coverage_is_complete,
            folder=folder,
        )
