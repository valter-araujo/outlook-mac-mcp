from pydantic import BaseModel, ConfigDict, Field

from outlook_mac_mcp.interface.mcp.email_page_view import FOLDER_DESCRIPTION


class EmailCountView(BaseModel):
    model_config = ConfigDict(frozen=True)

    total: int = Field(description="Exactly how many emails match the filters.")
    folder: str = Field(description=FOLDER_DESCRIPTION)
