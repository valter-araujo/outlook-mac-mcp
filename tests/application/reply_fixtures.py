from outlook_mac_mcp.domain.email_address import EmailAddress
from outlook_mac_mcp.domain.reply_source import ReplySource

ME = "me@example.com"
THREAD_MESSAGE = ReplySource(
    id="AAMk1",
    subject="Next steps",
    sender=EmailAddress(address="ana@example.com", display_name="Ana"),
    reply_to=(),
    to=(EmailAddress(address=ME), EmailAddress(address="bo@example.com")),
    cc=(EmailAddress(address="cy@example.com"), EmailAddress(address="di@example.com")),
)
NO_REPLY_MESSAGE = ReplySource(
    id="AAMk2",
    subject="Re: Your application",
    sender=EmailAddress(address="recruiting@email.example.com"),
    reply_to=(),
    to=(EmailAddress(address=ME),),
    cc=(),
)
REDIRECTED_MESSAGE = ReplySource(
    id="AAMk3",
    subject="Interview",
    sender=EmailAddress(address="ats@example.com"),
    reply_to=(EmailAddress(address="recruiter@example.com"),),
    to=(EmailAddress(address=ME),),
    cc=(),
)
