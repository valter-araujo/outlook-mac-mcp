def folder_scope_guidance() -> str:
    """The one wording every mail tool uses to state its folder scope plainly.

    A result like "showing 20 of 250" gives no indication it only covered one folder;
    this sentence and the output's own folder field are what make that explicit.
    """
    return (
        "Results are scoped to folder (default inbox) only: other folders are not "
        "included unless folder names one of them, or folder=all is passed to search "
        "every well-known folder and merge the results. The output carries folder so a "
        "reader can always tell which."
    )
