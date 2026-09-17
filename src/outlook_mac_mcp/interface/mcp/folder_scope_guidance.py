def folder_scope_guidance() -> str:
    """The one wording every mail tool uses to state its folder scope plainly.

    A result like "showing 20 of 250" gives no indication it only covered one folder;
    this sentence and the output's own folder field are what make that explicit.
    """
    return (
        "folder accepts a well-known name (default inbox), all (the five well-known "
        "folders only -- never custom folders), or a custom folder's full path exactly "
        "as shown in list_folders' custom list, e.g. folder=Entrevistas/Work/AWS. "
        "Results are scoped to folder only: other folders are not included unless "
        "folder names one of them. The output carries folder so a reader can always "
        "tell which."
    )
