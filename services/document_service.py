from pathlib import Path

from utilities.docxDisplayHelper import (
    listAllSubfolderInRepo
)

from utilities.docx_editor import (
    get_docx_as_html,
    save_html_as_docx
)

from utilities.file_safety import (
    acquire_editing_lock,
    release_editing_lock,
    refresh_editing_lock
)


def resolve_file(subpath):

    try:

        contents, is_file = (
            listAllSubfolderInRepo(subpath)
        )

    except FileNotFoundError:

        return None, None

    if not is_file:

        return None, None

    return (
        contents["parent"],
        contents["name"]
    )


def get_document_for_edit(
    subpath,
    email
):

    parent, name = resolve_file(
        subpath
    )

    if not parent:

        return {
            "success": False,
            "not_found": True
        }

    target_path = (
        Path(parent) / name
    )

    acquired, holder = (
        acquire_editing_lock(
            target_path,
            email
        )
    )

    if not acquired:

        return {
            "success": False,
            "locked": True,
            "filename": name,
            "holder": holder
        }

    html_content, version = (
        get_docx_as_html(
            parent,
            name
        )
    )

    return {
        "success": True,
        "filename": name,
        "doc_html": html_content,
        "version": version
    }


def save_document(
    subpath,
    html_content,
    expected_version,
    email
):

    parent, name = resolve_file(
        subpath
    )

    if not parent:

        return {
            "success": False,
            "not_found": True
        }

    result = save_html_as_docx(
        parent,
        name,
        html_content,
        expected_version,
        email
    )

    if not result["success"]:

        return {
            "success": False,
            "conflict": True,
            "filename": name
        }

    release_editing_lock(
        Path(parent) / name,
        email
    )

    return {
        "success": True
    }


def refresh_document_lock(
    subpath,
    email
):

    parent, name = resolve_file(
        subpath
    )

    if not parent:

        return {
            "success": False
        }

    refresh_editing_lock(
        Path(parent) / name,
        email
    )

    return {
        "success": True
    }


def cancel_document_edit(
    subpath,
    email
):

    parent, name = resolve_file(
        subpath
    )

    if not parent:

        return {
            "success": False
        }

    release_editing_lock(
        Path(parent) / name,
        email
    )

    return {
        "success": True
    }