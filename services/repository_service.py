from pathlib import Path

from utilities.docxDisplayHelper import (
    listAllTeamFolder,
    listAllSubfolderInRepo,
    readDocxContent
)

from utilities.docx_create import (
    resolve_folder_path,
    create_new_repo,
    create_new_folder,
    create_new_docx
)

from utilities.document_versioning import list_versions


DOCS_ROOT = (
    Path(__file__).resolve().parent.parent
    / "documents"
)


def get_root_repositories():

    folder_list, is_empty = listAllTeamFolder()

    return {
        "folder_list": folder_list,
        "is_empty": is_empty
    }


def get_repository_contents(subpath):

    try:

        contents, is_file = (
            listAllSubfolderInRepo(subpath)
        )

    except FileNotFoundError:

        return {
            "success": False,
            "error": "Path not found"
        }

    # Folder
    if not is_file:

        return {
            "success": True,
            "is_file": False,
            "contents": contents,
            "is_empty": len(contents) == 0
        }

    # File
    parent = contents["parent"]
    filename = contents["name"]

    target_path = DOCS_ROOT / subpath

    # Read document
    file_content = readDocxContent(
        parent,
        filename
    )

    # Get versions
    versions = list_versions(
        target_path
    )

    return {
        "success": True,
        "is_file": True,
        "parent": parent,
        "filename": filename,
        "subpath": subpath,
        "content": file_content,
        "versions": versions
    }


def create_repository(new_name):

    return create_new_repo(
        new_name
    )


def create_folder(subpath, new_name):

    try:

        directory_path = resolve_folder_path(
            subpath
        )

    except FileNotFoundError:

        return {
            "success": False,
            "not_found": True,
            "message": "Folder not found"
        }

    success, message, final_name = (
        create_new_folder(
            directory_path,
            new_name
        )
    )

    return {
        "success": success,
        "message": message,
        "final_name": final_name,
        "not_found": False
    }


def create_document(subpath, new_name, email):

    try:

        directory_path = resolve_folder_path(
            subpath
        )

    except FileNotFoundError:

        return {
            "success": False,
            "not_found": True,
            "message": "Folder not found"
        }

    success, message, final_name = (
        create_new_docx(
            directory_path,
            new_name,
            email
        )
    )

    return {
        "success": success,
        "message": message,
        "final_name": final_name,
        "not_found": False
    }