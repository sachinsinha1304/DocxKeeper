from pathlib import Path

from utilities.document_versioning import (
    restore_version, read_version_content, get_version_path
)


DOCS_ROOT = (
    Path(__file__).resolve().parent.parent
    / "documents"
)


def restore_document_version(
    subpath,
    version_id,
    restored_by
):

    target_path = (
        DOCS_ROOT / subpath
    )

    try:

        restore_version(
            target_path,
            version_id,
            restored_by=restored_by
        )

        return {
            "success": True
        }

    except Exception as e:

        return {
            "success": False,
            "error": str(e)
        }



def view_document_version(
    subpath,
    version_id
):

    target_path = DOCS_ROOT / subpath

    try:

        if not target_path.exists():

            return {
                "success": False,
                "error": "Document not found."
            }

        version_path = get_version_path(
            target_path,
            version_id
        )

        if not version_path.exists():

            return {
                "success": False,
                "error": "Requested version was not found."
            }

        content = read_version_content(
            version_path
        )

        return {
            "success": True,
            "filename": target_path.name,
            "content": content,
            "version": version_id
        }

    except Exception as e:

        return {
            "success": False,
            "error": str(e)
        }