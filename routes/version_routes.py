from flask import (
    Blueprint,
    request,
    session,
    redirect,
    url_for,
    flash, 
    render_template
)

from utilities.auth import login_required

from services.version_service import (
    restore_document_version, view_document_version
)


version_bp = Blueprint(
    "version",
    __name__
)


@version_bp.route(
    "/restore-docx/<version_id>",
    methods=["POST"]
)
@login_required
def restore_docx_route(version_id):
    subpath = request.form.get(
        "subpath"
    )
    if not subpath:
        flash(
            "Invalid document path.",
            "danger"
        )
        return redirect(
            url_for(
                "repository.show_docx_in_repo",
                subpath=""
            )
        )

    current_user = session["email"]
    result = restore_document_version(
        subpath,
        version_id,
        current_user
    )
    if result["success"]:
        flash(
            "Successfully restored version!",
            "success"
        )

    else:

        flash(
            f"Error restoring version: "
            f"{result['error']}",
            "danger"
        )

    return redirect(
        url_for(
            "repository.show_docx_in_repo",
            subpath=subpath
        )
    )


@version_bp.route(
    "/view-docx/<path:subpath>/<version_id>",
    methods=["GET"]
)
@login_required
def view_specific_version(subpath, version_id):
    print(subpath)
    print(version_id)

    if not subpath:
        flash(
            "Invalid document path.",
            "danger"
        )

        return redirect(
            url_for(
                "repository.show_docx_in_repo",
                subpath=""
            )
        )

    result = view_document_version(
        subpath,
        version_id
    )
    print("==================================================")
    print(f"{result}")

    if not result["success"]:

        flash(
            result["error"],
            "danger"
        )

        return redirect(
            url_for(
                "repository.show_docx_in_repo",
                subpath=subpath
            )
        )

    return render_template(
        "view_version.html",
        filename=result["filename"],
        subpath=subpath,
        version_id=version_id,
        content=result["content"],
        version=result["version"]
    )