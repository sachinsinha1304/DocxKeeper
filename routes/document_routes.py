from flask import (
    Blueprint,
    render_template,
    request,
    session,
    redirect,
    url_for
)

from utilities.auth import login_required

from services.document_service import (
    get_document_for_edit,
    save_document,
    refresh_document_lock,
    cancel_document_edit
)


document_bp = Blueprint(
    "document",
    __name__
)


# ============================================================
# Edit document page
# ============================================================

@document_bp.route(
    "/edit-docx/<path:subpath>",
    methods=["GET"]
)
@login_required
def edit_docx_view(subpath):

    email = session["email"]

    result = get_document_for_edit(
        subpath,
        email
    )

    # File doesn't exist
    if result.get("not_found"):

        return render_template(
            "not_found.html",
            subpath=subpath
        ), 404

    # Someone else is editing
    if result.get("locked"):

        return render_template(
            "edit_locked.html",
            filename=result["filename"],
            subpath=subpath,
            holder=result["holder"]
        )

    return render_template(
        "edit_docx.html",
        filename=result["filename"],
        subpath=subpath,
        doc_html=result["doc_html"],
        version=result["version"]
    )


# ============================================================
# Save document
# ============================================================

@document_bp.route(
    "/edit-docx/<path:subpath>",
    methods=["POST"]
)
@login_required
def edit_docx_save(subpath):

    html_content = request.form.get(
        "doc_html",
        ""
    )

    expected_version = request.form.get(
        "version"
    )

    email = session["email"]

    result = save_document(
        subpath,
        html_content,
        expected_version,
        email
    )

    # Document doesn't exist
    if result.get("not_found"):

        return render_template(
            "not_found.html",
            subpath=subpath
        ), 404

    # Version conflict
    if result.get("conflict"):

        return render_template(
            "edit_conflict.html",
            filename=result["filename"],
            subpath=subpath,
            doc_html=html_content
        ), 409

    return redirect(
        url_for(
            "repository.show_docx_in_repo",
            subpath=subpath
        )
    )


# ============================================================
# Heartbeat
# ============================================================

@document_bp.route(
    "/edit-docx/<path:subpath>/heartbeat",
    methods=["POST"]
)
@login_required
def edit_docx_heartbeat(subpath):

    email = session["email"]

    result = refresh_document_lock(
        subpath,
        email
    )

    if not result["success"]:

        return "", 404

    return "", 204


# ============================================================
# Cancel editing
# ============================================================

@document_bp.route(
    "/edit-docx/<path:subpath>/cancel",
    methods=["POST"]
)
@login_required
def edit_docx_cancel(subpath):

    email = session["email"]

    result = cancel_document_edit(
        subpath,
        email
    )

    if not result["success"]:

        return "", 404

    return "", 204