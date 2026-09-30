from flask import (
    Blueprint,
    render_template,
    request,
    redirect,
    url_for,
    session
)

from utilities.auth import login_required

from services.repository_service import (
    get_root_repositories,
    get_repository_contents,
    create_repository,
    create_folder,
    create_document
)


repository_bp = Blueprint(
    "repository",
    __name__,
    url_prefix="/all-docx"
)


# ============================================================
# Root repository page
# ============================================================

@repository_bp.route("")
@login_required
def show_docx():

    result = get_root_repositories()

    return render_template(
        "all_docx.html",
        folder_list=result["folder_list"],
        isEmpty=result["is_empty"]
    )


# ============================================================
# Repository / folder / document browsing
# ============================================================

@repository_bp.route("/", defaults={"subpath": ""})
@repository_bp.route("/<path:subpath>")
@login_required
def show_docx_in_repo(subpath):

    result = get_repository_contents(
        subpath
    )

    if not result["success"]:

        return render_template(
            "not_found.html",
            subpath=subpath
        ), 404

    # Document
    if result["is_file"]:

        return render_template(
            "view_file.html",
            filename=result["filename"],
            parent=result["parent"],
            subpath=result["subpath"],
            content=result["content"],
            versions=result["versions"]
        )

    # Folder
    return render_template(
        "repo.html",
        folder_list=result["contents"],
        isEmpty=result["is_empty"],
        current_path=subpath
    )


# ============================================================
# Create repository
# ============================================================

@repository_bp.route(
    "/create-repo",
    methods=["POST"]
)
@login_required
def create_repo_route():

    new_name = request.form.get(
        "new_repo_name",
        ""
    )

    success, message, final_name = (
        create_repository(new_name)
    )

    if not success:

        result = get_root_repositories()

        return render_template(
            "all_docx.html",
            folder_list=result["folder_list"],
            isEmpty=result["is_empty"],
            error=message
        ), 400

    return redirect(
        url_for("repository.show_docx")
    )


# ============================================================
# Create folder
# ============================================================

@repository_bp.route(
    "/<path:subpath>/create-folder",
    methods=["POST"]
)
@login_required
def create_folder_route(subpath):

    new_name = request.form.get(
        "new_folder_name",
        ""
    )

    result = create_folder(
        subpath,
        new_name
    )

    if result["not_found"]:

        return render_template(
            "not_found.html",
            subpath=subpath
        ), 404

    if not result["success"]:

        contents_result = (
            get_repository_contents(subpath)
        )

        return render_template(
            "repo.html",
            folder_list=contents_result["contents"],
            isEmpty=contents_result["is_empty"],
            current_path=subpath,
            error=result["message"]
        ), 400

    return redirect(
        url_for(
            "repository.show_docx_in_repo",
            subpath=subpath
        )
    )


# ============================================================
# Create document
# ============================================================

@repository_bp.route(
    "/<path:subpath>/create-docx",
    methods=["POST"]
)
@login_required
def create_docx_route(subpath):

    new_name = request.form.get(
        "new_filename",
        ""
    )

    email = session["email"]

    result = create_document(
        subpath,
        new_name,
        email
    )

    if result["not_found"]:

        return render_template(
            "not_found.html",
            subpath=subpath
        ), 404

    if not result["success"]:

        contents_result = (
            get_repository_contents(subpath)
        )

        return render_template(
            "repo.html",
            folder_list=contents_result["contents"],
            isEmpty=contents_result["is_empty"],
            current_path=subpath,
            error=result["message"]
        ), 400

    return redirect(
        url_for(
            "repository.show_docx_in_repo",
            subpath=subpath
        )
    )