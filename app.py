from flask import Flask, render_template, request, session, redirect, url_for, flash
from extension import socketio
from utilities.docxDisplayHelper import listAllTeamFolder, listAllSubfolderInRepo, readDocxContent
from utilities.docx_editor import get_docx_as_html, save_html_as_docx
from utilities.file_safety import acquire_editing_lock, release_editing_lock, refresh_editing_lock
from utilities.docx_create import (
    resolve_folder_path, create_new_repo, create_new_folder, create_new_docx
)
from pathlib import Path
from utilities.document_versioning import list_versions, restore_version
DOCS_ROOT = Path(__file__).resolve().parent / "documents"


app = Flask(__name__)
socketio.init_app(app)
app.secret_key = 'your_super_secret_and_random_string'


def resolve_file(subpath):
    """
    Resolves a subpath to (parent, name), or returns (None, None) if it's
    not a valid file (missing, or a folder rather than a document).
    """
    try:
        contents, isFile = listAllSubfolderInRepo(subpath)
    except FileNotFoundError:
        return None, None
    if not isFile:
        return None, None
    return contents["parent"], contents["name"]


@app.route("/", methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        session['email'] = request.form.get("email")
        return redirect("all-docx")
    return render_template("login.html")


@app.route("/all-docx")
def show_docx():
    folder_list, is_empty = listAllTeamFolder()
    return render_template("all_docx.html", folder_list=folder_list, isEmpty=is_empty)


@app.route("/all-docx/", defaults={"subpath": ""})
@app.route("/all-docx/<path:subpath>")
def show_docx_in_repo(subpath):
    if 'email' not in session:
        return redirect("/")
    try:
        contents, isFile = listAllSubfolderInRepo(subpath)
    except FileNotFoundError:
        return render_template("not_found.html", subpath=subpath), 404

    if isFile:
        parent_dir = contents['parent']
        file_name = contents['name']
        
        # FIX: Directly build the absolute path using DOCS_ROOT and subpath
        target_path = DOCS_ROOT / subpath
        
        # Fallback check if the file exists
        if not target_path.exists():
            # Try alternative resolution if your structure differs
            target_path = Path(parent_dir) / file_name

        file_content = readDocxContent(parent_dir, file_name)
        
        # Now list_versions will find the file correctly
        versions = list_versions(target_path)

        return render_template(
            "view_file.html",
            filename=file_name,
            parent=parent_dir,
            subpath=subpath,
            content=file_content,
            versions=versions
        )

    is_empty = len(contents) == 0
    return render_template(
        "repo.html",
        folder_list=contents,
        isEmpty=is_empty,
        current_path=subpath,
    )


@app.route("/home")
def home():
    if 'email' in session:
        return render_template("editor.html")
    else:
        return redirect("/")


@app.route("/edit-docx/<path:subpath>", methods=["GET"])
def edit_docx_view(subpath):
    if 'email'not in session:
            return redirect("/")
    parent, name = resolve_file(subpath)
    if not parent:
        return render_template("not_found.html", subpath=subpath), 404

    email = session["email"]
    target_path = Path(parent) / name

    acquired, holder = acquire_editing_lock(target_path, email)
    if not acquired:
        return render_template("edit_locked.html", filename=name, subpath=subpath, holder=holder)

    html_content, version = get_docx_as_html(parent, name)
    return render_template(
        "edit_docx.html",
        filename=name,
        subpath=subpath,
        doc_html=html_content,
        version=version,
    )


@app.route("/edit-docx/<path:subpath>", methods=["POST"])
def edit_docx_save(subpath):
    if 'email'not in session:
        return redirect("/")
    parent, name = resolve_file(subpath)
    if not parent:
        return render_template("not_found.html", subpath=subpath), 404

    html_content = request.form.get("doc_html", "")
    expected_version = request.form.get("version")
    email = session["email"]

    result = save_html_as_docx(parent, name, html_content, expected_version)

    if not result["success"]:
        return render_template(
            "edit_conflict.html", filename=name, subpath=subpath, doc_html=html_content
        ), 409

    release_editing_lock(Path(parent) / name, email)
    return redirect(url_for("show_docx_in_repo", subpath=subpath))


@app.route("/edit-docx/<path:subpath>/heartbeat", methods=["POST"])
def edit_docx_heartbeat(subpath):
    parent, name = resolve_file(subpath)
    if not parent:
        return "", 404
    email = session["email"]
    refresh_editing_lock(Path(parent) / name, email)
    return "", 204


@app.route("/edit-docx/<path:subpath>/cancel", methods=["POST"])
def edit_docx_cancel(subpath):
    if 'email' not in session:
        return redirect("/")
    parent, name = resolve_file(subpath)
    if not parent:
        return "", 404
    email = session["email"]
    release_editing_lock(Path(parent) / name, email)
    return "", 204

@app.route("/all-docx/create-repo", methods=["POST"])
def create_repo_route():
    """Root-level only: creates a brand new top-level repo folder."""
    new_name = request.form.get("new_repo_name", "")
    success, message, final_name = create_new_repo(new_name)

    if not success:
        folder_list, is_empty = listAllTeamFolder()
        return render_template("all_docx.html", folder_list=folder_list, isEmpty=is_empty, error=message), 400

    return redirect(url_for("show_docx"))


@app.route("/all-docx/<path:subpath>/create-folder", methods=["POST"])
def create_folder_route(subpath):
    """Inside a repo/folder: creates a new subfolder."""
    new_name = request.form.get("new_folder_name", "")

    try:
        directory_path = resolve_folder_path(subpath)
    except FileNotFoundError:
        return render_template("not_found.html", subpath=subpath), 404

    success, message, final_name = create_new_folder(directory_path, new_name)

    if not success:
        contents, _ = listAllSubfolderInRepo(subpath)
        return render_template(
            "repo.html", folder_list=contents, isEmpty=len(contents) == 0,
            current_path=subpath, error=message
        ), 400

    return redirect(url_for("show_docx_in_repo", subpath=subpath))


@app.route("/all-docx/<path:subpath>/create-docx", methods=["POST"])
def create_docx_route(subpath):
    """Inside a repo/folder: creates a new blank .docx."""
    new_name = request.form.get("new_filename", "")

    try:
        directory_path = resolve_folder_path(subpath)
    except FileNotFoundError:
        return render_template("not_found.html", subpath=subpath), 404

    success, message, final_name = create_new_docx(directory_path, new_name)

    if not success:
        contents, _ = listAllSubfolderInRepo(subpath)
        return render_template(
            "repo.html", folder_list=contents, isEmpty=len(contents) == 0,
            current_path=subpath, error=message
        ), 400

    return redirect(url_for("show_docx_in_repo", subpath=subpath))

# @app.route('/all-docx/<path:subpath>')
# def view_docx(subpath):
#     target_path = resolve_folder_path(subpath) # your path resolver
#     filename = target_path.name
    
#     # Fetch content for reading...
#     content = readDocxContent(target_path.parent, filename)
    
#     # Fetch version list for the history modal
#     versions = list_versions(target_path)

#     return render_template('view_docx.html', filename=filename, subpath=subpath, content=content, versions=versions)


@app.route('/restore-docx/<version_id>', methods=['POST'])
def restore_docx_route(version_id):
    if 'email' not in session:
        return redirect("/")
        
    # Get subpath from the hidden form field
    subpath = request.form.get('subpath')
    if not subpath:
        flash("Invalid document path.", "danger")
        return redirect(url_for('show_docx_in_repo', subpath=''))
        
    target_path = DOCS_ROOT / subpath
    current_user = session.get('email', 'Admin')
    
    try:
        # Restores the old snapshot and logs it automatically
        restore_version(target_path, version_id, restored_by=current_user)
        flash(f"Successfully restored version!", "success")
    except Exception as e:
        flash(f"Error restoring version: {e}", "danger")
        
    return redirect(url_for('show_docx_in_repo', subpath=subpath))



if __name__ == "__main__":
    app.run(debug=True)