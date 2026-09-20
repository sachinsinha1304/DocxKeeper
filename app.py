from flask import Flask, render_template, request, session, redirect, url_for
from extension import socketio
from utilities.docxDisplayHelper import listAllTeamFolder, listAllSubfolderInRepo, readDocxContent
from flask import request, redirect, render_template
from utilities.docx_editor import get_docx_as_html, save_html_as_docx



app = Flask(__name__)
socketio.init_app(app)
app.secret_key = 'your_super_secret_and_random_string'

@app.route("/", methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        session['email'] = request.form.get("email")
        return redirect(url_for("home"))                               
    return render_template("login.html")

@app.route("/all-docx")
def show_docx(): 
    folder_list, _ = listAllTeamFolder()   
    print()                         
    return render_template("all_docx.html", folder_list=folder_list, isEmpty= _)

@app.route("/all-docx/", defaults={"subpath": ""})
@app.route("/all-docx/<path:subpath>")
def show_docx_in_repo(subpath):   
    contents, isFile = listAllSubfolderInRepo(subpath)
    if isFile:
        print(contents)
        file_content = readDocxContent(contents['parent'], contents['name'])
        return render_template("view_file.html", filename=contents['name'],parent=contents['parent'], content=file_content)
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

@app.route("/edit-docx", methods=["GET"])
def edit_docx_view():
    parent = request.args.get("parent")
    name = request.args.get("name")
    if not parent or not name:
        return "Missing file information — please open this page via the document viewer.", 400

    html_content = get_docx_as_html(parent, name)
    return render_template("edit_docx.html", filename=name, parent=parent, doc_html=html_content)


@app.route("/edit-docx", methods=["POST"])
def edit_docx_save():
    parent = request.form.get("parent")
    name = request.form.get("name")
    return_url = request.form.get("return_url")
    html_content = request.form.get("doc_html", "")

    save_html_as_docx(parent, name, html_content)

    return redirect(return_url or f"/edit-docx?parent={parent}&name={name}")
if __name__ == "__main__":
    socketio.run(app)
    app.run(debug="True")
    