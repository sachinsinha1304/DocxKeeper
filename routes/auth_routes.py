from flask import (
    Blueprint,
    render_template,
    request,
    session,
    redirect,
    url_for,
    flash
)

from db.UserDataAccess import getUserByEmail


auth_bp = Blueprint(
    "auth",
    __name__
)


@auth_bp.route("/", methods=["GET", "POST"])
def login():
    # Already logged in
    if "email" in session:
        return redirect(
            url_for("repository.show_docx")
        )
    if request.method == "POST":
        email = request.form.get("email")
        password = request.form.get("password")
        user = getUserByEmail(
            email,
            password
        )
        if user:
            session["email"] = user["email"]
            session["role"] = user["role"]
            return redirect(
                url_for("repository.show_docx")
            )
        flash(
            "User not found in database.",
            "danger"
        )
    return render_template("login.html")


@auth_bp.route("/logout")
def logout():
    session.clear()
    return redirect(
        url_for("auth.login")
    )