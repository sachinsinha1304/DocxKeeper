from flask import Blueprint, render_template, request, redirect, url_for, session, flash
from functools import wraps
from models import db, User

from db.UserDataAccess import getAllUsers, createUser, editUser, deleteUser

user_bp = Blueprint('user_bp', __name__, url_prefix='/users')

def admin_required(f):
    """Decorator to restrict access strictly to admin users."""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'email' not in session or session.get('role') != 'admin':
            flash("Access denied. Administrator privileges required.", "danger")
            return redirect(url_for('show_docx'))
        return f(*args, **kwargs)
    return decorated_function

@user_bp.route('/manage', methods=['GET'])
@admin_required
def manage_users():
    users = getAllUsers()
    return render_template('manage.html', users=users)

@user_bp.route('/create', methods=['POST'])
@admin_required
def create_user():
    email = request.form.get('email')
    role = request.form.get('role', 'user')
    password = request.form.get('password')
    if not createUser(email, password, role):
        flash("User created successfully.", "success") 
        
    return redirect(url_for('user_bp.manage_users'))

@user_bp.route('/edit/<int:user_id>', methods=['POST'])
@admin_required
def edit_user(user_id):
    print(user_id)
    email = request.form.get('email')
    role = request.form.get('role')
    password = request.form.get('password')
    if not editUser(email, password, role, user_id):
        flash("User updated successfully.", "success")
    return redirect(url_for('user_bp.manage_users'))

@user_bp.route('/delete/<int:user_id>', methods=['POST'])
@admin_required
def delete_user(user_id):
    if not deleteUser(user_id):
        flash("User deleted successfully.", "success")
    return redirect(url_for('user_bp.manage_users'))