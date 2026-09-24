from flask import Blueprint, render_template, request, redirect, url_for, session, flash
from functools import wraps
from models import db, User

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
    users = User.query.all()
    return render_template('manage.html', users=users)

@user_bp.route('/create', methods=['POST'])
@admin_required
def create_user():
    email = request.form.get('email')
    role = request.form.get('role', 'user')
    
    if User.query.filter_by(email=email).first():
        flash("User with this email already exists.", "danger")
    else:
        new_user = User(email=email, role=role)
        db.session.add(new_user)
        db.session.commit()
        flash("User created successfully.", "success")
        
    return redirect(url_for('user_bp.manage_users'))

@user_bp.route('/edit/<int:user_id>', methods=['POST'])
@admin_required
def edit_user(user_id):
    user = User.query.get_or_404(user_id)
    user.email = request.form.get('email')
    user.role = request.form.get('role')
    db.session.commit()
    flash("User updated successfully.", "success")
    return redirect(url_for('user_bp.manage_users'))

@user_bp.route('/delete/<int:user_id>', methods=['POST'])
@admin_required
def delete_user(user_id):
    user = User.query.get_or_404(user_id)
    db.session.delete(user)
    db.session.commit()
    flash("User deleted successfully.", "success")
    return redirect(url_for('user_bp.manage_users'))