from flask import Blueprint, render_template, request, redirect, url_for, session, flash
from functools import wraps
from db.FilterDataAccess import getDocxByFilter

search_bp = Blueprint('search_bp', __name__, url_prefix='/search')

def login_required(f):
    """Decorator to restrict access strictly to admin users."""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'email' not in session:
            return redirect('/')
        return f(*args, **kwargs)
    return decorated_function

@search_bp.route('', methods=['GET'])
@login_required
def search():
    search = request.args.get('filter_text')
    print(search)
    docx_list = getDocxByFilter(search)
    return render_template('search.html', docx_list=docx_list, search_text=search)



