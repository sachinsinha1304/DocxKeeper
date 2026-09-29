from .db import get_db_connection
from pathlib import Path
from constants import DOCS_ROOT

def get_relative_subpath(file_path_str):
    base_path = Path(DOCS_ROOT)
    file_path = Path(file_path_str)
    
    # Get the relative path and convert it to POSIX format (forward slashes)
    return file_path.relative_to(base_path).as_posix()


def getDocxByFilter(text):
    con = get_db_connection()
        # Use dictionary=True if using mysql-connector so you can access fields by name
    cursor = con.cursor(dictionary=True) 
    try:
            # 1. Query ONLY by email using a safe placeholder (%s)
        query = f"SELECT distinct file_path from documents_db.document_versions WHERE file_path LIKE '%{text}%'"
        cursor.execute(query)
        selected_docx = cursor.fetchall()
        subpath = [get_relative_subpath(path['file_path']) for path in selected_docx]
        return subpath
    except Exception as ex:
        print(ex)