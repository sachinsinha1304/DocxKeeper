# db.py
import mysql.connector
from mysql.connector import pooling
from werkzeug.security import check_password_hash, generate_password_hash
from .db import get_db_connection

def getUserByEmail(email, password):
    """
    Fetches the user by email using a parameterized query, 
    then securely verifies the password using Werkzeug.
    """
    con = get_db_connection()
    # Use dictionary=True if using mysql-connector so you can access fields by name
    cursor = con.cursor(dictionary=True) 
    try:
        # 1. Query ONLY by email using a safe placeholder (%s)
        query = "SELECT email, password, role FROM documents_db.users WHERE email = %s"
        cursor.execute(query, (email,))
        user = cursor.fetchone()
        print(user)
        
        # 2. Verify the typed password against the stored database hash in Python
        if user and check_password_hash(user['password'], password):
            # Return only what is needed (exclude the password hash for safety)
            return {"email": user["email"], "role": user["role"]}
            
        return None
    finally:
        cursor.close()
        con.close()

def getAllUsers():
    con = get_db_connection()
    # Use dictionary=True if using mysql-connector so you can access fields by name
    cursor = con.cursor(dictionary=True) 
    try:
        # 1. Query ONLY by email using a safe placeholder (%s)
        query = "SELECT id, email, password, role FROM documents_db.users"
        cursor.execute(query)
        users = cursor.fetchall()

        return users
    finally:
        cursor.close()
        con.close()

def createUser(user, password, role):
    con = get_db_connection()
    cursor = con.cursor(dictionary=True) 
    try:
        query = "insert into documents_db.users(email, password, role) values(%s, %s, %s)"
        cursor.execute(query, (user, generate_password_hash(password), role))
        con.commit()
        return 0
    except Exception as ex:
        print(ex)

    finally:
        cursor.close()
        con.close()


def editUser(user, password, role, user_id):
    con = get_db_connection()
    cursor = con.cursor(dictionary=True) 
    try:
        query = 'update documents_db.users set email = %s, password = %s,  role = %s where id = %s'
        cursor.execute(query, (user, generate_password_hash(password), role, user_id))
        return 0
    except Exception as ex:
        print(ex)

    finally:
        cursor.close()
        con.close()


def deleteUser(user_id):
    con = get_db_connection()
    cursor = con.cursor(dictionary=True) 
    try:
        query = 'delete from documents_db.users where id = %s'
        cursor.execute(query, (user_id,))
        con.commit()
        return 0
    except Exception as ex:
        print(ex)

    finally:
        cursor.close()
        con.close()