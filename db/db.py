# db.py
import mysql.connector
from mysql.connector import pooling
from werkzeug.security import generate_password_hash

DB_CONFIG = {
    "host": "172.20.42.143",
    "user": "sa",
    "password": "sa"
}

db_pool = None

def init_db():
    """Runs on system startup to create the database/tables and initialize the pool."""
    temp_conn = mysql.connector.connect(**DB_CONFIG)
    temp_cursor = temp_conn.cursor()
    try:
        temp_cursor.execute("CREATE DATABASE IF NOT EXISTS documents_db;")
        temp_cursor.execute("USE documents_db;")
        
        # Create users table with password column
        temp_cursor.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INT AUTO_INCREMENT PRIMARY KEY,
                email VARCHAR(120) NOT NULL UNIQUE,
                password VARCHAR(255) NOT NULL,
                role VARCHAR(20) DEFAULT 'user'
            );
        """)
        
        # Seed default admin if table is empty
        temp_cursor.execute("SELECT COUNT(*) FROM users;")
        result = temp_cursor.fetchone()
        if result[0] == 0:
            hashed_password = generate_password_hash("admin123")
            temp_cursor.execute(
                "INSERT INTO users (email, password, role) VALUES (%s, %s, %s);",
                ("admin@example.com", hashed_password, "admin")
            )
            print("Default admin user created: admin@example.com (Password: admin123)")
            
        temp_conn.commit()
    finally:
        temp_cursor.close()
        temp_conn.close()

    # Initialize connection pool
    global db_pool
    pool_config = {**DB_CONFIG, "database": "documents_db"}
    db_pool = pooling.MySQLConnectionPool(
        pool_name="app_connection_pool",
        pool_size=5,
        pool_reset_session=True,
        **pool_config
    )
    print("MySQL connection pool initialized successfully!")

def get_db_connection():
    """Fetches an active connection from the pool."""
    if db_pool is None:
        raise RuntimeError("Database pool is not initialized. Run init_db() first.")
    return db_pool.get_connection()