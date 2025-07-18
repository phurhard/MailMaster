import sqlite3
from oauth2client.client import Credentials

def initialize_db():
    """Initialize the SQLite database and create the credentials table if it doesn't exist."""
    conn = sqlite3.connect('credentials.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS user_credentials (
            user_id TEXT PRIMARY KEY,
            credentials TEXT NOT NULL
        )
    ''')
    conn.commit()
    conn.close()

def store_credentials(user_id, credentials):
    """Store OAuth 2.0 credentials in the SQLite database."""
    conn = sqlite3.connect('credentials.db')
    cursor = conn.cursor()
    cursor.execute('''
        INSERT OR REPLACE INTO user_credentials (user_id, credentials)
        VALUES (?, ?)
    ''', (user_id, credentials.to_json()))
    conn.commit()
    conn.close()

def get_stored_credentials(user_id):
    """Retrieve stored credentials for the provided user ID."""
    conn = sqlite3.connect('credentials.db')
    cursor = conn.cursor()
    cursor.execute('SELECT credentials FROM user_credentials WHERE user_id = ?', (user_id,))
    row = cursor.fetchone()
    conn.close()
    if row:
        return Credentials.new_from_json(row[0])
    return None