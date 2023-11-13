import sqlite3

connection = sqlite3.connect("my_database.db")

cursor = connection.cursor()

cursor.execute(
    """
    CREATE TABLE IF NOT EXISTS email_types (
        type TEXT PRIMARY KEY
    )
"""
)

cursor.execute(
    """
    CREATE TABLE IF NOT EXISTS email_types (
        type TEXT PRIMARY KEY
    )
"""
)
