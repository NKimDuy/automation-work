import mysql.connector
from config.settings import host, user_db, password_db, database

def get_connection():
      conn = mysql.connector.connect(
            host=host,
            user=user_db,
            password=password_db,
            database=database
      )
      return conn

