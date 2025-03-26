import pyodbc
from dotenv import load_dotenv
import os

load_dotenv()

'''
    This function execute the database connection
    Returns the connection and the cursor
'''

def connect_db():
     
    try:
         
        # Get the protected database variables
        db_server= os.getenv("DB_SERVER")
        db_name= os.getenv("DB_NAME")
        db_user= os.getenv("DB_USER")
        db_password= os.getenv("DB_PASSWORD")
        
        # Connection string to SQL SERVER
        connection_string=f"""
            DRIVER=SQL Server;
            SERVER={db_server};
            DATABASE=SRLSQL;
            UID=juliandevuser;
            PWD=devUserdb@1;
        """
        #print(connection_string)
        # Connection
        connection= pyodbc.connect(connection_string)
        cursor= connection.cursor()
        
        print("Succesfully database connection")
        return connection, cursor
    
    except Exception as e:
        
        print(f"Error al conectar a SQL Server: {e}")
        return None, None
    
    
connect_db()