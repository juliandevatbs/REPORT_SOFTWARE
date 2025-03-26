from typing import Tuple, Union
import sys
import os
# Agrega el directorio raíz al PATH
sys.path.append(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))


from DB.connect import connect_db


def get_header_data() -> Union[Tuple[str, str, str, str, str, str], bool]:
    """
    Retrieves client header data from the database for Client_ID '1063'.
    
    Returns:
        Tuple containing (client, address, city, company_phone, zip_code, state) if successful,
        False if the connection fails or no data is found.
        
    Raises:
        Prints error messages but doesn't raise exceptions to caller
    """
    connection, cursor = None, None
    try:
        connection, cursor = connect_db()
        
        if not connection or not cursor:
            print("Connection module failed")
            return False
        
        query = """
        SELECT Client, Address_1, City, CompanyPhone, Postal_Code, State_Prov 
        FROM CLIENTS 
        WHERE Client_ID = '1063';
        """
        
        cursor.execute(query)
        results = cursor.fetchall()
        
        if not results:
            print("No data found for Client_ID 1063")
            return False
            
        if len(results[0]) != 6:
            print("Unexpected number of columns returned")
            return False
            
        client, address, city, company_phone, zip_code, state = results[0]
        
        return client, address, city, company_phone, zip_code, state
        
    except Exception as e:
        print(f"Database error occurred: {str(e)}")
        return False
        
    finally:
        try:
            if cursor:
                cursor.close()
            if connection:
                connection.close()
        except Exception as e:
            print(f"Error closing database resources: {str(e)}")


#get_header_data()