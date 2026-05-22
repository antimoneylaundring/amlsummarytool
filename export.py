import pandas as pd
import psycopg2
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

try:
    # Get database URL from .env file
    DB_URL = os.getenv("DB_URL")

    if DB_URL and 'postgresql+psycopg2://' in DB_URL:
        DB_URL = DB_URL.replace('postgresql+psycopg2://', 'postgresql://')
    
    # Connect to database
    conn = psycopg2.connect(DB_URL)
    
    # Fetch data from table
    df = pd.read_sql_query('SELECT * FROM "all_upiiD"', conn)
    
    # Close connection
    conn.close()
    
    # Export to Excel
    df.to_excel("upi_ids_export.xlsx", index=False, sheet_name='UPI IDs')
    
    print(f"Successfully exported {len(df):,} records to upi_ids_export.xlsx")

except Exception as e:
    print(f"Error: {e}")