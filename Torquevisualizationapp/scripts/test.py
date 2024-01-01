import sqlite3
import pandas as pd

# Connect to the SQLite database
conn = conn = sqlite3.connect('C:\\Program Files\\sqlite-tools-win32-x86-3420000\\test_db')

# Read data into a pandas DataFrame
df = pd.read_sql_query("SELECT * FROM ad_leads_data", conn)

# Display the data
print(df)

# Close the connection
conn.close()
