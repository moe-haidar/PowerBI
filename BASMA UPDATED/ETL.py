import pandas as pd
import pyodbc

# Step 1: Extract data from Excel file
excel_file = "campain_data.xlsx"  # Replace with the path to your Excel file
df = pd.read_excel(excel_file)

 
# Step 2: Load data into SQL Server
 
server = 'LAPTOP-L5ANLNV9'
database = 'moe'

 
# Establish connection to SQL Server using Windows Authentication 
conn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';Trusted_Connection=yes')

# Create a cursor object using the connection
cursor = conn.cursor()

# Define your SQL query to create a table and load data into it
create_table_query = '''
CREATE TABLE CampaignData (
    [Campaign name] VARCHAR(255),
    [Day] DATE,
    Platform VARCHAR(50),
    Reach INT,
    Impressions INT,
    [Amount Spent (USD)] DECIMAL(10, 2),
    [Link clicks] INT,
    [Unique Registrations] INT,
    [Unique Purchases] INT
)
'''

# Execute the create table query
cursor.execute(create_table_query)

# Convert DataFrame to list of tuples for insertion into SQL Server
data = [tuple(row) for row in df.values]

# Define the insert query
insert_query = '''
INSERT INTO CampaignData 
([Campaign name], [Day], Platform, Reach, Impressions, [Amount Spent (USD)], [Link clicks], [Unique Registrations], [Unique Purchases])
VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
'''


# Execute the insert query with the data
cursor.executemany(insert_query, data)

# Commit the transaction
conn.commit()

# Close the cursor and connection
cursor.close()
conn.close()

print("Data successfully loaded into SQL Server.")