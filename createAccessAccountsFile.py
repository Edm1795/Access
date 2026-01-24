import win32com.client # known as pywin32 on package installer
import pyodbc
import os
import random

#  Create the database file
db_path = r"D:\DataForProgramming\accounts.accdb"
os.makedirs(os.path.dirname(db_path), exist_ok=True)

cat = win32com.client.Dispatch("ADOX.Catalog")
cat.Create(f"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={db_path};")
print("Database created:", db_path)

# Connect with pyodbc
conn = pyodbc.connect(
    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
    + r"DBQ=" + db_path + r";"
)
cursor = conn.cursor()

# 3️⃣ Create table
cursor.execute("""
CREATE TABLE Accounts (
    AccountNumber AUTOINCREMENT PRIMARY KEY,
    Balance DOUBLE
)
""")
conn.commit()
print("Table created: Accounts")

# 4️⃣ Insert 10 accounts with random balances
for _ in range(10):
    balance = round(random.uniform(100, 10000), 2)  # random between 100 and 10,000
    cursor.execute("INSERT INTO Accounts (Balance) VALUES (?)", (balance,))
conn.commit()
print("Inserted 10 accounts")

# 5️⃣ Verify contents
for row in cursor.execute("SELECT * FROM Accounts"):
    print(row)

conn.close()
