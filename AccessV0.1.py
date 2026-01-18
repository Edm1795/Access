# Log of Actions

# Downloaded Microsoft Access Engine 64 bit version (https://www.microsoft.com/en-us/download/details.aspx?id=54920)
# (check that you get hte version matching the bitness of your Python Inter. x64 for 64 bit; .exe. for 32 bit)


import pyodbc

# Open the Access File
def openFile(filename):
    '''
    This function opens a Microsoft Access file and returns that opened file
    Input: str: the filename without any extension
    Output: the opened file
    '''

    try:
        conn = pyodbc.connect(
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            r"DBQ=D:\DataForProgramming" + "\\" + filename + r".accdb;")  # Access File Name ("\\" is used to give a single slash in the path)
        return conn
    except:
        print('There is an error in opening the file')



openedFile=openFile('accounts')

cursor = openedFile.cursor()
print("Connected successfully")

# Print off Contents of Access File
for row in cursor.tables():
    print(row.table_name)
