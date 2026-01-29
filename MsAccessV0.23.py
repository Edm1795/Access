# Log of Actions

# Version 0.23

## remember to run the createAccessAcountsFile to create the test Access file which this program uses. Make sure you save the Test Access file in the correct directory
## currently see for D:\DataForProgramming. Adjust the directory as you like inside the first function below (openFile)

# Downloaded Microsoft Access Engine 64 bit version (https://www.microsoft.com/en-us/download/details.aspx?id=54920)
# (check that you get hte version matching the bitness of your Python Inter. x64 for 64 bit; .exe. for 32 bit)
# Changed code into set of functions and added main() (not yet fully generalized)

# Version 0.22

# Now the openFile() needs to be given a filename with its extension .mdb or .accdb and it will open the file

# Ver. 0.23
# this simply opens a different account file which now has 20 accounts not just 10 for the test file. (accounts20.accdb)

import pyodbc
import matplotlib.pyplot as plt

# Open the Access File
def openFile(filename):
    '''
    This function opens a Microsoft Access file and returns that opened file
    Input: str: the filename with extension .mdb or .accdb
    Output: the opened file
    '''


    try:
        conn = pyodbc.connect(
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            r"DBQ=D:\DataForProgramming" + "\\" + filename)  # Access File Name ("\\" is used to give a single slash in the path)
        return conn
    except:
        print('There is an error in opening the file')


def initiateCursor(file):

    openedFile=file
    print("Connected successfully")

    return openedFile.cursor()



def closeFile(openedFile):
    '''
    This function closes an Access Database
    Input: openedFile: the file you want to close
    '''
    openedFile.close()

def fetchData(cursor):
    '''
    This function gets the raw data from the database, currently specified for the account number
    and the balances.
    Input: the cursor
    '''

    cursor.execute("SELECT AccountNumber, Balance FROM Accounts20 ORDER BY AccountNumber")

    return cursor.fetchall()  # returns all data fetched

def extractSpecifiedColumns(data):
    '''
    This function returns a tuple of lists ([list1],[list2]) of the data given to it from the
    fetchData(). Presumably the input data is some database type format not yet well organizaed
    input: rawdata from fetchData()
    return: a tuple of lists
    '''

    firstCol = [row.AccountNumber for row in data]
    secColumn = [row.Balance for row in data]

    return (firstCol,secColumn)

def plotData(data):

    plt.figure(figsize=(8, 5))
    plt.bar(data[0], data[1], color='skyblue')
    plt.xlabel("Account Number")
    plt.ylabel("Balance ($)")
    plt.title("Account Balances")
    plt.xticks(data[0])  # Show each account number on x-axis
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    plt.show()

def main():

    openedFile = openFile('accounts20.accdb')
    cursor=initiateCursor(openedFile)
    rawData=fetchData(cursor)
    closeFile(openedFile)
    extractedData=extractSpecifiedColumns(rawData)
    plotData(extractedData)

# Log of Actions

# Version 0.23

## remember to run the createAccessAcountsFile to create the test Access file which this program uses. Make sure you save the Test Access file in the correct directory
## currently see for D:\DataForProgramming. Adjust the directory as you like inside the first function below (openFile)

# Downloaded Microsoft Access Engine 64 bit version (https://www.microsoft.com/en-us/download/details.aspx?id=54920)
# (check that you get hte version matching the bitness of your Python Inter. x64 for 64 bit; .exe. for 32 bit)
# Changed code into set of functions and added main() (not yet fully generalized)

# Version 0.22

# Now the openFile() needs to be given a filename with its extension .mdb or .accdb and it will open the file

# Ver. 0.23
# this simply opens a different account file which now has 20 accounts not just 10 for the test file. (accounts20.accdb)

import pyodbc
import matplotlib.pyplot as plt

# Open the Access File
def openFile(filename):
    '''
    This function opens a Microsoft Access file and returns that opened file
    Input: str: the filename with extension .mdb or .accdb
    Output: the opened file
    '''


    try:
        conn = pyodbc.connect(
            r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
            r"DBQ=D:\DataForProgramming" + "\\" + filename)  # Access File Name ("\\" is used to give a single slash in the path)
        return conn
    except:
        print('There is an error in opening the file')


def initiateCursor(file):

    openedFile=file
    print("Connected successfully")

    return openedFile.cursor()



def closeFile(openedFile):
    '''
    This function closes an Access Database
    Input: openedFile: the file you want to close
    '''
    openedFile.close()

def fetchData(cursor):
    '''
    This function gets the raw data from the database, currently specified for the account number
    and the balances.
    Input: the cursor
    '''

    cursor.execute("SELECT AccountNumber, Balance FROM Accounts20 ORDER BY AccountNumber")

    return cursor.fetchall()  # returns all data fetched

def extractSpecifiedColumns(data):
    '''
    This function returns a tuple of lists ([list1],[list2]) of the data given to it from the
    fetchData(). Presumably the input data is some database type format not yet well organizaed
    input: rawdata from fetchData()
    return: a tuple of lists
    '''

    firstCol = [row.AccountNumber for row in data]
    secColumn = [row.Balance for row in data]

    return (firstCol,secColumn)

def plotData(data):

    plt.figure(figsize=(8, 5))
    plt.bar(data[0], data[1], color='skyblue')
    plt.xlabel("Account Number")
    plt.ylabel("Balance ($)")
    plt.title("Account Balances")
    plt.xticks(data[0])  # Show each account number on x-axis
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    plt.show()

def main():

    openedFile = openFile('accounts20.accdb')
    cursor=initiateCursor(openedFile)
    rawData=fetchData(cursor)
    closeFile(openedFile)
    extractedData=extractSpecifiedColumns(rawData)
    plotData(extractedData)

