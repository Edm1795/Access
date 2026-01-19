# Log of Actions

# Version 0.2

# Downloaded Microsoft Access Engine 64 bit version (https://www.microsoft.com/en-us/download/details.aspx?id=54920)
# (check that you get hte version matching the bitness of your Python Inter. x64 for 64 bit; .exe. for 32 bit)
# Changed code into set of functions and added main() (not yet fully generalized)

import pyodbc
import matplotlib.pyplot as plt

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


def initiateCursor(file):

    openedFile=file
    print("Connected successfully")

    return openedFile.cursor()



def closeFile(openedFile):
    openedFile.close()

def fetchData(cursor):


    cursor.execute("SELECT AccountNumber, Balance FROM Accounts ORDER BY AccountNumber")

    return cursor.fetchall()  # returns all data fetched

def extractSpecifiedColumns(data):

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

    openedFile = openFile('accounts')
    cursor=initiateCursor(openedFile)
    rawData=fetchData(cursor)
    closeFile(openedFile)
    extractedData=extractSpecifiedColumns(rawData)
    plotData(extractedData)

main()
