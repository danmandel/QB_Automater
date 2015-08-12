import csv
import win32com.client
import win32api
import win32con
import SendKeys

statement = 'stmt1.txt'

def send(text):
   win32com.client.Dispatch("WScript.Shell").SendKeys(text)

def click(x,y):
    win32api.SetCursorPos((x,y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)

def dataentry(date,vendor,amount):
    #in theory you would first check if you're in the deposits/credits screen
    #but rn i'm going to assume im at deposits.
    #starts off with "deposit to" highlighted
    #abs first thing to do is to select QB screen with a mouseclick

    #type in date according to statement row
    send("{TAB}")
    send(date)
    send("{TAB}")
    send("{TAB}")

    #type in Received from: vendor
    send(vendor)
    send("{TAB}")

    #type in From Account: always income here
    send("income")
    send("{TAB}")
    send("{TAB}")
    send("{TAB}")
    send("{TAB}")

    #type in Amount: amount
    send(amount)
    send("{ENTER}")
    send("{ENTER}")

def DVA(statement):
    with open(statement) as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',')
        
        for transaction in readCSV:
            date = transaction[0]
            vendor = transaction[1]
            amount = transaction[2]
            
            dataentry(date,vendor,amount)

           
        
click(100,120)



DVA(statement)    
