import csv
import SendKeys
import win32api
import win32com.client
import win32con


import time
statement = 'stmt2.txt'

def send(text):
    win32com.client.Dispatch("WScript.Shell").SendKeys(text)

def click(x,y):
    win32api.SetCursorPos((x,y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)

def dataentry(d,v,a):
    #first thing to add is to select QB screen with a mouseclick
    #add check if you're in the deposits/credits screen
    
    #starts off with "deposit to" highlighted
    send("{TAB}")
    
    #now location: Date.textbox
    send(d)
    send("{TAB}")
    send("{TAB}")
    #time.sleep(1)
    

    #type in Received from: vendor
    send(v)
    send("{TAB}")
    #time.sleep(1)

    #type in From Account: always income here
    send("income")
    send("{TAB}")
    send("{TAB}")
    send("{TAB}")
    send("{TAB}")
    #time.sleep(1)
    

    #type in Amount: amount
    send(a)
    time.sleep(1)
    send("{ENTER}")
    #send("{ENTER}")
    #time.sleep(1)
    
dates = []
vendors = []
amounts = []

def DVA(statement):
    with open(statement) as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',')
        
        for transaction in readCSV:
            date = transaction[0]
            vendor = transaction[1]
            amount = transaction[2]
            
            dates.append(date)
            vendors.append(vendor)
            amounts.append(amount)
            
            dataentry(dates,vendors,amounts)
            
            


click(100,170)
time.sleep(2)


DVA(statement)    

    
    
