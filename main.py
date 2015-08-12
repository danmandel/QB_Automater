import csv
import win32com.client
import win32api
import win32con
import SendKeys


statement = 'stmt1.txt'

shell = win32com.client.Dispatch("WScript.Shell")
shell.SendKeys("test")

def DVA(statement):
    with open(statement) as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',')
        
        for transaction in readCSV:
            date = transaction[0]
            vendor = transaction[1]
            amount = transaction[2]

            return (date, vendor, amount)
        
##    print ("dates: ", dates)
##    print "             "
##    print ("vendors: ", vendors)
##    print "             "
##    print ("amounts: ", amounts)


def click(x,y):
    win32api.SetCursorPos((x,y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)


click(100,120)




def dataentry(date,vendor,amount):
    #in theory you would first check if you're in the deposits/credits screen
    #but rn i'm going to assume im at deposits.
    #starts off with "deposit to" highlighted
    #abs first thing to do is to select QB screen with a mouseclick
    shell = win32com.client.Dispatch("WScript.Shell")
    tab = shell.SendKeys("{TAB}")

##    #type in date according to statement row
##    type something
##    tab
##    tab
##
##    #type in Received from: vendor
##    type something
##    tab
##
##    #type in From Account: always income here
##    type "income"
##    tab
##    tab
##    tab
##    tab
##
##    #type in Amount: amount
##    type "amount"
##    enter
##    enter
##    
##    
##    
##        
##for transaction in readCSV:
##    dataentry(date,vendor,amount)
##    

#
