import csv
import SendKeys
import win32api
from win32com.client import Dispatch
import win32con
import time

class Checkpoint(object):
    #square or rectangle ABCD.
    name = ""   
    a_coords = (0,0)
    b_coords = (0,0)
    c_coords = (0,0)
    d_coords = (0,0)
    mid_xaxis = 0
    mid_yaxis = 0
    midpoint = (0,0)
    #image = checkpoint.jpg
    
def make_checkpoint(name,a_coords,b_coords,c_coords,d_coords):
    checkpoint = Checkpoint()
    checkpoint.name = name
    checkpoint.a_coords = a_coords
    checkpoint.b_coords = b_coords
    checkpoint.c_coords = c_coords
    checkpoint.d_coords = d_coords

    #didn't have to include these in the args
    checkpoint.mid_xaxis = (a_coords[0]+b_coords[0])/2
    checkpoint.mid_yaxis = (a_coords[1]+d_coords[1])/2
    checkpoint.midpoint = (checkpoint.mid_xaxis, checkpoint.mid_yaxis)
    return checkpoint

def close_all_windows():
    Auto.WinActivate(apptitle)
    Auto.send("!w")
    Auto.send("{DOWN}")
    Auto.send("{ENTER}")

def open_home():
    Auto.MouseClick("left", 34, 77)
    
def setup():
    Auto.WinActivate(apptitle)
    Auto.WinMove(apptitle,"", 0, 0, 1000, 1000)
    close_all_windows()
    open_home()
    Auto.send("{TAB}")
    Auto.send("{ENTER}")
    time.sleep(1)
    Auto.send("!w")
    time.sleep(1)
    Auto.send("{DOWN 3}")
    time.sleep(1)
    Auto.send("{ENTER}") # right now leaves me with date selected in BOA

##def move_to(coordinates):
##    win32api.SetCursorPos(coordinates)
    
def click(coordinates):
    win32api.SetCursorPos((coordinates))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,coordinates[0],coordinates[1],0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,coordinates[0],coordinates[1],0,0)

def depositentry(d,v,a):
    #add check if you're in the deposits/credits screen

    Auto.send(d) #Date
    Auto.send("{TAB 2}")
    Auto.send(v) #Vendor
    Auto.send("{TAB 3}")
    Auto.send(a) #Amount
    Auto.send("{TAB}")
    Auto.send("income") #Income
    Auto.send("{TAB 2}") 
    Auto.send("{ENTER}")
    
def Record(statement):
    with open(statement) as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',')
        
        click(Account_Balances_Box.midpoint) ####FIX THIS
        Auto.send("{ENTER}") ####FIX THIS
        #time.sleep(2)
        
        for transaction in readCSV:
            date = transaction[0]
            vendor = transaction[1]
            amount = transaction[2]
            
            
            depositentry(date,vendor,amount)
            time.sleep(1)
            
        print "Done"
            


Deposit_To_Textbox = make_checkpoint("Deposit_To",
                             (78,169),(152,169),
                             (78,186),(152,186)) 

Date_Textbox = make_checkpoint("Date",
                             (214,169),(278,169),
                             (214,186),(278,186))

Received_From_Textbox = make_checkpoint("Received_From",
                             (20,258),(145,259),
                             (20,270),(145,270))

From_Account_Textbox = make_checkpoint("From_Account",
                             (160,254),(278,254),
                             (160,269),(278,269))

Amount_Textbox = make_checkpoint("Amount",
                             (639,254),(759,269),
                             (639,269),(759,269)) # check the coords here

Account_Balances_Box = make_checkpoint("Account_Balances",
                             (721,172),(880,172),
                             (721,183),(880,183))

apptitle = "Global"
statement = "stmt2.txt"
Auto = Dispatch("AutoItX3.Control")

setup()
Record(statement)    
