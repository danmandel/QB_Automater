import csv
import SendKeys
import win32api
from win32com.client import Dispatch
import win32con
import time

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
    time.sleep(1)
    Auto.send("{TAB}")
    Auto.send(bankcode)
    Auto.send("{ENTER}")
    time.sleep(1)
    Auto.send("!w")
    time.sleep(1)
    Auto.send("{DOWN 3}")
    time.sleep(1)
    Auto.send("{ENTER}")
    # right now leaves me with date selected in BOA

    

   
    #add test for if Home..basically if not grey background
    #grey = 0xABABAB

##def move_to(coordinates):
##    win32api.SetCursorPos(coordinates)
    
def click(coordinates):
    win32api.SetCursorPos((coordinates))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,coordinates[0],coordinates[1],0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,coordinates[0],coordinates[1],0,0)

def check_if_black(x,y):
    #if blue
    #coord = (191,412)
    color = 0x000000               # 0xFFFFFF         #(x,y#Auto.PixelGetColor(x,y #black
    if Auto.PixelGetColor(x,y) == color: 
        print "TRUE!!!!!!!!!!!!black!!!!!!!!!!!!!"
        ##change to black, general preferences automatically enter dropdown list while typing
        return True
        
    else:
        print "fals not black" 
        return False

def check_if_grey(x,y):
    #if blue
    #coord = (191,412)
    color = 0xABABAB             
    if Auto.PixelGetColor(x,y) == color: 
        print "TRUE!!!!!!!!grey!!!!!!!!!!!!!!!!!"
        return True
        
    else:
        print "false not grey" 
        return False
    

def send_vendor(v,a): # type 4 letters. if checkifblack == true: send. else: append skipped
    #black = 
    counter = 0
    for letter in v[0:3]:
        Auto.send(letter)
        counter += 1
        time.sleep(.5)
        print counter
    if check_if_black(300,450) == True: 
        Auto.send("{TAB}") # now in payment after one tab
        print "Sent Vendor: "
        Auto.send("{TAB 2}") #2 for deposit 0 for payment
        Auto.send(a) #Amount
        Auto.send("{TAB}")
        Auto.send("income") #Income. will need to be xpanded for credits to include lookup
        Auto.send("{TAB 2}")
        time.sleep(2)
        Auto.send("{ENTER}") # this en
        return True
        time.sleep(1)
        
    else:
        Skipped_List.append(v)
        #print ("Skipped list: " , Skipped_List)
        print "skipped a vendor"
        Auto.send("{ESC}")
        Auto.Send("+{TAB 2}") # now is at Date as if nothing ever happened.      
        return False

        
     
def DepositEntry(d,v,a): # Cycles for every transaction in statement if a > 0
    #add check if you're in the deposits/credits screen

    Auto.send(d) #Date
    Auto.send("{TAB 2}")
    #Auto.send(v) #Vendor
    send_vendor(v,a)
##    check = send_vendor(v)
##    if check == False:
##            Auto.send("{ESC}")
##            Send("+{TAB 2}") # now is at Date as if nothing ever happened.      
      

def Record(statement):
    with open(statement) as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',')
        Auto.send("{ENTER}") ####FIX THIS
        #time.sleep(2)
        
        for transaction in readCSV:
            date = transaction[0]
            vendor = transaction[1]
            amount = transaction[2]
            
            if float(amount) > 0:
                DepositEntry(date,vendor,amount)
                print ("Attempting to debit:  %s to [eventual location]: " % transaction)
                
            elif float(amount) < 0:
                #CreditEntry(date,vendor,amount)
                print ("Credited %s to [eventual location]: " % transaction)
            else:
                Skipped_List.append(transaction)
                print ("Added %s to Skipped_List: " % transaction)
            
            #time.sleep(1)
            
        print "Done"
            
def Process():
    setup()
    time.sleep(1)
    if check_if_grey(464, 326) == True:
        setup()
    Record(statement)

Skipped_List = []

apptitle = "Yuliya"
##statement = "stmt2.txt"
statement = "stmtsampleclean.txt"
Auto = Dispatch("AutoItX3.Control")
bankcode = "10030"

##setup()
##time.sleep(1)
##Record(statement)    
Process()
    
