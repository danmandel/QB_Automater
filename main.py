import csv
from win32com.client import Dispatch
import win32con
import time
import datetime

##def close_all_windows():
##    print "Closing all windows at: %s" % time.strftime("%H:%M:%S")
##    Auto.WinActivate("Yuliya")
##    Auto.WinActivate(apptitle)
##    Auto.send("!w")
##    Auto.send("{DOWN}")
##    Auto.send("{ENTER}")
##    print "Closed all windows at: %s" % time.strftime("%H:%M:%S")

def close_all_windows(): # Ends with is_color(50,50,grey == 1) aka blank screen
    print "Calling close_all_windows() at: %s" % time.strftime("%H:%M:%S")
    Auto.WinActivate(apptitle)
    Auto.send("!w")
    Auto.send("{DOWN}")
    time.sleep(1)
    Auto.send("{ENTER 2}")
    time.sleep(1)
    for x in range(4):
        if is_color(250,250,grey) == 0:
            Auto.send("{ESC}")
            print "Esc attempt %s" % x
    Auto.Send("{ENTER}")
    time.sleep(1)
    for x in range(4):
        if is_color(250,250,grey) == 0:
            Auto.send("{ESC}")
            print "Esc attempt %s" % x
    print "Ended close_all_windows() at: %s" % time.strftime("%H:%M:%S")
    print ("")

def open_home():
    print "Calling open_home() at: %s" % time.strftime("%H:%M:%S")
    Auto.MouseClick("left", 34, 77)
    print "Ended open_home() at: %s" % time.strftime("%H:%M:%S")

def tile_windows():
    Auto.send("!w") # Tile Windows()
    time.sleep(1)
    Auto.send("(DOWN 3)")
    time.sleep(1)
    Auto.send("{ENTER}")
    #ends wherever curlor left off. 

def setup(bankcode): # leaves you with Cursor in Date according to bankcode. 
    print "Calling setup() at: %s" % time.strftime("%H:%M:%S")
    Auto.WinActivate(apptitle)
    Auto.WinMove(apptitle,"", 0, 0, 1000, 1000)
    close_all_windows()
    open_home()
    time.sleep(1)
    Auto.send("{TAB}") # will eventually evolve into select_bank function
    Auto.send(bankcode) 
    Auto.send("{ENTER}")  
    #Auto.send("{BACKSPACE}")
    time.sleep(1)
    tile_windows()
    time.sleep(1)
    print "Ended setup() at: %s" % time.strftime("%H:%M:%S")
    print "Now cursor is at Date for bankcode: %s " % bankcode

def is_color(x,y,color):
    PositionColor = Auto.PixelGetColor(x,y)
    if color == PositionColor:
        return 1
    else:
        return 0

def attempt_send_vendor(v,Type): #starts at
    print ("Calling attempt_send_vendor() at: %s" % time.strftime("%H:%M:%S"))
    for letter in v[0:3]:
        Auto.send(letter)
        time.sleep(1)
    if is_color(325,452,black) == 1: #############found it## really needs be improved by making apptitle = "10030" move into upper left perfectly.
        Auto.send("{TAB}") # Now un-hilighted cursor is in Payment textbox after first tab.
        time.sleep(1)
        print "Vendor recognized in drop-down."
        if Type == "deposit":
            Auto.send("{TAB 2}") # Now un-highlighted cursor is in Deposit textbox after 2 more tabs.
            return 1
        elif Type == "credit":
            pass # Now highlighted cursor is still in Payment textbox.
            return 1
        else:
            print "Type passed though attempt_send_vendor(v,Type) is not recognized."
            return 0
    else:
        print "attempt_send_vendor_ failed" 
        #Highlight failed. Cursor now at end of Account textbox.
        return 0
    print ("Ended attempt_send_vendor() at: %s" % time.strftime("%H:%M:%S"))

        
def attempt_send_amount(a,Type):       
    Auto.send(a) #Amount
    if Type == "debit":
        Auto.send("{TAB}") # End up in Accounts after one tab from deposits
        print "amount entered for debit in attempt_send_amount(a,Type)"
        return 1
    elif Type == "credit":
        Auto.send("{TAB 3}")# End up in Accounts after 3 tabs from Payments
        print "amount entered for credit in attempt_send_amount(a,Type)"
        return 1
    else:
        print "Failure occured : %s" % time.strftime("%H:%M:%S")
        print "Function 'attempt_send_amount' failed."
        return 2            

def attempt_send_account(Type):
    if Type == "deposit":
        account = "income"
        Auto.send(account)
        print "account entered for deposit in attempt_send_account(Type)"
        return 1
        
    elif Type == "credit":
        paste_account()
        print "account entered for deposit in attempt_send_account(Type)"
        return 1
    
    
    else:
        print "Failure occured : %s" % time.strftime("%H:%M:%S")
        print "Function 'attempt_send_account' failed."
        return 0


def attempt_send_date(d):
    print ("Called attempt_send_date(d) at: %s" % time.strftime("%H:%M:%S"))
    time.sleep(1)
    Auto.send(d)
    Auto.send("{TAB 2}")
    print ("Ended attempt_send_date() at: %s" % time.strftime("%H:%M:%S"))
    print ("")
    time.sleep(3)
    

##add transaction Id
##    make Transaction a class

##class Transaction(object):
##    Transaction.
##
##def make_transaction(Date,Name,amount,Type,transaction) #lowercase transaction is the each in for each transaction

    
def DepositEntry(d,v,a,Type,transaction): # starts with Date in deposits highlighted
    #print ("Calling DepositEntry() for transaction number  %s: " % transaction) ######
    print ("Calling DepositEntry() at: %s" % time.strftime("%H:%M:%S"))
    #time.sleep(1)
    attempt_send_date(d) #Cursor ends up at Account (aka vendor)
    if attempt_send_vendor(v,Type) == 1: # also executes the function ### how to make it use value without executing?
        print "attempt_send_vendor(v,Type) == 1"
        time.sleep(1)
        if attempt_send_amount(v,Type) == 1:
            print "attempt_send_amount(a,Type) == 1"
            time.sleep(1)
            if attempt_send_account(Type) == 1:
                time.sleep()
                print "DepositEntry sucess"
                return 1           
    else:
        print "DepositEntry failure"
        return 0
        time.sleep(1)
        
def CreditEntry(d,v,a,Type,transaction):
    #print ("Attempting to credit:  %s to [bankcode]: " % transaction) #######
    print "CreditEntry pass on name: %s " % v

def Process(statement):
    #Auto.WinActivate(apptitle)
    #setup(bankcode)
    with open(statement) as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',')
        counter = 0
        for transaction in readCSV:
            
            setup(bankcode) ################### ################## remove this to avoid setup every turn #tile windows
            date = transaction[0]
            vendor = transaction[1]
            amount = transaction[2]
            
            if float(amount) > 0: # Deposit
                Type = "deposit"
                DepositEntry(date,vendor,amount,Type,transaction)
                
            elif float(amount) < 0: # Credit
                Type = "credit"
                #copy_account(vendor)
                CreditEntry(date,vendor,amount,Type,transaction)
                #print ("Credited %s to [eventual location]: " % transaction)

            else:
                Skipped_List.append(transaction)
                #print ("Added %s to Skipped_List: " % transaction) # does print run the function??
                print "How is this not a credit or deposit? Messed up in Process()"
            #time.sleep(1)
            print "Finished Transaction number: %s" % counter
            counter += 1 
            print ""
            print "______________________________"
            print ""
                
        print "Processed all transactions at time: "
         
    
Skipped_List = []

apptitle = "Yuliya"
##statement = "stmt2.txt"
statement = "stmtsampleclean.txt"
Auto = Dispatch("AutoItX3.Control")
bankcode = "10030"
current_time = time.strftime("%H:%M:%S")
black = 0x000000 ###
grey = 0xABABAB ###

#Auto.WinActivate(apptitle)
Process(statement)

##Auto.WinActivate("Yuliya")
##time.sleep(1)
print Auto.WinExists("Bank")

print "Finished Process"

#print Skipped_List

#print Auto.WinExists("Yuliya")
