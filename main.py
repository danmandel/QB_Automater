import csv
import SendKeys # maybe get rid of
from win32com.client import Dispatch
import win32con
import time
import datetime

def close_all_windows():
    print "Closing all windows at: %s" % time.strftime("%H:%M:%S")
    Auto.WinActivate(apptitle)
    Auto.send("!w")
    Auto.send("{DOWN}")
    Auto.send("{ENTER}")

def open_home():
    print "Calling open_home()"
    Auto.MouseClick("left", 34, 77)

##def delete_current_transaction(): #should only be used if you want to delete current transaction
##    Auto.WinActivate("Yuliya")
##    Auto.send("{CTRLDOWN}")
##    Auto.send("{d}")
##    Auto.send("{CTRLUP}")   
##    if Auto.WinExists("Past Transactions") == 1:
##        Auto.send("{TAB}")
##        Auto.send("{ENTER}")
##        print "Closed 'Past Transactions'"
##    print "Closed current transaction. Now_at_date"
    
    
##def select_bank(bankcode):
##    #Auto.WinActivate(apptitle)
##    BOA = 10030   # "Bank of America Business - Operating"
##    AmEx = 10400
##    Auto.send("{TAB}") # will eventually evolve into select_bank function
##    Auto.send(bankcode) #^
##    Auto.send("{ENTER}")
    
def setup(): # might eventually be a function like setup(bankcode)
    print "Calling setup()"
    Auto.WinActivate(apptitle)
    Auto.WinMove(apptitle,"", 0, 0, 1000, 1000)
    close_all_windows()
    open_home()
    time.sleep(1)
    Auto.send("{TAB}") # will eventually evolve into select_bank function
    Auto.send(bankcode) #^
    Auto.send("{ENTER}") #^
    #time.sleep(1)
    
def is_color(x,y,color):
    PositionColor = Auto.PixelGetColor(x,y)
    if color == PositionColor:
        return True
    else:
        return False     
   
def attempt_send_vendor(v,Type): 
    for letter in v[0:3]:
        Auto.send(letter)
    if is_color(300,450,black): 
        Auto.send("{TAB}") # Now un-hilighted cursor is in Payment textbox after first tab. 
        if Type == "deposit":
            Auto.send("{TAB 2}") # Now un-highlighted cursor is in Deposit textbox after 2 more tabs.
        elif Type == "credit":
            pass # Now highlighted cursor is still in Payment textbox.
        return 1   
    else:
        #Highlight failed. Cursor now at end of Account textbox.
        return 2
        
def attempt_send_amount(a,Type):       
    Auto.send(a) #Amount
    if Type == "debit":
        Auto.send("{TAB}") # End up in Accounts after one tab from deposits
        return 1
    elif Type == "credit":
        Auto.send("{TAB 3}")# End up in Accounts after 3 tabs from Payments
        return 1
    else:
        print "Failure occured : %s" % time.strftime("%H:%M:%S")
        print "Function 'attempt_send_amount' failed."
        return 2            

def attempt_send_account(Type):
    if Type == "deposit":
        account = "income"
        Auto.send(account)
        return 1
    elif Type == "credit":
        paste_account()
        return 1
    else:
        print "Failure occured : %s" % time.strftime("%H:%M:%S")
        print "Function 'attempt_send_account' failed."
        return 2
    
        
def copy_account(vendor):
    #Auto.WinActivate(apptitle)
    Auto.send("!g") # alt+g
    time.sleep(2)
    Auto.WinMove("Go To","", 0, 0, 1000, 1000)
    time.sleep(2)
    Auto.send("{TAB}")
    Auto.send(vendor) #Name
    Auto.send("{TAB}")
    
    Auto.send("ENTER}")#does back search
    if Auto.WinExists("Name Not Found"):
        Auto.send("{TAB 2}")
        Auto.send("ENTER")
        Auto.send("ESC")
    
    Auto.send("{TAB 2}")
    Auto.send("ENTER}")
    Auto.send("{TAB 4}")
    Auto.send("{CTRLDOWN}")
    Auto.send("{c}")
    Auto.send("{CTRLUP}")

def paste_account():
    Auto.send("{CTRLDOWN}")
    Auto.send("{v}")
    Auto.send("{CTRLUP}")

def exit_TransactionEntry():
    Auto.send("{ESC 2}")
    #bankcode
    #Auto.send(bankcode)
    #Auto.send("{ENTER}")
    
    
    
    
def TransactionEntry(d,v,a,Type): # Cycles for every transaction in statement if a > 0
    Auto.send(d) #Date
    Auto.send("{TAB 2}")
    attempt_send_vendor(v,Type) #if sucess, cursor ends up in Payment or Deposit. if failure, undefined.
    
    if attempt_send_vendor(v,Type) == 1: # does this also execute the commands in addition to returning 1? can i run without standalone command?
        print "attempt_send_vendor(v,Type) == 1"
        time.sleep(1)
        if attempt_send_amount(a,Type) == 1:
            print "attempt_send_amount(a,Type) == 1"
            time.sleep(1)
            if attempt_send_account(Type)==1:
                print "attempt_send_account(Type)==1"                
                Auto.send("{TAB 2}")
                print "TransactionEntry success."
                return 1
                time.sleep(1)
            else:
                Auto.send("{ESC 3}")
                setup()
                time.sleep(1)
                
                print "Method 'attempt_send_account(Type)' failed because Type =/= 1. Location: Level 3 if statement in TransactionEntry(*args) "
        else:
            Auto.send("{ESC 3}")
            setup()
            time.sleep(1)
            
            print "Method 'attempt_send_amount(a,Type)' failed because Type =/= 1."
    elif attempt_send_vendor == 2:
        Auto.send("{ESC 3}")
        #attempt_send_vendor failed. Cursor now at end of Account textbox. 
        Skipped_List.append(transaction)
        print "Failure occured : %s" % time.strftime("%H:%M:%S")
        print ("Added %s to Skipped_List: " % transaction)
        print "Info:  At 'attempt_send_vendor' : entry was not highlighted."
        print ""
        setup()
        time.sleep(1)
        

    else:
        print "what the fuck"
        time.sleep(1)
        
##    send_vendor_result = attempt_send_vendor(v,Type)
##    send_amount_result = attempt_send_amount(a,Type)
##    send_amount_result = attempt_send_account(Type)
##    print "send_vendor_result", send_vendor_result
##    print "amount_result", send_amount_result
##    print "send_account_result", send_account_result
    #print "TransactionEntry failed somewhere."
    print ""

    time.sleep(1)
        
def Record(statement):
    with open(statement) as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',')
        
        for transaction in readCSV:
            date = transaction[0]
            vendor = transaction[1]
            amount = transaction[2]
            
            if float(amount) > 0:
                Type = "deposit"
                #TransactionEntry(date,vendor,amount,Type)
                #print ("Attempting to debit:  %s to [eventual location]: " % transaction)
                print "Attempting to debit transaction.."
                if TransactionEntry(date,vendor,amount,Type) == 1:
                    print "TransactionEntry for deposit success"
                
            elif float(amount) < 0:
                Type = "credit"
                copy_account(vendor)
                #TransactionEntry(date,vendor,amount,Type)
                #print ("Credited %s to [eventual location]: " % transaction)
                print "Attempting to credit transaction.."
                if TransactionEntry(date,vendor,amount,Type) == 1:
                    print "TransactionEntry for credit success"

            else:
                Skipped_List.append(transaction)
                print ("Added %s to Skipped_List: " % transaction)      
            #time.sleep(1)
                
        print "Done"

def Process():
    setup()
    time.sleep(1)
    if is_color(464,326,grey):
        setup() #setup should leave you with date highlighted
    
    Record(statement)

    counter = 0
##    for transaction in Skipped_List:
##        print "skipped %s" % transaction 
##        counter +=1 
    
Skipped_List = []

apptitle = "Yuliya"
##statement = "stmt2.txt"
statement = "stmtsampleclean.txt"
Auto = Dispatch("AutoItX3.Control")
bankcode = "10030"
current_time = time.strftime("%H:%M:%S")
black = 0x000000 ###
grey = 0xABABAB ###

Process()
