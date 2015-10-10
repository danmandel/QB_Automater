import csv
from win32com.client import Dispatch
import win32con
import time
import datetime

def close_all_windows():
    # Starts anywhere. 
    # Ends with blank screen.
    Auto.WinActivate(apptitle)
    Auto.send("!w")
    Auto.send("a")
    if Auto.WinExists("Recording"):
        time.sleep(2)
        print "'Do you want to record this transaction?' warning message exists."
        Auto.send("n")      
    time.sleep(1)
    for x in range(5):
        if is_color(250,250,grey) == 0:
            Auto.send("{ESC}")
            print "Esc attempt %s" % x
    Auto.Send("{ENTER 2}")
    time.sleep(1)
    if Auto.WinExists("Past Transactions"):
        time.sleep(2)
        print "'Past Transactions' warning message exists."
        Auto.send("n") 
    for x in range(5):
        if is_color(250,250,grey) == 0:
            Auto.send("{ESC}")
            print "Esc attempt %s" % x
   
def check_checker():
    if is_color(35,480,blackish_grey) == 1:
        print "'1-Line' box is currently checked."
        return 1
    else:
        print "no check "
        return 0
    
def partially_type(text,n):
    # print "Entering first %s letters of %s" % (n, text)
    for letter in text[0:n]:
        Auto.send(letter)
        time.sleep(.5)
    
def open_register(bank_code):
    # Should be started at a blank screen.
    Auto.send("!c") # Opens "Company" menu.
    Auto.send("h") # Selects "home".
    Auto.send("{TAB}") # Activates the bank selection window.
    Auto.send(bank_code) # Types in bank_code.
    Auto.send("{ENTER}") # Brings up register.
    # Ends at bank register with "Date" textbox highlighted.
    
   
def setup():
     Auto.WinActivate(apptitle)
     Auto.WinMove(apptitle,"", 0, 0, 1000, 1000)
     close_all_windows()
     time.sleep(sleep)
     open_register(bank_code)
     if check_checker() == 1:
         Auto.send("!1") # "Alt + 1" to turn off 1_line box.
    # Ends at "Date" textbox.

def input_check(goal):
    testVar = raw_input("True?: x%s" % goal )
    if testVar == "y":
        print "Success"
        return 1
    elif testVar == "n":
        print "Failure"
        return 0
    else:
        print "Must be y/n" 
        return 0

def tile_windows():
    Auto.send("!w")
    Auto.send("h") # Chooses "home" dropdown option. Ends wherever curlor left off. 
     
def is_color(x,y,color):
    PositionColor = Auto.PixelGetColor(x,y)
    if color == PositionColor:
        return 1
    else:
        return 0

def attempt_send_vendor(v,Type): # Starts at "Payee" textbox.
    #print ("Attempting to enter vendor: %s" % v)
    partially_type(v,n) # n letters
   
    if is_color(325,452,black) == 1:
        Auto.send("{TAB}") # Now highlighted cursor is in "Payment" textbox.

        if Type == "debit":
            Auto.send("{TAB 2}") # Ends with un-highlighted cursor in "Payment" textbox.
            return 1
        elif Type == "credit":
            #Auto.send("{TAB}") # Ends with un-highlighted cursor is in "Charge" textbox.
            return 1
        else:
            print ("Error, attempted to pass type: %s through attempt_send_vendor()" % Type)
            return 0

        if Auto.WinExists("Name Not Found"):
            Auto.send("c") # Need to restart transaction now.
            print "NNF exists in attempt_send_vendor()"
            time.sleep(10)
            return 0
              
    else:
        print "attempt_send_vendor(v,Type) failed. Check: is_color()" 
        return 0

    if Auto.WinExists("Name Not Found"):
        Auto.send("c") 
        print "NNF exists in attempt_send_vendor()"

def attempt_send_amount(a,Type): 
    print "Entering amount: %s" % a      
    Auto.send(a) #Amount
    time.sleep(sleep)
    if Type == "debit":
        Auto.send("{TAB}") # Now at "Account" textbox.     
        return 1
    elif Type == "credit":
        Auto.send("{TAB 3}")# Now at "Account" textbox. 
        return 1
    else:
        print "Function 'attempt_send_amount()' failed."
        print "Entering Type: %s" % Type
        return 0  
    if Auto.WinExists("Warning"):
            Auto.send("{ENTER}") # Need to restart transaction now.
            return 0
    if Auto.WinExists("Name Not Found"):
         Auto.send("c") # Need to restart transaction now.
         print "NNF in send_amount()?"
         return 0
    else: 
        return 1
              
def attempt_send_account(Type):
    if Type == "debit":
        account = "income"
        Auto.send(account)
        print "account entered for deposit in attempt_send_account(Type): %s" % account
        Auto.send("{TAB 2}")
        if input_check("Enter this transaction?")==1:
            Auto.WinActivate(apptitle)
            Auto.send("{ENTER}")
        else:
            print "no"
        if Auto.WinExists("Account Not Found"):
            Auto.send("c") # Need to restart transaction now.
            print "ANF in send_account()?"
            return 0
        else: 
            return 1
        if Auto.WinExists("Name Not Found"):
            Auto.send("c") # Need to restart transaction now.
            print "NNF in send_account()?"
            return 0
        else: 
            return 1
        
    elif Type == "credit":
        paste_account()
        print "account entered for deposit in attempt_send_account(Type)"
        if Auto.WinExists("Account Not Found"):
            Auto.send("c") # Need to restart transaction now.
            print "ANF in send_account()?"
            return 0
        else:
            return 1
    
    else:
        print "Function 'attempt_send_account' failed."
        return 0 

def attempt_send_date(d,Type): 
    # Starts at "Date" textbox. 
    # Ends with cursor in "Received From" textbox.
    Auto.send(d)
    print "Entering date: %s" % d
    if Type == "debit":
        Auto.send("{TAB 2}")
        if Auto.WinExists("Warning"):
            Auto.send("{ENTER}") # Need to restart transaction now.
            print "Warning encountered. Moving onto next transaction."
            return 0
        else:
            return 1 
    elif Type == "credit":
        Auto.send("{TAB 2}")
        if Auto.WinExists("Warning"):
            Auto.send("{ENTER}") # Need to restart transaction now.
            return 0
        else:
            return 1
    else:
        print "Not credit or debit. Check Type. "
        return 0
   
def copy_account(vendor):
    # Start in vendor textbox for transaction.
    Auto.send("!g")
    Auto.send("!s") # Now highlighted cursor is in "Search for: "
    partially_type(vendor,n)
    Auto.send("{TAB}")
    time.sleep(sleep)
    if Auto.WinExists("Name Not Found"):
        print "Credit's Account name not found when copying"
        auto.send("c") 
        Auto.send("{ESC}")
        if Auto.WinExists("Recording Transaction"):
            auto.send("n")
            
def paste_account():
    #ctrl+v
    pass

def Transaction_Entry(d,v,a,Type,transaction):

    if attempt_send_date(d,Type) == 1:
        time.sleep(sleep)
        if attempt_send_vendor(v,Type) == 1:
            time.sleep(sleep)
            if attempt_send_amount(a,Type) == 1:
                time.sleep(sleep)
                if attempt_send_account(Type) == 1:
                    time.sleep(sleep)
                    print "success"
                    
                else: 
                    print "Send_Account() failed."        
                    Skipped_List.append(transaction)       
            else:
                print "Send_Amount() failed." 
                Skipped_List.append(transaction) 
        else:
            print "Send_Vendor() failed."
            Skipped_List.append(transaction)  
    else: 
        print "Send_Date() failed."
        Skipped_List.append(transaction) 

def Process(statement):
    setup() # Ends at "Date" textbox.

    with open(statement) as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',')
        debitcounter = 0
        creditcounter = 0
        for transaction in readCSV:
            date = transaction[0]
            vendor = transaction[1]
            amount = transaction[2]
             
            if float(amount) > 0: # Debit.
                Type = "debit"
                Transaction_Entry(date,vendor,amount,Type,transaction)
                
            elif float(amount) < 0: # Credit.
                Type = "credit"             
                Transaction_Entry(date,vendor,amount,Type,transaction)
            else:
                note_skipped_transaction(Type, Transaction)
                print "Error in Process(), amount = %s" % amount

        
            #print "Finished Transaction number: %s" % counter 
            print "______________________________"
            print ""
                
        print "Processed all transactions at: %s" % current_time
         
n = 7 # length for partially_type()         
Auto = Dispatch("AutoItX3.Control")
current_time = time.strftime("%H:%M:%S")

black = 0x000000 
blackish_grey = 0x484848
blue = 0x3399FF
grey = 0xABABAB 
white = 0xFFFFFF
#counter = 0
Skipped_List = []

#### Settings ####
apptitle = "Yuliya"
statement = "C:\Python27\Scripts\QB\stmtsampleclean.txt"
bank_code = "Bank of America Bus"
sleep = 1
#add option to request confirmation after every transaction
#### Settings ####

Process(statement)

print Skipped_List
