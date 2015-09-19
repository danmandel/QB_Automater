import csv
from win32com.client import Dispatch
import win32con
import time
import datetime

def close_all_windows():
    # Starts anywhere. 
    # Ends with blank screen.
    #print "Calling close_all_windows() at: %s" % time.strftime("%H:%M:%S")
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
    #print "Ended close_all_windows() at: %s" % time.strftime("%H:%M:%S")
def check_checker():
    if is_color(35,480,blackish_grey) == 1:
        print "'1-Line' box is currently checked."
        return 1
    else:
        print "no check "
        return 0
    
def open_make_deposits(bank_code): 
    # Starts at blank screen. 
    # Ends with cursor at "Date" textbox.
    #print ("Calling open_make_deposits(bank_code) at: %s" % time.strftime("%H:%M:%S"))
    Auto.send("!b")
    Auto.send("d")
    tile_windows()
    time.sleep(1)
    for letter in bank_code[0:3]:
        Auto.send(letter)
        time.sleep(1)
    if is_color(115,177,blue) == 1: # Can avoid this by typing out bank_code, tabbing, then checking for error message.
        Auto.send("{TAB}") # Now un-highlighted cursor is in "Date" textbox. 
        time.sleep(1)
    elif is_color(135,178,blue) == 1: # Backup highlight checker.
        Auto.send("{TAB}") # Now un-highlighted cursor is in "Date" textbox.
        time.sleep(1)
    else:
        print "Bank_code not recognized. Check is_color() coordinates."  ### get rid of this
    # print ("Ended open_make_deposits(bank_code) at: %s" % time.strftime("%H:%M:%S"))

def open_make_credits(bank_code):
    # Starts at blank screen. 
    # Ends with cursor at "Date" textbox.
    print ("Calling open_make_deposits(bank_code) at: %s" % time.strftime("%H:%M:%S"))
    Auto.send("{CTRLDOWN}")
    Auto.send("w")
    Auto.send("{CTRLUP}")
    Auto.send("!k") # Now at "Bank Account" textbox.
    tile_windows()
    time.sleep(1)
    print ("Ended open_make_deposits(bank_code) at: %s" % time.strftime("%H:%M:%S")) ### combine with open_make_deposits or delete both()

def open_register(bank_code):
    Auto.WinActivate(apptitle)
    Auto.WinMove(apptitle,"", 0, 0, 1000, 1000)
    #close_all_windows()
    Auto.send("!c")
    Auto.send("h") # chooses "home" dropdown option
    Auto.send("{TAB}") # activates the bank selection window
    Auto.send(bank_code) # Types in bank_code
    Auto.send("{ENTER}") # Brings up bank register
    # Ends at bank register with "Date" textbox highlighted.
   
def setup():
    #close_all_windows()

     Auto.WinActivate(apptitle)
     Auto.WinMove(apptitle,"", 0, 0, 1000, 1000)
     close_all_windows()
     time.sleep(sleep)
     open_home()
     open_register()
     if check_checker() == 1:
         Auto.send("!1") # "Alt + 1" to turn off 1_line box.

def input_check(goal):
    testVar = raw_input("True?: x%s" % goal )
    if testVar == "y":
        print "cool"
    elif testVar == "n":
        print "not okay man"
    else:
        print "wut"

def tile_windows():
    Auto.send("!w")
    Auto.send("h") # Chooses "home" dropdown option. Ends wherever curlor left off. 
     
def is_color(x,y,color):
    PositionColor = Auto.PixelGetColor(x,y)
    if color == PositionColor:
        return 1
    else:
        return 0

def attempt_send_vendor(v,Type): # Starts at Payee
    #print ("Calling attempt_send_vendor() at: %s" % time.strftime("%H:%M:%S"))
    print ("Attempting to enter vendor: %s" % v)
    for letter in v[0:3]:
        Auto.send(letter)
        time.sleep(1)
    if is_color(325,452,black) == 1: # or (325,452, black) or (330,468,black,Uglyregister)
        Auto.send("{TAB}") # Now highlighted cursor is in "Payment" textbox.
        time.sleep(1)
        if Type == "debit":
            Auto.send("{TAB 2}") # Ends with un-highlighted cursor in "Payment" textbox.
            return 1
        elif Type == "credit":
            #Auto.send("{TAB}") # Ends with un-highlighted cursor is in "Charge" textbox.
            return 1
        else:
            print ("Error, attempted to pass type: %s through attempt_send_vendor()" % Type)
            return 0
    else:
        print "attempt_send_vendor(v,Type) failed. Check: is_color()" 
        # Highlight failed. Cursor still at end of "Account" textbox.
        return 0
    #print ("Ended attempt_send_vendor() at: %s" % time.strftime("%H:%M:%S"))

def attempt_send_vendor_deposit(v): # Ends at "Amount" textbox
    print ("Calling attempt_send_vendor() at: %s" % time.strftime("%H:%M:%S"))
    print ("Attempting to enter vendor: %s" % v)
    for letter in v[0:3]:
        Auto.send(letter)
        time.sleep(1)
    if is_color(170,300,black) == 1:
        Auto.send("{TAB}") # Now hilighted cursor is in "From Account" textbox.
        Auto.send("Income")
        Auto.send("{TAB 4}") # Now un-highlighted cursor is in "Amount" textbox.
        time.sleep(2)
        return 1
    else:
        print "Attempt to send vendor failed." 
        return 0
    #print ("Ended attempt_send_vendor() at: %s" % time.strftime("%H:%M:%S"))
        
def attempt_send_amount(a,Type): 
    print "Entering amount: %s" %a      
    Auto.send(a) #Amount
    time.sleep(sleep)
    if Type == "debit":
        Auto.send("{TAB}") # Now at "Account" textbox for a new transaction.
        '''if Auto.WinExists("Past Transactions"):
            Auto.send("y")
            print "Saving transaction >30 days in the past." # Can just delete this message in preferences.'''          
        return 1
    elif Type == "credit":
        Auto.send("{TAB 3}")# End up in Accounts after 3 tabs 
        return 1
    else:
        #print "Failure occured : %s" % time.strftime("%H:%M:%S")
        print "Function 'attempt_send_amount' failed."
        print "LMAOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO"
        print "Entering Type: %s" % Type
        return 0            

def attempt_send_account(Type):
    if Type == "debit":
        account = "income"
        Auto.send(account)
        print "account entered for deposit in attempt_send_account(Type): %s" % account
        return 1
        
    elif Type == "credit":
        paste_account()
        print "account entered for deposit in attempt_send_account(Type)"
        return 1
    
    
    else:
        print "Failure occured : %s" % time.strftime("%H:%M:%S")
        print "Function 'attempt_send_account' failed."
        return 0 # Might make individual for Deposit and Credit

def attempt_send_date(d,Type): 
    # Starts  
    # Ends with cursor in "Received From" textbox.
    print ("Called attempt_send_date(d) at: %s" % time.strftime("%H:%M:%S"))
    time.sleep(1)
    Auto.send(d)
    if Type == "debit":
        Auto.send("{TAB 2}")
    elif Type == "credit":
        Auto.send("{TAB 2}")
        print "Should be at Account. Confirm?"
        #if inputcheck()=="y"

        time.sleep(2)
    else:
        print "Not credit or debit. Check Type. "
    #print ("Ended attempt_send_date() at: %s" % time.strftime("%H:%M:%S"))
    time.sleep(2)

def Transaction_Entry(d,v,a,Type,transaction): 
    #AKA Transaction_Entry
    #print ("Attempting to credit:  %s to [bank_code]: " % transaction) #######
    #print "CreditEntry pass on name: %s " % v
    attempt_send_date(d,Type) # Ends with cursor in "Pay to the order of (vendor)" textbox.
    if attempt_send_vendor(v,Type) == 1: # Ends with cursor at "Amount" textbox.
        print "attempt_send_vendor(v) == 1"
        time.sleep(1)
        if attempt_send_amount(a,Type) == 1:
            time.sleep(1)
            return 1 
            if attempt_send_account(Type) == 1:
                print "Transaction_Entry sucess"
                return 1
    else:
        print "Transaction_Entry failure" # needs to be printed for every if but will be replaced with assert
        return  0
        time.sleep(1)
   
def DepositEntry(d,v,a,Type,transaction): # starts with Date in deposits highlighted
    #print ("Calling DepositEntry() at: %s" % time.strftime("%H:%M:%S"))
    attempt_send_date(d, Type) # Ends with cursor in "Received From (vendor)" textbox.
    if attempt_send_vendor_deposit(v) == 1: # Ends with cursor at "Amount" textbox.
        print "attempt_send_vendor(v) == 1"
        time.sleep(1)
        if attempt_send_amount(a,Type) == 1:
            print "DepositEntry sucess"
            time.sleep(1)
            return 1 
    else:
        print "DepositEntry failure"
        return  0
        time.sleep(1)
    #print ("Ended DepositEntry() at: %s" % time.strftime("%H:%M:%S"))

    obsession
    noissesbso
        
def copy_account(vendor):
    #start in vendor textbox for transaction. or start in empty screen?
    #Auto.send("!g")
    #Perhaps press S to get into search for, then start typing vendor
    #if "warning..No more transactions found that match the criteria. Conitue searching from beginning?" Y/N 
    # if y it means you keep looking and window goes away
    #press esc to go to record and youll be in the date section
    #5 tabs and you shuld be in accounttextbox
    # [y/n] promt here
    #auto.send ctrl c
    #go to fresh transaction box <- can be a method. but can be replaced with resetup
    print 'poop'

    0x6B6B6B

def Process(statement):
    Auto.WinActivate(apptitle)
    with open(statement) as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',')
        counter = 0
        for transaction in readCSV:
            date = transaction[0]
            vendor = transaction[1]
            amount = transaction[2]

            #if option.print_transactiondate == 1,
                #print date, vendor, amount, etc
            #if option.print_transactiondate == 1,
            #if option.print_transactiondate == 1,
            
            
            if float(amount) > 0: # Debit.
                Type = "debit"
                close_all_windows()
                open_register(bank_code)
                #open_make_deposits(bank_code) # Ends with cursor at "Date" textbox.
                time.sleep(1)
                Transaction_Entry(date,vendor,amount,Type,transaction)
                
            elif float(amount) < 0: # Credit.
                Type = "credit"
                close_all_windows
                #copy_account(vendor) 3 goes before creditentry aka transactionentry
                #open_make_credits(bank_code) # Starts at blank screen. # Ends with cursor at "Date" textbox.
                Transaction_Entry(date,vendor,amount,Type,transaction)
                time.sleep(1)
                print ("Credited %s to [eventual location]: " % transaction)

            else:
                Skipped_List.append(transaction)
                #print ("Added %s to Skipped_List: " % transaction) # does print run the function??
                print "Error in in Process(), amount is not > or < 0."

            print "Finished Transaction number: %s" % counter
            counter += 1 
            print "______________________________"
            print ""
                
        print "Processed all transactions at: %s" % current_time
         
            
Auto = Dispatch("AutoItX3.Control")
current_time = time.strftime("%H:%M:%S")
black = 0x000000 
blackish_grey = 0x484848
blue = 0x3399FF
grey = 0xABABAB 
white = 0xFFFFFF

Skipped_List = []

#### Settings ####
apptitle = "Yuliya"
statement = "C:\Python27\Scripts\QB\stmtsampleclean.txt"
bank_code = "Bank of America Bus"
sleep = 2
#### Settings ####


Process(statement)




#check_checker()
