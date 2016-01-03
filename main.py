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
        if not is_color(250,250,grey):
            Auto.send("{ESC}")
            print "Esc attempt %s" % x
    Auto.Send("{ENTER 2}")
    time.sleep(1)
    if Auto.WinExists("Past Transactions"):
        time.sleep(2)
        print "'Past Transactions' warning message exists."
        Auto.send("n") 
    for x in range(5):
        if not is_color(250,250,grey):
            Auto.send("{ESC}")
            print "Esc attempt %s" % x
   
def check_checker():
    if is_color(35,480,blackish_grey):
        print "'1-Line' box is currently checked."
        return 1
    else:
        return 0
    
def partially_type(text,n):
    # print "Entering first %s letters of %s" % (n, text)
    for letter in text[0:n]:
        Auto.send(letter)
        time.sleep(.5)
    
def open_register(bank_code):
    # Should be started at a blank screen.
    # Ends at bank register with "Date" textbox highlighted.
    #Auto.send("!c") # Opens "Company" menu.
    #Auto.send("h") # Selects "home".
    #Auto.send("{TAB}") # Activates the bank selection window.
    Auto.send("^r") # Opens register.
    Auto.send(bank_code) # Types in bank_code.
    Auto.send("{ENTER}") # Brings up register.
      
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
    if request_confirmation: 
        Auto.WinActivate("C:\Python27")
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
    else:
        return 1
    Auto.WinActivate(apptitle)

def tile_windows():
    Auto.send("!w")
    Auto.send("h") # Chooses "home" dropdown option. Ends wherever curlor left off. 
     
def is_color(x,y,color):

    return (color == Auto.PixelGetColor(x,y))

def error_checker():
    if Auto.WinExists("Account Not Found"):
        return 1
    if Auto.WinExists("Name Not Found"):
        return 1
    if Auto.WinExists("Select Name Type"):
        return 1
    if Auto.WinExists("Warning"): 
        return 1   
    
def go_to_date():
    Auto.send("^d")
    if Auto.WinExists("Past Transactions"):
        Auto.send("n")
    if Auto.WinExists("Delete Transaction"):
        Auto.send("{ESC}") # lazy fix because this shouldnt show up if you do it correctly
   
def copy_account(vendor):
    # Start in vendor textbox for transaction.
    Auto.send("!g")
    Auto.send("!s") # Now highlighted cursor is in "Search for: "
    partially_type(vendor,n)
    Auto.send("{TAB}")
    time.sleep(sleep)
    if error_checker == 1:
        print "Credit's Account name not found when copying"
        auto.send("c") 
        Auto.send("{ESC}")
        return 0 # now in acc portion
        if Auto.WinExists("Recording Transaction"):
            auto.send("n")
            return 0
          
def paste_account():
    #ctrl+v
    pass

class Transaction(object):
    def __init__(self, transaction):
        self.date = transaction[0]
        self.vendor = transaction[1]
        self.amount = transaction[2]
            
    def Determine_Type(self):
        if float(self.amount) > 0:
            self.Type = "debit"
        elif float(self.amount) < 0:
            self.Type = "credit"
        else:
            self.Type = "Amount = 0"

    def Send_Date(self): 
    # Starts at "Date" textbox. 
    # Ends with cursor in "Received From" textbox.
        print "sending date"
        Auto.send("^d")
        time.sleep(sleep)
        Auto.send(self.date)
        print "Entering date: %s" % self.date
        if self.Type == "debit":
            Auto.send("{TAB 2}")
            if error_checker() == 1:
                Auto.send("{ESC}") # Need to restart transaction now.
                print "Warning encountered. Moving onto next transaction."
                return False
            else:
                return True
        elif self.Type == "credit":
            Auto.send("{TAB 2}")
            if error_checker() == 1:
                Auto.send("{ESC}") # Need to restart transaction now.
                return False
            else:
                return True
        else:
            print "Not credit or debit. Check Type. "
            return False

    def Send_Vendor(self): # Starts at "Payee" textbox.
        #print ("Attempting to enter vendor: %s" % v)
        print "sending vendor"
        partially_type(self.vendor,n) # n letters
   
        if is_color(325,452,black):
            Auto.send("{TAB}") # Now highlighted cursor is in "Payment" textbox.
            if self.Type == "debit":
                Auto.send("{TAB 2}") # Ends with un-highlighted cursor in "Payment" textbox.
                return True
            elif self.Type == "credit":
                #Auto.send("{TAB}") # Ends with un-highlighted cursor is in "Charge" textbox.
                return True
            else:
                print ("Error, attempted to pass type: %s through attempt_send_vendor()" % Type)
                return False

            if error_checker() == 1:
                Auto.send("{ESC}") # Need to restart transaction now.
                print "NNF exists in attempt_send_vendor()"
                time.sleep(10)
                return False
              
        else:
            print "attempt_send_vendor(v,Type) failed. Check: is_color()" 
            return False

        if Auto.WinExists("Name Not Found"):
            Auto.send("c") 
            print "NNF exists in attempt_send_vendor()"

    def Send_Amount(self): 
        print "Entering amount: %s" % self.amount      
        Auto.send(self.amount)
        time.sleep(sleep)
        if self.Type == "debit":
            Auto.send("{TAB}") # Now at "Account" textbox.     
            return True
        elif self.Type == "credit":
            Auto.send("{TAB 3}")# Now at "Account" textbox. 
            return True
        else:
            print "Function 'attempt_send_amount()' failed."
            print "Entering Type: %s" % self.Type
            return False  
        if error_checker() == 1:
             Auto.send("{ESC}") # Need to restart transaction now.
             print "NNF in send_amount()?"
             return False
        else:
             pass

    def Send_Account(self):
        if self.Type == "debit":
            self.account = "income"
            Auto.send(self.account)
            Auto.send("{TAB 2}")

            if input_check("Enter this transaction? y/n") == 1:
                Auto.WinActivate(apptitle)
                Auto.send("{ENTER}")
            else:
                Auto.WinActivate(apptitle)
                print "no"
                Auto.send("c")
                return False
            if error_checker == 1:
                Auto.send("c") # Need to restart transaction now.
                print "ANF error in send_account()"
                return False
            else: 
                return True
        
        elif self.Type == "credit":
            paste_account()
            print "account entered for deposit in attempt_send_account(Type)"
            if error_checker() == 1:
                Auto.send("c") # Need to restart transaction now.
                print "ANF error in send_account()"
                return False
            else:
                return True
        else:
            print "Function 'attempt_send_account' failed."
            return False 
   
    def Transaction_Entry(self):
       if (self.Send_Date() and self.Send_Vendor() and self.Send_Amount() and self.Send_Account()):
           print "Success"
           return True
       else:
           print "Failure"
           Skipped_List.append(self)
           return False
                
def Process(statement):
    setup() # Ends at "Date" textbox.  
    with open(statement) as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',') 
        for transaction in readCSV:
            Current_Transaction = Transaction(transaction) #if i could end this function here thatd be great
            Current_Transaction.Determine_Type()
            time.sleep(sleep)
            Current_Transaction.Transaction_Entry()                                 
            Transaction_List.append(Current_Transaction)                    
            #print "Finished Transaction number: %s" % counter 
            print "______________________________"            
                
    print "Processed all transactions at: %s" % current_time
         

n = 7 # length for partially_type()         
Auto = Dispatch("AutoItX3.Control")
current_time = time.strftime("%H:%M:%S")
black = 0x000000 
blackish_grey = 0x484848
blue = 0x3399FF
grey = 0xABABAB 
white = 0xFFFFFF
Transaction_List = []
Skipped_List = []


#### Settings ####
apptitle = "Yuliya"
statement = "C:\Python27\Scripts\QB\stmtsampleclean.txt"
bank_code = "Bank of America Bus"
do_credits = 0
request_confirmation = True
# ^ do for debits and credits separately
sleep = 1
#### Settings ####

Process(statement)

time.sleep(100)

