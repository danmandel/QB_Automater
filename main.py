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
            #print "Esc attempt %s" % x
    Auto.Send("{ENTER 2}")
    time.sleep(1)
    if Auto.WinExists("Past Transactions"):
        time.sleep(2)
        print "'Past Transactions' warning message exists."
        Auto.send("n") 
    for x in range(5):
        if not is_color(250,250,grey):
            Auto.send("{ESC}")
               
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

    Auto.send("^r") # Opens register.
    Auto.send(bank_code) # Types in bank_code.
    Auto.send("{ENTER}") # Brings up register.
      
def setup():
     Auto.WinActivate(apptitle)
     Auto.WinMove(apptitle,"", 0, 0, 1000, 1000)
     close_all_windows()
     time.sleep(sleep)
     open_register(bank_code)
     tile_windows()
     if check_checker() == 1:
         Auto.send("!1") # "Alt + 1" to turn off 1_line box.
    # Ends at "Date" textbox.
def entered_y():
    if Auto.WinExists("C:\Python27"):
        print "winexists"
    Auto.WinMove("C:\Python27","", 1100, 200, 500, 500)
    Auto.Send("^{TAB}")
    time.sleep(1)
    #Auto.WinActivate("C:\Python27")
    test_var = raw_input("Enter this Transaction? y/n")
    if test_var == "y":
        return True
    else:
        return False
    
def tile_windows():
    Auto.send("!w")
    Auto.send("h") # Chooses "home" dropdown option. Ends wherever curlor left off. 
     
def is_color(x,y,color):
    if color == Auto.PixelGetColor(x,y):
        return True
    else: 
        False

def errors_exist():
    if Auto.WinExists("Account Not Found"):
        print "ANF"
        return True
    if Auto.WinExists("Name Not Found"):
        print "NNF"
        return True
    if Auto.WinExists("Select Name Type"):
        print "SNT"
        return True
    if Auto.WinExists("Warning"): 
        print "W"
        return True   
         
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
    # Starts at register. 
    # Ends with cursor in "Payee" textbox.
        Auto.send("^d") # Moves cursor to "Date" textbox.
        #print "Entering date: %s" % self.date
        time.sleep(sleep)
        Auto.send(self.date)  
         
        Auto.send("{TAB 2}") # Moves cursor to "Payee" textbox
        if errors_exist(): # Need to restart transaction now. 
            print "errors in date"           
            return False        
        else:
            return True
        
    def Send_Vendor(self): 
        # Starts at "Payee" textbox.
        print "Entering vendor: %s " % self.vendor
        partially_type(self.vendor,n) # n letters
   
        if is_color(572,900,green):
            Auto.send("{TAB}") # Now highlighted cursor is in "Payment" textbox.
            if self.Type == "debit":
                Auto.send("{TAB 2}") # Ends with un-highlighted cursor in "Deposit" textbox.
                
            elif self.Type == "credit":
                Auto.send("{TAB}") # Ends with un-highlighted cursor is in "Payment" textbox.
                
            else:
                print ("Error, attempted to pass type: %s through Send_Vendor()" % self.Type)
                return False

            if errors_exist():  
                print "Errors in Send_Vendor()"        
                return False
            else:
                return True
              
        else:
            print "Send_Vendor failed. Dropdown box not detected." 
            return False

    def Send_Amount(self): 
        print "Entering amount: %s" % self.amount      
        Auto.send(self.amount)
        time.sleep(sleep)
        if self.Type == "debit":
            Auto.send("{TAB}") # Now at "Account" textbox.                
        elif self.Type == "credit":
            Auto.send("{TAB 3}")# Now at "Account" textbox.      
        else:
            print "Send_Amount failed. Amount == 0"          
            return False  
        if errors_exist():
             return False
        else:
             return True

    def Send_Account(self):
        if self.Type == "debit":
            self.account = "income"
            Auto.send(self.account)
            Auto.send("{TAB 2}")
            if request_confirmation:
                if entered_y():
                    print "entered y"
                    Auto.WinActivate(apptitle)
                    print "shouldnt see y et"
                    Auto.send("{ENTER}")
                else:
                    Auto.WinActivate(apptitle)
                    return False
            else:
                Auto.send("{ENTER}")
                                         
            if errors_exist():          
                return False
            else: 
                return True     
        elif self.Type == "credit":
            copy_account(self)
            paste_account()
            print "account entered for deposit in attempt_send_account(Type)"
            if errors_exist():
                return False
            else:
                return True
        else:
            print "Function 'attempt_send_account' failed."
            return False 

    def copy_account(self):
        # Start in vendor textbox for transaction.
        Auto.send("!g")
        Auto.send("!s") # Now highlighted cursor is in "Search for: "
        partially_type(self.vendor,n)
        Auto.send("{TAB}")
        time.sleep(sleep)
        if errors_exist():
            print "Credit's Account name not found when copying"           
            return False
            if Auto.WinExists("Recording Transaction"):
                auto.send("n")
                return 0

    def paste_account():
        #ctrl+v
        pass
   
    def Transaction_Entry(self):
       if (self.Send_Date() and self.Send_Vendor() and self.Send_Amount() and self.Send_Account()):
           print "Success"
           return True        
       else:
           print "Transaction_Entry Failure"
           Skipped_List.append(self)
           close_all_windows()
           open_register(bank_code)
           return False
                          
def Process(statement):
    setup() # Ends at "Date" textbox.  
    with open(statement) as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',') 
        for transaction in readCSV:
            Current_Transaction = Transaction(transaction) #if i could end this function here thatd be great
            Current_Transaction.Determine_Type()
            Current_Transaction.Transaction_Entry()                                 
            Transaction_List.append(Current_Transaction)                               
            #print "Finished Transaction number: %s" % counter 
            print "______________________________"            
                
    print "Processed all transactions at: %s" % current_time
         

n = 7 # length for partially_type()         
Auto = Dispatch("AutoItX3.Control")
current_time = time.strftime("%H:%M:%S")
green = 0x4E9E19
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
do_credits = True
request_confirmation = True
# ^ do for debits and credits separately
sleep = 1
#### Settings ####

Process(statement)
