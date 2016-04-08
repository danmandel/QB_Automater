import csv
from win32com.client import Dispatch
import time
import datetime

def close_all_windows():
    """Starts anywhere. Ends with blank screen."""
    Auto.WinActivate(apptitle)
    Auto.send("{ESC}") # Should this be a SendEsc function?
    if Auto.WinExists("Recording"):
        print "Closing recording message."
        Auto.send("n")     
    Auto.send("!w") # Alt+w selects Window dropdown menu.
    Auto.send("a") # A selects the "Close All" option.
       
    time.sleep(1)
    for x in range(5): # Checks for any non-grey windows.
        if not is_color(250,250,grey):
            Auto.send("{ESC}")
            
    Auto.Send("{ENTER 2}")
    #time.sleep(1)
    if Auto.WinExists("Past Transactions"):
        time.sleep(2)
        print "'Past Transactions' warning message exists."
        Auto.send("n") 
    for x in range(5):
        if not is_color(250,250,grey):
            Auto.send("{ESC}")
               
def check_exists():
    if is_color(35,480,blackish_grey):
        print "'1-Line' box is currently checked."
        return True
    else:
        return False
    
def partially_type(text,n):
    # print "Entering first %s letters of %s" % (n, text)
    for letter in text[0:n]:
        Auto.send(letter)
        time.sleep(.1)
    
def open_register(bank_code):
    """ Should be started at a blank screen.
        Ends at bank register with "Date" textbox highlighted."""
    Auto.send("^r") # "ctrl+r" opens register.
    Auto.send(bank_code) # Types in bank_code.
    Auto.send("{ENTER}") # Brings up register.
      
def setup():
     Auto.WinActivate(apptitle)
     Auto.WinMove(apptitle,"", 0, 0, 900, 700) # 0,0,x,y
     close_all_windows()
     time.sleep(sleep)
     open_register(bank_code)
     tile_windows()
     if check_exists():
         Auto.send("!1") # "Alt + 1" to untick 1_line box.
    # Ends at "Date" textbox.
    
def entered_y():
    #Auto.WinActivate("C:\Python27")
    Auto.WinMove("*Python 2.7.11 Shell*","", 901, 200, 460, 500)
    Auto.Send("!{TAB}")
    time.sleep(sleep)
    test_var = raw_input("Enter this transaction? y/n: ")
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
    """test docstring"""
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
        if self.Type == "credit":
            print "starting copy_account"
            self.copy_account()
            time.sleep(sleep)
        print "date %s" % self.date 
        Auto.send("!d") # Moves cursor to "Date" textbox.
        #print "Entering date: %s" % self.date
        time.sleep(3)
        Auto.send(self.date)  
         
        Auto.send("{TAB 2}") # Moves cursor to "Payee" textbox
        time.sleep(sleep)
        if errors_exist(): # Need to restart transaction now. 
            print "errors in date"           
            return False        
        else:
            return True
               
    def Send_Vendor(self):     
        # Starts at "Payee" textbox.
        print "Entering vendor: %s " % self.vendor
        time.sleep(sleep)
        partially_type(self.vendor,n) # n letters
        time.sleep(sleep)
        Auto.send("{TAB}") # Now highlighted cursor is in "Payment" textbox.
        if errors_exist():
            print "Send_Vendor() failed. Dropdown box not detected." 
            return False
        else:
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
           
    def Send_Amount(self): 
        #print "Entering amount: %s" % self.amount      
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
            Auto.send("{TAB}")
            if request_confirmation:
                if entered_y():               
                    Auto.WinActivate(apptitle)                          
                    Auto.send("{TAB}")
                    if errors_exist():          
                        return False
                    else: 
                        return True                    
                else:                  
                    Auto.WinActivate(apptitle)
                    Auto.send("n")
                    time.sleep(sleep)
                    close_all_windows()
                    return False
            else:
                Auto.send("{TAB}")                          
                if errors_exist():          
                    return False
                else: 
                    return True     
        elif self.Type == "credit":
            self.paste_account()
            print "account entered for deposit in attempt_send_account(Type)"
            if errors_exist():
                return False
            else:
                return True
        else:
            print "Send_Account() failed."
            return False 

    def copy_account(self):
        # Start in vendor textbox for transaction. or date tb?
        time.sleep(sleep)
        Auto.send("!g")
        Auto.send("!s") # Now highlighted cursor is in "Search for: "
        time.sleep(sleep)
        partially_type(self.vendor,n)
        Auto.send("{TAB}")
        time.sleep(sleep)
        if Auto.WinExists("Name Not Found"):
            print "Credit's account name not found when copying"
            Auto.send("{ESC 3}")
            time.sleep(sleep)
            return False
        time.sleep(sleep)
      
        Auto.send("!k")
        time.sleep(sleep)
        Auto.WinActivate(apptitle) #not sure why this seems to be necessary
        if Auto.WinExists("Warning"):
            print "warning exists"
            Auto.send("{ENTER}")
            Auto.send("{ESC 2}")
            return False
        else:
            time.sleep(sleep)
            Auto.send("{ESC}")
            time.sleep(sleep)
            Auto.send("{TAB 4}")
           
            time.sleep(sleep)
            Auto.send("{c}")
            time.sleep(3)
            Auto.send("{ESC}")#
            time.sleep(sleep)
            Auto.send("n")#
            time.sleep(sleep)
            open_register(bank_code)
            time.sleep(sleep)
        
        

        '''#messag
        alt d
        continue date'''

    def paste_account(self):
        Auto.send("{^v}")
               
    def Transaction_Entry(self):

        if self.Type == "debit" and do_debits:
            if (self.Send_Date() and self.Send_Vendor() and self.Send_Amount() and self.Send_Account()):
               return True
            else:
                print "Skipping: %s" % self.amount
                Skipped_List.append(self)
                time.sleep(sleep)
                close_all_windows()
                open_register(bank_code)
                return False                
        elif self.Type == "credit" and do_credits:
            if (self.Send_Date() and self.Send_Vendor() and self.Send_Amount() and self.Send_Account()):
                return True
            else:
                print "Skipping: %s" % self.amount
                Skipped_List.append(self)
                time.sleep(sleep)
                close_all_windows()
                open_register(bank_code)
                return False
        else:
            print "Skipping: %s" % self.amount
            Skipped_List.append(self)
            time.sleep(sleep)
            close_all_windows()
            time.sleep(1)
            open_register(bank_code)
            return False
        
   
def Process(statement):
    setup() # Ends at "Date" textbox.  
    with open(statement) as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',') 
        for transaction in readCSV:
            Auto.WinActivate(apptitle)
            Current_Transaction = Transaction(transaction) #if i could end this function here thatd be great
            Current_Transaction.Determine_Type()
            Current_Transaction.Transaction_Entry()                                 
            #Transaction_List.append(Current_Transaction)                               
            #print "Finished Transaction number: %s" % counter 
            time.sleep(sleep)
            print "______________________________"            
                
    print "Processed all transactions at: %s" % current_time

def Process1(statement):
    setup()
    readCSV = open(statement)
         

n = 6 # length for partially_type()         
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
   

 
##### SETTINGS #####
apptitle = "Yuliya"
statement = "C:\Python27\Scripts\QB\credit_test.txt"
bank_code = "Bank of America Bus"
do_debits = False
do_credits = True
request_confirmation = True
sleep = 1 #seconds
#option to remove 30 sec
##### SETTINGS #####

Process(statement)
