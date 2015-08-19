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
    Auto.MouseClick("left", 34, 77)

def quit_current_transaction(): # ends with date highlighted
    Auto.WinActivate("Yuliya")
    Auto.send("{CTRLDOWN}")
    Auto.send("{d}")
    Auto.send("{CTRLUP}")   
    if Auto.WinExists("Past Transactions"):
        Auto.send("{TAB}")
        Auto.send("{ENTER}")
        print "Closed 'Past Transactions'"
    print "Now_at_date"

##def select_bank(bankcode):
##    #Auto.WinActivate(apptitle)
##    BOA = 10030   # "Bank of America Business - Operating"
##    AmEx = 10400
##    Auto.send("{TAB}") # will eventually evolve into select_bank function
##    Auto.send(bankcode) #^
##    Auto.send("{ENTER}")
    
def setup():
    Auto.WinActivate(apptitle)
    Auto.WinMove(apptitle,"", 0, 0, 1000, 1000)
    close_all_windows()
    open_home()
    time.sleep(1)
    Auto.send("{TAB}") # will eventually evolve into select_bank function
    Auto.send(bankcode) #^
    Auto.send("{ENTER}") #^
    #time.sleep(1)
    #right now leaves me with date selected in BOA

def check_if_black(x,y):
    #if blue
    #coord = (191,412)
    color = 0x000000               # 0xFFFFFF         #(x,y#Auto.PixelGetColor(x,y #black
    if Auto.PixelGetColor(x,y) == color: 
        #print "TRUE!!!!!!!!!!!!black!!!!!!!!!!!!!"
        return True     
    else:
        print "No suggestions detected." 
        return False

def check_if_grey(x,y):
    color = 0xABABAB             
    if Auto.PixelGetColor(x,y) == color: 
        print "Attempting to restart setup..."
        return True      
    else:
        #print "false not grey" 
        return False
    
def send_vendor_deposit(v,a): # type 4 letters. if checkifblack == true: send. else: append skipped
    counter = 0
    for letter in v[0:3]:
        Auto.send(letter)
        counter += 1
        #time.sleep(.5)
        #print counter
    if check_if_black(300,450) == True: 
        Auto.send("{TAB}") # now in payment after one tab
        print "Sent Vendor: "
        Auto.send("{TAB 2}") #2 for deposit 0 for payment
        Auto.send(a) #Amount
        Auto.send("{TAB}")
        Auto.send("income") #Income. will need to be xpanded for credits to include lookup
        Auto.send("{TAB 2}")
        #time.sleep(2)
        Auto.send("{ENTER}") # this en
        return True
        #time.sleep(1)
    else:
        Skipped_List.append(v)
        #print ("Skipped list: " , Skipped_List)
        print "skipped a vendor"
        Auto.send("{ESC}")
        Auto.Send("+{TAB 2}") # now is at Date as if nothing ever happened.      
        return False
     
def DepositEntry(d,v,a): # Cycles for every transaction in statement if a > 0
    Auto.send(d) #Date
    Auto.send("{TAB 2}")
    send_vendor_deposit(v,a)

##def find_account(vendor): should be here
##     = vendor
##    Auto.send("!e")
##    #return acc
##    return True


def send_vendor_credit(v,a): # type 4 letters. if checkifblack == true: send. else: append skipped
    counter = 0
    account = find_account(v) ####################
    for letter in v[0:3]:
        Auto.send(letter)
        counter += 1
        #time.sleep(.5)
        print counter
    if check_if_black(300,450) == True: 
        Auto.send("{TAB}") # now in payment after one tab
        print "Sent Vendor: "
        Auto.send("{TAB 2}") #2 for deposit 0 for payment
        Auto.send(a) #Amount
        Auto.send("{TAB}")
        Auto.send(account) #Income. will need to be xpanded for credits to include lookup
        Auto.send("{TAB 2}")
        #time.sleep(2)
        Auto.send("{ENTER}") # this en
        return True
        #time.sleep(1)
    else:
        Skipped_List.append(v)
        #print ("Skipped list: " , Skipped_List)
        print "skipped a vendor"
        Auto.send("{ESC}")
        Auto.Send("+{TAB 2}") # now is at Date as if nothing ever happened.      
        return False
    
def CreditEntry(d,v,a): # Cycles for every transaction in statement if a > 0
    Auto.send(d) #Date
    Auto.send("{TAB 2}")
    send_vendor_credit(v,a)   

def Record(statement):
    with open(statement) as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',')
        #Auto.send("{ENTER}") ####FIX THIS
        #time.sleep(2)
        
        for transaction in readCSV:
            date = transaction[0]
            vendor = transaction[1]
            amount = transaction[2]
            
            if float(amount) > 0:
                DepositEntry(date,vendor,amount)
                print ("Attempting to debit:  %s to [eventual location]: " % transaction)
                
            elif float(amount) < 0:
                CreditEntry(date,vendor,amount)
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
        setup() #setup should leave you with date highlighted
    
    Record(statement)

    counter = 0
    for transaction in Skipped_List:
        print "skipped %s" % transaction 
        counter +=1 
    
        #print transaction

def find_account(vendor):
    Auto.send("!e")
    Auto.send("f")
    openFile
    #some other stuff to go into advanced and export

def openFile():
    pass
    #return acc
    


Skipped_List = []

apptitle = "Yuliya"
##statement = "stmt2.txt"
statement = "stmtsampleclean.txt"
Auto = Dispatch("AutoItX3.Control")
bankcode = "10030"
current_time = time.strftime("%H:%M:%S")

Process()


##Auto.WinActivate(apptitle)
##answer =  Auto.WinGetState(apptitle)
##print answer
