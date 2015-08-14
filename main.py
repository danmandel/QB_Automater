import csv
import SendKeys
import win32api
import win32com.client
import win32con
import time

def move_to(coordinates):
    win32api.SetCursorPos(coordinates)
    
def click(coordinates):
    win32api.SetCursorPos((coordinates))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,coordinates[0],coordinates[1],0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,coordinates[0],coordinates[1],0,0)

def send(text):
    win32com.client.Dispatch("WScript.Shell").SendKeys(text)


def dataentry(d,v,a):
    #first thing to add is to select QB screen with a mouseclick
    #add check if you're in the deposits/credits screen
    
    #Date_Textbox
    click(Date_Textbox.midpoint)
    click(Date_Textbox.midpoint)
    click(Date_Textbox.midpoint)
    send(d)
    time.sleep(1)
    

    #Received_From_Textbox
    click(Received_From_Textbox.midpoint)

    send(v)
    time.sleep(1)

    #From_Account_Textbox
    click(From_Account_Textbox.midpoint)
    send("income")
    time.sleep(1)
    
    #Amount_Textbox
    click(Amount_Textbox.midpoint)
    send(a)
    time.sleep(1)
    send("{ENTER}")

    
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
    #picture = checkpoint.jpg
    
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


def DVA(statement):
    with open(statement) as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',')
        
        for transaction in readCSV:
            date = transaction[0]
            vendor = transaction[1]
            amount = transaction[2]        

            dataentry(date,vendor,amount)
            print(date,vendor,amount)
            print "  "
            
def get_cursor_location():
    x, y = win32api.GetCursorPos()
    print x,y

    
statement = "stmt2.txt"    

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
                             (639,269),(759,269))

#get_cursor_location()


DVA(statement)    
