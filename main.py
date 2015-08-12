import csv


with open('stmt1.txt') as csvfile:
    readCSV = csv.reader(csvfile, delimiter=',')

    dates = []
    vendors = []
    amounts = []
    
    for transaction in readCSV:
        date = transaction[0]
        vendor = transaction[1]
        amount = transaction[2]

        print (date, vendor, amount)

        dates.append(date)
        vendors.append(vendor)
        amounts.append(amount)
        
##    print ("dates: ", dates)
##    print "             "
##    print ("vendors: ", vendors)
##    print "             "
##    print ("amounts: ", amounts)
        

'''def click(x,y):
    win32api.SetCursorPos((x,y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)
click(10,10)'''

#
