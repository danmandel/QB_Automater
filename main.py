import win32api, win32
import numpy as np
#first you must get a sample csv loaded into ListOfTransactions



ListOfTransactions = []

date,description, amount = np.loadtxt('stmt.csv',
                                      unpack=True,
                                      delimiter=',')
                                      



'''def click(x,y):
    win32api.SetCursorPos((x,y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)
click(10,10)'''

#
