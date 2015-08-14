import win32api

class Checkpoint(object):
    name = ""
    #square or rectangle ABCD. 
    a_coords = (0,0)
    b_coords = (0,0)
    c_coords = (0,0)
    d_coords = (0,0)
    mid_xaxis = 0
    mid_yaxis = 0
    midpoint = (0,0)
    #picture = checkpoint.jpg
    
def make_checkpoint(name,a_coords,b_coords,c_coords,d_coords,mid_xaxis,mid_yaxis,midpoint):
    checkpoint = Checkpoint()
    checkpoint.name = name
    checkpoint.a_coords = a_coords
    checkpoint.b_coords = b_coords
    checkpoint.c_coords = c_coords
    checkpoint.d_coords = d_coords
    checkpoint.mid_xaxis = mid_xaxis # eventually make this a formula of ((a[0]+b[0])/2)
    checkpoint.mid_yaxis = mid_yaxis
    checkpoint.midpoint = midpoint # eventually make this a formula of mid x/y
    return checkpoint

Deposit_To_Textbox = make_checkpoint("Deposit_To",
                             (78,169),(152,169),
                             (78,186),(152,186),115,178,(115,178))



import win32api
#x, y = win32api.GetCursorPos()

def move_to(coordinates):
    win32api.SetCursorPos(coordinates)

move_to(Deposit_To_Textbox.midpoint)

#print Deposit_To_Textbox.d_coords

