import win32api

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

def move_to(coordinates):
    win32api.SetCursorPos(coordinates)

Deposit_To_Textbox = make_checkpoint("Deposit_To",
                             (78,169),(152,169),
                             (78,186),(152,186)) 

Date_Textbox = make_checkpoint("Date",
                             (214,169),(278,169),
                             (214,186),(278,186))

Received_From_Textbox = make_checkpoint("Received_From",
                             (17,254),(142,254),
                             (17,270),(142,270))

From_Account_Textbox = make_checkpoint("From_Account",
                             (160,254),(278,254),
                             (160,269),(278,269))

Amount_Textbox = make_checkpoint("Amount",
                             (639,254),(759,269),
                             (639,269),(759,269))
