conf = {
    "colors": {
        "black" : 0x000000,
        "blackish_grey" : 0x484848,
        "blue" : 0x3399FF,
        "green" : 0x4E9E,
        "grey" : 0xABABAB,
        "white" : 0xFFFFFF
        },
    
    "transaction_types": {
        "do_debits" : True,
        "do_credit" : False
        },

    "misc": {
        "apptitle" : "Yuliya",
        "statement" : "C:\Python27\Scripts\QB\credit_test.txt",
        "bank_code" : "Bank of America Bus",
        "request_confirmation" : True,
        "sleep" : 1,
        },
    
    "methods": {
        "print_x" : print ("x")
        },
    
    "messages": {
        "error_1" : "error",
        }

    
}

class Config(object):
    def __init__(self):
        self.sleep = 0
        self.config = conf

    def get_property(self, property_name):
        if property_name not in self.config.keys():
            return None # Avoids KeyError.
        return self.config[property_name]
