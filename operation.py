#class def
class Operations:
    def __init__(self, name, type,amount,eglise):
        self.name = name
        self.type = type
        self.amount = amount
        self.eglise = eglise

    # getter method
    def get_name(self):
        return self.name
      
    # setter method
    def set_name(self, x):
        self.name = x

    # getter method
    def get_type(self):
        return self.type
      
    # setter method
    def set_type(self, x):
        self.type = x

    def get_amount(self):
        return self.amount
      
    # setter method
    def set_amount(self, x):
        self.amount = x
 