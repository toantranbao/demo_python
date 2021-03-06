class Car:
    color = ""
    brand = ""

    def di_chuyen(self):
        print("Xe co the di chuyen")

    #Ham tao co tham so
    def __init__(self,b="",c=""):
        print("The object has been initialized...!")
        self.color = c
        self.brand = b

#Adding one more row
