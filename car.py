class Car:
    color = ""
    brand = ""


    def di_chuyen(self):
        print("Xe co the di chuyen")

    #Ham tao co tham so
    def __init__(self,brand="",color="", max_speed:int = 0):
        print("The object has been initialized...!")
        self.color = color
        self.brand = brand
        self.max_speed = max_speed

#Adding one more row
#phattatsuken