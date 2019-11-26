import winsound

def makeNoise():
    winsound.Beep(2000, 500)
    winsound.Beep(2500, 500)
    winsound.Beep(2000, 500)
    winsound.Beep(2500, 500)
    
a = str(input("enter 1 to make a noise or 2 make no noise: "))
if a == "1":
    makeNoise()
elif a == "2":
    print("no noise")

else:
    print("Incorrect Input please enter a 1 or 2")