import datetime as date
from datetime import datetime


print("Store")
today = date.date.today()
x = today.strftime('%Y-%m-%d')
print("date_string =", x)
print("type of date_string =", type(x))

print("Store")
today = date.date.today()
x = today.strftime('%Y-%m-%d')
print("date_string =", x)
print("type of date_string =", type(x))

now = datetime.today()
time_string = now.strftime("%H:%M")
print("time_string =", time_string)
print("type of time_string =", type(time_string))



print("Recover")
date_object = datetime.strptime(x, "%Y-%m-%d")
print("date_string =", date_object)
print("type of date_string =", type(date_object))

print("test")
sixmonths = (date.date.today() - date.timedelta(182))
if date_object.date() >= sixmonths:
  print(date_object.date(), " was after ", sixmonths)
else:
  print(date_object.date(), " was before ", sixmonths)
print("")

