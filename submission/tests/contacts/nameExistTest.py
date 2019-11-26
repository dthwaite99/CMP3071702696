import openpyxl

nameArray = []
try:
    wb1 = openpyxl.load_workbook('contacts.xlsx')
    ws1 = wb1.get_sheet_by_name('Sheet1')
except:
    print("Your contacts File is Missing Please Contact an Administrator")
    
    
for cell in ws1['A']:
    nameArray.append(str(cell.value))
    
print(nameArray)
x = input("Enter your user name:")
if x in nameArray:
    print("You have logged in as '" + x + "' the program will now launch")
else:
    print("No match the username you entered either doesn't exist or you need an updated contacts file, please contact an administrator")
