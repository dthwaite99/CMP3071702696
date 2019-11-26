from PyQt5.QtCore import QDateTime, Qt, QTimer
from PyQt5.QtWidgets import (QApplication, QCheckBox, QComboBox, QDateTimeEdit,
        QDial, QDialog, QGridLayout, QGroupBox, QHBoxLayout, QLabel, QLineEdit,
        QProgressBar, QPushButton, QRadioButton, QScrollBar, QSizePolicy,
        QSlider, QSpinBox, QStyleFactory, QTableWidget, QTabWidget, QTextEdit,
        QVBoxLayout, QWidget, QMessageBox)
from socket import AF_INET, socket, SOCK_STREAM
from threading import Thread
from simplecrypt import encrypt, decrypt
import pickle
import openpyxl
import time
import sys
import copy



HEADERSIZE = 10 #headersize for transmissions
password = 'password' # decryption password
unreadMessList = [] #list of all messages used in the changeView function

# GUI:
app = QApplication([])

text_area = QTextEdit()
text_area.setFocusPolicy(Qt.NoFocus)

unread_mess = QTextEdit()
unread_mess.setFocusPolicy(Qt.NoFocus)

message = QLineEdit()

styleComboBox = QComboBox()

button = QPushButton("Important")
button.setCheckable(True)

layout = QVBoxLayout()
layout.addWidget(styleComboBox)
layout.addWidget(button)
layout.addWidget(text_area)
layout.addWidget(unread_mess)
layout.addWidget(message)

window = QWidget()
window.setLayout(layout)
window.show()

#Event handlers:
def send_message():  #function for sending message
    global activeChat#active chat is who the message is going to
    global user#user is where the message is coming from
    d = {}#messages are sent as pickled dictionaries
    d[0] = message.text()#message text
    d[1] = activeChat#recipient
    d[2] = user#sender
    if button.isChecked(): #checkkks whether or not message is important is makrs message accordingly
        d[3] = "Y"#y for yes
        text_area.append(str("You :" + message.text()+" IMPORTANT"))
    else:
        d[3] = "N" # n for no
        text_area.append(str("You :" + message.text()+" NON IMPORTANT"))

    writeDic = copy.deepcopy(d) #write unencrypted message to the list for speed
    unreadMessList.append(writeDic)#pickle the dictionary
    d[0] = encrypt(password, message.text()) #encrypt message text    
    msg = pickle.dumps(d)#pickle the dictionary
    msg = bytes(f"{len(msg):<{HEADERSIZE}}", 'utf-8')+msg#make message into bytes
    client_socket.send(msg)
    message.clear()

def receive():
    global user
    global unreadMessList
    full_msg = b'' #variables for LAN
    new_msg = True
    while True:
        msg = client_socket.recv(BUFSIZ) #waiting to recieve a transmission
        if new_msg:
            msglen = int(msg[:HEADERSIZE])
            new_msg = False


        full_msg += msg


        if len(full_msg)-HEADERSIZE == msglen:
            x = pickle.loads(full_msg[HEADERSIZE:])
            if x[1] == user: #if recipient is user then decypt and display message if not then do not process it
                display = decrypt(password, x[0]) #decrypts message
                if x[2] == activeChat: # if message sender is the active chat
                    if x[3] == "Y":#display code for if message is important
                        text_area.append(str("Them :" +display.decode("utf-8")+" IMPORTANT"))
                    else:#if message unimporant then display message using this code
                        text_area.append(str("Them :" +display.decode("utf-8")+" NON IMPORTANT"))
                else:
                    if x[3] == "Y":
                        unread_mess.append("You missed an important message from " + x[2])
                    else:
                        unread_mess.append("You missed a message from " + x[2])
                x[0] = display.decode("utf-8")
                unreadMessList.append(x) #writes decoded message into message list
        new_msg = True
        full_msg = b""
          
             
    
def checkUnreadMessage(name): #iterates through the message list and displays appropriate messages
    global unreadMessList
    for x in unreadMessList:
        if x[2] == name:
            display = x[0] 
            if x[3] == "Y":
                text_area.append(str("Them :" +display+" IMPORTANT")) # stupid w
            else:
                text_area.append(str("Them :" +display+" NON IMPORTANT"))
        if x[2] == user and x[1] == name: #if the user sent a message to the name variable then display
            display = x[1] 
            if x[3] == "Y":
                text_area.append(str("You :" +display+" IMPORTANT")) # stupid w
            else:
                text_area.append(str("You :" +display+" NON IMPORTANT"))
        
        
def changeView():  #tells you who yoiu are chatting with and displays all previous messages
    global activeChat
    text_area.clear()#clears chatbox
    text_area.append("You are now chatting with " + str(nameArray[styleComboBox.currentIndex()]))
    activeChat = nameArray[styleComboBox.currentIndex()]
    checkUnreadMessage(activeChat) #calls function with the activechat as an argument
    

            
# Signals:
message.returnPressed.connect(send_message) #when you hit enter message is sent
styleComboBox.activated.connect(changeView) #when you change the dropdown box the chat gets updated


HOST = "127.0.0.1" #IP of server
PORT = 33000 #Port used by server

nameArray = [] #array of all contacts used for sign in and chats

try:
    wb1 = openpyxl.load_workbook('contacts.xlsx')
    ws1 = wb1.get_sheet_by_name('Sheet1')
    for cell in ws1['A']:
        nameArray.append(str(cell.value)) #loads names from contacts sheet
    print("Contacts loaded properly")
except:
    print("Your contacts File is Missing So We have Provided you with a defualt one Please Contact an Administrator")
    nameArray.append("jeffA@dogefin.jp")
    nameArray.append("jeffB@dogefin.jp")
    nameArray.append("jeffC@dogefin.jp")


user = input('Enter username: ') #only accepts users if they enter a username from the excell file 
if user in nameArray: #if entered username dosn't exist then reject log in
    print("You have logged in as '" + user + "' the program will now launch")
    nameArray.remove(user)
else:
    print("No match the username you entered either doesn't exist or you need an updated contacts file, please contact an administrator")
    time.sleep(5)
    sys.exit()
    
styleComboBox.addItems(nameArray) #adds all names bar the users to the chat selector so that they cannot message themselves
activeChat = nameArray[0]
if not PORT:
    PORT = 33000
else:
    PORT = int(PORT)

BUFSIZ = 1024 #buffersize
ADDR = (HOST, PORT) #address of server

client_socket = socket(AF_INET, SOCK_STREAM) #creates socket
client_socket.connect(ADDR) #binds address to socket

receive_thread = Thread(target=receive) #starts thread so that it can always recieve messages no matter what
receive_thread.start() #starts thread

#time.sleep(5)
app.exec_() #starts GUI
