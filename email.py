import tkinter

import pythoncom
import win32com.client
import os

from datetime import datetime, timedelta

from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog

    # TKinter UI
root = Tk()
root.geometry('300x150')
root.resizable(False, False)
root.title('Attachment Downloader')
root.iconbitmap("Mail10.ico")

    # Global VARs
global enableLoop
enableLoop = IntVar()

    # Outlook Access
outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")

inbox = mapi.GetDefaultFolder(6)

    #Backup Management
with open('downloadlocal.txt', 'w+') as f:
    outputDir = f.readline()

    #Output Folder
global folderPath

if(outputDir == ""):
    outputDir = os.getcwd() + "\\attach"

    with open('downloadlocal.txt', 'w+') as f:
        f.write(outputDir)

if not os.path.exists(outputDir):
    os.makedirs(outputDir)
folderPath = outputDir

    # Filter Email
global received_dt

def DefineMessages():
    received_dt = (datetime.now() - timedelta(days = 1)).strftime('%d/%m/%Y %H:%M %p')
    global messages
    messages = inbox.Items
    messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")

    #  Options - Select Folder
def FindFolder():
    saveLocationEntry["state"] = NORMAL
    saveLocationEntry.delete(0, tkinter.END)
    folderPath = filedialog.askdirectory()
    outputDir = folderPath
    saveLocationEntry.insert(0, folderPath)
    saveLocationEntry["state"] = DISABLED

    with open('downloadlocal.txt', 'w+') as f:
        f.write(outputDir)

    # Manage Downloads
def DownloadAttachment():
    try:
        for message in list(messages):
            if message.unread == True:
                try:
                    with open('downloadlocal.txt', 'r') as f:
                        outputDir = f.readline()

                    for attachment in message.Attachments:
                        attachment.SaveASFile(os.path.join(outputDir, attachment.FileName))

                    message.unread = False

                except Exception as e:
                    messagebox.showerror(title = "Error", message = "Error when saving the attachment: " + str(e))

    except Exception as e:
        messagebox.showerror(title = "Error", message = "Error when processing emails messages: " + str(e))

    DefineMessages()

    #Menu Management
def LoadMenu():
    ForgetAll()

    btn_generate.pack(side = TOP, fill = "both")

    if enableLoop.get() == 1:
        btn_generate["state"] = DISABLED
        tl_autoDownload.pack()
        root.after((5 * 1000), DownloadAttachment)
    else:
        btn_generate["state"] = NORMAL
        tl_autoDownload.pack_forget()

    btn_options.pack(side = TOP, fill = "both")

    #Options Menu
def LoadOptions():
    ForgetAll()

    global win
    win = Toplevel(root)

    win.geometry('300x250')
    win.resizable(False, False)
    win.title('Options')

    win.iconbitmap("Note01.ico")

    # Altera o que acontece ao clicar no X da janela
    win.protocol("WM_DELETE_WINDOW", ReturnToMenu)

    Label(win, text = "Options", font = "italic 15 bold").pack(pady = 10)

    btn_return = ttk.Button(win, text = "Retornar ao Início", command = ReturnToMenu)
    btn_file = ttk.Button(win, text = "Selecionar Pasta", command = FindFolder)
    chk_enableLoop = Checkbutton(win, text = "Baixar Automaticamente", variable = enableLoop, onvalue = 1, offvalue = 0)
    global tl_saveLocation
    tl_saveLocation = Label(win, text = "Escolha um Local para Salvar:", font = "italic 8")
    tl_emptyLineOne = Label(win, text = "", font = "italic 10")
    tl_emptyLineTwo = Label(win, text = "", font = "italic 10")

    global saveLocationEntry
    saveLocationEntry = Entry(win, width=60)
    saveLocationEntry.insert(0, outputDir)

    chk_enableLoop.pack()
    tl_emptyLineOne.pack()
    tl_saveLocation.pack()
    saveLocationEntry.pack()
    saveLocationEntry["state"] = DISABLED
    btn_file.pack(fill = BOTH)
    tl_emptyLineTwo.pack()
    btn_return.pack(side = TOP, fill = "both")

def ReturnToMenu():
    LoadMenu()
    win.destroy()

def ForgetAll():
    tl_emptyLine.pack_forget()
    tl_autoDownload.pack_forget()

    btn_options.pack_forget()
    btn_generate.pack_forget()

    # TKinter UI - Global Buttons
Label(root, text = "Attachment Downloader", font = "italic 15 bold").pack(pady = 10)

btn_generate = ttk.Button(root, text = "Baixar Anexos", command = DownloadAttachment)
btn_options = ttk.Button(root, text = "Abrir Opções", command = LoadOptions)

tl_emptyLine = Label(root, text = "", font = "italic 10")
tl_autoDownload = Label(root, text = "Download Automático Ativo...", font = "italic 8")

    # END OPERATIONS
LoadMenu()
DefineMessages()

root.mainloop()