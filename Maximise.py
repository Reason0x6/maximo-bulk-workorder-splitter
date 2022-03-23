# Prototype UI for Maximo Workorder Export Splitter
# Designed by github\reason0x6

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter.constants import HORIZONTAL, LEFT
from time import sleep
import os
import threading
from PyPDF2 import PdfFileReader, PdfFileWriter
import asyncio
import concurrent.futures
import win32api
import subprocess
from datetime import date
import win32print
global keepvalue
global logfile_name

class runner(threading.Thread):

    def __init__(self, function_that_downloads):
        threading.Thread.__init__(self)
        self.runnable = function_that_downloads
        self.daemon = True

    def run(self):
        self.runnable()
        


#printer def
currentprinter = win32print.GetDefaultPrinter()

if not os.path.isdir('C:\Support\MaximoOutput'):
    os.mkdir('C:\Support\MaximoOutput')

if not os.path.isdir('C:\Support\MaximoSplitLogs'):
    os.mkdir('C:\Support\MaximoSplitLogs')

logfile_name = 'C:\Support\MaximoSplitLogs\{}.log'.format(date.today())

if not os.path.isfile(logfile_name):
    fp = open(logfile_name, 'x')
    fp.close()

files = []
defaultFile = "C:\\temp\\printjobspoolernew.pdf"
def title_page_find(text):
    if "Location" in text and \
    "Asset" in text and \
    "Job Plan" in text and \
    "WORK ORDER" in text and \
    "Reported By" in text:
        return True

def pdf_splitter(path):
    fname = os.path.splitext(os.path.basename(path))[0]
    printed = 1

    pdf = PdfFileReader(path)
    page = -1

    while page < pdf.getNumPages() -1:
        page += 1
        test = title_page_find(pdf.getPage(page).extractText())
        
        if test:
            output_filename = 'C:\Support\MaximoOutput\{}_CoverPage_{}.pdf'.format(fname, printed)
            
            
            pdf_writer = PdfFileWriter()
            pdf_writer.addPage(pdf.getPage(page))
            files.append(output_filename)
            output_filename
            with open(output_filename, 'wb') as out:
                pdf_writer.write(out)
            printed += 1

        elif len(pdf.getPage(page).extractText()) == 8:
            continue

        else:  
            pdf_writer.addPage(pdf.getPage(page))
            with open(output_filename, 'ab') as out:
                pdf_writer.write(out)

        
    if drop.get() == 'Print Work Orders':
        title = ttk.Label(widgets_frame, justify=LEFT, text="Printing Work Orders (This may take some time)")
        title.grid(row=0, column=0,  padx=5, pady=(0, 20), sticky="ew")

        for file in files:
            try:
                GHOSTSCRIPT_PATH = "_gscript\\bin\\gswin32.exe"
                GSPRINT_PATH = "_gsprint\\gsprint.exe"

                params = '-ghostscript "'+ GHOSTSCRIPT_PATH  +'" -printer "'+currentprinter+'" -all "'+file+'"'
                print(params) #for debugging in console 

                win32api.ShellExecute(0, 'open', GSPRINT_PATH, params, '.',0)
                
               
            except Exception as e: 
                out = open(logfile_name, 'a')
                out.write(str(e))
                out.close()

            sleep(0.02)

    else:
        os.startfile(os.path.realpath("C:\Support\MaximoOutput"))
        title.grid_forget()
        title = ttk.Label(widgets_frame, justify=LEFT, text="Error When Printing", foreground='red')
        title.grid(row=0, column=0,  padx=5, pady=(0, 20), sticky="ew")
    
    sleep(2)
    title.grid_forget()
    title = ttk.Label(widgets_frame, justify=LEFT, text="Printing Finish Sending", foreground='green')
    title.grid(row=0, column=0,  padx=5, pady=(0, 20), sticky="ew")
    
    
def openFIles():
    os.startfile(os.path.realpath("C:\Support\MaximoOutput"))

def file_spliter():
     global title
     global widgets_frame
     keepvalue = drop.get()
     try:
        title = ttk.Label(widgets_frame, justify=LEFT, text="Pages Proccessing...")
        title.grid(row=0, column=0,  padx=5, pady=(0, 20), sticky="ew")

        button.grid_forget()
        drop.grid_forget()
        title2.grid_forget()
        widgets_frame.grid_forget()
        
        widgets_frame = ttk.Frame(root, padding=(0, 0, 0, 10))
        widgets_frame.grid(row=0, column=1, padx=10, pady=(40, 0), sticky="nsew", rowspan=5)
        widgets_frame.columnconfigure(index=0, weight=1)

        titleS = ttk.Label(widgets_frame, justify=LEFT, text="If there has been a print error:\n  Please select 'View Documents' to be taken to the Work Order Pdf's & manually print the errored document.\n\nIf there has been no issue:\n  Please select 'Finish Print Run' to remove the temporay documents")
        titleS.grid(row=1, column=0,  padx=5, pady=10, sticky="ew")

        accentbutton.grid_forget()
        accentbuttonE = ttk.Button(widgets_frame, text="Finish Print Run", style="Accent.TButton", width=15, command=processDel)
        accentbuttonE.grid(row=2, column=0, padx=5, pady=0, sticky="w")  
        accentbuttonD = ttk.Button(widgets_frame, text="View Documents", style="Accent.TButton", width=15, command=openFIles)
        accentbuttonD.grid(row=2, column=0, padx=0, pady=0, sticky="")   
        
        pdf_splitter(defaultFile)
     except Exception as e: 
        out = open(logfile_name, 'a')
        out.write(e)
        out.close()

        title = ttk.Label(widgets_frame, justify=LEFT, text="Error When Exporting Pages", foreground='red')
        title.grid(row=0, column=0,  padx=5, pady=(0, 20), sticky="ew")
        accentbutton.grid(row=2, column=0, padx=5, pady=0, sticky="nsew")     

def process():
        widgets_frame.focus_set()
        thread = runner(file_spliter)
        thread.start()
   
def processDel():
        for file in files:
            try:
                os.remove(file)
            except Exception as e: 
                out = open(logfile_name, 'a')
                out.write(e)
                out.close()
                continue
            

        root.destroy()

def defocus(event):
    event.widget.master.focus_set()

def file_picker():
    global defaultFile
    accentbutton.grid_forget()
    baseFile = filedialog.askopenfile(initialdir = "/temp",title = "Select file",filetypes = (("PDF files","*.pdf"),("all files","*.*")))
    defaultFile = baseFile.name
    accentbutton.grid(row=3, column=0, padx=5, pady=10, sticky="nsew")
    entry = ttk.Entry(widgets_frame, width=50)
    entry.insert(0, baseFile.name)
    entry.grid(row=1, column=0,  padx=5, pady=(0, 10), sticky="ew")
    
    title.grid_forget()

    strr = "File Selected: " + os.path.splitext(os.path.basename(defaultFile))[0] + "\nValidated: " + countPage(defaultFile) 
    titles = ttk.Label(widgets_frame, justify=LEFT, text=strr, foreground='light green')
    titles.grid(row=0, column=0,  padx=5, pady=(0, 20), sticky="ew")

    
def countPage(name):
    
    fnames = os.path.splitext(os.path.basename(name))[0]
    printable= 0
    pdfs = PdfFileReader(name)
    page = 0
    while page < pdfs.getNumPages():
       if title_page_find(pdfs.getPage(page).extractText()):
           printable += 1
       page += 1

    return str(printable) + " Workorders Found accross " + str(pdfs.getNumPages()) + " pages"

root = tk.Tk()
root.title("Process Workorder")
root.option_add("*tearOff", False) # This is always a good idea

# Make the app responsive
root.columnconfigure(index=0, weight=1)
root.columnconfigure(index=1, weight=1)
root.columnconfigure(index=2, weight=1)
root.rowconfigure(index=0, weight=1)
root.rowconfigure(index=1, weight=1)
root.rowconfigure(index=2, weight=1)

# Create a style
style = ttk.Style(root)

# Import the tcl file
root.tk.call("source", "azure.tcl")

# Set the theme with the theme_use method
root.tk.call("set_theme", "dark")

# Create a Frame for input widgets
widgets_frame = ttk.Frame(root, padding=(0, 0, 0, 10))
widgets_frame.grid(row=0, column=1, padx=10, pady=(40, 0), sticky="nsew", rowspan=5)
widgets_frame.columnconfigure(index=0, weight=1)

# Create a Frame for input widgets
widgets_frame2 = ttk.Frame(root, padding=(0, 0, 0, 10))
widgets_frame2.grid(row=1, column=1, padx=10, pady=(160, 10), sticky="nsew", rowspan=3)
widgets_frame2.columnconfigure(index=1, weight=1)

strng = "Please Select the .pdf exported by Maximo"

title = ttk.Label(widgets_frame, justify=LEFT, text=strng)
title.grid(row=0, column=0,  padx=5, pady=(0, 20), sticky="ew")

title2 = ttk.Label(widgets_frame, justify=LEFT, text="--------------------------------------------\n")
title2.grid(row=0, column=0,  padx=5, pady=(40, 0), sticky="ew")

# Entry
entry = ttk.Entry(widgets_frame, width=50)
entry.insert(0, defaultFile)
entry.grid(row=1, column=0,  padx=5, pady=(0, 10), sticky="ew")


ComboSelect = tk.StringVar()
keepvalue = ComboSelect.get()
drop = ttk.Combobox(widgets_frame2, textvariable=keepvalue, values=('Print Work Orders', 'Create Test Files'), state='readonly')

keepvalue = drop.get()
drop.grid(row=4, column=0, padx=5, pady=10, sticky="    ew")
drop.current(0)


drop.SelectionLength = 0

# Accentbutton
accentbutton = ttk.Button(widgets_frame2, text="Print", style="Accent.TButton", width=10, command=process)
accentbutton.grid(row=3, column=0, padx=5, pady=10, sticky="nsew")

button = ttk.Button(widgets_frame, text="Select File", command=file_picker)
button.grid(row=1, column=20, padx=0, pady=(0, 10), sticky="nsew")

# Sizegrip
sizegrip = ttk.Sizegrip(root)
sizegrip.grid(row=100, column=100, padx=(0, 5), pady=(0, 5))

# Center the window, and set minsize
root.update()
root.minsize(root.winfo_width(), root.winfo_height())
x_cordinate = int((root.winfo_screenwidth()/2) - (root.winfo_width()/2))
y_cordinate = int((root.winfo_screenheight()/2) - (root.winfo_height()/2))
root.geometry("+{}+{}".format(x_cordinate, y_cordinate))


# Start the main loop
root.mainloop()
