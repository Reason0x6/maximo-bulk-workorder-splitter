# Prototype UI for Maximo Workorder Export Splitter
# Designed by github\reason0x6

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter.constants import HORIZONTAL, LEFT
import os
import threading
from PyPDF2 import PdfFileReader, PdfFileWriter
import asyncio
import concurrent.futures


class runner(threading.Thread):

    def __init__(self, function_that_downloads):
        threading.Thread.__init__(self)
        self.runnable = function_that_downloads
        self.daemon = True

    def run(self):
        self.runnable()


files = []
global defaultFile
defaultFile =  "C:/temp/SilentPrintServlet.pdf"

def title_page_find(text):
    if "Location" in text and "Asset" in text and "Job Plan" in text and "Reported By" in text:
        return True

def pdf_splitter(path):
    
    fname = os.path.splitext(os.path.basename(path))[0]
    printed = 1

    pdf = PdfFileReader(path)
    for page in range(pdf.getNumPages()):

        test = title_page_find(pdf.getPage(page).extractText())
     
        if test:
            pdf_writer = PdfFileWriter()
            pdf_writer.addPage(pdf.getPage(page))

            output_filename = '{}_Section_{}.pdf'.format(
                fname, printed)
            
            with open(output_filename, 'wb') as out:
                pdf_writer.write(out)
                files.append(output_filename)
                printed += 1
        elif pdf.getPage(page).extractText().isspace():
                continue
        else:
            pdf_writer.addPage(pdf.getPage(page))
            with open(output_filename, 'ab') as out:
                pdf_writer.write(out)

def file_picker():
    global defaultFile
    baseFile = filedialog.askopenfile(initialdir = "/temp",title = "Select file",filetypes = (("PDF files","*.pdf"),("all files","*.*")))
    defaultFile = baseFile.name
    entry = ttk.Entry(widgets_frame, width=50)
    entry.insert(0, baseFile.name)
    entry.grid(row=1, column=0,  padx=5, pady=(0, 10), sticky="ew")

def file_spliter():
     global title
     try:
        title.grid_forget()
        title = ttk.Label(widgets_frame, justify=LEFT, text="Pages Exporting......")
        title.grid(row=0, column=0,  padx=5, pady=(0, 20), sticky="ew")
        accentbutton.grid_forget()
        
        pdf_splitter(defaultFile)
     except:
        title = ttk.Label(widgets_frame, justify=LEFT, text="Error When Exporting Pages")
        title.grid(row=0, column=0,  padx=5, pady=(0, 20), sticky="ew")
        accentbutton.grid(row=2, column=0, padx=5, pady=10, sticky="nsew")     

def process():
        thread = runner(file_spliter)
        thread.start()
   

    
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
root.tk.call("source", "azure-dark.tcl")

# Set the theme with the theme_use method
style.theme_use("azure-dark")



# Create a Frame for input widgets
widgets_frame = ttk.Frame(root, padding=(0, 0, 0, 10))
widgets_frame.grid(row=0, column=1, padx=10, pady=(40, 0), sticky="nsew", rowspan=3)
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



# Accentbutton
accentbutton = ttk.Button(widgets_frame2, text="Process", style="AccentButton", width=10, command=process)
accentbutton.grid(row=2, column=0, padx=5, pady=10, sticky="nsew")

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
