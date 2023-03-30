import logging
import docx
from docx import Document
from docx.shared import RGBColor
from docx.shared import Inches
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from typing import Tuple

from tkinter import scrolledtext
import re
import copy
import time

# This libraries are for opening word document automatically
import os
import platform
import subprocess

# This library is for opening excel document automatically
import xlwings as xw
import pandas as pd
import matplotlib.pyplot as plt

import Targest2

from tkinter import messagebox

def GUI1():
   
    try:
        

        # Creates the gui
        window = Tk(className=' TARGEST v.1.10.1 ')
        # set window size #
        window.geometry("900x650")
        window['background'] = '#afeae6'

        #icon = tk.PhotoImage(file='itachiakatttt.png')
        icon = PhotoImage(file='itachiakatttt.png')
        window.iconphoto(True, icon)

        # Create a style for the widgets
        style = ttk.Style()
        #style.configure('Emergency.TButton', font='helvetica 24', foreground='red', padding=10)
        style.configure("TButton", font=("Segoe UI", 10), background="#afeae6", foreground="green")


        # Create a canvas on the left side of the window
        #canvas = tk.Canvas(window, width=200, height=400)
        #canvas.pack(side='left')


        # Load your anime figures as image files
        #figure1 = tk.PhotoImage(file='eren.png')
        #figure2 = tk.PhotoImage(file='eren.png')

        # Place your anime figures on the canvas
        #canvas.create_image(50, 50, image=figure1)

        # button 1
        ttk.Button(window, text="Choose list of Documents", command=Targest2.generateReport, width = 20).place(x=100, y=10)

        # button 2
        global genRep
        genRep = ttk.Button(window, text="Generate Reports", state= DISABLED, command=Targest2.generateReport2, width = 20)
        genRep.place(x=280, y=10)

        # button 3
        global allTagsButton
        allTagsButton = ttk.Button(text="Open All Tags Table Report", state= DISABLED, command=Targest2.getDocumentTable, width = 30)
        allTagsButton.place(x=600, y=10)

        # button 4
        global getDoc
        getDoc = ttk.Button(window, text="Open Child and Parent Tags Report", state= DISABLED, command=Targest2.getDocument, width = 30)
        getDoc.place(x=600, y=35)

        # button 5
        global getOrphanDoc
        getOrphanDoc = ttk.Button(text="Open Orphan Tags Report", state= DISABLED, command=Targest2.getOrphanDocument, width = 30)
        getOrphanDoc.place(x=600, y=60)

        # button 6
        global getChildlessDoc
        getChildlessDoc = ttk.Button(text="Open Childless Tags Report", state= DISABLED, command=Targest2.getChildlessDocument, width = 30)
        getChildlessDoc.place(x=600, y=85)

        # button 7
        global getTBVdoc
        getTBVdoc = ttk.Button(text="Open TBV Word Report", state= DISABLED, command=Targest2.getTBV, width = 30)
        getTBVdoc.place(x=600, y=110)

        # button 8
        global getTBDdoc
        getTBDdoc = ttk.Button(text="Open TBD Word Report", state= DISABLED, command=Targest2.getTBD, width = 30)
        getTBDdoc.place(x=600, y=135)

        # button 9
        global getExcel
        getExcel = ttk.Button(text="Open Tags and Requirements Excel Report", state= DISABLED, command=Targest2.createExcel, width = 30)
        getExcel.place(x=600, y=160)

        # button 10
        global getExcel2
        getExcel2 = ttk.Button(text="Open All Tags Excel Report", state= DISABLED, command=Targest2.createExcel2, width = 30)
        getExcel2.place(x=600, y=185)

        #global button
        #button = Button(text="End Program", command=window.destroy, width = 30, font=("Segoe UI", 10), background="#4CAF50", foreground="white")
        #button.place()

        # button 11
        global button
        button = ttk.Button(text="End Program", command=lambda:[window.destroy(), Targest2.closeReports(), Targest2.closeExcelWorkbooks()], width = 20)
        button.place(x=190, y=40)

        # Create text widget and specify size.
        global Txt
        Txt = Text(window, height = 40, width = 75)
        Txt.place(x=30, y=80)
        Txt.configure(bg='grey', fg='white')

        
        msg3 = ('You need a text file with paths to your documents\n 1. Please choose your documents by clicking on \n    the "Choose list of Documents" button.\n 2. Once your documents are displayed, Click "Generate Reports"\n\n')
        Txt.insert(tk.END, msg3) #print in GUI

        # show a pop-up message
        #messagebox.showinfo("Welcome to TARGEST",  "Make sure you have closed all your previous Word Reports and Excel Reports, before running this application")
        messagebox.showinfo("Welcome to TARGEST",  "Make sure to save a text file with the paths to the documents you want to use, if you haven't already")

    except Exception as e:
        # Log an error message
        logging.exception('main(): ERROR', exc_info=True)
    else:
        # Log a success message
        logging.info('main(): PASS')

        window.mainloop()
