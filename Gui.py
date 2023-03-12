# from debug import debug
import logging
# import pdb
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

def GUI1():
   
    try:
       
            # pdb.set_trace()
        # Creates a word document, saves it as "report 3, and also adds a heading
        
        # Creates the gui
        window = Tk(className=' TARGEST v.1.7.x ')
        # set window size #
        window.geometry("600x600")
        window['background'] = '#afeae6'

        # Create a style for the widgets
        style = ttk.Style()
        #style.configure('Emergency.TButton', font='helvetica 24', foreground='red', padding=10)
        style.configure("TButton", font=("Segoe UI", 10), background="#4CAF50", foreground="green")


        # Create a canvas on the left side of the window
        #canvas = tk.Canvas(window, width=200, height=400)
        #canvas.pack(side='left')


        # Load your anime figures as image files
        #figure1 = tk.PhotoImage(file='eren.png')
        #figure2 = tk.PhotoImage(file='eren.png')

        # Place your anime figures on the canvas
        #canvas.create_image(50, 50, image=figure1)

        # Creates button 1pi
        ttk.Button(window, text="Choose list of Documents", command=Targest2.generateReport, width = 26).pack()
        # Creates button 2
        global genRep
        genRep = ttk.Button(window, text="Generate Reports ", state= DISABLED, command=Targest2.generateReport2, width = 26)
        genRep.pack()
        global allTagsButton
        allTagsButton = ttk.Button(text="Open all tags Report", state= DISABLED, command=Targest2.getDocumentTable, width = 26)
        allTagsButton.pack()
        # Creates button 3
        global getDoc
        getDoc = ttk.Button(window, text="Open Generated Report", state= DISABLED, command=Targest2.getDocument, width = 26)
        getDoc.pack()
        # Creates button 4
        global getOrphan
        getOrphan = ttk.Button(text="Generate Orphan Report", state= DISABLED, command=Targest2.orphanGenReport, width = 26)
        getOrphan.pack()
        # Creates button 5
        global getOrphanDoc
        getOrphanDoc = ttk.Button(text="Open Orphan Tags Report", state= DISABLED, command=Targest2.getOrphanDocument, width = 26)
        getOrphanDoc.pack()
        # Creates button 6
        global getChildlessDoc
        getChildlessDoc = ttk.Button(text="Open childless tags Report", state= DISABLED, command=Targest2.getChildlessDocument, width = 26)
        getChildlessDoc.pack()
        # Creates Excel button button 7
        global getExcel
        getExcel = ttk.Button(text="Open Generated Excel Report", state= DISABLED, command=Targest2.createExcel, width = 26)
        getExcel.pack()

        # Creates Excel button button 7
        global getExcel2
        getExcel2 = ttk.Button(text="Open Generated Excel Report2", state= DISABLED, command=Targest2.createExcel2, width = 26)
        getExcel2.pack()
        
        # Creates button 9
        #global button
        #button = Button(text="End Program", command=window.destroy, width = 26, font=("Segoe UI", 10), background="#4CAF50", foreground="white")
        #button.pack()
        global button
        button = ttk.Button(text="End Program", command=window.destroy, width = 26)
        button.pack()


        # Create text widget and specify size.
        global Txt
        Txt = Text(window, height = 25, width = 55)
        Txt.pack()

        
        msg3 = ('You need a text file with paths to your documents\n 1. Please choose your documents by clicking on \n    the "choose document" button.\n 2. Click "Generate Reports".  \n\n')
        Txt.insert(tk.END, msg3) #print in GUI
        
    except Exception as e:
        # Log an error message
        logging.exception('main(): ERROR', exc_info=True)
    else:
        # Log a success message
        logging.info('main(): PASS')

        window.mainloop()
