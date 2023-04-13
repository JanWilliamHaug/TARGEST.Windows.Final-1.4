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
from tkinter.scrolledtext import ScrolledText
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

import Targest
global scrolled_text_box

from tkinter import messagebox

import webbrowser


def GUI1():
   
    try:
        # Creates the gui
        window = Tk(className=' TARGEST v.1.16.1 ')
        # set window size #
        window.geometry("1000x620")

        canvas = Canvas(window, width=1000, height=620)
        canvas.pack()

        # Create a horizontal gradient
        for i in range(1000):
            r = int(i/1000 * 152)
            g = 75  # fixed green value
            b = 255 - int(i/1000 * 108)
            color = '#{:02x}{:02x}{:02x}'.format(r, g, b)
            canvas.create_rectangle(i, 0, i+1, 1000, fill=color, outline='')

        icon = PhotoImage(file='TARGEST.png')
        window.iconphoto(True, icon)

        # Create a style for the widgets
        style = ttk.Style()
        #style.configure('Emergency.TButton', font='helvetica 24', foreground='red', padding=10)
        style.configure("TButton", font=("Segoe UI", 10, "bold"), background="#b2d8ff", foreground="black")


        # Create a canvas 
        #canvas = tk.Canvas(window, width=200, height=400)
        #canvas.place(x=500, y=500)
        #canvas.pack(side='left')


        # Load your anime figures as image files
        #figure1 = tk.PhotoImage(file='TARGEST2.png', width=200, height=400)
        #figure2 = tk.PhotoImage(file='eren.png')

        # Place your anime figures on the canvas
        #canvas.create_image(100, 100, image=figure1)

        # button 1
        ttk.Button(window, text="Choose list of Documents", command=Targest2.generateReport, width = 22).place(x=136, y=10)

        # button 2
        global genRep
        genRep = ttk.Button(window, text="Generate Reports", state= DISABLED, command=Targest2.generateReport2, width = 22)
        genRep.place(x=330, y=10)

        # button 3
        global allTagsButton
        allTagsButton = ttk.Button(text="Open All Tags Table Report", state= DISABLED, command=Targest2.getDocumentTable, width = 34)
        allTagsButton.place(x=627, y=10)

        # button 4
        global getDoc
        getDoc = ttk.Button(window, text="Open Child and Parent Tags Report", state= DISABLED, command=Targest2.getDocument, width = 34)
        getDoc.place(x=627, y=35)

        # button 5
        global getOrphanDoc
        getOrphanDoc = ttk.Button(text="Open Orphan Tags Report", state= DISABLED, command=Targest2.getOrphanDocument, width = 34)
        getOrphanDoc.place(x=627, y=60)

        # button 6
        global getChildlessDoc
        getChildlessDoc = ttk.Button(text="Open Childless Tags Report", state= DISABLED, command=Targest2.getChildlessDocument, width = 34)
        getChildlessDoc.place(x=627, y=85)

        # button 7
        global getTBVdoc
        getTBVdoc = ttk.Button(text="Open TBV Word Report", state= DISABLED, command=Targest2.getTBV, width = 34)
        getTBVdoc.place(x=627, y=110)

        # button 8
        global getTBDdoc
        getTBDdoc = ttk.Button(text="Open TBD Word Report", state= DISABLED, command=Targest2.getTBD, width = 34)
        getTBDdoc.place(x=627, y=135)

        # button 9
        global getExcel
        getExcel = ttk.Button(text="Open Requirements Excel Report", state= DISABLED, command=Targest2.createExcel, width = 34)
        getExcel.place(x=627, y=160)

        # button 10
        #global getExcel2
        #getExcel2 = ttk.Button(text="Open All Tags Excel Report", state= DISABLED, command=Targest2.createExcel2, width = 30)
        #getExcel2.place(x=620, y=185)

        # button 10
        global getExcel2
        getExcel2 = ttk.Button(text="Open Relationship Trees Excel Report", state= DISABLED, command=Targest2.createExcel3, width = 34)
        getExcel2.place(x=627, y=185)
        
        # button 10
        global TreeDiagram
        TreeDiagram = ttk.Button(text="Create Family Trees", state= DISABLED, command =lambda: Targest.text3(window), width = 34)
        TreeDiagram.place(x=627, y=210)

        # button 11
        global Website
        Website = ttk.Button(text="Visit our Website", state= ACTIVE, command =lambda: open_website(), width = 30)
        Website.place(x=205, y=62)
       

        #global button
        #button = Button(text="End Program", command=window.destroy, width = 30, font=("Segoe UI", 10), background="#4CAF50", foreground="white")
        #button.place()

        # button 11
        global button
        button = ttk.Button(text="End Program", command=lambda:[window.destroy(), Targest2.closeReports(), Targest2.closeExcelWorkbooks()], width = 30)
        button.place(x=205, y=38)

        # Create text widget and specify size.
        global Txt
        Txt = ScrolledText(window, wrap=tk.WORD, height = 30, width = 60)
        Txt.place(x=25, y=120)
        Txt.configure(bg='grey', fg='white')

        # Create a label for the developers
        labelDevs = Label(window, text="Developers:\nJan William Haug\nAdrian Bernardino\nStephania Rey", font=("Segoe UI", 10, "bold"), bg="#E5CCFF")
        labelDevs.place(x=690, y=490)
        labelDevs.config(borderwidth=2, relief="groove", padx=10, pady=5, fg="black")
        

        # Create ScrolledText widget
        scrolled_text_box = ScrolledText(window, wrap=tk.WORD, height=15, width=47)
        scrolled_text_box.place(x=566, y=240)
        scrolled_text_box.configure(bg='grey', fg='white') 

        # Load the image file
        global imageLogo
        imageLogo = PhotoImage(file="TARGEST3.png")

        # Create a label to display the image
        label2 = Label(window, image=imageLogo)
        label2.place(x=25, y=10)

        
        msg3 = ('You need a text file with paths to your documents\n 1. Please choose your documents by clicking on \n    the "Choose list of Documents" button.\n 2. Once the documents are displayed, Click "Generate Reports"\n\n')
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


    
def open_website():
    webbrowser.open("https://targest-website.vercel.app/")