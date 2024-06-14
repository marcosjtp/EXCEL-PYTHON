import tkinter
#from tkinter import *
from tkinter import  filedialog ,Tk,Entry,StringVar
import openpyxl

#from fpdf.fpdf import FPDF
#import subprocess
import csv
from tkinter import messagebox
import os
import subprocess    
varia1=""
wre = ""
peso = ""
cliente=""
desc=""
cobrar=""
listas=[["WR","PESO"]]
lista=[]
hoja=""
root=tkinter.Tk(screenName="Convertidor", baseName=None, className='Tix')
root.geometry("370x300+500+200")
root.resizable(0,0)
root.wm_attributes("-topmost", 0)
#root.iconbitmap(default='Libros.ico')
Entry1=Entry(root)
Entry1.place(rely=0.03,relx=0.045,relheight=0.1,relwidth=0.90)
Entry1.config(background="grey",textvariable=varia1)

Entry2=Entry(root)
Entry2.place(rely=0.5,relx=0.28,relheight=0.1,relwidth=0.40)
Entry2.config(background="grey",foreground="red")




root.mainloop()





