#!/usr/bin/env python
# Data Entry 
# Sourav Bhatti.
# Date 05/16/1522
# Version 1.00



from datetime import datetime
import tkinter as tk
from tkinter import W,S, Button, StringVar, ttk
from tkinter import font

from numpy import place
import openpyxl
from tkcalendar import DateEntry, Calendar
import pandas as pd

    



root = tk.Tk()
root.geometry("850x450+500+250")
root.title("FAI DATA LOG")
root.configure(bg='steelblue')
root.resizable(0,0)

style = ttk.Style()
style.theme_use("clam")

fa_num = StringVar()
bc=StringVar()
fa_reason=StringVar()
obj_fa=StringVar()
exp_out=StringVar()
assem_h=StringVar()
pcb=StringVar()
pcba=StringVar()
next_steps=StringVar()



#left 
lblL1= tk.Label(root, text="       FA#", font=("Times new roman", 12, 'bold'), bg="steelblue").place(x=50 , y=50,anchor='center')
lblL2= tk.Label(root, text="    BC Part #", font=("Times new roman", 12, 'bold'), bg="steelblue").place(x=50 , y=100,anchor='center')
lblL3= tk.Label(root, text="    PCB MFG:  ", font=("Times new roman", 12, 'bold'), bg="steelblue").place(x=50 , y=150,anchor='center')
#lblL4= tk.Label(root, text="    Obj. of FAI: ", font=("Times new roman", 12, 'bold'), bg="steelblue").place(x=50 , y=200,anchor='center')
lblL5= tk.Label(root, text="\tExpected Outcome", font=("Times new roman", 12, 'bold'), bg="steelblue").place(x=90 , y=260, anchor='w')


entry_FAN = tk.Entry(root, textvariable= fa_num, width=35, relief="ridge", bd= 5)
entry_FAN.place(x=120 , y=50,anchor='w')

entry_BC =tk.Entry(root, textvariable= bc ,width=35, relief="ridge", bd= 5)
entry_BC.place(x=120 , y=100,anchor='w')

entry_FAR =tk.Entry(root, textvariable= pcb, width=35, relief="ridge", bd= 5)
entry_FAR.place(x=120 , y=150,anchor='w')

entry_FAI =tk.Entry(root, textvariable= obj_fa, width=35, relief="ridge", bd= 5)
#entry_FAI.place(x=120 , y=200,anchor='w')

entry_EO =tk.Text(root,height = 5,width = 36, relief="ridge", bd= 5)
entry_EO.place(x=80 , y=270)



# right 
lblR1= tk.Label(root, text="\tFA Reason", font=("Times new roman", 12, 'bold'), bg="steelblue").place(x=460 , y=50, anchor='center')
lblR2= tk.Label(root, text="\tAssem. House", font=("Times new roman", 12, 'bold'), bg="steelblue").place(x=460 , y=100,anchor='center')
lblR3= tk.Label(root, text="\tPCBA mfg", font=("Times new roman", 12, 'bold'), bg="steelblue").place(x=460 , y=150,anchor='center')
#lblR4= tk.Label(root, text="\tPCBA mfg", font=("Times new roman", 12, 'bold'), bg="steelblue").place(x=460 , y=200,anchor='center')
lblR5= tk.Label(root, text="  Next Steps", font=("Times new roman", 12, 'bold'), bg="steelblue").place(x=610 , y=260,anchor='w')

entry_AH = tk.Entry(root, textvariable= fa_reason, width=35, relief="ridge", bd= 5)
entry_AH.place(x=550 , y=50,anchor='w')

entry_FAN = tk.Entry(root, textvariable= assem_h, width=35, relief="ridge", bd= 5)
entry_FAN.place(x=550 , y=100,anchor='w')

entry_PCB = tk.Entry(root, textvariable= pcba, width=35, relief="ridge", bd= 5)
entry_PCB.place(x=550 , y=150,anchor='w')

entry_PCBA = tk.Entry(root, textvariable= pcba, width=35, relief="ridge", bd= 5)
#entry_PCBA.place(x=550 , y=200,anchor='w')

entry_NXS = tk.Text(root,height = 5,width = 36, relief="ridge", bd= 5)
entry_NXS.place(x= 520, y= 270 )


def savetoexcel():

    timestamp = str(datetime.now().strftime("%Y-%m-%d"))
    data = ({"Date": [timestamp], "FA Number": [fa_num.get()], "FA Reason": [entry_FAR.get()], "Assem. House": [entry_AH.get()], " BC Part#": [entry_BC.get()], 
    "PCB Manu.": [entry_PCB.get()],"PCBA Manu.": [entry_PCBA.get()], "Exp. OUtcome": [entry_EO.get('1.0', 'end')], "Next Steps": [entry_NXS.get('1.0', 'end')] })
    f = openpyxl.load_workbook('Data_log.xlsx')
    sheet = f.active
    maxrow = sheet.max_row
    dataframe = pd.DataFrame(data)
    with pd.ExcelWriter('Data_log.xlsx', mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
            dataframe.to_excel(writer, startrow=maxrow, startcol=0, header=None, index=False)

Button(root, width=10, text="Save", relief='raised', bd=5, bg="#bcaf67", command= lambda : savetoexcel()).place(x=300, y= 400, anchor='center')

def print():
    pass

Button(root, width=10, text="Print", relief='raised', bd=5, bg="#edba44").place(x=580, y= 400, anchor='center')


root.mainloop()



