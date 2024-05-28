import os
from tkinter import*
import tkinter.messagebox
import GeomergeAUXTB_BackEnd
import tkinter.ttk as ttk
import tkinter as tk
import sqlite3
from tkinter.filedialog import asksaveasfile
from tkinter.filedialog import askopenfilenames
from tkinter import simpledialog
import pandas as pd
import numpy as np
import openpyxl
import csv
import time
import datetime
import sys
import PySimpleGUI as sg

def VibAuxProfileSetup():
    layout = [[sg.Text('Enter Number Of Vibes :',    size=(40,1)), sg.InputText()],
              [sg.Text('Enter AUX Unit Device Type:',    size=(40,1)), sg.InputText(257)],
              [sg.Submit(), sg.Cancel()]]

    WindowRead = sg.Window('Input Number Of Vibes',
             auto_size_text=True, default_element_size=(10, 1)).Layout(layout)
    event, values = WindowRead.read()    

    if event is None or event == 'Cancel':
        sg.PopupAutoClose('Exiting Vib Number Input',line_width=60)
        
    else:
        TotalNumberVibes = values[0]
        AuxDeviceType    = values[1]
        
        if (len(TotalNumberVibes) == 0)| (len(AuxDeviceType) == 0):
            tkinter.messagebox.showinfo("Aux profile Error Message","TotalNumberVibes or AuxDeviceType can not be empty")
        else:
            try:
                TotalNumberVibes = int(TotalNumberVibes)
                TotalNumberVibes = TotalNumberVibes + 1
                layout = []
                for i in range(1,TotalNumberVibes):
                    layout += [sg.Text(f'{i}. ' "Enter Vib" f'{i}. ' "PSS Profile ID :" , size=(25,1)),  sg.In(i),
                               sg.Text(f'      {i}. ' "Enter Vib" f'{i}. ' "AUX Unit No :" , size=(25,1)),  sg.In()],
                layout += [[sg.Button('Submit'), sg.Button('Cancel')]]

                ReadwindowVibInfo = sg.Window('Vib Input Profile', auto_size_text=True, default_element_size=(10, 1)).Layout(layout)
                event, values = ReadwindowVibInfo.read()
            
                if event is None or event == 'Cancel':
                        sg.PopupAutoClose('Exiting Vib Input for TB QC',line_width=60)
                        
                else:                       
                    Vib_ProfileID     = []
                    VibAUXUnitList    = []
                    
                    Len_Values       = len(values)

                    for i in range(0, Len_Values, 2):
                        Vib_ProfileID.append(values[i])
                        
                    for i in range(1, Len_Values, 2):
                        VibAUXUnitList.append(values[i])

                    VibAUXProfile = pd.DataFrame({'ProfileId': Vib_ProfileID, 'AUXUnitNumber': VibAUXUnitList})
                    if (VibAUXProfile['AUXUnitNumber'].duplicated().values.any() == True)| (VibAUXProfile['ProfileId'].duplicated().values.any() == True)| (VibAUXProfile['AUXUnitNumber'].isnull().values.any() == True)| (VibAUXProfile['ProfileId'].isnull().values.any() == True):
                            tkinter.messagebox.showinfo("Aux profile Error Message","Duplicate Or Empty ProfileId Or AUXUnitNumber")
                    else:
                        VibAUXProfile["DeviceType"] = VibAUXProfile.shape[0]*[AuxDeviceType]
                        VibAUXProfile['ProfileId']  = VibAUXProfile['ProfileId'].astype(int)
                        VibAUXProfile['AUXUnitNumber']  = VibAUXProfile['AUXUnitNumber'].astype(int)
                        con= sqlite3.connect("GeomergeAUXTB.db")
                        cur=con.cursor()                
                        VibAUXProfile.to_sql('AUXBOX_Profile', con, if_exists="replace", index=False)                    
                        con.commit()
                        con.close()
                        tkinter.messagebox.showinfo("Vib Aux Unit profile Message","Vib AUX Unit Profile Created And Stored in DB")
                ReadwindowVibInfo.Close()

            except ValueError:
                tkinter.messagebox.showinfo("Aux profile Error Message","Use Interger Number")

    WindowRead.Close()
        
                






                        
