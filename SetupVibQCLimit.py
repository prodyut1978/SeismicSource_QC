import PySimpleGUI as sg
import os
import pandas as pd
import numpy as np
import glob
import datetime
from datetime import datetime
import csv
import openpyxl
import matplotlib.pyplot as plt
import pickle
import tkinter.ttk as ttk
import tkinter as tk
import tkinter.messagebox
from tkinter.filedialog import asksaveasfile
from tkinter.filedialog import askopenfilenames
from tkinter import simpledialog

def VibQCLimitParameter():
    if not os.path.exists("C:\VibRestrictedFolder\VibIFQC\VibQCLimitParameter"):
        os.makedirs("C:\VibRestrictedFolder\VibIFQC")
        layout = [
                    [sg.Text('A : THD Parameters Setup',                           size=(25, 1), font=('Helvetica', 12), background_color='blue', text_color='white')],
                    [sg.Text('1. Enter Desired THDAvg Min Limit (Always = 0):',    size=(40, 1)), sg.Slider(range=(0, 0), orientation='h', size=(6.7, 10), font=('Helvetica', 8) ,default_value=0)],          
                    [sg.Text('2. Enter Desired THDAvg Max Limit (Default = 25):',  size=(40, 1)), sg.InputText(25)],
                    [sg.Text('3. Enter Desired THDMax Min Limit (Always = 0):',    size=(40, 1)), sg.Slider(range=(0, 0), orientation='h', size=(6.7, 10), font=('Helvetica', 8) ,default_value=0)],
                    [sg.Text('4. Enter Desired THDMax Max Limit (Default = 50):',  size=(40, 1)), sg.InputText(50)],

                    [sg.Text('B : Force Parameters Setup',                          size=(25, 1), font=('Helvetica', 12), background_color='blue', text_color='white')],              
                    [sg.Text('1. Enter Desired ForceAvg Min Limit (Default = 30):', size=(40, 1)), sg.InputText(30)],
                    [sg.Text('2. Enter Desired ForceAvg Max Limit (Default = 100):',size=(40, 1)), sg.InputText(100)],
                    [sg.Text('3. Enter Desired ForceMax Min Limit (Default =40):',  size=(40, 1)), sg.InputText(40)],
                    [sg.Text('4. Enter Desired ForceMax Max Limit (Default = 100):', size=(40, 1)), sg.InputText(100)],

                    [sg.Text('C : Phase Parameters Setup',                          size=(25, 1), font=('Helvetica', 12), background_color='blue', text_color='white')],
                    [sg.Text('1. Enter Desired PhaseAvg Max Limit (Default = 2):', size=(40, 1)), sg.InputText(2)],              
                    [sg.Text('2. Enter Desired PhaseMax Max Limit (Default = 10):',size=(40, 1)), sg.InputText(10)],

                  [sg.Submit(), sg.Cancel()]
                 ]

        window = sg.Window('Please Input Vibroseis QC Parameters Limit:',auto_size_text=True, default_element_size=(10, 1)).Layout(layout)
        event, values = window.read()

        if event is None or event == 'Cancel':
            Low_THDAvg_Limit        = 0
            High_THDAvg_Limit       = 25

            Low_THDMax_Limit        = 0
            High_THDMax_Limit       = 50

            Low_ForceAvg_Limit      = 30
            High_ForceAvg_Limit     = 100

            Low_ForceMax_Limit      = 40
            High_ForceMax_Limit     = 100

            Low_PhaseAvg_Limit      = -2
            High_PhaseAvg_Limit     = 2

            Low_PhaseMax_Limit      = -10
            High_PhaseMax_Limit     = 10

            pickle_dict = {1:Low_THDAvg_Limit,   2:High_THDAvg_Limit,
                           3:Low_THDMax_Limit,   4:High_THDMax_Limit,
                           5:Low_ForceAvg_Limit, 6:High_ForceAvg_Limit,
                           7:Low_ForceMax_Limit, 8:High_ForceMax_Limit,
                           9:Low_PhaseAvg_Limit, 10:High_PhaseAvg_Limit,
                           11:Low_PhaseMax_Limit,12:High_PhaseMax_Limit}
            pickle_out  = open("C:\VibRestrictedFolder\VibIFQC\VibQCLimitParameter","wb")
            pickle.dump(pickle_dict,pickle_out)
            pickle_out.close()
            sg.PopupAutoClose('Exiting VibroseisQC Limit Input',line_width=60)

        else:    
            Low_THDAvg_Limit        = float(values[0])
            High_THDAvg_Limit       = float(values[1])

            Low_THDMax_Limit        = float(values[2])
            High_THDMax_Limit       = float(values[3])

            Low_ForceAvg_Limit      = float(values[4])
            High_ForceAvg_Limit     = float(values[5])

            Low_ForceMax_Limit      = float(values[6])
            High_ForceMax_Limit     = float(values[7])

            Low_PhaseAvg_Limit      = -(float(values[8]))
            High_PhaseAvg_Limit     = float(values[8])

            Low_PhaseMax_Limit      = -(float(values[9]))
            High_PhaseMax_Limit     = float(values[9])

            pickle_dict = {1:Low_THDAvg_Limit,   2:High_THDAvg_Limit,
                           3:Low_THDMax_Limit,   4:High_THDMax_Limit,
                           5:Low_ForceAvg_Limit, 6:High_ForceAvg_Limit,
                           7:Low_ForceMax_Limit, 8:High_ForceMax_Limit,
                           9:Low_PhaseAvg_Limit, 10:High_PhaseAvg_Limit,
                           11:Low_PhaseMax_Limit,12:High_PhaseMax_Limit}
            pickle_out  = open("C:\VibRestrictedFolder\VibIFQC\VibQCLimitParameter","wb")
            pickle.dump(pickle_dict,pickle_out)
            pickle_out.close()
            tkinter.messagebox.showinfo("Vib QC Limit Input Message","Vib QC Limit Input Profile Is Created")            
        window.close()    

    else:
        Pickle_in       = open("C:\VibRestrictedFolder\VibIFQC\VibQCLimitParameter","rb")
        pickle_dict     = pickle.load(Pickle_in)

        Low_THDAvg_Limit  = pickle_dict[1]
        High_THDAvg_Limit = pickle_dict[2]
        Low_THDMax_Limit  = pickle_dict[3]
        High_THDMax_Limit = pickle_dict[4]

        Low_ForceAvg_Limit  = pickle_dict[5]
        High_ForceAvg_Limit = pickle_dict[6]
        Low_ForceMax_Limit  = pickle_dict[7]
        High_ForceMax_Limit = pickle_dict[8]

        Low_PhaseAvg_Limit  = pickle_dict[9]
        High_PhaseAvg_Limit = pickle_dict[10]
        Low_PhaseMax_Limit  = pickle_dict[11]
        High_PhaseMax_Limit = pickle_dict[12]
        
        Pickle_in.close()
        
        layout = [
                    [sg.Text('A : THD Parameters Setup',                           size=(25, 1), font=('Helvetica', 12), background_color='blue', text_color='white')],
                    [sg.Text('1. Enter Desired THDAvg Min Limit (Always = 0):',    size=(40, 1)), sg.Slider(range=(0, 0), orientation='h', size=(6.7, 10), font=('Helvetica', 8) ,default_value=Low_THDAvg_Limit)],          
                    [sg.Text('2. Enter Desired THDAvg Max Limit (Default = 25):',  size=(40, 1)), sg.InputText(High_THDAvg_Limit)],
                    [sg.Text('3. Enter Desired THDMax Min Limit (Always = 0):',    size=(40, 1)), sg.Slider(range=(0, 0), orientation='h', size=(6.7, 10), font=('Helvetica', 8) ,default_value=Low_THDMax_Limit)],
                    [sg.Text('4. Enter Desired THDMax Max Limit (Default = 50):',  size=(40, 1)), sg.InputText(High_THDMax_Limit)],

                    [sg.Text('B : Force Parameters Setup',                          size=(25, 1), font=('Helvetica', 12), background_color='blue', text_color='white')],              
                    [sg.Text('1. Enter Desired ForceAvg Min Limit (Default = 30):', size=(40, 1)), sg.InputText(Low_ForceAvg_Limit)],
                    [sg.Text('2. Enter Desired ForceAvg Max Limit (Default = 100):',size=(40, 1)), sg.InputText(High_ForceAvg_Limit)],
                    [sg.Text('3. Enter Desired ForceMax Min Limit (Default =40):',  size=(40, 1)), sg.InputText(Low_ForceMax_Limit)],
                    [sg.Text('4. Enter Desired ForceMax Max Limit (Default = 100):', size=(40, 1)), sg.InputText(High_ForceMax_Limit)],

                    [sg.Text('C : Phase Parameters Setup',                          size=(25, 1), font=('Helvetica', 12), background_color='blue', text_color='white')],
                    [sg.Text('1. Enter Desired PhaseAvg Max Limit (Default = 2):', size=(40, 1)), sg.InputText(High_PhaseAvg_Limit)],              
                    [sg.Text('2. Enter Desired PhaseMax Max Limit (Default = 10):',size=(40, 1)), sg.InputText(High_PhaseMax_Limit)],

                  [sg.Submit(), sg.Cancel()]
                 ]


        window = sg.Window('Please Input Vibroseis QC Parameters Limit:',auto_size_text=True, default_element_size=(10, 1)).Layout(layout)
        event, values = window.Read()

        if event is None or event == 'Cancel':
            sg.PopupAutoClose('Exiting VibroseisQC Limit Input',line_width=60)
            
        else:        
            Low_THDAvg_Limit        = float(values[0])
            High_THDAvg_Limit       = float(values[1])

            Low_THDMax_Limit        = float(values[2])
            High_THDMax_Limit       = float(values[3])

            Low_ForceAvg_Limit      = float(values[4])
            High_ForceAvg_Limit     = float(values[5])

            Low_ForceMax_Limit      = float(values[6])
            High_ForceMax_Limit     = float(values[7])

            Low_PhaseAvg_Limit      = -(float(values[8]))
            High_PhaseAvg_Limit     = float(values[8])

            Low_PhaseMax_Limit      = -(float(values[9]))
            High_PhaseMax_Limit     = float(values[9])


            pickle_dict = {1:Low_THDAvg_Limit,   2:High_THDAvg_Limit,
                           3:Low_THDMax_Limit,   4:High_THDMax_Limit,
                           5:Low_ForceAvg_Limit, 6:High_ForceAvg_Limit,
                           7:Low_ForceMax_Limit, 8:High_ForceMax_Limit,
                           9:Low_PhaseAvg_Limit, 10:High_PhaseAvg_Limit,
                           11:Low_PhaseMax_Limit,12:High_PhaseMax_Limit}
            pickle_out  = open("C:\VibRestrictedFolder\VibIFQC\VibQCLimitParameter","wb")
            pickle.dump(pickle_dict,pickle_out)
            pickle_out.close()
            tkinter.messagebox.showinfo("Vib QC Limit Input Message","Vib QC Limit Input Profile Is Created")
        window.Close()  # Don't forget to close your window!



