#Front End
import os
import sys
from tkinter import*
import tkinter.messagebox
import Eagle_SourceLink_PSSLog_Analysis_QC_BackEnd
import SetupVibQCLimit
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
from datetime import datetime
import pickle

def SourceLink_AnalyzePSS_LogIMPORT():
    window = Tk()
    window.title ("SourceLink PSS Log QC & Analysis")
    window.geometry("1350x800+10+0")
    window.config(bg="cadet blue")
    window.resizable(0, 0)
    window.grid()
    OffsetTimeUTC = StringVar(window, value=float(+07.00))

    if not os.path.exists("C:\VibRestrictedFolder\VibIFQC\VibQCLimitParameter"):
        THDAvg_Limit_Min    = 0
        THDAvg_Limit_Max    = 25

        THDMax_Limit_Min    = 0
        THDMax_Limit_Max    = 50

        ForceAvg_Limit_Min  = 30
        ForceAvg_Limit_Max  = 100

        ForceMax_Limit_Min  = 40
        ForceMax_Limit_Max  = 100

        PhaseAvg_Limit_Min  = -2
        PhaseAvg_Limit_Max  = 2

        PhaseMax_Limit_Min  = -10
        PhaseMax_Limit_Max  =  10
    else:
        Pickle_in       = open("C:\VibRestrictedFolder\VibIFQC\VibQCLimitParameter","rb")
        pickle_dict     = pickle.load(Pickle_in)

        THDAvg_Limit_Min    = pickle_dict[1]
        THDAvg_Limit_Max    = pickle_dict[2]

        THDMax_Limit_Min    = pickle_dict[3]
        THDMax_Limit_Max    = pickle_dict[4]

        ForceAvg_Limit_Min  = pickle_dict[5]
        ForceAvg_Limit_Max  = pickle_dict[6]

        ForceMax_Limit_Min  = pickle_dict[7]
        ForceMax_Limit_Max  = pickle_dict[8]

        PhaseAvg_Limit_Min  = pickle_dict[9]
        PhaseAvg_Limit_Max  = pickle_dict[10]
        
        PhaseMax_Limit_Min  = pickle_dict[11]
        PhaseMax_Limit_Max  = pickle_dict[12]

    Low_THDAvg_Limit     = float(THDAvg_Limit_Min)
    High_THDAvg_Limit    = float(THDAvg_Limit_Max)

    Low_THDMax_Limit     = float(THDMax_Limit_Min)
    High_THDMax_Limit    = float(THDMax_Limit_Max)
            
    Low_ForceAvg_Limit   = float(ForceAvg_Limit_Min)
    High_ForceAvg_Limit  = float(ForceAvg_Limit_Max)
            
    Low_ForceMax_Limit   = float(ForceMax_Limit_Min)
    High_ForceMax_Limit  = float(ForceMax_Limit_Max)

    Low_PhaseAvg_Limit  = float(PhaseAvg_Limit_Min)
    High_PhaseAvg_Limit = float(PhaseAvg_Limit_Max)
        
    Low_PhaseMax_Limit  = float(PhaseMax_Limit_Min)
    High_PhaseMax_Limit = float(PhaseMax_Limit_Max)

    PhaseMax_Limit      = float(PhaseMax_Limit_Max)
    PhaseMax_Limit      = abs(PhaseMax_Limit)

    QCLimit_Summary =("Vib QC Limit Summary : " + '\n' +  '\n' +

                    "Low THD Average Limit (Min) : " + str(Low_THDAvg_Limit)+ '\n' +
                    "High THD Average Limit (Max) : " + str(High_THDAvg_Limit) + '\n' + '\n' +

                    "Low THD Amplitude Peak Limit (Min) : " + str(Low_THDMax_Limit)+ '\n' +                          
                    "High THD Amplitude Peak Limit (Max) : " + str(High_THDMax_Limit)+ '\n' + '\n' +
                      
                    "Low Force Average Limit (Min) : " + str(Low_ForceAvg_Limit)+ '\n' +
                    "High Force Average Limit (Max): " + str(High_ForceAvg_Limit)+ '\n' +  '\n' +
                      
                    "Low Force Max Limit (Min): " + str(Low_ForceMax_Limit)+ '\n' +
                    "High Force Max Limit (Max): " + str(High_ForceMax_Limit)+ '\n' + '\n' +
                      
                    "Low Phase Average Limit (Min): " + str(Low_PhaseAvg_Limit)+ '\n' +
                    "High Phase Average Limit (Max): " + str(High_PhaseAvg_Limit)+ '\n' +  '\n' +

                    "Low Phase Amplitude Peak Limit (Min): " + str(Low_PhaseMax_Limit)+ '\n' +                          
                    "High Phase Amplitude Peak Limit (Max) : " + str(High_PhaseMax_Limit)+ '\n' + '\n' +

                    "GPS Quality : " + " != No Fix")

    DataFrameTOP = LabelFrame(window, bd = 2, width = 1350, height = 8, padx= 0, pady= 1,relief = RIDGE,
                                       bg = "cadet blue",font=('aerial', 12, 'bold'))
    DataFrameTOP.pack(side=TOP)

    DataFrameBOTTOM_ACTIONS = LabelFrame(window, bd = 2, width = 1350, height = 38, padx= 0, pady= 1,relief = RIDGE,
                                       bg = "cadet blue",font=('aerial', 12, 'bold'))
    DataFrameBOTTOM_ACTIONS.place(x=0,y=612)

    DataFrameBOTTOM_IFQC = LabelFrame(window, bd = 2, width = 1350, height = 48, padx= 0, pady= 1,relief = RIDGE,
                                       bg = "cadet blue",font=('aerial', 12, 'bold'))

    DataFrameBOTTOM_IFQC.place(x=0,y=640)
    Label_IFQC_Summary = Label(DataFrameBOTTOM_IFQC, text = "PSS Analysis On Valid Entries Summary:", font=("arial", 10,'bold'),bg = "#E0EEEE")
    Label_IFQC_Summary.grid(row =0, column = 0, sticky ="W", padx= 1)

    DataFrameBOTTOM_INVALIDQC = LabelFrame(window, bd = 2, width = 1350, height = 48, padx= 0, pady= 1,relief = RIDGE,
                                       bg = "cadet blue",font=('aerial', 12, 'bold'))
    DataFrameBOTTOM_INVALIDQC.place(x=800,y=640)

    Label_INVALIDQC_Summary = Label(DataFrameBOTTOM_INVALIDQC, text = "PSS Analysis On Invalid Removed Entries Summary:", font=("arial", 10,'bold'),bg = "#E0EEEE")
    Label_INVALIDQC_Summary.grid(row =0, column = 0, sticky ="W", padx= 1)

    ##### Table Define
    TableMargin = Frame(window, bd = 2, padx= 5, pady= 3, relief = RIDGE)
    TableMargin.pack(side=TOP)
    scrollbarx = Scrollbar(TableMargin, orient=HORIZONTAL)
    scrollbary = Scrollbar(TableMargin, orient=VERTICAL)
    tree = ttk.Treeview(TableMargin, column=("column1", "column2", "column3", "column4", "column5",
                                             "column6", "column7", "column8", "column9", "column10",
                                             "column11", "column12", "column13", "column14", "column15",
                                             "column16", "column17", "column18", "column19", "column20",
                                             "column21", "column22", "column23", "column24", "column25",
                                             "column26", "column27", "column28", "column29", "column30",
                                             "column31", "column32", "column33", "column34", "column35",
                                             "column36", "column37", "column38", "column39", "column40",
                                             "column41", "column42"), height=26, show='headings')
    scrollbary.config(command=tree.yview)
    scrollbary.pack(side=RIGHT, fill=Y)
    scrollbarx.config(command=tree.xview)
    scrollbarx.pack(side=BOTTOM, fill=X)
    tree.heading("#1", text="DB_ID", anchor=W) 
    tree.heading("#2", text="ShotID", anchor=W)
    tree.heading("#3", text="FileNum", anchor=W)
    tree.heading("#4", text="EP", anchor=W)        
    tree.heading("#5", text="Line", anchor=W)
    tree.heading("#6", text="Station", anchor=W)
    tree.heading("#7", text="LocalDate", anchor=W)        
    tree.heading("#8", text="LocalTime" ,anchor=W)
    tree.heading("#9", text="PSSComment", anchor=W)
    tree.heading("#10", text="ShotStatus", anchor=W)        
    tree.heading("#11", text="PhaseMax", anchor=W)
    tree.heading("#12", text="PhaseAvg" ,anchor=W)
    tree.heading("#13", text="ForceMax", anchor=W)
    tree.heading("#14", text="ForceAvg", anchor=W)        
    tree.heading("#15", text="THDMax", anchor=W)
    tree.heading("#16", text="THDAvg", anchor=W)        
    tree.heading("#17", text="SwCksm", anchor=W)
    tree.heading("#18", text="PmCksm", anchor=W)
    tree.heading("#19", text="GPSQuality", anchor=W)
    tree.heading("#20", text="Unit_ID", anchor=W)
    tree.heading("#21", text="TBDate", anchor=W)
    tree.heading("#22", text="TBTime", anchor=W)        
    tree.heading("#23", text="TBMicro", anchor=W)            
    tree.heading("#24", text="SigFile#", anchor=W)
    tree.heading("#25", text="Latitude" ,anchor=W)
    tree.heading("#26", text="Longitude", anchor=W)        
    tree.heading("#27", text="Altitude", anchor=W)
    tree.heading("#28", text="EncoderIndex", anchor=W)
    tree.heading("#29", text="RecordIndex" ,anchor=W)        
    tree.heading("#30", text="EPCount", anchor=W)
    tree.heading("#31", text="CrewID", anchor=W)
    tree.heading("#32", text="StartCode", anchor=W)
    tree.heading("#33", text="ForceOut", anchor=W)        
    tree.heading("#34", text="GPSTime", anchor=W)
    tree.heading("#35", text="GPSAltitude", anchor=W)
    tree.heading("#36", text="Sats", anchor=W)
    tree.heading("#37", text="PDOP", anchor=W)
    tree.heading("#38", text="HDOP" ,anchor=W)        
    tree.heading("#39", text="VDOP", anchor=W)
    tree.heading("#40", text="Age", anchor=W)
    tree.heading("#41", text="StartTimeDelta", anchor=W)
    tree.heading("#42", text="SweepNumber", anchor=W)     
    tree.column('#1', stretch=NO, minwidth=0, width=0)            
    tree.column('#2', stretch=NO, minwidth=0, width=43)
    tree.column('#3', stretch=NO, minwidth=0, width=55)
    tree.column('#4', stretch=NO, minwidth=0, width=30)
    tree.column('#5', stretch=NO, minwidth=0, width=50)
    tree.column('#6', stretch=NO, minwidth=0, width=55)
    tree.column('#7', stretch=NO, minwidth=0, width=75)
    tree.column('#8', stretch=NO, minwidth=0, width=73)
    tree.column('#9', stretch=NO, minwidth=0, width=100)
    tree.column('#10', stretch=NO, minwidth=0, width=70)
    tree.column('#11', stretch=NO, minwidth=0, width=62)
    tree.column('#12', stretch=NO, minwidth=0, width=60)
    tree.column('#13', stretch=NO, minwidth=0, width=60)
    tree.column('#14', stretch=NO, minwidth=0, width=60)
    tree.column('#15', stretch=NO, minwidth=0, width=53)
    tree.column('#16', stretch=NO, minwidth=0, width=53)
    tree.column('#17', stretch=NO, minwidth=0, width=52)
    tree.column('#18', stretch=NO, minwidth=0, width=56)            
    tree.column('#19', stretch=NO, minwidth=0, width=70)
    tree.column('#20', stretch=NO, minwidth=0, width=45)
    tree.column('#21', stretch=NO, minwidth=0, width=68)
    tree.column('#22', stretch=NO, minwidth=0, width=60)
    tree.column('#23', stretch=NO, minwidth=0, width=60)
    tree.column('#24', stretch=NO, minwidth=0, width=60)
    tree.column('#25', stretch=NO, minwidth=0, width=60)
    tree.column('#26', stretch=NO, minwidth=0, width=90)
    tree.column('#27', stretch=NO, minwidth=0, width=60)
    tree.column('#28', stretch=NO, minwidth=0, width=90)
    tree.column('#29', stretch=NO, minwidth=0, width=90)
    tree.column('#30', stretch=NO, minwidth=0, width=60)
    tree.column('#31', stretch=NO, minwidth=0, width=60)
    tree.column('#32', stretch=NO, minwidth=0, width=70)
    tree.column('#33', stretch=NO, minwidth=0, width=70)
    tree.column('#34', stretch=NO, minwidth=0, width=60)
    tree.column('#35', stretch=NO, minwidth=0, width=90)
    tree.column('#36', stretch=NO, minwidth=0, width=40)
    tree.column('#37', stretch=NO, minwidth=0, width=40)
    tree.column('#38', stretch=NO, minwidth=0, width=40)
    tree.column('#39', stretch=NO, minwidth=0, width=40)
    tree.column('#40', stretch=NO, minwidth=0, width=40)
    tree.column('#41', stretch=NO, minwidth=0, width=100)
    tree.column('#42', stretch=NO, minwidth=0, width=100)
    style = ttk.Style()
    style.theme_use("clam")
    style.configure(".", font=('aerial', 8), foreground="black")
    style.configure("Treeview", foreground='black')
    style.configure("Treeview.Heading",font=('aerial', 7,'bold'), background='Ghost White', foreground='blue',fieldbackground='Ghost White')
    ##TableMargin.pack(side=LEFT)
    tree.pack()

    # All Functions defining 

    def iExit():
        iExit= tkinter.messagebox.askyesno("Eagle PSS Analysis Widget", "Confirm if you want to exit")
        if iExit >0:
            window.destroy()
            return

    def iExit_Must():
        iExit= tkinter.messagebox.askyesno("Must Exit Message", "Confirm YES Only")
        if iExit >0:
            window.destroy()
            return

    def SetupVibQCLimitParameter():
        SetupVibQCLimit.VibQCLimitParameter()
        iExit_Must()

    def FixedCorruptedPSSImport():
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_TEMP ORDER BY `ShotID` ASC ;", conn)
        conn.commit()
        conn.close()
        Corrupted_PSS_DF = pd.DataFrame(Complete_df)
        Corrupted_PSS_DF = Corrupted_PSS_DF.reset_index(drop=True)

        fileList = askopenfilenames(initialdir = "/", title = "Import SourceLink Time Break Files To Fix PSS UTC TB_Date-TB_Time-TB_Microseconds Columns" , filetypes=[('CSV File', '*.csv'), ('Excel File', ('*.xls', '*.xlsx'))])
        Length_fileList  =  len(fileList)
        if Length_fileList >0:
            if fileList:
                df_TBList = []           
                for filename in fileList:
                    if filename.endswith('.csv'):
                        df_TB             = pd.read_csv(filename, sep=',' , low_memory=False)
                        df_TB             = df_TB.iloc[:,:]
                        Unit_ID           = df_TB.loc[:,' ProfileId']
                        FileNum           = df_TB.loc[:,' ShotNumber']
                        EPNumber          = df_TB.loc[:,' EpNumber']
                        SourceLine        = df_TB.loc[:,' ShotLine']
                        SourceStation     = df_TB.loc[:,' ShotStation']
                        TB_Date           = df_TB.loc[:,' ShotUtcDateTime']
                        TB_Time           = df_TB.loc[:,' ShotUtcDateTime']
                        TB_Micro          = df_TB.loc[:,' ShotUtcDateTime']
                                   
                        column_names = [Unit_ID,FileNum,EPNumber,SourceLine,SourceStation,TB_Date,TB_Time,TB_Micro]
                        catdf_TB = pd.concat (column_names,axis=1,ignore_index =True)
                        df_TBList.append(catdf_TB) 
                    else:
                        df_TB             = pd.read_excel(filename)
                        df_TB             = df_TB.iloc[:,:]
                        Unit_ID           = df_TB.loc[:,' ProfileId']
                        FileNum           = df_TB.loc[:,' ShotNumber']
                        EPNumber          = df_TB.loc[:,' EpNumber']
                        SourceLine        = df_TB.loc[:,' ShotLine']
                        SourceStation     = df_TB.loc[:,' ShotStation']
                        TB_Date           = df_TB.loc[:,' ShotUtcDateTime']
                        TB_Time           = df_TB.loc[:,' ShotUtcDateTime']
                        TB_Micro          = df_TB.loc[:,' ShotUtcDateTime']
                        
                        column_names = [Unit_ID,FileNum,EPNumber,SourceLine,SourceStation,TB_Date,TB_Time,TB_Micro]
                        catdf_TB = pd.concat (column_names,axis=1,ignore_index =True)
                        df_TBList.append(catdf_TB) 

                concatdf_TB = pd.concat(df_TBList,axis=0, ignore_index =True)
                concatdf_TB.rename(columns={0:'Unit_ID', 1:'FileNum', 2:'EPNumber', 3:'SourceLine', 4:'SourceStation',
                                         5:'TB_Date', 6:'TB_Time', 7:'TB_Micro'},inplace = True)
                concatdf_TB = concatdf_TB.reset_index(drop=True)

            Valid_TB_df = pd.DataFrame(concatdf_TB)        
            Valid_TB_df= Valid_TB_df[pd.to_numeric(Valid_TB_df.FileNum,errors='coerce').notnull()]        
            Valid_TB_df["SourceLine"].fillna(0, inplace = True)
            Valid_TB_df["SourceStation"].fillna(0, inplace = True)
            Valid_TB_df["EPNumber"].fillna(1, inplace = True)
            Valid_TB_df["Unit_ID"].fillna(0, inplace = True)
            Valid_TB_df['SourceLine']       = (Valid_TB_df.loc[:,['SourceLine']]).astype(int)
            Valid_TB_df['SourceStation']    = (Valid_TB_df.loc[:,['SourceStation']]).astype(float)
            Valid_TB_df['FileNum']          = (Valid_TB_df.loc[:,['FileNum']]).astype(int)            
            Valid_TB_df['EPNumber']         = (Valid_TB_df.loc[:,['EPNumber']]).astype(int)
            Valid_TB_df['Unit_ID']          = (Valid_TB_df.loc[:,['Unit_ID']]).astype(int)
            
            Valid_TB_df['TB_Date']                = pd.to_datetime(Valid_TB_df['TB_Date']).dt.strftime('%Y/%m/%d')
            Valid_TB_df['TB_Time']                = pd.to_datetime(Valid_TB_df['TB_Time']).dt.strftime('%H:%M:%S')
            Valid_TB_df['TB_Micro']               = pd.to_datetime(Valid_TB_df['TB_Micro']).dt.strftime('%f')

            Valid_TB_df                        = Valid_TB_df.reset_index(drop=True)   
            DATA_VALID_TB                      = pd.DataFrame(Valid_TB_df)
            DATA_VALID_TB['DuplicatedEntries'] = DATA_VALID_TB .sort_values(by =['FileNum','SourceLine','SourceStation']).duplicated(['Unit_ID','FileNum','EPNumber','SourceLine','SourceStation'],keep='last')
            DATA_VALID_TB                      = DATA_VALID_TB.loc[DATA_VALID_TB.DuplicatedEntries == False, 'Unit_ID': 'TB_Micro']
            DATA_VALID_TB                      = DATA_VALID_TB.reset_index(drop=True)
            DATA_VALID_TB                      = pd.DataFrame(DATA_VALID_TB)

            Merge_SourceLink_TB_CorruptedPSS = pd.merge(Corrupted_PSS_DF, DATA_VALID_TB, on =['FileNum', 'EPNumber', 'SourceLine', 'SourceStation', 'Unit_ID'] ,how ='left',
               left_index = False, right_index = False, sort = True, indicator = False )
            Merge_SourceLink_TB_CorruptedPSS.drop(['DataBase_ID','TB_Date_x', 'TB_Time_x','TB_Micro_x'], axis=1, inplace=True)


            Merge_SourceLink_TB_CorruptedPSS.rename(columns={'TB_Date_y':'TB_Date', 'TB_Time_y':'TB_Time', 'TB_Micro_y':'TB_Micro'},inplace = True)
            Merge_SourceLink_TB_CorruptedPSS = Merge_SourceLink_TB_CorruptedPSS.loc[:,['ShotID','FileNum','EPNumber','SourceLine','SourceStation',
                                                                                       'Local_Date','Local_Time','Observer_Comment','ShotStatus',
                                                                                       'PhaseMax','PhaseAvg','ForceMax','ForceAvg','THDMax','THDAvg',
                                                                                       'SwCksm','PmCksm','GPS_Quality','Unit_ID','TB_Date','TB_Time','TB_Micro',
                                                                                       'Signature_File_Number','Latitude','Longitude','Altitude',
                                                                                       'Encoder_Index','Record_Index','EP_Count','Crew_ID','Start_Code',
                                                                                       'Force_Out','GPS_Time','GPS_Altitude','Sats','PDOP','HDOP','VDOP',
                                                                                       'Age','Start_Time_Delta','Sweep_Number']]    
            Merge_SourceLink_TB_CorruptedPSS = Merge_SourceLink_TB_CorruptedPSS.reset_index(drop=True)

            DATA_VALID_PSS = pd.DataFrame(Merge_SourceLink_TB_CorruptedPSS)

            ## Getting Void Shot DataFrame
            VIB_Rep_VOID        = DATA_VALID_PSS[(DATA_VALID_PSS.ShotStatus.notnull())]
            VIB_Rep_VOID        = VIB_Rep_VOID.reset_index(drop=True)

            ## Getting Not Void Shot DataFrame
            VIB_Rep_VALID_NOT_VOID    = DATA_VALID_PSS[(DATA_VALID_PSS.ShotStatus.isnull())]
            VIB_Rep_VALID_NOT_VOID    = VIB_Rep_VALID_NOT_VOID[pd.to_numeric(VIB_Rep_VALID_NOT_VOID.PhaseAvg,errors='coerce').notnull()]
            VIB_Rep_VALID_NOT_VOID    = VIB_Rep_VALID_NOT_VOID[pd.to_numeric(VIB_Rep_VALID_NOT_VOID.ForceAvg,errors='coerce').notnull()]
            VIB_Rep_VALID_NOT_VOID    = VIB_Rep_VALID_NOT_VOID[pd.to_numeric(VIB_Rep_VALID_NOT_VOID.THDAvg,errors='coerce').notnull()]
            VIB_Rep_VALID_NOT_VOID    = VIB_Rep_VALID_NOT_VOID[pd.to_numeric(VIB_Rep_VALID_NOT_VOID.THDMax,errors='coerce').notnull()]
            VIB_Rep_VALID_NOT_VOID    = VIB_Rep_VALID_NOT_VOID[pd.to_numeric(VIB_Rep_VALID_NOT_VOID.PhaseMax,errors='coerce').notnull()]
            VIB_Rep_VALID_NOT_VOID    = VIB_Rep_VALID_NOT_VOID.reset_index(drop=True)

            ## Getting QC Passed Shot Dataframe    
            VIB_Rep_QC_Passed   =  pd.DataFrame(VIB_Rep_VALID_NOT_VOID)
            VIB_Rep_QC_Passed   =  VIB_Rep_QC_Passed[(VIB_Rep_QC_Passed.PhaseMax <= PhaseMax_Limit)&
                                                 (VIB_Rep_QC_Passed.PhaseAvg <= High_PhaseAvg_Limit)&
                                                 (VIB_Rep_QC_Passed.PhaseAvg >= Low_PhaseAvg_Limit)&
                                                 (VIB_Rep_QC_Passed.ForceMax <= High_ForceMax_Limit)&
                                                 (VIB_Rep_QC_Passed.ForceMax >= Low_ForceMax_Limit)&                                     
                                                 (VIB_Rep_QC_Passed.ForceAvg <= High_ForceAvg_Limit)&
                                                 (VIB_Rep_QC_Passed.ForceAvg >= Low_ForceAvg_Limit)&
                                                 (VIB_Rep_QC_Passed.THDAvg   <= High_THDAvg_Limit)&
                                                 (VIB_Rep_QC_Passed.THDMax   <= High_THDMax_Limit)&  
                                                 (VIB_Rep_QC_Passed.GPS_Quality != "No Fix")]
            VIB_Rep_QC_Passed   = VIB_Rep_QC_Passed.reset_index(drop=True)
            VIB_Rep_QC_Passed   =  pd.DataFrame(VIB_Rep_QC_Passed)

            ## Getting QC Failed Shot Dataframe    
            VIB_Rep_QC_Failed   =  pd.DataFrame(VIB_Rep_VALID_NOT_VOID)
            VIB_Rep_QC_Failed   =  VIB_Rep_QC_Failed[(VIB_Rep_QC_Failed.PhaseMax > PhaseMax_Limit)|
                                          (VIB_Rep_QC_Failed.ForceAvg > High_ForceAvg_Limit)|
                                          (VIB_Rep_QC_Failed.ForceAvg < Low_ForceAvg_Limit)|
                                          (VIB_Rep_QC_Failed.THDAvg   > High_THDAvg_Limit)|
                                          (VIB_Rep_QC_Failed.THDMax   > High_THDMax_Limit)|
                                          (VIB_Rep_QC_Failed.GPS_Quality == "No Fix")|                              
                                          (VIB_Rep_QC_Failed.PhaseAvg > High_PhaseAvg_Limit)|
                                          (VIB_Rep_QC_Failed.PhaseAvg < Low_PhaseAvg_Limit)|
                                          (VIB_Rep_QC_Failed.ForceMax > High_ForceMax_Limit)|
                                          (VIB_Rep_QC_Failed.ForceMax < Low_ForceMax_Limit)]
            VIB_Rep_QC_Failed  = VIB_Rep_QC_Failed.reset_index(drop=True)
            VIB_Rep_QC_Failed  =  pd.DataFrame(VIB_Rep_QC_Failed)
                
            con = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
            cur=con.cursor()
            Merge_SourceLink_TB_CorruptedPSS.to_sql('Eagle_PSSLog_TEMP',con, if_exists="replace",  index_label='DataBase_ID')                
            VIB_Rep_VOID.to_sql('Eagle_PSSLog_VOID',con, if_exists="replace", index_label='DataBase_ID')
            VIB_Rep_QC_Passed.to_sql('Eagle_PSSLog_QCPassed',con, if_exists="replace", index_label='DataBase_ID')
            VIB_Rep_QC_Failed.to_sql('Eagle_PSSLog_QCFailed',con, if_exists="replace", index_label='DataBase_ID')                
            con.commit()
            cur.close()
            con.close()
            ViewTotalImport()
            txtSingleVoidReport.delete(0,END)

    def ViewTotalImport():
        tree.delete(*tree.get_children())
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedShotID.delete(0,END)
        txtVoidEntries.delete(0,END)
        txtQCPassedEntries.delete(0,END)
        txtQCFailedEntries.delete(0,END)
        txtTotalRAWImport.delete(0,END)
        txtTotalinvalidRemoved.delete(0,END)
        txtDuplicatedShotExchange.delete(0,END)
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_TEMP ORDER BY `ShotID` ASC ;", conn)
        data = pd.DataFrame(Complete_df)
        data = data.reset_index(drop=True)
        for each_rec in range(len(data)):
            tree.insert("", tk.END, values=list(data.loc[each_rec]))
        TotalEntries = len(data)       
        txtTotalEntries.insert(tk.END,TotalEntries)              
        conn.commit()
        conn.close()
        TotalRAW_ImportedEntries()
        TotalInvalidRemoved()
        InvalidEntries()
        DuplicatedShotIDEntries()
        VoidShotIDEntries()
        QCPassedShotIDEntries()
        QCFailedShotIDEntries()

    def ViewInvalidImport():
        UTC_Offset_Hours = (Entrytxt_TimeOffset.get())
        UTC_Offset_Hours = float(UTC_Offset_Hours)
        tree.delete(*tree.get_children())
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedShotID.delete(0,END)
        txtVoidEntries.delete(0,END)
        txtQCPassedEntries.delete(0,END)
        txtQCFailedEntries.delete(0,END)
        txtTotalRAWImport.delete(0,END)
        txtTotalinvalidRemoved.delete(0,END)
        txtDuplicatedShotExchange.delete(0,END)
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_INVALID_NULL ORDER BY `ShotID` ASC ;", conn)        
        try:                    
            Complete_df['TB_Date']                = pd.to_datetime(Complete_df['TB_Date']).dt.strftime('%Y/%m/%d')
            Complete_df['TB_Time']                = pd.to_datetime(Complete_df['TB_Time']).dt.strftime('%H:%M:%S')
            Complete_df['TB_Micro']               = pd.to_datetime(Complete_df['TB_Micro']).dt.strftime('%f')
            
        except:
            try:
                Complete_df['TB_Date']            = pd.to_datetime(Complete_df['Local_Date']).dt.strftime('%Y/%m/%d')
                Complete_df['TB_Time']            = pd.to_datetime(Complete_df['Local_Time']).dt.strftime('%H:%M:%S')
                Complete_df['TB_DateTime']        = pd.to_datetime(Complete_df.TB_Date.astype(str)+' '+Complete_df.TB_Time.astype(str))                                                                        
                Complete_df['TB_DateTime']        = pd.to_datetime(Complete_df['TB_DateTime'].astype(str)) + pd.DateOffset(hours=UTC_Offset_Hours)
                Complete_df['TB_Date']            = pd.to_datetime(Complete_df['TB_DateTime']).dt.strftime('%Y/%m/%d')
                Complete_df['TB_Time']            = pd.to_datetime(Complete_df['TB_DateTime']).dt.strftime('%H:%M:%S')                                                
                Complete_df['TB_Micro']           = 0
                Complete_df.drop(['TB_DateTime'], axis=1, inplace=True)
            except:
                Complete_df['TB_Date']                = (Complete_df['TB_Date'])
                Complete_df['TB_Time']                = (Complete_df['TB_Time'])
                Complete_df['TB_Micro']               = (Complete_df['TB_Micro'])
        
        data = pd.DataFrame(Complete_df)
        data = data.reset_index(drop=True)
        for each_rec in range(len(data)):
            tree.insert("", tk.END, values=list(data.loc[each_rec]))
        TotalInvalidEntries = len(data)       
        txtInvalidEntries.insert(tk.END,TotalInvalidEntries)              
        conn.commit()
        conn.close()
        TotalRAW_ImportedEntries()
        TotalEntries()
        TotalInvalidRemoved()
        DuplicatedShotIDEntries()
        VoidShotIDEntries()
        QCPassedShotIDEntries()
        QCFailedShotIDEntries()

    def ViewTotalRAWImport():
        UTC_Offset_Hours = (Entrytxt_TimeOffset.get())
        UTC_Offset_Hours = float(UTC_Offset_Hours)
        tree.delete(*tree.get_children())
        txtTotalRAWImport.delete(0,END)
        txtTotalinvalidRemoved.delete(0,END)
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedShotID.delete(0,END)
        txtVoidEntries.delete(0,END)
        txtQCPassedEntries.delete(0,END)
        txtQCFailedEntries.delete(0,END)
        txtDuplicatedShotExchange.delete(0,END)
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_RAWDUMP ORDER BY `ShotID` ASC ;", conn)               
        try:                  
            Complete_df['TB_Date']                = pd.to_datetime(Complete_df['TB_Date']).dt.strftime('%Y/%m/%d')
            Complete_df['TB_Time']                = pd.to_datetime(Complete_df['TB_Time']).dt.strftime('%H:%M:%S')
            Complete_df['TB_Micro']               = pd.to_datetime(Complete_df['TB_Micro']).dt.strftime('%f')
        except:
            try:
                Complete_df['TB_Date']            = pd.to_datetime(Complete_df['Local_Date']).dt.strftime('%Y/%m/%d')
                Complete_df['TB_Time']            = pd.to_datetime(Complete_df['Local_Time']).dt.strftime('%H:%M:%S')
                Complete_df['TB_DateTime']        = pd.to_datetime(Complete_df.TB_Date.astype(str)+' '+Complete_df.TB_Time.astype(str))                                                                        
                Complete_df['TB_DateTime']        = pd.to_datetime(Complete_df['TB_DateTime'].astype(str)) + pd.DateOffset(hours=UTC_Offset_Hours)
                Complete_df['TB_Date']            = pd.to_datetime(Complete_df['TB_DateTime']).dt.strftime('%Y/%m/%d')
                Complete_df['TB_Time']            = pd.to_datetime(Complete_df['TB_DateTime']).dt.strftime('%H:%M:%S')                                                
                Complete_df['TB_Micro']           = 0
                Complete_df.drop(['TB_DateTime'], axis=1, inplace=True)
            except:
                Complete_df['TB_Date']                = (Complete_df['TB_Date'])
                Complete_df['TB_Time']                = (Complete_df['TB_Time'])
                Complete_df['TB_Micro']               = (Complete_df['TB_Micro'])
        data = pd.DataFrame(Complete_df)
        data = data.reset_index(drop=True)
        for each_rec in range(len(data)):
            tree.insert("", tk.END, values=list(data.loc[each_rec]))
        TotalRAWEntries = len(data)       
        txtTotalRAWImport.insert(tk.END,TotalRAWEntries)              
        conn.commit()
        conn.close()
        TotalEntries()
        TotalInvalidRemoved()
        DuplicatedShotIDEntries()
        VoidShotIDEntries()
        QCPassedShotIDEntries()
        QCFailedShotIDEntries()
        InvalidEntries()

    def ViewDuplicatedShotIDImport():
        tree.delete(*tree.get_children())
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedShotID.delete(0,END)
        txtVoidEntries.delete(0,END)
        txtQCPassedEntries.delete(0,END)
        txtQCFailedEntries.delete(0,END)
        txtTotalRAWImport.delete(0,END)
        txtTotalinvalidRemoved.delete(0,END)
        txtDuplicatedShotExchange.delete(0,END)
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_DuplicatedShotID ORDER BY `ShotID` ASC ;", conn)
        data = pd.DataFrame(Complete_df)
        data = data.reset_index(drop=True)
        for each_rec in range(len(data)):
            tree.insert("", tk.END, values=list(data.loc[each_rec]))
        TotalDuplicatedEntries = len(data)       
        txtDuplicatedShotID.insert(tk.END,TotalDuplicatedEntries)              
        conn.commit()
        conn.close()
        TotalRAW_ImportedEntries()
        InvalidEntries()
        TotalEntries()
        VoidShotIDEntries()
        QCPassedShotIDEntries()
        QCFailedShotIDEntries()
        TotalInvalidRemoved()

    def ViewAllInvalidRemoved():
        UTC_Offset_Hours = (Entrytxt_TimeOffset.get())
        UTC_Offset_Hours = float(UTC_Offset_Hours)
        tree.delete(*tree.get_children())
        txtTotalRAWImport.delete(0,END)
        txtTotalinvalidRemoved.delete(0,END)
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedShotID.delete(0,END)
        txtVoidEntries.delete(0,END)
        txtQCPassedEntries.delete(0,END)
        txtQCFailedEntries.delete(0,END)
        txtDuplicatedShotExchange.delete(0,END)
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")    
        Complete_df_PSSLog_DuplicatedShotID = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_DuplicatedShotID ORDER BY `ShotID` ASC ;", conn)
        Complete_df_PSSLog_DuplicatedShotID = pd.DataFrame(Complete_df_PSSLog_DuplicatedShotID)
        Complete_df_PSSLog_DuplicatedShotID = Complete_df_PSSLog_DuplicatedShotID.reset_index(drop=True)
        Complete_df_PSSLog_INVALID_NULL= pd.read_sql_query("SELECT * FROM Eagle_PSSLog_INVALID_NULL ORDER BY `ShotID` ASC ;", conn)
        try:                  
            Complete_df_PSSLog_INVALID_NULL['TB_Date']                = pd.to_datetime(Complete_df_PSSLog_INVALID_NULL['TB_Date']).dt.strftime('%Y/%m/%d')
            Complete_df_PSSLog_INVALID_NULL['TB_Time']                = pd.to_datetime(Complete_df_PSSLog_INVALID_NULL['TB_Time']).dt.strftime('%H:%M:%S')
            Complete_df_PSSLog_INVALID_NULL['TB_Micro']               = pd.to_datetime(Complete_df_PSSLog_INVALID_NULL['TB_Micro']).dt.strftime('%f')
        except:
            try:
                Complete_df_PSSLog_INVALID_NULL['TB_Date']            = pd.to_datetime(Complete_df_PSSLog_INVALID_NULL['Local_Date']).dt.strftime('%Y/%m/%d')
                Complete_df_PSSLog_INVALID_NULL['TB_Time']            = pd.to_datetime(Complete_df_PSSLog_INVALID_NULL['Local_Time']).dt.strftime('%H:%M:%S')
                Complete_df_PSSLog_INVALID_NULL['TB_DateTime']        = pd.to_datetime(Complete_df_PSSLog_INVALID_NULL.TB_Date.astype(str)+' '+Complete_df_PSSLog_INVALID_NULL.TB_Time.astype(str))                                                                        
                Complete_df_PSSLog_INVALID_NULL['TB_DateTime']        = pd.to_datetime(Complete_df_PSSLog_INVALID_NULL['TB_DateTime'].astype(str)) + pd.DateOffset(hours=UTC_Offset_Hours)
                Complete_df_PSSLog_INVALID_NULL['TB_Date']            = pd.to_datetime(Complete_df_PSSLog_INVALID_NULL['TB_DateTime']).dt.strftime('%Y/%m/%d')
                Complete_df_PSSLog_INVALID_NULL['TB_Time']            = pd.to_datetime(Complete_df_PSSLog_INVALID_NULL['TB_DateTime']).dt.strftime('%H:%M:%S')                                                
                Complete_df_PSSLog_INVALID_NULL['TB_Micro']           = 0
                Complete_df_PSSLog_INVALID_NULL.drop(['TB_DateTime'], axis=1, inplace=True)
            except:
                Complete_df_PSSLog_INVALID_NULL['TB_Date']                = (Complete_df_PSSLog_INVALID_NULL['TB_Date'])
                Complete_df_PSSLog_INVALID_NULL['TB_Time']                = (Complete_df_PSSLog_INVALID_NULL['TB_Time'])
                Complete_df_PSSLog_INVALID_NULL['TB_Micro']               = (Complete_df_PSSLog_INVALID_NULL['TB_Micro'])
        
        Complete_df_PSSLog_INVALID_NULL = pd.DataFrame(Complete_df_PSSLog_INVALID_NULL)
        Complete_df_PSSLog_INVALID_NULL = Complete_df_PSSLog_INVALID_NULL.reset_index(drop=True)
        Complete_df = Complete_df_PSSLog_DuplicatedShotID.append(Complete_df_PSSLog_INVALID_NULL, ignore_index=True)    
        data = pd.DataFrame(Complete_df)
        data = data.reset_index(drop=True)
        for each_rec in range(len(data)):
            tree.insert("", tk.END, values=list(data.loc[each_rec]))
        TotalinvalidRemoved = len(data)       
        txtTotalinvalidRemoved.insert(tk.END,TotalinvalidRemoved)              
        conn.commit()
        conn.close()
        TotalRAW_ImportedEntries()
        InvalidEntries()
        TotalEntries()
        VoidShotIDEntries()
        QCPassedShotIDEntries()
        QCFailedShotIDEntries()
        DuplicatedShotIDEntries()

    def ViewVoidShotIDImport():
        tree.delete(*tree.get_children())
        txtSingleVoidReport.delete(0,END)
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedShotID.delete(0,END)
        txtVoidEntries.delete(0,END)
        txtQCPassedEntries.delete(0,END)
        txtQCFailedEntries.delete(0,END)
        txtTotalRAWImport.delete(0,END)
        txtTotalinvalidRemoved.delete(0,END)
        txtDuplicatedShotExchange.delete(0,END)
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_VOID ORDER BY `ShotID` ASC ;", conn)
        data = pd.DataFrame(Complete_df)
        data = data.reset_index(drop=True)
        for each_rec in range(len(data)):
            tree.insert("", tk.END, values=list(data.loc[each_rec]))
        TotalVoidEntries = len(data)       
        txtVoidEntries.insert(tk.END,TotalVoidEntries)              
        conn.commit()
        conn.close()
        TotalRAW_ImportedEntries()
        InvalidEntries()
        TotalEntries()
        DuplicatedShotIDEntries()
        QCPassedShotIDEntries()
        QCFailedShotIDEntries()
        TotalInvalidRemoved()


    def ExportAllVoidImport():
        txtSingleVoidReport.delete(0,END)
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
        Complete_ExportAllVoidImport = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_VOID ORDER BY `ShotID` ASC ;", conn)
        data_ExportAllVoidImport = pd.DataFrame(Complete_ExportAllVoidImport)
        data_ExportAllVoidImport.drop(['DataBase_ID'], axis=1, inplace=True)
        data_ExportAllVoidImport.sort_values(by =['SourceLine', 'SourceStation'], inplace =True)
        data_ExportAllVoidImport ['DuplicatedEntries']=data_ExportAllVoidImport.sort_values(by =['SourceLine', 'SourceStation']).duplicated(['SourceLine','SourceStation'],keep='last')
        data_ExportAllVoidImport = data_ExportAllVoidImport.loc[:,['ShotID','FileNum','EPNumber','SourceLine','SourceStation','Local_Date','Local_Time','Observer_Comment',
                                                                   'ShotStatus', 'PhaseMax','PhaseAvg','ForceMax','ForceAvg','THDMax','THDAvg','SwCksm','PmCksm',
                                                                   'GPS_Quality','Unit_ID','TB_Date','TB_Time','TB_Micro','DuplicatedEntries']]    
        data_ExportAllVoidImport = data_ExportAllVoidImport.reset_index(drop=True)

        Complete_df_AllValidImport = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_TEMP ORDER BY `ShotID` ASC ;", conn)
        data_AllValidImport = pd.DataFrame(Complete_df_AllValidImport)
        data_AllValidImport.drop(['DataBase_ID'], axis=1, inplace=True)
        data_AllValidImport    = data_AllValidImport[(data_AllValidImport.ShotStatus.isnull())]
        
        data_AllValidImport.sort_values(by =['SourceLine', 'SourceStation'], inplace =True)
        data_AllValidImport ['DuplicatedEntries']=data_AllValidImport.sort_values(by =['SourceLine', 'SourceStation']).duplicated(['SourceLine','SourceStation'],keep='last')
        data_AllValidImport    = data_AllValidImport [data_AllValidImport.DuplicatedEntries == False]

        data_AllValidImport    =  data_AllValidImport.loc[:,['SourceLine','SourceStation']]
                                                                                                      
        data_AllValidImport = data_AllValidImport.reset_index(drop=True)

        VIB_Rep_Fail_QC_SingleVOID      =  pd.merge(data_ExportAllVoidImport , data_AllValidImport , how='outer', on = ['SourceLine','SourceStation'], indicator=True).query('_merge == "left_only"').drop(columns=['_merge'])
        VIB_Rep_Fail_QC_DuplicatedVOID   =  pd.merge(data_ExportAllVoidImport , data_AllValidImport , how='outer', on = ['SourceLine','SourceStation'], indicator=True).query('_merge == "both"').drop(columns=['_merge'])

        VIB_Rep_Fail_QC_SingleVOID = VIB_Rep_Fail_QC_SingleVOID.loc[:,['ShotID','FileNum','EPNumber','SourceLine','SourceStation','Local_Date','Local_Time','Observer_Comment',
                                                                   'ShotStatus', 'PhaseMax','PhaseAvg','ForceMax','ForceAvg','THDMax','THDAvg','SwCksm','PmCksm',
                                                                   'GPS_Quality','Unit_ID','TB_Date','TB_Time','TB_Micro']]
        VIB_Rep_Fail_QC_DuplicatedVOID = VIB_Rep_Fail_QC_DuplicatedVOID.loc[:,['ShotID','FileNum','EPNumber','SourceLine','SourceStation','Local_Date','Local_Time','Observer_Comment',
                                                                   'ShotStatus', 'PhaseMax','PhaseAvg','ForceMax','ForceAvg','THDMax','THDAvg','SwCksm','PmCksm',
                                                                   'GPS_Quality','Unit_ID','TB_Date','TB_Time','TB_Micro']]          
        conn.commit()
        conn.close()

        def get_QC_VoidReport_datetime():
            return " - ViB Void Shots Report Detail -" + datetime.now().strftime("%Y%m%d-%H%M%S") + ".xlsx"
        filename = tkinter.filedialog.asksaveasfilename(initialdir = "/",title = "Select File Name For Void Detail Report" ,
                                                        filetypes = (("Excel file",".xlsx"),("Excel file",".xlsx")))
        if len(filename) >0:
            QC_VibVOID   = get_QC_VoidReport_datetime()
            outfile_QC_VibVOIDSummary = filename + QC_VibVOID
            XLSX_writer = pd.ExcelWriter(outfile_QC_VibVOIDSummary)
            data_ExportAllVoidImport.to_excel(XLSX_writer, 'VIB_All_VOIDS', index=False)
            VIB_Rep_Fail_QC_DuplicatedVOID.to_excel(XLSX_writer, 'VOID_With_ValidShot', index=False)
            VIB_Rep_Fail_QC_SingleVOID.to_excel(XLSX_writer, 'Single VOID_With_NoValidShot', index=False)        
            XLSX_writer.save()
            XLSX_writer.close()
            tkinter.messagebox.showinfo("QC VOID Detailed PSS Report","QC VOID Detailed PSS Report Saved as Excel")
        else:
            tkinter.messagebox.showinfo("Export QC VOID Detailed PSS Report Message","Please Select File Name To Export")
        

    def GenerateQCSingleVOIDReport():
        tree.delete(*tree.get_children())
        txtSingleVoidReport.delete(0,END)
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
        Complete_ExportAllVoidImport = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_VOID ORDER BY `ShotID` ASC ;", conn)
        data_ExportAllVoidImport = pd.DataFrame(Complete_ExportAllVoidImport)
        data_ExportAllVoidImport.drop(['DataBase_ID'], axis=1, inplace=True)
        data_ExportAllVoidImport.sort_values(by =['SourceLine', 'SourceStation'], inplace =True)
        data_ExportAllVoidImport ['DuplicatedEntries']=data_ExportAllVoidImport.sort_values(by =['SourceLine', 'SourceStation']).duplicated(['SourceLine','SourceStation'],keep='last')
        data_ExportAllVoidImport = data_ExportAllVoidImport.loc[:,['ShotID','FileNum','EPNumber','SourceLine','SourceStation','Local_Date','Local_Time','Observer_Comment',
                                                                   'ShotStatus', 'PhaseMax','PhaseAvg','ForceMax','ForceAvg','THDMax','THDAvg','SwCksm','PmCksm',
                                                                   'GPS_Quality','Unit_ID','TB_Date','TB_Time','TB_Micro','Signature_File_Number',
                                                                     'Latitude','Longitude','Altitude','Encoder_Index','Record_Index',
                                                                     'EP_Count','Crew_ID','Start_Code','Force_Out','GPS_Time',
                                                                     'GPS_Altitude','Sats','PDOP','HDOP','VDOP','Age','Start_Time_Delta',
                                                                     'Sweep_Number','DuplicatedEntries']]    
        data_ExportAllVoidImport = data_ExportAllVoidImport.reset_index(drop=True)
        Complete_df_AllValidImport = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_TEMP ORDER BY `ShotID` ASC ;", conn)
        data_AllValidImport = pd.DataFrame(Complete_df_AllValidImport)
        data_AllValidImport.drop(['DataBase_ID'], axis=1, inplace=True)
        data_AllValidImport    = data_AllValidImport[(data_AllValidImport.ShotStatus.isnull())]        
        data_AllValidImport.sort_values(by =['SourceLine', 'SourceStation'], inplace =True)
        data_AllValidImport ['DuplicatedEntries']=data_AllValidImport.sort_values(by =['SourceLine', 'SourceStation']).duplicated(['SourceLine','SourceStation'],keep='last')
        data_AllValidImport    = data_AllValidImport [data_AllValidImport.DuplicatedEntries == False]
        data_AllValidImport    =  data_AllValidImport.loc[:,['SourceLine','SourceStation']]                                                                                                      
        data_AllValidImport = data_AllValidImport.reset_index(drop=True)
        VIB_Rep_Fail_QC_SingleVOID      =  pd.merge(data_ExportAllVoidImport , data_AllValidImport , how='outer', on = ['SourceLine','SourceStation'], indicator=True).query('_merge == "left_only"').drop(columns=['_merge'])        
        VIB_Rep_Fail_QC_SingleVOID = VIB_Rep_Fail_QC_SingleVOID.loc[:,['ShotID','ShotID','FileNum','EPNumber','SourceLine','SourceStation','Local_Date','Local_Time','Observer_Comment',
                                                                   'ShotStatus', 'PhaseMax','PhaseAvg','ForceMax','ForceAvg','THDMax','THDAvg','SwCksm','PmCksm',
                                                                   'GPS_Quality','Unit_ID','TB_Date','TB_Time','TB_Micro','Signature_File_Number',
                                                                     'Latitude','Longitude','Altitude','Encoder_Index','Record_Index',
                                                                     'EP_Count','Crew_ID','Start_Code','Force_Out','GPS_Time',
                                                                     'GPS_Altitude','Sats','PDOP','HDOP','VDOP','Age','Start_Time_Delta',
                                                                     'Sweep_Number']]
        VIB_Rep_Fail_QC_SingleVOID = pd.DataFrame(VIB_Rep_Fail_QC_SingleVOID)
        TotalSingleVOID = len(VIB_Rep_Fail_QC_SingleVOID)       
        txtSingleVoidReport.insert(tk.END,TotalSingleVOID)
        if TotalSingleVOID > 0:
            try:
                VIB_Rep_Fail_QC_SingleVOID['ShotID'] = (VIB_Rep_Fail_QC_SingleVOID.loc[:,['ShotID']]).astype(int)
            except:
                VIB_Rep_Fail_QC_SingleVOID['ShotID'] = (VIB_Rep_Fail_QC_SingleVOID.loc[:,['ShotID']])

            try:
                VIB_Rep_Fail_QC_SingleVOID['FileNum']= (VIB_Rep_Fail_QC_SingleVOID.loc[:,['FileNum']]).astype(int)
            except:
                VIB_Rep_Fail_QC_SingleVOID['FileNum']= (VIB_Rep_Fail_QC_SingleVOID.loc[:,['FileNum']])

            try:
                VIB_Rep_Fail_QC_SingleVOID['EPNumber']= (VIB_Rep_Fail_QC_SingleVOID.loc[:,['EPNumber']]).astype(int)
            except:
                VIB_Rep_Fail_QC_SingleVOID['EPNumber']= (VIB_Rep_Fail_QC_SingleVOID.loc[:,['EPNumber']])

            try:
                VIB_Rep_Fail_QC_SingleVOID['Unit_ID']= (VIB_Rep_Fail_QC_SingleVOID.loc[:,['Unit_ID']]).astype(int)
            except:
                VIB_Rep_Fail_QC_SingleVOID['Unit_ID']= (VIB_Rep_Fail_QC_SingleVOID.loc[:,['Unit_ID']])

            try:
                VIB_Rep_Fail_QC_SingleVOID['Signature_File_Number'] = (VIB_Rep_Fail_QC_SingleVOID.loc[:,['Signature_File_Number']]).astype(int)
            except:
                VIB_Rep_Fail_QC_SingleVOID['Signature_File_Number'] = (VIB_Rep_Fail_QC_SingleVOID.loc[:,['Signature_File_Number']])

            try:
                VIB_Rep_Fail_QC_SingleVOID['Encoder_Index']= (VIB_Rep_Fail_QC_SingleVOID.loc[:,['Encoder_Index']]).astype(int)
            except:
                VIB_Rep_Fail_QC_SingleVOID['Encoder_Index']= (VIB_Rep_Fail_QC_SingleVOID.loc[:,['Encoder_Index']])

            try:
                VIB_Rep_Fail_QC_SingleVOID['Record_Index']= (VIB_Rep_Fail_QC_SingleVOID.loc[:,['Record_Index']]).astype(int)
            except:
                VIB_Rep_Fail_QC_SingleVOID['Record_Index']= (VIB_Rep_Fail_QC_SingleVOID.loc[:,['Record_Index']])

            try:
                VIB_Rep_Fail_QC_SingleVOID['EP_Count']= (VIB_Rep_Fail_QC_SingleVOID.loc[:,['EP_Count']]).astype(int)
            except:
                VIB_Rep_Fail_QC_SingleVOID['EP_Count']= (VIB_Rep_Fail_QC_SingleVOID.loc[:,['EP_Count']])

            try:
                VIB_Rep_Fail_QC_SingleVOID['Crew_ID']= (VIB_Rep_Fail_QC_SingleVOID.loc[:,['Crew_ID']]).astype(int)
            except:
                VIB_Rep_Fail_QC_SingleVOID['Crew_ID']= (VIB_Rep_Fail_QC_SingleVOID.loc[:,['Crew_ID']])

            try:
                VIB_Rep_Fail_QC_SingleVOID['Start_Code']= (VIB_Rep_Fail_QC_SingleVOID.loc[:,['Start_Code']]).astype(int)
            except:
                VIB_Rep_Fail_QC_SingleVOID['Start_Code']= (VIB_Rep_Fail_QC_SingleVOID.loc[:,['Start_Code']])

            try:
                VIB_Rep_Fail_QC_SingleVOID['Sats']= (VIB_Rep_Fail_QC_SingleVOID.loc[:,['Sats']]).astype(int)
            except:
                VIB_Rep_Fail_QC_SingleVOID['Sats']= (VIB_Rep_Fail_QC_SingleVOID.loc[:,['Sats']])

            try:
                VIB_Rep_Fail_QC_SingleVOID['Sweep_Number']= (VIB_Rep_Fail_QC_SingleVOID.loc[:,['Sweep_Number']]).astype(int)  
            except:
                VIB_Rep_Fail_QC_SingleVOID['Sweep_Number']= (VIB_Rep_Fail_QC_SingleVOID.loc[:,['Sweep_Number']]) 
                              
            VIB_Rep_Fail_QC_SingleVOID = VIB_Rep_Fail_QC_SingleVOID.reset_index(drop=True)                
            for each_rec in range(len(VIB_Rep_Fail_QC_SingleVOID)):
                tree.insert("", tk.END, values=list(VIB_Rep_Fail_QC_SingleVOID.loc[each_rec]))
        else:
            VIB_Rep_Fail_QC_SingleVOID = VIB_Rep_Fail_QC_SingleVOID.reset_index(drop=True)        
            for each_rec in range(len(VIB_Rep_Fail_QC_SingleVOID)):
                tree.insert("", tk.END, values=list(VIB_Rep_Fail_QC_SingleVOID.loc[each_rec]))
        
        conn.commit()
        conn.close()

    def ViewQCPassedImport():
        tree.delete(*tree.get_children())
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedShotID.delete(0,END)
        txtVoidEntries.delete(0,END)
        txtQCPassedEntries.delete(0,END)
        txtQCFailedEntries.delete(0,END)
        txtTotalRAWImport.delete(0,END)
        txtTotalinvalidRemoved.delete(0,END)
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_QCPassed ORDER BY `ShotID` ASC ;", conn)
        data = pd.DataFrame(Complete_df)
        data = data.reset_index(drop=True)
        for each_rec in range(len(data)):
            tree.insert("", tk.END, values=list(data.loc[each_rec]))
        TotalQCPassedEntries = len(data)       
        txtQCPassedEntries.insert(tk.END,TotalQCPassedEntries)              
        conn.commit()
        conn.close()
        TotalRAW_ImportedEntries()
        InvalidEntries()
        TotalEntries()
        DuplicatedShotIDEntries()
        VoidShotIDEntries()
        QCFailedShotIDEntries()
        TotalInvalidRemoved()

    def ViewQCFailedImport():
        tree.delete(*tree.get_children())
        txtSingleFailedReport.delete(0,END)
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedShotID.delete(0,END)
        txtVoidEntries.delete(0,END)
        txtQCPassedEntries.delete(0,END)
        txtQCFailedEntries.delete(0,END)
        txtTotalRAWImport.delete(0,END)
        txtTotalinvalidRemoved.delete(0,END)
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_QCFailed ORDER BY `ShotID` ASC ;", conn)
        data = pd.DataFrame(Complete_df)
        data = data.reset_index(drop=True)
        for each_rec in range(len(data)):
            tree.insert("", tk.END, values=list(data.loc[each_rec]))
        TotalQCFailedEntries = len(data)       
        txtQCFailedEntries.insert(tk.END,TotalQCFailedEntries)              
        conn.commit()
        conn.close()
        TotalRAW_ImportedEntries()
        InvalidEntries()
        TotalEntries()
        DuplicatedShotIDEntries()
        VoidShotIDEntries()
        QCPassedShotIDEntries()
        TotalInvalidRemoved()

    def GenerateQCPassedReport():
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
        Complete_df_QC_Passed = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_QCPassed ORDER BY `ShotID` ASC ;", conn)
        VIB_Rep_QC_Passed  = pd.DataFrame(Complete_df_QC_Passed)        
        VIB_Rep_QC_Passed.drop(['DataBase_ID'], axis=1, inplace=True)        
        VIB_Rep_QC_Passed  = VIB_Rep_QC_Passed.reset_index(drop=True)
        VIB_Rep_Passed_QC  =  pd.DataFrame(VIB_Rep_QC_Passed)
        conn.commit()
        conn.close()

        def get_QC_PassedReport_datetime():
            return " - ViB QC Passed Report Detail -" + datetime.now().strftime("%Y%m%d-%H%M%S") + ".xlsx"
        filename = tkinter.filedialog.asksaveasfilename(initialdir = "/",title = "Select File Name For ViB QC Passed Detail Report" ,
                                                        filetypes = (("Excel file",".xlsx"),("Excel file",".xlsx")))
        if len(filename) >0:
            QC_VibPass   = get_QC_PassedReport_datetime()
            outfile_QC_VibPassSummary = filename + QC_VibPass
            XLSX_writer = pd.ExcelWriter(outfile_QC_VibPassSummary)
            VIB_Rep_Passed_QC.to_excel(XLSX_writer, 'VIBPassedQC', index=False)
            XLSX_writer.save()
            XLSX_writer.close()
            tkinter.messagebox.showinfo("QC Passed Detailed PSS Report","QC Passed Detailed PSS Report Saved as Excel")
        else:
            tkinter.messagebox.showinfo("Export QC Passed Detailed PSS Report Message","Please Select File Name To Export")

    def GenerateQCFailedReport():
        txtSingleFailedReport.delete(0,END)
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
        Complete_df_QC_Passed = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_QCPassed ORDER BY `ShotID` ASC ;", conn)
        Complete_df_QC_Failed = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_QCFailed ORDER BY `ShotID` ASC ;", conn)

        VIB_Rep_QC_Passed  = pd.DataFrame(Complete_df_QC_Passed)
        VIB_Rep_QC_Passed  = VIB_Rep_QC_Passed.reset_index(drop=True)
        VIB_Rep_Passed_QC  =  pd.DataFrame(VIB_Rep_QC_Passed)

        VIB_Rep_QC_Failed  = pd.DataFrame(Complete_df_QC_Failed)
        VIB_Rep_QC_Failed  = VIB_Rep_QC_Failed.reset_index(drop=True)
        VIB_Rep_Fail_QC    =  pd.DataFrame(VIB_Rep_QC_Failed)
        conn.commit()
        conn.close()

        def trans_PhaseMax_Limit(x):
                if x > PhaseMax_Limit:
                    return 'PhaseMax_Limit_Fail'
                else:
                    return '(PhaseMax_OK)'

        def trans_PhaseAvg_Limit(q):
            if q > High_PhaseAvg_Limit:
                return 'High_PhaseAvg_Limit_Fail'
            elif q < Low_PhaseAvg_Limit:
                return 'Low_PhaseAvg_Limit_Fail'
            else:
                return '(PhaseAvg_OK)'

        def trans_ForceMax_Limit(p):
            if p > High_ForceMax_Limit:
                return 'High_ForceMax_Limit_Fail'
            elif p < Low_ForceMax_Limit:
                return 'Low_ForceMax_Limit_Fail'
            else:
                return '(ForceMax_OK)'

        def trans_ForceAvg_Limit(y):
            if y > High_ForceAvg_Limit:
                return 'High_ForceAvg_Limit_Fail'
            elif y < Low_ForceAvg_Limit:
                return 'Low_ForceAvg_Limit_Fail'
            else:
                return '(ForceAvg_OK)'

        def trans_THDAvg_Limit(z):
            if z > High_THDAvg_Limit:
                return 'High_THDAvg_Limit_Fail'
            else:
                return '(THDAvg_OK)'  

        def trans_THDMax_Limit(u):
            if u > High_THDMax_Limit:
                return 'High_THDMax_Limit_Fail'
            else:
                return '(THDMax_OK)'  

        def trans_GPS_Quality(w):
            if w == "No Fix":
                return 'GPS_NoFIX'
            else:
                return '(GPS_OK)'

        VIB_Rep_Fail_QC['PhaseMax_Limit_FLAG']  = VIB_Rep_Fail_QC['PhaseMax'].apply(trans_PhaseMax_Limit)
        VIB_Rep_Fail_QC['PhaseMax_Limit_FLAG']  = VIB_Rep_Fail_QC.PhaseMax_Limit_FLAG.astype (object)

        VIB_Rep_Fail_QC['PhaseAvg_Limit_FLAG']  = VIB_Rep_Fail_QC['PhaseAvg'].apply(trans_PhaseAvg_Limit)
        VIB_Rep_Fail_QC['PhaseAvg_Limit_FLAG']  = VIB_Rep_Fail_QC.PhaseAvg_Limit_FLAG.astype (object)

        VIB_Rep_Fail_QC['ForceMax_Limit_FLAG']  = VIB_Rep_Fail_QC['ForceMax'].apply(trans_ForceMax_Limit)
        VIB_Rep_Fail_QC['ForceMax_Limit_FLAG']  = VIB_Rep_Fail_QC.ForceMax_Limit_FLAG.astype (object)

        VIB_Rep_Fail_QC['ForceAvg_Limit_FLAG']  = VIB_Rep_Fail_QC['ForceAvg'].apply(trans_ForceAvg_Limit)
        VIB_Rep_Fail_QC['ForceAvg_Limit_FLAG']  = VIB_Rep_Fail_QC.ForceAvg_Limit_FLAG.astype (object)

        VIB_Rep_Fail_QC['THDAvg_Limit_FLAG']    = VIB_Rep_Fail_QC['THDAvg'].apply(trans_THDAvg_Limit)
        VIB_Rep_Fail_QC['THDAvg_Limit_FLAG']    = VIB_Rep_Fail_QC.THDAvg_Limit_FLAG.astype (object)

        VIB_Rep_Fail_QC['THDMax_Limit_FLAG']    = VIB_Rep_Fail_QC['THDMax'].apply(trans_THDMax_Limit)
        VIB_Rep_Fail_QC['THDMax_Limit_FLAG']    = VIB_Rep_Fail_QC.THDMax_Limit_FLAG.astype (object)

        VIB_Rep_Fail_QC['GPS_FLAG']             = VIB_Rep_Fail_QC['GPS_Quality'].apply(trans_GPS_Quality)
        VIB_Rep_Fail_QC['GPS_FLAG']             = VIB_Rep_Fail_QC.GPS_FLAG.astype (object)

        VIB_Rep_Fail_QC    =  pd.DataFrame(VIB_Rep_Fail_QC)
        VIB_Rep_Fail_QC    =  VIB_Rep_Fail_QC.loc[:,['ShotID','FileNum','ShotStatus','EPNumber','SourceLine','SourceStation',
                                                     'Local_Date','Local_Time','TB_Date','TB_Time','TB_Micro',
                                                     'Unit_ID','Observer_Comment','SwCksm','PmCksm',
                                                     'PhaseMax','PhaseAvg', 'ForceMax','ForceAvg','THDMax','THDAvg','GPS_Quality',
                                                     'PhaseMax_Limit_FLAG', 'PhaseAvg_Limit_FLAG',
                                                     'ForceMax_Limit_FLAG', 'ForceAvg_Limit_FLAG',
                                                     'THDMax_Limit_FLAG',   'THDAvg_Limit_FLAG',
                                                     'GPS_FLAG']]
        VIB_Rep_Fail_QC                 =  VIB_Rep_Fail_QC.reset_index(drop=True)
        VIB_Rep_Fail_QC_SingleFail      =  pd.merge(VIB_Rep_Fail_QC , VIB_Rep_Passed_QC , how='outer', on = ['SourceLine','SourceStation'], indicator=True).query('_merge == "left_only"').drop(columns=['_merge'])
        VIB_Rep_Fail_QC_DuplicatedFail  =  pd.merge(VIB_Rep_Fail_QC , VIB_Rep_Passed_QC , how='outer', on = ['SourceLine','SourceStation'], indicator=True).query('_merge == "both"').drop(columns=['_merge'])

        VIB_Rep_Fail_QC_SingleFail      =  VIB_Rep_Fail_QC_SingleFail.loc[:,['ShotID_x','FileNum_x','ShotStatus_x','EPNumber_x','SourceLine','SourceStation',
                                                     'Local_Date_x','Local_Time_x','TB_Date_x','TB_Time_x','TB_Micro_x',
                                                     'Unit_ID_x','Observer_Comment_x','SwCksm_x','PmCksm_x',
                                                     'PhaseMax_x','PhaseAvg_x','ForceMax_x','ForceAvg_x','THDMax_x','THDAvg_x','GPS_Quality_x',
                                                     'PhaseMax_Limit_FLAG', 'PhaseAvg_Limit_FLAG',
                                                     'ForceMax_Limit_FLAG', 'ForceAvg_Limit_FLAG',
                                                     'THDMax_Limit_FLAG',   'THDAvg_Limit_FLAG',
                                                     'GPS_FLAG']]    
        VIB_Rep_Fail_QC_SingleFail      =  VIB_Rep_Fail_QC_SingleFail.reset_index(drop=True)
        
        VIB_Rep_Fail_QC_DuplicatedFail  =  VIB_Rep_Fail_QC_DuplicatedFail.loc[:,['ShotID_x','FileNum_x','ShotStatus_x','EPNumber_x','SourceLine','SourceStation',
                                                     'Local_Date_x','Local_Time_x','TB_Date_x','TB_Time_x','TB_Micro_x',
                                                     'Unit_ID_x','Observer_Comment_x','SwCksm_x','PmCksm_x',
                                                     'PhaseMax_x','PhaseAvg_x','ForceMax_x','ForceAvg_x','THDMax_x','THDAvg_x','GPS_Quality_x',
                                                     'PhaseMax_Limit_FLAG', 'PhaseAvg_Limit_FLAG',
                                                     'ForceMax_Limit_FLAG', 'ForceAvg_Limit_FLAG',
                                                     'THDMax_Limit_FLAG',   'THDAvg_Limit_FLAG',
                                                     'GPS_FLAG']]
        VIB_Rep_Fail_QC_DuplicatedFail  =  VIB_Rep_Fail_QC_DuplicatedFail.reset_index(drop=True)

        VIB_Rep_Fail_QC_SingleFail.rename(columns={'ShotID_x':'ShotID', 'FileNum_x':'FileNum','ShotStatus_x':'ShotStatus',
                                                   'EPNumber_x':'EPNumber', 'SourceLine':'SourceLine', 'SourceStation':'SourceStation',
                                                   'Local_Date_x':'Local_Date','Local_Time_x':'Local_Time','TB_Date_x':'TB_Date',
                                                   'TB_Time_x':'TB_Time','TB_Micro_x':'TB_Micro',
                                                   'Unit_ID_x':'Unit_ID','Observer_Comment_x':'Observer_Comment','SwCksm_x':'SwCksm','PmCksm_x':'PmCksm',
                                                   'PhaseMax_x':'PhaseMax','PhaseAvg_x':'PhaseAvg','ForceMax_x':'ForceMax',
                                                   'ForceAvg_x':'ForceAvg','THDMax_x':'THDMax','THDAvg_x':'THDAvg',
                                                   'GPS_Quality_x':'GPS_Quality', 'PhaseMax_Limit_FLAG':'PhaseMax_Limit_FLAG',
                                                   'PhaseAvg_Limit_FLAG':'PhaseAvg_Limit_FLAG','ForceMax_Limit_FLAG':'ForceMax_Limit_FLAG',
                                                   'ForceAvg_Limit_FLAG':'ForceAvg_Limit_FLAG', 'THDMax_Limit_FLAG':'THDMax_Limit_FLAG',
                                                   'THDAvg_Limit_FLAG':'THDAvg_Limit_FLAG',
                                                   'GPS_FLAG':'GPS_FLAG'},inplace = True)

        VIB_Rep_Fail_QC_DuplicatedFail.rename(columns={'ShotID_x':'ShotID', 'FileNum_x':'FileNum','ShotStatus_x':'ShotStatus',
                                                   'EPNumber_x':'EPNumber', 'SourceLine':'SourceLine', 'SourceStation':'SourceStation',
                                                   'Local_Date_x':'Local_Date','Local_Time_x':'Local_Time','TB_Date_x':'TB_Date',
                                                   'TB_Time_x':'TB_Time','TB_Micro_x':'TB_Micro',
                                                   'Unit_ID_x':'Unit_ID','Observer_Comment_x':'Observer_Comment','SwCksm_x':'SwCksm','PmCksm_x':'PmCksm',
                                                   'PhaseMax_x':'PhaseMax','PhaseAvg_x':'PhaseAvg','ForceMax_x':'ForceMax',
                                                   'ForceAvg_x':'ForceAvg','THDMax_x':'THDMax','THDAvg_x':'THDAvg',
                                                   'GPS_Quality_x':'GPS_Quality', 'PhaseMax_Limit_FLAG':'PhaseMax_Limit_FLAG',
                                                   'PhaseAvg_Limit_FLAG':'PhaseAvg_Limit_FLAG','ForceMax_Limit_FLAG':'ForceMax_Limit_FLAG',
                                                   'ForceAvg_Limit_FLAG':'ForceAvg_Limit_FLAG','THDMax_Limit_FLAG':'THDMax_Limit_FLAG',
                                                   'THDAvg_Limit_FLAG':'THDAvg_Limit_FLAG',
                                                   'GPS_FLAG':'GPS_FLAG'},inplace = True)

        def get_QC_FailReport_datetime():
            return " - ViB QC Fail Report Detail -" + datetime.now().strftime("%Y%m%d-%H%M%S") + ".xlsx"
        filename = tkinter.filedialog.asksaveasfilename(initialdir = "/",title = "Select File Name For ViB QC Fail Detail Report" ,
                                                        filetypes = (("Excel file",".xlsx"),("Excel file",".xlsx")))
        if len(filename) >0:
            QC_VibFail   = get_QC_FailReport_datetime()
            outfile_QC_VibFailSummary = filename + QC_VibFail
            XLSX_writer = pd.ExcelWriter(outfile_QC_VibFailSummary)
            VIB_Rep_Fail_QC.to_excel(XLSX_writer, 'VIBFailQC', index=False)
            VIB_Rep_Fail_QC_SingleFail.to_excel(XLSX_writer, 'VIBFailQC_Without_QCPassedShot', index=False)
            VIB_Rep_Fail_QC_DuplicatedFail.to_excel(XLSX_writer, 'VIBFailQC_With_QCPassedShot', index=False)
            XLSX_writer.save()
            XLSX_writer.close()
            tkinter.messagebox.showinfo("QC Failed Detailed PSS Report","QC Failed Detailed PSS Report Saved as Excel")
        else:
            tkinter.messagebox.showinfo("Export QC Failed Detailed PSS Report Message","Please Select File Name To Export")
        
    def GenerateQCSingleFailedReport():
        tree.delete(*tree.get_children())
        txtSingleFailedReport.delete(0,END)
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
        Complete_df_QC_Passed = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_QCPassed ORDER BY `ShotID` ASC ;", conn)
        Complete_df_QC_Failed = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_QCFailed ORDER BY `ShotID` ASC ;", conn)

        VIB_Rep_QC_Passed  = pd.DataFrame(Complete_df_QC_Passed)
        VIB_Rep_QC_Passed  = VIB_Rep_QC_Passed.reset_index(drop=True)
        VIB_Rep_Passed_QC  =  pd.DataFrame(VIB_Rep_QC_Passed)

        VIB_Rep_QC_Failed  = pd.DataFrame(Complete_df_QC_Failed)
        VIB_Rep_QC_Failed  = VIB_Rep_QC_Failed.reset_index(drop=True)
        VIB_Rep_Fail_QC    =  pd.DataFrame(VIB_Rep_QC_Failed)
        conn.commit()
        conn.close()

        def trans_PhaseMax_Limit(x):
                if x > PhaseMax_Limit:
                    return 'PhaseMax_Limit_Fail'
                else:
                    return '(PhaseMax_OK)'

        def trans_PhaseAvg_Limit(q):
            if q > High_PhaseAvg_Limit:
                return 'High_PhaseAvg_Limit_Fail'
            elif q < Low_PhaseAvg_Limit:
                return 'Low_PhaseAvg_Limit_Fail'
            else:
                return '(PhaseAvg_OK)'

        def trans_ForceMax_Limit(p):
            if p > High_ForceMax_Limit:
                return 'High_ForceMax_Limit_Fail'
            elif p < Low_ForceMax_Limit:
                return 'Low_ForceMax_Limit_Fail'
            else:
                return '(ForceMax_OK)'

        def trans_ForceAvg_Limit(y):
            if y > High_ForceAvg_Limit:
                return 'High_ForceAvg_Limit_Fail'
            elif y < Low_ForceAvg_Limit:
                return 'Low_ForceAvg_Limit_Fail'
            else:
                return '(ForceAvg_OK)'

        def trans_THDAvg_Limit(z):
            if z > High_THDAvg_Limit:
                return 'High_THDAvg_Limit_Fail'
            else:
                return '(THDAvg_OK)'  

        def trans_THDMax_Limit(u):
            if u > High_THDMax_Limit:
                return 'High_THDMax_Limit_Fail'
            else:
                return '(THDMax_OK)'  

        def trans_GPS_Quality(w):
            if w == "No Fix":
                return 'GPS_NoFIX'
            else:
                return '(GPS_OK)'

        VIB_Rep_Fail_QC['PhaseMax_Limit_FLAG']  = VIB_Rep_Fail_QC['PhaseMax'].apply(trans_PhaseMax_Limit)
        VIB_Rep_Fail_QC['PhaseMax_Limit_FLAG']  = VIB_Rep_Fail_QC.PhaseMax_Limit_FLAG.astype (object)

        VIB_Rep_Fail_QC['PhaseAvg_Limit_FLAG']  = VIB_Rep_Fail_QC['PhaseAvg'].apply(trans_PhaseAvg_Limit)
        VIB_Rep_Fail_QC['PhaseAvg_Limit_FLAG']  = VIB_Rep_Fail_QC.PhaseAvg_Limit_FLAG.astype (object)

        VIB_Rep_Fail_QC['ForceMax_Limit_FLAG']  = VIB_Rep_Fail_QC['ForceMax'].apply(trans_ForceMax_Limit)
        VIB_Rep_Fail_QC['ForceMax_Limit_FLAG']  = VIB_Rep_Fail_QC.ForceMax_Limit_FLAG.astype (object)

        VIB_Rep_Fail_QC['ForceAvg_Limit_FLAG']  = VIB_Rep_Fail_QC['ForceAvg'].apply(trans_ForceAvg_Limit)
        VIB_Rep_Fail_QC['ForceAvg_Limit_FLAG']  = VIB_Rep_Fail_QC.ForceAvg_Limit_FLAG.astype (object)

        VIB_Rep_Fail_QC['THDAvg_Limit_FLAG']    = VIB_Rep_Fail_QC['THDAvg'].apply(trans_THDAvg_Limit)
        VIB_Rep_Fail_QC['THDAvg_Limit_FLAG']    = VIB_Rep_Fail_QC.THDAvg_Limit_FLAG.astype (object)

        VIB_Rep_Fail_QC['THDMax_Limit_FLAG']    = VIB_Rep_Fail_QC['THDMax'].apply(trans_THDMax_Limit)
        VIB_Rep_Fail_QC['THDMax_Limit_FLAG']    = VIB_Rep_Fail_QC.THDMax_Limit_FLAG.astype (object)

        VIB_Rep_Fail_QC['GPS_FLAG']             = VIB_Rep_Fail_QC['GPS_Quality'].apply(trans_GPS_Quality)
        VIB_Rep_Fail_QC['GPS_FLAG']             = VIB_Rep_Fail_QC.GPS_FLAG.astype (object)

        VIB_Rep_Fail_QC    =  pd.DataFrame(VIB_Rep_Fail_QC)
        VIB_Rep_Fail_QC    =  VIB_Rep_Fail_QC.loc[:,['ShotID','FileNum','ShotStatus','EPNumber','SourceLine','SourceStation',
                                                     'Local_Date','Local_Time','TB_Date','TB_Time','TB_Micro',
                                                     'Unit_ID','Observer_Comment','SwCksm','PmCksm',
                                                     'PhaseMax','PhaseAvg','ForceMax','ForceAvg','THDMax','THDAvg','GPS_Quality',
                                                     'PhaseMax_Limit_FLAG', 'PhaseAvg_Limit_FLAG',
                                                     'ForceMax_Limit_FLAG', 'ForceAvg_Limit_FLAG',
                                                     'THDAvg_Limit_FLAG', 'THDMax_Limit_FLAG', 'GPS_FLAG']]
        VIB_Rep_Fail_QC                 =  VIB_Rep_Fail_QC.reset_index(drop=True)
        VIB_Rep_Fail_QC_SingleFail      =  pd.merge(VIB_Rep_Fail_QC , VIB_Rep_Passed_QC , how='outer', on = ['SourceLine','SourceStation'], indicator=True).query('_merge == "left_only"').drop(columns=['_merge'])
        VIB_Rep_Fail_QC_DuplicatedFail  =  pd.merge(VIB_Rep_Fail_QC , VIB_Rep_Passed_QC , how='outer', on = ['SourceLine','SourceStation'], indicator=True).query('_merge == "both"').drop(columns=['_merge'])

        VIB_Rep_Fail_QC_SingleFail      =  VIB_Rep_Fail_QC_SingleFail.loc[:,['ShotID_x','FileNum_x','ShotStatus_x','EPNumber_x','SourceLine','SourceStation',
                                                     'Local_Date_x','Local_Time_x','TB_Date_x','TB_Time_x','TB_Micro_x',
                                                     'Unit_ID_x','Observer_Comment_x','SwCksm_x','PmCksm_x',
                                                     'PhaseMax_x','PhaseAvg_x','ForceMax_x','ForceAvg_x','THDMax_x','THDAvg_x','GPS_Quality_x',
                                                     'PhaseMax_Limit_FLAG', 'PhaseAvg_Limit_FLAG',
                                                     'ForceMax_Limit_FLAG', 'ForceAvg_Limit_FLAG',
                                                     'THDAvg_Limit_FLAG', 'THDMax_Limit_FLAG', 'GPS_FLAG']]    
        VIB_Rep_Fail_QC_SingleFail      =  VIB_Rep_Fail_QC_SingleFail.reset_index(drop=True)
        
        
        VIB_Rep_Fail_QC_SingleFail.rename(columns={'ShotID_x':'ShotID', 'FileNum_x':'FileNum','ShotStatus_x':'ShotStatus',
                                                   'EPNumber_x':'EPNumber', 'SourceLine':'SourceLine', 'SourceStation':'SourceStation',
                                                   'Local_Date_x':'Local_Date','Local_Time_x':'Local_Time','TB_Date_x':'TB_Date',
                                                   'TB_Time_x':'TB_Time','TB_Micro_x':'TB_Micro',
                                                   'Unit_ID_x':'Unit_ID','Observer_Comment_x':'Observer_Comment','SwCksm_x':'SwCksm','PmCksm_x':'PmCksm',
                                                   'PhaseMax_x':'PhaseMax','PhaseAvg_x':'PhaseAvg','ForceMax_x':'ForceMax',
                                                   'ForceAvg_x':'ForceAvg','THDMax_x':'THDMax','THDAvg_x':'THDAvg',
                                                   'GPS_Quality_x':'GPS_Quality', 'PhaseMax_Limit_FLAG':'PhaseMax_Limit_FLAG',
                                                   'PhaseAvg_Limit_FLAG':'PhaseAvg_Limit_FLAG','ForceMax_Limit_FLAG':'ForceMax_Limit_FLAG',
                                                   'ForceAvg_Limit_FLAG':'ForceAvg_Limit_FLAG','THDAvg_Limit_FLAG':'THDAvg_Limit_FLAG',
                                                   'THDMax_Limit_FLAG':'THDMax_Limit_FLAG', 'GPS_FLAG':'GPS_FLAG'},inplace = True)

        VIB_Rep_Fail_QC_SingleFail    =  VIB_Rep_Fail_QC_SingleFail.loc[:,['ShotID', 'ShotID','FileNum',
                                                                            'EPNumber', 'SourceLine','SourceStation',
                                                                            'Local_Date','Local_Time', 'Observer_Comment','ShotStatus',
                                                                            'PhaseMax','PhaseAvg','ForceMax','ForceAvg','THDMax','THDAvg',
                                                                            'SwCksm','PmCksm','GPS_Quality','Unit_ID',
                                                                            'TB_Date','TB_Time','TB_Micro']]
        data_VIB_Rep_Fail_QC_SingleFail = pd.DataFrame(VIB_Rep_Fail_QC_SingleFail)
        TotalQC_SingleFail= len(data_VIB_Rep_Fail_QC_SingleFail)       
        txtSingleFailedReport.insert(tk.END,TotalQC_SingleFail)
        if TotalQC_SingleFail >0:    
            data_VIB_Rep_Fail_QC_SingleFail['ShotID'] = (data_VIB_Rep_Fail_QC_SingleFail.loc[:,['ShotID']]).astype(int)
            data_VIB_Rep_Fail_QC_SingleFail['FileNum']= (data_VIB_Rep_Fail_QC_SingleFail.loc[:,['FileNum']]).astype(int)
            data_VIB_Rep_Fail_QC_SingleFail['EPNumber']= (data_VIB_Rep_Fail_QC_SingleFail.loc[:,['EPNumber']]).astype(int)
            data_VIB_Rep_Fail_QC_SingleFail['Unit_ID']= (data_VIB_Rep_Fail_QC_SingleFail.loc[:,['Unit_ID']]).astype(int)            
            data_VIB_Rep_Fail_QC_SingleFail = data_VIB_Rep_Fail_QC_SingleFail.reset_index(drop=True)    
            for each_rec in range(len(data_VIB_Rep_Fail_QC_SingleFail)):
                tree.insert("", tk.END, values=list(data_VIB_Rep_Fail_QC_SingleFail.loc[each_rec]))
        else:
            data_VIB_Rep_Fail_QC_SingleFail = data_VIB_Rep_Fail_QC_SingleFail.reset_index(drop=True)    
            for each_rec in range(len(data_VIB_Rep_Fail_QC_SingleFail)):
                tree.insert("", tk.END, values=list(data_VIB_Rep_Fail_QC_SingleFail.loc[each_rec]))
            
    def ClearView():
        txtTotalRAWImport.delete(0,END)
        txtTotalEntries.delete(0,END)
        txtTotalinvalidRemoved.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedShotID.delete(0,END)
        txtVoidEntries.delete(0,END)
        txtQCPassedEntries.delete(0,END)
        txtQCFailedEntries.delete(0,END)
        txtSingleVoidReport.delete(0,END)
        txtSingleFailedReport.delete(0,END)
        txtDuplicatedShotExchange.delete(0,END)
        tree.delete(*tree.get_children())

    def TotalEntries():
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_TEMP ;", conn)
        data = pd.DataFrame(Complete_df)
        TotalEntries = len(data)       
        txtTotalEntries.insert(tk.END,TotalEntries)              
        conn.commit()
        conn.close()

    def TotalInvalidRemoved():
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")    
        Complete_df_PSSLog_DuplicatedShotID = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_DuplicatedShotID ORDER BY `ShotID` ASC ;", conn)
        Complete_df_PSSLog_INVALID_NULL= pd.read_sql_query("SELECT * FROM Eagle_PSSLog_INVALID_NULL ORDER BY `ShotID` ASC ;", conn)
        Complete_df = Complete_df_PSSLog_DuplicatedShotID.append(Complete_df_PSSLog_INVALID_NULL, ignore_index=True)    
        data = pd.DataFrame(Complete_df)
        data = data.reset_index(drop=True)
        TotalinvalidRemoved = len(data)       
        txtTotalinvalidRemoved.insert(tk.END,TotalinvalidRemoved)              
        conn.commit()
        conn.close()
        
    def TotalRAW_ImportedEntries():
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_RAWDUMP ;", conn)
        data = pd.DataFrame(Complete_df)
        TotalEntries = len(data)       
        txtTotalRAWImport.insert(tk.END,TotalEntries)              
        conn.commit()
        conn.close()

    def InvalidEntries():
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_INVALID_NULL ;", conn)
        data = pd.DataFrame(Complete_df)
        TotalInvalidEntries = len(data)       
        txtInvalidEntries.insert(tk.END,TotalInvalidEntries)              
        conn.commit()
        conn.close()

    def DuplicatedShotIDEntries():
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_DuplicatedShotID ;", conn)
        data = pd.DataFrame(Complete_df)
        TotalDuplicatedEntries = len(data)       
        txtDuplicatedShotID.insert(tk.END,TotalDuplicatedEntries)              
        conn.commit()
        conn.close()

    def VoidShotIDEntries():
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_VOID ;", conn)
        data = pd.DataFrame(Complete_df)
        TotalVoidEntries = len(data)       
        txtVoidEntries.insert(tk.END,TotalVoidEntries)              
        conn.commit()
        conn.close()

    def QCPassedShotIDEntries():
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_QCPassed ;", conn)
        data = pd.DataFrame(Complete_df)
        TotalQCPassedShotIDEntries = len(data)       
        txtQCPassedEntries.insert(tk.END,TotalQCPassedShotIDEntries)              
        conn.commit()
        conn.close()

    def QCFailedShotIDEntries():
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_QCFailed ;", conn)
        data = pd.DataFrame(Complete_df)
        TotalQCFailedShotIDEntries = len(data)       
        txtQCFailedEntries.insert(tk.END,TotalQCFailedShotIDEntries)              
        conn.commit()
        conn.close()

    def UpdateDuplicatedShotID():
        cur_id = tree.focus()
        selvalue = tree.item(cur_id)['values']
        Length_Selected  =  (len(selvalue))
        if Length_Selected != 0:
            for item in tree.selection():
                list_item = (tree.item(item, 'values'))
                txtDuplicatedShotExchange.delete(0,END)
                txtDuplicatedShotExchange.insert(tk.END,list_item[1])
                con= sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
                cur=con.cursor()
                cur.execute("UPDATE Eagle_PSSLog_TEMP SET FileNum = ?,  EPNumber = ?, SourceLine = ?, SourceStation = ?, Local_Date = ?, Local_Time = ?, Observer_Comment = ?, ShotStatus = ?,\
                            PhaseMax = ?, PhaseAvg = ?, ForceMax = ?, ForceAvg = ?, THDMax = ?, THDAvg = ?, SwCksm = ?, PmCksm = ?, GPS_Quality = ?, Unit_ID = ?,\
                            TB_Date = ?, TB_Time = ?, TB_Micro = ?, Signature_File_Number = ?,\
                            Latitude = ?, Longitude = ?, Altitude = ?, Encoder_Index = ?, Record_Index = ?, EP_Count = ?, Crew_ID = ?,\
                            Start_Code = ?, Force_Out = ?, GPS_Time = ?, GPS_Altitude = ?, Sats = ?, PDOP = ?, HDOP = ?, VDOP = ?, Age = ?, Start_Time_Delta = ?, Sweep_Number= ? WHERE ShotID =?", 
                            (list_item[2],list_item[3],list_item[4],list_item[5],
                             list_item[6],list_item[7],list_item[8],list_item[9],list_item[10],list_item[11],
                             list_item[12],list_item[13],list_item[14],list_item[15],list_item[16],list_item[17],
                             list_item[18],list_item[19],list_item[20],list_item[21],list_item[22],list_item[23],
                             list_item[24],list_item[25],list_item[26],list_item[27],list_item[28],list_item[29],
                             list_item[30],list_item[31],list_item[32],list_item[33],list_item[34],list_item[35],
                             list_item[36],list_item[37],list_item[38],list_item[39],list_item[40],list_item[41], list_item[1]))   
                con.commit()
                con.close()      
            tkinter.messagebox.showinfo("Update Duplicate ShotID Message","Selected List of Entries Added To Imported Valid Database")
            ViewTotalImport()
        else:
            tkinter.messagebox.showinfo("Update Duplicate ShotID Message","Please Select List of Entries To Update Imported Valid Database")

      
    def DeleteSelectedImportData():
        iDelete = tkinter.messagebox.askyesno("Delete Entry", "Confirm if you want to Delete")
        if iDelete >0:
            txtTotalEntries.delete(0,END)
            conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
            cur = conn.cursor()                
            for selected_item in tree.selection():
                cur.execute("DELETE FROM Eagle_PSSLog_TEMP WHERE DataBase_ID =? " ,(tree.set(selected_item, '#1'),)) 
                conn.commit()
                tree.delete(selected_item)
            conn.commit()
            conn.close()
            TotalEntries()
            return

    def ExportValidForTest_i_Fy():
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_TEMP ORDER BY `ShotID` ASC ;", conn)            
        conn.commit()
        conn.close()
        Export_TestifyDF = pd.DataFrame(Complete_df)
        Export_TestifyDF = Export_TestifyDF.loc[:,['Encoder_Index','ShotStatus','ShotID','FileNum','EPNumber','SourceLine','SourceStation',
                                                               'Local_Date','Local_Time','Observer_Comment','TB_Date','TB_Time','TB_Micro',
                                                               'Record_Index','EP_Count','Crew_ID','Unit_ID','Start_Code','SwCksm','PmCksm',
                                                               'PhaseMax','PhaseAvg','ForceMax','ForceAvg','THDMax','THDAvg','Force_Out','GPS_Time',
                                                               'Latitude','Longitude','Altitude','GPS_Altitude','Sats','PDOP','HDOP',
                                                               'VDOP','Age','GPS_Quality', 'Start_Time_Delta','Sweep_Number','Signature_File_Number']]

        Export_TestifyDF.rename(columns={'Encoder_Index':'Encoder Index', 'ShotStatus':'Void', 'ShotID':'Shot ID', 'FileNum':'File Num', 'EPNumber':'EP ID', 'SourceLine':'Line',
                             'SourceStation':'Station','Local_Date':'Date','Local_Time':'Time','Observer_Comment':'Comment','TB_Date':'TB Date','TB_Time':'TB Time',
                             'TB_Micro':'TB Micro','Record_Index':'Record Index','EP_Count':'EP Count','Crew_ID':'Crew ID','Unit_ID':'Unit ID','Start_Code':'Start Code',                        
                             'SwCksm':'Sweep Checksum','PmCksm':'Param Checksum','PhaseMax':'Phase Max','PhaseAvg':'Phase Avg','ForceMax':'Force Max',
                             'ForceAvg':'Force Avg','THDMax':'THD Max','THDAvg':'THD Avg','Force_Out':'Force Out','GPS_Time':'GPS Time',
                             'Latitude':'Lat','Longitude':'Lon','Altitude':'Altitude','GPS_Altitude':'GPS Altitude','Sats':'Sats',
                             'PDOP':'PDOP','HDOP':'HDOP','VDOP':'VDOP','Age':'Age','GPS_Quality':'Quality','Start_Time_Delta':'Start Time Delta','Sweep_Number':'Sweep Number',
                             'Signature_File_Number':'Signature File Number'},inplace = True)
        Export_TestifyDF['Date']    = pd.to_datetime(Export_TestifyDF['Date']).dt.strftime('%m/%d/%Y')
        Export_TestifyDF['TB Date'] = pd.to_datetime(Export_TestifyDF['TB Date']).dt.strftime('%m/%d/%Y')
        
        Export_TestifyDF = Export_TestifyDF.reset_index(drop=True)
        Export_TestifyDF_With_VOID = pd.DataFrame(Export_TestifyDF)

        Export_TestifyDF_With_No_VOID = pd.DataFrame(Export_TestifyDF)
        Export_TestifyDF_With_No_VOID = Export_TestifyDF_With_No_VOID[(Export_TestifyDF_With_No_VOID.Void.isnull())]
        

        ExportVoidorNotQuestion = tkinter.messagebox.askquestion("Export Testif-y Message",
                                "Do you Like To Export Testif-y Input File Including All Void Shots?"+ '\n' +
                                'Yes For Export With Including All Void Shots' + '\n' +
                                'No For Export With Excluding (Removing) All Void Shots')


        if ExportVoidorNotQuestion == "yes":
            filename = tkinter.filedialog.asksaveasfilename(initialdir = "/" ,title = "Save Valid PSS Export For Testif-y As CSV Or Excel" ,\
                   defaultextension='.csv', filetypes = (("CSV file",".csv"),("Excel file",".xlsx")))

            if len(filename) >0:
                if filename.endswith('.csv'):
                    Export_TestifyDF_With_VOID.to_csv((filename),index=None)
                    tkinter.messagebox.showinfo("Valid PSS Export For Testif-y","PSS For Test-i-fy Saved as CSV")
                else:
                    Export_TestifyDF_With_VOID.to_excel(filename, sheet_name='PSS Testif-y', index=False)
                    tkinter.messagebox.showinfo("Valid PSS Export For Testif-y","PSS For Testif-y Saved as Excel")
            else:
                tkinter.messagebox.showinfo("Export Testif-y Message","Please Select File Name To Export")
        else:
            filename = tkinter.filedialog.asksaveasfilename(initialdir = "/" ,title = "Save Valid PSS Export For Testif-y As CSV Or Excel" ,\
                   defaultextension='.csv', filetypes = (("CSV file",".csv"),("Excel file",".xlsx")))

            if len(filename) >0:
                if filename.endswith('.csv'):
                    Export_TestifyDF_With_No_VOID.to_csv((filename),index=None)
                    tkinter.messagebox.showinfo("Valid PSS Export For Testif-y","PSS For Test-i-fy Saved as CSV")
                else:
                    Export_TestifyDF_With_No_VOID.to_excel(filename, sheet_name='PSS Testif-y', index=False)
                    tkinter.messagebox.showinfo("Valid PSS Export For Testif-y","PSS For Testif-y Saved as Excel")
            else:
                tkinter.messagebox.showinfo("Export Testif-y Message","Please Select File Name To Export")

    def ExportValidForGlobalMapper():
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_TEMP ORDER BY `ShotID` ASC ;", conn)
        conn.commit()
        conn.close()
        Export_GMapperDF = pd.DataFrame(Complete_df)
        Export_GMapperDF = Export_GMapperDF.loc[:,['Encoder_Index','ShotStatus','ShotID','FileNum','EPNumber','SourceLine','SourceStation',
                                                               'Local_Date','Local_Time','Observer_Comment','TB_Date','TB_Time','TB_Micro',
                                                               'Record_Index','EP_Count','Crew_ID','Unit_ID','Start_Code','SwCksm','PmCksm',
                                                               'PhaseMax','PhaseAvg','ForceMax','ForceAvg','THDMax','THDAvg','Force_Out','GPS_Time',
                                                               'Latitude','Longitude','Altitude','GPS_Altitude','Sats','PDOP','HDOP',
                                                               'VDOP','Age','GPS_Quality', 'Start_Time_Delta','Sweep_Number','Signature_File_Number']]

        Export_GMapperDF.rename(columns={'Encoder_Index':'Encoder Index', 'ShotStatus':'Void', 'ShotID':'Shot ID', 'FileNum':'File Num', 'EPNumber':'EP ID', 'SourceLine':'Line',
                             'SourceStation':'Station','Local_Date':'Date','Local_Time':'Time','Observer_Comment':'Comment','TB_Date':'TB Date','TB_Time':'TB Time',
                             'TB_Micro':'TB Micro','Record_Index':'Record Index','EP_Count':'EP Count','Crew_ID':'Crew ID','Unit_ID':'Unit ID','Start_Code':'Start Code',                        
                             'SwCksm':'Sweep Checksum','PmCksm':'Param Checksum','PhaseMax':'Phase Max','PhaseAvg':'Phase Avg','ForceMax':'Force Max',
                             'ForceAvg':'Force Avg','THDMax':'THD Max','THDAvg':'THD Avg','Force_Out':'Force Out','GPS_Time':'GPS Time',
                             'Latitude':'Lat','Longitude':'Lon','Altitude':'Altitude','GPS_Altitude':'GPS Altitude','Sats':'Sats',
                             'PDOP':'PDOP','HDOP':'HDOP','VDOP':'VDOP','Age':'Age','GPS_Quality':'Quality','Start_Time_Delta':'Start Time Delta','Sweep_Number':'Sweep Number',
                             'Signature_File_Number':'Signature File Number'},inplace = True)

        Export_GMapperDF = Export_GMapperDF.loc[:,['Lat', 'Lon','Altitude','GPS Altitude','Sats',
                                                   'Encoder Index', 'Void', 'Shot ID', 'File Num', 'EP ID',
                                                   'Line', 'Station','Date','Time','Comment','TB Date','TB Time',
                                                   'TB Micro','Record Index','EP Count','Crew ID','Unit ID','Start Code',                        
                                                   'Sweep Checksum','Param Checksum','Phase Max','Phase Avg','Force Max',
                                                   'Force Avg','THD Max','THD Avg','Force Out','GPS Time',
                                                   'PDOP','HDOP','VDOP','Age','Quality','Start Time Delta','Sweep Number','Signature File Number']]
        Export_GMapperDF_Check      = pd.DataFrame(Export_GMapperDF)        
        Export_GMapperDF_Check      = Export_GMapperDF_Check[pd.to_numeric(Export_GMapperDF_Check.Lat,errors='coerce').notnull()]
        Export_GMapperDF_Check      = Export_GMapperDF_Check[pd.to_numeric(Export_GMapperDF_Check.Lon,errors='coerce').notnull()]
        Export_GMapperDF_Check      = Export_GMapperDF[(Export_GMapperDF.Lat > 0)&
                                                  (abs(Export_GMapperDF.Lon) > 0)]        
        Export_GMapperDF_Check      = Export_GMapperDF_Check.reset_index(drop=True)
        Export_GMapperDF            = pd.DataFrame(Export_GMapperDF_Check)
        Export_GMapperDF['Date']    = pd.to_datetime(Export_GMapperDF['Date']).dt.strftime('%m/%d/%Y')
        Export_GMapperDF['TB Date'] = pd.to_datetime(Export_GMapperDF['TB Date']).dt.strftime('%m/%d/%Y')    
        Export_GMapperDF            = Export_GMapperDF.reset_index(drop=True)
        Export_GMapperDF            = pd.DataFrame(Export_GMapperDF)
        
        filename = tkinter.filedialog.asksaveasfilename(initialdir = "/" ,title = "Save Valid PSS Export For Global Mapper As CSV Or Excel" ,\
               defaultextension='.csv', filetypes = (("CSV file",".csv"),("Excel file",".xlsx")))

        if len(filename) >0:
            if filename.endswith('.csv'):
                Export_GMapperDF.to_csv((filename),index=None)
                tkinter.messagebox.showinfo("Valid PSS Export For Global Mapper","PSS For Global Mapper Saved as CSV")
            else:
                Export_GMapperDF.to_excel(filename, sheet_name='PSS Testif-y', index=False)
                tkinter.messagebox.showinfo("Valid PSS Export For Global Mapper","PSS For Global Mapper Saved as Excel")
        else:
            tkinter.messagebox.showinfo("Export Global Mapper Message","Please Select File Name To Export")
        

    def ExportRAW_ImportedPSS():
        conn = sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_RAWDUMP ORDER BY `ShotID` ASC ;", conn)
        Complete_df['Local_Date']             = pd.to_datetime(Complete_df['Local_Date']).dt.strftime('%Y/%m/%d')
        Complete_df['TB_Date']                = pd.to_datetime(Complete_df['TB_Date']).dt.strftime('%Y/%m/%d')
        Complete_df['TB_Time']                = pd.to_datetime(Complete_df['TB_Time']).dt.strftime('%H:%M:%S')
        Complete_df['TB_Micro']               = pd.to_datetime(Complete_df['TB_Micro']).dt.strftime('%f')
        Export_RAW_ImportedPSS = pd.DataFrame(Complete_df)
        Export_RAW_ImportedPSS = Export_RAW_ImportedPSS.loc[:,['Encoder_Index','ShotStatus','ShotID','FileNum','EPNumber','SourceLine','SourceStation',
                                                               'Local_Date','Local_Time','Observer_Comment','TB_Date','TB_Time','TB_Micro',
                                                               'Record_Index','EP_Count','Crew_ID','Unit_ID','Start_Code','SwCksm','PmCksm',
                                                               'PhaseMax','PhaseAvg','ForceMax','ForceAvg','THDMax','THDAvg','Force_Out','GPS_Time',
                                                               'Latitude','Longitude','Altitude','GPS_Altitude','Sats','PDOP','HDOP',
                                                               'VDOP','Age','GPS_Quality', 'Start_Time_Delta','Sweep_Number','Signature_File_Number']]

        Export_RAW_ImportedPSS.rename(columns={'Encoder_Index':'Encoder Index', 'ShotStatus':'Void', 'ShotID':'Shot ID', 'FileNum':'File Num', 'EPNumber':'EP ID', 'SourceLine':'Line',
                             'SourceStation':'Station','Local_Date':'Date','Local_Time':'Time','Observer_Comment':'Comment','TB_Date':'TB Date','TB_Time':'TB Time',
                             'TB_Micro':'TB Micro','Record_Index':'Record Index','EP_Count':'EP Count','Crew_ID':'Crew ID','Unit_ID':'Unit ID','Start_Code':'Start Code',                        
                             'SwCksm':'Sweep Checksum','PmCksm':'Param Checksum','PhaseMax':'Phase Max','PhaseAvg':'Phase Avg','ForceMax':'Force Max',
                             'ForceAvg':'Force Avg','THDMax':'THD Max','THDAvg':'THD Avg','Force_Out':'Force Out','GPS_Time':'GPS Time',
                             'Latitude':'Lat','Longitude':'Lon','Altitude':'Altitude','GPS_Altitude':'GPS Altitude','Sats':'Sats',
                             'PDOP':'PDOP','HDOP':'HDOP','VDOP':'VDOP','Age':'Age','GPS_Quality':'Quality','Start_Time_Delta':'Start Time Delta','Sweep_Number':'Sweep Number',
                             'Signature_File_Number':'Signature File Number'},inplace = True)
        Export_RAW_ImportedPSS['Date']    = pd.to_datetime(Export_RAW_ImportedPSS['Date']).dt.strftime('%m/%d/%Y')
        Export_RAW_ImportedPSS['TB Date'] = pd.to_datetime(Export_RAW_ImportedPSS['TB Date']).dt.strftime('%m/%d/%Y')
        
        Export_RAW_ImportedPSS = Export_RAW_ImportedPSS.reset_index(drop=True)
        
        filename = tkinter.filedialog.asksaveasfilename(initialdir = "/" ,title = "Save RAW Impoted PSS Export As Excel" ,\
                   defaultextension='.csv', filetypes = (("CSV file",".csv"),("Excel file",".xlsx")))
        if len(filename) >0:
            if filename.endswith('.csv'):
                Export_RAW_ImportedPSS.to_csv(filename,index=None)
                tkinter.messagebox.showinfo("RAW Impoted PSS Export","RAW Impoted PSS Saved as CSV")
            else:
                Export_RAW_ImportedPSS.to_excel(filename, sheet_name='RAW Impoted PSS', index=False)
                tkinter.messagebox.showinfo("RAW Impoted PSS","RAW Impoted PSS Saved as Excel")
        else:
            tkinter.messagebox.showinfo("Export RAW Impoted PSS Message","Please Select File Name To Export")
                    
        conn.commit()
        conn.close()

    def ExportListBoxPSS():
        dfList =[] 
        for child in tree.get_children():
            df = tree.item(child)["values"]
            dfList.append(df)
        ListBox_DF = pd.DataFrame(dfList)
        ListBox_DF.rename(columns={0:'DB_ID', 1:'Shot ID', 2:'File Num', 3:'EP ID', 4:'Line', 5:'Station', 6:'Date',
                                       7:'Time',8:'Comment',9:'Void',10:'Phase Max',11:'Phase Avg',12:'Force Max',
                                       13:'Force Avg',14:'THD Max',15:'THD Avg',16:'Sweep Checksum',17:'Param Checksum',18:'Quality',
                                       19:'Unit ID',20:'TB Date',21:'TB Time',22:'TB Micro',23:'Signature File Number',
                                       24:'Lat',25:'Lon',26:'Altitude',27:'Encoder Index',28:'Record Index',
                                       29:'EP Count',30:'Crew ID',31:'Start Code',32:'Force Out',33:'GPS Time',
                                       34:'GPS Altitude',35:'Sats',36:'PDOP',37:'HDOP',38:'VDOP',39:'Age',40:'Start Time Delta',
                                       41:'Sweep Number'},inplace = True)
                        
        Export_ListBox  = pd.DataFrame(ListBox_DF)
        TotalListBox= len(Export_ListBox)       
        if TotalListBox >0:        
            Export_ListBox.drop(['DB_ID'], axis=1, inplace=True)
            Export_ListBox['Void'].replace('None', np.nan, inplace=True)
            Export_ListBox  = Export_ListBox.reset_index(drop=True)    
            filename = tkinter.filedialog.asksaveasfilename(initialdir = "/" ,title = "Select File Name to Export" ,\
                           defaultextension='.csv', filetypes = (("CSV file",".csv"),("Excel file",".xlsx")))
            if len(filename) >0:
                if filename.endswith('.csv'):
                    Export_ListBox.to_csv(filename,index=None)
                    tkinter.messagebox.showinfo("ListBox PSS Export","ListBox PSS Entries Saved as CSV")
                else:
                    Export_ListBox.to_excel(filename, sheet_name='ListBoxTB', index=False)
                    tkinter.messagebox.showinfo("ListBox PSS Export","ListBox PSS Entries Saved as Excel")
            else:
                tkinter.messagebox.showinfo("ListBox PSS Export Message","Please Select File Name To Export")

    def SortbyLineStation():
        dfList =[] 
        for child in tree.get_children():
            df = tree.item(child)["values"]
            dfList.append(df)
        ListBox_DF = pd.DataFrame(dfList)
        ListBox_DF.rename(columns={0:'DB_ID', 1:'ShotID', 2:'FileNum', 3:'EPNumber', 4:'SourceLine', 5:'SourceStation', 6:'Local_Date',
                                   7:'Local_Time',8:'Observer_Comment',9:'ShotStatus',10:'PhaseMax',11:'PhaseAvg',12:'ForceMax',
                                   13:'ForceAvg',14:'THDMax',15:'THDAvg',16:'SwCksm',17:'PmCksm',18:'GPS_Quality',
                                   19:'Unit_ID',20:'TB_Date',21:'TB_Time',22:'TB_Micro',23:'Signature_File_Number',
                                   24:'Latitude',25:'Longitude',26:'Altitude',27:'Encoder_Index',28:'Record_Index',
                                   29:'EP_Count',30:'Crew_ID',31:'Start_Code',32:'Force_Out',33:'GPS_Time',
                                   34:'GPS_Altitude',35:'Sats',36:'PDOP',37:'HDOP',38:'VDOP',39:'Age',40:'Start_Time_Delta',
                                   41:'Sweep_Number'},inplace = True)
                    
        SortbyLineStation_ListBox  = pd.DataFrame(ListBox_DF)    
        TotalListBox= len(SortbyLineStation_ListBox)       
        if TotalListBox >0:
            SortbyLineStation_ListBox  = SortbyLineStation_ListBox.sort_values(by =['SourceLine', 'SourceStation'])
            SortbyLineStation_ListBox  = SortbyLineStation_ListBox.reset_index(drop=True)
            tree.delete(*tree.get_children())
            for each_rec in range(len(SortbyLineStation_ListBox)):
                    tree.insert("", tk.END, values=list(SortbyLineStation_ListBox.loc[each_rec]))

    def ViewQCLimit():
        tkinter.messagebox.showinfo("Vib QC Limit Parameters", QCLimit_Summary)

    def ImportPSSLogFile():
        ClearView()            
        UTC_Offset_Hours = (Entrytxt_TimeOffset.get())
        UTC_Offset_Hours = float(UTC_Offset_Hours)
        
        fileList = askopenfilenames(initialdir = "/", title = "Import SourceLink PSS Files" , filetypes=[('CSV File', '*.csv'), ('Excel File', ('*.xls', '*.xlsx'))])
        Length_fileList  =  len(fileList)
        if Length_fileList >0:            
            if fileList:
                dfList =[]            
                for filename in fileList:
                    if filename.endswith('.csv'):
                        df = pd.read_csv(filename, sep=',' , low_memory=False)
                        df = df.iloc[:,:]
                        ShotID                = df.loc[:,'Shot ID']
                        FileNum               = df.loc[:,'File Num']
                        EPNumber              = df.loc[:,'EP ID']
                        SourceLine            = df.loc[:,'Line']
                        SourceStation         = df.loc[:,'Station']
                        Local_Date            = df.loc[:,'Date']
                        Local_Time            = df.loc[:,'Time']
                        Observer_Comment      = df.loc[:,'Comment']
                        ShotStatus            = df.loc[:,'Void']
                        PhaseMax              = df.loc[:,'Phase Max']
                        PhaseAvg              = df.loc[:,'Phase Avg']
                        ForceMax              = df.loc[:,'Force Max']
                        ForceAvg              = df.loc[:,'Force Avg']
                        THDMax                = df.loc[:,'THD Max']
                        THDAvg                = df.loc[:,'THD Avg']
                        SwCksm                = df.loc[:,'Sweep Checksum']
                        PmCksm                = df.loc[:,'Param Checksum']
                        GPS_Quality           = df.loc[:,'Quality']
                        Unit_ID               = df.loc[:,'Unit ID']
                        try:
                            TB_Date               = df.loc[:,'TB UTC Time']
                            TB_Time               = df.loc[:,'TB UTC Time']
                            TB_Micro              = df.loc[:,'TB UTC Time']
                        except:
                            TB_Date               = df.loc[:,'TB Date']
                            TB_Time               = df.loc[:,'TB Time']
                            TB_Micro              = (df.loc[:,'TB Micro'])*1000
                            
                        Signature_File_Number = df.loc[:,'Signature File Number']
                        Latitude              = df.loc[:,'Lat']
                        Longitude             = df.loc[:,'Lon']
                        Altitude              = df.loc[:,'Altitude']
                        Encoder_Index         = df.loc[:,'Encoder Index']
                        Record_Index          = df.loc[:,'Record Index']
                        EP_Count              = df.loc[:,'EP Count']
                        Crew_ID               = df.loc[:,'Crew ID']
                        Start_Code            = df.loc[:,'Start Code']
                        Force_Out             = df.loc[:,'Force Out']
                        GPS_Time              = df.loc[:,'GPS Time']
                        GPS_Altitude          = df.loc[:,'GPS Altitude']
                        Sats                  = df.loc[:,'Sats']
                        PDOP                  = df.loc[:,'PDOP']
                        HDOP                  = df.loc[:,'HDOP']
                        VDOP                  = df.loc[:,'VDOP']
                        Age                   = df.loc[:,'Age']
                        Start_Time_Delta      = df.loc[:,'Start Time Delta']
                        Sweep_Number          = df.loc[:,'Sweep Number']
                        column_names = [ShotID, FileNum, EPNumber, SourceLine, SourceStation, Local_Date,Local_Time, Observer_Comment, ShotStatus,
                            PhaseMax, PhaseAvg, ForceMax, ForceAvg, THDMax, THDAvg, SwCksm,
                            PmCksm,GPS_Quality,Unit_ID,TB_Date,TB_Time,TB_Micro,Signature_File_Number,
                            Latitude,Longitude,Altitude,Encoder_Index,Record_Index,EP_Count,Crew_ID,
                            Start_Code,Force_Out,GPS_Time,GPS_Altitude,Sats,PDOP,HDOP,VDOP,Age,Start_Time_Delta,
                            Sweep_Number]
                        catdf = pd.concat (column_names,axis=1,ignore_index =True)
                        dfList.append(catdf) 
                    else:
                        df = pd.read_excel(filename)
                        df = df.iloc[:,:]
                        ShotID                = df.loc[:,'Shot ID']
                        FileNum               = df.loc[:,'File Num']
                        EPNumber              = df.loc[:,'EP ID']
                        SourceLine            = df.loc[:,'Line']
                        SourceStation         = df.loc[:,'Station']
                        Local_Date            = df.loc[:,'Date']
                        Local_Time            = df.loc[:,'Time']
                        Observer_Comment      = df.loc[:,'Comment']
                        ShotStatus            = df.loc[:,'Void']
                        PhaseMax              = df.loc[:,'Phase Max']
                        PhaseAvg              = df.loc[:,'Phase Avg']
                        ForceMax              = df.loc[:,'Force Max']
                        ForceAvg              = df.loc[:,'Force Avg']
                        THDMax                = df.loc[:,'THD Max']
                        THDAvg                = df.loc[:,'THD Avg']
                        SwCksm                = df.loc[:,'Sweep Checksum']
                        PmCksm                = df.loc[:,'Param Checksum']
                        GPS_Quality           = df.loc[:,'Quality']
                        Unit_ID               = df.loc[:,'Unit ID']
                        try:
                            TB_Date               = df.loc[:,'TB UTC Time']
                            TB_Time               = df.loc[:,'TB UTC Time']
                            TB_Micro              = df.loc[:,'TB UTC Time']
                        except:
                            TB_Date               = df.loc[:,'TB Date']
                            TB_Time               = df.loc[:,'TB Time']
                            TB_Micro              = (df.loc[:,'TB Micro'])*1000
                        Signature_File_Number = df.loc[:,'Signature File Number']
                        Latitude              = df.loc[:,'Lat']
                        Longitude             = df.loc[:,'Lon']
                        Altitude              = df.loc[:,'Altitude']
                        Encoder_Index         = df.loc[:,'Encoder Index']
                        Record_Index          = df.loc[:,'Record Index']
                        EP_Count              = df.loc[:,'EP Count']
                        Crew_ID               = df.loc[:,'Crew ID']
                        Start_Code            = df.loc[:,'Start Code']
                        Force_Out             = df.loc[:,'Force Out']
                        GPS_Time              = df.loc[:,'GPS Time']
                        GPS_Altitude          = df.loc[:,'GPS Altitude']
                        Sats                  = df.loc[:,'Sats']
                        PDOP                  = df.loc[:,'PDOP']
                        HDOP                  = df.loc[:,'HDOP']
                        VDOP                  = df.loc[:,'VDOP']
                        Age                   = df.loc[:,'Age']
                        Start_Time_Delta      = df.loc[:,'Start Time Delta']
                        Sweep_Number          = df.loc[:,'Sweep Number']
                        column_names = [ShotID, FileNum, EPNumber, SourceLine, SourceStation, Local_Date,Local_Time, Observer_Comment, ShotStatus,
                            PhaseMax, PhaseAvg, ForceMax, ForceAvg, THDMax, THDAvg, SwCksm,
                            PmCksm,GPS_Quality,Unit_ID,TB_Date,TB_Time,TB_Micro,Signature_File_Number,
                            Latitude,Longitude,Altitude,Encoder_Index,Record_Index,EP_Count,Crew_ID,
                            Start_Code,Force_Out,GPS_Time,GPS_Altitude,Sats,PDOP,HDOP,VDOP,Age,Start_Time_Delta,
                            Sweep_Number]
                        catdf = pd.concat (column_names,axis=1,ignore_index =True)
                        dfList.append(catdf) 

                concatDf = pd.concat(dfList,axis=0, ignore_index =True)
                concatDf.rename(columns={0:'ShotID', 1:'FileNum', 2:'EPNumber', 3:'SourceLine', 4:'SourceStation', 5:'Local_Date',
                                 6:'Local_Time',7:'Observer_Comment',8:'ShotStatus',9:'PhaseMax',10:'PhaseAvg',11:'ForceMax',
                                 12:'ForceAvg',13:'THDMax',14:'THDAvg',15:'SwCksm',16:'PmCksm',17:'GPS_Quality',
                                 18:'Unit_ID',19:'TB_Date',20:'TB_Time',21:'TB_Micro',22:'Signature_File_Number',
                                 23:'Latitude',24:'Longitude',25:'Altitude',26:'Encoder_Index',27:'Record_Index',
                                 28:'EP_Count',29:'Crew_ID',30:'Start_Code',31:'Force_Out',32:'GPS_Time',
                                 33:'GPS_Altitude',34:'Sats',35:'PDOP',36:'HDOP',37:'VDOP',38:'Age',39:'Start_Time_Delta',
                                 40:'Sweep_Number'},inplace = True)
                # RAW DUMP Total PSS imported
                RAW_DUMP_ImportedPSS_DF    = pd.DataFrame(concatDf)

                # Separating InValid with Shot ID is Null
                Invalid_PSS_DF    = pd.DataFrame(concatDf)
                Invalid_PSS_DF    = Invalid_PSS_DF[pd.to_numeric(Invalid_PSS_DF.ShotID,errors='coerce').isnull()]                    
                Invalid_PSS_DF    = Invalid_PSS_DF.reset_index(drop=True)
                Data_Invalid_PSS  = pd.DataFrame(Invalid_PSS_DF)
                
                # Separating Valid with Shot ID Not Null
                Valid_PSS_DF = pd.DataFrame(concatDf)
                Valid_PSS_DF = Valid_PSS_DF[pd.to_numeric(Valid_PSS_DF.ShotID, errors='coerce').notnull()]                  
                Valid_PSS_DF["SourceLine"].fillna(0, inplace = True)
                Valid_PSS_DF["SourceStation"].fillna(0, inplace = True)
                Valid_PSS_DF["FileNum"].fillna(0, inplace = True)
                Valid_PSS_DF["EPNumber"].fillna(1, inplace = True)
                Valid_PSS_DF["Local_Date"].fillna('1900/1/01', inplace = True)
                Valid_PSS_DF["TB_Date"].fillna('1900/1/01', inplace = True)                
                Valid_PSS_DF['SourceLine']             = (Valid_PSS_DF.loc[:,['SourceLine']]).astype(int)
                Valid_PSS_DF['SourceStation']          = (Valid_PSS_DF.loc[:,['SourceStation']]).astype(float)
                Valid_PSS_DF['ShotID']                 = (Valid_PSS_DF.loc[:,['ShotID']]).astype(int)
                Valid_PSS_DF['FileNum']                = (Valid_PSS_DF.loc[:,['FileNum']]).astype(int)
                Valid_PSS_DF['EPNumber']               = (Valid_PSS_DF.loc[:,['EPNumber']]).astype(int)
                Valid_PSS_DF['Local_Date']             = pd.to_datetime(Valid_PSS_DF['Local_Date']).dt.strftime('%Y/%m/%d')
                try:                    
                    Valid_PSS_DF['TB_Date']            = pd.to_datetime(Valid_PSS_DF['TB_Date']).dt.strftime('%Y/%m/%d')
                    Valid_PSS_DF['TB_Time']            = pd.to_datetime(Valid_PSS_DF['TB_Time']).dt.strftime('%H:%M:%S')
                    Valid_PSS_DF['TB_Micro']           = pd.to_datetime(Valid_PSS_DF['TB_Micro']).dt.strftime('%f')
                except:
                    try:
                        Valid_PSS_DF['TB_Date']            = pd.to_datetime(Valid_PSS_DF['Local_Date']).dt.strftime('%Y/%m/%d')
                        Valid_PSS_DF['TB_Time']            = pd.to_datetime(Valid_PSS_DF['Local_Time']).dt.strftime('%H:%M:%S')
                        Valid_PSS_DF['TB_DateTime']        = pd.to_datetime(Valid_PSS_DF.TB_Date.astype(str)+' '+Valid_PSS_DF.TB_Time.astype(str))                                                                        
                        Valid_PSS_DF['TB_DateTime']        = pd.to_datetime(Valid_PSS_DF['TB_DateTime'].astype(str)) + pd.DateOffset(hours=UTC_Offset_Hours)
                        Valid_PSS_DF['TB_Date']            = pd.to_datetime(Valid_PSS_DF['TB_DateTime']).dt.strftime('%Y/%m/%d')
                        Valid_PSS_DF['TB_Time']            = pd.to_datetime(Valid_PSS_DF['TB_DateTime']).dt.strftime('%H:%M:%S')                                                
                        Valid_PSS_DF['TB_Micro']           = 0
                        Valid_PSS_DF.drop(['TB_DateTime'], axis=1, inplace=True)
                        tkinter.messagebox.showinfo("Imported PSS File Message","PSS Column Name : [TB UTC Time] Is Corrupted" + '\n' +  '\n' +
                                                    "Columns [TB_Date], [TB_Time] Is Fixed From Local Time <<<>>> [TB_Micro] Column Is Corrupted" + '\n' +  '\n' +
                                                    " To Fix All TB Columns Go To Advanced Option and Select >> Fix Corrupted PSS From TB Import >> ")
                    except:
                        Valid_PSS_DF['TB_Date']            = Valid_PSS_DF['TB_Date']
                        Valid_PSS_DF['TB_Time']            = Valid_PSS_DF['TB_Time']
                        Valid_PSS_DF['TB_Micro']           = Valid_PSS_DF['TB_Micro']
                        tkinter.messagebox.showinfo("Imported PSS File Message","PSS Column Name : [TB UTC Time] Is Corrupted" + '\n' +  '\n' +
                                                    "[TB_Date], [TB_Time] and [TB_Micro] Columns are Corrupted" + '\n' +  '\n' +
                                                    " To Fix All TB Columns Go To Advanced Option and Select >> Fix Corrupted PSS From TB Import >> ")
                        
                
                Valid_PSS_DF['DuplicatedEntries']      = Valid_PSS_DF.sort_values(by =['ShotID', 'Unit_ID','Crew_ID',
                                                         'SwCksm','PmCksm','PhaseMax','PhaseAvg','ForceMax','ForceAvg',
                                                         'THDMax','THDAvg']).duplicated(['ShotID','FileNum','SourceLine','SourceStation'],keep='last')
                Valid_PSS_DF                           = Valid_PSS_DF.reset_index(drop=True)
                Valid_PSS_DF                           = pd.DataFrame(Valid_PSS_DF)

                # Separating Valid with Shot ID Duplicated
                DATA_DuplicatedShotID = Valid_PSS_DF.loc[Valid_PSS_DF.DuplicatedEntries == True, 'ShotID': 'Sweep_Number']
                DATA_DuplicatedShotID = DATA_DuplicatedShotID.reset_index(drop=True)
                DATA_DuplicatedShotID = pd.DataFrame(DATA_DuplicatedShotID)

                # Separating Valid with Shot ID Not Duplicated
                DATA_VALID_PSS = Valid_PSS_DF.loc[Valid_PSS_DF.DuplicatedEntries == False, 'ShotID': 'Sweep_Number']
                DATA_VALID_PSS = DATA_VALID_PSS.reset_index(drop=True)
                DATA_VALID_PSS = pd.DataFrame(DATA_VALID_PSS)

                ## Getting Void Shot DataFrame
                VIB_Rep_VOID        = DATA_VALID_PSS[(DATA_VALID_PSS.ShotStatus.notnull())]
                VIB_Rep_VOID        = VIB_Rep_VOID.reset_index(drop=True)

                ## Getting Not Void Shot DataFrame
                VIB_Rep_VALID_NOT_VOID    = DATA_VALID_PSS[(DATA_VALID_PSS.ShotStatus.isnull())]
                VIB_Rep_VALID_NOT_VOID    = VIB_Rep_VALID_NOT_VOID[pd.to_numeric(VIB_Rep_VALID_NOT_VOID.PhaseAvg,errors='coerce').notnull()]
                VIB_Rep_VALID_NOT_VOID    = VIB_Rep_VALID_NOT_VOID[pd.to_numeric(VIB_Rep_VALID_NOT_VOID.ForceAvg,errors='coerce').notnull()]
                VIB_Rep_VALID_NOT_VOID    = VIB_Rep_VALID_NOT_VOID[pd.to_numeric(VIB_Rep_VALID_NOT_VOID.THDAvg,errors='coerce').notnull()]
                VIB_Rep_VALID_NOT_VOID    = VIB_Rep_VALID_NOT_VOID[pd.to_numeric(VIB_Rep_VALID_NOT_VOID.THDMax,errors='coerce').notnull()]
                VIB_Rep_VALID_NOT_VOID    = VIB_Rep_VALID_NOT_VOID[pd.to_numeric(VIB_Rep_VALID_NOT_VOID.PhaseMax,errors='coerce').notnull()]
                VIB_Rep_VALID_NOT_VOID    = VIB_Rep_VALID_NOT_VOID.reset_index(drop=True)

                ## Getting QC Passed Shot Dataframe    
                VIB_Rep_QC_Passed   =  pd.DataFrame(VIB_Rep_VALID_NOT_VOID)
                VIB_Rep_QC_Passed['PhaseMax'] = VIB_Rep_QC_Passed['PhaseMax'].abs()
                VIB_Rep_QC_Passed   =  VIB_Rep_QC_Passed[(VIB_Rep_QC_Passed.PhaseMax <= PhaseMax_Limit)&
                                                     (VIB_Rep_QC_Passed.PhaseAvg <= High_PhaseAvg_Limit)&
                                                     (VIB_Rep_QC_Passed.PhaseAvg >= Low_PhaseAvg_Limit)&
                                                     (VIB_Rep_QC_Passed.ForceMax <= High_ForceMax_Limit)&
                                                     (VIB_Rep_QC_Passed.ForceMax >= Low_ForceMax_Limit)&                                     
                                                     (VIB_Rep_QC_Passed.ForceAvg <= High_ForceAvg_Limit)&
                                                     (VIB_Rep_QC_Passed.ForceAvg >= Low_ForceAvg_Limit)&
                                                     (VIB_Rep_QC_Passed.THDAvg   <= High_THDAvg_Limit)&
                                                     (VIB_Rep_QC_Passed.THDMax   <= High_THDMax_Limit)&  
                                                     (VIB_Rep_QC_Passed.GPS_Quality != "No Fix")]
                VIB_Rep_QC_Passed   = VIB_Rep_QC_Passed.reset_index(drop=True)
                VIB_Rep_QC_Passed   =  pd.DataFrame(VIB_Rep_QC_Passed)


                ## Getting QC Failed Shot Dataframe    
                VIB_Rep_QC_Failed   =  pd.DataFrame(VIB_Rep_VALID_NOT_VOID)
                VIB_Rep_QC_Failed['PhaseMax'] = VIB_Rep_QC_Failed['PhaseMax'].abs()
                VIB_Rep_QC_Failed   =  VIB_Rep_QC_Failed[(VIB_Rep_QC_Failed.PhaseMax > PhaseMax_Limit)|
                                              (VIB_Rep_QC_Failed.ForceAvg > High_ForceAvg_Limit)|
                                              (VIB_Rep_QC_Failed.ForceAvg < Low_ForceAvg_Limit)|
                                              (VIB_Rep_QC_Failed.THDAvg   > High_THDAvg_Limit)|
                                              (VIB_Rep_QC_Failed.THDMax   > High_THDMax_Limit)|
                                              (VIB_Rep_QC_Failed.GPS_Quality == "No Fix")|                              
                                              (VIB_Rep_QC_Failed.PhaseAvg > High_PhaseAvg_Limit)|
                                              (VIB_Rep_QC_Failed.PhaseAvg < Low_PhaseAvg_Limit)|
                                              (VIB_Rep_QC_Failed.ForceMax > High_ForceMax_Limit)|
                                              (VIB_Rep_QC_Failed.ForceMax < Low_ForceMax_Limit)]
                VIB_Rep_QC_Failed  = VIB_Rep_QC_Failed.reset_index(drop=True)
                VIB_Rep_QC_Failed  =  pd.DataFrame(VIB_Rep_QC_Failed)

                # Connect To Database and Export DF  
                con= sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
                cur=con.cursor()
                RAW_DUMP_ImportedPSS_DF.to_sql('Eagle_PSSLog_RAWDUMP',con, if_exists="replace",  index_label='DataBase_ID')
                DATA_VALID_PSS.to_sql('Eagle_PSSLog_TEMP',con, if_exists="replace",  index_label='DataBase_ID')
                Data_Invalid_PSS.to_sql('Eagle_PSSLog_INVALID_NULL',con, if_exists="replace", index_label='DataBase_ID')
                DATA_DuplicatedShotID.to_sql('Eagle_PSSLog_DuplicatedShotID',con, if_exists="replace", index_label='DataBase_ID')
                VIB_Rep_VOID.to_sql('Eagle_PSSLog_VOID',con, if_exists="replace", index_label='DataBase_ID')
                VIB_Rep_QC_Passed.to_sql('Eagle_PSSLog_QCPassed',con, if_exists="replace", index_label='DataBase_ID')
                VIB_Rep_QC_Failed.to_sql('Eagle_PSSLog_QCFailed',con, if_exists="replace", index_label='DataBase_ID')
                con.commit()
                cur.close()
                con.close()
                ViewTotalImport()
                txtSingleVoidReport.delete(0,END)
        else:
            tkinter.messagebox.showinfo("Import PSS File Message","Please Select PSS Files To Import")


    ## Adding File Menu 
    menu = Menu(window)
    window.config(menu=menu)
    filemenu  = Menu(menu, tearoff=0)       
    Advanced  = Menu(menu, tearoff=0)
    menu.add_cascade(label="File", menu=filemenu)                
    menu.add_cascade(label="Advanced", menu=Advanced)

    filemenu.add_command(label="Exit", command=iExit)
    Advanced.add_command(label="Fix Corrupted PSS From TB Import", command=FixedCorruptedPSSImport)



    ## DataFrame TOP 
    btnTotalRAWImport= Button(DataFrameTOP, text="Total Raw Imported", font=('aerial', 10, 'bold'),
                              bg = '#FFFAF0', height =1, width=16, bd=1, command = ViewTotalRAWImport)
    btnTotalRAWImport.grid(row =2, column = 0, sticky ="W", padx= 1)
    txtTotalRAWImport  = Entry(DataFrameTOP, font=('aerial', 10, 'bold'),textvariable=IntVar(), width = 10, bd=2)
    txtTotalRAWImport.grid(row =2, column = 1, sticky ="W", padx= 2)

    btnTotalValidAnalyzed = Button(DataFrameTOP, text="Total Valid Analyzed", font=('aerial', 10, 'bold'),
                                   bg = '#FFFAF0', height =1, width=18, bd=1, command = ViewTotalImport)
    btnTotalValidAnalyzed.grid(row =2, column = 3, sticky ="E", padx= 180)

    btnExportValidAnalyzedTestiFy = Button(DataFrameTOP, text="Export For Testif-i", font=('aerial', 10, 'bold'),
                                           bg = '#FFFAF0', height =1, width=15, bd=1, command = ExportValidForTest_i_Fy)
    btnExportValidAnalyzedTestiFy.grid(row =2, column = 3, sticky ="E", padx= 47)
    txtTotalEntries  = Entry(DataFrameTOP, font=('aerial', 10, 'bold'),textvariable=IntVar(), width = 10 , bd=2)
    txtTotalEntries.grid(row =2, column = 3, sticky ="E", padx= 340)


    txtTotalinvalidRemoved  = Entry(DataFrameTOP, font=('aerial', 10, 'bold'),textvariable=IntVar(), width = 10, bd=2)
    txtTotalinvalidRemoved.grid(row =2, column = 5, sticky ="W", padx= 140)
    btnTotalInValidRemoved = Button(DataFrameTOP, text="Total Invalid Removed", font=('aerial', 10, 'bold'),
                                   bg = '#FFFAF0', height =1, width=18, bd=1, command = ViewAllInvalidRemoved)
    btnTotalInValidRemoved.grid(row =2, column = 5, sticky ="E", padx= 220)
    DataFrameTOP.pack()

    ## DataFrame BOTTOM ACTIONS

    btnImportPSSLog = Button(DataFrameBOTTOM_ACTIONS, text="Import Raw PSS", font=('aerial', 10, 'bold'), height =1, width=14, bd=2, command = ImportPSSLogFile)
    btnImportPSSLog.grid(row =2, column = 0, sticky ="W", padx= 1)

    btnExportTotalRAWImport = Button(DataFrameBOTTOM_ACTIONS, text="Export Raw", font=('aerial', 10, 'bold'),
                                        height =1, width=9, bd=2, command = ExportRAW_ImportedPSS)
    btnExportTotalRAWImport.grid(row =2, column = 1, sticky ="E", padx= 1)

    btnShowQCLimit = Button(DataFrameBOTTOM_ACTIONS, text="View QC Limit", font=('aerial', 9, 'bold'), height =1, width=12, bd=2, command = ViewQCLimit)
    btnShowQCLimit.grid(row =2, column = 2, sticky ="W", padx= 1)
    btnSetupQCLimit = Button(DataFrameBOTTOM_ACTIONS, text="Setup QC Limit", font=('aerial', 9, 'bold'), height =1, width=12, bd=2, command = SetupVibQCLimitParameter)
    btnSetupQCLimit.grid(row =2, column = 3, sticky ="W", padx= 1)
    btnExportLB = Button(DataFrameBOTTOM_ACTIONS, text="Export ListBox", font=('aerial', 9, 'bold'), height =1, width=12, bd=2, command = ExportListBoxPSS)
    btnExportLB.grid(row =2, column = 4, sticky ="W", padx= 1)
    btnDelete = Button(DataFrameBOTTOM_ACTIONS, text="Delete SelectedValid", font=('aerial', 9, 'bold'), height =1, width=17, bd=2, command = DeleteSelectedImportData)
    btnDelete.grid(row =2, column = 5, sticky ="W", padx= 1)
    btnClearView = Button(DataFrameBOTTOM_ACTIONS, text="Clear View", font=('aerial', 9, 'bold'), height =1, width=10, bd=2, command = ClearView)
    btnClearView.grid(row =2, column = 6, sticky ="W", padx= 1)
    btnSortLineStation = Button(DataFrameBOTTOM_ACTIONS, text="Sort By Line-Station", font=('aerial', 9, 'bold'), height =1, width=16, bd=2, command = SortbyLineStation)
    btnSortLineStation.grid(row =2, column = 7, sticky ="W", padx= 1)
    btnExportForGlobalMapper = Button(DataFrameBOTTOM_ACTIONS, text="Export For GlobalMapper", font=('aerial', 9, 'bold'), height =1, width=20, bd=2, command = ExportValidForGlobalMapper)
    btnExportForGlobalMapper.grid(row =2, column = 8, sticky ="W", padx= 1)    
    Label_TimeOffset = Label(DataFrameBOTTOM_ACTIONS, text = "UTC Time Offset (Default: +7 hrs) :", font=("arial", 10,'bold'), bg = 'cadet blue')
    Label_TimeOffset.grid(row =2, column = 9, sticky ="W", padx= 70)
    Entrytxt_TimeOffset  = Entry(DataFrameBOTTOM_ACTIONS, font=('aerial', 12, 'bold'),textvariable=OffsetTimeUTC, width = 6)
    Entrytxt_TimeOffset.grid(row =2, column = 9, sticky ="W", padx= 300)

    DataFrameBOTTOM_ACTIONS.pack()


    ## DataFrame BOTTOM IFQC

    Label_txtVoidEntries = Label(DataFrameBOTTOM_IFQC, text = "Total Void Shots :", font=("arial", 10,'bold'), bg = 'cadet blue')
    Label_txtVoidEntries.grid(row =3, column = 0, sticky ="W", padx= 1, pady =5)
    txtVoidEntries  = Entry(DataFrameBOTTOM_IFQC, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
    txtVoidEntries.grid(row =3, column = 0, sticky ="W", padx= 170, pady =5)
    btnVoidEntries= Button(DataFrameBOTTOM_IFQC, text="View All Void Shots", font=('aerial', 9, 'bold'), bg = 'white', height =1, width=17, bd=2, command = ViewVoidShotIDImport)
    btnVoidEntries.grid(row =3, column = 0, sticky ="W", padx= 255, pady =5)
    btnGenVoidReport= Button(DataFrameBOTTOM_IFQC, text="ExportReport", font=('aerial', 9, 'bold'), bg = 'white', height =1, width=11, bd=2, command = ExportAllVoidImport)
    btnGenVoidReport.grid(row =3, column = 0, sticky ="W", padx= 390, pady =5)
    btnSingleVoidReport= Button(DataFrameBOTTOM_IFQC, text="SingleVoidShots", font=('aerial', 9, 'bold'), bg = 'white', height =1, width=13, bd=2, command = GenerateQCSingleVOIDReport)
    btnSingleVoidReport.grid(row =3, column = 0, sticky ="W", padx= 480, pady =5)
    txtSingleVoidReport  = Entry(DataFrameBOTTOM_IFQC, font=('aerial', 12, 'bold'),textvariable=None, width = 4)
    txtSingleVoidReport.grid(row =3, column = 0, sticky ="W", padx= 583, pady =5)

    Label_txtQCFailedEntries= Label(DataFrameBOTTOM_IFQC, text = "Total QC Failed Shots :", font=("arial", 10,'bold'), bg = 'cadet blue')
    Label_txtQCFailedEntries.grid(row =5, column = 0, sticky ="W", padx= 1, pady =5)
    txtQCFailedEntries  = Entry(DataFrameBOTTOM_IFQC, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
    txtQCFailedEntries.grid(row =5, column = 0, sticky ="W", padx= 170)
    btnQCFailedEntries= Button(DataFrameBOTTOM_IFQC, text="View All Failed Shots", font=('aerial', 9, 'bold'), bg = 'white', height =1, width=17, bd=2, command = ViewQCFailedImport)
    btnQCFailedEntries.grid(row =5, column = 0, sticky ="W", padx= 255, pady =5)
    btnGenQCFailedReport= Button(DataFrameBOTTOM_IFQC, text="ExportReport", font=('aerial', 9, 'bold'), bg = 'white', height =1, width=11, bd=2, command = GenerateQCFailedReport)
    btnGenQCFailedReport.grid(row =5, column = 0, sticky ="W", padx= 390, pady =5)
    btnGenQCFailedSingle= Button(DataFrameBOTTOM_IFQC, text="SingleQCFailed", font=('aerial', 9, 'bold'), bg = 'white', height =1, width=13, bd=2, command = GenerateQCSingleFailedReport)
    btnGenQCFailedSingle.grid(row =5, column = 0, sticky ="W", padx= 480, pady =5)
    txtSingleFailedReport  = Entry(DataFrameBOTTOM_IFQC, font=('aerial', 12, 'bold'),textvariable=None, width = 4)
    txtSingleFailedReport.grid(row =5, column = 0, sticky ="W", padx= 583, pady =5)

    Label_txtQCPassedEntries= Label(DataFrameBOTTOM_IFQC, text = "Total QC Passed Shots :", font=("arial", 10,'bold'), bg = 'cadet blue')
    Label_txtQCPassedEntries.grid(row =7, column = 0, sticky ="W", padx= 1, pady =5)
    txtQCPassedEntries  = Entry(DataFrameBOTTOM_IFQC, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
    txtQCPassedEntries.grid(row =7, column = 0, sticky ="W", padx= 170)
    btnQCPassedEntries= Button(DataFrameBOTTOM_IFQC, text="View Passed Shots", font=('aerial', 9, 'bold'), bg = 'white', height =1, width=17, bd=2, command = ViewQCPassedImport)
    btnQCPassedEntries.grid(row =7, column = 0, sticky ="W", padx= 255, pady =5)
    btnGenQCPassedReport= Button(DataFrameBOTTOM_IFQC, text="- Export All QC Passed Shots -", font=('aerial', 9, 'bold'), bg = 'white', height =1, width=26, bd=2, command = GenerateQCPassedReport)
    btnGenQCPassedReport.grid(row =7, column = 0, sticky ="W", padx= 390, pady =5)

    ## DataFrame BOTTOM INVAILD REMOVED

    Label_txtInvalidNullEntries = Label(DataFrameBOTTOM_INVALIDQC, text = "Total Invalid Null Shot ID / FFID :", font=("arial", 10,'bold'), bg = 'cadet blue')
    Label_txtInvalidNullEntries.grid(row =3, column = 0, sticky ="W", padx= 1, pady =5)
    txtInvalidEntries  = Entry(DataFrameBOTTOM_INVALIDQC, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
    txtInvalidEntries.grid(row =3, column = 0, sticky ="W", padx= 275, pady =5)
    btnInvalidNullEntries= Button(DataFrameBOTTOM_INVALIDQC, text="View Invalid Null Shot ID/FFID", font=('aerial', 9, 'bold'), bg = '#FF3030', height =1, width=25, bd=2, command = ViewInvalidImport)
    btnInvalidNullEntries.grid(row =3, column = 0, sticky ="W", padx= 360, pady =5)

    Label_txtDuplicatedShotID= Label(DataFrameBOTTOM_INVALIDQC, text = "Total Invalid Duplicated Shot ID/FFID :", font=("arial", 10,'bold'), bg = 'cadet blue')
    Label_txtDuplicatedShotID.grid(row =5, column = 0, sticky ="W", padx= 1, pady =5)
    txtDuplicatedShotID  = Entry(DataFrameBOTTOM_INVALIDQC, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
    txtDuplicatedShotID.grid(row =5, column = 0, sticky ="W", padx= 275, pady =5)
    btnDuplicatedShotIDImport= Button(DataFrameBOTTOM_INVALIDQC, text="View Invalid Duplicated ShotID", font=('aerial', 9, 'bold'), bg = '#FF3030', height =1, width=25, bd=2, command = ViewDuplicatedShotIDImport)
    btnDuplicatedShotIDImport.grid(row =5, column = 0, sticky ="W", padx= 360, pady =5)

    Label_DuplicatedShotExchange= Label(DataFrameBOTTOM_INVALIDQC, text = "Update  Duplicated Shot ID/FFID:", font=("arial", 10,'bold'), bg = 'cadet blue')
    Label_DuplicatedShotExchange.grid(row =7, column = 0, sticky ="W", padx= 1, pady =5)
    txtDuplicatedShotExchange  = Entry(DataFrameBOTTOM_INVALIDQC, font=('aerial', 12, 'bold'),textvariable=None, width = 8)
    txtDuplicatedShotExchange.grid(row =7, column = 0, sticky ="W", padx= 275)
    btnDuplicatedShotExchange= Button(DataFrameBOTTOM_INVALIDQC, text="Update Duplicated Shot ID", font=('aerial', 9, 'bold'), bg = '#FF3030', height =1, width=25, bd=2, command = UpdateDuplicatedShotID)
    btnDuplicatedShotExchange.grid(row =7, column = 0, sticky ="W", padx= 360, pady =5)


