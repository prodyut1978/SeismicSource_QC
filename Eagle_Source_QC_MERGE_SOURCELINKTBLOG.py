#Front End
import os
from tkinter import*
import tkinter.messagebox
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
import Eagle_SourceLink_Vibroseis_Log_BackEnd
import Eagle_SourceLink_TB_Vibroseis_Import_Module
import Eagle_SourceLink_PSS_Import_Module
import Eagle_SourceLink_VIBPositionImport_Module

def Merge_Sourcelink_TBLog():
    Vib_QC_path = r'C:\VIB_QC_Report'
    if not os.path.exists(Vib_QC_path):
        os.makedirs(Vib_QC_path)

    Default_Date_today   = datetime.date.today()
    try:
        ### Connect To Sourcelink TB DB
        connSOURCELINK_TBLog = sqlite3.connect("SourceLink_Log.db")
        Complete_SOURCELINK_TBLog = pd.read_sql_query("SELECT * FROM Eagle_SOURCELINK_TB_TEMP ORDER BY `FileNum` ASC ;", connSOURCELINK_TBLog)
        Complete_SOURCELINK_TBLog["Sourcelink TB < - > PSS & VibPosition"] = Complete_SOURCELINK_TBLog.shape[0]*["<<TB>> ---- << PSS & VibPosition >>"]
        Complete_SOURCELINK_TBlOG_DF = pd.DataFrame(Complete_SOURCELINK_TBLog)            
        Complete_SOURCELINK_TBlOG_DF = Complete_SOURCELINK_TBLog.reset_index(drop=True)
        Length_SOURCELINK_TBlOG_DF   = len(Complete_SOURCELINK_TBLog)

        ### Connect To PSS DB
        connSourceLink_PSSLog = sqlite3.connect("SourceLink_Log.db")
        Complete_SourceLink_PSSLog = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_TEMP ORDER BY `FileNum` ASC ;", connSourceLink_PSSLog)
        Complete_SourceLink_PSSLog = Complete_SourceLink_PSSLog.loc[:,['ShotID','FileNum','EPNumber','SourceLine','SourceStation',
                                                                       'Unit_ID','SwCksm','PmCksm', 'PhaseMax','PhaseAvg',
                                                                       'ForceMax','ForceAvg','THDMax','THDAvg',
                                                                       'Force_Out','GPS_Quality']]
        Complete_SourceLink_PSSLog_DF = pd.DataFrame(Complete_SourceLink_PSSLog)            
        Complete_SourceLink_PSSLog_DF = Complete_SourceLink_PSSLog_DF.reset_index(drop=True)
        Length_SourceLink_PSSLog_DF   = len(Complete_SourceLink_PSSLog_DF)


         ### Connect To Vib Position DB
        connSourceLink_VIB_POSITION       = sqlite3.connect("SourceLink_Log.db")
        Complete_SourceLink_VIB_POSITION  = pd.read_sql_query("SELECT * FROM Eagle_VIB_COG_TEMP ORDER BY `FileNum` ASC ;", connSourceLink_VIB_POSITION)
        Complete_SourceLink_VIB_POSITION  = Complete_SourceLink_VIB_POSITION.loc[:,['ShotID','FileNum','EPNumber','SourceLine','SourceStation',
                                                                       'Unit_ID','DistanceCOG','NearFlagLine', 'NearFlagStation','DistanceNearFlag',
                                                                       'GPS_Quality','Near_Flag_Message']]
        Complete_SourceLink_VIB_POSITION  = pd.DataFrame(Complete_SourceLink_VIB_POSITION) 
        Complete_SourceLink_VIB_POSITION  = Complete_SourceLink_VIB_POSITION.reset_index(drop=True)
        Length_SourceLink_VIB_POSITION_DF = len(Complete_SourceLink_VIB_POSITION)

        ## Merging Summary
        Merging_Summary =  ("Merging Summary: " + '\n' + 
                            "Total Sourcelink Timebreak Entries: " + str(Length_SOURCELINK_TBlOG_DF) + '\n' + 
                            "Total PSS Entries: " + str(Length_SourceLink_PSSLog_DF)+ '\n' +
                            "Total VIB Position Entries: " + str(Length_SourceLink_VIB_POSITION_DF))


        ## Merging TB and PSS DB    
        Merge_SourceLink_TB_PSS = pd.merge(Complete_SOURCELINK_TBlOG_DF, Complete_SourceLink_PSSLog_DF, on =['FileNum', 'EPNumber', 'SourceLine', 'SourceStation', 'Unit_ID'] ,how ='outer',
                   left_index = False, right_index = False, sort = True, indicator = True )

        def trans_SHOT_MISSING(x):
            if x   == 'both':
                return np.nan

            elif x == 'right_only':
                return 'TB MISSING SHOT'

            elif x == 'left_only':
                return 'PSS MISSING SHOT'
            
            else:
                return x      

        Merge_SourceLink_TB_PSS['_merge']   = Merge_SourceLink_TB_PSS['_merge'].apply(trans_SHOT_MISSING)
        Merge_SourceLink_TB_PSS.rename(columns={'_merge':'PSS Flags'},inplace = True)
        Merge_SourceLink_TB_PSS             = Merge_SourceLink_TB_PSS.reset_index(drop=True)
        Merge_SourceLink_TB_PSS             = pd.DataFrame(Merge_SourceLink_TB_PSS) 

        ## Merging TB and PSS And VIB Position DB
        Merge_SourceLink_TB_PSS_VIBPosition = pd.merge(Merge_SourceLink_TB_PSS, Complete_SourceLink_VIB_POSITION, on =['FileNum', 'EPNumber', 'SourceLine', 'SourceStation', 'Unit_ID'] ,how ='outer', suffixes =('_VIB PSS', '_VIB Position'),
                   left_index = False, right_index = False, sort = True, indicator = True )

        def trans_VIBPosition_MISSING(x):
            if x   == 'both':
                return np.nan 

            elif x == 'right_only':
                return 'PSS-TB SHOT MISSING'

            elif x == 'left_only':
                return 'MISSING_VIB_POSITION'
            
            else:
                return x      

        Merge_SourceLink_TB_PSS_VIBPosition['_merge']     = Merge_SourceLink_TB_PSS_VIBPosition['_merge'].apply(trans_VIBPosition_MISSING)
        Merge_SourceLink_TB_PSS_VIBPosition.rename(columns={'_merge':'Position Flags'},inplace = True)        
        Merge_SourceLink_TB_PSS_VIBPosition.rename(columns={'TriggerIndex':'TriggerIndex', 'Unit_ID':' ProfileId', 'FileNum':' ShotNumber',
                                                            'EPNumber':' EpNumber', 'SourceLine':' ShotLine', 'SourceStation':' ShotStation',
                                                            'ShotUtcDateTime':' ShotUtcDateTime','Latitude':' Latitude','Longitude':' Longitude',
                                                            'ShotStatus':' ShotStatus', 'TBComment':' Comment'},inplace = True)
        Merge_SourceLink_TB_PSS_VIBPosition  = Merge_SourceLink_TB_PSS_VIBPosition.reset_index(drop=True)
        Merge_SourceLink_TB_PSS_VIBPosition  = pd.DataFrame(Merge_SourceLink_TB_PSS_VIBPosition)

        MergedDuplicatedLineStation = Merge_SourceLink_TB_PSS_VIBPosition.drop_duplicates([' ShotNumber'],keep='last')
        MergedDuplicatedLineStation = MergedDuplicatedLineStation.reset_index(drop=True)
        MergedDuplicatedLineStation ['DuplicatedShot']= MergedDuplicatedLineStation.sort_values(by =[' ShotLine',' ShotStation']).duplicated([' ShotLine',' ShotStation'], keep=False)
        MergedDuplicatedLineStation = MergedDuplicatedLineStation.loc[MergedDuplicatedLineStation.DuplicatedShot == True]
        MergedDuplicatedLineStation  = MergedDuplicatedLineStation.loc[:,[' ShotNumber',' ShotLine',' ShotStation']]
        MergedDuplicatedLineStation = MergedDuplicatedLineStation.reset_index(drop=True)
        Merge_SourceLink_TB_PSS_VIBPosition = pd.merge(Merge_SourceLink_TB_PSS_VIBPosition, MergedDuplicatedLineStation, on =[' ShotNumber',' ShotLine',' ShotStation'] ,how ='left', indicator = True )

        def trans_DuplicatedCheck(x):
            if x   == 'both':
                return 'DuplicateShot'

            elif x == 'left_only':
                return np.nan
            
            else:
                return x

        Merge_SourceLink_TB_PSS_VIBPosition['_merge']     = Merge_SourceLink_TB_PSS_VIBPosition['_merge'].apply(trans_DuplicatedCheck)
        Merge_SourceLink_TB_PSS_VIBPosition.rename(columns={'_merge':'DuplicateCheck'},inplace = True)

        outfile_Merge_TB_PSS_VIBPosition=("C:\\VIB_QC_Report\\Combined_Sourcelink TB_MERGE_Report.csv")
        Merge_SourceLink_TB_PSS_VIBPosition.to_csv(outfile_Merge_TB_PSS_VIBPosition,index=None)
        
        connSOURCELINK_TBLog.commit()
        connSOURCELINK_TBLog.close()
        connSourceLink_PSSLog.commit()
        connSourceLink_PSSLog.close()
        connSourceLink_VIB_POSITION.commit()
        connSourceLink_VIB_POSITION.close()
        tkinter.messagebox.showinfo("Merging Database Complete Message", Merging_Summary)

        ## Export Report         
        filename = tkinter.filedialog.asksaveasfilename(initialdir = "/" ,title = "Select File Name To Export Output" ,\
                           defaultextension='.csv', filetypes = (("CSV file",".csv"),("Excel file",".xlsx")))
        if len(filename) >0:
            if filename.endswith('.csv'):
                Merge_SourceLink_TB_PSS_VIBPosition.to_csv(filename,index=None)
                tkinter.messagebox.showinfo("Merged Report Export For QC","SourceLink PSS and TB and VIB Position Merged Report Saved as CSV")
            else:
                Merge_SourceLink_TB_PSS_VIBPosition.to_excel(filename, sheet_name='Merged PSS-TB-Vib Position', index=False)
                tkinter.messagebox.showinfo("Merged Report Export For QC","SourceLink PSS and TB and VIB Position Merged Report Saved as Excel")
        else:
            tkinter.messagebox.showinfo("Export Message","Please Select File Name To Export Merged Report")


    except:
        tkinter.messagebox.showinfo("Error In Merging Databases","Please Check Imported TB Log Or PSS Log Or Vib Position Log To Merge Correctly")
        
