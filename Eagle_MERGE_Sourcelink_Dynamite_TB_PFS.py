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
import Eagle_SourceLink_Dynamite_Log_BackEnd
import Eagle_SourceLink_TB_Dynamite_Import_Module
import Eagle_SourceLink_PFS_Import_Module


def Merge_Sourcelink_Dynamite_TBLog():
    Dynamite_QC_path = r'C:\Dynamite_QC_Report'
    if not os.path.exists(Dynamite_QC_path):
        os.makedirs(Dynamite_QC_path)

    Default_Date_today   = datetime.date.today()
    try:
        ### Connect To Sourcelink Dynamite TB DB
        connSOURCELINK_TBLog = sqlite3.connect("SourceLink_Dynamite_Log.db")
        Complete_SOURCELINK_TBLog = pd.read_sql_query("SELECT * FROM Eagle_SOURCELINK_TB_Dynamite ORDER BY `FileNum` ASC ;", connSOURCELINK_TBLog)
        Complete_SOURCELINK_TBLog["Sourcelink TB < - > Sourcelink PFS"] = Complete_SOURCELINK_TBLog.shape[0]*["<<TB>> ---- << PFS >>"]
        Complete_SOURCELINK_TBlOG_DF = pd.DataFrame(Complete_SOURCELINK_TBLog)            
        Complete_SOURCELINK_TBlOG_DF = Complete_SOURCELINK_TBLog.reset_index(drop=True)
        Length_SOURCELINK_TBlOG_DF   = len(Complete_SOURCELINK_TBLog)

        ### Connect To PFS DB
        connSourceLink_PFSLog = sqlite3.connect("SourceLink_Dynamite_Log.db")
        Complete_SourceLink_PFSLog = pd.read_sql_query("SELECT * FROM Eagle_PFSLog_TEMP ORDER BY `FileNum` ASC ;", connSourceLink_PFSLog)
        Complete_SourceLink_PFSLog = Complete_SourceLink_PFSLog.loc[:,['ShotID','FileNum','EPNumber','SourceLine','SourceStation',
                                                                       'Unit_ID','Battery','CapRes', 'GeoRes', 'FlagNum', 'UpholeWindow', 'FiredOK',
                                                                       'BatteryOK', 'GeoOK', 'CapOK','Timebreak', 'FirstBreak',
                                                                       'CapSerialNumber','GPS_Time','GPS_Quality', 'Latitude', 'Longitude', 'Altitude']]
        Complete_SourceLink_PFSLog_DF = pd.DataFrame(Complete_SourceLink_PFSLog)            
        Complete_SourceLink_PFSLog_DF = Complete_SourceLink_PFSLog_DF.reset_index(drop=True)
        Length_SourceLink_PFSLog_DF   = len(Complete_SourceLink_PFSLog_DF)


        ## Merging Summary
        Merging_Summary =  ("Merging Summary: " + '\n' + 
                            "Total Sourcelink Timebreak Entries: " + str(Length_SOURCELINK_TBlOG_DF) + '\n' + 
                            "Total PFS Entries: " + str(Length_SourceLink_PFSLog_DF))


        ## Merging TB and PFS DB    
        Merge_SourceLink_TB_PFS = pd.merge(Complete_SOURCELINK_TBlOG_DF, Complete_SourceLink_PFSLog_DF, on =['FileNum', 'EPNumber', 'SourceLine', 'SourceStation'] ,how ='outer',
                   left_index = False, right_index = False, sort = True, indicator = True )

        def trans_SHOT_MISSING(x):
            if x   == 'both':
                return np.nan

            elif x == 'right_only':
                return 'TB MISSING SHOT'

            elif x == 'left_only':
                return 'PFS MISSING SHOT'
            
            else:
                return x      

        Merge_SourceLink_TB_PFS['_merge']   = Merge_SourceLink_TB_PFS['_merge'].apply(trans_SHOT_MISSING)
        Merge_SourceLink_TB_PFS.rename(columns={'_merge':'PFS Flags'},inplace = True)
        Merge_SourceLink_TB_PFS             = Merge_SourceLink_TB_PFS.reset_index(drop=True)
        Merge_SourceLink_TB_PFS             = pd.DataFrame(Merge_SourceLink_TB_PFS)

        Merge_SourceLink_TB_PFS.rename(columns={'TriggerIndex':'TriggerIndex', 'Unit_ID_x':' ProfileId', 'FileNum':' ShotNumber',
                                                'EPNumber':' EpNumber', 'SourceLine':' ShotLine', 'SourceStation':' ShotStation',
                                                'ShotUtcDateTime':' ShotUtcDateTime','Latitude':' Latitude','Longitude':' Longitude',
                                                'ShotStatus':' ShotStatus', 'Uphole':' Uphole', 'TBComment':' Comment',
                                                'Process':' Process', 'Unit_ID_y':'Unit ID', 'Latitude_y':'PFS_Lat', 'Longitude_y':'PFS_Lon'},inplace = True)
        Merge_SourceLink_TB_PFS  = Merge_SourceLink_TB_PFS.reset_index(drop=True)
        Merge_SourceLink_TB_PFS  = pd.DataFrame(Merge_SourceLink_TB_PFS)

        MergedDuplicatedLineStation = Merge_SourceLink_TB_PFS.drop_duplicates([' ShotNumber'],keep='last')
        MergedDuplicatedLineStation = MergedDuplicatedLineStation.reset_index(drop=True)
        MergedDuplicatedLineStation ['DuplicatedShot']= MergedDuplicatedLineStation.sort_values(by =[' ShotLine',' ShotStation']).duplicated([' ShotLine',' ShotStation'], keep=False)
        MergedDuplicatedLineStation = MergedDuplicatedLineStation.loc[MergedDuplicatedLineStation.DuplicatedShot == True]
        MergedDuplicatedLineStation  = MergedDuplicatedLineStation.loc[:,[' ShotNumber',' ShotLine',' ShotStation']]
        MergedDuplicatedLineStation = MergedDuplicatedLineStation.reset_index(drop=True)
        Merge_SourceLink_TB_PFS     = pd.merge(Merge_SourceLink_TB_PFS, MergedDuplicatedLineStation, on =[' ShotNumber',' ShotLine',' ShotStation'] ,how ='left', indicator = True )

        def trans_DuplicatedCheck(x):
            if x   == 'both':
                return 'DuplicateShot'

            elif x == 'left_only':
                return np.nan
            
            else:
                return x

        Merge_SourceLink_TB_PFS['_merge']     = Merge_SourceLink_TB_PFS['_merge'].apply(trans_DuplicatedCheck)
        Merge_SourceLink_TB_PFS.rename(columns={'_merge':'DuplicateCheck'},inplace = True)
        Merge_SourceLink_TB_PFS             = Merge_SourceLink_TB_PFS.reset_index(drop=True)
        Merge_SourceLink_TB_PFS             = pd.DataFrame(Merge_SourceLink_TB_PFS) 

        outfile_Merge_SourceLink_TB_PFS=("C:\\Dynamite_QC_Report\\Combined_Sourcelink TB_MERGE_Report.csv")
        Merge_SourceLink_TB_PFS.to_csv(outfile_Merge_SourceLink_TB_PFS,index=None)

        connSOURCELINK_TBLog.commit()
        connSOURCELINK_TBLog.close()
        connSourceLink_PFSLog.commit()
        connSourceLink_PFSLog.close()

        tkinter.messagebox.showinfo("Merging PFS -TB Database Complete Message", Merging_Summary)

        ## Export Report         
        filename = tkinter.filedialog.asksaveasfilename(initialdir = "/" ,title = "Select File Name To Export Output" ,\
                           defaultextension='.csv', filetypes = (("CSV file",".csv"),("Excel file",".xlsx")))
        if len(filename) >0:
            if filename.endswith('.csv'):
                Merge_SourceLink_TB_PFS.to_csv(filename,index=None)
                tkinter.messagebox.showinfo("Merged Report Export For QC","SourceLink Dynamite PFS and TB Merged Report Saved as CSV")
            else:
                Merge_SourceLink_TB_PFS.to_excel(filename, sheet_name='Merged PFS-TB', index=False)
                tkinter.messagebox.showinfo("Merged Report Export For QC","SourceLink Dynamite PFS and TB Merged Report Saved as Excel")
        else:
            tkinter.messagebox.showinfo("Export Message","Please Select File Name To Export Merged Report")


    except:
        tkinter.messagebox.showinfo("Error In Merging Databases","Please Check Imported TB Log Or PFS Log To Merge Correctly")
        
