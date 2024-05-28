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
import Eagle_HAWK_OBLog_BackEnd
import Eagle_HAWK_OBLog_Import_Module
import Eagle_SourceLink_Vibroseis_Log_BackEnd
import Eagle_SourceLink_PSS_Import_Module
import Eagle_SourceLink_VIBPositionImport_Module

def Merge_HAWK_OBLog():
    Vib_QC_path = r'C:\VIB_QC_Report'
    if not os.path.exists(Vib_QC_path):
        os.makedirs(Vib_QC_path)

    Default_Date_today   = datetime.date.today()
    try:
        ### Connect To HAWK TB DB
        connHAWK_OBLog = sqlite3.connect("HAWK_OBLog.db")
        Complete_HAWK_OBLog = pd.read_sql_query("SELECT * FROM Eagle_HAWK_OBLog_TEMP ORDER BY `ShotID` ASC ;", connHAWK_OBLog)
        Complete_HAWK_OBLog["INOVA TB < - > SOURCELINK PSS & VIB POSITION"] = Complete_HAWK_OBLog.shape[0]*["<< INOVA TB>> ---- << SOURCELINK PSS & VIB POSITION >>"]
        Complete_HAWK_OBlOG_DF = pd.DataFrame(Complete_HAWK_OBLog)            
        Complete_HAWK_OBlOG_DF = Complete_HAWK_OBlOG_DF.reset_index(drop=True)
        Length_HAWK_OBlOG_DF   = len(Complete_HAWK_OBlOG_DF)

        ### Connect To PSS DB
        connSourceLink_PSSLog = sqlite3.connect("SourceLink_Log.db")
        Complete_SourceLink_PSSLog = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_TEMP ORDER BY `ShotID` ASC ;", connSourceLink_PSSLog)
        Complete_SourceLink_PSSLog = Complete_SourceLink_PSSLog.loc[:,['ShotID','FileNum','EPNumber','SourceLine','SourceStation',
                                                                       'Unit_ID','SwCksm','PmCksm', 'PhaseMax','PhaseAvg',
                                                                       'ForceMax','ForceAvg','THDMax','THDAvg',
                                                                       'Force_Out','GPS_Quality']]
        Complete_SourceLink_PSSLog_DF = pd.DataFrame(Complete_SourceLink_PSSLog)            
        Complete_SourceLink_PSSLog_DF = Complete_SourceLink_PSSLog_DF.reset_index(drop=True)
        Length_SourceLink_PSSLog_DF   = len(Complete_SourceLink_PSSLog_DF)


         ### Connect To Vib Position DB
        connSourceLink_VIB_POSITION       = sqlite3.connect("SourceLink_Log.db")
        Complete_SourceLink_VIB_POSITION  = pd.read_sql_query("SELECT * FROM Eagle_VIB_COG_TEMP ORDER BY `ShotID` ASC ;", connSourceLink_VIB_POSITION)
        Complete_SourceLink_VIB_POSITION  = Complete_SourceLink_VIB_POSITION.loc[:,['ShotID','FileNum','EPNumber','SourceLine','SourceStation',
                                                                       'Unit_ID','DistanceCOG','NearFlagLine', 'NearFlagStation','DistanceNearFlag',
                                                                       'GPS_Quality','Near_Flag_Message']]
        Complete_SourceLink_VIB_POSITION  = pd.DataFrame(Complete_SourceLink_VIB_POSITION) 
        Complete_SourceLink_VIB_POSITION  = Complete_SourceLink_VIB_POSITION.reset_index(drop=True)
        Length_SourceLink_VIB_POSITION_DF = len(Complete_SourceLink_VIB_POSITION)

        ## Merging Summary
        Merging_Summary =  ("Merging Summary: " + '\n' + 
                            "Total HAWK Timebreak Entries: " + str(Length_HAWK_OBlOG_DF) + '\n' + 
                            "Total PSS Entries: " + str(Length_SourceLink_PSSLog_DF)+ '\n' +
                            "Total VIB Position Entries: " + str(Length_SourceLink_VIB_POSITION_DF))


        ## Merging HAWK TB and PSS DB    
        Merge_HAWK_TB_PSS = pd.merge(Complete_HAWK_OBlOG_DF, Complete_SourceLink_PSSLog_DF, on =['ShotID','EPNumber','SourceLine','SourceStation'] ,how ='outer',
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

        Merge_HAWK_TB_PSS['_merge']   = Merge_HAWK_TB_PSS['_merge'].apply(trans_SHOT_MISSING)
        Merge_HAWK_TB_PSS.rename(columns={'_merge':'PSS Flags'},inplace = True)
        Merge_HAWK_TB_PSS             = Merge_HAWK_TB_PSS.reset_index(drop=True)
        Merge_HAWK_TB_PSS             = pd.DataFrame(Merge_HAWK_TB_PSS)
       
        ## Merging HAWK TB and PSS And VIB Position DB
        Merge_HAWK_TB_PSS_VIBPosition = pd.merge(Merge_HAWK_TB_PSS, Complete_SourceLink_VIB_POSITION, on =['ShotID','FileNum','EPNumber','SourceLine','SourceStation','Unit_ID'] ,how ='outer', suffixes =('_VIB PSS', '_VIB Position'),
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

        Merge_HAWK_TB_PSS_VIBPosition['_merge']     = Merge_HAWK_TB_PSS_VIBPosition['_merge'].apply(trans_VIBPosition_MISSING)
        Merge_HAWK_TB_PSS_VIBPosition.rename(columns={'_merge':'Position Flags'},inplace = True)
        Merge_HAWK_TB_PSS_VIBPosition  = Merge_HAWK_TB_PSS_VIBPosition.reset_index(drop=True)
        Merge_HAWK_TB_PSS_VIBPosition.rename(columns = {'MasterSystemFieldRecordID':'Master System Field Record ID', 'EPNumber':'EP Number',
        'ShotID':'Shot ID', 'Omit':'Omit', 'FileType':'File Type', 'File_CorrBeforeStack':'File - Corr Before Stack',
        'File_CorrAfterStack':'File - Corr After Stack', 'File_UncorrStack':'File - Uncorr Stack', 'File_CorrEP':'File - Corr EP',
        'File_UncorrEP':'File - Uncorr EP', 'Timebreak_SecondUnixTimeStamp':'Timebreak Second (Unix TimeStamp or DateTime)',
        'Timebreak_mSecs':'Timebreak (mSecs)', 'Timebreak_uSecs':'Timebreak (uSecs)', 'RecordLength_mSecs':'Record Length (mSecs)',
        'Acquisition_Time_mSecs':'Acquisition Time (mSecs)', 'SourceLine':'Source Line', 'SourceStation':'Source Station',
        'SourceType_DynamiteorVibroseis':'Source Type (Dynamite or Vibroseis)', 'Vibes64_bit_mask':'Vibes (64-bit mask)',
        'SampleRateuSecs':'Sample Rate (uSecs)', 'SourceX':'Source X', 'SourceY':'Source Y', 'SourceZ':'Source Z', 'GridUnits':'Grid Units',
        'SweepFile':'Sweep File', 'SweepID':'Sweep ID', 'SweepType':'Sweep Type (ShotPro. Linear. dbHz. dbOct. etc)', 'SweepStartFrequency':'Sweep Start Frequency (Hz)',
        'SweepEndFrequency':'Sweep End Frequency (Hz)', 'SweepLength':'Sweep Length (mSecs)', 'TaperType':'Taper Type (BlackMan or Cosine)',
        'StartTaperDuration':'Start Taper Duration (mSecs)', 'EndTaperDuration':'End Taper Duration (mSecs)',
        'Comment':'Comment'},inplace = True)     
        Merge_HAWK_TB_PSS_VIBPosition  = pd.DataFrame(Merge_HAWK_TB_PSS_VIBPosition)

        MergedDuplicatedLineStation = Merge_HAWK_TB_PSS_VIBPosition.drop_duplicates(['Master System Field Record ID'],keep='last')
        MergedDuplicatedLineStation = MergedDuplicatedLineStation.reset_index(drop=True)
        MergedDuplicatedLineStation ['DuplicatedShot']= MergedDuplicatedLineStation.sort_values(by =['Source Line','Source Station']).duplicated(['Source Line','Source Station'], keep=False)
        MergedDuplicatedLineStation = MergedDuplicatedLineStation.loc[MergedDuplicatedLineStation.DuplicatedShot == True]
        MergedDuplicatedLineStation  = MergedDuplicatedLineStation.loc[:,['Master System Field Record ID','Source Line','Source Station']]
        MergedDuplicatedLineStation = MergedDuplicatedLineStation.reset_index(drop=True)
        Merge_HAWK_TB_PSS_VIBPosition = pd.merge(Merge_HAWK_TB_PSS_VIBPosition, MergedDuplicatedLineStation, on =['Master System Field Record ID','Source Line','Source Station'] ,how ='left', indicator = True )

        def trans_DuplicatedCheck(x):
            if x   == 'both':
                return 'DuplicateShot'

            elif x == 'left_only':
                return np.nan
            
            else:
                return x

        Merge_HAWK_TB_PSS_VIBPosition['_merge']     = Merge_HAWK_TB_PSS_VIBPosition['_merge'].apply(trans_DuplicatedCheck)
        Merge_HAWK_TB_PSS_VIBPosition.rename(columns={'_merge':'DuplicateCheck'},inplace = True)

        outfile_Merge_HAWK_TB_PSS=("C:\\VIB_QC_Report\\Combined_HAWK TB_MERGE_Report.csv")
        Merge_HAWK_TB_PSS_VIBPosition.to_csv(outfile_Merge_HAWK_TB_PSS,index=None)

        connHAWK_OBLog.commit()
        connHAWK_OBLog.close()
        connSourceLink_PSSLog.commit()
        connSourceLink_PSSLog.close()
        connSourceLink_VIB_POSITION.commit()
        connSourceLink_VIB_POSITION.close()
        tkinter.messagebox.showinfo("Merging Database Complete Message", Merging_Summary)

        ## Export report        
        filename = tkinter.filedialog.asksaveasfilename(initialdir = "/" ,title = "Select File Name To Export Output" ,\
                           defaultextension='.csv', filetypes = (("CSV file",".csv"),("Excel file",".xlsx")))
        if len(filename) >0:
            if filename.endswith('.csv'):
                Merge_HAWK_TB_PSS_VIBPosition.to_csv(filename,index=None)
                tkinter.messagebox.showinfo("Merged Report Export For QC","PSS HAWK TB And VIB Position Merged Report Saved as CSV")
            else:
                Merge_HAWK_TB_PSS_VIBPosition.to_excel(filename, sheet_name='Merged PSS-TB-Vib Position', index=False)
                tkinter.messagebox.showinfo("Merged Report Export For QC","PSS HAWK TB And VIB Position Merged Report Saved as Excel")
        else:
            tkinter.messagebox.showinfo("Export Message","Please Select File Name To Export Merged Report")


    except:
        tkinter.messagebox.showinfo("Error In Merging Databases","Please Check Imported TB Log Or PSS Log Or Vib Position Log To Merge Correctly")
            
