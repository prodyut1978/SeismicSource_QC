import os
import PySimpleGUI as sg
from tkinter import*
import tkinter.messagebox
import Eagle_Sourcefile_Dynamite_SPS_BackEnd
import Eagle_SourceLink_Dynamite_Log_BackEnd
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

def GenerateSourceLinkTB_Dynamite_MissingShotReport():    
    con= sqlite3.connect("DynamiteSourceSPS.db")
    cur=con.cursor()
    Complete_df_SPS = pd.read_sql_query("SELECT * FROM SourceFileSPS ORDER BY `SourceLineStationCombined` ASC;", con)
    MasterSPS_DF_Merge = pd.DataFrame(Complete_df_SPS)
    con.commit()
    cur.close()
    con.close()

    conn = sqlite3.connect("SourceLink_Dynamite_Log.db")
    cur=conn.cursor()
    Complete_df  = pd.read_sql_query("SELECT * FROM Eagle_SOURCELINK_DYNAMITE_TBMASTER ORDER BY `FileNum` ASC ;", conn)
    MasterTB_DF  = pd.DataFrame(Complete_df)
    conn.commit()
    cur.close()
    conn.close()

    if (len(MasterSPS_DF_Merge) >0) & (len(MasterTB_DF) >0):
        MasterSPS_DF_Merge = MasterSPS_DF_Merge.drop_duplicates(['SourceLineStationCombined'],keep='last')
        MasterSPS_DF_Merge = MasterSPS_DF_Merge.reset_index(drop=True)
        MasterTB_DF  = MasterTB_DF.loc[:,['FileNum','SourceLine','SourceStation', 'ShotStatus', 'TBComment']]
        MasterTB_DF  = MasterTB_DF.reset_index(drop=True)
        MasterTB_DF  = pd.DataFrame(MasterTB_DF)
        MasterTB_DF_Line_M      = MasterTB_DF['SourceLine'].astype(int)
        MasterTB_DF_Station_M   = MasterTB_DF['SourceStation'].astype(float)
        MasterTBLine_Station_Combined        = (MasterTB_DF_Line_M.map(str) + MasterTB_DF_Station_M.map(str)).astype(float)
        CombinedMasterTBLine_Station         =  pd.DataFrame(MasterTBLine_Station_Combined)
        CombinedMasterTBLine_Station.rename(columns={0:'SourceLineStationCombined'},inplace = True)
        MasterTB_DF_Merge                              = pd.concat([MasterTB_DF,CombinedMasterTBLine_Station],axis=1)
        MasterTB_DF_Merge['SourceLineStationCombined'] = MasterTB_DF_Merge['SourceLineStationCombined'].astype(float)
        MasterTB_DF_Merge                              = MasterTB_DF_Merge.drop_duplicates(['SourceLineStationCombined'],keep='last')
        MasterTB_DF_Merge                              = MasterTB_DF_Merge.reset_index(drop=True)

        # Production Statistics
        Production_Line_Station_Minimum    = (MasterTB_DF_Merge.groupby('SourceLine').SourceStation.min()).astype(float)
        Production_Line_Station_Maximum    = (MasterTB_DF_Merge.groupby('SourceLine').SourceStation.max()).astype(float)
        Production_Station_Count           = (MasterTB_DF_Merge.groupby('SourceLine').SourceStation.count()).astype(int)
        Production_Statistics              = [Production_Line_Station_Minimum, Production_Line_Station_Maximum, Production_Station_Count]
        Production_Statistics_Combined     = pd.concat(Production_Statistics,axis=1,ignore_index =True)
        Production_Statistics_Combined.reset_index(inplace=True)
        Production_Statistics_Combined.rename(columns={0:'ProductionStartSP', 1:'ProductionEndSP',2:'CompletedShotsCount'},inplace = True)

        ## SPS Statistics
        SPS_Line_Station_Minimum    = (MasterSPS_DF_Merge.groupby('SourceLine').SourceStation.min()).astype(float)
        SPS_Line_Station_Maximum    = (MasterSPS_DF_Merge.groupby('SourceLine').SourceStation.max()).astype(float)
        SPS_Station_Count           = (MasterSPS_DF_Merge.groupby('SourceLine').SourceStation.count()).astype(int)
        SPS_Statistics              = [SPS_Line_Station_Minimum, SPS_Line_Station_Maximum, SPS_Station_Count]
        SPS_Statistics_Combined     = pd.concat(SPS_Statistics,axis=1,ignore_index =True)
        SPS_Statistics_Combined.reset_index(inplace=True)
        SPS_Statistics_Combined.rename(columns={0:'SPSBeginPoint', 1:'SPSEndPoint',2:'SPSTotalShotsCount'},inplace = True)

        ## Incomplete Production Lines
        IncompleteProductionLine       = pd.merge(SPS_Statistics_Combined, Production_Statistics_Combined, how='left',on='SourceLine')
        IncompleteProductionLine       = IncompleteProductionLine[(IncompleteProductionLine.CompletedShotsCount.isnull())]
        IncompleteProductionLine.reset_index(drop=True)

        ## Calculating SP Increment
        layout = [[sg.Text('Please Enter SPS Source Point Increment (Default = 1):',         size=(40, 1)), sg.InputText(1)],
                  [sg.Submit(), sg.Cancel()] ]

        window = sg.Window('Input SPS Source Point Increment:',auto_size_text=True, default_element_size=(10, 1)).Layout(layout)      
        event, values = window.Read()

        if event is None or event == 'Cancel':
            sg.PopupAutoClose('Exiting Without Processing Missing Shot Report', line_width=60)
            
        else:
            Src_Station_increment = values[0]
            Src_Station_increment = float(Src_Station_increment)

            ## Merging Production Statistcs and SPS Statistics
            Production_Statistics_Combined = pd.merge(Production_Statistics_Combined, SPS_Statistics_Combined, how='left',on='SourceLine')
            Production_Statistics_Combined.reset_index(drop=True)

            ## Generate Vib Total planned SPS
            QCList = []
            List_SLine    = (list(Production_Statistics_Combined.SourceLine))

            for i in range(len(List_SLine)):
                ListSL  = List_SLine[i]
                List_St = list(np.arange((Production_Statistics_Combined.ProductionStartSP[i]),(Production_Statistics_Combined.ProductionEndSP[i]+1), Src_Station_increment))
                QC_List = {'SPSLine': ListSL, 'SPSStation': List_St}
                QC_DF   = pd.DataFrame(data=QC_List,index=None)
                QC_DF1  = QC_DF.iloc[:,:]
                SPSLine    = (QC_DF1.loc[:,'SPSLine']).astype(int)
                SPSStation = (QC_DF1.loc[:,'SPSStation']).astype(float)
                LN_ST   = [SPSLine,SPSStation]
                QCcatdf = pd.concat (LN_ST,axis=1,ignore_index =True)
                QCList.append(QCcatdf)
            concatQCList = pd.concat(QCList,axis=0)
            concatQCList.rename(columns={0:'SourceLine', 1:'SourceStation'},inplace = True)
            QC_Planned_Rep = pd.DataFrame(concatQCList)
            QC_Planned_Rep['SourceLineStationCombined'] = (QC_Planned_Rep['SourceLine'].map(str) + QC_Planned_Rep['SourceStation'].map(str)).astype(float)
            QC_Planned_Rep['SourceLineStationCombined'] = QC_Planned_Rep['SourceLineStationCombined'].astype(float)
            QC_Planned_Rep                              = QC_Planned_Rep.reset_index(drop=True)

            ### Merging DF QC_Planned_Rep, MasterTB_DF_Merge
            QC_Missing_Rep = pd.merge(QC_Planned_Rep, MasterTB_DF_Merge,
                                how='left', on ='SourceLineStationCombined',
                                suffixes=('_Planned', '_Accomplished'))
            QC_Missing_Rep = QC_Missing_Rep[(QC_Missing_Rep.SourceLine_Accomplished.isnull())|
                                            (QC_Missing_Rep.SourceStation_Accomplished.isnull())]
            QC_Missing_Rep = QC_Missing_Rep.loc[:,['SourceLine_Planned','SourceStation_Planned','SourceLineStationCombined']]

            ### Merging DF QC_Missing_Rep, MasterSPS_DF_Merge_Final Missing 
            QC_Missing_Rep = pd.merge(QC_Missing_Rep, MasterSPS_DF_Merge,
                                how='left', on='SourceLineStationCombined',
                                suffixes=('_Missing', '_Preplot'))
            QC_Missing_Rep = QC_Missing_Rep[(QC_Missing_Rep.SourceLine.notnull())&
                                            (QC_Missing_Rep.SourceStation.notnull())]
            QC_Missing_Rep = pd.DataFrame(QC_Missing_Rep)
            QC_Missing_Rep.rename(columns={'SourceLine_Planned':'Line_Missing_Shots', 'SourceStation_Planned':'Station_Missing_Shots','SourceLineStationCombined':'SourceLineStationCombined',
                                           'SourceLine':'Line_Preplot_SP1', 'SourceStation':'Station_Preplot_SP1'},inplace = True)
            
            Missing_Station_Count  = (QC_Missing_Rep.groupby('Line_Missing_Shots').Station_Missing_Shots.count()).astype(int)
            Missing_Station_Count  = pd.DataFrame(Missing_Station_Count)
            Missing_Station_Count.reset_index(inplace=True)
            Missing_Station_Count.rename(columns={0:'Missing_Shot_Count'},inplace = True)
            #Missing_Shots_Row  = QC_Missing_Rep.groupby('Line_Missing_Shots').Station_Missing_Shots.unique()
            Missing_Shots_Row = QC_Missing_Rep.groupby('Line_Missing_Shots').agg({'Station_Missing_Shots':lambda x: list(x)})
            Missing_Shots_Row  = pd.DataFrame(Missing_Shots_Row)
            
            Missing_Station_Count = pd.DataFrame(Missing_Station_Count)
            Missing_Station_Count.rename(columns={'Line_Missing_Shots':'SourceLine','Missing_Shot_Count':'Missing_Shot_Countx'},inplace = True)
        
            Production_Statistics_Combined  = pd.merge(Production_Statistics_Combined, Missing_Station_Count, how='left',on='SourceLine')
            Production_Statistics_Combined  = Production_Statistics_Combined.loc[:,
                                   ['SourceLine', 'ProductionStartSP', 'ProductionEndSP', 'CompletedShotsCount', 'Station_Missing_Shots',
                                    'SPSBeginPoint', 'SPSEndPoint', 'SPSTotalShotsCount']]
            Production_Statistics_Combined.rename(columns={'Station_Missing_Shots':'MissingShots BetweenProductionStart&End'},inplace = True)
            Production_Statistics_Combined.reset_index(drop=True)

            Missing_Station_Count = pd.merge(Missing_Station_Count, Production_Statistics_Combined, how='left',on='SourceLine')
            Missing_Station_Count = Missing_Station_Count.loc[:,
                                   ['SourceLine','ProductionStartSP','ProductionEndSP','CompletedShotsCount','MissingShots BetweenProductionStart&End']]

            def get_QC_Missing_Rep_datetime():
                return " - MissingShotReport -" + datetime.now().strftime("%Y%m%d-%H%M%S") + ".xlsx"
            filename = tkinter.filedialog.asksaveasfilename(initialdir = "/",title = "Select file name to Export Missing Shot Record" ,
                                       filetypes = (("Excel file",".xlsx"),("Excel file",".xlsx")))
            if len(filename) >0:
                FileQC_Missing_Rep   = get_QC_Missing_Rep_datetime()
                outfile_QC_Missing_Rep = filename + FileQC_Missing_Rep 
                XLSX_writer = pd.ExcelWriter(outfile_QC_Missing_Rep)
                Production_Statistics_Combined.to_excel(XLSX_writer,'ProductionSummary',index=False)
                Missing_Station_Count.to_excel(XLSX_writer,'MissingShotsSummary',index=False)
                Missing_Shots_Row.to_excel(XLSX_writer,'MissingShotsQuickView',index=True)
                IncompleteProductionLine.to_excel(XLSX_writer,'IncompleteProductionLine',index=False)
                XLSX_writer.save()
                XLSX_writer.close()
                tkinter.messagebox.showinfo("Missing Shot Report Export","Missing Shot Report Saved as Excel")
            else:            
                tkinter.messagebox.showinfo("Missing Shot Report Export Message","Please Select File Name To Export")
        window.Close()            
    else:
        tkinter.messagebox.showinfo("Error In Generating Missing Shot Report","Please Check Imported Sourcelink TB Log or Imported SPS File to Generate Missing Shot Report Correctly")







