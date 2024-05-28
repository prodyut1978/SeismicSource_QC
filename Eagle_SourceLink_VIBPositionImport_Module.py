#Front End
import os
from tkinter import*
import tkinter.messagebox
import Eagle_SourceLink_Vibroseis_Log_BackEnd
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

def SourceLink_COG_LogIMPORT():
    Default_Date_today   = datetime.date.today()
    window = Tk()
    window.title ("Eagle SourceLink VIB Position Report Import Wizard")
    window.geometry("1250x650+10+0")
    window.config(bg="cadet blue")
    window.resizable(0, 0)
    TableMargin = Frame(window, bd = 2, padx= 10, pady= 8, relief = RIDGE)
    TableMargin.pack(side=TOP)
    TableMargin.pack(side=LEFT)
    scrollbarx = Scrollbar(TableMargin, orient=HORIZONTAL)
    scrollbary = Scrollbar(TableMargin, orient=VERTICAL)
    tree = ttk.Treeview(TableMargin, column=("column1", "column2", "column3", "column4", "column5",
                                             "column6", "column7", "column8", "column9", "column10",
                                             "column11","column12"), height=26, show='headings')
    scrollbary.config(command=tree.yview)
    scrollbary.pack(side=RIGHT, fill=Y)
    scrollbarx.config(command=tree.xview)
    scrollbarx.pack(side=BOTTOM, fill=X)        
    tree.heading("#1", text="File Number", anchor=W)
    tree.heading("#2", text="ShotID (Encoder Index)", anchor=W)
    tree.heading("#3", text="EPNumber", anchor=W)
    tree.heading("#4", text="Source Point Line", anchor=W)
    tree.heading("#5", text="Source Point Station", anchor=W)        
    tree.heading("#6", text="Distance to Source Point", anchor=W)
    tree.heading("#7", text="Near Flag Line", anchor=W)
    tree.heading("#8", text="Near Flag Station", anchor=W)        
    tree.heading("#9", text="Distance to Near Flag" ,anchor=W)
    tree.heading("#10", text="GPS Quality", anchor=W)
    tree.heading("#11", text="Unit_ID", anchor=W)
    tree.heading("#12", text="Near Flag Message", anchor=W)
    tree.column('#1', stretch=NO, minwidth=0, width=80)            
    tree.column('#2', stretch=NO, minwidth=0, width=140)
    tree.column('#3', stretch=NO, minwidth=0, width=70)
    tree.column('#4', stretch=NO, minwidth=0, width=115)
    tree.column('#5', stretch=NO, minwidth=0, width=130)
    tree.column('#6', stretch=NO, minwidth=0, width=150)
    tree.column('#7', stretch=NO, minwidth=0, width=100)
    tree.column('#8', stretch=NO, minwidth=0, width=110)
    tree.column('#9', stretch=NO, minwidth=0, width=140)
    tree.column('#10', stretch=NO, minwidth=0, width=80)
    tree.column('#11', stretch=NO, minwidth=0, width=60)
    tree.column('#12', stretch=NO, minwidth=0, width=120)
    style = ttk.Style()
    style.theme_use("clam")
    style.configure(".", font=('aerial', 8), foreground="black")
    style.configure("Treeview", foreground='black')
    style.configure("Treeview.Heading",font=('aerial', 8,'bold'), background='Ghost White', foreground='blue',fieldbackground='Ghost White')
    tree.pack()

    # All Functions defining 

    def iExit():
        iExit= tkinter.messagebox.askyesno("Eagle VIB Position Import Wizard", "Confirm if you want to exit")
        if iExit >0:
            window.destroy()
            return

    def ViewTotalImport():
        tree.delete(*tree.get_children())
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedShotID.delete(0,END)
        conn = sqlite3.connect("SourceLink_Log.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_VIB_COG_TEMP ORDER BY `ShotID` ASC ;", conn)
        data = pd.DataFrame(Complete_df)
        data = data.reset_index(drop=True)
        for each_rec in range(len(data)):
            tree.insert("", tk.END, values=list(data.loc[each_rec]))
        TotalEntries = len(data)       
        txtTotalEntries.insert(tk.END,TotalEntries)              
        conn.commit()
        conn.close()
        InvalidEntries()
        DuplicatedShotIDEntries()

    def ViewInvalidImport():
        tree.delete(*tree.get_children())
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedShotID.delete(0,END)
        conn = sqlite3.connect("SourceLink_Log.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_VIB_COG_INVALID ORDER BY `ShotID` ASC ;", conn)
        data = pd.DataFrame(Complete_df)
        data = data.reset_index(drop=True)
        for each_rec in range(len(data)):
            tree.insert("", tk.END, values=list(data.loc[each_rec]))
        TotalInvalidEntries = len(data)       
        txtInvalidEntries.insert(tk.END,TotalInvalidEntries)              
        conn.commit()
        conn.close()
        TotalEntries()
        DuplicatedShotIDEntries()

    def ViewDuplicatedShotIDImport():
        tree.delete(*tree.get_children())
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedShotID.delete(0,END)
        conn = sqlite3.connect("SourceLink_Log.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_VIB_COG_DUPLICATEDSHOTID ORDER BY `ShotID` ASC ;", conn)
        data = pd.DataFrame(Complete_df)
        data = data.reset_index(drop=True)
        for each_rec in range(len(data)):
            tree.insert("", tk.END, values=list(data.loc[each_rec]))
        TotalDuplicatedEntries = len(data)       
        txtDuplicatedShotID.insert(tk.END,TotalDuplicatedEntries)              
        conn.commit()
        conn.close()
        InvalidEntries()
        TotalEntries()

    def ClearView():
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedShotID.delete(0,END)
        tree.delete(*tree.get_children())

    def TotalEntries():
        conn = sqlite3.connect("SourceLink_Log.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_VIB_COG_TEMP ;", conn)
        data = pd.DataFrame(Complete_df)
        TotalEntries = len(data)       
        txtTotalEntries.insert(tk.END,TotalEntries)              
        conn.commit()
        conn.close()

    def InvalidEntries():
        conn = sqlite3.connect("SourceLink_Log.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_VIB_COG_INVALID ;", conn)
        data = pd.DataFrame(Complete_df)
        TotalInvalidEntries = len(data)       
        txtInvalidEntries.insert(tk.END,TotalInvalidEntries)              
        conn.commit()
        conn.close()

    def DuplicatedShotIDEntries():
        conn = sqlite3.connect("SourceLink_Log.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_VIB_COG_DUPLICATEDSHOTID ;", conn)
        data = pd.DataFrame(Complete_df)
        TotalDuplicatedEntries = len(data)       
        txtDuplicatedShotID.insert(tk.END,TotalDuplicatedEntries)              
        conn.commit()
        conn.close()
      
    def DeleteSelectedImportData():
        iDelete = tkinter.messagebox.askyesno("Delete Entry", "Confirm if you want to Delete")
        if iDelete >0:
            txtTotalEntries.delete(0,END)
            conn = sqlite3.connect("SourceLink_Log.db")
            cur = conn.cursor()                
            for selected_item in tree.selection():
                cur.execute("DELETE FROM Eagle_VIB_COG_TEMP WHERE ShotID =? AND SourceLine =? AND \
                            SourceStation =? ",\
                            (tree.set(selected_item, '#2'), tree.set(selected_item, '#4'),tree.set(selected_item, '#5'),)) 
                conn.commit()
                tree.delete(selected_item)
            conn.commit()
            conn.close()
            TotalEntries()
            return

    def ExportValidCOG():
        conn = sqlite3.connect("SourceLink_Log.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_VIB_COG_TEMP ORDER BY `ShotID` ASC ;", conn)
        Export_COGDF = pd.DataFrame(Complete_df)
        Export_COGDF = Export_COGDF.reset_index(drop=True)
        filename = tkinter.filedialog.asksaveasfilename(initialdir = "/" ,title = "Export Vib Position file" ,\
                   defaultextension='.csv', filetypes = (("CSV file",".csv"),("Excel file",".xlsx")))
        if len(filename) >0:
            if filename.endswith('.csv'):
                Export_COGDF.to_csv(filename,index=None)
                tkinter.messagebox.showinfo("Valid COG Export","Valid COG Saved as CSV")
            else:
                Export_COGDF.to_excel(filename, sheet_name='COG Export', index=False)
                tkinter.messagebox.showinfo("Valid COG Export","Valid COG Saved as Excel")
        else:
            tkinter.messagebox.showinfo("Export COG Message","Please Select File Name To Export")
                    
        conn.commit()
        conn.close()

    def ImportVIBPositionFile():
        tree.delete(*tree.get_children())
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedShotID.delete(0,END)
        fileList = askopenfilenames(initialdir = "/", title = "Import SourceLink Vib Position Files" , filetypes=[('CSV File', '*.csv'), ('Excel File', ('*.xls', '*.xlsx'))])
        Length_fileList  =  len(fileList)
        if Length_fileList >0:
            if fileList:
                dfList_COG = []           
                for filename in fileList:
                    if filename.endswith('.csv'):
                        df_COG          = pd.read_csv(filename, sep=',' , low_memory=False)
                        df_COG          = df_COG.iloc[:,:]
                        FileNum         = df_COG .loc[:,'FF ID']
                        ShotID          = df_COG .loc[:,'Encoder Index']
                        EPNumber        = df_COG .loc[:,'EP']
                        SourceLine      = df_COG .loc[:,'Source Point Line']
                        SourceStation   = df_COG .loc[:,'Source Point Station']
                        DistanceCOG     = df_COG .loc[:,'Distance to Source Point']
                        NearFlagLine    = df_COG .loc[:,'Near Flag Line']
                        NearFlagStation = df_COG .loc[:,'Near Flag Station']                
                        DistanceNearFlag= df_COG .loc[:,'Distance to Near Flag']
                        GPS_Quality     = df_COG .loc[:,'GPS Quality']
                        Unit_ID         = df_COG .loc[:,'Unit ID']
                        column_names = [FileNum, ShotID, EPNumber, SourceLine, SourceStation, DistanceCOG, NearFlagLine, NearFlagStation, DistanceNearFlag, GPS_Quality, Unit_ID]
                        catdf = pd.concat (column_names,axis=1,ignore_index =True)
                        dfList_COG.append(catdf) 
                    else:
                        df_COG  = pd.read_excel(filename)
                        df_COG  = df_COG.iloc[:,:]
                        FileNum                 = df_COG .loc[:,'FF ID']
                        ShotID                  = df_COG .loc[:,'Encoder Index']
                        EPNumber                = df_COG .loc[:,'EP']
                        SourceLine              = df_COG .loc[:,'Source Point Line']
                        SourceStation           = df_COG .loc[:,'Source Point Station']
                        DistanceCOG             = df_COG .loc[:,'Distance to Source Point']
                        NearFlagLine            = df_COG .loc[:,'Near Flag Line']
                        NearFlagStation         = df_COG .loc[:,'Near Flag Station']                
                        DistanceNearFlag        = df_COG .loc[:,'Distance to Near Flag']
                        GPS_Quality             = df_COG .loc[:,'GPS Quality']
                        Unit_ID                 = df_COG .loc[:,'Unit ID']
                        column_names = [FileNum, ShotID, EPNumber, SourceLine, SourceStation, DistanceCOG, NearFlagLine, NearFlagStation, DistanceNearFlag, GPS_Quality, Unit_ID]
                        catdf = pd.concat (column_names,axis=1,ignore_index =True)
                        dfList_COG.append(catdf) 

                concatDf = pd.concat(dfList_COG,axis=0, ignore_index =True)
                concatDf.rename(columns={0:'FileNum', 1:'ShotID', 2:'EPNumber', 3:'SourceLine', 4:'SourceStation',
                                         5:'DistanceCOG', 6:'NearFlagLine', 7:'NearFlagStation', 8:'DistanceNearFlag',
                                         9:'GPS_Quality', 10:'Unit_ID' },inplace = True)
                concatDf = concatDf.reset_index(drop=True)

                ## Separating Invalid COG
                Invalid_COG_DF   = pd.DataFrame(concatDf)
                Invalid_COG_DF   = Invalid_COG_DF[pd.to_numeric(Invalid_COG_DF.ShotID,errors='coerce').isnull()]               
                Invalid_COG_DF   = Invalid_COG_DF.reset_index(drop=True)
                Data_Invalid_COG = pd.DataFrame(Invalid_COG_DF)

                ## Separating Valid COG 
                Valid_COG_DF = pd.DataFrame(concatDf)        
                Valid_COG_DF = Valid_COG_DF[pd.to_numeric(Valid_COG_DF.ShotID,errors='coerce').notnull()]        
                Valid_COG_DF["SourceLine"].fillna(0, inplace = True)
                Valid_COG_DF["SourceStation"].fillna(0, inplace = True)
                Valid_COG_DF["NearFlagLine"].fillna(0, inplace = True)
                Valid_COG_DF["NearFlagStation"].fillna(0, inplace = True)
                Valid_COG_DF["EPNumber"].fillna(1, inplace = True)
                Valid_COG_DF["Unit_ID"].fillna(0, inplace = True)
                Valid_COG_DF['SourceLine']       = (Valid_COG_DF.loc[:,['SourceLine']]).astype(int)
                Valid_COG_DF['SourceStation']    = (Valid_COG_DF.loc[:,['SourceStation']]).astype(float)
                Valid_COG_DF['FileNum']          = (Valid_COG_DF.loc[:,['FileNum']]).astype(int)
                Valid_COG_DF['NearFlagLine']     = (Valid_COG_DF.loc[:,['NearFlagLine']]).astype(int)
                Valid_COG_DF['NearFlagStation']  = (Valid_COG_DF.loc[:,['NearFlagStation']]).astype(int)
                Valid_COG_DF['EPNumber']         = (Valid_COG_DF.loc[:,['EPNumber']]).astype(int)
                Valid_COG_DF['Unit_ID']          = (Valid_COG_DF.loc[:,['Unit_ID']]).astype(int)
                Valid_COG_DF['ShotID']           = (Valid_COG_DF.loc[:,['ShotID']]).astype(int)
                Valid_COG_DF['DuplicatedEntries']=  Valid_COG_DF.sort_values(by =['ShotID','FileNum','SourceLine','SourceStation']).duplicated(['ShotID'],keep='last')
                

                # Separeating Valid with Duplicated Shot ID        
                DATA_DuplicatedShotID              = Valid_COG_DF.loc[Valid_COG_DF.DuplicatedEntries == True, 'FileNum': 'Unit_ID']
                DATA_DuplicatedShotID              = DATA_DuplicatedShotID.reset_index(drop=True)   
                DATA_DuplicatedShotID              = pd.DataFrame(DATA_DuplicatedShotID)

                # Separeating Valid with No Duplicated Shot ID        
                DATA_VALID_COG                     = Valid_COG_DF.loc[Valid_COG_DF.DuplicatedEntries == False, 'FileNum': 'Unit_ID']
                DATA_VALID_COG                     = DATA_VALID_COG.reset_index(drop=True)   
                DATA_VALID_COG                     = pd.DataFrame(DATA_VALID_COG)
                
                ## Finding Near Flag Mismatch on DATA_VALID_COG
                Line_COG_M             = DATA_VALID_COG['SourceLine'].astype(int)
                Station_COG_M          = DATA_VALID_COG['SourceStation'].astype(int)
                Line_Station_COG       = (Line_COG_M.map(str) + Station_COG_M.map(str)).astype(int)
                Combined_LS_COG        = pd.DataFrame(Line_Station_COG)
                Combined_LS_COG.rename(columns     = {0:'COG_Line_Station_Combined'},inplace = True)
                Line_COG_N             = DATA_VALID_COG['NearFlagLine'].astype(int)
                Station_COG_N          = DATA_VALID_COG['NearFlagStation'].astype(int)
                Line_Station_COG_N     = (Line_COG_N.map(str) + Station_COG_N.map(str)).astype(int)
                Combined_LS_COG_N      = pd.DataFrame(Line_Station_COG_N)
                Combined_LS_COG_N.rename(columns = {0:'COG_Near_Flag_LS_Combined'},inplace = True)
                VIB_Rep_COG_Merge      = pd.concat([DATA_VALID_COG , Combined_LS_COG, Combined_LS_COG_N], axis=1)
                VIB_Rep_COG_Merge      = pd.DataFrame(VIB_Rep_COG_Merge)
                VIB_Rep_COG_Merge['Check_Near_Flag']= VIB_Rep_COG_Merge['COG_Line_Station_Combined'] == VIB_Rep_COG_Merge['COG_Near_Flag_LS_Combined']

                def trans_COG_MISSING(x):
                    if x   == False:
                        return 'NEAR_FLAG_MISMATCH'
                    elif x == True:
                        return np.nan    
                    else:
                        return x

                VIB_Rep_COG_Merge['Near_Flag_Message'] = VIB_Rep_COG_Merge['Check_Near_Flag'].apply(trans_COG_MISSING)
                VIB_Rep_COG_Merge = VIB_Rep_COG_Merge.loc[:,
                            ['FileNum', 'ShotID', 'EPNumber', 'SourceLine', 'SourceStation',
                             'DistanceCOG', 'NearFlagLine', 'NearFlagStation',
                             'DistanceNearFlag','GPS_Quality','Unit_ID','Near_Flag_Message']]
                VIB_Rep_COG_Merge = VIB_Rep_COG_Merge.reset_index(drop=True)
                con= sqlite3.connect("SourceLink_Log.db")
                cur=con.cursor()                
                VIB_Rep_COG_Merge.to_sql('Eagle_VIB_COG_TEMP',con, if_exists="replace", index=False)
                Data_Invalid_COG.to_sql ('Eagle_VIB_COG_INVALID',con, if_exists="replace", index=False)
                DATA_DuplicatedShotID.to_sql ('Eagle_VIB_COG_DUPLICATEDSHOTID',con, if_exists="replace", index=False)
                con.commit()
                cur.close()
                con.close()
                ViewTotalImport()
        else:
            tkinter.messagebox.showinfo("Import VIB Position Message","Please Select VIB Position Files To Import")
                

    ##### Entry Wizard
    txtInvalidEntries  = Entry(window, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
    txtInvalidEntries.place(x=425,y=6)
    txtTotalEntries  = Entry(window, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 10)
    txtTotalEntries.place(x=1150,y=6)
    txtDuplicatedShotID  = Entry(window, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
    txtDuplicatedShotID.place(x=475,y=620)

    L1 = Label(window, text = "Source Link VIB Position Details:", font=("arial", 10,'bold'),bg = "green").place(x=2,y=6)
    L2 = Label(window, text = "No ShotID Or File Number", font=("arial", 10,'bold'),bg = "red").place(x=505,y=7)
    L3 = Label(window, text = "Duplicated ShotID", font=("arial", 10,'bold'),bg = "red").place(x=555,y=620)

    ### Button Wizard
    btnInValidImport= Button(window, text="View Invalid Import", font=('aerial', 9, 'bold'), height =1, width=16, bd=1, command = ViewInvalidImport)
    btnInValidImport.place(x=300,y=6)
    btnValidImport = Button(window, text="Valid For Analysis", font=('aerial', 9, 'bold'), height =1, width=15, bd=1, command = ViewTotalImport)
    btnValidImport.place(x=1032,y=6)
    btnExportValid = Button(window, text="Export Valid Position", font=('aerial', 9, 'bold'), height =1, width=17, bd=1, command = ExportValidCOG)
    btnExportValid.place(x=900,y=6)
    btnImportVIBPosition= Button(window, text="Import Vib Position", font=('aerial', 9, 'bold'), height =1, width=16, bd=4, command = ImportVIBPositionFile)
    btnImportVIBPosition.place(x=2,y=620)
    btnDuplicatedShotIDImport= Button(window, text="View Duplicated ShotID", font=('aerial', 9, 'bold'), height =1, width=19, bd=2, command = ViewDuplicatedShotIDImport)
    btnDuplicatedShotIDImport.place(x=332,y=620)
    btnDelete = Button(window, text="Delete Selected Valid", font=('aerial', 9, 'bold'), height =1, width=18, bd=4, command = DeleteSelectedImportData)
    btnDelete.place(x=920,y=620)
    btnClearView = Button(window, text="Clear View", font=('aerial', 9, 'bold'), height =1, width=10, bd=4, command = ClearView)
    btnClearView.place(x=1077,y=620)
    btnExit = Button(window, text="Exit Import", font=('aerial', 9, 'bold'), height =1, width=10, bd=4, command = iExit)
    btnExit.place(x=1165,y=620)













