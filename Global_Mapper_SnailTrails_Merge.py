#Front End
import os
from tkinter import*
import tkinter.messagebox
import Global_Mapper_SnailTrails_Merge_BackEnd
import tkinter.ttk as ttk
import tkinter as tk
import sqlite3
from tkinter.filedialog import asksaveasfile
from tkinter.filedialog import askopenfilenames
from tkinter import simpledialog
import pandas as pd
import openpyxl
import csv
import time
import datetime
import numpy as np

def Vib_SnailTrail_Merge():
    Default_Date_today   = datetime.date.today()
    window = Tk()
    window.title ("Global Mapper SnailTrails Merge Import Wizard")
    window.geometry("1350x645+10+0")
    window.config(bg="cadet blue")
    window.resizable(0, 0)
    window.grid()
    OffsetTimeUTC = StringVar(window, value=float(+07.00))
    DataFrameTOP = LabelFrame(window, bd = 2, width = 400, height = 8, padx= 0, pady= 1,relief = RIDGE,labelanchor='nw',
                                               bg = "cadet blue",font=('aerial', 12, 'bold'))
    DataFrameTOP.pack(side=TOP)
    TableMargin = Frame(window, bd = 2, padx= 4, pady= 4, relief = RIDGE)
    TableMargin.pack(side=TOP)
    scrollbarx = Scrollbar(TableMargin, orient=HORIZONTAL)
    scrollbary = Scrollbar(TableMargin, orient=VERTICAL)
    tree = ttk.Treeview(TableMargin, column=("column1", "column2", "column3", "column4", "column5",
                                             "column6", "column7", "column8", "column9", "column10",
                                             "column11", "column12", "column13"), height=26, show='headings')
    scrollbary.config(command=tree.yview)
    scrollbary.pack(side=RIGHT, fill=Y)
    scrollbarx.config(command=tree.xview)
    scrollbarx.pack(side=BOTTOM, fill=X)
    tree.heading("#1", text="Unit ID", anchor=W) 
    tree.heading("#2", text="Latitude ", anchor=W)
    tree.heading("#3", text="Longitude", anchor=W)
    tree.heading("#4", text="Elevation", anchor=W)        
    tree.heading("#5", text="Status", anchor=W)
    tree.heading("#6", text="GPS Quality #", anchor=W)
    tree.heading("#7", text=" GPS Quality Verbal", anchor=W)        
    tree.heading("#8", text="Satellites" ,anchor=W)
    tree.heading("#9", text="UTC Date", anchor=W)
    tree.heading("#10", text=" UTC Time", anchor=W)        
    tree.heading("#11", text="Local Date", anchor=W)
    tree.heading("#12", text=" Local Time", anchor=W)
    tree.heading("#13", text=" Merge Flags", anchor=W)
    tree.column('#1', stretch=NO, minwidth=0, width=60)            
    tree.column('#2', stretch=NO, minwidth=0, width=100)
    tree.column('#3', stretch=NO, minwidth=0, width=100)
    tree.column('#4', stretch=NO, minwidth=0, width=100)
    tree.column('#5', stretch=NO, minwidth=0, width=100)
    tree.column('#6', stretch=NO, minwidth=0, width=120)
    tree.column('#7', stretch=NO, minwidth=0, width=150)
    tree.column('#8', stretch=NO, minwidth=0, width=80)
    tree.column('#9', stretch=NO, minwidth=0, width=100)
    tree.column('#10', stretch=NO, minwidth=0, width=100)
    tree.column('#11', stretch=NO, minwidth=0, width=100)
    tree.column('#12', stretch=NO, minwidth=0, width=100)
    tree.column('#13', stretch=NO, minwidth=0, width=100)
    style = ttk.Style()
    style.theme_use("clam")
    style.configure(".", font=('aerial', 10), foreground="black")
    style.configure("Treeview", foreground='black')
    style.configure("Treeview.Heading",font=('aerial', 10,'bold'), background='Ghost White', foreground='blue',fieldbackground='Ghost White')
    tree.pack()
    # All Functions defining 

    def iExit():
        iExit= tkinter.messagebox.askyesno("Vib Snail Trail Merge Module", "Confirm if you want to exit")
        if iExit >0:
            window.destroy()
            return

    def MergeVibOBLogSnailTrails():
        tree.delete(*tree.get_children())
        txtTotalMergedEntries.delete(0,END)
        txttotalVibSnailTrails.delete(0,END)
        txttotalOBSnailTrails.delete(0,END)
        conn = sqlite3.connect("Global_Mapper_SnailTrails_Log.db")
        DATA_VALID_OBLog_SNAILTRAIL = pd.read_sql_query("SELECT * FROM Global_Mapper_SnailTrails_OBLog;", conn)
        DATA_VALID_OBLog_SNAILTRAIL = pd.DataFrame(DATA_VALID_OBLog_SNAILTRAIL)
        DATA_VALID_OBLog_SNAILTRAIL = DATA_VALID_OBLog_SNAILTRAIL.reset_index(drop=True)
        DATA_VALID_OBLog_SNAILTRAIL["MergeFlags"]     = DATA_VALID_OBLog_SNAILTRAIL.shape[0]*["OBLog SnailTrail"]
        Length_OBLog_SNAILTRAIL_DF     = len(DATA_VALID_OBLog_SNAILTRAIL)
        DATA_VALID_VIB_SNAILTRAIL = pd.read_sql_query("SELECT * FROM Global_Mapper_SnailTrails_Vib;", conn)
        DATA_VALID_VIB_SNAILTRAIL = pd.DataFrame(DATA_VALID_VIB_SNAILTRAIL)
        DATA_VALID_VIB_SNAILTRAIL = DATA_VALID_VIB_SNAILTRAIL.reset_index(drop=True)
        DATA_VALID_VIB_SNAILTRAIL["MergeFlags"] = DATA_VALID_VIB_SNAILTRAIL.shape[0]*["Vib SnailTrail"]
        Length_VIB_SNAILTRAIL_DF = len(DATA_VALID_VIB_SNAILTRAIL)
        conn.commit()
        conn.close()
        MERGE_OBLOG_VIB_SNAILTRAIL = DATA_VALID_VIB_SNAILTRAIL.append(DATA_VALID_OBLog_SNAILTRAIL, ignore_index=True)
        MERGE_OBLOG_VIB_SNAILTRAIL  = pd.DataFrame(MERGE_OBLOG_VIB_SNAILTRAIL)
        MERGE_OBLOG_VIB_SNAILTRAIL['DuplicatedEntries_Chk_1']   = MERGE_OBLOG_VIB_SNAILTRAIL.sort_values(by =['UnitID', 'Lat','Lon']).duplicated(['UnitID','Lat','Lon'],keep='last')
        MERGE_OBLOG_VIB_SNAILTRAIL                              = MERGE_OBLOG_VIB_SNAILTRAIL.reset_index(drop=True)
        MERGE_OBLOG_VIB_SNAILTRAIL                              = MERGE_OBLOG_VIB_SNAILTRAIL.loc[MERGE_OBLOG_VIB_SNAILTRAIL.DuplicatedEntries_Chk_1 == False, 'UnitID': 'LocalTime']
        MERGE_OBLOG_VIB_SNAILTRAIL                              = MERGE_OBLOG_VIB_SNAILTRAIL.reset_index(drop=True)
        MERGE_OBLOG_VIB_SNAILTRAIL                              = pd.DataFrame(MERGE_OBLOG_VIB_SNAILTRAIL)    
        MERGE_OBLOG_VIB_SNAILTRAIL.sort_values(by =['Lat', 'Lon'], inplace =True)    
        MERGE_OBLOG_VIB_SNAILTRAIL                              = MERGE_OBLOG_VIB_SNAILTRAIL.reset_index(drop=True)
        MERGE_OBLOG_VIB_SNAILTRAIL                              = pd.DataFrame(MERGE_OBLOG_VIB_SNAILTRAIL)  

        for each_rec in range(len(MERGE_OBLOG_VIB_SNAILTRAIL)):
            tree.insert("", tk.END, values=list(MERGE_OBLOG_VIB_SNAILTRAIL.loc[each_rec]))
        TotalMergedEntries = len(MERGE_OBLOG_VIB_SNAILTRAIL)       
        txtTotalMergedEntries.insert(tk.END,TotalMergedEntries)
        txttotalVibSnailTrails.insert(tk.END,Length_VIB_SNAILTRAIL_DF)
        txttotalOBSnailTrails.insert(tk.END,Length_OBLog_SNAILTRAIL_DF)
        con = sqlite3.connect("Global_Mapper_SnailTrails_Log.db")
        cur=con.cursor()                
        MERGE_OBLOG_VIB_SNAILTRAIL.to_sql('Global_Mapper_Merged_SnailTrails',con, if_exists="replace", index=False)            
        con.commit()
        cur.close()
        con.close()
        

    def View_GM_SnailTrails_OBLog():
        tree.delete(*tree.get_children())
        txtTotalMergedEntries.delete(0,END)
        txttotalVibSnailTrails.delete(0,END)
        txttotalOBSnailTrails.delete(0,END)
        con = sqlite3.connect("Global_Mapper_SnailTrails_Log.db")
        cur=con.cursor() 
        DATA_VALID_OBLog_SNAILTRAIL = pd.read_sql_query("SELECT * FROM Global_Mapper_SnailTrails_OBLog;", con)
        DATA_VALID_OBLog_SNAILTRAIL = pd.DataFrame(DATA_VALID_OBLog_SNAILTRAIL)
        DATA_VALID_OBLog_SNAILTRAIL = DATA_VALID_OBLog_SNAILTRAIL.reset_index(drop=True)
        TotalEntries = len(DATA_VALID_OBLog_SNAILTRAIL)
        for each_rec in range(len(DATA_VALID_OBLog_SNAILTRAIL)):
            tree.insert("", tk.END, values=list(DATA_VALID_OBLog_SNAILTRAIL.loc[each_rec]))
        txttotalOBSnailTrails.insert(tk.END,TotalEntries)
        con.commit()
        cur.close()
        con.close()


    def View_GM_SnailTrails_Vib():
        tree.delete(*tree.get_children())
        txtTotalMergedEntries.delete(0,END)
        txttotalVibSnailTrails.delete(0,END)
        txttotalOBSnailTrails.delete(0,END)
        con = sqlite3.connect("Global_Mapper_SnailTrails_Log.db")
        cur=con.cursor() 
        DATA_VALID_VIBLog_SNAILTRAIL = pd.read_sql_query("SELECT * FROM Global_Mapper_SnailTrails_Vib;", con)
        DATA_VALID_VIBLog_SNAILTRAIL = pd.DataFrame(DATA_VALID_VIBLog_SNAILTRAIL)
        DATA_VALID_VIBLog_SNAILTRAIL = DATA_VALID_VIBLog_SNAILTRAIL.reset_index(drop=True)
        TotalEntries = len(DATA_VALID_VIBLog_SNAILTRAIL)
        for each_rec in range(len(DATA_VALID_VIBLog_SNAILTRAIL)):
            tree.insert("", tk.END, values=list(DATA_VALID_VIBLog_SNAILTRAIL.loc[each_rec]))
        txttotalVibSnailTrails.insert(tk.END,TotalEntries)
        con.commit()
        cur.close()
        con.close()

    def ViewMergeVibOBLogSnailTrails():
        tree.delete(*tree.get_children())
        txtTotalMergedEntries.delete(0,END)
        txttotalVibSnailTrails.delete(0,END)
        txttotalOBSnailTrails.delete(0,END)
        con = sqlite3.connect("Global_Mapper_SnailTrails_Log.db")
        cur=con.cursor() 
        MERGE_OBLOG_VIB_SNAILTRAIL = pd.read_sql_query("SELECT * FROM Global_Mapper_Merged_SnailTrails;", con)
        MERGE_OBLOG_VIB_SNAILTRAIL = pd.DataFrame(MERGE_OBLOG_VIB_SNAILTRAIL)
        MERGE_OBLOG_VIB_SNAILTRAIL = MERGE_OBLOG_VIB_SNAILTRAIL.reset_index(drop=True)
        TotalEntries = len(MERGE_OBLOG_VIB_SNAILTRAIL)
        for each_rec in range(len(MERGE_OBLOG_VIB_SNAILTRAIL)):
            tree.insert("", tk.END, values=list(MERGE_OBLOG_VIB_SNAILTRAIL.loc[each_rec]))
        txtTotalMergedEntries.insert(tk.END,TotalEntries)
        con.commit()
        cur.close()
        con.close()

    def ExportForGlobalMapper():
        conn = sqlite3.connect("Global_Mapper_SnailTrails_Log.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Global_Mapper_Merged_SnailTrails;", conn)
        Export_GlobalMapper = pd.DataFrame(Complete_df)
        Export_GlobalMapper = Export_GlobalMapper.loc[:,['Lat','Lon','Elevation','Status',
                                                         'GPSQuality_Number','GPSQuality_Verbal','Satellites',
                                                         'UTCDate','UTCTime','LocalDate','LocalTime']]
        Export_GlobalMapper.rename(columns={'Lat':'Lat', 'Lon':'Lon', 'Elevation':'Elevation',
                                         'Status':'Status', 'GPSQuality_Number':'GPS Quality#',
                                         'GPSQuality_Verbal':' GPS Quality Verbal','Satellites':'Satellites',
                                         'UTCDate':'UTC Date','UTCTime':' UTC Time',
                                         'LocalDate':'Local Date','LocalTime':' Local Time'},inplace = True)        
        Export_GlobalMapper = Export_GlobalMapper.reset_index(drop=True)
        filename = tkinter.filedialog.asksaveasfilename(initialdir = "/" ,title = "Save Merged SnailTrail Export For Global Mapper As CSV" ,\
                   defaultextension='.csv', filetypes = (("CSV file",".csv"),("Excel file",".xlsx")))
        if len(filename) >0:
            if filename.endswith('.csv'):
                Export_GlobalMapper.to_csv(filename,index=None)
                tkinter.messagebox.showinfo("Valid Merged SnailTrail Export For Global Mapper","Merged SnailTrail Export For Global Mapper Saved as CSV")
            else:
                Export_GlobalMapper.to_excel(filename, sheet_name='Merged Vib SnailTrail', index=False)
                tkinter.messagebox.showinfo("Valid Merged SnailTrail Export For Global Mapper","Merged SnailTrail Export For Global Mapper Saved as Excel")
        else:
            tkinter.messagebox.showinfo("Export Merged SnailTrail Message","Please Select File Name To Export")
                    
        conn.commit()
        conn.close()


    def ClearView():
        txttotalVibSnailTrails.delete(0,END)
        txttotalOBSnailTrails.delete(0,END)
        txtTotalMergedEntries.delete(0,END)
        tree.delete(*tree.get_children())    

    def ImportVibSnailTrail():
        tree.delete(*tree.get_children())
        txttotalVibSnailTrails.delete(0,END)
        txtTotalMergedEntries.delete(0,END)
        fileList = askopenfilenames(initialdir = "/", title = "Import Vib Snail Trails Files" , filetypes=[('CSV File', '*.csv'), ('Excel File', ('*.xls', '*.xlsx'))])
        Length_fileList  =  len(fileList)
        if Length_fileList >0:            
            if fileList:
                dfList =[]            
                for filename in fileList:
                    if filename.endswith('.csv'):
                        df = pd.read_csv(filename, sep=',' , low_memory=False)
                        df = df.iloc[:,:]
                        UnitID                  = df.loc[:,'UnitID']
                        Lat                     = df.loc[:,'Lat']
                        Lon                     = df.loc[:,'Lon']
                        Elevation               = df.loc[:,'Elevation']
                        Status                  = df.loc[:,'Status']
                        GPSQuality_Number       = df.loc[:,'GPS Quality#']
                        GPSQuality_Verbal       = df.loc[:,' GPS Quality Verbal']
                        Satellites              = df.loc[:,'Satellites']
                        UTCDate                 = df.loc[:,'UTC Date']
                        UTCTime                 = df.loc[:,' UTC Time']
                        LocalDate               = df.loc[:,'Local Date']
                        LocalTime               = df.loc[:,' Local Time']
                        
                        column_names = [UnitID, Lat, Lon, Elevation, Status, GPSQuality_Number,  GPSQuality_Verbal, Satellites, UTCDate,  UTCTime, LocalDate,  LocalTime]
                        catdf = pd.concat (column_names,axis=1,ignore_index =True)
                        dfList.append(catdf) 
                    else:
                        df = pd.read_excel(filename)
                        df = df.iloc[:,:]
                        UnitID                  = df.loc[:,'UnitID']
                        Lat                     = df.loc[:,'Lat']
                        Lon                     = df.loc[:,'Lon']
                        Elevation               = df.loc[:,'Elevation']
                        Status                  = df.loc[:,'Status']
                        GPSQuality_Number       = df.loc[:,'GPS Quality#']
                        GPSQuality_Verbal       = df.loc[:,' GPS Quality Verbal']
                        Satellites              = df.loc[:,'Satellites']
                        UTCDate                 = df.loc[:,'UTC Date']
                        UTCTime                 = df.loc[:,' UTC Time']
                        LocalDate               = df.loc[:,'Local Date']
                        LocalTime               = df.loc[:,' Local Time']
                        
                        column_names = [UnitID, Lat, Lon, Elevation, Status, GPSQuality_Number,  GPSQuality_Verbal, Satellites, UTCDate,  UTCTime, LocalDate,  LocalTime]
                        catdf = pd.concat (column_names,axis=1,ignore_index =True)
                        dfList.append(catdf) 

                concatDf = pd.concat(dfList,axis=0, ignore_index =True)
                concatDf.rename(columns={0:'UnitID',                1:'Lat',            2:'Lon',        3:'Elevation',      4:'Status',         5:'GPSQuality_Number',
                                         6:'GPSQuality_Verbal',     7:'Satellites',     8:'UTCDate',    9:'UTCTime',        10:'LocalDate',     11:'LocalTime'},inplace = True)
                
                # Separating Valid with Shot ID Not Null
                Valid_VibSnailTrail_DF = pd.DataFrame(concatDf)
                Valid_VibSnailTrail_DF = Valid_VibSnailTrail_DF[pd.to_numeric(Valid_VibSnailTrail_DF.UnitID, errors='coerce').notnull()]
                Valid_VibSnailTrail_DF = Valid_VibSnailTrail_DF[pd.to_numeric(Valid_VibSnailTrail_DF.Lat, errors='coerce').notnull()]
                Valid_VibSnailTrail_DF = Valid_VibSnailTrail_DF[pd.to_numeric(Valid_VibSnailTrail_DF.Lon, errors='coerce').notnull()]
                Valid_VibSnailTrail_DF['DuplicatedEntries_Chk_1'] = Valid_VibSnailTrail_DF.sort_values(by =['UnitID', 'Lat','Lon']).duplicated(['UnitID','Lat','Lon'],keep='last')            
                Valid_VibSnailTrail_DF                      = Valid_VibSnailTrail_DF.loc[Valid_VibSnailTrail_DF.DuplicatedEntries_Chk_1 == False, 'UnitID': 'LocalTime']
                DATA_VALID_VIB_SNAILTRAIL                   = pd.DataFrame(Valid_VibSnailTrail_DF)
                DATA_VALID_VIB_SNAILTRAIL                   = DATA_VALID_VIB_SNAILTRAIL[(DATA_VALID_VIB_SNAILTRAIL.Lat > 0)&
                                                              (abs(DATA_VALID_VIB_SNAILTRAIL.Lon) > 0)]            
                DATA_VALID_VIB_SNAILTRAIL  = DATA_VALID_VIB_SNAILTRAIL[(DATA_VALID_VIB_SNAILTRAIL.GPSQuality_Number > 0)]
                DATA_VALID_VIB_SNAILTRAIL  = DATA_VALID_VIB_SNAILTRAIL.reset_index(drop=True)
                DATA_VALID_VIB_SNAILTRAIL  = pd.DataFrame(DATA_VALID_VIB_SNAILTRAIL)            
                Length_VIB_SNAILTRAIL_DF = len(DATA_VALID_VIB_SNAILTRAIL)

                ## Connect To Database and Export DF  
                con= sqlite3.connect("Global_Mapper_SnailTrails_Log.db")
                cur=con.cursor()                
                DATA_VALID_VIB_SNAILTRAIL.to_sql('Global_Mapper_SnailTrails_Vib',con, if_exists="replace", index=False)            
                con.commit()
                cur.close()
                con.close()

                ## Tree view populate
                txttotalVibSnailTrails.insert(tk.END,Length_VIB_SNAILTRAIL_DF)
                for each_rec in range(len(DATA_VALID_VIB_SNAILTRAIL)):
                    tree.insert("", tk.END, values=list(DATA_VALID_VIB_SNAILTRAIL.loc[each_rec]))
        else:
            tkinter.messagebox.showinfo("Import Vib SnailTrail File Message","Please Select Vib Snail Trail Files To Import")

    def ImportOBLogSnailTrail():
        tree.delete(*tree.get_children())
        txttotalOBSnailTrails.delete(0,END)
        txtTotalMergedEntries.delete(0,END)
        fileList = askopenfilenames(initialdir = "/", title = "Import Observer Snail Trails Files" , filetypes=[('CSV File', '*.csv'), ('Excel File', ('*.xls', '*.xlsx'))])
        Length_fileList  =  len(fileList)
        if Length_fileList >0:            
            if fileList:
                dfList =[]            
                for filename in fileList:
                    if filename.endswith('.csv'):
                        df = pd.read_csv(filename, sep=',' , low_memory=False)
                        df = df.iloc[:,:]
                        UnitID                  = df.loc[:,'UnitID']
                        Lat                     = df.loc[:,'Lat']
                        Lon                     = df.loc[:,'Lon']
                        Elevation               = df.loc[:,'Elevation']
                        Status                  = df.loc[:,'Status']
                        GPSQuality_Number       = df.loc[:,'GPS Quality#']
                        GPSQuality_Verbal       = df.loc[:,' GPS Quality Verbal']
                        Satellites              = df.loc[:,'Satellites']
                        UTCDate                 = df.loc[:,'UTC Date']
                        UTCTime                 = df.loc[:,' UTC Time']
                        LocalDate               = df.loc[:,'Local Date']
                        LocalTime               = df.loc[:,' Local Time']
                        
                        column_names = [UnitID, Lat, Lon, Elevation, Status, GPSQuality_Number,  GPSQuality_Verbal, Satellites, UTCDate,  UTCTime, LocalDate,  LocalTime]
                        catdf = pd.concat (column_names,axis=1,ignore_index =True)
                        dfList.append(catdf) 
                    else:
                        df = pd.read_excel(filename)
                        df = df.iloc[:,:]
                        UnitID                  = df.loc[:,'UnitID']
                        Lat                     = df.loc[:,'Lat']
                        Lon                     = df.loc[:,'Lon']
                        Elevation               = df.loc[:,'Elevation']
                        Status                  = df.loc[:,'Status']
                        GPSQuality_Number       = df.loc[:,'GPS Quality#']
                        GPSQuality_Verbal       = df.loc[:,' GPS Quality Verbal']
                        Satellites              = df.loc[:,'Satellites']
                        UTCDate                 = df.loc[:,'UTC Date']
                        UTCTime                 = df.loc[:,' UTC Time']
                        LocalDate               = df.loc[:,'Local Date']
                        LocalTime               = df.loc[:,' Local Time']
                        
                        column_names = [UnitID, Lat, Lon, Elevation, Status, GPSQuality_Number,  GPSQuality_Verbal, Satellites, UTCDate,  UTCTime, LocalDate,  LocalTime]
                        catdf = pd.concat (column_names,axis=1,ignore_index =True)
                        dfList.append(catdf) 

                concatDf = pd.concat(dfList,axis=0, ignore_index =True)
                concatDf.rename(columns={0:'UnitID',                1:'Lat',            2:'Lon',        3:'Elevation',      4:'Status',         5:'GPSQuality_Number',
                                         6:'GPSQuality_Verbal',     7:'Satellites',     8:'UTCDate',    9:'UTCTime',        10:'LocalDate',     11:'LocalTime'},inplace = True)
                
                # Separating Valid with Shot ID Not Null
                Valid_OBLogSnailTrail_DF = pd.DataFrame(concatDf)            
                Valid_OBLogSnailTrail_DF = Valid_OBLogSnailTrail_DF[pd.to_numeric(Valid_OBLogSnailTrail_DF.UnitID, errors='coerce').notnull()]
                Valid_OBLogSnailTrail_DF = Valid_OBLogSnailTrail_DF[pd.to_numeric(Valid_OBLogSnailTrail_DF.Lat, errors='coerce').notnull()]
                Valid_OBLogSnailTrail_DF = Valid_OBLogSnailTrail_DF[pd.to_numeric(Valid_OBLogSnailTrail_DF.Lon, errors='coerce').notnull()]
                Valid_OBLogSnailTrail_DF = Valid_OBLogSnailTrail_DF[pd.to_numeric(Valid_OBLogSnailTrail_DF.GPSQuality_Number, errors='coerce').notnull()]
                Valid_OBLogSnailTrail_DF['DuplicatedEntries_Chk_1'] = Valid_OBLogSnailTrail_DF.sort_values(by =['UnitID', 'Lat','Lon']).duplicated(['UnitID','Lat','Lon'],keep='last')
                Valid_OBLogSnailTrail_DF                      = Valid_OBLogSnailTrail_DF.reset_index(drop=True)
                Valid_OBLogSnailTrail_DF                      = Valid_OBLogSnailTrail_DF.loc[Valid_OBLogSnailTrail_DF.DuplicatedEntries_Chk_1 == False, 'UnitID': 'LocalTime']
                Valid_OBLogSnailTrail_DF                      = Valid_OBLogSnailTrail_DF.reset_index(drop=True)
                Valid_OBLogSnailTrail_DF                      = pd.DataFrame(Valid_OBLogSnailTrail_DF)    
                DATA_VALID_OBLog_SNAILTRAIL  = pd.DataFrame(Valid_OBLogSnailTrail_DF)
                DATA_VALID_OBLog_SNAILTRAIL      = DATA_VALID_OBLog_SNAILTRAIL[(DATA_VALID_OBLog_SNAILTRAIL.Lat > 0)&
                                                  (abs(DATA_VALID_OBLog_SNAILTRAIL.Lon) > 0)]
                DATA_VALID_OBLog_SNAILTRAIL  = DATA_VALID_OBLog_SNAILTRAIL[(DATA_VALID_OBLog_SNAILTRAIL.GPSQuality_Number > 0)]
                DATA_VALID_OBLog_SNAILTRAIL  = DATA_VALID_OBLog_SNAILTRAIL.reset_index(drop=True)
                Length_OBLog_SNAILTRAIL_DF     = len(DATA_VALID_OBLog_SNAILTRAIL)
                DATA_VALID_OBLog_SNAILTRAIL  = pd.DataFrame(DATA_VALID_OBLog_SNAILTRAIL) 

                ## Connect To Database and Export DF  
                con= sqlite3.connect("Global_Mapper_SnailTrails_Log.db")
                cur=con.cursor()                
                DATA_VALID_OBLog_SNAILTRAIL.to_sql('Global_Mapper_SnailTrails_OBLog',con, if_exists="replace", index=False)            
                con.commit()
                cur.close()
                con.close()

                ## Tree view populate
                txttotalOBSnailTrails.insert(tk.END,Length_OBLog_SNAILTRAIL_DF)
                for each_rec in range(len(DATA_VALID_OBLog_SNAILTRAIL)):
                    tree.insert("", tk.END, values=list(DATA_VALID_OBLog_SNAILTRAIL.loc[each_rec]))
        else:
            tkinter.messagebox.showinfo("Import Vib SnailTrail File Message","Please Select Vib Snail Trail Files To Import")

    ### top Actions Menu
    txttotalVibSnailTrails  = Entry(DataFrameTOP, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 12)
    txttotalVibSnailTrails.grid(row=0,column=0, sticky ="W", padx= 1)
    btntotalVibSnailTrails= Button(DataFrameTOP, text="View Vib SnailTrails Import", font=('aerial', 9, 'bold'), height =1, width=22, bg = "ghost white", bd=2, command = View_GM_SnailTrails_Vib)
    btntotalVibSnailTrails.grid(row=0,column=1, sticky ="W", padx= 2)

    txttotalOBSnailTrails  = Entry(DataFrameTOP, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 12)
    txttotalOBSnailTrails.grid(row=0,column=2, sticky ="W", padx= 450)
    btntotalOBSnailTrails = Button(DataFrameTOP, text="View OBLog SnailTrails Import", font=('aerial', 9, 'bold'), height =1, width=25, bg = "ghost white", bd=2, command = View_GM_SnailTrails_OBLog)
    btntotalOBSnailTrails.grid(row=0,column=2, sticky ="W", padx= 262)

    btnExportGlobalMapper = Button(DataFrameTOP, text="View Merged Report", font=('aerial', 9, 'bold'), height =1, width=18, bd=2, command = ViewMergeVibOBLogSnailTrails)
    btnExportGlobalMapper.grid(row=0,column=2, sticky ="W", padx= 700)

    btnExportGlobalMapper = Button(DataFrameTOP, text="Export For Global Mapper", font=('aerial', 9, 'bold'), height =1, width=22, bd=2, command = ExportForGlobalMapper)
    btnExportGlobalMapper.grid(row=0,column=2, sticky ="W", padx= 900)

    ### Bottom Actions Menu
    btnImportVibSnailTrails = Button(window, text="Import Vib SnailTrails Log", font=('aerial', 9, 'bold'), height =1, width=23, bd=4, command = ImportVibSnailTrail)
    btnImportVibSnailTrails.place(x=2,y=612)

    btnImportOBLogSnailTrails= Button(window, text="Import OBLog SnailTrails Log", font=('aerial', 9, 'bold'), height =1, width=25, bd=4, command = ImportOBLogSnailTrail)
    btnImportOBLogSnailTrails.place(x=181,y=612)

    btnMergeVib_OBLog_SnailTrails= Button(window, text="Merge Vib - OBLog SnailTrails Import", font=('aerial', 9, 'bold'), height =1, width=30, bd=4, command = MergeVibOBLogSnailTrails)
    btnMergeVib_OBLog_SnailTrails.place(x=374,y=612)

    txtTotalMergedEntries  = Entry(window, font=('aerial', 12, 'bold'),textvariable=IntVar(),  bd=4, width = 12)
    txtTotalMergedEntries.place(x=602,y=612)

    btnExit = Button(window, text="Exit Widget", font=('aerial', 9, 'bold'), height =1, width=10, bd=3, command = iExit)
    btnExit.place(x=1268,y=612)






