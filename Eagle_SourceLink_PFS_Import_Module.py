#Front End
import os
from tkinter import*
import tkinter.messagebox
import Eagle_SourceLink_Dynamite_Log_BackEnd
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

def SourceLink_PFS_LogIMPORT():
    Default_Date_today   = datetime.date.today()
    window = Tk()
    window.title ("Eagle SourceLink PFS Log Import Wizard")
    window.geometry("1350x730+10+0")
    window.config(bg="cadet blue")
    window.resizable(0, 0)
    OffsetTimeUTC = StringVar(window, value=float(+07.00))
    TableMargin = Frame(window, bd = 2, padx= 10, pady= 8, relief = RIDGE)
    TableMargin.pack(side=TOP)
    TableMargin.pack(side=LEFT)
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
                                             "column41"), height=27, show='headings')
    scrollbary.config(command=tree.yview)
    scrollbary.pack(side=RIGHT, fill=Y)
    scrollbarx.config(command=tree.xview)
    scrollbarx.pack(side=BOTTOM, fill=X)
    tree.heading("#1", text="DB_ID", anchor=W) 
    tree.heading("#2", text="ShotID", anchor=W)
    tree.heading("#3", text="FileNum", anchor=W)
    tree.heading("#4", text="EPNumber", anchor=W)        
    tree.heading("#5", text="Line", anchor=W)
    tree.heading("#6", text="Station", anchor=W)
    tree.heading("#7", text="LocalDate", anchor=W)        
    tree.heading("#8", text="LocalTime" ,anchor=W)
    tree.heading("#9", text="PFSComment", anchor=W)
    tree.heading("#10", text="ShotStatus", anchor=W)        
    tree.heading("#11", text="Timebreak", anchor=W)
    tree.heading("#12", text="FirstBreak" ,anchor=W)
    tree.heading("#13", text="Battery", anchor=W)
    tree.heading("#14", text="CapRes", anchor=W)        
    tree.heading("#15", text="GeoRes", anchor=W)
    tree.heading("#16", text="FlagNum", anchor=W)        
    tree.heading("#17", text="UpholeWindow", anchor=W)
    tree.heading("#18", text="FiredOK", anchor=W)
    tree.heading("#19", text="BatteryOK", anchor=W)
    tree.heading("#20", text="GeoOK", anchor=W)
    tree.heading("#21", text="CapOK", anchor=W)
    tree.heading("#22", text="GPSQuality", anchor=W)
    tree.heading("#23", text="Unit_ID", anchor=W)
    tree.heading("#24", text="TBDate", anchor=W)
    tree.heading("#25", text="TBTime", anchor=W)        
    tree.heading("#26", text="TBMicro", anchor=W)    
    tree.heading("#27", text="CapSerialNumber", anchor=W)
    tree.heading("#28", text="Latitude" ,anchor=W)
    tree.heading("#29", text="Longitude", anchor=W)        
    tree.heading("#30", text="Altitude", anchor=W)
    tree.heading("#31", text="EncoderIndex", anchor=W)
    tree.heading("#32", text="RecordIndex" ,anchor=W)        
    tree.heading("#33", text="EPCount", anchor=W)
    tree.heading("#34", text="CrewID", anchor=W)        
    tree.heading("#35", text="GPSTime", anchor=W)
    tree.heading("#36", text="GPSAltitude", anchor=W)
    tree.heading("#37", text="Sats", anchor=W)
    tree.heading("#38", text="PDOP", anchor=W)
    tree.heading("#39", text="HDOP" ,anchor=W)        
    tree.heading("#40", text="VDOP", anchor=W)
    tree.heading("#41", text="Age", anchor=W)   
    tree.column('#1', stretch=NO, minwidth=0, width=0)            
    tree.column('#2', stretch=NO, minwidth=0, width=80)
    tree.column('#3', stretch=NO, minwidth=0, width=80)
    tree.column('#4', stretch=NO, minwidth=0, width=80)
    tree.column('#5', stretch=NO, minwidth=0, width=80)
    tree.column('#6', stretch=NO, minwidth=0, width=80)
    tree.column('#7', stretch=NO, minwidth=0, width=80)
    tree.column('#8', stretch=NO, minwidth=0, width=80)
    tree.column('#9', stretch=NO, minwidth=0, width=120)
    tree.column('#10', stretch=NO, minwidth=0, width=80)
    tree.column('#11', stretch=NO, minwidth=0, width=80)
    tree.column('#12', stretch=NO, minwidth=0, width=70)
    tree.column('#13', stretch=NO, minwidth=0, width=70)
    tree.column('#14', stretch=NO, minwidth=0, width=70)
    tree.column('#15', stretch=NO, minwidth=0, width=70)
    tree.column('#16', stretch=NO, minwidth=0, width=60)
    tree.column('#17', stretch=NO, minwidth=0, width=60)
    tree.column('#18', stretch=NO, minwidth=0, width=90)            
    tree.column('#19', stretch=NO, minwidth=0, width=60)
    tree.column('#20', stretch=NO, minwidth=0, width=90)
    tree.column('#21', stretch=NO, minwidth=0, width=60)
    tree.column('#22', stretch=NO, minwidth=0, width=60)
    tree.column('#23', stretch=NO, minwidth=0, width=60)
    tree.column('#24', stretch=NO, minwidth=0, width=60)
    tree.column('#25', stretch=NO, minwidth=0, width=60)
    tree.column('#26', stretch=NO, minwidth=0, width=90)
    tree.column('#27', stretch=NO, minwidth=0, width=90)
    tree.column('#28', stretch=NO, minwidth=0, width=90)
    tree.column('#29', stretch=NO, minwidth=0, width=90)
    tree.column('#30', stretch=NO, minwidth=0, width=90)
    tree.column('#31', stretch=NO, minwidth=0, width=90)
    tree.column('#32', stretch=NO, minwidth=0, width=90)
    tree.column('#33', stretch=NO, minwidth=0, width=90)
    tree.column('#34', stretch=NO, minwidth=0, width=60)
    tree.column('#35', stretch=NO, minwidth=0, width=90)
    tree.column('#36', stretch=NO, minwidth=0, width=40)
    tree.column('#37', stretch=NO, minwidth=0, width=40)
    tree.column('#38', stretch=NO, minwidth=0, width=40)
    tree.column('#39', stretch=NO, minwidth=0, width=40)
    tree.column('#40', stretch=NO, minwidth=0, width=40)
    tree.column('#41', stretch=NO, minwidth=0, width=100)

    style = ttk.Style()
    style.theme_use("clam")
    style.configure(".", font=('aerial', 8), foreground="black")
    style.configure("Treeview", foreground='black')
    style.configure("Treeview.Heading",font=('aerial', 8,'bold'), background='Ghost White', foreground='blue',fieldbackground='Ghost White')
    tree.pack()

    ### All Functions defining 
    def iExit():
        iExit= tkinter.messagebox.askyesno("Eagle PSS Import Wizard", "Confirm if you want to exit")
        if iExit >0:
            window.destroy()
            return

    def ViewTotalImport():
        tree.delete(*tree.get_children())
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedShotID.delete(0,END)
        conn = sqlite3.connect("SourceLink_Dynamite_Log.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PFSLog_TEMP ORDER BY `ShotID` ASC ;", conn)
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
        conn = sqlite3.connect("SourceLink_Dynamite_Log.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PFSLog_INVALID_NULL ORDER BY `ShotID` ASC ;", conn)
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
        conn = sqlite3.connect("SourceLink_Dynamite_Log.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PFSLog_DuplicatedShotID ORDER BY `ShotID` ASC ;", conn)
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
        conn = sqlite3.connect("SourceLink_Dynamite_Log.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PFSLog_TEMP ;", conn)
        data = pd.DataFrame(Complete_df)
        TotalEntries = len(data)       
        txtTotalEntries.insert(tk.END,TotalEntries)              
        conn.commit()
        conn.close()

    def InvalidEntries():
        conn = sqlite3.connect("SourceLink_Dynamite_Log.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PFSLog_INVALID_NULL ;", conn)
        data = pd.DataFrame(Complete_df)
        TotalInvalidEntries = len(data)       
        txtInvalidEntries.insert(tk.END,TotalInvalidEntries)              
        conn.commit()
        conn.close()

    def DuplicatedShotIDEntries():
        conn = sqlite3.connect("SourceLink_Dynamite_Log.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PFSLog_DuplicatedShotID ;", conn)
        data = pd.DataFrame(Complete_df)
        TotalDuplicatedEntries = len(data)       
        txtDuplicatedShotID.insert(tk.END,TotalDuplicatedEntries)              
        conn.commit()
        conn.close()


    def DeleteSelectedImportData():
        iDelete = tkinter.messagebox.askyesno("Delete Entry", "Confirm if you want to Delete")
        if iDelete >0:
            txtTotalEntries.delete(0,END)
            conn = sqlite3.connect("SourceLink_Dynamite_Log.db")
            cur = conn.cursor()                
            for selected_item in tree.selection():
                cur.execute("DELETE FROM Eagle_PFSLog_TEMP WHERE DataBase_ID =? " ,(tree.set(selected_item, '#1'),)) 
                conn.commit()
                tree.delete(selected_item)
            conn.commit()
            conn.close()
            TotalEntries()
            return

    def ExportValidForTest_i_Fy():
        conn = sqlite3.connect("SourceLink_Dynamite_Log.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PFSLog_TEMP ORDER BY `ShotID` ASC ;", conn)
        Export_TestifyDF = pd.DataFrame(Complete_df)
        Export_TestifyDF = Export_TestifyDF.loc[:,['Encoder_Index','ShotStatus','ShotID','FileNum','EPNumber','SourceLine','SourceStation',
                                                               'Local_Date','Local_Time','Observer_Comment','TB_Date','TB_Time','TB_Micro',
                                                               'Record_Index','EP_Count','Crew_ID', 'Timebreak', 'FirstBreak', 'Battery',
                                                               'CapRes', 'GeoRes', 'FlagNum', 'UpholeWindow', 'FiredOK','BatteryOK', 'GeoOK', 'CapOK','GPS_Time',
                                                               'Latitude','Longitude','Altitude','GPS_Altitude','Sats','PDOP','HDOP',
                                                               'VDOP','Age','GPS_Quality','Unit_ID','CapSerialNumber']]

        Export_TestifyDF.rename(columns={'Encoder_Index':'Encoder Index', 'ShotStatus':'Void', 'ShotID':'Shot ID', 'FileNum':'File Num', 'EPNumber':'EP ID', 'SourceLine':'Line',
                             'SourceStation':'Station','Local_Date':'Date','Local_Time':'Time','Observer_Comment':'Comment','TB_Date':'TB Date','TB_Time':'TB Time',
                             'TB_Micro':'TB Micro','Record_Index':'Record Index','EP_Count':'EP Count','Crew_ID':'Crew ID','Timebreak':'Timebreak', 'FirstBreak':'First Break', 'Battery':'Battery',
                             'CapRes':'Cap Res', 'GeoRes':'Geo Res', 'FlagNum':'Flag Num', 'UpholeWindow':'Uphole Window', 'FiredOK':'Fired OK','BatteryOK':'Battery OK', 'GeoOK':'Geo OK', 'CapOK':'Cap OK',
                             'GPS_Time':'GPS Time','Latitude':'Lat','Longitude':'Lon','Altitude':'Altitude','GPS_Altitude':'GPS Altitude','Sats':'Sats','PDOP':'PDOP','HDOP':'HDOP','VDOP':'VDOP','Age':'Age',
                             'GPS_Quality':'Quality', 'Unit_ID':'Unit ID','CapSerialNumber':'Cap Serial Number'},inplace = True)

        Export_TestifyDF['Date']    = pd.to_datetime(Export_TestifyDF['Date']).dt.strftime('%m/%d/%Y')
        Export_TestifyDF['TB Date'] = pd.to_datetime(Export_TestifyDF['TB Date']).dt.strftime('%m/%d/%Y')    
        Export_TestifyDF = Export_TestifyDF.reset_index(drop=True)    
        Export_TestifyDF_With_VOID      = pd.DataFrame(Export_TestifyDF)
        Export_TestifyDF_With_No_VOID   = pd.DataFrame(Export_TestifyDF)
        Export_TestifyDF_With_No_VOID   = Export_TestifyDF_With_No_VOID[(Export_TestifyDF_With_No_VOID.Void.isnull())]
        ExportVoidorNotQuestion = tkinter.messagebox.askquestion("Export Testif-y Message",
                                "Do you Like To Export Testif-y Input File Including All Void Shots?"+ '\n' +
                                'Yes For Export With Including All Void Shots' + '\n' +
                                'No For Export With Excluding (Removing) All Void Shots')
        if ExportVoidorNotQuestion == "yes":
            filename = tkinter.filedialog.asksaveasfilename(initialdir = "/" ,title = "Save Valid PFS Export For Testif-y As CSV Or Excel" ,\
                   defaultextension='.csv', filetypes = (("CSV file",".csv"),("Excel file",".xlsx")))
            if len(filename) >0:
                if filename.endswith('.csv'):
                    Export_TestifyDF_With_VOID.to_csv((filename),index=None)
                    tkinter.messagebox.showinfo("Valid PFS Export For Testif-y","PFS For Test-i-fy Saved as CSV")
                else:
                    Export_TestifyDF_With_VOID.to_excel(filename, sheet_name='PFS Testif-y', index=False)
                    tkinter.messagebox.showinfo("Valid PFS Export For Testif-y","PFS For Testif-y Saved as Excel")
            else:
                tkinter.messagebox.showinfo("Export Testif-y Message","Please Select File Name To Export")
        else:
            filename = tkinter.filedialog.asksaveasfilename(initialdir = "/" ,title = "Save Valid PFS Export For Testif-y As CSV Or Excel" ,\
                   defaultextension='.csv', filetypes = (("CSV file",".csv"),("Excel file",".xlsx")))

            if len(filename) >0:
                if filename.endswith('.csv'):
                    Export_TestifyDF_With_No_VOID.to_csv((filename),index=None)
                    tkinter.messagebox.showinfo("Valid PFS Export For Testif-y","PFS For Test-i-fy Saved as CSV")
                else:
                    Export_TestifyDF_With_No_VOID.to_excel(filename, sheet_name='PFS Testif-y', index=False)
                    tkinter.messagebox.showinfo("Valid PFS Export For Testif-y","PFS For Testif-y Saved as Excel")
            else:
                tkinter.messagebox.showinfo("Export Testif-y Message","Please Select PFS File Name To Export")
                    
        conn.commit()
        conn.close()

    def ExportListBoxPFS():
        dfList =[] 
        for child in tree.get_children():
            df = tree.item(child)["values"]
            dfList.append(df)
        ListBox_DF = pd.DataFrame(dfList)
        ListBox_DF.rename(columns={0:'DB_ID', 1:'Shot ID', 2:'File Num', 3:'EP ID', 4:'Line', 5:'Station', 6:'Date',
                                 7:'Time', 8:'Comment', 9:'Void', 10:'Timebreak', 11:'First Break', 12:'Battery',
                                 13:'Cap Res', 14:'Geo Res', 15:'Flag Num', 16:'Uphole Window', 17:'Fired OK', 18:'Battery OK', 19:'Geo OK',
                                 20:'Cap OK', 21:'Quality',
                                 22:'Unit ID', 23:'TB Date', 24:'TB Time', 25:'TB Micro', 26:'Cap Serial Number',
                                 27:'Lat', 28:'Lon', 29:'Altitude', 30:'Encoder Index', 31:'Record Index',
                                 32:'EP Count', 33:'Crew ID', 34:'GPS Time',
                                 35:'GPS Altitude', 36:'Sats', 37:'PDOP', 38:'HDOP', 39:'VDOP', 40:'Age'},inplace = True)
        Export_ListBox  = pd.DataFrame(ListBox_DF)
        Length_Export_ListBox  =  len(Export_ListBox)
        if Length_Export_ListBox >0:                
            Export_ListBox.drop(['DB_ID'], axis=1, inplace=True)
            Export_ListBox['Void'].replace('None', np.nan, inplace=True)
            Export_ListBox  = Export_ListBox.reset_index(drop=True)    
            filename = tkinter.filedialog.asksaveasfilename(initialdir = "/" ,title = "Select PFS File Name to Export" ,\
                           defaultextension='.csv', filetypes = (("CSV file",".csv"),("Excel file",".xlsx")))
            if len(filename) >0:
                if filename.endswith('.csv'):
                    Export_ListBox.to_csv(filename,index=None)
                    tkinter.messagebox.showinfo("ListBox PFS Export","ListBox PFS Entries Saved as CSV")
                else:
                    Export_ListBox.to_excel(filename, sheet_name='ListBoxTB', index=False)
                    tkinter.messagebox.showinfo("ListBox PFS Export","ListBox PFS Entries Saved as Excel")
            else:
                tkinter.messagebox.showinfo("ListBox PFS Export Message","Please Select File Name To Export")

    def ImportPFSLogFile():
        tree.delete(*tree.get_children())
        txtTotalRAWImportedPFS.delete(0,END)
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedShotID.delete(0,END)
        UTC_Offset_Hours = (Entrytxt_TimeOffset.get())
        UTC_Offset_Hours = float(UTC_Offset_Hours)
        fileList = askopenfilenames(initialdir = "/", title = "Import SourceLink PFS Files" , filetypes=[('CSV File', '*.csv'), ('Excel File', ('*.xls', '*.xlsx'))])
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
                        Timebreak             = df.loc[:,'Timebreak']
                        FirstBreak            = df.loc[:,'First Break']
                        Battery               = df.loc[:,'Battery']
                        CapRes                = df.loc[:,'Cap Res']
                        GeoRes                = df.loc[:,'Geo Res']
                        FlagNum               = df.loc[:,'Flag Num']
                        UpholeWindow          = df.loc[:,'Uphole Window']
                        FiredOK               = df.loc[:,'Fired OK']
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
                        CapSerialNumber       = df.loc[:,'Cap Serial Number']
                        Latitude              = df.loc[:,'Lat']
                        Longitude             = df.loc[:,'Lon']
                        Altitude              = df.loc[:,'Altitude']
                        Encoder_Index         = df.loc[:,'Encoder Index']
                        Record_Index          = df.loc[:,'Record Index']
                        EP_Count              = df.loc[:,'EP Count']
                        Crew_ID               = df.loc[:,'Crew ID']
                        BatteryOK             = df.loc[:,'Battery OK']
                        GeoOK                 = df.loc[:,'Geo OK']
                        GPS_Time              = df.loc[:,'GPS Time']
                        GPS_Altitude          = df.loc[:,'GPS Altitude']
                        Sats                  = df.loc[:,'Sats']
                        PDOP                  = df.loc[:,'PDOP']
                        HDOP                  = df.loc[:,'HDOP']
                        VDOP                  = df.loc[:,'VDOP']
                        Age                   = df.loc[:,'Age']
                        CapOK                 = df.loc[:,'Cap OK']
                        
                        column_names = [ShotID, FileNum, EPNumber, SourceLine, SourceStation, Local_Date, Local_Time, Observer_Comment, ShotStatus,
                            Timebreak, FirstBreak, Battery, CapRes, GeoRes, FlagNum, UpholeWindow,
                            FiredOK,  BatteryOK, GeoOK, CapOK, GPS_Quality, Unit_ID, TB_Date, TB_Time, TB_Micro, CapSerialNumber,
                            Latitude, Longitude, Altitude, Encoder_Index, Record_Index, EP_Count, Crew_ID,
                            GPS_Time, GPS_Altitude, Sats, PDOP, HDOP, VDOP, Age ]
                        
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
                        Timebreak             = df.loc[:,'Timebreak']
                        FirstBreak            = df.loc[:,'First Break']
                        Battery               = df.loc[:,'Battery']
                        CapRes                = df.loc[:,'Cap Res']
                        GeoRes                = df.loc[:,'Geo Res']
                        FlagNum               = df.loc[:,'Flag Num']
                        UpholeWindow          = df.loc[:,'Uphole Window']
                        FiredOK               = df.loc[:,'Fired OK']
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
                        CapSerialNumber       = df.loc[:,'Cap Serial Number']
                        Latitude              = df.loc[:,'Lat']
                        Longitude             = df.loc[:,'Lon']
                        Altitude              = df.loc[:,'Altitude']
                        Encoder_Index         = df.loc[:,'Encoder Index']
                        Record_Index          = df.loc[:,'Record Index']
                        EP_Count              = df.loc[:,'EP Count']
                        Crew_ID               = df.loc[:,'Crew ID']
                        BatteryOK             = df.loc[:,'Battery OK']
                        GeoOK                 = df.loc[:,'Geo OK']
                        GPS_Time              = df.loc[:,'GPS Time']
                        GPS_Altitude          = df.loc[:,'GPS Altitude']
                        Sats                  = df.loc[:,'Sats']
                        PDOP                  = df.loc[:,'PDOP']
                        HDOP                  = df.loc[:,'HDOP']
                        VDOP                  = df.loc[:,'VDOP']
                        Age                   = df.loc[:,'Age']
                        CapOK                 = df.loc[:,'Cap OK']
                        
                        column_names = [ShotID, FileNum, EPNumber, SourceLine, SourceStation, Local_Date, Local_Time, Observer_Comment, ShotStatus,
                            Timebreak, FirstBreak, Battery, CapRes, GeoRes, FlagNum, UpholeWindow,
                            FiredOK,  BatteryOK, GeoOK, CapOK, GPS_Quality, Unit_ID, TB_Date, TB_Time, TB_Micro, CapSerialNumber,
                            Latitude, Longitude, Altitude, Encoder_Index, Record_Index, EP_Count, Crew_ID,
                            GPS_Time, GPS_Altitude, Sats, PDOP, HDOP, VDOP, Age ]
                        
                        catdf = pd.concat (column_names,axis=1,ignore_index =True)
                        dfList.append(catdf) 

                concatDf = pd.concat(dfList,axis=0, ignore_index =True)
                concatDf.rename(columns={0:'ShotID', 1:'FileNum', 2:'EPNumber', 3:'SourceLine', 4:'SourceStation', 5:'Local_Date',
                                 6:'Local_Time', 7:'Observer_Comment', 8:'ShotStatus', 9:'Timebreak', 10:'FirstBreak', 11:'Battery',
                                 12:'CapRes', 13:'GeoRes', 14:'FlagNum', 15:'UpholeWindow', 16:'FiredOK', 17:'BatteryOK', 18:'GeoOK',
                                 19:'CapOK', 20:'GPS_Quality',
                                 21:'Unit_ID', 22:'TB_Date', 23:'TB_Time', 24:'TB_Micro', 25:'CapSerialNumber',
                                 26:'Latitude', 27:'Longitude', 28:'Altitude', 29:'Encoder_Index', 30:'Record_Index',
                                 31:'EP_Count', 32:'Crew_ID', 33:'GPS_Time',
                                 34:'GPS_Altitude', 35:'Sats', 36:'PDOP', 37:'HDOP', 38:'VDOP', 39:'Age'},inplace = True)

                # RAW DUMP Total PSS imported
                RAW_DUMP_ImportedPFS_DF = pd.DataFrame(concatDf)
                len_RAW_Dump            = len(RAW_DUMP_ImportedPFS_DF)

                # Separating InValid with Shot ID is Null
                Invalid_PFS_DF    = pd.DataFrame(concatDf)
                Invalid_PFS_DF    = Invalid_PFS_DF[pd.to_numeric(Invalid_PFS_DF.ShotID,errors='coerce').isnull()]                    
                Invalid_PFS_DF    = Invalid_PFS_DF.reset_index(drop=True)
                Data_Invalid_PFS  = pd.DataFrame(Invalid_PFS_DF)
                
                # Separating Valid with Shot ID Not Null
                Valid_PFS_DF = pd.DataFrame(concatDf)
                Valid_PFS_DF = Valid_PFS_DF[pd.to_numeric(Valid_PFS_DF.ShotID, errors='coerce').notnull()]                  
                Valid_PFS_DF["SourceLine"].fillna(0, inplace = True)
                Valid_PFS_DF["SourceStation"].fillna(0, inplace = True)
                Valid_PFS_DF["FileNum"].fillna(0, inplace = True)
                Valid_PFS_DF["EPNumber"].fillna(1, inplace = True)
                Valid_PFS_DF["Local_Date"].fillna('1900/1/01', inplace = True)
                Valid_PFS_DF["TB_Date"].fillna('1900/1/01', inplace = True)                
                Valid_PFS_DF['SourceLine']             = (Valid_PFS_DF.loc[:,['SourceLine']]).astype(int)
                Valid_PFS_DF['SourceStation']          = (Valid_PFS_DF.loc[:,['SourceStation']]).astype(float)
                Valid_PFS_DF['ShotID']                 = (Valid_PFS_DF.loc[:,['ShotID']]).astype(int)
                Valid_PFS_DF['FileNum']                = (Valid_PFS_DF.loc[:,['FileNum']]).astype(int)
                Valid_PFS_DF['EPNumber']               = (Valid_PFS_DF.loc[:,['EPNumber']]).astype(int)
                Valid_PFS_DF['Local_Date']             = pd.to_datetime(Valid_PFS_DF['Local_Date']).dt.strftime('%Y/%m/%d')
                try:                    
                    Valid_PFS_DF['TB_Date']            = pd.to_datetime(Valid_PFS_DF['TB_Date']).dt.strftime('%Y/%m/%d')
                    Valid_PFS_DF['TB_Time']            = pd.to_datetime(Valid_PFS_DF['TB_Time']).dt.strftime('%H:%M:%S')
                    Valid_PFS_DF['TB_Micro']           = pd.to_datetime(Valid_PFS_DF['TB_Micro']).dt.strftime('%f')
                except:
                    try:
                        Valid_PFS_DF['TB_Date']            = pd.to_datetime(Valid_PFS_DF['Local_Date']).dt.strftime('%Y/%m/%d')
                        Valid_PFS_DF['TB_Time']            = pd.to_datetime(Valid_PFS_DF['Local_Time']).dt.strftime('%H:%M:%S')
                        Valid_PFS_DF['TB_DateTime']        = pd.to_datetime(Valid_PFS_DF.TB_Date.astype(str)+' '+Valid_PFS_DF.TB_Time.astype(str))                                                                        
                        Valid_PFS_DF['TB_DateTime']        = pd.to_datetime(Valid_PFS_DF['TB_DateTime'].astype(str)) + pd.DateOffset(hours=UTC_Offset_Hours)
                        Valid_PFS_DF['TB_Date']            = pd.to_datetime(Valid_PFS_DF['TB_DateTime']).dt.strftime('%Y/%m/%d')
                        Valid_PFS_DF['TB_Time']            = pd.to_datetime(Valid_PFS_DF['TB_DateTime']).dt.strftime('%H:%M:%S')                                                
                        Valid_PFS_DF['TB_Micro']           = 0
                        Valid_PFS_DF.drop(['TB_DateTime'], axis=1, inplace=True)
                        tkinter.messagebox.showinfo("Imported PFS File Message","PFS Column Name : [TB UTC Time] Is Corrupted" + '\n' +  '\n' +
                                                    "Columns [TB_Date], [TB_Time] Is Fixed From Local Time <<<>>> [TB_Micro] Column Is Corrupted" + '\n' +  '\n' +
                                                    " To Fix All TB Columns Go To Advanced Option and Select >> Fix Corrupted PFS From TB Import >> ")
                    except:
                        Valid_PFS_DF['TB_Date']            = Valid_PFS_DF['TB_Date']
                        Valid_PFS_DF['TB_Time']            = Valid_PFS_DF['TB_Time']
                        Valid_PFS_DF['TB_Micro']           = Valid_PFS_DF['TB_Micro']
                        tkinter.messagebox.showinfo("Imported PFS File Message","PFS Column Name : [TB UTC Time] Is Corrupted" + '\n' +  '\n' +
                                                    "[TB_Date], [TB_Time] and [TB_Micro] Columns are Corrupted" + '\n' +  '\n' +
                                                    " To Fix All TB Columns Go To Advanced Option and Select >> Fix Corrupted PFS From TB Import >> ")        

                Valid_PFS_DF['DuplicatedEntries']      = Valid_PFS_DF.sort_values(by =['ShotID', 'FileNum', 'Unit_ID', 'Crew_ID']).duplicated(['ShotID','FileNum','SourceLine','SourceStation'],keep='last')
                Valid_PFS_DF                           = Valid_PFS_DF.reset_index(drop=True)
                Valid_PFS_DF                           = pd.DataFrame(Valid_PFS_DF)

                # Separating Valid with Shot ID Not Duplicated
                DATA_VALID_PFS = Valid_PFS_DF.loc[Valid_PFS_DF.DuplicatedEntries == False, 'ShotID': 'Age']
                DATA_VALID_PFS = DATA_VALID_PFS.reset_index(drop=True)
                DATA_VALID_PFS = pd.DataFrame(DATA_VALID_PFS)

                # Separating Valid with Shot ID Duplicated
                DATA_DuplicatedShotID = Valid_PFS_DF.loc[Valid_PFS_DF.DuplicatedEntries == True, 'ShotID': 'Age']
                DATA_DuplicatedShotID = DATA_DuplicatedShotID.reset_index(drop=True)
                DATA_DuplicatedShotID = pd.DataFrame(DATA_DuplicatedShotID)

                # Connect To Database and Export DF
                txtTotalRAWImportedPFS.insert(tk.END,len_RAW_Dump)
                con= sqlite3.connect("SourceLink_Dynamite_Log.db")
                cur=con.cursor()                
                DATA_VALID_PFS.to_sql('Eagle_PFSLog_TEMP',con, if_exists="replace",  index_label='DataBase_ID')
                Data_Invalid_PFS.to_sql('Eagle_PFSLog_INVALID_NULL',con, if_exists="replace", index_label='DataBase_ID')
                DATA_DuplicatedShotID.to_sql('Eagle_PFSLog_DuplicatedShotID',con, if_exists="replace", index_label='DataBase_ID')
                con.commit()
                cur.close()
                con.close()
                ViewTotalImport()
        else:
            tkinter.messagebox.showinfo("Import PFS File Message","Please Select PFS Files To Import")

    ### Top 
    L1 = Label(window, text = "Total Raw PFS Imported:", font=("arial", 10,'bold'),bg = "green").place(x=2,y=6)
    txtTotalRAWImportedPFS  = Entry(window, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
    txtTotalRAWImportedPFS.place(x=2,y=35)

    L2 = Label(window, text = "Null ShotID Or FFID", font=("arial", 10,'bold'),bg = "red").place(x=555,y=7)
    txtInvalidEntries  = Entry(window, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
    txtInvalidEntries.place(x=475,y=6)
    btnInValidImport= Button(window, text="View Invalid Null ShotID", font=('aerial', 9, 'bold'), height =1, width=19, bd=1, command = ViewInvalidImport)
    btnInValidImport.place(x=332,y=6)

    L3 = Label(window, text = "Duplicated ShotID", font=("arial", 10,'bold'),bg = "red").place(x=555,y=35)
    txtDuplicatedShotID  = Entry(window, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
    txtDuplicatedShotID.place(x=475,y=35)
    btnDuplicatedShotIDImport= Button(window, text="View Duplicated ShotID", font=('aerial', 9, 'bold'), height =1, width=19, bd=2, command = ViewDuplicatedShotIDImport)
    btnDuplicatedShotIDImport.place(x=332,y=35)
    btnExportTestiFy = Button(window, text="Export For Testif-i", font=('aerial', 9, 'bold'), height =1, width=16, bd=1, command = ExportValidForTest_i_Fy)
    btnExportTestiFy.place(x=1000,y=6)
    btnValidImport = Button(window, text="Valid For Analysis", font=('aerial', 9, 'bold'), height =1, width=15, bd=1, command = ViewTotalImport)
    btnValidImport.place(x=1134,y=6)
    txtTotalEntries  = Entry(window, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 10)
    txtTotalEntries.place(x=1250,y=6)
    L4 = Label(window, text = "UTC Time Offset (+7 Hours from MST):", font=("arial", 10,'bold'),bg = "green").place(x=1000,y=35)
    Entrytxt_TimeOffset  = Entry(window, font=('aerial', 12, 'bold'),textvariable=OffsetTimeUTC, width = 10)
    Entrytxt_TimeOffset.place(x=1250,y=35)

    ## Bottom
    btnImportPFSLog = Button(window, text="Import PFS Log", font=('aerial', 9, 'bold'), height =1, width=14, bd=4, command = ImportPFSLogFile)
    btnImportPFSLog.place(x=2,y=670)
    btnExportLB = Button(window, text="Export ListBox", font=('aerial', 9, 'bold'), height =1, width=12, bd=4, command = ExportListBoxPFS)
    btnExportLB.place(x=940,y=670)
    btnDelete = Button(window, text="Delete Selected Valid", font=('aerial', 9, 'bold'), height =1, width=18, bd=4, command = DeleteSelectedImportData)
    btnDelete.place(x=1040,y=670)
    btnClearView = Button(window, text="Clear View", font=('aerial', 9, 'bold'), height =1, width=10, bd=4, command = ClearView)
    btnClearView.place(x=1181,y=670)
    btnExit = Button(window, text="Exit Import", font=('aerial', 9, 'bold'), height =1, width=10, bd=4, command = iExit)
    btnExit.place(x=1267,y=670)












