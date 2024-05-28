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
import openpyxl
import csv
import time
import datetime
import numpy as np

def SourceLink_PSS_LogIMPORT():
    Default_Date_today   = datetime.date.today()
    window = Tk()
    window.title ("Eagle SourceLink PSS Log Import Wizard")
    window.geometry("1350x680+10+0")
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
    tree.pack()
    # All Functions defining 

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
        conn = sqlite3.connect("SourceLink_Log.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_TEMP ORDER BY `ShotID` ASC ;", conn)
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
        UTC_Offset_Hours = (Entrytxt_TimeOffset.get())
        UTC_Offset_Hours = float(UTC_Offset_Hours)
        tree.delete(*tree.get_children())
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedShotID.delete(0,END)
        conn = sqlite3.connect("SourceLink_Log.db")
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
        TotalEntries()
        DuplicatedShotIDEntries()


    def ViewDuplicatedShotIDImport():
        tree.delete(*tree.get_children())
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedShotID.delete(0,END)
        conn = sqlite3.connect("SourceLink_Log.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_DuplicatedShotID ORDER BY `ShotID` ASC ;", conn)
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
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_TEMP ;", conn)
        data = pd.DataFrame(Complete_df)
        TotalEntries = len(data)       
        txtTotalEntries.insert(tk.END,TotalEntries)              
        conn.commit()
        conn.close()

    def InvalidEntries():
        conn = sqlite3.connect("SourceLink_Log.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_INVALID_NULL ;", conn)
        data = pd.DataFrame(Complete_df)
        TotalInvalidEntries = len(data)       
        txtInvalidEntries.insert(tk.END,TotalInvalidEntries)              
        conn.commit()
        conn.close()

    def DuplicatedShotIDEntries():
        conn = sqlite3.connect("SourceLink_Log.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_DuplicatedShotID ;", conn)
        data = pd.DataFrame(Complete_df)
        TotalDuplicatedEntries = len(data)       
        txtDuplicatedShotID.insert(tk.END,TotalDuplicatedEntries)              
        conn.commit()
        conn.close()

    def UpdateDuplicatedShotID():
        cur_id = tree.focus()
        selvalue = tree.item(cur_id)['values']
        Length_Selected  =  (len(selvalue))
        if Length_Selected != 0:
            for item in tree.selection():
                list_item = (tree.item(item, 'values'))
                con= sqlite3.connect("SourceLink_Log.db")
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
            tkinter.messagebox.showinfo("Exchange Duplicate ShotID Message","Selected List of Entries Added To Imported Database")
            ViewTotalImport()
        else:
            tkinter.messagebox.showinfo("Exchange Duplicate ShotID Message","Please Select List of Entries To Exchange Imported Database")

      
    def DeleteSelectedImportData():
        iDelete = tkinter.messagebox.askyesno("Delete Entry", "Confirm if you want to Delete")
        if iDelete >0:
            txtTotalEntries.delete(0,END)
            conn = sqlite3.connect("SourceLink_Log.db")
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
        conn = sqlite3.connect("SourceLink_Log.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_PSSLog_TEMP ORDER BY `ShotID` ASC ;", conn)
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
        Length_Export_ListBox  =  len(Export_ListBox)
        if Length_Export_ListBox >0:
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

    def FixedCorruptedPSSImport():
        conn = sqlite3.connect("SourceLink_Log.db")
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
            con= sqlite3.connect("SourceLink_Log.db")
            cur=con.cursor()                
            DATA_VALID_PSS.to_sql('Eagle_PSSLog_TEMP',con, if_exists="replace",  index_label='DataBase_ID')
            con.commit()
            cur.close()
            con.close()
            ViewTotalImport()

    def ImportPSSLogFile():
        tree.delete(*tree.get_children())
        txttotalRAWEntries.delete(0,END)
        txtInvalidRemoved.delete(0,END)
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedShotID.delete(0,END)
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
                RAW_DUMP_ImportedPSS_DF = pd.DataFrame(concatDf)
                len_RAW_Dump            = len(RAW_DUMP_ImportedPSS_DF)

                # Separating InValid with Shot ID is Null
                Invalid_PSS_DF    = pd.DataFrame(concatDf)
                Invalid_PSS_DF    = Invalid_PSS_DF[pd.to_numeric(Invalid_PSS_DF.ShotID,errors='coerce').isnull()]                    
                Invalid_PSS_DF    = Invalid_PSS_DF.reset_index(drop=True)
                Data_Invalid_PSS  = pd.DataFrame(Invalid_PSS_DF)
                length_Data_Invalid_PSS = len(Data_Invalid_PSS)
                
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

                # Separating Valid with Shot ID Not Duplicated
                DATA_VALID_PSS = Valid_PSS_DF.loc[Valid_PSS_DF.DuplicatedEntries == False, 'ShotID': 'Sweep_Number']
                DATA_VALID_PSS = DATA_VALID_PSS.reset_index(drop=True)
                DATA_VALID_PSS = pd.DataFrame(DATA_VALID_PSS)

                # Separating Valid with Shot ID Duplicated
                DATA_DuplicatedShotID = Valid_PSS_DF.loc[Valid_PSS_DF.DuplicatedEntries == True, 'ShotID': 'Sweep_Number']
                DATA_DuplicatedShotID = DATA_DuplicatedShotID.reset_index(drop=True)
                DATA_DuplicatedShotID = pd.DataFrame(DATA_DuplicatedShotID)
                len_DATA_DuplicatedShotID = len(DATA_DuplicatedShotID)

                # Connect To Database and Export DF
                Invalid_Removed = length_Data_Invalid_PSS+len_DATA_DuplicatedShotID
                txttotalRAWEntries.insert(tk.END,len_RAW_Dump)
                txtInvalidRemoved.insert(tk.END,Invalid_Removed)
                con= sqlite3.connect("SourceLink_Log.db")
                cur=con.cursor()                
                DATA_VALID_PSS.to_sql('Eagle_PSSLog_TEMP',con, if_exists="replace",  index_label='DataBase_ID')
                Data_Invalid_PSS.to_sql('Eagle_PSSLog_INVALID_NULL',con, if_exists="replace", index_label='DataBase_ID')
                DATA_DuplicatedShotID.to_sql('Eagle_PSSLog_DuplicatedShotID',con, if_exists="replace", index_label='DataBase_ID')
                con.commit()
                cur.close()
                con.close()
                ViewTotalImport()
        else:
            tkinter.messagebox.showinfo("Import PSS File Message","Please Select PSS Files To Import")

    ### top Actions Menu
    LabeltotalRAWEntries = Label(DataFrameTOP, text = "Total Raw Imported : ", font=("arial", 10,'bold'),bg = "cadet blue").grid(row=0,column=0, sticky ="W", padx= 2)
    txttotalRAWEntries  = Entry(DataFrameTOP, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
    txttotalRAWEntries.grid(row=0,column=1, sticky ="W", padx= 2)

    LabelInvalidRemoved = Label(DataFrameTOP, text = "Total Invalid Removed : ", font=("arial", 10,'bold'),bg = "cadet blue").grid(row=0,column=2, sticky ="W", padx= 180)
    txtInvalidRemoved  = Entry(DataFrameTOP, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
    txtInvalidRemoved.grid(row=0,column=2, sticky ="W", padx= 340)

    btnExportTestiFy = Button(DataFrameTOP, text="Export For Testif-i", font=('aerial', 9, 'bold'), height =1, width=16, bd=2, command = ExportValidForTest_i_Fy)
    btnExportTestiFy.grid(row=0,column=3, sticky ="W", padx= 1)

    btnValidImport = Button(DataFrameTOP, text="Valid For Analysis", font=('aerial', 9, 'bold'), height =1, width=15, bd=2, command = ViewTotalImport)
    btnValidImport.grid(row=0,column=5, sticky ="W", padx= 125)
    txtTotalEntries  = Entry(DataFrameTOP, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 9)
    txtTotalEntries.grid(row=0,column=5, sticky ="W", padx= 35)

    ### Bottom Actions Menu
    btnImportPSSLog = Button(window, text="Import PSS Log", font=('aerial', 9, 'bold'), height =1, width=14, bd=4, command = ImportPSSLogFile)
    btnImportPSSLog.place(x=2,y=608)

    L1 = Label(window, text = "UTC Time Offset :", font=("arial", 10,'bold'),bg = "cadet blue").place(x=2,y=645)
    Entrytxt_TimeOffset  = Entry(window, font=('aerial', 12, 'bold'),textvariable=OffsetTimeUTC, width = 6)
    Entrytxt_TimeOffset.place(x=120,y=645)

    btnFixedCorruptedPSSImport= Button(window, text="Fix Corrupted PSS Import", font=('aerial', 9, 'bold'), height =1, width=23, bd=4, command = FixedCorruptedPSSImport)
    btnFixedCorruptedPSSImport.place(x=116,y=608)

    btnDuplicatedShotExchange= Button(window, text="Exchange Duplicated ShotID", font=('aerial', 9, 'bold'), height =1, width=23, bd=4, command = UpdateDuplicatedShotID)
    btnDuplicatedShotExchange.place(x=296,y=608)

    txtDuplicatedShotID  = Entry(window, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
    txtDuplicatedShotID.place(x=600,y=608)
    btnDuplicatedShotIDImport= Button(window, text="Duplicated Invalid ShotID", font=('aerial', 9, 'bold'), height =1, width=20, bg = "red", bd=2, command = ViewDuplicatedShotIDImport)
    btnDuplicatedShotIDImport.place(x=680,y=608)

    btnExportLB = Button(window, text="Export ListBox", font=('aerial', 9, 'bold'), height =1, width=12, bd=4, command = ExportListBoxPSS)
    btnExportLB.place(x=940,y=608)
    btnDelete = Button(window, text="Delete Selected Valid", font=('aerial', 9, 'bold'), height =1, width=18, bd=4, command = DeleteSelectedImportData)
    btnDelete.place(x=1040,y=608)
    btnClearView = Button(window, text="Clear View", font=('aerial', 9, 'bold'), height =1, width=10, bd=4, command = ClearView)
    btnClearView.place(x=1181,y=608)
    btnExit = Button(window, text="Exit Import", font=('aerial', 9, 'bold'), height =1, width=10, bd=4, command = iExit)
    btnExit.place(x=1267,y=608)

    txtInvalidEntries  = Entry(window, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
    txtInvalidEntries.place(x=600,y=638)
    btnInValidImport= Button(window, text="Null Invalid ShotID/FFID", font=('aerial', 9, 'bold'), height =1, width=20, bg = "red", bd=2, command = ViewInvalidImport)
    btnInValidImport.place(x=680,y=638)







