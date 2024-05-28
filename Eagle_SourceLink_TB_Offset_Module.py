#Front End
import os
from tkinter import*
import tkinter.messagebox
import Eagle_SourceLink_TB_Offset_Module_BackEnd
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
from pandas.tseries.offsets import DateOffset

def SourceLink_TB_Offset():
    Default_Date_today   = datetime.date.today()
    window = Tk()
    window.title ("Eagle SourceLink Timebreak Report Import Wizard")
    window.geometry("1246x685+10+0")
    window.config(bg="cadet blue")
    window.resizable(0, 0)
    TableMargin = Frame(window, bd = 2, padx= 1, pady= 1, relief = RIDGE)
    TableMargin.place(x=1, y =32)
    scrollbarx = Scrollbar(TableMargin, orient=HORIZONTAL)
    scrollbary = Scrollbar(TableMargin, orient=VERTICAL)
    tree = ttk.Treeview(TableMargin, column=("column1", "column2", "column3", "column4", "column5",
                                             "column6", "column7", "column8", "column9", "column10", "column11"), height=27, show='headings')
    scrollbary.config(command=tree.yview)
    scrollbary.pack(side=RIGHT, fill=Y)
    scrollbarx.config(command=tree.xview)
    scrollbarx.pack(side=BOTTOM, fill=X)        
    tree.heading("#1", text="Trigger Index", anchor=W)
    tree.heading("#2", text="Profile Id", anchor=W)
    tree.heading("#3", text="Shot Number", anchor=W)
    tree.heading("#4", text="Ep Number", anchor=W)
    tree.heading("#5", text="Shot Line", anchor=W)        
    tree.heading("#6", text="Shot Station", anchor=W)
    tree.heading("#7", text="Shot UtcDateTime", anchor=W)
    tree.heading("#8", text="Latitude", anchor=W)        
    tree.heading("#9", text="Longitude" ,anchor=W)
    tree.heading("#10", text="Shot Status", anchor=W)
    tree.heading("#11", text="Comments", anchor=W)
    tree.column('#1', stretch=NO, minwidth=0, width=90)            
    tree.column('#2', stretch=NO, minwidth=0, width=80)
    tree.column('#3', stretch=NO, minwidth=0, width=100)
    tree.column('#4', stretch=NO, minwidth=0, width=100)
    tree.column('#5', stretch=NO, minwidth=0, width=100)
    tree.column('#6', stretch=NO, minwidth=0, width=120)
    tree.column('#7', stretch=NO, minwidth=0, width=180)
    tree.column('#8', stretch=NO, minwidth=0, width=90)
    tree.column('#9', stretch=NO, minwidth=0, width=100)
    tree.column('#10', stretch=NO, minwidth=0, width=90)
    tree.column('#11', stretch=NO, minwidth=0, width=170)
    style = ttk.Style()
    style.theme_use("clam")
    style.configure(".", font=('aerial', 8), foreground="black")
    style.configure("Treeview", foreground='black')
    style.configure("Treeview.Heading",font=('aerial', 8,'bold'), background='Ghost White', foreground='blue',fieldbackground='Ghost White')
    tree.pack()

    # All Functions defining 

    def iExit():
        iExit= tkinter.messagebox.askyesno("Eagle SourceLink TB Import Wizard", "Confirm if you want to exit")
        if iExit >0:
            window.destroy()
            return

    def ViewTotalImport():
        tree.delete(*tree.get_children())
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedEntries.delete(0,END)
        conn = sqlite3.connect("SourceLink_Microseconds_Offset.db")
        Complete_df_TB = pd.read_sql_query("SELECT * FROM Eagle_SOURCELINK_TB_TEMP_Vib ORDER BY `FileNum` ASC ;", conn)
        data = pd.DataFrame(Complete_df_TB)
        data = data.reset_index(drop=True)
        for each_rec in range(len(data)):
            tree.insert("", tk.END, values=list(data.loc[each_rec]))
        TotalEntries = len(data)       
        txtTotalEntries.insert(tk.END,TotalEntries)              
        conn.commit()
        conn.close()
        InvalidEntries()
        DuplicatedEntries()


    def GenerateMicroSeconds_Time_Offset_TB():
        OffsetMicroSecValues = txtMicrosecondsOffsetEntries.get()     
        if(len(OffsetMicroSecValues)!=0):
            OffsetMicroSecValues = int(txtMicrosecondsOffsetEntries.get())
            tree.delete(*tree.get_children())
            txtTotalEntries.delete(0,END)
            txtInvalidEntries.delete(0,END)
            txtDuplicatedEntries.delete(0,END)
            conn = sqlite3.connect("SourceLink_Microseconds_Offset.db")
            Complete_df_TB = pd.read_sql_query("SELECT * FROM Eagle_SOURCELINK_TB_TEMP_Vib ORDER BY `FileNum` ASC ;", conn)
            data = pd.DataFrame(Complete_df_TB)
            data = data.reset_index(drop=True)
            data['ShotUtcDateTime'] = pd.to_datetime(data['ShotUtcDateTime']) + pd.DateOffset(microseconds=OffsetMicroSecValues)
            data['ShotUtcDateTime'] = pd.to_datetime(data['ShotUtcDateTime']).dt.strftime('%Y/%m/%d %H:%M:%S.%f')        
            data = data.reset_index(drop=True)
            MicroSecOffset_TB_DF = pd.DataFrame(data)

            ## Connecting to SQL DB        
            con= sqlite3.connect("SourceLink_Microseconds_Offset.db")
            cur=con.cursor()                
            MicroSecOffset_TB_DF.to_sql('Eagle_SOURCELINKTB_Microseconds_Offset_Vib',con, if_exists="replace", index=False)            
            con.commit()
            cur.close()
            con.close()
            ## Populate Offset Data In GUI 
            TotalEntries = len(MicroSecOffset_TB_DF)
            txtTotalEntries.insert(tk.END,TotalEntries)            
            for each_rec in range(len(MicroSecOffset_TB_DF)):
                tree.insert("", tk.END, values=list(MicroSecOffset_TB_DF.loc[each_rec]))
        else:
            tkinter.messagebox.showinfo("Time Offset Input Error Message","Please Input Time Offset Value In Microseconds To Generate Time Offset TB")

    def ViewMicroSeconds_Time_Offset_TB():
        tree.delete(*tree.get_children())
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedEntries.delete(0,END)
        conn = sqlite3.connect("SourceLink_Microseconds_Offset.db")
        Complete_df_TB = pd.read_sql_query("SELECT * FROM Eagle_SOURCELINKTB_Microseconds_Offset_Vib ORDER BY `FileNum` ASC ;", conn)
        data = pd.DataFrame(Complete_df_TB)
        data = data.reset_index(drop=True)
        for each_rec in range(len(data)):
            tree.insert("", tk.END, values=list(data.loc[each_rec]))
        TotalEntries = len(data)       
        txtTotalEntries.insert(tk.END,TotalEntries)              
        conn.commit()
        conn.close()


    def ExportMicroSeconds_Time_Offset_TB():
        conn = sqlite3.connect("SourceLink_Microseconds_Offset.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_SOURCELINKTB_Microseconds_Offset_Vib ORDER BY `FileNum` ASC ;", conn)
        Export_MicroSeconds_Time_Offset_TB  = pd.DataFrame(Complete_df)
        Export_MicroSeconds_Time_Offset_TB.rename(columns={'TriggerIndex':'TriggerIndex', 'Unit_ID':' ProfileId', 'FileNum':' ShotNumber',
                                                                'EPNumber':' EpNumber', 'SourceLine':' ShotLine', 'SourceStation':' ShotStation',
                                                                'ShotUtcDateTime':' ShotUtcDateTime','Latitude':' Latitude','Longitude':' Longitude',
                                                                'ShotStatus':' ShotStatus', 'TBComment':' Comment'},inplace = True)
        Export_MicroSeconds_Time_Offset_TB  = Export_MicroSeconds_Time_Offset_TB.reset_index(drop=True)
        filename = tkinter.filedialog.asksaveasfilename(initialdir = "/" ,title = "Select file" ,\
                    defaultextension='.csv', filetypes = (("CSV file",".csv"),("Excel file",".xlsx")))
        if len(filename) >0:
            if filename.endswith('.csv'):
                Export_MicroSeconds_Time_Offset_TB.to_csv(filename, index=None, date_format = '%Y/%m/%d %H:%M:%S.%f')
                tkinter.messagebox.showinfo("MicroSeconds Time Offset TB Export","MicroSeconds Time Offset TB Saved as CSV")
            else:
                Export_MicroSeconds_Time_Offset_TB.to_excel(filename, sheet_name='MasterTB', index=False)
                tkinter.messagebox.showinfo("MicroSeconds Time Offset TB Export","MicroSeconds Time Offset TB Saved as Excel")
        else:
            tkinter.messagebox.showinfo("MicroSeconds Time Offset TB Export Message","Please Select File Name To Export")
                    
        conn.commit()
        conn.close()


    def ViewInvalidImport():
        tree.delete(*tree.get_children())
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedEntries.delete(0,END)
        conn = sqlite3.connect("SourceLink_Microseconds_Offset.db")
        Complete_df_TB = pd.read_sql_query("SELECT * FROM Eagle_SOURCELINK_TB_INVALID_Vib ORDER BY `FileNum` ASC ;", conn)
        data = pd.DataFrame(Complete_df_TB)
        data = data.reset_index(drop=True)
        for each_rec in range(len(data)):
            tree.insert("", tk.END, values=list(data.loc[each_rec]))
        TotalInvalidEntries = len(data)       
        txtInvalidEntries.insert(tk.END,TotalInvalidEntries)              
        conn.commit()
        conn.close()
        TotalEntries()
        DuplicatedEntries()

    def ViewDuplicatedImport():
        tree.delete(*tree.get_children())
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedEntries.delete(0,END)
        conn = sqlite3.connect("SourceLink_Microseconds_Offset.db")
        Complete_df_TB = pd.read_sql_query("SELECT * FROM Eagle_SOURCELINK_TB_Duplicated_Vib ORDER BY `FileNum` ASC ;", conn)
        data = pd.DataFrame(Complete_df_TB)
        data = data.reset_index(drop=True)
        for each_rec in range(len(data)):
            tree.insert("", tk.END, values=list(data.loc[each_rec]))
        TotalDuplicatedEntries = len(data)       
        txtDuplicatedEntries.insert(tk.END,TotalDuplicatedEntries)              
        conn.commit()
        conn.close()
        InvalidEntries()
        TotalEntries()
        

    def ClearView():
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedEntries.delete(0,END)
        tree.delete(*tree.get_children())

    def TotalEntries():
        conn = sqlite3.connect("SourceLink_Microseconds_Offset.db")
        Complete_df_TB = pd.read_sql_query("SELECT * FROM Eagle_SOURCELINK_TB_TEMP_Vib ;", conn)
        data = pd.DataFrame(Complete_df_TB)
        TotalEntries = len(data)       
        txtTotalEntries.insert(tk.END,TotalEntries)              
        conn.commit()
        conn.close()

    def InvalidEntries():
        conn = sqlite3.connect("SourceLink_Microseconds_Offset.db")
        Complete_df_TB = pd.read_sql_query("SELECT * FROM Eagle_SOURCELINK_TB_INVALID_Vib ;", conn)
        data = pd.DataFrame(Complete_df_TB)
        TotalInvalidEntries = len(data)       
        txtInvalidEntries.insert(tk.END,TotalInvalidEntries)              
        conn.commit()
        conn.close()

    def DuplicatedEntries():
        conn = sqlite3.connect("SourceLink_Microseconds_Offset.db")
        Complete_df_TB = pd.read_sql_query("SELECT * FROM Eagle_SOURCELINK_TB_Duplicated_Vib ;", conn)
        data = pd.DataFrame(Complete_df_TB)
        TotalDuplicatedEntries = len(data)       
        txtDuplicatedEntries.insert(tk.END,TotalDuplicatedEntries)              
        conn.commit()
        conn.close()
      
    def DeleteSelectedImportData():
        iDelete = tkinter.messagebox.askyesno("Delete Entry", "Confirm if you want to Delete")
        if iDelete >0:
            txtTotalEntries.delete(0,END)
            conn = sqlite3.connect("SourceLink_Microseconds_Offset.db")
            cur = conn.cursor()                
            for selected_item in tree.selection():
                cur.execute("DELETE FROM Eagle_SOURCELINK_TB_TEMP_Vib WHERE TriggerIndex =? AND FileNum =? AND EPNumber =? AND SourceLine =? AND \
                            SourceStation =? ",\
                            (tree.set(selected_item, '#1'), tree.set(selected_item, '#3'),tree.set(selected_item, '#4'),
                             tree.set(selected_item, '#5'), tree.set(selected_item, '#6'),)) 
                conn.commit()
                tree.delete(selected_item)
            conn.commit()
            conn.close()
            TotalEntries()
            return


    def ImportTBFile():
        tree.delete(*tree.get_children())
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        fileList = askopenfilenames(initialdir = "/", title = "Import SourceLink Time Break Files" , filetypes=[('CSV File', '*.csv'), ('Excel File', ('*.xls', '*.xlsx'))])
        Length_fileList  =  len(fileList)
        if Length_fileList >0:
            if fileList:
                df_TBList = []           
                for filename in fileList:
                    if filename.endswith('.csv'):
                        df_TB             = pd.read_csv(filename, sep=',' , low_memory=False)
                        df_TB             = df_TB.iloc[:,:]
                        TriggerIndex      = df_TB.loc[:,'TriggerIndex']
                        Unit_ID           = df_TB.loc[:,' ProfileId']
                        FileNum           = df_TB.loc[:,' ShotNumber']
                        EPNumber          = df_TB.loc[:,' EpNumber']
                        SourceLine        = df_TB.loc[:,' ShotLine']
                        SourceStation     = df_TB.loc[:,' ShotStation']
                        ShotUtcDateTime   = df_TB.loc[:,' ShotUtcDateTime']
                        Latitude          = df_TB.loc[:,' Latitude']
                        Longitude         = df_TB.loc[:,' Longitude']
                        ShotStatus        = df_TB.loc[:,' ShotStatus']
                        TBComment         = df_TB.loc[:,' Comment']                    
                        column_names = [TriggerIndex, Unit_ID, FileNum, EPNumber, SourceLine, SourceStation, ShotUtcDateTime,
                                        Latitude, Longitude, ShotStatus, TBComment]
                        catdf_TB = pd.concat (column_names,axis=1,ignore_index =True)
                        df_TBList.append(catdf_TB) 
                    else:
                        df_TB             = pd.read_excel(filename)
                        df_TB             = df_TB.iloc[:,:]
                        TriggerIndex      = df_TB.loc[:,'TriggerIndex']
                        Unit_ID           = df_TB.loc[:,' ProfileId']
                        FileNum           = df_TB.loc[:,' ShotNumber']
                        EPNumber          = df_TB.loc[:,' EpNumber']
                        SourceLine        = df_TB.loc[:,' ShotLine']
                        SourceStation     = df_TB.loc[:,' ShotStation']
                        ShotUtcDateTime   = df_TB.loc[:,' ShotUtcDateTime']
                        Latitude          = df_TB.loc[:,' Latitude']
                        Longitude         = df_TB.loc[:,' Longitude']
                        ShotStatus        = df_TB.loc[:,' ShotStatus']
                        TBComment         = df_TB.loc[:,' Comment']                    
                        column_names = [TriggerIndex, Unit_ID, FileNum, EPNumber, SourceLine, SourceStation, ShotUtcDateTime,
                                        Latitude, Longitude, ShotStatus, TBComment]
                        catdf_TB = pd.concat (column_names,axis=1,ignore_index =True)
                        df_TBList.append(catdf_TB) 

                concatdf_TB = pd.concat(df_TBList,axis=0, ignore_index =True)
                concatdf_TB.rename(columns={0:'TriggerIndex', 1:'Unit_ID', 2:'FileNum', 3:'EPNumber', 4:'SourceLine',
                                         5:'SourceStation', 6:'ShotUtcDateTime', 7:'Latitude', 8:'Longitude',
                                         9:'ShotStatus', 10:'TBComment' },inplace = True)
                concatdf_TB = concatdf_TB.reset_index(drop=True)
                
                ## Separating Invalid COG
                Invalid_TB_df   = pd.DataFrame(concatdf_TB)
                Invalid_TB_df   = Invalid_TB_df[pd.to_numeric(Invalid_TB_df.FileNum,errors='coerce').isnull()]               
                Invalid_TB_df   = Invalid_TB_df.reset_index(drop=True)
                Data_Invalid_TB = pd.DataFrame(Invalid_TB_df)
                
                ## Separating Valid COG 
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
                Valid_TB_df                     = Valid_TB_df.reset_index(drop=True)   
                DATA_VALID_TB                   = pd.DataFrame(Valid_TB_df)
                DATA_VALID_TB['DuplicatedEntries'] = DATA_VALID_TB .sort_values(by =['ShotUtcDateTime']).duplicated(['ShotUtcDateTime','TriggerIndex','Unit_ID','FileNum','EPNumber','SourceLine','SourceStation'],keep='last')
                DATA_VALID_TB                      = DATA_VALID_TB.loc[DATA_VALID_TB.DuplicatedEntries == False, 'TriggerIndex': 'TBComment']
                DATA_VALID_TB                      = DATA_VALID_TB.reset_index(drop=True)
                DATA_VALID_TB                      = pd.DataFrame(DATA_VALID_TB)

                DATA_VALID_TB_DuplicatedShotID     = pd.DataFrame(Valid_TB_df)
                DATA_VALID_TB_DuplicatedShotID['DuplicatedEntries'] = DATA_VALID_TB_DuplicatedShotID.sort_values(by =['ShotUtcDateTime']).duplicated(['ShotUtcDateTime','TriggerIndex','Unit_ID','FileNum','EPNumber','SourceLine','SourceStation'],keep='last')
                DATA_VALID_TB_DuplicatedShotID                      = DATA_VALID_TB_DuplicatedShotID.loc[DATA_VALID_TB_DuplicatedShotID.DuplicatedEntries == True, 'TriggerIndex': 'TBComment']
                DATA_VALID_TB_DuplicatedShotID                      = DATA_VALID_TB_DuplicatedShotID.reset_index(drop=True)
                DATA_VALID_TB_DuplicatedShotID                      = pd.DataFrame(DATA_VALID_TB_DuplicatedShotID)

                ## Connecting to SQL DB        
                con= sqlite3.connect("SourceLink_Microseconds_Offset.db")
                cur=con.cursor()                
                DATA_VALID_TB.to_sql('Eagle_SOURCELINK_TB_TEMP_Vib',con, if_exists="replace", index=False)
                Data_Invalid_TB.to_sql ('Eagle_SOURCELINK_TB_INVALID_Vib',con, if_exists="replace", index=False)
                DATA_VALID_TB_DuplicatedShotID.to_sql ('Eagle_SOURCELINK_TB_Duplicated_Vib',con, if_exists="replace", index=False)                
                con.commit()
                cur.close()
                con.close()
                ViewTotalImport()
        else:
            tkinter.messagebox.showinfo("Import TB File Message","Please Select TimeBreak Files To Import")
            


    ##### Entry Wizard
    txtInvalidEntries  = Entry(window, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
    txtInvalidEntries.place(x=425,y=6)
    txtTotalEntries  = Entry(window, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 10)
    txtTotalEntries.place(x=1150,y=6)
    txtDuplicatedEntries  = Entry(window, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
    txtDuplicatedEntries.place(x=711,y=6)
    txtMicrosecondsOffsetEntries  = Entry(window, font=('aerial', 13, 'bold'),textvariable=IntVar(), width = 24)
    txtMicrosecondsOffsetEntries.place(x=500,y=626)
    L1 = Label(window, text = "Source Link Timebreak Details:", font=("arial", 10,'bold'),bg = "green").place(x=2,y=6)
    L2 = Label(window, text = "Enter Time Offset Value In Microseconds (+ or -) :", font=("arial", 12,'bold'),bg = "orange").place(x=120,y=626)


    ### Button Wizard
    btnInValidImport= Button(window, text="View Invalid Import", font=('aerial', 9, 'bold'), height =1, width=16, bd=1, command = ViewInvalidImport)
    btnInValidImport.place(x=300,y=6)
    btnValidImport = Button(window, text="View Valid Import", font=('aerial', 9, 'bold'), height =1, width=15, bd=1, command = ViewTotalImport)
    btnValidImport.place(x=1032,y=6)
    btnDuplicatedImport= Button(window, text="View Duplicated", font=('aerial', 9, 'bold'), height =1, width=14, bd=1, command = ViewDuplicatedImport)
    btnDuplicatedImport.place(x=600,y=6)
    btnClearView = Button(window, text="Clear View", font=('aerial', 9, 'bold'), height =1, width=9, bd=1, command = ClearView)
    btnClearView.place(x=950,y=6)

    btnImportTBFile = Button(window, text="Import TB File", font=('aerial', 9, 'bold'), height =1, width=12, bd=2, command = ImportTBFile)
    btnImportTBFile.place(x=2,y=626)

    btnTimeOffsetTBFile = Button(window, text="Generate Time Offset TB", font=('aerial', 9, 'bold'), height =1, width=20, bd=2, command = GenerateMicroSeconds_Time_Offset_TB)
    btnTimeOffsetTBFile.place(x=730,y=626)

    btnExportTimeOffsetTBFile = Button(window, text="Export Time Offset TB", font=('aerial', 9, 'bold'), height =1, width=20, bd=2, command = ExportMicroSeconds_Time_Offset_TB)
    btnExportTimeOffsetTBFile.place(x=730,y=656)

    btnViewTimeOffsetTBFile = Button(window, text="View Time Offset TB", font=('aerial', 9, 'bold'), height =1, width=18, bd=2, command = ViewMicroSeconds_Time_Offset_TB)
    btnViewTimeOffsetTBFile.place(x=890,y=626)

    btnDelete = Button(window, text="Delete Valid", font=('aerial', 9, 'bold'), height =1, width=10, bd=2, command = DeleteSelectedImportData)
    btnDelete.place(x=1080,y=626)

    btnExit = Button(window, text="Exit Import", font=('aerial', 9, 'bold'), height =1, width=10, bd=2, command = iExit)
    btnExit.place(x=1165,y=626)













