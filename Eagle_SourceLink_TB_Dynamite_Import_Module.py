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
import numpy as np
import openpyxl
import csv
import time
import datetime

def SourceLink_TB_Dynamite_LogIMPORT():
    Default_Date_today   = datetime.date.today()
    window = Tk()
    window.title ("Eagle SourceLink Dynamite Timebreak Report Import Wizard")
    window.geometry("1250x650+10+0")
    window.config(bg="cadet blue")
    window.resizable(0, 0)
    TableMargin = Frame(window, bd = 2, padx= 10, pady= 8, relief = RIDGE)
    TableMargin.pack(side=TOP)
    TableMargin.pack(side=LEFT)
    scrollbarx = Scrollbar(TableMargin, orient=HORIZONTAL)
    scrollbary = Scrollbar(TableMargin, orient=VERTICAL)
    tree = ttk.Treeview(TableMargin, column=("column1", "column2", "column3", "column4", "column5",
                                             "column6", "column7", "column8", "column9", "column10", "column11",
                                             "column12", "column13"), height=26, show='headings')
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
    tree.heading("#11", text="Uphole", anchor=W)
    tree.heading("#12", text="Comments", anchor=W)
    tree.heading("#13", text="Process", anchor=W)

    tree.column('#1', stretch=NO, minwidth=0, width=90)            
    tree.column('#2', stretch=NO, minwidth=0, width=80)
    tree.column('#3', stretch=NO, minwidth=0, width=100)
    tree.column('#4', stretch=NO, minwidth=0, width=80)
    tree.column('#5', stretch=NO, minwidth=0, width=90)
    tree.column('#6', stretch=NO, minwidth=0, width=100)
    tree.column('#7', stretch=NO, minwidth=0, width=170)
    tree.column('#8', stretch=NO, minwidth=0, width=90)
    tree.column('#9', stretch=NO, minwidth=0, width=100)
    tree.column('#10', stretch=NO, minwidth=0, width=90)
    tree.column('#11', stretch=NO, minwidth=0, width=70)
    tree.column('#12', stretch=NO, minwidth=0, width=90)
    tree.column('#13', stretch=NO, minwidth=0, width=60)

    style = ttk.Style()
    style.theme_use("clam")
    style.configure(".", font=('aerial', 8), foreground="black")
    style.configure("Treeview", foreground='black')
    style.configure("Treeview.Heading",font=('aerial', 8,'bold'), background='Ghost White', foreground='blue',fieldbackground='Ghost White')
    tree.pack()

    # All Functions defining 

    def iExit():
        iExit= tkinter.messagebox.askyesno("Eagle Dynamite TB Import Wizard", "Confirm if you want to exit")
        if iExit >0:
            window.destroy()
            return

    def ViewTotalImport():
        tree.delete(*tree.get_children())
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedEntries.delete(0,END)
        conn = sqlite3.connect("SourceLink_Dynamite_Log.db")
        Complete_df_TB = pd.read_sql_query("SELECT * FROM Eagle_SOURCELINK_TB_Dynamite ORDER BY `FileNum` ASC ;", conn)
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
        

    def ViewInvalidImport():
        tree.delete(*tree.get_children())
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedEntries.delete(0,END)
        conn = sqlite3.connect("SourceLink_Dynamite_Log.db")
        Complete_df_TB = pd.read_sql_query("SELECT * FROM Eagle_SOURCELINK_TB_Dynamite_INVALID ORDER BY `FileNum` ASC ;", conn)
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
        conn = sqlite3.connect("SourceLink_Dynamite_Log.db")
        Complete_df_TB = pd.read_sql_query("SELECT * FROM Eagle_SOURCELINK_TB_Dynamite_Duplicated ORDER BY `FileNum` ASC ;", conn)
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
        conn = sqlite3.connect("SourceLink_Dynamite_Log.db")
        Complete_df_TB = pd.read_sql_query("SELECT * FROM Eagle_SOURCELINK_TB_Dynamite ;", conn)
        data = pd.DataFrame(Complete_df_TB)
        TotalEntries = len(data)       
        txtTotalEntries.insert(tk.END,TotalEntries)              
        conn.commit()
        conn.close()

    def InvalidEntries():
        conn = sqlite3.connect("SourceLink_Dynamite_Log.db")
        Complete_df_TB = pd.read_sql_query("SELECT * FROM Eagle_SOURCELINK_TB_Dynamite_INVALID ;", conn)
        data = pd.DataFrame(Complete_df_TB)
        TotalInvalidEntries = len(data)       
        txtInvalidEntries.insert(tk.END,TotalInvalidEntries)              
        conn.commit()
        conn.close()

    def DuplicatedEntries():
        conn = sqlite3.connect("SourceLink_Dynamite_Log.db")
        Complete_df_TB = pd.read_sql_query("SELECT * FROM Eagle_SOURCELINK_TB_Dynamite_Duplicated ;", conn)
        data = pd.DataFrame(Complete_df_TB)
        TotalDuplicatedEntries = len(data)       
        txtDuplicatedEntries.insert(tk.END,TotalDuplicatedEntries)              
        conn.commit()
        conn.close()
      
    def DeleteSelectedImportData():
        iDelete = tkinter.messagebox.askyesno("Delete Entry", "Confirm if you want to Delete")
        if iDelete >0:
            txtTotalEntries.delete(0,END)
            conn = sqlite3.connect("SourceLink_Dynamite_Log.db")
            cur = conn.cursor()                
            for selected_item in tree.selection():
                cur.execute("DELETE FROM Eagle_SOURCELINK_TB_Dynamite WHERE TriggerIndex =? AND FileNum =? AND EPNumber =? AND SourceLine =? AND \
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
        fileList = askopenfilenames(initialdir = "/", title = "Import SourceLink Dynamite Time Break Files" , filetypes=[('CSV File', '*.csv'), ('Excel File', ('*.xls', '*.xlsx'))])
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
                        Uphole            = df_TB.loc[:,' Uphole']
                        TBComment         = df_TB.loc[:,' Comment']
                        Process           = df_TB.loc[:,' Process'] 
                        column_names = [TriggerIndex, Unit_ID, FileNum, EPNumber, SourceLine, SourceStation, ShotUtcDateTime,
                                        Latitude, Longitude, ShotStatus, Uphole, TBComment, Process]
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
                        Uphole            = df_TB.loc[:,' Uphole']
                        TBComment         = df_TB.loc[:,' Comment']
                        Process           = df_TB.loc[:,' Process'] 
                        column_names = [TriggerIndex, Unit_ID, FileNum, EPNumber, SourceLine, SourceStation, ShotUtcDateTime,
                                        Latitude, Longitude, ShotStatus, Uphole, TBComment, Process]
                        catdf_TB = pd.concat (column_names,axis=1,ignore_index =True)
                        df_TBList.append(catdf_TB) 

                concatdf_TB = pd.concat(df_TBList,axis=0, ignore_index =True)
                concatdf_TB.rename(columns={0:'TriggerIndex', 1:'Unit_ID', 2:'FileNum', 3:'EPNumber', 4:'SourceLine',
                                         5:'SourceStation', 6:'ShotUtcDateTime', 7:'Latitude', 8:'Longitude',
                                         9:'ShotStatus', 10:'Uphole', 11:'TBComment',12:'Process'},inplace = True)
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
                DATA_VALID_TB                      = DATA_VALID_TB.loc[DATA_VALID_TB.DuplicatedEntries == False, 'TriggerIndex': 'Process']
                DATA_VALID_TB                      = DATA_VALID_TB.reset_index(drop=True)
                DATA_VALID_TB                      = pd.DataFrame(DATA_VALID_TB)

                DATA_VALID_TB_DuplicatedShotID     = pd.DataFrame(Valid_TB_df)
                DATA_VALID_TB_DuplicatedShotID['DuplicatedEntries'] = DATA_VALID_TB_DuplicatedShotID.sort_values(by =['ShotUtcDateTime']).duplicated(['ShotUtcDateTime','TriggerIndex','Unit_ID','FileNum','EPNumber','SourceLine','SourceStation'],keep='last')
                DATA_VALID_TB_DuplicatedShotID                      = DATA_VALID_TB_DuplicatedShotID.loc[DATA_VALID_TB_DuplicatedShotID.DuplicatedEntries == True, 'TriggerIndex': 'Process']
                DATA_VALID_TB_DuplicatedShotID                      = DATA_VALID_TB_DuplicatedShotID.reset_index(drop=True)
                DATA_VALID_TB_DuplicatedShotID                      = pd.DataFrame(DATA_VALID_TB_DuplicatedShotID)

                ## Connecting to SQL DB        
                con= sqlite3.connect("SourceLink_Dynamite_Log.db")
                cur=con.cursor()                
                DATA_VALID_TB.to_sql('Eagle_SOURCELINK_TB_Dynamite',con, if_exists="replace", index=False)
                Data_Invalid_TB.to_sql ('Eagle_SOURCELINK_TB_Dynamite_INVALID',con, if_exists="replace", index=False)
                DATA_VALID_TB_DuplicatedShotID.to_sql ('Eagle_SOURCELINK_TB_Dynamite_Duplicated',con, if_exists="replace", index=False)                
                con.commit()
                cur.close()
                con.close()
                ViewTotalImport()
        else:
            tkinter.messagebox.showinfo("Import Dynamite TB File Message","Please Select Dynamite TimeBreak Files To Import")
            


    ##### Entry Wizard
    txtInvalidEntries  = Entry(window, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
    txtInvalidEntries.place(x=425,y=6)
    txtTotalEntries  = Entry(window, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 10)
    txtTotalEntries.place(x=1150,y=6)
    txtDuplicatedEntries  = Entry(window, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
    txtDuplicatedEntries.place(x=425,y=620)
    L1 = Label(window, text = "Source Link Dynamite Timebreak Details:", font=("arial", 10,'bold'),bg = "green").place(x=2,y=6)
    L2 = Label(window, text = "Null Dynamite ShotID Or File Number", font=("arial", 10,'bold'),bg = "red").place(x=505,y=7)
    L3 = Label(window, text = "Duplicated Dynamite ShotID Or File Number", font=("arial", 10,'bold'),bg = "red").place(x=505,y=621)

    ### Button Wizard
    btnInValidImport= Button(window, text="View Invalid Import", font=('aerial', 9, 'bold'), height =1, width=16, bd=1, command = ViewInvalidImport)
    btnInValidImport.place(x=300,y=6)
    btnValidImport = Button(window, text="View Valid Import", font=('aerial', 9, 'bold'), height =1, width=15, bd=1, command = ViewTotalImport)
    btnValidImport.place(x=1032,y=6)
    btnImportTBFile = Button(window, text="Import Dynamite TB File", font=('aerial', 9, 'bold'), height =1, width=20, bd=4, command = ImportTBFile)
    btnImportTBFile.place(x=2,y=620)

    btnDuplicatedImport= Button(window, text="View Duplicated", font=('aerial', 9, 'bold'), height =1, width=14, bd=1, command = ViewDuplicatedImport)
    btnDuplicatedImport.place(x=314,y=620)
    btnDelete = Button(window, text="Delete Selected Valid", font=('aerial', 9, 'bold'), height =1, width=18, bd=4, command = DeleteSelectedImportData)
    btnDelete.place(x=920,y=620)
    btnClearView = Button(window, text="Clear View", font=('aerial', 9, 'bold'), height =1, width=10, bd=4, command = ClearView)
    btnClearView.place(x=1077,y=620)
    btnExit = Button(window, text="Exit Widget", font=('aerial', 9, 'bold'), height =1, width=10, bd=4, command = iExit)
    btnExit.place(x=1165,y=620)













