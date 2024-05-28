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

def SourceLink_DynamiteTB_SubmitToMasterDB():
    Default_Date_today   = datetime.date.today()
    window = Tk()
    window.title ("Eagle SourceLink Dynamite Timebreak Master Import Wizard")
    window.geometry("1250x780+10+0")
    window.config(bg="cadet blue")
    window.resizable(0, 0)
    TableMargin = Frame(window, bd = 2, padx= 2, pady= 30, relief = RIDGE)
    TableMargin.pack(side=TOP)
    TableMargin.config(bg="cadet blue")
    scrollbarx = Scrollbar(TableMargin, orient=HORIZONTAL)
    scrollbary = Scrollbar(TableMargin, orient=VERTICAL)
    tree = ttk.Treeview(TableMargin, column=("column1", "column2", "column3", "column4", "column5",
                                             "column6", "column7", "column8", "column9", "column10", "column11",
                                             "column12", "column13"), height=28, show='headings')
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
    tree.column('#6', stretch=NO, minwidth=0, width=95)
    tree.column('#7', stretch=NO, minwidth=0, width=170)
    tree.column('#8', stretch=NO, minwidth=0, width=90)
    tree.column('#9', stretch=NO, minwidth=0, width=100)
    tree.column('#10', stretch=NO, minwidth=0, width=90)
    tree.column('#11', stretch=NO, minwidth=0, width=70)
    tree.column('#12', stretch=NO, minwidth=0, width=110)
    tree.column('#13', stretch=NO, minwidth=0, width=68)

    style = ttk.Style()
    style.theme_use("clam")
    style.configure(".", font=('aerial', 8), foreground="black")
    style.configure("Treeview", foreground='black')
    style.configure("Treeview.Heading",font=('aerial', 8,'bold'), background='Ghost White', foreground='blue',fieldbackground='Ghost White')
    tree.pack()

    TitFrame = Frame(window, bd = 2, padx= 5, pady= 4, bg = "orange", relief = RIDGE)
    TitFrame.pack(side = TOP)

    EditTableMargin = Frame(window, bd = 2, width = 1250, height = 80, padx= 1, pady= 1, relief = RIDGE)
    EditTableMargin.pack(side=TOP)


    # All Functions defining

    def update():
        if(len(txtShotNumber.get())!=0):
            conn = sqlite3.connect("SourceLink_Dynamite_Log.db")
            cur = conn.cursor()
            for selected_item in tree.selection():
                cur.execute("DELETE FROM Eagle_SOURCELINK_DYNAMITE_TBMASTER WHERE FileNum =? AND SourceLine = ? AND SourceStation = ?", (tree.set(selected_item, '#3'), tree.set(selected_item, '#5'), tree.set(selected_item, '#6'),))
                conn.commit()
                tree.delete(selected_item)
                conn.close()

        if(len(txtShotNumber.get())!=0):
            try:
                Eagle_SourceLink_Dynamite_Log_BackEnd.addInvRec(txtTriggerIndex.get(), txtProfileID.get(), txtShotNumber.get(), txtEpNumber.get(), txtShotLine.get(), txtShotStation.get(),
                                                           txtShotUtcDateTime.get(), txtLatitude.get(), txtLongitude.get(), txtShotStatus.get(), txtUphole.get(),
                                                           txtComment.get(), txtProcess.get())
                tree.delete(*tree.get_children())
                tree.insert("", tk.END,values=(txtTriggerIndex.get(), txtProfileID.get(), txtShotNumber.get(), txtEpNumber.get(), txtShotLine.get(), txtShotStation.get(),
                                                           txtShotUtcDateTime.get(), txtLatitude.get(), txtLongitude.get(), txtShotStatus.get(), txtUphole.get(),
                                                           txtComment.get(), txtProcess.get()))
            except:
                tkinter.messagebox.showinfo("Update Error","Invalid Data Type")
        else:
                tkinter.messagebox.showinfo("Update Error","Shot ID can not be empty")

    def DynamiteTB_MasterRec_Event(event):
        for nm in tree.selection():
            sd = tree.item(nm, 'values')
            txtTriggerIndex.delete(0,END)
            txtTriggerIndex.insert(tk.END,sd[0])                
            txtProfileID.delete(0,END)
            txtProfileID.insert(tk.END,sd[1])
            txtShotNumber.delete(0,END)
            txtShotNumber.insert(tk.END,sd[2])
            txtEpNumber.delete(0,END)
            txtEpNumber.insert(tk.END,sd[3])
            txtShotLine.delete(0,END)
            txtShotLine.insert(tk.END,sd[4])                
            txtShotStation.delete(0,END)
            txtShotStation.insert(tk.END,sd[5])
            txtShotUtcDateTime.delete(0,END)
            txtShotUtcDateTime.insert(tk.END,sd[6])
            txtLatitude.delete(0,END)
            txtLatitude.insert(tk.END,sd[7])
            txtLongitude.delete(0,END)
            txtLongitude.insert(tk.END,sd[8])
            txtShotStatus.delete(0,END)
            txtShotStatus.insert(tk.END,sd[9])
            txtUphole.delete(0,END)
            txtUphole.insert(tk.END,sd[10])
            txtComment.delete(0,END)
            txtComment.insert(tk.END,sd[11])
            txtProcess.delete(0,END)
            txtProcess.insert(tk.END,sd[12])


    def ClearEditEntry():
        txtTriggerIndex.delete(0,END)            
        txtProfileID.delete(0,END)
        txtShotNumber.delete(0,END)
        txtEpNumber.delete(0,END)
        txtShotLine.delete(0,END)          
        txtShotStation.delete(0,END)
        txtShotUtcDateTime.delete(0,END)    
        txtLatitude.delete(0,END)    
        txtLongitude.delete(0,END)    
        txtShotStatus.delete(0,END)    
        txtUphole.delete(0,END)    
        txtComment.delete(0,END)    
        txtProcess.delete(0,END)
        

    def iExit():
        iExit= tkinter.messagebox.askyesno("Eagle Clean Dynamite TB Import Wizard", "Confirm if you want to exit")
        if iExit >0:
            window.destroy()
            return

    def UpdateMasterTB():
        conn = sqlite3.connect("SourceLink_Dynamite_Log.db")
        Complete_df_TB = pd.read_sql_query("SELECT * FROM Eagle_SOURCELINK_DYNAMITE_TBMASTER ORDER BY `FileNum` ASC;", conn)
        data = pd.DataFrame(Complete_df_TB)
        data ['DuplicatedEntries']=data.sort_values(by =['ShotUtcDateTime']).duplicated(['ShotUtcDateTime','TriggerIndex','Unit_ID','FileNum','EPNumber', 'SourceLine','SourceStation'],keep='last')
        data = data.loc[data.DuplicatedEntries == False, 'TriggerIndex': 'Process']
        data = data.reset_index(drop=True)
        data = pd.DataFrame(data)
        data.to_sql('Eagle_SOURCELINK_DYNAMITE_TBMASTER',conn, if_exists="replace", index=False)
        conn.commit()
        conn.close()    
        TotalEntries()
          
    def ViewMasterTB():
        ClearView()
        conn = sqlite3.connect("SourceLink_Dynamite_Log.db")
        Complete_df_TB = pd.read_sql_query("SELECT * FROM Eagle_SOURCELINK_DYNAMITE_TBMASTER ORDER BY `FileNum` ASC;", conn)
        data = pd.DataFrame(Complete_df_TB)
        data = data.reset_index(drop=True)
        for each_rec in range(len(data)):
            tree.insert("", tk.END, values=list(data.loc[each_rec]))
        TotalEntries = len(data)       
        txtTotalEntries.insert(tk.END,TotalEntries)              
        conn.commit()
        conn.close()

    def ExportMasterTB_DB():
        conn = sqlite3.connect("SourceLink_Dynamite_Log.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_SOURCELINK_DYNAMITE_TBMASTER ORDER BY `FileNum` ASC ;", conn)
        Export_MasterTB_DF  = pd.DataFrame(Complete_df)
        Export_MasterTB_DF.rename(columns={'TriggerIndex':'TriggerIndex', 'Unit_ID':' ProfileId', 'FileNum':' ShotNumber',
                                                                'EPNumber':' EpNumber', 'SourceLine':' ShotLine', 'SourceStation':' ShotStation',
                                                                'ShotUtcDateTime':' ShotUtcDateTime','Latitude':' Latitude','Longitude':' Longitude',
                                                                'ShotStatus':' ShotStatus', 'Uphole':' Uphole','TBComment':' Comment', 'Process':' Process'},inplace = True)
        Export_MasterTB_DF  = Export_MasterTB_DF.reset_index(drop=True)
        filename = tkinter.filedialog.asksaveasfilename(initialdir = "/" ,title = "Select file" ,\
                    defaultextension='.csv', filetypes = (("CSV file",".csv"),("Excel file",".xlsx")))
        if len(filename) >0:
            if filename.endswith('.csv'):
                Export_MasterTB_DF.to_csv(filename,index=None)
                tkinter.messagebox.showinfo("Master TimeBreak Database Export","Master Timebreak DB Saved as CSV")
            else:
                Export_MasterTB_DF.to_excel(filename, sheet_name='MasterTB', index=False)
                tkinter.messagebox.showinfo("Master TimeBreak Database Export","Master Timebreak DB Saved as Excel")
        else:
            tkinter.messagebox.showinfo("Master TimeBreak Database Export Message","Please Select File Name To Export")
                    
        conn.commit()
        conn.close()

    def ExportListBoxTB():
        dfList =[] 
        for child in tree.get_children():
            df = tree.item(child)["values"]
            dfList.append(df)
        ListBox_DF = pd.DataFrame(dfList)
        ListBox_DF.rename(columns = {0:'TriggerIndex', 1:' ProfileId', 2:' ShotNumber', 3:' EpNumber', 4:' ShotLine',
                                      5: ' ShotStation', 6:' ShotUtcDateTime', 7:' Latitude', 8:' Longitude', 9:' ShotStatus', 10:' Uphole', 11:' Comment', 12:' Process'},inplace = True)
                    
        Export_ListBox  = pd.DataFrame(ListBox_DF)
        Export_ListBox  = Export_ListBox.reset_index(drop=True)    
        filename = tkinter.filedialog.asksaveasfilename(initialdir = "/" ,title = "Select File Name to Export" ,\
                       defaultextension='.csv', filetypes = (("CSV file",".csv"),("Excel file",".xlsx")))
        if len(filename) >0:
            if filename.endswith('.csv'):
                Export_ListBox.to_csv(filename,index=None)
                tkinter.messagebox.showinfo("ListBox TimeBreak Export","ListBox Timebreak Entries Saved as CSV")
            else:
                Export_ListBox.to_excel(filename, sheet_name='ListBoxTB', index=False)
                tkinter.messagebox.showinfo("ListBox TimeBreak Export","ListBox Timebreak Entries Saved as Excel")
        else:
            tkinter.messagebox.showinfo("ListBox TimeBreak Export Message","Please Select File Name To Export") 

    def SortListBoxView():
        try:
            dfList =[] 
            for child in tree.get_children():
                df = tree.item(child)["values"]
                dfList.append(df)
            ListBox_DF = pd.DataFrame(dfList)
            ListBox_DF.rename(columns = {0:'TriggerIndex', 1:'Unit_ID', 2:'FileNum', 3:'EPNumber', 4:'SourceLine',
                                          5: 'SourceStation', 6:'ShotUtcDateTime', 7:'Latitude', 8:'Longitude',9:'ShotStatus', 10:' Uphole', 11:'TBComment', 12:' Process'},inplace = True)                        
            Sort_ListBox  = pd.DataFrame(ListBox_DF)
            Sort_ListBox  = Sort_ListBox.sort_values(by =['SourceLine','SourceStation'])
            Sort_ListBox  = Sort_ListBox.reset_index(drop=True)
            tree.delete(*tree.get_children())
            for each_rec in range(len(Sort_ListBox)):
                tree.insert("", tk.END, values=list(Sort_ListBox.loc[each_rec]))
        except:
            tkinter.messagebox.showinfo("ListBox Sort Message","List Box Is Empty No Data to Sort") 
                    
    def ClearView():
        txtTotalEntries.delete(0,END)
        txtDuplicatedEntries.delete(0,END)
        tree.delete(*tree.get_children())
        ClearEditEntry()

    def TotalEntries():
        txtTotalEntries.delete(0,END)
        conn = sqlite3.connect("SourceLink_Dynamite_Log.db")
        Complete_df_TB = pd.read_sql_query("SELECT * FROM Eagle_SOURCELINK_DYNAMITE_TBMASTER ORDER BY `FileNum` ASC;", conn)
        data = pd.DataFrame(Complete_df_TB)
        data = data.reset_index(drop=True)
        TotalEntries = len(data)       
        txtTotalEntries.insert(tk.END,TotalEntries)              
        conn.commit()
        conn.close()

    def ViewDuplicatedEntries():
        ClearView()
        conn = sqlite3.connect("SourceLink_Dynamite_Log.db")
        cur = conn.cursor()  
        Complete_df_TB = pd.read_sql_query("SELECT * FROM Eagle_SOURCELINK_DYNAMITE_TBMASTER ORDER BY `FileNum` ASC;", conn)    
        data = pd.DataFrame(Complete_df_TB)
        data = data.drop_duplicates(['FileNum'],keep='last')
        data = data.reset_index(drop=True)
        data ['DuplicatedShot']=data.sort_values(by =['SourceLine','SourceStation']).duplicated(['SourceLine','SourceStation'],keep=False)
        data = data.loc[data.DuplicatedShot == True]
        data = data.reset_index(drop=True)
        List_DuplicatedFFID = (list(data['FileNum']))
        for i in range(len(List_DuplicatedFFID)):
            cur.execute("SELECT * FROM Eagle_SOURCELINK_DYNAMITE_TBMASTER WHERE FileNum =?", (List_DuplicatedFFID[i],))
            rows=cur.fetchall()        
            for each_rec in rows:
                tree.insert("", tk.END, values=each_rec)    
        conn.commit()
        conn.close()
        DuplicatedEntries = len(tree.get_children())       
        txtDuplicatedEntries.insert(tk.END,DuplicatedEntries)

      
    def DeleteSelectedImportData():
        iDelete = tkinter.messagebox.askyesno("Delete Entry", "Confirm if you want to Delete")
        if iDelete >0:
            txtTotalEntries.delete(0,END)
            txtDuplicatedEntries.delete(0,END)
            conn = sqlite3.connect("SourceLink_Dynamite_Log.db")
            cur = conn.cursor()                
            for selected_item in tree.selection():
                cur.execute("DELETE FROM Eagle_SOURCELINK_DYNAMITE_TBMASTER WHERE TriggerIndex =? AND FileNum =? AND EPNumber =? AND SourceLine =? AND \
                            SourceStation =? ",\
                            (tree.set(selected_item, '#1'), tree.set(selected_item, '#3'),tree.set(selected_item, '#4'),
                             tree.set(selected_item, '#5'), tree.set(selected_item, '#6'),)) 
                tree.delete(selected_item)
                conn.commit()
            conn.commit()
            conn.close()
            TotalEntries()
            return
        

    def ImportTBFile():
        tree.delete(*tree.get_children())
        txtNewEntries.delete(0,END)
        txtTotalEntries.delete(0,END)
        txtDuplicatedEntries.delete(0,END)
        fileList = askopenfilenames(initialdir = "/", title = "Import Dynamite SourceLink Time Break Files" , filetypes=[('CSV File', '*.csv'), ('Excel File', ('*.xls', '*.xlsx'))])
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
                concatdf_TB                     = concatdf_TB[pd.to_numeric(concatdf_TB.FileNum,errors='coerce').notnull()]                
                concatdf_TB['SourceLine']       = (concatdf_TB.loc[:,['SourceLine']]).astype(int)
                concatdf_TB['SourceStation']    = (concatdf_TB.loc[:,['SourceStation']]).astype(float)
                concatdf_TB['FileNum']          = (concatdf_TB.loc[:,['FileNum']]).astype(int)            
                concatdf_TB['EPNumber']         = (concatdf_TB.loc[:,['EPNumber']]).astype(int)
                concatdf_TB['Unit_ID']          = (concatdf_TB.loc[:,['Unit_ID']]).astype(int)
                concatdf_TB['TriggerIndex']     = (concatdf_TB.loc[:,['TriggerIndex']]).astype(int)
                concatdf_TB                     = concatdf_TB.reset_index(drop=True)
                DATA_VALID_TB                   = pd.DataFrame(concatdf_TB)
                Len_NewTB_Entries               = len(DATA_VALID_TB)

                ## Connecting to SQL DB        
                con= sqlite3.connect("SourceLink_Dynamite_Log.db")
                cur=con.cursor()
                txtNewEntries.insert(tk.END,Len_NewTB_Entries)
                DATA_VALID_TB.to_sql('Eagle_SOURCELINK_DYNAMITE_TBMASTER',con, if_exists="append", index=False)
                for each_rec in range(len(DATA_VALID_TB)):
                    tree.insert("", tk.END, values=list(DATA_VALID_TB.loc[each_rec]))
                con.commit()
                con.close()
                UpdateMasterTB()        
        else:
            tkinter.messagebox.showinfo("Import Dynamite TB File Message","Please Select Clean Dynamite TimeBreak Files To Import")
            

    ##### Entry Wizard
    txtDuplicatedEntries  = Entry(window, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
    txtDuplicatedEntries.place(x=425,y=6)
    txtTotalEntries  = Entry(window, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 10)
    txtTotalEntries.place(x=1150,y=6)

    ## Entry For Edit
    L3 = Label(TitFrame, text = "Edit Selected Master Database Dynamite Timebreak Entry : ", font=("arial", 10,'bold'),bg = "orange")
    L3.pack(side = TOP)

    LabelTriggerIndex = Label(EditTableMargin, text = "Index", font=("arial", 10,'bold'))
    LabelTriggerIndex.grid(row =2, column = 0 , sticky = W, padx= 0)
    txtTriggerIndex  = Entry(EditTableMargin, font=('aerial', 10, 'bold'),textvariable=StringVar(), width = 10)
    txtTriggerIndex.grid(row =4, column = 0 , sticky = W, padx= 0)

    LabelProfileID = Label(EditTableMargin, text = " Profile", font=("arial", 10,'bold'))
    LabelProfileID.grid(row =2, column = 1 , sticky = W, padx= 0)
    txtProfileID  = Entry(EditTableMargin, font=('aerial', 10, 'bold'),textvariable=StringVar(), width = 10)
    txtProfileID.grid(row =4, column = 1 , sticky = W, padx= 0)

    LabelShotNumber = Label(EditTableMargin, text = "ShotNumber", font=("arial", 10,'bold'))
    LabelShotNumber.grid(row =2, column = 2 , sticky = W, padx= 0)
    txtShotNumber  = Entry(EditTableMargin, font=('aerial', 10, 'bold'),textvariable=StringVar(), width = 15)
    txtShotNumber.grid(row =4, column = 2 , sticky = W, padx= 0)


    LabelEpNumber = Label(EditTableMargin, text = "EP", font=("arial", 10,'bold'))
    LabelEpNumber.grid(row =2, column = 3 , sticky = W, padx= 0)
    txtEpNumber  = Entry(EditTableMargin, font=('aerial', 10, 'bold'),textvariable=StringVar(), width = 8)
    txtEpNumber.grid(row =4, column = 3 , sticky = W, padx= 0)


    LabelShotLine = Label(EditTableMargin, text = "Line", font=("arial", 10,'bold'))
    LabelShotLine.grid(row =2, column = 4 , sticky = W, padx= 0)
    txtShotLine  = Entry(EditTableMargin, font=('aerial', 10, 'bold'),textvariable=StringVar(), width = 12)
    txtShotLine.grid(row =4, column = 4 , sticky = W, padx= 0)


    LabelShotStation = Label(EditTableMargin, text = "Station", font=("arial", 10,'bold'))
    LabelShotStation.grid(row =2, column = 5 , sticky = W, padx= 0)
    txtShotStation  = Entry(EditTableMargin, font=('aerial', 10, 'bold'),textvariable=StringVar(), width = 12)
    txtShotStation.grid(row =4, column = 5 , sticky = W, padx= 0)

    LabelShotUtcDateTime = Label(EditTableMargin, text = "UtcDateTime", font=("arial", 10,'bold'))
    LabelShotUtcDateTime.grid(row =2, column = 6 , sticky = W, padx= 0)
    txtShotUtcDateTime  = Entry(EditTableMargin, font=('aerial', 10, 'bold'),textvariable=StringVar(), width = 25)
    txtShotUtcDateTime.grid(row =4, column = 6 , sticky = W, padx= 0)


    LabelLatitude = Label(EditTableMargin, text = "Lat", font=("arial", 10,'bold'))
    LabelLatitude.grid(row =2, column = 7 , sticky = W, padx= 0)
    txtLatitude  = Entry(EditTableMargin, font=('aerial', 10, 'bold'),textvariable=StringVar(), width = 13)
    txtLatitude.grid(row =4, column = 7 , sticky = W, padx= 0)

    LabelLongitude = Label(EditTableMargin, text = "Long", font=("arial", 10,'bold'))
    LabelLongitude.grid(row =2, column = 8 , sticky = W, padx= 0)
    txtLongitude  = Entry(EditTableMargin, font=('aerial', 10, 'bold'),textvariable=StringVar(), width = 13)
    txtLongitude.grid(row =4, column = 8 , sticky = W, padx= 0)

    LabelShotStatus = Label(EditTableMargin, text = "Status", font=("arial", 10,'bold'))
    LabelShotStatus.grid(row =2, column = 9 , sticky = W, padx= 0)
    txtShotStatus  = Entry(EditTableMargin, font=('aerial', 10, 'bold'),textvariable=StringVar(), width = 14)
    txtShotStatus.grid(row =4, column = 9 , sticky = W, padx= 0)

    LabelUphole = Label(EditTableMargin, text = "Uphole", font=("arial", 10,'bold'))
    LabelUphole.grid(row =2, column = 10 , sticky = W, padx= 0)
    txtUphole  = Entry(EditTableMargin, font=('aerial', 10, 'bold'),textvariable=StringVar(), width = 12)
    txtUphole.grid(row =4, column = 10 , sticky = W, padx= 0)

    LabelComment = Label(EditTableMargin, text = "Comment", font=("arial", 10,'bold'))
    LabelComment.grid(row =2, column = 11 , sticky = W, padx= 0)
    txtComment  = Entry(EditTableMargin, font=('aerial', 10, 'bold'),textvariable=StringVar(), width = 15)
    txtComment.grid(row =4, column = 11 , sticky = W, padx= 0)


    LabelProcess = Label(EditTableMargin, text = "Process", font=("arial", 10,'bold'))
    LabelProcess.grid(row =2, column = 12 , sticky = W, padx= 0)
    txtProcess  = Entry(EditTableMargin, font=('aerial', 10, 'bold'),textvariable=StringVar(), width = 12)
    txtProcess.grid(row =4, column = 12 , sticky = W, padx= 0)

    #----------------- Tree View Select Event------------
            
    tree.bind('<<TreeviewSelect>>',DynamiteTB_MasterRec_Event)


    ##### Labeling
    L1 = Label(window, text = " MasterDB Dynamite Timebreak Details:", font=("arial", 10,'bold'),bg = "green").place(x=2,y=6)
    L2 = Label(window, text = "Duplicated Dynamite Shot", font=("arial", 10,'bold'),bg = "red").place(x=505,y=7)

    ### Top Button Wizard
    btnViewDuplicatedShot= Button(window, text="View Duplicated Shot", font=('aerial', 9, 'bold'), height =1, width=18, bd=1, command = ViewDuplicatedEntries)
    btnViewDuplicatedShot.place(x=290,y=6)
    btnViewSort= Button(window, text="Sort View", font=('aerial', 9, 'bold'), height =1, width=8, bd=1, command = SortListBoxView)
    btnViewSort.place(x=680,y=6)
    btnViewMasterDB = Button(window, text="View Master DB", font=('aerial', 9, 'bold'), height =1, width=15, bd=1, command = ViewMasterTB)
    btnViewMasterDB.place(x=1032,y=6)
    btnExportMasterDB = Button(window, text="Export Master DB", font=('aerial', 9, 'bold'), height =1, width=15, bd=1, command = ExportMasterTB_DB)
    btnExportMasterDB.place(x=900,y=6)

    ## Bottom Buttons
    btnImportCleanTB = Button(TableMargin, text="Import Clean Dynamite TB File", font=('aerial', 9, 'bold'), height =1, width=25, bd=2, command = ImportTBFile)
    btnImportCleanTB.place(x=0,y=607)
    btnUpdateMasterDb = Button(TableMargin, text="Update Dynamite MasterDB", font=('aerial', 9, 'bold'), height =1, width=23, bd=2, command = UpdateMasterTB)
    btnUpdateMasterDb.place(x=190,y=607)
    L3 = Label(TableMargin, text = "New TB Entries :", font=("arial", 10,'bold'),bg = "cadet blue").place(x=500,y=607)
    txtNewEntries  = Entry(TableMargin, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
    txtNewEntries.place(x=612,y=607)
    btnExportListBoxView = Button(TableMargin, text="Export List Box", font=('aerial', 9, 'bold'), height =1, width=13, bd=2, command = ExportListBoxTB)
    btnExportListBoxView.place(x=877,y=607)
    btnDelete = Button(TableMargin, text="Delete Selected", font=('aerial', 9, 'bold'), height =1, width=13, bd=2, command = DeleteSelectedImportData)
    btnDelete.place(x=980,y=607)
    btnClearView = Button(TableMargin, text="Clear View", font=('aerial', 9, 'bold'), height =1, width=10, bd=2, command = ClearView)
    btnClearView.place(x=1083,y=607)
    btnExit = Button(TableMargin, text="Exit Widget", font=('aerial', 9, 'bold'), height =1, width=10, bd=2, command = iExit)
    btnExit.place(x=1165,y=607)

    ## Bottom Last Buttons
    btnEditMasterDBEntry = Button(window, text="Edit MasterDB Entry", font=('aerial', 10, 'bold'), height =1, width=16, bd=2, command = update)
    btnEditMasterDBEntry.place(x=1112,y=753)












