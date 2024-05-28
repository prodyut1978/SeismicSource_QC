#Front End
import os
from tkinter import*
import tkinter.messagebox
import GeomergeAUXTB_BackEnd
import SetupVibAuxProfile
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

def GenerateGeomergeTBVibSignature():
    Default_Date_today   = datetime.date.today()
    window = Tk()
    window.title ("Eagle SourceLink Timebreak AUX TimeBreak Wizard")
    window.geometry("1350x655+10+0")
    window.config(bg="cadet blue")
    window.resizable(0, 0)
    TableMargin = Frame(window, bd = 2, padx= 2, pady= 2, relief = RIDGE)
    TableMargin1 = Frame(window, bd = 2, padx= 2, pady= 2, relief = RIDGE)
    TableMargin.pack(side=TOP)
    TableMargin.pack(side=LEFT)
    TableMargin1.pack(side=TOP)
    TableMargin1.pack(side=LEFT)
    scrollbarx = Scrollbar(TableMargin, orient=HORIZONTAL)
    scrollbary = Scrollbar(TableMargin, orient=VERTICAL)
    scrollbarx1 = Scrollbar(TableMargin1, orient=HORIZONTAL)
    scrollbary1 = Scrollbar(TableMargin1, orient=VERTICAL)
    tree = ttk.Treeview(TableMargin, column=("column1", "column2", "column3", "column4", "column5",
                                             "column6", "column7", "column8", "column9", "column10", "column11" , "column12", "column13"), height=27, show='headings')

    tree1 = ttk.Treeview(TableMargin1, column=("column1", "column2", "column3"), height=12, show='headings')
    scrollbary.config(command=tree.yview)
    scrollbary.pack(side=RIGHT, fill=Y)
    scrollbarx.config(command=tree.xview)
    scrollbarx.pack(side=BOTTOM, fill=X)
    scrollbary1.config(command=tree1.yview)
    scrollbary1.pack(side=RIGHT, fill=Y)
    scrollbarx1.config(command=tree1.xview)
    scrollbarx1.pack(side=BOTTOM, fill=X)
    tree.heading("#1", text="TriggerIndex", anchor=W)
    tree.heading("#2", text="ProfileId", anchor=W)
    tree.heading("#3", text="ShotNumber", anchor=W)
    tree.heading("#4", text="EpNumber", anchor=W)
    tree.heading("#5", text="ShotLine", anchor=W)        
    tree.heading("#6", text="ShotStation", anchor=W)
    tree.heading("#7", text="ShotUtcDateTime", anchor=W)
    tree.heading("#8", text="Latitude", anchor=W)        
    tree.heading("#9", text="Longitude" ,anchor=W)
    tree.heading("#10", text="ShotStatus", anchor=W)
    tree.heading("#11", text="Comments", anchor=W)
    tree.heading("#12", text="AUXUnitNumber", anchor=W)
    tree.heading("#13", text="DeviceType", anchor=W)

    tree1.heading("#1", text="ProfileId", anchor=W)
    tree1.heading("#2", text="AUXUnitNumber", anchor=W)
    tree1.heading("#3", text="DeviceType", anchor=W)
    tree.column('#1', stretch=NO, minwidth=0, width=85)            
    tree.column('#2', stretch=NO, minwidth=0, width=70)
    tree.column('#3', stretch=NO, minwidth=0, width=80)
    tree.column('#4', stretch=NO, minwidth=0, width=80)
    tree.column('#5', stretch=NO, minwidth=0, width=75)
    tree.column('#6', stretch=NO, minwidth=0, width=80)
    tree.column('#7', stretch=NO, minwidth=0, width=105)
    tree.column('#8', stretch=NO, minwidth=0, width=65)
    tree.column('#9', stretch=NO, minwidth=0, width=65)
    tree.column('#10', stretch=NO, minwidth=0, width=80)
    tree.column('#11', stretch=NO, minwidth=0, width=80)
    tree.column('#12', stretch=NO, minwidth=0, width=100)
    tree.column('#13', stretch=NO, minwidth=0, width=80)
    tree1.column('#1', stretch=NO, minwidth=0, width=70)            
    tree1.column('#2', stretch=NO, minwidth=0, width=100)
    tree1.column('#3', stretch=NO, minwidth=0, width=80)
    style = ttk.Style()
    style.theme_use("clam")
    style.configure(".", font=('aerial', 8), foreground="black")
    style.configure("Treeview", foreground='black')
    style.configure("Treeview.Heading",font=('aerial', 8,'bold'), background='Ghost White', foreground='blue',fieldbackground='Ghost White')
    tree.pack()
    tree1.pack()

    ProfileId      = IntVar()
    AUXUnitNumber  = IntVar()
    DeviceType     = IntVar(window, value=257)

    lblProfileID = Label(window, font=('aerial', 10, 'bold'), text = "1. Vib Profile ID :", padx =2, pady= 2, bg = "cadet blue")
    lblProfileID.place(x=1077,y=50)
    txtProfileID  = Entry(window, font=('aerial', 12, 'bold'),textvariable= ProfileId, width = 6)
    txtProfileID.place(x=1230,y=50)

    lblAUXUnitNumber = Label(window, font=('aerial', 10, 'bold'), text = "2. AUX UnitNumber :", padx =2, pady= 2, bg = "cadet blue")
    lblAUXUnitNumber.place(x=1077,y=80)
    txtAUXUnitNumber  = Entry(window, font=('aerial', 12, 'bold'),textvariable = AUXUnitNumber, width = 6)
    txtAUXUnitNumber.place(x=1230,y=80)

    lblDeviceType = Label(window, font=('aerial', 10, 'bold'), text = "3. AUX DeviceType :", padx =2, pady= 2, bg = "cadet blue")
    lblDeviceType.place(x=1077,y=110)
    txtDeviceType  = Entry(window, font=('aerial', 12, 'bold'),textvariable = DeviceType, width = 6)
    txtDeviceType.place(x=1230,y=110)


    # All Functions defining 

    def iExit():
        iExit= tkinter.messagebox.askyesno("Eagle Aux TB Import Wizard", "Confirm if you want to exit")
        if iExit >0:
            window.destroy()
            return

    def AddData():
        if((txtProfileID.get())!=0) & ((txtAUXUnitNumber.get())!=0) & ((txtDeviceType.get())!=0):
            GeomergeAUXTB_BackEnd.addInvRec(txtProfileID.get(), txtAUXUnitNumber.get(), txtDeviceType.get())
            tree1.delete(*tree1.get_children())
            tree1.insert("", tk.END,values=(txtProfileID.get(), txtAUXUnitNumber.get(), txtDeviceType.get()))
        else:
            tkinter.messagebox.showinfo("Add Error","Entries can not be empty")

    def update():
        if((txtProfileID.get())!=0) & ((txtAUXUnitNumber.get())!=0) & ((txtDeviceType.get())!=0):
            conn = sqlite3.connect("GeomergeAUXTB.db")
            cur = conn.cursor()
            for selected_item in tree1.selection():
                cur.execute("DELETE FROM AUXBOX_Profile WHERE ProfileId =?", (tree1.set(selected_item, '#1'),))
                conn.commit()
                tree1.delete(selected_item)
                conn.close()

        if((txtProfileID.get())!=0) & ((txtAUXUnitNumber.get())!=0) & ((txtDeviceType.get())!=0):
            GeomergeAUXTB_BackEnd.addInvRec(txtProfileID.get(), txtAUXUnitNumber.get(), txtDeviceType.get())
            tree1.delete(*tree1.get_children())
            tree1.insert("", tk.END,values=(txtProfileID.get(), txtAUXUnitNumber.get(), txtDeviceType.get()))
        else:
            tkinter.messagebox.showinfo("Add Error","Entries can not be empty")


    def DeleteSelected():
        if((txtProfileID.get())!=0) & ((txtAUXUnitNumber.get())!=0) & ((txtDeviceType.get())!=0):
            conn = sqlite3.connect("GeomergeAUXTB.db")
            cur = conn.cursor()
            for selected_item in tree1.selection():
                cur.execute("DELETE FROM AUXBOX_Profile WHERE ProfileId =?", (tree1.set(selected_item, '#1'),))
                conn.commit()
                tree1.delete(selected_item)
            conn.commit()
            conn.close()
            
            
    def ViewAuxTBProfile():
        tree1.delete(*tree1.get_children())
        conn = sqlite3.connect("GeomergeAUXTB.db")
        Complete_df = pd.read_sql_query("SELECT * FROM AUXBOX_Profile ORDER BY `ProfileId` ASC;", conn)
        data = pd.DataFrame(Complete_df)
        data = data.reset_index(drop=True)
        for each_rec in range(len(data)):
            tree1.insert("", tk.END, values=list(data.loc[each_rec]))        
        conn.commit()
        conn.close()

    def SetupAuxTBProfile():
        tree1.delete(*tree1.get_children())
        txtProfileID.delete(0,END)
        txtAUXUnitNumber.delete(0,END)
        txtDeviceType.delete(0,END)            
        SetupVibAuxProfile.VibAuxProfileSetup()
        ViewAuxTBProfile()
        
        
    def ExportAUXTB_Report():
        conn = sqlite3.connect("GeomergeAUXTB.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Geomerge_AUXTB_MERGED ORDER BY `ShotNumber` ASC ;", conn)
        Export_MasterTB_DF  = pd.DataFrame(Complete_df)
        Export_MasterTB_DF.rename(columns={'TriggerIndex':'TriggerIndex', 'ProfileId':' ProfileId', 'ShotNumber':' ShotNumber',
                                                                'EpNumber':' EpNumber', 'ShotLine':' ShotLine', 'ShotStation':' ShotStation',
                                                                'ShotUtcDateTime':' ShotUtcDateTime','Latitude':' Latitude','Longitude':' Longitude',
                                                                'ShotStatus':' ShotStatus', 'Comment':' Comment', 'DeviceType':'DeviceType', 'AUXUnitNumber':'AUXUnitNumber'},inplace = True)

        Export_MasterTB_DF = Export_MasterTB_DF.loc[:,['DeviceType','AUXUnitNumber','TriggerIndex',' ProfileId',' ShotNumber',
                          ' EpNumber',' ShotLine',' ShotStation',' ShotUtcDateTime',                                  
                          ' Latitude',' Longitude',' ShotStatus',' Comment']]
        Export_MasterTB_DF  = Export_MasterTB_DF.reset_index(drop=True)
        filename = tkinter.filedialog.asksaveasfilename(initialdir = "/" ,title = "Select SourceLink TimeBreak File" ,\
                       defaultextension='.csv', filetypes = (("CSV file",".csv"),("Excel file",".xlsx")))
        if len(filename) >0:            
            if filename.endswith('.csv'):
                Export_MasterTB_DF.to_csv(filename,index=None)
                tkinter.messagebox.showinfo("SourceLink AUX TB Export","SourceLink AUX Timebreak Report Saved as CSV")
            else:
                Export_MasterTB_DF.to_excel(filename, sheet_name='MasterTB', index=False)
                tkinter.messagebox.showinfo("SourceLink AUX TB Export","SourceLink AUX Timebreak Report Saved as Excel")
        else:
            tkinter.messagebox.showinfo("SourceLink AUX TB Export Export Message","Please Select File Name To Export")
                    
        conn.commit()
        conn.close()


    def SortByShotNumber():
        try:
            dfList =[] 
            for child in tree.get_children():
                df = tree.item(child)["values"]
                dfList.append(df)
            ListBox_DF = pd.DataFrame(dfList)
            ListBox_DF.rename(columns={0:'TriggerIndex', 1:'ProfileId', 2:'ShotNumber', 3:'EpNumber', 4:'ShotLine',
                                         5:'ShotStation', 6:'ShotUtcDateTime', 7:'Latitude', 8:'Longitude',
                                         9:'ShotStatus', 10:'Comment' },inplace = True)                      
            Sort_ListBox  = pd.DataFrame(ListBox_DF)
            Sort_ListBox  = Sort_ListBox.sort_values(by =['ShotNumber'])
            Sort_ListBox  = Sort_ListBox.reset_index(drop=True)
            tree.delete(*tree.get_children())
            for each_rec in range(len(Sort_ListBox)):
                tree.insert("", tk.END, values=list(Sort_ListBox.loc[each_rec]))
        except:
            tkinter.messagebox.showinfo("ListBox Sort Message","List Box Is Empty No Data to Sort")

    def SortByProfileID():
        try:
            dfList =[] 
            for child in tree.get_children():
                df = tree.item(child)["values"]
                dfList.append(df)
            ListBox_DF = pd.DataFrame(dfList)
            ListBox_DF.rename(columns={0:'TriggerIndex', 1:'ProfileId', 2:'ShotNumber', 3:'EpNumber', 4:'ShotLine',
                                         5:'ShotStation', 6:'ShotUtcDateTime', 7:'Latitude', 8:'Longitude',
                                         9:'ShotStatus', 10:'Comment' },inplace = True)                      
            Sort_ListBox  = pd.DataFrame(ListBox_DF)
            Sort_ListBox  = Sort_ListBox.sort_values(by =['ProfileId'])
            Sort_ListBox  = Sort_ListBox.reset_index(drop=True)
            tree.delete(*tree.get_children())
            for each_rec in range(len(Sort_ListBox)):
                tree.insert("", tk.END, values=list(Sort_ListBox.loc[each_rec]))
        except:
            tkinter.messagebox.showinfo("ListBox Sort Message","List Box Is Empty No Data to Sort") 
                    
    def ClearView():
        txtTotalEntries.delete(0,END)
        txtTotalGeneratedEntries.delete(0,END)
        txtProfileID.delete(0,END)
        txtAUXUnitNumber.delete(0,END)
        txtDeviceType.delete(0,END)
        tree.delete(*tree.get_children())
        tree1.delete(*tree1.get_children())

    def AuxCreateClearEntry():
        txtProfileID.delete(0,END)
        txtAUXUnitNumber.delete(0,END)
        txtDeviceType.delete(0,END)


    def GenerateVibAUX_TB():
        tree.delete(*tree.get_children())
        txtTotalGeneratedEntries.delete(0,END)
        conn = sqlite3.connect("GeomergeAUXTB.db")
        Complete_df_TB = pd.read_sql_query("SELECT * FROM Geomerge_AUXTB_TEMP ORDER BY `ShotNumber` ASC;", conn)    
        Complete_df_TB = pd.DataFrame(Complete_df_TB)
        Complete_df_TB = Complete_df_TB.reset_index(drop=True)

        df_AUXProfile  = pd.read_sql_query("SELECT * FROM AUXBOX_Profile ORDER BY `ProfileId` ASC;", conn)
        df_AUXProfile  = pd.DataFrame(df_AUXProfile)
        df_AUXProfile  = df_AUXProfile.reset_index(drop=True)

        GeneratedAUXTB_Report = pd.merge(Complete_df_TB, df_AUXProfile, on ='ProfileId',how ='inner', sort = True)
        GeneratedAUXTB_Report = pd.DataFrame(GeneratedAUXTB_Report)
        GeneratedAUXTB_Report = GeneratedAUXTB_Report.reset_index(drop=True)
        GeneratedAUXTB_Report.to_sql('Geomerge_AUXTB_MERGED',conn, if_exists="replace", index=False)
        for each_rec in range(len(GeneratedAUXTB_Report)):
            tree.insert("", tk.END, values=list(GeneratedAUXTB_Report.loc[each_rec]))  
        conn.commit()
        conn.close()
        TotalGeneratedEntries = len(GeneratedAUXTB_Report)       
        txtTotalGeneratedEntries.insert(tk.END,TotalGeneratedEntries)


    def InventoryRec(event):
        for nm in tree1.selection():
            sd = tree1.item(nm, 'values')
            txtProfileID.delete(0,END)
            txtProfileID.insert(tk.END,sd[0])                
            txtAUXUnitNumber.delete(0,END)
            txtAUXUnitNumber.insert(tk.END,sd[1])
            txtDeviceType.delete(0,END)
            txtDeviceType.insert(tk.END,sd[2])

    def ImportTBFile():
        tree.delete(*tree.get_children())
        txtTotalEntries.delete(0,END)
        txtTotalGeneratedEntries.delete(0,END)
        fileList = askopenfilenames(initialdir = "/", title = "Import SourceLink TimeBreak Files" , filetypes=[('CSV File', '*.csv'), ('Excel File', ('*.xls', '*.xlsx'))])
        Length_fileList  =  len(fileList)
        if Length_fileList >0:            
            if fileList:
                df_TBList = []           
                for filename in fileList:
                    if filename.endswith('.csv'):
                        df_TB             = pd.read_csv(filename, sep=',' , low_memory=False)
                        df_TB             = df_TB.iloc[:,:]
                        TriggerIndex      = df_TB.loc[:,'TriggerIndex']
                        ProfileId         = df_TB.loc[:,' ProfileId']
                        ShotNumber        = df_TB.loc[:,' ShotNumber']
                        EpNumber          = df_TB.loc[:,' EpNumber']
                        ShotLine          = df_TB.loc[:,' ShotLine']
                        ShotStation       = df_TB.loc[:,' ShotStation']
                        ShotUtcDateTime   = df_TB.loc[:,' ShotUtcDateTime']
                        Latitude          = df_TB.loc[:,' Latitude']
                        Longitude         = df_TB.loc[:,' Longitude']
                        ShotStatus        = df_TB.loc[:,' ShotStatus']
                        Comment           = df_TB.loc[:,' Comment']                    
                        column_names = [TriggerIndex, ProfileId, ShotNumber, EpNumber, ShotLine, ShotStation, ShotUtcDateTime,
                                        Latitude, Longitude, ShotStatus, Comment]
                        catdf_TB = pd.concat (column_names,axis=1,ignore_index =True)
                        df_TBList.append(catdf_TB) 
                    else:
                        df_TB             = pd.read_excel(filename)
                        df_TB             = df_TB.iloc[:,:]
                        TriggerIndex      = df_TB.loc[:,'TriggerIndex']
                        ProfileId         = df_TB.loc[:,' ProfileId']
                        ShotNumber        = df_TB.loc[:,' ShotNumber']
                        EpNumber          = df_TB.loc[:,' EpNumber']
                        ShotLine          = df_TB.loc[:,' ShotLine']
                        ShotStation       = df_TB.loc[:,' ShotStation']
                        ShotUtcDateTime   = df_TB.loc[:,' ShotUtcDateTime']
                        Latitude          = df_TB.loc[:,' Latitude']
                        Longitude         = df_TB.loc[:,' Longitude']
                        ShotStatus        = df_TB.loc[:,' ShotStatus']
                        Comment           = df_TB.loc[:,' Comment']                    
                        column_names = [TriggerIndex, ProfileId, ShotNumber, EpNumber, ShotLine, ShotStation, ShotUtcDateTime,
                                        Latitude, Longitude, ShotStatus, Comment]
                        catdf_TB = pd.concat (column_names,axis=1,ignore_index =True)
                        df_TBList.append(catdf_TB) 

                concatdf_TB = pd.concat(df_TBList,axis=0, ignore_index =True)
                concatdf_TB.rename(columns={0:'TriggerIndex', 1:'ProfileId', 2:'ShotNumber', 3:'EpNumber', 4:'ShotLine',
                                         5:'ShotStation', 6:'ShotUtcDateTime', 7:'Latitude', 8:'Longitude',
                                         9:'ShotStatus', 10:'Comment' },inplace = True)
                concatdf_TB= concatdf_TB[pd.to_numeric(concatdf_TB.ShotNumber,errors='coerce').notnull()]
                
                concatdf_TB['ShotLine']       = (concatdf_TB.loc[:,['ShotLine']]).astype(int)
                concatdf_TB['ShotStation']    = (concatdf_TB.loc[:,['ShotStation']]).astype(float)
                concatdf_TB['ShotNumber']     = (concatdf_TB.loc[:,['ShotNumber']]).astype(int)            
                concatdf_TB['EpNumber']       = (concatdf_TB.loc[:,['EpNumber']]).astype(int)
                concatdf_TB['ProfileId']      = (concatdf_TB.loc[:,['ProfileId']]).astype(int)
                concatdf_TB['TriggerIndex']   = (concatdf_TB.loc[:,['TriggerIndex']]).astype(int)
                concatdf_TB                   = concatdf_TB.reset_index(drop=True)
                DATA_VALID_TB                 = pd.DataFrame(concatdf_TB)
                TotalEntries                  = len(DATA_VALID_TB)

                ## Connecting to SQL DB        
                con= sqlite3.connect("GeomergeAUXTB.db")
                cur=con.cursor()                
                DATA_VALID_TB.to_sql('Geomerge_AUXTB_TEMP',con, if_exists="replace", index=False)
                for each_rec in range(len(DATA_VALID_TB)):
                    tree.insert("", tk.END, values=list(DATA_VALID_TB.loc[each_rec]))
                con.commit()
                con.close()
                txtTotalEntries.insert(tk.END,TotalEntries)
                        
        else:
            tkinter.messagebox.showinfo("Import TB File Error Message","Please Check TimeBreak Imported File")
            

    ##### Entry Wizard

    tree1.bind('<<TreeviewSelect>>',InventoryRec)


    ##### Labeling
    L1 = Label(window, text = "Timebreak Details :", font=("arial", 10,'bold'),bg = "cadet blue").place(x=2,y=6)

    L2 = Label(window, text = "Total Imported Entries :", font=("arial", 12,'bold'),bg = "cadet blue").place(x=820,y=4)
    txtTotalEntries  = Entry(window, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 7)
    txtTotalEntries.place(x=1008,y=4)

    L5 = Label(window, text = "Total Generated Entries :", font=("arial", 12,'bold'),bg = "cadet blue").place(x=500,y=4)
    txtTotalGeneratedEntries  = Entry(window, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 7)
    txtTotalGeneratedEntries.place(x=700,y=4)

    L3 = Label(window, text = "Vib Aux Unit Profile :", font=("arial", 10,'bold'),bg = "cadet blue").place(x=1077,y=155)
    L4 = Label(window, text = "Vib Profile Info :", font=("arial", 10,'bold'),bg = "green").place(x=1077,y=25)

    ### Button Wizard

    btnImportTB = Button(window, text="Import Clean TB File", font=('aerial', 9, 'bold'), height =1, width=17, bd=4, command = ImportTBFile)
    btnImportTB.place(x=2,y=625)

    btnGenerateAUXTB = Button(window, text="Generate AUX TB", font=('aerial', 9, 'bold'), height =1, width=15, bd=4, command = GenerateVibAUX_TB)
    btnGenerateAUXTB.place(x=140,y=625)

    btnExportGeneratedAUXTB = Button(window, text="Export Generated AUX TB", font=('aerial', 9, 'bold'), height =1, width=22, bd=4, command = ExportAUXTB_Report)
    btnExportGeneratedAUXTB.place(x=340,y=625)

    btnSortByShotNumber = Button(window, text="Sort By Shot Number", font=('aerial', 9, 'bold'), height =1, width=17, bd=1, command = SortByShotNumber)
    btnSortByShotNumber.place(x=150,y=5)

    btnSortByProfileID = Button(window, text="Sort By Profile ID", font=('aerial', 9, 'bold'), height =1, width=14, bd=1, command = SortByProfileID)
    btnSortByProfileID.place(x=300,y=5)

    btnViewAuxProfile = Button(window, text="View Aux Profile", font=('aerial', 9, 'bold'), height =1, width=14, bd=2, command = ViewAuxTBProfile)
    btnViewAuxProfile.place(x=1075,y=475)

    btnSetupAuxProfile = Button(window, text="Setup Aux Profile", font=('aerial', 9, 'bold'), height =1, width=14, bd=2, command = SetupAuxTBProfile)
    btnSetupAuxProfile.place(x=1075,y=505)

    btnUpdateSelected = Button(window, text="Update Selected", font=('aerial', 9, 'bold'), height =1, width=14, bd=2, command = update)
    btnUpdateSelected.place(x=1242,y=475)

    btnDeleteSelected = Button(window, text="Delete Selected", font=('aerial', 9, 'bold'), height =1, width=14, bd=2, command = DeleteSelected)
    btnDeleteSelected.place(x=1242,y=505)

    btnClearView = Button(window, text="ClearView", font=('aerial', 9, 'bold'), height =1, width=10, bd=4, command = ClearView)
    btnClearView.place(x=920,y=625)

    btnExit = Button(window, text="Exit", font=('aerial', 9, 'bold'), height =1, width=6, bd=4, command = iExit)
    btnExit.place(x=1020,y=625)

    btnAddProfile = Button(window, text="Add New", font=('aerial', 9, 'bold'), height =1, width=7, bd=2, command = AddData)
    btnAddProfile.place(x=1290,y=140)

    btnAUXCreateClearAll = Button(window, text="Clear All", font=('aerial', 9, 'bold'), height =1, width=7, bd=2, command = AuxCreateClearEntry)
    btnAUXCreateClearAll.place(x=1290,y=25)












