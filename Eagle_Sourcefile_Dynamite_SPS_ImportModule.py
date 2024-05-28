import sqlite3
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
import Eagle_Sourcefile_Dynamite_SPS_BackEnd

def SourceFile_Dynamite_ImportModule():
    window = Tk()
    window.title ("Eagle Dynamite Source File (SPS) Import Wizard")
    window.geometry("550x660+10+10")
    window.config(bg="cadet blue")
    window.resizable(0, 0)
    TableMargin = Frame(window, bd = 2, padx= 10, pady= 2, relief = RIDGE)
    TableMargin.pack(side=TOP)
    TableMargin.pack(side=LEFT)
    scrollbarx = Scrollbar(TableMargin, orient=HORIZONTAL)
    scrollbary = Scrollbar(TableMargin, orient=VERTICAL)
    tree = ttk.Treeview(TableMargin, column=("column1", "column2", "column3"), height=27, show='headings')
    scrollbary.config(command=tree.yview)
    scrollbary.pack(side=RIGHT, fill=Y)
    scrollbarx.config(command=tree.xview)
    scrollbarx.pack(side=BOTTOM, fill=X)        
    tree.heading("#1", text="SourceLine", anchor=W)
    tree.heading("#2", text="SourceStation", anchor=W)
    tree.heading("#3", text="SourceLineStationCombined", anchor=W)
    tree.column('#1', stretch=NO, minwidth=0, width=150)            
    tree.column('#2', stretch=NO, minwidth=0, width=150)
    tree.column('#3', stretch=NO, minwidth=0, width=200)
    style = ttk.Style()
    style.theme_use("clam")
    style.configure(".", font=('aerial', 8), foreground="black")
    style.configure("Treeview", foreground='black')
    style.configure("Treeview.Heading",font=('aerial', 8,'bold'), background='Ghost White', foreground='blue',fieldbackground='Ghost White')
    tree.pack()

    def UpdateimportedSourceSPS():
        con= sqlite3.connect("DynamiteSourceSPS.db")
        cur=con.cursor()
        Complete_df_SPS = pd.read_sql_query("SELECT * FROM SourceFileSPS ORDER BY `SourceLineStationCombined` ASC;", con)
        data = pd.DataFrame(Complete_df_SPS)
        data = data.drop_duplicates(['SourceLineStationCombined'], keep='last')
        data = data.reset_index(drop=True)
        data.to_sql('SourceFileSPS',con, if_exists="replace", index=False)
        con.commit()
        cur.close()
        con.close()

    def importSourceSPS():
        txtTotalEntries.delete(0,END)
        SourcefileList = askopenfilenames(initialdir = "/", title = "Import Dynamite Source Files" , filetypes=[('CSV File', '*.csv'), ('Excel File', ('*.xls', '*.xlsx'))])
        Length_fileList  =  len(SourcefileList)
        if Length_fileList >0:        
            if SourcefileList:
                df_SourceList = []           
                for filename in SourcefileList:
                    if filename.endswith('.csv'):
                        dfS               = pd.read_csv(filename, sep=',',low_memory=False)
                        dfS               = dfS.iloc[:,:]
                        SourceLine        = dfS.loc[:,'SourceLine']
                        SourceStation     = dfS.loc[:,'SourceStation']        
                        column_names      = [SourceLine, SourceStation]                                    
                        catdf_SPS         = pd.concat (column_names,axis=1,ignore_index =True)
                        df_SourceList.append(catdf_SPS) 
                    else:
                        dfS               = pd.read_excel(filename)
                        dfS               = dfS.iloc[:,:]
                        SourceLine        = dfS.loc[:,'SourceLine']
                        SourceStation     = dfS.loc[:,'SourceStation']        
                        column_names      = [SourceLine, SourceStation]                                    
                        catdf_SPS         = pd.concat (column_names,axis=1,ignore_index =True)
                        df_SourceList.append(catdf_SPS)

                concatDfS = pd.concat(df_SourceList,axis=0)
                concatDfS.rename(columns={0:'SourceLine', 1:'SourceStation'},inplace = True)
                SourceDF = pd.DataFrame(concatDfS)
                SourceDF = SourceDF.reset_index(drop=True)
                SourceDF['SourceLine']       = (SourceDF.loc[:,['SourceLine']]).astype(int)
                SourceDF['SourceStation']    = (SourceDF.loc[:,['SourceStation']]).astype(float)

                SourceDF_Line_M        = SourceDF['SourceLine'].astype(int)
                SourceDF_Station_M     = SourceDF['SourceStation'].astype(float)
                SourceDF_Line_Station_Combined = (SourceDF_Line_M.map(str) + SourceDF_Station_M.map(str)).astype(float)
                SourceDF_Combined_LS           =  pd.DataFrame(SourceDF_Line_Station_Combined)
                SourceDF_Combined_LS.rename(columns={0:'SourceLineStationCombined'},inplace = True)
                SourceDF_DB                       = pd.concat([SourceDF, SourceDF_Combined_LS],axis=1)
                SourceDF_DB                       = SourceDF_DB.drop_duplicates(['SourceLineStationCombined'],keep='last')
                SourceDF_DB['SourceLineStationCombined'] = (SourceDF_DB.loc[:,['SourceLineStationCombined']]).astype(float)  
                SourceDF_DB                       = SourceDF_DB.reset_index(drop=True)
                SourceDF_DB                       = pd.DataFrame(SourceDF_DB)
                SourceDF_DB['SourceLine']         = (SourceDF_DB.loc[:,['SourceLine']]).astype(int)
                SourceDF_DB['SourceStation']      = (SourceDF_DB.loc[:,['SourceStation']]).astype(float)
                for each_rec in range(len(SourceDF_DB)):
                    tree.insert("", tk.END, values=list(SourceDF_DB.loc[each_rec]))
                TotalEntries = len(SourceDF_DB)       
                txtTotalEntries.insert(tk.END,TotalEntries)   
                ## Connecting to SQL DB        
                con= sqlite3.connect("DynamiteSourceSPS.db")
                cur=con.cursor()                
                SourceDF_DB.to_sql('SourceFileSPS',con, if_exists="append", index=False)
                con.commit()
                cur.close()
                con.close()
                UpdateimportedSourceSPS()
                tkinter.messagebox.showinfo("Import Dynamite Source File Message","Source File Imported to Database Successfully")
        else:
            tkinter.messagebox.showinfo("Import Dynamite Source File Error Message","Please Select Dynamite Source File")


    def ViewMasterSPS():
        txtTotalEntries.delete(0,END)
        conn = sqlite3.connect("DynamiteSourceSPS.db")
        Complete_df_TB = pd.read_sql_query("SELECT * FROM SourceFileSPS ORDER BY `SourceLineStationCombined` ASC;", conn)
        data = pd.DataFrame(Complete_df_TB)
        data = data.reset_index(drop=True)
        for each_rec in range(len(data)):
            tree.insert("", tk.END, values=list(data.loc[each_rec]))
        TotalEntries = len(data)       
        txtTotalEntries.insert(tk.END,TotalEntries)              
        conn.commit()
        conn.close()

    def TotalEntries():
        txtTotalEntries.delete(0,END)
        conn = sqlite3.connect("DynamiteSourceSPS.db")
        Complete_df_TB = pd.read_sql_query("SELECT * FROM SourceFileSPS ORDER BY `SourceLineStationCombined` ASC;", conn)
        data = pd.DataFrame(Complete_df_TB)
        data = data.reset_index(drop=True)
        TotalEntries = len(data)       
        txtTotalEntries.insert(tk.END,TotalEntries)              
        conn.commit()
        conn.close()

    def DeleteSelectedImportData():
        iDelete = tkinter.messagebox.askyesno("Delete Entry", "Confirm if you want to Delete")
        if iDelete >0:
            txtTotalEntries.delete(0,END)
            conn = sqlite3.connect("DynamiteSourceSPS.db")
            cur = conn.cursor()                
            for selected_item in tree.selection():
                cur.execute("DELETE FROM SourceFileSPS WHERE SourceLine =? AND SourceStation =? AND \
                            SourceLineStationCombined =? ",\
                            (tree.set(selected_item, '#1'), tree.set(selected_item, '#2'),tree.set(selected_item, '#3'),)) 
                tree.delete(selected_item)
                conn.commit()
            conn.commit()
            conn.close()
            TotalEntries()
            return

    def ClearView():
        txtTotalEntries.delete(0,END)
        tree.delete(*tree.get_children())


    def iExit():
        iExit= tkinter.messagebox.askyesno("Eagle Dynamite Source File Import Wizard", "Confirm if you want to exit")
        if iExit >0:
            window.destroy()
            return


    ##### Entry Wizard
    txtTotalEntries  = Entry(window, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 10)
    txtTotalEntries.place(x=360,y=6)

    ##### Labeling
    L1 = Label(window, text = " Source File Dynamite:", font=("arial", 10,'bold'),bg = "green").place(x=2,y=6)
    L2 = Label(window, text = "Total Entries", font=("arial", 10,'bold'),bg = "yellow").place(x=460,y=7)

    btnImportSPSFile = Button(window, text="Import Dynamite Source File", font=('aerial', 9, 'bold'), height =1, width=23, bd=4, command = importSourceSPS)
    btnImportSPSFile.place(x=2,y=630)

    btnViewMasterSPSFile = Button(window, text="View Master SPS File", font=('aerial', 9, 'bold'), height =1, width=17, bd=4, command = ViewMasterSPS)
    btnViewMasterSPSFile.place(x=211,y=630)

    btnDeleteSelected = Button(window, text="Delete Selected", font=('aerial', 9, 'bold'), height =1, width=16, bd=4, command = DeleteSelectedImportData)
    btnDeleteSelected.place(x=420,y=630)

    btnClear = Button(window, text="ClearView", font=('aerial', 9, 'bold'), height =1, width=8, bd=1, command = ClearView)
    btnClear.place(x=195,y=6)

    btnExit = Button(window, text="Exit", font=('aerial', 9, 'bold'), height =1, width=4, bd=1, command = iExit)
    btnExit.place(x=270,y=6)
















