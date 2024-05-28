#Front End
import os
from tkinter import*
import tkinter.messagebox
import Eagle_HAWK_OBLog_BackEnd
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

def HAWK_OB_LogIMPORT():
    Default_Date_today   = datetime.date.today()
    window = Tk()
    window.title ("Eagle INOVA HAWK OBLog Import Wizard")
    window.geometry("1350x650+10+0")
    window.config(bg="cadet blue")
    window.resizable(0, 0)
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
                                             "column31", "column32", "column33", "column34" ), height=26, show='headings')
    scrollbary.config(command=tree.yview)
    scrollbary.pack(side=RIGHT, fill=Y)
    scrollbarx.config(command=tree.xview)
    scrollbarx.pack(side=BOTTOM, fill=X)        
    tree.heading("#1", text="Field Rec ID", anchor=W)
    tree.heading("#2", text="EP", anchor=W)
    tree.heading("#3", text="ShotID", anchor=W)        
    tree.heading("#4", text="Omit", anchor=W)
    tree.heading("#5", text="FileType", anchor=W)            
    tree.heading("#6", text="File CBS", anchor=W)        
    tree.heading("#7", text="File CAS" ,anchor=W)
    tree.heading("#8", text="UnCorrStack", anchor=W)
    tree.heading("#9", text="Cor EP", anchor=W)        
    tree.heading("#10", text="Uncor EP", anchor=W)
    tree.heading("#11", text="TB Unix" ,anchor=W)
    tree.heading("#12", text="TB (mSecs)", anchor=W)
    tree.heading("#13", text="TB (uSecs)", anchor=W)        
    tree.heading("#14", text="RecLength (mSecs)", anchor=W)
    tree.heading("#15", text="Acquisition (mSecs)", anchor=W)        
    tree.heading("#16", text="Source Line", anchor=W)
    tree.heading("#17", text="Source Station", anchor=W)
    tree.heading("#18", text="Source Type", anchor=W)
    tree.heading("#19", text="Vibes64", anchor=W)
    tree.heading("#20", text="SampleRate (uSecs)", anchor=W)
    tree.heading("#21", text="SourceX", anchor=W)        
    tree.heading("#22", text="SourceY", anchor=W)            
    tree.heading("#23", text="SourceZ", anchor=W)
    tree.heading("#24", text="GridUnits" ,anchor=W)
    tree.heading("#25", text="SweepFile", anchor=W)        
    tree.heading("#26", text="SweepID", anchor=W)
    tree.heading("#27", text="SweepType", anchor=W)
    tree.heading("#28", text="SweepStartFreq" ,anchor=W)        
    tree.heading("#29", text="SweepEndFreq", anchor=W)
    tree.heading("#30", text="SweepLength", anchor=W)
    tree.heading("#31", text="TaperType", anchor=W)
    tree.heading("#32", text="StartTaperDuration", anchor=W)        
    tree.heading("#33", text="EndTaperDuration", anchor=W)
    tree.heading("#34", text="Comment", anchor=W)        
    tree.column('#1', stretch=NO, minwidth=0, width=90)            
    tree.column('#2', stretch=NO, minwidth=0, width=40)
    tree.column('#3', stretch=NO, minwidth=0, width=60)
    tree.column('#4', stretch=NO, minwidth=0, width=40)
    tree.column('#5', stretch=NO, minwidth=0, width=60)
    tree.column('#6', stretch=NO, minwidth=0, width=60)
    tree.column('#7', stretch=NO, minwidth=0, width=60)
    tree.column('#8', stretch=NO, minwidth=0, width=90)
    tree.column('#9', stretch=NO, minwidth=0, width=60)
    tree.column('#10', stretch=NO, minwidth=0, width=60)
    tree.column('#11', stretch=NO, minwidth=0, width=110)
    tree.column('#12', stretch=NO, minwidth=0, width=90)
    tree.column('#13', stretch=NO, minwidth=0, width=90)
    tree.column('#14', stretch=NO, minwidth=0, width=120)
    tree.column('#15', stretch=NO, minwidth=0, width=120)
    tree.column('#16', stretch=NO, minwidth=0, width=90)
    tree.column('#17', stretch=NO, minwidth=0, width=90)
    tree.column('#18', stretch=NO, minwidth=0, width=90)            
    tree.column('#19', stretch=NO, minwidth=0, width=90)
    tree.column('#20', stretch=NO, minwidth=0, width=90)
    tree.column('#21', stretch=NO, minwidth=0, width=90)
    tree.column('#22', stretch=NO, minwidth=0, width=90)
    tree.column('#23', stretch=NO, minwidth=0, width=90)
    tree.column('#24', stretch=NO, minwidth=0, width=90)
    tree.column('#25', stretch=NO, minwidth=0, width=90)
    tree.column('#26', stretch=NO, minwidth=0, width=90)
    tree.column('#27', stretch=NO, minwidth=0, width=90)
    tree.column('#28', stretch=NO, minwidth=0, width=90)
    tree.column('#29', stretch=NO, minwidth=0, width=90)
    tree.column('#30', stretch=NO, minwidth=0, width=90)
    tree.column('#31', stretch=NO, minwidth=0, width=90)
    tree.column('#32', stretch=NO, minwidth=0, width=90)
    tree.column('#33', stretch=NO, minwidth=0, width=90)
    tree.column('#34', stretch=NO, minwidth=0, width=90)
    style = ttk.Style()
    style.theme_use("clam")
    style.configure(".", font=('aerial', 8), foreground="black")
    style.configure("Treeview", foreground='black')
    style.configure("Treeview.Heading",font=('aerial', 8,'bold'), background='Ghost White', foreground='blue',fieldbackground='Ghost White')
    tree.pack()

    ### All Functions defining 

    def iExit():
        iExit= tkinter.messagebox.askyesno("Eagle HAWK OBLog Import", "Confirm if you want to exit")
        if iExit >0:
            window.destroy()
            return

    def ViewTotalImport():
        tree.delete(*tree.get_children())
        txtInvalidEntries.delete(0,END)
        txtDuplicatedShotID.delete(0,END)
        txtTotalEntries.delete(0,END)
        conn = sqlite3.connect("HAWK_OBLog.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_HAWK_OBLog_TEMP ORDER BY `ShotID` ASC ;", conn)
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
        conn = sqlite3.connect("HAWK_OBLog.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_HAWK_OBLog_INVALID ORDER BY `ShotID` ASC ;", conn)
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
        conn = sqlite3.connect("HAWK_OBLog.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_HAWK_OBLog_DUPLICATESHOTID ORDER BY `ShotID` ASC ;", conn)
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
        conn = sqlite3.connect("HAWK_OBLog.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_HAWK_OBLog_TEMP;", conn)
        data = pd.DataFrame(Complete_df)
        TotalEntries = len(data)       
        txtTotalEntries.insert(tk.END,TotalEntries)              
        conn.commit()
        conn.close()

    def InvalidEntries():
        conn = sqlite3.connect("HAWK_OBLog.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_HAWK_OBLog_INVALID;", conn)
        data = pd.DataFrame(Complete_df)
        TotalInvalidEntries = len(data)       
        txtInvalidEntries.insert(tk.END,TotalInvalidEntries)              
        conn.commit()
        conn.close()

    def DuplicatedShotIDEntries():
        conn = sqlite3.connect("HAWK_OBLog.db")
        Complete_df = pd.read_sql_query("SELECT * FROM Eagle_HAWK_OBLog_DUPLICATESHOTID;", conn)
        data = pd.DataFrame(Complete_df)
        TotalDuplicatedEntries = len(data)       
        txtDuplicatedShotID.insert(tk.END,TotalDuplicatedEntries)              
        conn.commit()
        conn.close()

      
    def DeleteSelectedImportData():
        iDelete = tkinter.messagebox.askyesno("Delete Entry", "Confirm if you want to Delete")
        if iDelete >0:
            txtTotalEntries.delete(0,END)
            conn = sqlite3.connect("HAWK_OBLog.db")
            cur = conn.cursor()                
            for selected_item in tree.selection():
                cur.execute("DELETE FROM Eagle_HAWK_OBLog_TEMP WHERE ShotID =? AND SourceLine =? AND \
                            SourceStation =? AND SweepID =? ",\
                            (tree.set(selected_item, '#3'), tree.set(selected_item, '#16'),tree.set(selected_item, '#17'),\
                             tree.set(selected_item, '#26'),)) 
                conn.commit()
                tree.delete(selected_item)
            conn.commit()
            conn.close()
            TotalEntries()
            return

    def ImportHAWKOBLogFile():
        tree.delete(*tree.get_children())
        txtTotalEntries.delete(0,END)
        txtInvalidEntries.delete(0,END)
        txtDuplicatedShotID.delete(0,END)
        fileList = askopenfilenames(initialdir = "/", title = "Import HAWK Time Break Files" , filetypes=[('CSV File', '*.csv'), ('Excel File', ('*.xls', '*.xlsx'))])
        Length_fileList  =  len(fileList)
        if Length_fileList >0:              
            if fileList:
                dfList =[]            
                for filename in fileList:
                    if filename.endswith('.csv'):
                        df = pd.read_csv(filename, header = None, skiprows = {0})
                    else:
                        df = pd.read_excel(filename, header = None, skiprows = {0})
                    dfList.append(df)
                concatDf = pd.concat(dfList,axis=0, ignore_index =True)
                concatDf.rename(columns = {0:'MasterSystemFieldRecordID', 1:'EPNumber', 2:'ShotID', 3:'Omit', 4:'FileType', 5:'File_CorrBeforeStack',
                                           6:'File_CorrAfterStack', 7: 'File_UncorrStack', 8: 'File_CorrEP', 9: 'File_UncorrEP', 10: 'Timebreak_SecondUnixTimeStamp' ,
                                           11: 'Timebreak_mSecs', 12: 'Timebreak_uSecs' , 13: 'RecordLength_mSecs', 14: 'Acquisition_Time_mSecs', 15: 'SourceLine',
                                           16:'SourceStation', 17:'SourceType_DynamiteorVibroseis', 18:'Vibes64_bit_mask', 19:'SampleRateuSecs', 20:'SourceX',
                                           21:'SourceY', 22:'SourceZ', 23:'GridUnits', 24:'SweepFile', 25:'SweepID', 26:'SweepType', 27:'SweepStartFrequency',
                                           28:'SweepEndFrequency', 29:'SweepLength', 30:'TaperType', 31:'StartTaperDuration', 32:'EndTaperDuration',
                                           33:'Comment'},inplace = True)        
                concatDf = concatDf.reset_index(drop=True)

                # Separating InValid with Shot ID is Null
                Invalid_TB_DF     = pd.DataFrame(concatDf)
                Invalid_TB_DF     = Invalid_TB_DF[pd.to_numeric(Invalid_TB_DF.ShotID,errors='coerce').isnull()]                    
                Invalid_TB_DF     = Invalid_TB_DF.reset_index(drop=True)
                Data_Invalid_TB   = pd.DataFrame(Invalid_TB_DF)

                # Separating Valid with Shot ID Not Null
                Valid_TB_DF       = pd.DataFrame(concatDf)
                Valid_TB_DF       = Valid_TB_DF[pd.to_numeric(Valid_TB_DF.ShotID, errors='coerce').notnull()]                  
                Valid_TB_DF["SourceLine"].fillna(0, inplace = True)
                Valid_TB_DF["SourceStation"].fillna(0, inplace = True)
                Valid_TB_DF["MasterSystemFieldRecordID"].fillna(0, inplace = True)
                Valid_TB_DF["EPNumber"].fillna(1, inplace = True)
                Valid_TB_DF['SourceLine']   = (Valid_TB_DF.loc[:,['SourceLine']]).astype(int)
                Valid_TB_DF['SourceStation']= (Valid_TB_DF.loc[:,['SourceStation']]).astype(float)
                Valid_TB_DF['ShotID']       = (Valid_TB_DF.loc[:,['ShotID']]).astype(int)
                Valid_TB_DF['MasterSystemFieldRecordID'] = (Valid_TB_DF.loc[:,['MasterSystemFieldRecordID']]).astype(int)
                Valid_TB_DF['EPNumber']                  = (Valid_TB_DF.loc[:,['EPNumber']]).astype(int)
                Valid_TB_DF['DuplicatedEntries']         = Valid_TB_DF.sort_values(by =['ShotID','MasterSystemFieldRecordID', 'EPNumber']).duplicated(['ShotID'],keep='last')
                Valid_TB_DF                              = Valid_TB_DF.reset_index(drop=True)
                Valid_TB_DF                              = pd.DataFrame(Valid_TB_DF)

                # Separating Valid with Shot ID Not Duplicated
                DATA_VALID_TB  = Valid_TB_DF.loc[Valid_TB_DF.DuplicatedEntries == False, 'MasterSystemFieldRecordID': 'Comment']
                DATA_VALID_TB  = DATA_VALID_TB.reset_index(drop=True)
                DATA_VALID_TB  = pd.DataFrame(DATA_VALID_TB)

                # Separating Valid with Shot ID Duplicated
                DATA_DuplicatedShotID = Valid_TB_DF.loc[Valid_TB_DF.DuplicatedEntries == True, 'MasterSystemFieldRecordID': 'Comment']
                DATA_DuplicatedShotID = DATA_DuplicatedShotID.reset_index(drop=True)
                DATA_DuplicatedShotID = pd.DataFrame(DATA_DuplicatedShotID)

                # Connect To Database and Export DF              
                con= sqlite3.connect("HAWK_OBLog.db")
                cur=con.cursor()                
                DATA_VALID_TB.to_sql('Eagle_HAWK_OBLog_TEMP',con, if_exists="replace", index=False)
                Data_Invalid_TB.to_sql('Eagle_HAWK_OBLog_INVALID',con, if_exists="replace", index=False)
                DATA_DuplicatedShotID.to_sql('Eagle_HAWK_OBLog_DUPLICATESHOTID',con, if_exists="replace", index=False)
                con.commit()
                cur.close()
                con.close()

                # Populating Valid Import to TreeView
                ViewTotalImport()
        else:
            tkinter.messagebox.showinfo("Import HAWK TB Error Message","Please Select HAWK TB Imported Files")

    ### Entry Wizard
    txtTotalEntries  = Entry(window, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 10)
    txtTotalEntries.place(x=1250,y=6)
    txtInvalidEntries  = Entry(window, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
    txtInvalidEntries.place(x=475,y=6)
    txtDuplicatedShotID  = Entry(window, font=('aerial', 12, 'bold'),textvariable=IntVar(), width = 8)
    txtDuplicatedShotID.place(x=475,y=620)

    L1 = Label(window, text = "Inova HAWK OB Log Details:", font=("arial", 10,'bold'),bg = "green").place(x=2,y=6)
    L2 = Label(window, text = "No ShotID Or File Number", font=("arial", 10,'bold'),bg = "red").place(x=555,y=7)
    L3 = Label(window, text = "Duplicated ShotID", font=("arial", 10,'bold'),bg = "red").place(x=555,y=620)

    ### Button Wizard  
    btnImportHAWKOBLog = Button(window, text="Import InovaHAWK OBLog", font=('aerial', 9, 'bold'), height =1, width=21, bd=4, command = ImportHAWKOBLogFile)
    btnImportHAWKOBLog.place(x=2,y=620)
    btnDuplicatedShotIDImport= Button(window, text="View Duplicated ShotID", font=('aerial', 9, 'bold'), height =1, width=19, bd=2, command =ViewDuplicatedShotIDImport)
    btnDuplicatedShotIDImport.place(x=332,y=620)
    btnInValidImport= Button(window, text="View Invalid Import", font=('aerial', 9, 'bold'), height =1, width=16, bd=1, command = ViewInvalidImport)
    btnInValidImport.place(x=350,y=6)
    btnViewImportHAWKOBLog = Button(window, text="View Valid Import", font=('aerial', 9, 'bold'), height =1, width=14, bd=1, command = ViewTotalImport)
    btnViewImportHAWKOBLog.place(x=1140,y=6)

    btnExit = Button(window, text="Exit Import", font=('aerial', 9, 'bold'), height =1, width=10, bd=4, command = iExit)
    btnExit.place(x=1267,y=620)
    btnClearView = Button(window, text="Clear View", font=('aerial', 9, 'bold'), height =1, width=10, bd=4, command = ClearView)
    btnClearView.place(x=1181,y=620)
    btnDelete = Button(window, text="Delete Selected Valid", font=('aerial', 9, 'bold'), height =1, width=18, bd=4, command = DeleteSelectedImportData)
    btnDelete.place(x=1040,y=620)
















