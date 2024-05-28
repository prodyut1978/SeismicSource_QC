#Import Python Modules
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
import openpyxl
import csv
import time
import datetime

# Import Back End File DB Create
import Eagle_Sourcefile_Vibroseis_SPS_BackEnd
import Eagle_Sourcefile_Dynamite_SPS_BackEnd
import Eagle_HAWK_OBLog_BackEnd
import GeomergeAUXTB_BackEnd
import Eagle_SourceLink_Vibroseis_Log_BackEnd
import Eagle_SourceLink_Dynamite_Log_BackEnd
import Eagle_SourceLink_PSSLog_Analysis_QC_BackEnd
import GeoSpaceOutputLog_BackEnd
import Global_Mapper_SnailTrails_Merge_BackEnd

# Import SourceLink Vib QC and Daily production Reporting
import SetupVibQCLimit
import Eagle_SourceLink_AnalyzePSS_Import_Module
import Eagle_SourceLink_Vib_ProductionReport_InitialSetup
import Eagle_SourceLink_Vib_ProductionReport_ExistingSetup
import Vib_ProductionQuality_Report

# Import SourceLink Vib Miscellaneous Utility Modules
import SetupVibAuxProfile
import GeomergeAUXTB
import GeospaceOutputLog_TraceYield_Module
import Global_Mapper_SnailTrails_Merge

# Import SPS Source File
import Eagle_Sourcefile_Vibroseis_SPS_ImportModule
import Eagle_Sourcefile_Dynamite_SPS_ImportModule

# Import Advanced Modules
import Eagle_SourceLink_TB_Offset_Module_BackEnd
import Eagle_SourceLink_TB_Offset_Module

# Import HAWK Modules
import Eagle_HAWK_OBLog_Import_Module
import Eagle_Source_QC_MERGE_HAWKOBLOG
import Eagle_HAWK_OBLog_MasterDBWizard
import Eagle_HAWK_OBLog_MissingShotReport

# Import SourceLink VIB Import and Merge TB - PSS - Vib Position
import Eagle_SourceLink_TB_Vibroseis_Import_Module
import Eagle_SourceLink_PSS_Import_Module
import Eagle_SourceLink_VIBPositionImport_Module
import Eagle_Source_QC_MERGE_SOURCELINKTBLOG
import Eagle_SourceLink_TB_Vibroseis_MasterDBWizard
import Eagle_SourceLink_TB_MissingShotReport

# Import SourceLink Dynamite Import and Merge TB - PFS - DrillingLog
import Eagle_SourceLink_TB_Dynamite_Import_Module
import Eagle_SourceLink_PFS_Import_Module
import Eagle_MERGE_Sourcelink_Dynamite_TB_PFS
import Eagle_SourceLink_TB_Dynamite_MasterDBWizard
import Eagle_SourceLink_TB_Dynamite_MissingShotReport



class SourceQCModule:    
    def __init__(self,root):

        Default_Date_today   = datetime.date.today()
        self.root =root
        self.root.title ("In Field Seismic Source QC")
        self.root.geometry("1180x700+10+0")
        self.root.config(bg="cadet blue")
        self.root.resizable(0, 0)
        

        ## Define Function
        def iExit():
            iExit= tkinter.messagebox.askyesno("Exit Source QC Widget", "Confirm If You Want To Exit")
            if iExit >0:
                self.root.destroy()
            return


        def ClearVibSourceLinkTBMasterDB():
            try:
                conn = sqlite3.connect("SourceLink_Log.db")
                cur=conn.cursor()
                cur.execute("DELETE FROM Eagle_SOURCELINKTBMASTER")
                conn.commit()
                conn.close()
                tkinter.messagebox.showinfo("Clear Vib SourceLink TB Master DB","SourceLink Vib Master TB Database is cleared")
            except:
                tkinter.messagebox.showinfo("Clear Vib SourceLink TB Master DB","SourceLink Vib Master TB Database is Already Cleared")


        def ClearDynSourceLinkTBMasterDB():
            try:
                conn = sqlite3.connect("SourceLink_Dynamite_Log.db")
                cur=conn.cursor()
                cur.execute("DELETE FROM Eagle_SOURCELINK_DYNAMITE_TBMASTER")
                conn.commit()
                conn.close()
                tkinter.messagebox.showinfo("Clear Dynamite SourceLink TB Master DB","SourceLink Dynamite Master TB Database is cleared")
            except:
                tkinter.messagebox.showinfo("Clear Dynamite SourceLink TB Master DB","SourceLink Dynamite Master TB Database is Already Cleared")


        def ClearVibInovaHAWKTBMasterDB():
            try:
                conn = sqlite3.connect("HAWK_OBLog.db")
                cur=conn.cursor()
                cur.execute("DELETE FROM Eagle_HAWK_OBLog_MASTER")
                conn.commit()
                conn.close()
                tkinter.messagebox.showinfo("Clear Vib HAWK TB Master DB","HAWK Vib Master TB Database is cleared")
            except:
                tkinter.messagebox.showinfo("Clear Vib HAWK TB Master DB","HAWK Vib Master TB Database is Already Cleared")

        def ClearSPSMasterDB():
            try:
                conn = sqlite3.connect("SourceSPS.db")
                cur=conn.cursor()
                cur.execute("DELETE FROM SourceFileSPS")
                conn.commit()
                conn.close()
                tkinter.messagebox.showinfo("Clear Source Master DB","Source Database is cleared")
            except:
                tkinter.messagebox.showinfo("Clear Source Master DB","Source Database is Already Cleared")


        ## Advanced Menu
        def SourceLink_TB_Offset_Microseconds():
            Eagle_SourceLink_TB_Offset_Module.SourceLink_TB_Offset()

        ## SourceLink Vib Modules        
        def importVIBSourceSPS():
            Eagle_Sourcefile_Vibroseis_SPS_ImportModule.SourceFile_Vibroseis_ImportModule()

        def SourceLink_TB_LogIMPORT():
            Eagle_SourceLink_TB_Vibroseis_Import_Module.SourceLink_TB_LogIMPORT()

        def SourceLink_PSS_LogIMPORT():
            Eagle_SourceLink_PSS_Import_Module.SourceLink_PSS_LogIMPORT()

        def SourceLink_Vib_PositionIMPORT():
            Eagle_SourceLink_VIBPositionImport_Module.SourceLink_COG_LogIMPORT()

        def Merge_SourceLINK_TB():
            Eagle_Source_QC_MERGE_SOURCELINKTBLOG.Merge_Sourcelink_TBLog()

        def SourceLink_Vibroseis_MasterTB_LogIMPORT():
            Eagle_SourceLink_TB_Vibroseis_MasterDBWizard.SourceLink_TB_SubmitToMasterDB()

        def SourceLink_Vibroseis_TB_MissingShot():
            Eagle_SourceLink_TB_MissingShotReport.GenerateSourceLinkTBMissingShotReport()

        ## Hawk Vib Modules
        def Inova_HAWK_TB_IMPORT():
            Eagle_HAWK_OBLog_Import_Module.HAWK_OB_LogIMPORT()

        def Merge_InovaHAWK_TB():
            Eagle_Source_QC_MERGE_HAWKOBLOG.Merge_HAWK_OBLog()

        def HAWK_MasterTB_LogIMPORT():
            Eagle_HAWK_OBLog_MasterDBWizard.HAWK_OB_LogMasterDBIMPORT()

        def HAWK_TB_MissingShot():
            Eagle_HAWK_OBLog_MissingShotReport.GenerateHawkTBMissingShotReport()


        ## SourceLink Dynamite Modules
        def importDynSourceSPS():
            Eagle_Sourcefile_Dynamite_SPS_ImportModule.SourceFile_Dynamite_ImportModule()

        def SourceLink_Dynamite_TB_LogIMPORT_Module():
            Eagle_SourceLink_TB_Dynamite_Import_Module.SourceLink_TB_Dynamite_LogIMPORT()

        def SourceLink_Dynamite_PFS_LogIMPORT_Module():
            Eagle_SourceLink_PFS_Import_Module.SourceLink_PFS_LogIMPORT()

        def SourceLink_Dynamite_TB_PFS_Merge_Module():
            Eagle_MERGE_Sourcelink_Dynamite_TB_PFS.Merge_Sourcelink_Dynamite_TBLog()

        def SourceLink_Dynamite_Master_TB_LogIMPORT_Module():
            Eagle_SourceLink_TB_Dynamite_MasterDBWizard.SourceLink_DynamiteTB_SubmitToMasterDB()

        def SourceLink_Dynamite_TB_MissingShot():
            Eagle_SourceLink_TB_Dynamite_MissingShotReport.GenerateSourceLinkTB_Dynamite_MissingShotReport()

        ## Misc Modules
        def GeomergeOutputLogTraceyield():
            GeospaceOutputLog_TraceYield_Module.TraceYieldReport()

        def Generate_VibAUX_SignatureTB():
            GeomergeAUXTB.GenerateGeomergeTBVibSignature()

        def Vib_SnailTrail_MergeOBLog():
            Global_Mapper_SnailTrails_Merge.Vib_SnailTrail_Merge()
            

        ## Vib QC Analysis
        def SetupVibQCLimitParameter():
            SetupVibQCLimit.VibQCLimitParameter()
            
        def VibQCAnalysisTestiFyInput():
            Eagle_SourceLink_AnalyzePSS_Import_Module.SourceLink_AnalyzePSS_LogIMPORT()

        def Generate_VibproductionQualityReport():
            Vib_ProductionQuality_Report.VibProductionQuality()
        
        def Generate_VibproductionStatReport():
            if not os.path.exists("C:\VibRestrictedFolder\VibProductionModule\VibProfileParameter"):
                tkinter.messagebox.showinfo("Vib Profile Input Message","Please Provide job Specific Vib Input Profile")
                Eagle_SourceLink_Vib_ProductionReport_InitialSetup.VibProductionReportInitialSetup()
            else:
                Eagle_SourceLink_Vib_ProductionReport_ExistingSetup.VibProductionReportExistingSetup()
            
        def Reset_VibproductionStatReport():
            if not os.path.exists("C:\VibRestrictedFolder\VibProductionModule\VibProfileParameter"):
                tkinter.messagebox.showinfo("Vib Profile Input Message","Please Provide job Specific Vib Input Profile")
            else:
                os.remove("C:\VibRestrictedFolder\VibProductionModule\VibProfileParameter")
                tkinter.messagebox.showinfo("Vib Production Report Reset","Vib Production Report Reset And Now Please Provide job Specific Vib Input Profile")

        def Help_VibQCLimitParameter():
            HSE_Event_Help =  ("Procedure To Setup Vib QC Limit: " + '\n' + '\n' + 
                   "1. Vib QC Parameters ForceAvg/ForceMax Value Range : "  + '\n' + '\n' + 
                   "    ForceAvg QC Failed <   [ Min ForceAvg Value] " + '\n' + '\n' +
                   "    ForceAvg QC Failed >   [ Max ForceAvg Value]" +  '\n' + '\n' +
                   "    ForceAvg QC Passed >= [ Min ForceAvg Value]" +  '\n' + '\n' +
                   "    ForceAvg QC Passed <= [ Max ForceAvg Value]" +  '\n' + '\n' +'\n' +

                   "    ForceMax QC Failed <   [ Min ForceMax Value] " + '\n' + '\n'  +
                   "    ForceMax QC Failed >   [ Max ForceMax Value] " +  '\n' + '\n' +
                   "    ForceMax QC Passed >= [ Min ForceMax Value]" +  '\n' + '\n' +
                   "    ForceMax QC Passed <= [ Max ForceMax Value]" +  '\n' + '\n' +'\n' +

                   "2. Vib QC Parameters THDAvg/THDMax Value Range : "  + '\n' + '\n' + 
                   "    THDAvg QC Failed >   [ Max THDAvg Value]" +  '\n' + '\n' +
                   "    THDAvg QC Passed <= [ Max THDAvg Value]" +  '\n' + '\n' +'\n' +

                   "    THDMax QC Failed >   [ Max THDMax Value]" +  '\n' + '\n' +
                   "    THDMax QC Passed <= [ Max THDMax Value]" +  '\n' + '\n' +'\n' +
                               
                   "3. Vib QC Parameters PhaseAvg/PhaseMax Value Range : "  + '\n' + '\n' + 
                   "    PhaseAvg QC Failed >  Absolute[ Max PhaseAvg Value]" +  '\n' + '\n' +
                   "    PhaseAvg QC Passed <= Absolute[ Max PhaseAvg Value]" +  '\n' + '\n' +'\n' +

                   "    PhaseMax QC Failed >   Absolute[ Max PhaseMax Value]" +  '\n' + '\n' +
                   "    PhaseMax QC Passed <= Absolute[ Max PhaseMax Value]" +  '\n' + '\n' +'\n' +
                                                  
                   "4. GPSQuality QC Passed != No Fix" )
            

            tkinter.messagebox.showinfo("Vib QC Limit Setup Procedure", HSE_Event_Help)
            

        def Help_To_Make_SourceFile():
            SourceFile_Help =  ("Procedure To Make Source File From SPS File: " + '\n' + '\n' + 
                   "1. Open Excel/CSV File With Column Name:" + '\n' +  "   Column-1: SourceLine And Column-2: SourceStation  "  + '\n' + '\n' + 
                   "2. Copy Source Line And Station From SPS And Paste in the Proper Column " + '\n' + '\n' +
                   "3. Continue Appending Col-1. SourceLine Col-2. SourceStation From Updated Version Of SPS Line And Points" +  '\n' + '\n' +
                   "4. Before Start New Job Always Reset SPS Database (Reset Master Database > Clear Source Master SPS) And Import New Source File")
            tkinter.messagebox.showinfo("Source File Procedure", SourceFile_Help)
        def Under_Construction():
            tkinter.messagebox.showinfo("Under Construction", " Will Be Added later Version")

        def HAWK_QC_Module():
            window = Tk()
            window.title("HAWK In-Field QC")
            window.config(bg="cadet blue")
            width = 360
            height = 640
            screen_width = window.winfo_screenwidth()
            screen_height = window.winfo_screenheight()
            x = (screen_width/2) - (width/2)
            y = (screen_height/2) - (height/2)
            window.geometry("%dx%d+%d+%d" % (width, height, x, y))
            window.resizable(0, 0)
            window.grid()

            DataFrameLEFT = Frame(window)
            DataFrameLEFT.grid(row=2,column=1 ,padx= 15, pady= 10)        
            lblTitVibShotQC = Label(DataFrameLEFT, bd= 4, font=('aerial', 10, 'bold'), width = 40, fg = 'green',
                    bg = "orange", underline =-1, text="HAWK DAILY IN-FIELD VIB QC MODULES")
            lblTitVibShotQC.grid(row=1,column=1)       
            L1 = Label(DataFrameLEFT, text = "A: Daily In-Field Vib Import Modules:", justify ="left", font=("arial", 10,'bold'), fg = 'blue',
                    bg= "ghost white", padx= 1, pady= 2, underline =0).grid(row=2,column=1, sticky =W , padx= 20, pady= 5)
            L2 = Label(DataFrameLEFT, text = "B: Daily In-Field Merge & QC Modules:", justify ="left", font=("arial", 10,'bold'), fg = 'blue',
                    bg= "ghost white", padx= 1, pady= 2, underline =0).grid(row=6,column=1, sticky =W , padx= 20, pady= 5)        
            btnImport_HAWK_TB = Button(DataFrameLEFT, text="HAWK Vib TB Import", font=('aerial', 10, 'bold'),
                    bg= "ghost white" , height =1, width=24, bd=2, padx= 1, pady= 5, command = Inova_HAWK_TB_IMPORT)
            btnImport_HAWK_TB.grid(row = 3, column=1, sticky =W , padx= 40, pady= 5)        
            btnImport_SourceLINK_PSS = Button(DataFrameLEFT, text="SourceLink Vib PSS Import", font=('aerial', 10, 'bold'),
                    bg= "ghost white" , height =1, width=24, bd=2,  padx= 1, pady= 5, command = SourceLink_PSS_LogIMPORT)
            btnImport_SourceLINK_PSS.grid(row=4,column=1, sticky =W , padx= 40, pady= 5)
            btnImport_SourceLINK_COG = Button(DataFrameLEFT, text="SourceLink Position Import", font=('aerial', 10, 'bold'),
                    bg= "ghost white" , height =1, width=24, bd=2,  padx= 1, pady= 5, command = SourceLink_Vib_PositionIMPORT)
            btnImport_SourceLINK_COG.grid(row=5,column=1, sticky =W , padx= 40, pady= 5)        
            btnMerge_HAWK_TB = Button(DataFrameLEFT, text="Merge TB - PSS - VibPosition", font=('aerial', 10, 'bold'),
                    bg= "ghost white" , height =1, width=24, bd=2,  padx= 1, pady= 5, command = Merge_InovaHAWK_TB)
            btnMerge_HAWK_TB.grid(row=7,column=1 , sticky =W , padx= 40, pady= 5)

            DataFrameLEFT_CENTRE = Frame(window)
            DataFrameLEFT_CENTRE.grid(row=20,column=1, padx= 15, pady= 10)
            lblSubmitCleanTBMasterDB = Label(DataFrameLEFT_CENTRE, bd= 4, font=('aerial', 10, 'bold'), width = 40, fg = 'green',
                        bg = "orange", underline =-1, text="ACCUMULATE HAWK TB AND MISSING SHOTS")
            lblSubmitCleanTBMasterDB.grid(row=1,column=1)
            L3 = Label(DataFrameLEFT_CENTRE, text = "A: Daily HAWK TB Submit To MasterDB:", justify ="left", font=("arial", 10,'bold'), fg = 'blue',
                       bg= "ghost white", padx= 1, pady= 2, underline =0).grid(row=2,column=1, sticky =W , padx= 20, pady= 5)
            L4 = Label(DataFrameLEFT_CENTRE, text = "B: Import Vib Source (Line & Station) File:", justify ="left", font=("arial", 10,'bold'), fg = 'blue',
                       bg= "ghost white", padx= 1, pady= 2, underline =0).grid(row=4,column=1, sticky =W , padx= 20, pady= 5)
            L5 = Label(DataFrameLEFT_CENTRE, text = "C: Generate Vib Daily Missing Shot Report:", justify ="left", font=("arial", 10,'bold'), fg = 'blue',
                       bg= "ghost white", padx= 1, pady= 2, underline =0).grid(row=7,column=1, sticky =W , padx= 20, pady= 5)
            btnMaster_HAWK_TB = Button(DataFrameLEFT_CENTRE, text="HAWK Clean TB Import", font=('aerial', 10, 'bold'),
                       bg= "ghost white" ,height =1, width=24, bd=2, padx= 1, pady= 5, command = HAWK_MasterTB_LogIMPORT)
            btnMaster_HAWK_TB.grid(row=3,column=1 , sticky =W , padx= 40, pady= 5)
            btnHelp_To_Make_SourceFile= Button(DataFrameLEFT_CENTRE, text="Help To Make Source File", font=('aerial', 10, 'bold'),
                      bg= "ghost white" ,height =1, width=24, bd=2, padx= 1, pady= 5, command = Help_To_Make_SourceFile)
            btnHelp_To_Make_SourceFile.grid(row=5,column=1 , sticky =W , padx= 40, pady= 5)
            btnimportSourceSPS= Button(DataFrameLEFT_CENTRE, text="Vib Source (SPS) File Import", font=('aerial', 10, 'bold'),
                      bg= "ghost white" ,height =1, width=24, bd=2, padx= 1, pady= 5, command = importVIBSourceSPS)
            btnimportSourceSPS.grid(row=6,column=1 , sticky =W , padx= 40, pady= 5)
            btnGenerateHAWK_TBMissingShot = Button(DataFrameLEFT_CENTRE, text="Generate Missing Shot Report", font=('aerial', 10, 'bold'),
                     bg= "ghost white" ,height =1, width=24, bd=2, padx= 1, pady= 5, command = HAWK_TB_MissingShot)
            btnGenerateHAWK_TBMissingShot.grid(row=8,column=1 , sticky =W , padx= 40, pady= 5)

        def Enable_Advance():
            iEnable_Advance = tkinter.messagebox.askyesno("Enable Advance Mode", "Confirm if you want to Enable Advance Mode")
            if iEnable_Advance >0:
                application_window = self.root
                EnableCodeEntry = simpledialog.askstring("Input Advance Mode Code", "Please Input Advance Mode Code To Enable Advance Mode",
                                  parent=application_window)
                
                if EnableCodeEntry == "EagleQCAdmin":
                    menu.add_cascade(label="Advanced", state=tk.NORMAL, menu=Advanced)
                    Advanced.add_command(label="Offset ShotTime ± MicroSeconds",  state=tk.NORMAL, command=SourceLink_TB_Offset_Microseconds)

                else:
                    tkinter.messagebox.showinfo("Advanced Mode Enable Code", " Please Input Advanced Mode Enable Code")
        
        

        ## Adding File Menu 
        menu = Menu(self.root)
        self.root.config(menu=menu)
        filemenu  = Menu(menu, tearoff=0)
        ImportSPS = Menu(menu, tearoff=0)
        ResetDB   = Menu(menu, tearoff=0)
        Advanced  = Menu(menu, tearoff=0)
        SeismicRecSystem = Menu(menu, tearoff=0)

        menu.add_cascade(label="File", menu=filemenu)        
        menu.add_cascade(label="Import", menu=ImportSPS)
        menu.add_cascade(label="Reset Database", menu=ResetDB)
        menu.add_cascade(label="Seismic System", menu=SeismicRecSystem)
                
        filemenu.add_command(label="Enable Advance Mode", command=Enable_Advance)
        filemenu.add_command(label="Exit", command=iExit)

        ImportSPS.add_command(label="Import Vib Source (SPS) File", command = importVIBSourceSPS)
        ImportSPS.add_command(label="Import Dynamite Source (SPS) File", command = importDynSourceSPS)
        
        ImportSPS.add_command(label="Help To Make Source (SPS) File", command = Help_To_Make_SourceFile)
        
        
        ResetDB.add_command(label="Clear Vib SourceLink Master TB", command=ClearVibSourceLinkTBMasterDB)
        ResetDB.add_command(label="Clear Dynamite SourceLink Master TB", command=ClearDynSourceLinkTBMasterDB)
        ResetDB.add_command(label="Clear Vib HAWK Master TB", command=ClearVibInovaHAWKTBMasterDB)
        ResetDB.add_command(label="Clear Source Master SPS", command=ClearSPSMasterDB)

        SeismicRecSystem.add_command(label="INOVA-HAWK QC Modules", command=HAWK_QC_Module)

        ##  Define SourceLink VIB Import and Merge TB - PSS - Vib Position 
        DataFrameLEFT = Frame(self.root)
        DataFrameLEFT.grid(row=2,column=1 ,padx= 15, pady= 10)        
        lblTitVibShotQC = Label(DataFrameLEFT, bd= 4, font=('aerial', 10, 'bold'), width = 40, fg = 'green',
                    bg = "orange", underline =-1, text="DAILY IN-FIELD VIB QC MODULES")
        lblTitVibShotQC.grid(row=1,column=1)       
        L1 = Label(DataFrameLEFT, text = "A: Daily In-Field Vib Import Modules:", justify ="left", font=("arial", 10,'bold'), fg = 'blue',
                   bg= "ghost white", padx= 1, pady= 2, underline =0).grid(row=2,column=1, sticky =W , padx= 20, pady= 5)
        L2 = Label(DataFrameLEFT, text = "B: Daily In-Field Merge & QC Modules:", justify ="left", font=("arial", 10,'bold'), fg = 'blue',
                   bg= "ghost white", padx= 1, pady= 2, underline =0).grid(row=6,column=1, sticky =W , padx= 20, pady= 5)        
        btnImportSourceLink_TB = Button(DataFrameLEFT, text="SourceLink Vib TB Import", font=('aerial', 10, 'bold'),
                    bg= "ghost white" , height =1, width=24, bd=2, padx= 1, pady= 5, command = SourceLink_TB_LogIMPORT)
        btnImportSourceLink_TB.grid(row = 3, column=1, sticky =W , padx= 40, pady= 5)        
        btnImport_SourceLINK_PSS = Button(DataFrameLEFT, text="SourceLink Vib PSS Import", font=('aerial', 10, 'bold'),
                    bg= "ghost white" , height =1, width=24, bd=2,  padx= 1, pady= 5, command = SourceLink_PSS_LogIMPORT)
        btnImport_SourceLINK_PSS.grid(row=4,column=1, sticky =W , padx= 40, pady= 5)
        btnImport_SourceLINK_COG = Button(DataFrameLEFT, text="SourceLink Position Import", font=('aerial', 10, 'bold'),
                    bg= "ghost white" , height =1, width=24, bd=2,  padx= 1, pady= 5, command = SourceLink_Vib_PositionIMPORT)
        btnImport_SourceLINK_COG.grid(row=5,column=1, sticky =W , padx= 40, pady= 5)        
        btnMerge_SourceLINK_TB = Button(DataFrameLEFT, text="Merge TB - PSS - VibPosition", font=('aerial', 10, 'bold'),
                    bg= "ghost white" , height =1, width=24, bd=2,  padx= 1, pady= 5, command = Merge_SourceLINK_TB)
        btnMerge_SourceLINK_TB.grid(row=7,column=1 , sticky =W , padx= 40, pady= 5)

       ##  Define SourceLink VIB Clean TB and PSS submit to Master DB 
        DataFrameLEFT_CENTRE = Frame(self.root)
        DataFrameLEFT_CENTRE.grid(row=20,column=1, padx= 15, pady= 10)
        lblSubmitCleanTBMasterDB = Label(DataFrameLEFT_CENTRE, bd= 4, font=('aerial', 10, 'bold'),
                                         width = 40, fg = 'green', bg = "orange", underline =-1, text="ACCUMULATE VIB CLEAN TB AND MISSING SHOTS")
        lblSubmitCleanTBMasterDB.grid(row=1,column=1)
        L3 = Label(DataFrameLEFT_CENTRE, text = "A: Daily Clean Vib TB Submit To MasterDB:", justify ="left", font=("arial", 10,'bold'), fg = 'blue',
                   bg= "ghost white", padx= 1, pady= 2, underline =0).grid(row=2,column=1, sticky =W , padx= 20, pady= 5)
        L4 = Label(DataFrameLEFT_CENTRE, text = "B: Import Vib Source (Line & Station) File:", justify ="left", font=("arial", 10,'bold'), fg = 'blue',
                   bg= "ghost white", padx= 1, pady= 2, underline =0).grid(row=4,column=1, sticky =W , padx= 20, pady= 5)
        L5 = Label(DataFrameLEFT_CENTRE, text = "C: Generate Vib Daily Missing Shot Report:", justify ="left", font=("arial", 10,'bold'), fg = 'blue',
                   bg= "ghost white", padx= 1, pady= 2, underline =0).grid(row=7,column=1, sticky =W , padx= 20, pady= 5)
        btnMaster_SourceLINK_TB = Button(DataFrameLEFT_CENTRE, text="SourceLink Clean TB Import", font=('aerial', 10, 'bold'), bg= "ghost white" ,
                                         height =1, width=24, bd=2, padx= 1, pady= 5, command = SourceLink_Vibroseis_MasterTB_LogIMPORT)
        btnMaster_SourceLINK_TB.grid(row=3,column=1 , sticky =W , padx= 40, pady= 5)
        btnHelp_To_Make_SourceFile= Button(DataFrameLEFT_CENTRE, text="Help To Make Source File", font=('aerial', 10, 'bold'), bg= "ghost white" ,
                                           height =1, width=24, bd=2, padx= 1, pady= 5, command = Help_To_Make_SourceFile)
        btnHelp_To_Make_SourceFile.grid(row=5,column=1 , sticky =W , padx= 40, pady= 5)
        btnimportSourceSPS= Button(DataFrameLEFT_CENTRE, text="Vib Source (SPS) File Import", font=('aerial', 10, 'bold'), bg= "ghost white" ,
                                           height =1, width=24, bd=2, padx= 1, pady= 5, command = importVIBSourceSPS)
        btnimportSourceSPS.grid(row=6,column=1 , sticky =W , padx= 40, pady= 5)
        btnGenerateSourceLINK_TBMissingShot = Button(DataFrameLEFT_CENTRE, text="Generate Missing Shot Report", font=('aerial', 10, 'bold'), bg= "ghost white" ,
                                            height =1, width=24, bd=2, padx= 1, pady= 5, command = SourceLink_Vibroseis_TB_MissingShot)
        btnGenerateSourceLINK_TBMissingShot.grid(row=8,column=1 , sticky =W , padx= 40, pady= 5)


       ##  Define SourceLink Vib QC and Daily production Reporting 
        DataFrameMIDDLE = Frame(self.root)
        DataFrameMIDDLE.grid(row=2,column=2 , padx= 60, pady= 10)
        lblTitVibProductionReport = Label(DataFrameMIDDLE, bd= 4, font=('aerial', 10, 'bold'), width = 40, fg = 'green',
                    bg = "orange", underline =-1, text="VIB PRODUCTION REPORT MODULES")
        lblTitVibProductionReport.grid(row=1,column=1)
        L6 = Label(DataFrameMIDDLE, text = "A: Vib Production QC Daily - Sourcelink PSS:", justify ="left", font=("arial", 10,'bold'), fg = 'blue',
                   bg= "ghost white", padx= 1, pady= 2, underline =0).grid(row=2,column=1, sticky =W , padx= 20, pady= 5)
        L7 = Label(DataFrameMIDDLE, text = "B: Vib End Of Job Stat - Sourcelink PSS:", justify ="left", font=("arial", 10,'bold'), fg = 'blue',
                   bg= "ghost white", padx= 1, pady= 2, underline =0).grid(row=6,column=1, sticky =W , padx= 20, pady= 5)
        btnSetupVibQCLimit = Button(DataFrameMIDDLE, text="Setup Vib QC Parameter Limit", font=('aerial', 10, 'bold'),
                    bg= "ghost white" , height =1, width=26, bd=2, padx= 1, pady= 5, command = SetupVibQCLimitParameter)
        btnSetupVibQCLimit.grid(row = 3, column=1, sticky =W , padx= 40, pady= 5)
        btnHelpVibQCLimit = Button(DataFrameMIDDLE, text="Help", font=('aerial', 10, 'bold'),
                    bg= "ghost white" , height =1, width=6, bd=2, padx= 1, pady= 5, command = Help_VibQCLimitParameter)
        btnHelpVibQCLimit.grid(row = 3, column=1, sticky =E , padx= 5, pady= 5)        
        btnVibQCAnalysis = Button(DataFrameMIDDLE, text="PSS Analysis - Testif-i Report", font=('aerial', 10, 'bold'),
                    bg= "ghost white" , height =1, width=26, bd=2,  padx= 1, pady= 5, command = VibQCAnalysisTestiFyInput)
        btnVibQCAnalysis.grid(row=4,column=1, sticky =W , padx= 40, pady= 5)
        btnGenerateProductionQuality = Button(DataFrameMIDDLE, text="Generate Vib Production Report", font=('aerial', 10, 'bold'),
                    bg= "ghost white" , height =1, width=26, bd=2,  padx= 1, pady= 5, command = Generate_VibproductionQualityReport)
        btnGenerateProductionQuality.grid(row=5,column=1, sticky =W , padx= 40, pady= 5)        
        btnGenerateProductionStat = Button(DataFrameMIDDLE, text="Generate Vib Production Stat", font=('aerial', 10, 'bold'),
                    bg= "ghost white" , height =1, width=26, bd=2,  padx= 1, pady= 5, command = Generate_VibproductionStatReport)
        btnGenerateProductionStat.grid(row=7,column=1 , sticky =W , padx= 40, pady= 5)
        btnResetProductionStat = Button(DataFrameMIDDLE, text="Reset", font=('aerial', 10, 'bold'),
                    bg= "ghost white" , height =1, width=6, bd=2, padx= 1, pady= 5, command = Reset_VibproductionStatReport)
        btnResetProductionStat.grid(row = 7, column=1, sticky =E , padx= 5, pady= 5)

        ##  Define SourceLink Vib Miscellaneous Utility Modules
        DataFrameMIDDLE_CENTRE = Frame(self.root)
        DataFrameMIDDLE_CENTRE.grid(row=20,column=2, padx= 60, pady= 10)
        lblVibMiscUtility = Label(DataFrameMIDDLE_CENTRE, bd= 4, font=('aerial', 10, 'bold'), width = 40, fg = 'green',
                bg = "orange", underline =-1, text="MISCELLANEOUS UTILITY MODULES")
        lblVibMiscUtility.grid(row=1,column=1)
        L8 = Label(DataFrameMIDDLE_CENTRE, text = "A: Geomerge AUX Triggers - Sourcelink TB:", justify ="left", font=("arial", 10,'bold'), fg = 'blue',
                bg= "ghost white", padx= 1, pady= 2, underline =0).grid(row=2,column=1, sticky =W , padx= 20, pady= 5)

        L9 = Label(DataFrameMIDDLE_CENTRE, text = "B: Geomerge OutputLog Trace Yield Report:", justify ="left", font=("arial", 10,'bold'), fg = 'blue',
                bg= "ghost white", padx= 1, pady= 2, underline =0).grid(row=4,column=1, sticky =W , padx= 20, pady= 5)

        L10 = Label(DataFrameMIDDLE_CENTRE, text = "C: Vib SnailTrail Merge & GPX-CSV Convert:", justify ="left", font=("arial", 10,'bold'), fg = 'blue',
                bg= "ghost white", padx= 1, pady= 2, underline =0).grid(row=6,column=1, sticky =W , padx= 20, pady= 5)
           
        btnGenerateAUX_TB = Button(DataFrameMIDDLE_CENTRE, text="Generate Vib Signature TB", font=('aerial', 10, 'bold'),
                bg= "ghost white" , height =1, width=26, bd=2,  padx= 1, pady= 5, command = Generate_VibAUX_SignatureTB)
        btnGenerateAUX_TB.grid(row = 3, column=1, sticky =W , padx= 40, pady= 5)
        btnMaster_SourceLINK_PSS = Button(DataFrameMIDDLE_CENTRE, text="Generate Trace Yield Report", font=('aerial', 10, 'bold'),
                bg= "ghost white" ,height =1, width=26, bd=2,  padx= 1, pady= 5, command = GeomergeOutputLogTraceyield)
        btnMaster_SourceLINK_PSS.grid(row = 5, column=1, sticky =W , padx= 40, pady= 5)

        btnVib_SnailTrail_Merge = Button(DataFrameMIDDLE_CENTRE, text="Merge Vib-Observer Snail Trail", font=('aerial', 10, 'bold'),
                bg= "ghost white" , height =1, width=26, bd=2,  padx= 1, pady= 5, command = Vib_SnailTrail_MergeOBLog)
        btnVib_SnailTrail_Merge.grid(row = 7, column=1, sticky =W , padx= 40, pady= 5)        
        btnGPX_Fileconversion = Button(DataFrameMIDDLE_CENTRE, text="Garmin GPX- CSV File Convert", font=('aerial', 10, 'bold'),
                bg= "ghost white" , height =1, width=26, bd=2,  padx= 1, pady= 5, command = Under_Construction)
        btnGPX_Fileconversion.grid(row = 8, column=1, sticky =W , padx= 40, pady= 5)


        ##  Define SourceLink Dynamite Import and Merge TB - PFS - Drilling Log 
        DataFrameRIGHT = Frame(self.root)
        DataFrameRIGHT.grid(row=2,column=3 ,padx= 20, pady= 10)        
        lblTitDynamiteShotQC = Label(DataFrameRIGHT, bd= 4, font=('aerial', 10, 'bold'), width = 40, fg = 'green', bg = "orange", underline =-1, text="DAILY IN-FIELD DYNAMITE QC MODULES")
        lblTitDynamiteShotQC.grid(row=1,column=1)       
        L11 = Label(DataFrameRIGHT, text = "A: Daily In-Field Dynamite Import Modules:", justify ="left", font=("arial", 10,'bold'), fg = 'blue',
                   bg= "ghost white", padx= 1, pady= 2, underline =0).grid(row=2,column=1, sticky =W , padx= 20, pady= 5)
        L12 = Label(DataFrameRIGHT, text = "B: Daily In-Field Merge & QC Modules:", justify ="left", font=("arial", 10,'bold'), fg = 'blue',
                   bg= "ghost white", padx= 1, pady= 2, underline =0).grid(row=6,column=1, sticky =W , padx= 20, pady= 5)


        btnImportSourceLink_Dynamite_TB = Button(DataFrameRIGHT, text="SourceLink Dynamite TB Import", font=('aerial', 10, 'bold'),
                                bg= "ghost white" , height =1, width=27, bd=2, padx= 1, pady= 5, command = SourceLink_Dynamite_TB_LogIMPORT_Module)
        btnImportSourceLink_Dynamite_TB.grid(row = 3, column=1, sticky =W , padx= 40, pady= 5)

        btnImport_SourceLINK_PFS = Button(DataFrameRIGHT, text="SourceLink Dynamite PFS Import", font=('aerial', 10, 'bold'),
                                bg= "ghost white" , height =1, width=27, bd=2,  padx= 1, pady= 5, command = SourceLink_Dynamite_PFS_LogIMPORT_Module)
        btnImport_SourceLINK_PFS.grid(row=4,column=1, sticky =W , padx= 40, pady= 5)


        btnImport_DrillingLog = Button(DataFrameRIGHT, text="Dynamite Drilling Log Import", font=('aerial', 10, 'bold'),
                                bg= "ghost white" , height =1, width=27, bd=2,  padx= 1, pady= 5, command = Under_Construction)
        btnImport_DrillingLog.grid(row=5,column=1, sticky =W , padx= 40, pady= 5)

        btnMerge_SourceLINK_Dynamite_TB = Button(DataFrameRIGHT, text="Merge Dynamite TB - PFS Log", font=('aerial', 10, 'bold'),
                                bg= "ghost white" , height =1, width=27, bd=2,  padx= 1, pady= 5, command = SourceLink_Dynamite_TB_PFS_Merge_Module)
        btnMerge_SourceLINK_Dynamite_TB.grid(row=7,column=1 , sticky =W , padx= 40, pady= 5)


        ##  Define SourceLink Dynamite Clean TB and PFS submit to Master DB 
        DataFrameRIGHT_CENTRE = Frame(self.root)
        DataFrameRIGHT_CENTRE.grid(row=20,column=3, padx= 20, pady= 10)
        lblSubmitCleanTBDynDB = Label(DataFrameRIGHT_CENTRE, bd= 4, font=('aerial', 10, 'bold'), width = 40, fg = 'green',
                    bg = "orange", underline =-1, text="ACCUMULATE DYNAMITE TB AND MISSING SHOTS")
        lblSubmitCleanTBDynDB.grid(row=1,column=1)
        L13 = Label(DataFrameRIGHT_CENTRE, text = "A: Daily Dynamite TB Submit To MasterDB:", justify ="left", font=("arial", 10,'bold'), fg = 'blue',
                   bg= "ghost white", padx= 1, pady= 2, underline =0).grid(row=2,column=1, sticky =W , padx= 20, pady= 5)
        L14 = Label(DataFrameRIGHT_CENTRE, text = "B: Import Dynamite Source File:", justify ="left", font=("arial", 10,'bold'), fg = 'blue',
                   bg= "ghost white", padx= 1, pady= 2, underline =0).grid(row=4,column=1, sticky =W , padx= 20, pady= 5)
        L15 = Label(DataFrameRIGHT_CENTRE, text = "C: Generate Dynamite Missing Shot Report:", justify ="left", font=("arial", 10,'bold'), fg = 'blue',
                   bg= "ghost white", padx= 1, pady= 2, underline =0).grid(row=7,column=1, sticky =W , padx= 20, pady= 5)
   
        btnMaster_SourceLINK_Dynamite_TB = Button(DataFrameRIGHT_CENTRE, text="SourceLink Clean TB Import", font=('aerial', 10, 'bold'),
                                        bg= "ghost white" ,height =1, width=27, bd=2, padx= 1, pady= 5, command = SourceLink_Dynamite_Master_TB_LogIMPORT_Module)
        btnMaster_SourceLINK_Dynamite_TB.grid(row=3,column=1 , sticky =W , padx= 40, pady= 5)

        btnHelp_To_Make_SourceFile= Button(DataFrameRIGHT_CENTRE, text="Help To Make Source File", font=('aerial', 10, 'bold'),
                                    bg= "ghost white" ,height =1, width=27, bd=2, padx= 1, pady= 5, command = Help_To_Make_SourceFile)
        btnHelp_To_Make_SourceFile.grid(row=5,column=1 , sticky =W , padx= 40, pady= 5)
        btnimportSourceSPS= Button(DataFrameRIGHT_CENTRE, text="Dynamite Source File Import", font=('aerial', 10, 'bold'),
                                   bg= "ghost white" ,height =1, width=27, bd=2, padx= 1, pady= 5, command = importDynSourceSPS)
        btnimportSourceSPS.grid(row=6,column=1 , sticky =W , padx= 40, pady= 5)


        btnGeneSourceLINK_DynTBMissShot = Button(DataFrameRIGHT_CENTRE, text="Generate Missing Shot Report", font=('aerial', 10, 'bold'),
                                bg= "ghost white" ,height =1, width=27, bd=2, padx= 1, pady= 5, command = SourceLink_Dynamite_TB_MissingShot)
        btnGeneSourceLINK_DynTBMissShot.grid(row=8,column=1 , sticky =W , padx= 40, pady= 5)



if __name__ == '__main__':
    root = Tk()
    application  = SourceQCModule (root)
    root.mainloop()
