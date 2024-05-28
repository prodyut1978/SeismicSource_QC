#Front End
import os
from tkinter import*
import tkinter.messagebox
import tkinter.ttk as ttk
import tkinter as tk
from tkinter.filedialog import asksaveasfile
from tkinter.filedialog import askopenfilenames
from tkinter import simpledialog
import pandas as pd
import numpy as np
import openpyxl
import csv
import time
import datetime
from datetime import datetime
import pickle


def VibProductionQuality():
    if not os.path.exists("C:\VibRestrictedFolder\VibIFQC\VibQCLimitParameter"):
        THDAvg_Limit_Min    = 0
        THDAvg_Limit_Max    = 25

        THDMax_Limit_Min    = 0
        THDMax_Limit_Max    = 50

        ForceAvg_Limit_Min  = 30
        ForceAvg_Limit_Max  = 100

        ForceMax_Limit_Min  = 40
        ForceMax_Limit_Max  = 100

        PhaseAvg_Limit_Min  = -2
        PhaseAvg_Limit_Max  = 2

        PhaseMax_Limit_Min  = -10
        PhaseMax_Limit_Max  =  10
    else:
        Pickle_in       = open("C:\VibRestrictedFolder\VibIFQC\VibQCLimitParameter","rb")
        pickle_dict     = pickle.load(Pickle_in)

        THDAvg_Limit_Min    = pickle_dict[1]
        THDAvg_Limit_Max    = pickle_dict[2]

        THDMax_Limit_Min    = pickle_dict[3]
        THDMax_Limit_Max    = pickle_dict[4]

        ForceAvg_Limit_Min  = pickle_dict[5]
        ForceAvg_Limit_Max  = pickle_dict[6]

        ForceMax_Limit_Min  = pickle_dict[7]
        ForceMax_Limit_Max  = pickle_dict[8]

        PhaseAvg_Limit_Min  = pickle_dict[9]
        PhaseAvg_Limit_Max  = pickle_dict[10]
        
        PhaseMax_Limit_Min  = pickle_dict[11]
        PhaseMax_Limit_Max  = pickle_dict[12]


    Low_THDAvg_Limit     = float(THDAvg_Limit_Min)
    High_THDAvg_Limit    = float(THDAvg_Limit_Max)

    Low_THDMax_Limit     = float(THDMax_Limit_Min)
    High_THDMax_Limit    = float(THDMax_Limit_Max)
            
    Low_ForceAvg_Limit   = float(ForceAvg_Limit_Min)
    High_ForceAvg_Limit  = float(ForceAvg_Limit_Max)
            
    Low_ForceMax_Limit   = float(ForceMax_Limit_Min)
    High_ForceMax_Limit  = float(ForceMax_Limit_Max)

    Low_PhaseAvg_Limit  = float(PhaseAvg_Limit_Min)
    High_PhaseAvg_Limit = float(PhaseAvg_Limit_Max)
        
    Low_PhaseMax_Limit  = float(PhaseMax_Limit_Min)
    High_PhaseMax_Limit = float(PhaseMax_Limit_Max)

    PhaseMax_Limit      = float(PhaseMax_Limit_Max)
    PhaseMax_Limit      = abs(PhaseMax_Limit)

    PhaseMax_Limit_Set  = str(- PhaseMax_Limit)    + "   to   "    + str(PhaseMax_Limit)
    PhaseAvgLimit_Set   = str(Low_PhaseAvg_Limit)  + "   to   "    + str(High_PhaseAvg_Limit)

    ForceAvgLimit_Set    = str(Low_ForceAvg_Limit) + "   to   "    + str(High_ForceAvg_Limit)
    ForceMaxLimit_Set    = str(Low_ForceMax_Limit) + "   to   "    + str(High_ForceMax_Limit)

    THDAvgLimit_Set      = str(Low_THDAvg_Limit)   + "   to   "    + str(High_THDAvg_Limit)
    THDMaxLimit_Set      = str(Low_THDMax_Limit)   + "   to   "    + str(High_THDMax_Limit) 

    Default_Date_today   = "Date : " + datetime.now().strftime("%Y-%m-%d")
    Vib_path = r'C:\VIB_ProductionQuality_Report'
    if not os.path.exists(Vib_path):
        os.makedirs(Vib_path)
      
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
                    ProductionDayLocal    = df.loc[:,'Date']
                    LocalTime             = df.loc[:,'Time']
                    ObserverComment       = df.loc[:,'Comment']
                    ShotStatus            = df.loc[:,'Void']
                    PhaseMax              = df.loc[:,'Phase Max']
                    PhaseAvg              = df.loc[:,'Phase Avg']
                    ForceMax              = df.loc[:,'Force Max']
                    ForceAvg              = df.loc[:,'Force Avg']
                    THDMax                = df.loc[:,'THD Max']
                    THDAvg                = df.loc[:,'THD Avg']
                    SwCksm                = df.loc[:,'Sweep Checksum']
                    PmCksm                = df.loc[:,'Param Checksum']
                    GPSQuality            = df.loc[:,'Quality']
                    UnitID                = df.loc[:,'Unit ID']
                    Sats                  = df.loc[:,'Sats']
                    PDOP                  = df.loc[:,'PDOP']
                    HDOP                  = df.loc[:,'HDOP']
                    VDOP                  = df.loc[:,'VDOP']
                    Age                   = df.loc[:,'Age']
                    GPSAltitude           = df.loc[:,'GPS Altitude']
                    column_names = [ShotID, FileNum, EPNumber, SourceLine, SourceStation, ProductionDayLocal, LocalTime, ObserverComment, ShotStatus,
                                    PhaseMax, PhaseAvg, ForceMax, ForceAvg, THDMax, THDAvg, SwCksm, PmCksm, GPSQuality, UnitID, Sats, PDOP, HDOP, VDOP, Age, GPSAltitude]
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
                    ProductionDayLocal    = df.loc[:,'Date']
                    LocalTime             = df.loc[:,'Time']
                    ObserverComment       = df.loc[:,'Comment']
                    ShotStatus            = df.loc[:,'Void']
                    PhaseMax              = df.loc[:,'Phase Max']
                    PhaseAvg              = df.loc[:,'Phase Avg']
                    ForceMax              = df.loc[:,'Force Max']
                    ForceAvg              = df.loc[:,'Force Avg']
                    THDMax                = df.loc[:,'THD Max']
                    THDAvg                = df.loc[:,'THD Avg']
                    SwCksm                = df.loc[:,'Sweep Checksum']
                    PmCksm                = df.loc[:,'Param Checksum']
                    GPSQuality            = df.loc[:,'Quality']
                    UnitID                = df.loc[:,'Unit ID']
                    Sats                  = df.loc[:,'Sats']
                    PDOP                  = df.loc[:,'PDOP']
                    HDOP                  = df.loc[:,'HDOP']
                    VDOP                  = df.loc[:,'VDOP']
                    Age                   = df.loc[:,'Age']
                    GPSAltitude           = df.loc[:,'GPS Altitude']
                    column_names = [ShotID, FileNum, EPNumber, SourceLine, SourceStation, ProductionDayLocal, LocalTime, ObserverComment, ShotStatus,
                                    PhaseMax, PhaseAvg, ForceMax, ForceAvg, THDMax, THDAvg, SwCksm,PmCksm, GPSQuality, UnitID, Sats, PDOP, HDOP, VDOP, Age, GPSAltitude]
                    catdf = pd.concat (column_names,axis=1,ignore_index =True)
                    dfList.append(catdf) 

            concatDf = pd.concat(dfList,axis=0, ignore_index =True)
            concatDf.rename(columns={0:'ShotID', 1:'FileNum', 2:'EPNumber', 3:'SourceLine', 4:'SourceStation', 5:'ProductionDayLocal',
                             6:'LocalTime',7:'ObserverComment',8:'ShotStatus',9:'PhaseMax',10:'PhaseAvg',11:'ForceMax',
                             12:'ForceAvg',13:'THDMax',14:'THDAvg',15:'SwCksm',16:'PmCksm',17:'GPSQuality',
                             18:'UnitID', 19:'Sats', 20:'PDOP', 21:'HDOP', 22:'VDOP', 23:'Age', 24:'GPSAltitude'},inplace = True)
            
        # Separating Valid with Shot ID Not Null
        Valid_PSS_DF = pd.DataFrame(concatDf)
        Valid_PSS_DF = Valid_PSS_DF[pd.to_numeric(Valid_PSS_DF.ShotID, errors='coerce').notnull()]                  
        Valid_PSS_DF["SourceLine"].fillna(0, inplace = True)
        Valid_PSS_DF["SourceStation"].fillna(0, inplace = True)
        Valid_PSS_DF["FileNum"].fillna(0, inplace = True)
        Valid_PSS_DF["EPNumber"].fillna(1, inplace = True)
        Valid_PSS_DF["ProductionDayLocal"].fillna('1900/1/01', inplace = True)                 
        Valid_PSS_DF['SourceLine']             = (Valid_PSS_DF.loc[:,['SourceLine']]).astype(int)
        Valid_PSS_DF['SourceStation']          = (Valid_PSS_DF.loc[:,['SourceStation']]).astype(float)
        Valid_PSS_DF['ShotID']                 = (Valid_PSS_DF.loc[:,['ShotID']]).astype(int)
        Valid_PSS_DF['FileNum']                = (Valid_PSS_DF.loc[:,['FileNum']]).astype(int)
        Valid_PSS_DF['EPNumber']               = (Valid_PSS_DF.loc[:,['EPNumber']]).astype(int)
        Valid_PSS_DF['ProductionDayLocal']     = pd.to_datetime(Valid_PSS_DF['ProductionDayLocal']).dt.strftime('%Y/%m/%d')   
        Valid_PSS_DF['DuplicatedEntries']      = Valid_PSS_DF.sort_values(by =['ShotID', 'UnitID','SwCksm','PmCksm','PhaseMax','PhaseAvg','ForceMax','ForceAvg','THDMax','THDAvg']).duplicated(['ShotID'],keep='last')
        Valid_PSS_DF                           = Valid_PSS_DF.reset_index(drop=True)
        Valid_PSS_DF                           = pd.DataFrame(Valid_PSS_DF)

        # Separating Shot with Shot ID Not Duplicated And Not Null 
        DATA_VALID_PSS = Valid_PSS_DF.loc[Valid_PSS_DF.DuplicatedEntries == False, 'ShotID': 'GPSAltitude']
        DATA_VALID_PSS = DATA_VALID_PSS.reset_index(drop=True)
        DATA_VALID_PSS = pd.DataFrame(DATA_VALID_PSS)
        outfile_DATA_VALID_PSS =("C:\\VIB_ProductionQuality_Report\\DATA_VALID_PSS_Report.csv")
        DATA_VALID_PSS.to_csv(outfile_DATA_VALID_PSS,index=None)

        ## Getting Void Shot DataFrame
        VIB_Rep_VOID        = DATA_VALID_PSS[(DATA_VALID_PSS.ShotStatus.notnull())]
        VIB_Rep_VOID        = VIB_Rep_VOID.reset_index(drop=True)
        VIB_Rep_VOID_Count  = VIB_Rep_VOID.groupby('UnitID').ShotID.count()
        VIB_Rep_VOID_Count  = VIB_Rep_VOID_Count.reset_index(drop=False)
        VIB_Rep_VOID_Count.rename(columns = {'UnitID':'UnitID', 'ShotID':'Void Sweeps Count'}, inplace = True)

        ## Getting Not Void Shot DataFrame
        VIB_Rep_VALID_NOT_VOID    = DATA_VALID_PSS[(DATA_VALID_PSS.ShotStatus.isnull())]
        VIB_Rep_VALID_NOT_VOID    = VIB_Rep_VALID_NOT_VOID[pd.to_numeric(VIB_Rep_VALID_NOT_VOID.PhaseAvg,errors='coerce').notnull()]
        VIB_Rep_VALID_NOT_VOID    = VIB_Rep_VALID_NOT_VOID[pd.to_numeric(VIB_Rep_VALID_NOT_VOID.ForceAvg,errors='coerce').notnull()]
        VIB_Rep_VALID_NOT_VOID    = VIB_Rep_VALID_NOT_VOID[pd.to_numeric(VIB_Rep_VALID_NOT_VOID.THDAvg,errors='coerce').notnull()]
        VIB_Rep_VALID_NOT_VOID    = VIB_Rep_VALID_NOT_VOID[pd.to_numeric(VIB_Rep_VALID_NOT_VOID.THDMax,errors='coerce').notnull()]
        VIB_Rep_VALID_NOT_VOID    = VIB_Rep_VALID_NOT_VOID[pd.to_numeric(VIB_Rep_VALID_NOT_VOID.PhaseMax,errors='coerce').notnull()]
        VIB_Rep_VALID_NOT_VOID    = VIB_Rep_VALID_NOT_VOID.reset_index(drop=True)

        ## Getting QC Passed Shot Dataframe    
        VIB_Rep_QC_Passed   =  pd.DataFrame(VIB_Rep_VALID_NOT_VOID)
        VIB_Rep_QC_Passed['PhaseMax'] = VIB_Rep_QC_Passed['PhaseMax'].abs()
        VIB_Rep_QC_Passed   =  VIB_Rep_QC_Passed[(VIB_Rep_QC_Passed.PhaseMax <= PhaseMax_Limit)&
                                             (VIB_Rep_QC_Passed.PhaseAvg <= High_PhaseAvg_Limit)&
                                             (VIB_Rep_QC_Passed.PhaseAvg >= Low_PhaseAvg_Limit)&
                                             (VIB_Rep_QC_Passed.ForceMax <= High_ForceMax_Limit)&
                                             (VIB_Rep_QC_Passed.ForceMax >= Low_ForceMax_Limit)&                                     
                                             (VIB_Rep_QC_Passed.ForceAvg <= High_ForceAvg_Limit)&
                                             (VIB_Rep_QC_Passed.ForceAvg >= Low_ForceAvg_Limit)&
                                             (VIB_Rep_QC_Passed.THDAvg   <= High_THDAvg_Limit)&
                                             (VIB_Rep_QC_Passed.THDMax   <= High_THDMax_Limit)&  
                                             (VIB_Rep_QC_Passed.GPSQuality != "No Fix")]
        VIB_Rep_QC_Passed        = VIB_Rep_QC_Passed.reset_index(drop=True)
        VIB_Rep_QC_Passed_Count  = VIB_Rep_QC_Passed.groupby('UnitID').agg({'ProductionDayLocal':lambda x : x.iloc[0], 
                                                                            'ShotID' :'count'})
        VIB_Rep_QC_Passed_Count  = VIB_Rep_QC_Passed_Count.reset_index(drop=False)
        VIB_Rep_QC_Passed_Count.rename(columns = {'UnitID':'UnitID', 'ProductionDayLocal': 'Production Day', 'ShotID':'QC Passed Sweeps Count'}, inplace = True)

        ## Getting QC Failed Shot Dataframe    
        VIB_Rep_QC_Failed   =  pd.DataFrame(VIB_Rep_VALID_NOT_VOID)
        VIB_Rep_QC_Failed['PhaseMax'] = VIB_Rep_QC_Failed['PhaseMax'].abs()
        VIB_Rep_QC_Failed   =  VIB_Rep_QC_Failed[(VIB_Rep_QC_Failed.PhaseMax > PhaseMax_Limit)|
                                      (VIB_Rep_QC_Failed.ForceAvg > High_ForceAvg_Limit)|
                                      (VIB_Rep_QC_Failed.ForceAvg < Low_ForceAvg_Limit)|
                                      (VIB_Rep_QC_Failed.THDAvg   > High_THDAvg_Limit)|
                                      (VIB_Rep_QC_Failed.THDMax   > High_THDMax_Limit)|
                                      (VIB_Rep_QC_Failed.GPSQuality == "No Fix")|                              
                                      (VIB_Rep_QC_Failed.PhaseAvg > High_PhaseAvg_Limit)|
                                      (VIB_Rep_QC_Failed.PhaseAvg < Low_PhaseAvg_Limit)|
                                      (VIB_Rep_QC_Failed.ForceMax > High_ForceMax_Limit)|
                                      (VIB_Rep_QC_Failed.ForceMax < Low_ForceMax_Limit)]
        VIB_Rep_QC_Failed  = VIB_Rep_QC_Failed.reset_index(drop=True)
        VIB_Rep_Fail_QC    =  pd.DataFrame(VIB_Rep_QC_Failed)

        def trans_PhaseMax_Limit(x):
            if x > PhaseMax_Limit:
                return 'PhaseMax_Limit_Fail'
            else:
                return '(PhaseMax_OK)'

        def trans_PhaseAvg_Limit(q):
            if q > High_PhaseAvg_Limit:
                return 'High_PhaseAvg_Limit_Fail'
            elif q < Low_PhaseAvg_Limit:
                return 'Low_PhaseAvg_Limit_Fail'
            else:
                return '(PhaseAvg_OK)'

        def trans_ForceMax_Limit(p):
            if p > High_ForceMax_Limit:
                return 'High_ForceMax_Limit_Fail'
            elif p < Low_ForceMax_Limit:
                return 'Low_ForceMax_Limit_Fail'
            else:
                return '(ForceMax_OK)'

        def trans_ForceAvg_Limit(y):
            if y > High_ForceAvg_Limit:
                return 'High_ForceAvg_Limit_Fail'
            elif y < Low_ForceAvg_Limit:
                return 'Low_ForceAvg_Limit_Fail'
            else:
                return '(ForceAvg_OK)'

        def trans_THDAvg_Limit(z):
            if z > High_THDAvg_Limit:
                return 'High_THDAvg_Limit_Fail'
            else:
                return '(THDAvg_OK)'  

        def trans_THDMax_Limit(u):
            if u > High_THDMax_Limit:
                return 'High_THDMax_Limit_Fail'
            else:
                return '(THDMax_OK)'  

        def trans_GPS_Quality(w):
            if w == "No Fix":
                return 'GPS_NoFIX'
            else:
                return '(GPS_OK)'

        VIB_Rep_Fail_QC['PhaseMax_Limit_FLAG']  = VIB_Rep_Fail_QC['PhaseMax'].apply(trans_PhaseMax_Limit)
        VIB_Rep_Fail_QC['PhaseMax_Limit_FLAG']  = VIB_Rep_Fail_QC.PhaseMax_Limit_FLAG.astype (object)

        VIB_Rep_Fail_QC['PhaseAvg_Limit_FLAG']  = VIB_Rep_Fail_QC['PhaseAvg'].apply(trans_PhaseAvg_Limit)
        VIB_Rep_Fail_QC['PhaseAvg_Limit_FLAG']  = VIB_Rep_Fail_QC.PhaseAvg_Limit_FLAG.astype (object)

        VIB_Rep_Fail_QC['ForceMax_Limit_FLAG']  = VIB_Rep_Fail_QC['ForceMax'].apply(trans_ForceMax_Limit)
        VIB_Rep_Fail_QC['ForceMax_Limit_FLAG']  = VIB_Rep_Fail_QC.ForceMax_Limit_FLAG.astype (object)

        VIB_Rep_Fail_QC['ForceAvg_Limit_FLAG']  = VIB_Rep_Fail_QC['ForceAvg'].apply(trans_ForceAvg_Limit)
        VIB_Rep_Fail_QC['ForceAvg_Limit_FLAG']  = VIB_Rep_Fail_QC.ForceAvg_Limit_FLAG.astype (object)

        VIB_Rep_Fail_QC['THDAvg_Limit_FLAG']    = VIB_Rep_Fail_QC['THDAvg'].apply(trans_THDAvg_Limit)
        VIB_Rep_Fail_QC['THDAvg_Limit_FLAG']    = VIB_Rep_Fail_QC.THDAvg_Limit_FLAG.astype (object)

        VIB_Rep_Fail_QC['THDMax_Limit_FLAG']    = VIB_Rep_Fail_QC['THDMax'].apply(trans_THDMax_Limit)
        VIB_Rep_Fail_QC['THDMax_Limit_FLAG']    = VIB_Rep_Fail_QC.THDMax_Limit_FLAG.astype (object)

        VIB_Rep_Fail_QC['GPS_FLAG']             = VIB_Rep_Fail_QC['GPSQuality'].apply(trans_GPS_Quality)
        VIB_Rep_Fail_QC['GPS_FLAG']             = VIB_Rep_Fail_QC.GPS_FLAG.astype (object)

        VIB_Rep_Fail_QC    =  pd.DataFrame(VIB_Rep_Fail_QC)
        VIB_Rep_Fail_QC    = VIB_Rep_Fail_QC.reset_index(drop=True)

        outfile_VIB_Rep_Fail_QC =("C:\\VIB_ProductionQuality_Report\\VIB_Rep_Fail_QC.csv")
        VIB_Rep_Fail_QC.to_csv(outfile_VIB_Rep_Fail_QC,index=None)

        # Generating Vib Fail_QC_Statistics
        VIB_Rep_Fail_QC_Stat       = pd.DataFrame(VIB_Rep_Fail_QC)

        PhaseMax_Fail_Count        = VIB_Rep_Fail_QC_Stat[(VIB_Rep_Fail_QC_Stat.PhaseMax_Limit_FLAG=='PhaseMax_Limit_Fail')]
        PhaseMax_Limit_Fail_Count  = PhaseMax_Fail_Count.groupby('UnitID').ShotID.count()

        PhaseAvg_Fail_Count        = VIB_Rep_Fail_QC_Stat[(VIB_Rep_Fail_QC_Stat.PhaseAvg_Limit_FLAG =='High_PhaseAvg_Limit_Fail')|
                                                          (VIB_Rep_Fail_QC_Stat.PhaseAvg_Limit_FLAG =='Low_PhaseAvg_Limit_Fail')]
        PhaseAvg_Limit_Fail_Count  = PhaseAvg_Fail_Count.groupby('UnitID').ShotID.count()

        Force_Max_Fail_Count       = VIB_Rep_Fail_QC_Stat[(VIB_Rep_Fail_QC_Stat.ForceMax_Limit_FLAG =='High_ForceMax_Limit_Fail')|
                                                          (VIB_Rep_Fail_QC_Stat.ForceMax_Limit_FLAG =='Low_ForceMax_Limit_Fail')]
        Force_Max_Limit_Fail_Count = Force_Max_Fail_Count.groupby('UnitID').ShotID.count()

        Force_Avg_Fail_Count       = VIB_Rep_Fail_QC_Stat[(VIB_Rep_Fail_QC_Stat.ForceAvg_Limit_FLAG =='High_ForceAvg_Limit_Fail')|
                                                          (VIB_Rep_Fail_QC_Stat.ForceAvg_Limit_FLAG =='Low_ForceAvg_Limit_Fail')]
        Force_Avg_Limit_Fail_Count = Force_Avg_Fail_Count.groupby('UnitID').ShotID.count()

        THDMax_Fail_Count        =  VIB_Rep_Fail_QC_Stat[(VIB_Rep_Fail_QC_Stat.THDMax_Limit_FLAG=='High_THDMax_Limit_Fail')]
        THDMax_Limit_Fail_Count  =  THDMax_Fail_Count.groupby('UnitID').ShotID.count()
        THDAvg_Fail_Count        =  VIB_Rep_Fail_QC_Stat[(VIB_Rep_Fail_QC_Stat.THDAvg_Limit_FLAG=='High_THDAvg_Limit_Fail')]
        THDAvg_Limit_Fail_Count  =  THDAvg_Fail_Count.groupby('UnitID').ShotID.count()

        GPS_Fail_Count           =  VIB_Rep_Fail_QC_Stat[(VIB_Rep_Fail_QC_Stat.GPS_FLAG=='GPS_NoFIX')]
        GPS_Limit_Fail_Count     =  GPS_Fail_Count.groupby('UnitID').ShotID.count()

        Start_Day                =   VIB_Rep_Fail_QC_Stat.groupby('UnitID').ProductionDayLocal.nth(0)
        Failed_Sweeps_Per_vib    =   VIB_Rep_Fail_QC_Stat.groupby('UnitID').ShotID.count()

        VibFailure_Stat  = [Start_Day, Failed_Sweeps_Per_vib,
                            PhaseMax_Limit_Fail_Count,  PhaseAvg_Limit_Fail_Count,
                            Force_Max_Limit_Fail_Count, Force_Avg_Limit_Fail_Count,
                            THDMax_Limit_Fail_Count,    THDAvg_Limit_Fail_Count,
                            GPS_Limit_Fail_Count]
        Vib_Failure_Stat = pd.concat (VibFailure_Stat,axis=1,ignore_index =True)
        Vib_Failure_Stat.rename(columns = {0:'Start Date',   1:'FailedCount',
                                           2:'PhaseMaxFail', 3:'PhaseAvgFail',
                                           4:'ForceMaxFail', 5:'ForceAvgFail',                                   
                                           6:'THDMaxFail',   7:'THDAvgFail',
                                           8:'GPS_NoFix'},inplace = True)
        Vib_Failure_Stat    = pd.DataFrame(Vib_Failure_Stat)
        Vib_Failure_Stat    = Vib_Failure_Stat.reset_index(drop=False)
        

        ProductionCount = VIB_Rep_VALID_NOT_VOID.groupby(['UnitID']).agg({ 'ShotID'   :'count'})
        ProductionCount.rename(columns = {'ShotID'  :'TotalShots'},inplace = True)
        ProductionCount    = ProductionCount.reset_index(drop=False)
        

        Vib_Failure_Stat    = pd.merge(Vib_Failure_Stat, ProductionCount,
                                    how='outer', on = ['UnitID'])
        Vib_Failure_Stat    = Vib_Failure_Stat.fillna(0)
        Vib_Failure_Stat    = Vib_Failure_Stat.sort_values('UnitID')
        
        Vib_Failure_Stat['%GPS_NoFix'] = (((Vib_Failure_Stat['GPS_NoFix'])/(Vib_Failure_Stat['TotalShots']))*100).round(1)

        Vib_Failure_Stat    = Vib_Failure_Stat.loc[:,[ 'UnitID', 'TotalShots', 'FailedCount',                                                            
                                                       'PhaseAvgFail', 'PhaseMaxFail',
                                                       'THDAvgFail', 'THDMaxFail',
                                                       'ForceAvgFail', 'ForceMaxFail', 'GPS_NoFix', '%GPS_NoFix', 'Start Date']]

        Vib_Failure_Stat_Quantity    = pd.DataFrame(Vib_Failure_Stat)
        Vib_Failure_Stat_Percent     = pd.DataFrame(Vib_Failure_Stat)

        Vib_Failure_Stat_Percent['%PhaseAvgFail'] = (((Vib_Failure_Stat_Percent['PhaseAvgFail'])/(Vib_Failure_Stat_Percent['TotalShots']))*100).round(1)
        Vib_Failure_Stat_Percent['%PhaseMaxFail'] = (((Vib_Failure_Stat_Percent['PhaseMaxFail'])/(Vib_Failure_Stat_Percent['TotalShots']))*100).round(1)

        Vib_Failure_Stat_Percent['%THDAvgFail'] = (((Vib_Failure_Stat_Percent['THDAvgFail'])/(Vib_Failure_Stat_Percent['TotalShots']))*100).round(1)
        Vib_Failure_Stat_Percent['%THDMaxFail'] = (((Vib_Failure_Stat_Percent['THDMaxFail'])/(Vib_Failure_Stat_Percent['TotalShots']))*100).round(1)

        Vib_Failure_Stat_Percent['%ForceAvgFail'] = (((Vib_Failure_Stat_Percent['ForceAvgFail'])/(Vib_Failure_Stat_Percent['TotalShots']))*100).round(1)
        Vib_Failure_Stat_Percent['%ForceMaxFail'] = (((Vib_Failure_Stat_Percent['ForceMaxFail'])/(Vib_Failure_Stat_Percent['TotalShots']))*100).round(1)
        Vib_Failure_Stat_Percent    = Vib_Failure_Stat_Percent.loc[:,[ 'UnitID', 'TotalShots', 'FailedCount',                                                         
                                                                       '%PhaseAvgFail', '%PhaseMaxFail',
                                                                       '%THDAvgFail', '%THDMaxFail',
                                                                       '%ForceAvgFail', '%ForceMaxFail']]
        
        ##  Vib Failure Quantity 
        Failure_Per_Vib_Stat    = pd.DataFrame(Vib_Failure_Stat_Quantity)
        Failure_Per_Vib_Stat    = Failure_Per_Vib_Stat.reset_index(drop=True)
        Failure_Per_Vib_Stat    = Failure_Per_Vib_Stat.loc[:,['UnitID', 'TotalShots', 'FailedCount',                                                            
                                                              'PhaseAvgFail', 'PhaseMaxFail',
                                                              'THDAvgFail', 'THDMaxFail',
                                                              'ForceAvgFail', 'ForceMaxFail']]
        Failure_Per_Vib_Stat['PhaseAvgFail']  = (Failure_Per_Vib_Stat.loc[:,['PhaseAvgFail']]).astype(int)
        Failure_Per_Vib_Stat['PhaseMaxFail']  = (Failure_Per_Vib_Stat.loc[:,['PhaseMaxFail']]).astype(int)

        Failure_Per_Vib_Stat['THDAvgFail']    = (Failure_Per_Vib_Stat.loc[:,['THDAvgFail']]).astype(int)
        Failure_Per_Vib_Stat['THDMaxFail']    = (Failure_Per_Vib_Stat.loc[:,['THDMaxFail']]).astype(int)

        Failure_Per_Vib_Stat['ForceAvgFail']  = (Failure_Per_Vib_Stat.loc[:,['ForceAvgFail']]).astype(int)
        Failure_Per_Vib_Stat['ForceMaxFail']  = (Failure_Per_Vib_Stat.loc[:,['ForceMaxFail']]).astype(int)    
        Failure_Per_Vib_Stat    = Failure_Per_Vib_Stat.fillna(0)
        Failure_Per_Vib_Stat    = Failure_Per_Vib_Stat.sort_values('UnitID')
        Length_Vib_Stat         = len(Failure_Per_Vib_Stat)

        ##  Vib Failure Percentage
        Failure_Per_Vib_percent    = pd.DataFrame(Vib_Failure_Stat_Percent)
        Failure_Per_Vib_percent    = Failure_Per_Vib_percent.reset_index(drop=True)
        Failure_Per_Vib_percent    = Failure_Per_Vib_percent.fillna(0)
        Failure_Per_Vib_percent    = Failure_Per_Vib_percent.sort_values('UnitID')
        Length_Vib_percent         = len(Failure_Per_Vib_percent)

        ##  Vib GPS Positional QC
        GPS_NoFix_Vib_Stat    = pd.DataFrame(Vib_Failure_Stat)
        GPS_NoFix_Vib_Stat    = GPS_NoFix_Vib_Stat.reset_index(drop=False)
        GPS_NoFix_Vib_Stat    = GPS_NoFix_Vib_Stat.loc[:,['UnitID', 'FailedCount',                                                            
                                                          'GPS_NoFix', '%GPS_NoFix']]
        GPS_NoFix_Vib_Stat    = GPS_NoFix_Vib_Stat.fillna(0)
        GPS_NoFix_Vib_Stat    = GPS_NoFix_Vib_Stat.sort_values('UnitID')
        Vib_GPS_PositionQC    = pd.DataFrame(VIB_Rep_VALID_NOT_VOID)    
        Vib_GPS_PositionQC    = Vib_GPS_PositionQC.groupby(['UnitID']).agg({'Sats' :'mean',                        
                                                                            'PDOP' :'mean',
                                                                            'HDOP' :'mean',
                                                                            'VDOP' :'mean',
                                                                            'Age'  :'mean',
                                                                            'ShotID'   :'count'})    
        Vib_GPS_PositionQC.rename(columns = {'ShotID'  :'TotalShots',
                                             'Sats'    :'Mean-Sats#',
                                             'PDOP'    :'Mean-PDOP',
                                             'HDOP'    :'Mean-HDOP',
                                             'VDOP'    :'Mean-VDOP',
                                             'Age'     :'Mean-Age'},inplace = True)
        Vib_GPS_PositionQC    = Vib_GPS_PositionQC.reset_index(drop=False)
        GPS_NoFix_Vib_Stat    = pd.merge(GPS_NoFix_Vib_Stat, Vib_GPS_PositionQC,
                                    how='outer', on = ['UnitID'])
        GPS_NoFix_Vib_Stat    = GPS_NoFix_Vib_Stat.loc[:,[ 'UnitID', 'TotalShots', 'FailedCount',                                                            
                                                           'GPS_NoFix','Mean-Sats#', 'Mean-PDOP',
                                                           'Mean-HDOP', 'Mean-VDOP', 'Mean-Age']]
        
        GPS_NoFix_Vib_Stat    = GPS_NoFix_Vib_Stat.fillna(0)
        GPS_NoFix_Vib_Stat    = GPS_NoFix_Vib_Stat.sort_values('UnitID', ascending=True)
        
        GPS_NoFix_Vib_Stat['Mean-Sats#']     = (GPS_NoFix_Vib_Stat.loc[:,['Mean-Sats#']]).round(0)
        GPS_NoFix_Vib_Stat['Mean-Sats#']     = (GPS_NoFix_Vib_Stat.loc[:,['Mean-Sats#']]).astype(int)
        GPS_NoFix_Vib_Stat['Mean-PDOP']      = (GPS_NoFix_Vib_Stat.loc[:,['Mean-PDOP']]).round(1)
        GPS_NoFix_Vib_Stat['Mean-HDOP']      = (GPS_NoFix_Vib_Stat.loc[:,['Mean-HDOP']]).round(1)
        GPS_NoFix_Vib_Stat['Mean-VDOP']      = (GPS_NoFix_Vib_Stat.loc[:,['Mean-VDOP']]).round(1)
        GPS_NoFix_Vib_Stat['Mean-Age']       = (GPS_NoFix_Vib_Stat.loc[:,['Mean-Age']]).round(1)
        GPS_NoFix_Vib_Stat = pd.DataFrame(GPS_NoFix_Vib_Stat)
        Length_GPS_NoFix      = len(GPS_NoFix_Vib_Stat)

        ##  Vib GPS Quality
        Vib_GPS_QC    = pd.DataFrame(VIB_Rep_VALID_NOT_VOID)
        Vib_GPS_QC    = Vib_GPS_QC.reset_index(drop=False)
        Vib_GPS_QC    = Vib_GPS_QC.groupby(['GPSQuality']).agg({'Sats' :'mean',
                                                                'PDOP' :'mean',
                                                                'HDOP' :'mean',
                                                                'VDOP' :'mean',
                                                                'Age' :'mean',
                                                                'GPSAltitude':'mean',
                                                                'ShotID'   :'count'})
        Vib_GPS_QC.rename(columns = {'GPSQuality': 'GPSQuality', 'Sats' : 'Mean-Sats#',
                                     'PDOP' : 'Mean-PDOP','HDOP' : 'Mean-HDOP','VDOP' : 'Mean-VDOP','Age' : 'Mean-Age','GPSAltitude' : 'Mean-Altitude',
                                     'ShotID'  :'TotalShots'},inplace = True)

        Vib_GPS_QC['Percentage (%)'] = (((Vib_GPS_QC['TotalShots'])/((Vib_GPS_QC['TotalShots']).sum()))*100).round(1)
        Vib_GPS_QC['Mean-Sats#']     = (Vib_GPS_QC.loc[:,['Mean-Sats#']]).round(0)
        Vib_GPS_QC['Mean-Sats#']     = (Vib_GPS_QC.loc[:,['Mean-Sats#']]).astype(int)
        Vib_GPS_QC['Mean-PDOP']      = (Vib_GPS_QC.loc[:,['Mean-PDOP']]).round(1)
        Vib_GPS_QC['Mean-HDOP']      = (Vib_GPS_QC.loc[:,['Mean-HDOP']]).round(1)
        Vib_GPS_QC['Mean-VDOP']      = (Vib_GPS_QC.loc[:,['Mean-VDOP']]).round(1)
        Vib_GPS_QC['Mean-Age']       = (Vib_GPS_QC.loc[:,['Mean-Age']]).round(1)
        Vib_GPS_QC['Mean-Altitude']  = (Vib_GPS_QC.loc[:,['Mean-Altitude']]).round(2)    
        Vib_GPS_QC    = Vib_GPS_QC.reset_index(drop=False)
        Vib_GPS_QC    = Vib_GPS_QC.sort_values('TotalShots', ascending=False)
        Length_Vib_GPS_QC      = len(Vib_GPS_QC)
        
        outfile_Failure_Per_Vib_Stat =("C:\\VIB_ProductionQuality_Report\\Number of Failure_Per_Vib_Stat.csv")
        Vib_Failure_Stat.to_csv(outfile_Failure_Per_Vib_Stat,index=None)

        ## Generating Vib Production Quality Report
        ProductionQualityReport = VIB_Rep_VALID_NOT_VOID.groupby(['UnitID']).agg({'ProductionDayLocal':lambda x : x.iloc[0], 
                                                                 'ShotID'   :'count',
                                                                 'PhaseMax' :'mean', 
                                                                 'PhaseAvg' :'mean', 
                                                                 'ForceMax' :'mean', 
                                                                 'ForceAvg' :'mean',
                                                                 'THDMax'   :'mean',
                                                                 'THDAvg'   :'mean'})
        ProductionQualityReport.rename(columns = {'ProductionDayLocal': 'Start Date',
                                         'ShotID'  :'TotalShots',
                                         'PhaseMax':'PhaseMaxMean',
                                         'PhaseAvg':'PhaseAvgMean',
                                         'ForceMax':'ForceMaxMean',
                                         'ForceAvg':'ForceAvgMean',
                                         'THDMax'  :'THDMaxMean',
                                         'THDAvg'  :'THDAvgMean'},inplace = True)
        ProductionQualityReport    = ProductionQualityReport.reset_index(drop=False)
        ProductionQualityReport    = ProductionQualityReport.loc[:,['Start Date','UnitID','TotalShots', 
                                                               'PhaseAvgMean','PhaseMaxMean',
                                                               'THDAvgMean', 'THDMaxMean',
                                                               'ForceAvgMean', 'ForceMaxMean']]
        ProductionQualityReport = ProductionQualityReport.round(1)
        Length_ProductionQuality = len(ProductionQualityReport)    
        outfile_ProductionQualityReport =("C:\\VIB_ProductionQuality_Report\\ProductionQualityReport.csv")
        ProductionQualityReport.to_csv(outfile_ProductionQualityReport,index=None)

        ## Generating Vib Production Summary Report
        ProductionDay_Start  = DATA_VALID_PSS['ProductionDayLocal'].iloc[0]
        ProductionDay_End  = DATA_VALID_PSS['ProductionDayLocal'].iloc[-1]
        ProductionDay_StartTime  = DATA_VALID_PSS['LocalTime'].iloc[0]
        ProductionDay_EndTime  = DATA_VALID_PSS['LocalTime'].iloc[-1]
        ProductionStart_DateTime = pd.to_datetime (ProductionDay_Start + ' ' + ProductionDay_StartTime)
        ProductionEnd_DateTime = pd.to_datetime (ProductionDay_End + ' ' + ProductionDay_EndTime)
        production_DURATION    =  (ProductionEnd_DateTime - ProductionStart_DateTime)/(np.timedelta64(1,'h'))
        production_DURATION    =  round(production_DURATION,1)
        TotalShot      = len(DATA_VALID_PSS)
        VoidShot       = len(VIB_Rep_VOID)
        ProductionShot = TotalShot - VoidShot
        QCPassedShots  = len(VIB_Rep_QC_Passed)
        QCFailedShots  = len(VIB_Rep_QC_Failed)
        PercentDailyQCPassed   = round(100*(QCPassedShots)/(ProductionShot),1)
        PercentDailyQCFailed   = 100- PercentDailyQCPassed

        Production_Summary = pd.DataFrame({'Production Start':[ProductionDay_Start],
                                            'Production End':[ProductionDay_End], 
                                            'Total Shots':[TotalShot],
                                            'VOID Shots':[VoidShot],
                                            'Production Shot':[ProductionShot],
                                            'QC Passed Shots':[QCPassedShots],
                                            'QC Failed Shots':[QCFailedShots],
                                            '% Daily QC Pass':[PercentDailyQCPassed],
                                            '% Daily QC Fail':[PercentDailyQCFailed]
                                           },index=None)

        NumberofVibes = DATA_VALID_PSS['UnitID'].nunique()
        

        ## Exporting Generated Report
        def get_VibProductionSummary_Rep_datetime():
            return "Vib Production Quality Report -" + datetime.now().strftime("%Y%m%d-%H%M%S") + ".xlsx"

        filename               = tkinter.filedialog.asksaveasfilename(initialdir = "/",
                                 title = "Save Vib Production Quality Report As Excel",
                                 filetypes = (("Excel file",".xlsx"),("Excel file",".xlsx")))
        if len(filename) >0:
            VibProductionSummary   = get_VibProductionSummary_Rep_datetime()
            outfile_VibProductionSummary  = filename + VibProductionSummary
            XLSX_writer = pd.ExcelWriter(outfile_VibProductionSummary)
            Production_Summary.to_excel(XLSX_writer,'VibProductionQuality', startrow = 4 ,  index=False)
            ProductionQualityReport.to_excel(XLSX_writer,'VibProductionQuality', startrow = 28,  index=False)
            Failure_Per_Vib_Stat.to_excel(XLSX_writer,'VibProductionQuality', startrow = 28 +Length_ProductionQuality+2,  index=False)
            Failure_Per_Vib_percent.to_excel(XLSX_writer,'VibProductionQuality', startrow = 28 +Length_ProductionQuality+Length_Vib_Stat+4,  index=False)
            GPS_NoFix_Vib_Stat.to_excel(XLSX_writer,'VibProductionQuality', startrow = 28 +Length_ProductionQuality+Length_Vib_Stat+Length_Vib_percent+6,  index=False)
            Vib_GPS_QC.to_excel(XLSX_writer,'VibProductionQuality', startrow = 28 +Length_ProductionQuality+Length_Vib_Stat+Length_Vib_percent+ Length_GPS_NoFix+8, startcol =0, index=False)
            workbook       = XLSX_writer.book
            worksheet_VibProductionQuality  = XLSX_writer.sheets['VibProductionQuality']        
            header1  = '&L&G'+'&CEagle Canada Seismic Services ULC' + '\n' + '6806 Railway Street SE' + '\n' + 'Calgary, AB T2H 3A8' + '\n' + 'Ph: (403) 263-7770' +  '&R&U&14&"cambria, bold"Vib Production Report' + '\n'+  'Date : &D'
            worksheet_VibProductionQuality.set_header(header1,{'image_left':'eagle logo.jpg'})
            footer1  = ('&LDate : &D')        
            worksheet_VibProductionQuality.set_footer(footer1)        
            worksheet_VibProductionQuality.set_landscape()
            worksheet_VibProductionQuality.set_margins(0.6, 0.6, 1.6, 1.1)                                   
            worksheet_VibProductionQuality.set_paper(9)
            worksheet_VibProductionQuality.set_start_page(1)
            worksheet_VibProductionQuality.hide_gridlines(1)
            worksheet_VibProductionQuality.set_page_view()
            workbook.formats[0].set_align('center')
            workbook.formats[0].set_font_size(11)
            workbook.formats[0].set_bold(True)
            workbook.formats[0].set_border(1)
            worksheet_VibProductionQuality.set_column(0, 0, 15)   
            worksheet_VibProductionQuality.set_column(1, 1, 14)
            worksheet_VibProductionQuality.set_column(2, 2, 11)
            worksheet_VibProductionQuality.set_column(3, 3, 14)
            worksheet_VibProductionQuality.set_column(4, 7, 14)
            worksheet_VibProductionQuality.set_column(8, 9, 13)
            
            cell_format_Left = workbook.add_format({
                                                    'bold': True,
                                                    'text_wrap': False,
                                                    'valign': 'top',
                                                    'border': 1})
            cell_format_Left.set_align('left')
            cell_format_Left.set_font_size(11)
            cell_format_Center = workbook.add_format({
                                                    'bold': True,
                                                    'text_wrap': False,
                                                    'valign': 'top',
                                                    'border': 1})
            cell_format_Center.set_align('center')
            cell_format_Center.set_font_size(11)
            cell_format_Header = workbook.add_format({
                                                    'bold': True,
                                                    'text_wrap': False,
                                                    'valign': 'top',
                                                    'border': 1})
            cell_format_Header.set_align('left')
            cell_format_Header.set_font_size(12)
            cell_format_Header.set_underline(1)
            cell_format_Footnote = workbook.add_format({
                                                    'bold': True,
                                                    'text_wrap': False,
                                                    'valign': 'top',
                                                    'border': 1})
            cell_format_Footnote.set_align('left')
            cell_format_Footnote.set_font_size(12)
                
            ##  Vib QC Parameters Specifications
            worksheet_VibProductionQuality.merge_range('A1:I1', " Vib QC Parameters Specifications :", cell_format_Header)
            worksheet_VibProductionQuality.write('A2', " PhaseMax (Deg) :", cell_format_Left)
            worksheet_VibProductionQuality.write('A3', " PhaseAvg (Deg) :", cell_format_Left)
            worksheet_VibProductionQuality.write('D2', " ForceMax (%) :", cell_format_Left)
            worksheet_VibProductionQuality.write('D3', " ForceAvg (%) :", cell_format_Left)
            worksheet_VibProductionQuality.write('G2', " THDMax (%):", cell_format_Left)
            worksheet_VibProductionQuality.write('G3', " THDAvg (%) :", cell_format_Left)
            worksheet_VibProductionQuality.merge_range('B2:C2', PhaseMax_Limit_Set, cell_format_Center)
            worksheet_VibProductionQuality.merge_range('B3:C3', PhaseAvgLimit_Set, cell_format_Center)
            worksheet_VibProductionQuality.merge_range('E2:F2', ForceMaxLimit_Set, cell_format_Center)
            worksheet_VibProductionQuality.merge_range('E3:F3', ForceAvgLimit_Set, cell_format_Center)
            worksheet_VibProductionQuality.merge_range('H2:I2', THDMaxLimit_Set, cell_format_Center)
            worksheet_VibProductionQuality.merge_range('H3:I3', THDAvgLimit_Set, cell_format_Center)
            
            ##  Vib Production Overview        
            worksheet_VibProductionQuality.merge_range('A4:I4', " Vib Production Overview :", cell_format_Header)
            productionOverview_chart_Bar = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})
            productionOverview_chart_Pie = workbook.add_chart({'type': 'pie'})

            productionOverview_chart_Bar.add_series({  
                        'categories': ['VibProductionQuality', 4, 2, 4, 6],    
                        'values':     ['VibProductionQuality', 5, 2, 5, 6], 'gap': 50})
            productionOverview_chart_Bar.set_size({'width': 500, 'height': 353})                                  
            productionOverview_chart_Bar.set_y_axis({
                'name': 'Number Of Shots',
                'name_font': {'size': 10, 'bold': True},
                'num_font':  {'size': 8, 'bold': True, 'rotation': - 45},
                'major_gridlines': {
                              'visible': True,
                              'line': {'width': 1.25, 'dash_type': 'dash'}},})                        
            productionOverview_chart_Bar.set_x_axis({'name': 'Vib Shot Classification', 'num_font':  {'size': 8, 'bold': True, 'rotation': - 45}, 'name_font': {'size': 12, 'bold': True},
                              'major_gridlines': {
                              'visible': True,
                              'line': {'width': 1.25, 'dash_type': 'dash'}},})
            productionOverview_chart_Bar.set_style(10)
            productionOverview_chart_Bar.set_legend({'none': True}) 
            worksheet_VibProductionQuality.insert_chart('A7', productionOverview_chart_Bar,  
                            {'x_offset': 1, 'y_offset': 10})
            productionOverview_chart_Pie.add_series({   
                        'categories': ['VibProductionQuality', 4, 7, 4, 8],    
                        'values':     ['VibProductionQuality', 5, 7, 5, 8],
                        'data_labels': {'category': True, 'position': 'best_fit'}}) 
            productionOverview_chart_Pie.set_title({'name': 'Vib Production Overview', 'name_font': {'size': 12, 'bold': True},'overlay': True, 'layout': {'x': 0.70, 'y': 0.03}})
            productionOverview_chart_Pie.set_size({'width': 404, 'height': 353})                                              
            productionOverview_chart_Pie.set_style(10)
            productionOverview_chart_Pie.set_rotation(90)                    
            worksheet_VibProductionQuality.insert_chart('F7', productionOverview_chart_Pie,  
                            {'x_offset': 0, 'y_offset': 10})

            worksheet_VibProductionQuality.merge_range('A26:C26', " Vib Production Start DateTime :", cell_format_Footnote)
            worksheet_VibProductionQuality.merge_range('D26:E26', str(ProductionStart_DateTime), cell_format_Footnote)
            worksheet_VibProductionQuality.merge_range('A27:C27', " Vib Production End DateTime  :", cell_format_Footnote)
            worksheet_VibProductionQuality.merge_range('D27:E27', str(ProductionEnd_DateTime), cell_format_Footnote)
            worksheet_VibProductionQuality.merge_range('F26:H26', " Number Of Vibs In Production :", cell_format_Footnote)
            worksheet_VibProductionQuality.write('I26',NumberofVibes, cell_format_Footnote)
            worksheet_VibProductionQuality.merge_range('F27:H27', " Number Of Hours In Production :", cell_format_Footnote)
            worksheet_VibProductionQuality.write('I27',production_DURATION, cell_format_Footnote)
            
             ##  Vib Production Quality
            worksheet_VibProductionQuality.merge_range('A28:I28', " Vib Production Quality (***) :", cell_format_Header)

             ##  Vib Failure Stat
            worksheet_VibProductionQuality.merge_range((28 +Length_ProductionQuality+1), 0, (28 +Length_ProductionQuality+1), 8, " Vib QC Parameters Failure Quantity (#) :", cell_format_Header)

            ##  Vib Failure %%
            worksheet_VibProductionQuality.merge_range((28 +Length_ProductionQuality+Length_Vib_Stat+3), 0, (28 +Length_ProductionQuality+Length_Vib_Stat+3), 8, " Vib QC Parameters Failure Percentage (%) :", cell_format_Header)

            ##   Vib GPS Positional QC
            worksheet_VibProductionQuality.merge_range((28 +Length_ProductionQuality+Length_Vib_Stat+Length_Vib_percent+5), 0, (28 +Length_ProductionQuality+Length_Vib_Stat+Length_Vib_percent+5), 8,
                                                       " Vib GPS Positional QC :", cell_format_Header)

            ##  Vib GPS Quality
            worksheet_VibProductionQuality.merge_range((28 +Length_ProductionQuality+Length_Vib_Stat+Length_Vib_percent+Length_GPS_NoFix+7), 0, (28 +Length_ProductionQuality+Length_Vib_Stat+Length_Vib_percent+Length_GPS_NoFix+7), 8,
                                                       " Vib Overall GPS Quality :", cell_format_Header)
            XLSX_writer.save()
            XLSX_writer.close()
            tkinter.messagebox.showinfo("Vib Production Report Export Message"," Vib Production Report Saved As Excel")
        else:
            tkinter.messagebox.showinfo("Vib Production Report Export Message","Please Select Vib Production Report File Name To Save")
    else:
        tkinter.messagebox.showinfo("SourceLink PSS File Import Message","Please Select PSS Files From Imported Folder")






