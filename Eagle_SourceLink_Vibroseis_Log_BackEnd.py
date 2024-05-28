import sqlite3
#backend

def SOURCELINK_LogFiles():
    con= sqlite3.connect("SourceLink_Log.db")
    cur=con.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_PSSLog_MASTER(DataBase_ID INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE NOT NULL, ShotID integer, FileNum integer, EPNumber integer, SourceLine integer,\
                SourceStation integer, Local_Date, Local_Time, Observer_Comment, ShotStatus, PhaseMax,\
                PhaseAvg, ForceMax, ForceAvg, THDMax, THDAvg,\
                SwCksm, PmCksm, GPS_Quality, Unit_ID, TB_Date, TB_Time,\
                TB_Micro, Signature_File_Number, Latitude, Longitude, Altitude,\
                Encoder_Index, Record_Index integer, EP_Count, Crew_ID, Start_Code,\
                Force_Out, GPS_Time, GPS_Altitude, Sats, PDOP,\
                HDOP, VDOP, Age, Start_Time_Delta, Sweep_Number)")

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_PSSLog_TEMP(DataBase_ID INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE NOT NULL, ShotID integer, FileNum integer, EPNumber integer, SourceLine integer,\
                SourceStation integer, Local_Date, Local_Time, Observer_Comment, ShotStatus, PhaseMax,\
                PhaseAvg, ForceMax, ForceAvg, THDMax, THDAvg,\
                SwCksm, PmCksm, GPS_Quality, Unit_ID, TB_Date, TB_Time,\
                TB_Micro, Signature_File_Number, Latitude, Longitude, Altitude,\
                Encoder_Index, Record_Index integer, EP_Count, Crew_ID, Start_Code,\
                Force_Out, GPS_Time, GPS_Altitude, Sats, PDOP,\
                HDOP, VDOP, Age, Start_Time_Delta, Sweep_Number)")
   

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_PSSLog_INVALID_NULL(DataBase_ID INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE NOT NULL, ShotID integer, FileNum integer, EPNumber integer, SourceLine integer,\
                SourceStation integer, Local_Date, Local_Time, Observer_Comment, ShotStatus, PhaseMax,\
                PhaseAvg, ForceMax, ForceAvg, THDMax, THDAvg,\
                SwCksm, PmCksm, GPS_Quality, Unit_ID, TB_Date, TB_Time,\
                TB_Micro, Signature_File_Number, Latitude, Longitude, Altitude,\
                Encoder_Index, Record_Index integer, EP_Count, Crew_ID, Start_Code,\
                Force_Out, GPS_Time, GPS_Altitude, Sats, PDOP,\
                HDOP, VDOP, Age, Start_Time_Delta, Sweep_Number)")
    

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_PSSLog_DuplicatedShotID (DataBase_ID INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE NOT NULL, ShotID integer, FileNum integer, EPNumber integer, SourceLine integer,\
                SourceStation integer, Local_Date, Local_Time, Observer_Comment, ShotStatus, PhaseMax,\
                PhaseAvg, ForceMax, ForceAvg, THDMax, THDAvg,\
                SwCksm, PmCksm, GPS_Quality, Unit_ID, TB_Date, TB_Time,\
                TB_Micro, Signature_File_Number, Latitude, Longitude, Altitude,\
                Encoder_Index, Record_Index integer, EP_Count, Crew_ID, Start_Code,\
                Force_Out, GPS_Time, GPS_Altitude, Sats, PDOP,\
                HDOP, VDOP, Age, Start_Time_Delta, Sweep_Number)")

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_PSSLog_VOID (DataBase_ID INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE NOT NULL, ShotID integer, FileNum integer, EPNumber integer, SourceLine integer,\
                SourceStation integer, Local_Date, Local_Time, Observer_Comment, ShotStatus, PhaseMax,\
                PhaseAvg, ForceMax, ForceAvg, THDMax, THDAvg,\
                SwCksm, PmCksm, GPS_Quality, Unit_ID, TB_Date, TB_Time,\
                TB_Micro, Signature_File_Number, Latitude, Longitude, Altitude,\
                Encoder_Index, Record_Index integer, EP_Count, Crew_ID, Start_Code,\
                Force_Out, GPS_Time, GPS_Altitude, Sats, PDOP,\
                HDOP, VDOP, Age, Start_Time_Delta, Sweep_Number)")

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_PSSLog_QCPassed (DataBase_ID INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE NOT NULL, ShotID integer, FileNum integer, EPNumber integer, SourceLine integer,\
                SourceStation integer, Local_Date, Local_Time, Observer_Comment, ShotStatus, PhaseMax,\
                PhaseAvg, ForceMax, ForceAvg, THDMax, THDAvg,\
                SwCksm, PmCksm, GPS_Quality, Unit_ID, TB_Date, TB_Time,\
                TB_Micro, Signature_File_Number, Latitude, Longitude, Altitude,\
                Encoder_Index, Record_Index integer, EP_Count, Crew_ID, Start_Code,\
                Force_Out, GPS_Time, GPS_Altitude, Sats, PDOP,\
                HDOP, VDOP, Age, Start_Time_Delta, Sweep_Number)")

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_PSSLog_QCFailed (DataBase_ID INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE NOT NULL, ShotID integer, FileNum integer, EPNumber integer, SourceLine integer,\
                SourceStation integer, Local_Date, Local_Time, Observer_Comment, ShotStatus, PhaseMax,\
                PhaseAvg, ForceMax, ForceAvg, THDMax, THDAvg,\
                SwCksm, PmCksm, GPS_Quality, Unit_ID, TB_Date, TB_Time,\
                TB_Micro, Signature_File_Number, Latitude, Longitude, Altitude,\
                Encoder_Index, Record_Index integer, EP_Count, Crew_ID, Start_Code,\
                Force_Out, GPS_Time, GPS_Altitude, Sats, PDOP,\
                HDOP, VDOP, Age, Start_Time_Delta, Sweep_Number)")
    

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_VIB_COG_MASTER(FileNum integer, ShotID integer, EPNumber integer, SourceLine integer, SourceStation integer, DistanceCOG, NearFlagLine, NearFlagStation, DistanceNearFlag,\
                GPS_Quality, Unit_ID, Near_Flag_Message text)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_VIB_COG_TEMP(FileNum integer, ShotID integer, EPNumber integer, SourceLine integer, SourceStation integer, DistanceCOG, NearFlagLine, NearFlagStation, DistanceNearFlag,\
                GPS_Quality, Unit_ID, Near_Flag_Message text)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_VIB_COG_INVALID(FileNum integer, ShotID integer, EPNumber integer, SourceLine integer, SourceStation integer, DistanceCOG, NearFlagLine, NearFlagStation, DistanceNearFlag,\
                GPS_Quality, Unit_ID)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_VIB_COG_DUPLICATEDSHOTID(FileNum integer, ShotID integer, EPNumber integer, SourceLine integer, SourceStation integer, DistanceCOG, NearFlagLine, NearFlagStation, DistanceNearFlag,\
                GPS_Quality, Unit_ID)")

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_SOURCELINKTBMASTER(TriggerIndex integer, Unit_ID integer, FileNum integer, EPNumber integer,\
                 SourceLine integer, SourceStation integer,  ShotUtcDateTime,  Latitude,  Longitude,  ShotStatus,  TBComment, Process, SweepNumber)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_SOURCELINK_TB_TEMP(TriggerIndex integer, Unit_ID integer, FileNum integer, EPNumber integer,\
                 SourceLine integer, SourceStation integer,  ShotUtcDateTime,  Latitude,  Longitude,  ShotStatus,  TBComment)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_SOURCELINK_TB_INVALID(TriggerIndex integer, Unit_ID integer, FileNum integer, EPNumber integer,\
                 SourceLine integer, SourceStation integer,  ShotUtcDateTime,  Latitude,  Longitude,  ShotStatus,  TBComment)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_SOURCELINK_TB_Duplicated(TriggerIndex integer, Unit_ID integer, FileNum integer, EPNumber integer,\
                 SourceLine integer, SourceStation integer,  ShotUtcDateTime,  Latitude,  Longitude,  ShotStatus,  TBComment)")
    


    
    con.commit()
    cur.close()
    con.close()

def addInvRec(TriggerIndex, Unit_ID, FileNum, EPNumber, SourceLine, SourceStation, ShotUtcDateTime, Latitude, Longitude, ShotStatus, TBComment, Process, SweepNumber):
    con= sqlite3.connect("SourceLink_Log.db")
    cur=con.cursor()    
    cur.execute("INSERT INTO Eagle_SOURCELINKTBMASTER VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",(TriggerIndex, Unit_ID, FileNum,
                EPNumber, SourceLine, SourceStation,  ShotUtcDateTime, Latitude, Longitude,  ShotStatus, TBComment,  Process, SweepNumber))
    con.commit()
    con.close()

SOURCELINK_LogFiles()

