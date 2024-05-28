import sqlite3
#backend

def SOURCELINK_PSSLog_Analysis_QC():
    con= sqlite3.connect("SourceLink_PSSLog_Analysis_QC.db")
    cur=con.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_PSSLog_MASTER(DataBase_ID INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE NOT NULL, ShotID integer, FileNum integer, EPNumber integer, SourceLine integer,\
                SourceStation integer, Local_Date, Local_Time, Observer_Comment, ShotStatus, PhaseMax,\
                PhaseAvg, ForceMax, ForceAvg, THDMax, THDAvg,\
                SwCksm, PmCksm, GPS_Quality, Unit_ID, TB_Date, TB_Time,\
                TB_Micro, Signature_File_Number, Latitude, Longitude, Altitude,\
                Encoder_Index, Record_Index integer, EP_Count, Crew_ID, Start_Code,\
                Force_Out, GPS_Time, GPS_Altitude, Sats, PDOP,\
                HDOP, VDOP, Age, Start_Time_Delta, Sweep_Number)")

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_PSSLog_RAWDUMP(DataBase_ID INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE NOT NULL, ShotID integer, FileNum integer, EPNumber integer, SourceLine integer,\
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
    

    
    


    
    con.commit()
    cur.close()
    con.close()



SOURCELINK_PSSLog_Analysis_QC()

