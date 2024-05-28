import sqlite3
#backend

def SOURCELINK_Dynamite_LogFiles():
    con= sqlite3.connect("SourceLink_Dynamite_Log.db")
    cur=con.cursor()

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_PFSLog_MASTER(DataBase_ID INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE NOT NULL, ShotID integer, FileNum integer, EPNumber integer, SourceLine integer,\
                SourceStation integer, Local_Date, Local_Time, Observer_Comment, ShotStatus, Timebreak, FirstBreak, Battery,\
                CapRes, GeoRes, FlagNum, UpholeWindow, FiredOK,\
                BatteryOK, GeoOK, CapOK, GPS_Quality, Unit_ID, TB_Date, TB_Time,\
                TB_Micro, CapSerialNumber, Latitude, Longitude, Altitude,\
                Encoder_Index, Record_Index integer, EP_Count, Crew_ID, \
                GPS_Time, GPS_Altitude, Sats, PDOP,\
                HDOP, VDOP, Age)")

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_PFSLog_TEMP(DataBase_ID INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE NOT NULL, ShotID integer, FileNum integer, EPNumber integer, SourceLine integer,\
                SourceStation integer, Local_Date, Local_Time, Observer_Comment, ShotStatus, Timebreak, FirstBreak, Battery,\
                CapRes, GeoRes, FlagNum, UpholeWindow, FiredOK,\
                BatteryOK, GeoOK, CapOK, GPS_Quality, Unit_ID, TB_Date, TB_Time,\
                TB_Micro, CapSerialNumber, Latitude, Longitude, Altitude,\
                Encoder_Index, Record_Index integer, EP_Count, Crew_ID, \
                GPS_Time, GPS_Altitude, Sats, PDOP,\
                HDOP, VDOP, Age)")
   

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_PFSLog_INVALID_NULL(DataBase_ID INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE NOT NULL, ShotID integer, FileNum integer, EPNumber integer, SourceLine integer,\
                SourceStation integer, Local_Date, Local_Time, Observer_Comment, ShotStatus, Timebreak, FirstBreak, Battery,\
                CapRes, GeoRes, FlagNum, UpholeWindow, FiredOK,\
                BatteryOK, GeoOK, CapOK, GPS_Quality, Unit_ID, TB_Date, TB_Time,\
                TB_Micro, CapSerialNumber, Latitude, Longitude, Altitude,\
                Encoder_Index, Record_Index integer, EP_Count, Crew_ID, \
                GPS_Time, GPS_Altitude, Sats, PDOP,\
                HDOP, VDOP, Age)")
    

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_PFSLog_DuplicatedShotID (DataBase_ID INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE NOT NULL, ShotID integer, FileNum integer, EPNumber integer, SourceLine integer,\
                SourceStation integer, Local_Date, Local_Time, Observer_Comment, ShotStatus, Timebreak, FirstBreak, Battery,\
                CapRes, GeoRes, FlagNum, UpholeWindow, FiredOK,\
                BatteryOK, GeoOK, CapOK, GPS_Quality, Unit_ID, TB_Date, TB_Time,\
                TB_Micro, CapSerialNumber, Latitude, Longitude, Altitude,\
                Encoder_Index, Record_Index integer, EP_Count, Crew_ID, \
                GPS_Time, GPS_Altitude, Sats, PDOP,\
                HDOP, VDOP, Age)")

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_PFSLog_VOID (DataBase_ID INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE NOT NULL, ShotID integer, FileNum integer, EPNumber integer, SourceLine integer,\
                SourceStation integer, Local_Date, Local_Time, Observer_Comment, ShotStatus, Timebreak, FirstBreak, Battery,\
                CapRes, GeoRes, FlagNum, UpholeWindow, FiredOK,\
                BatteryOK, GeoOK, CapOK, GPS_Quality, Unit_ID, TB_Date, TB_Time,\
                TB_Micro, CapSerialNumber, Latitude, Longitude, Altitude,\
                Encoder_Index, Record_Index integer, EP_Count, Crew_ID, \
                GPS_Time, GPS_Altitude, Sats, PDOP,\
                HDOP, VDOP, Age)")


    

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_SOURCELINK_DYNAMITE_TBMASTER(TriggerIndex integer, Unit_ID integer, FileNum integer, EPNumber integer,\
                 SourceLine integer, SourceStation integer,  ShotUtcDateTime,  Latitude,  Longitude,  ShotStatus,  Uphole, TBComment,  Process)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_SOURCELINK_TB_Dynamite(TriggerIndex integer, Unit_ID integer, FileNum integer, EPNumber integer,\
                 SourceLine integer, SourceStation integer,  ShotUtcDateTime,  Latitude,  Longitude,  ShotStatus,   Uphole, TBComment,  Process)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_SOURCELINK_TB_Dynamite_INVALID(TriggerIndex integer, Unit_ID integer, FileNum integer, EPNumber integer,\
                 SourceLine integer, SourceStation integer,  ShotUtcDateTime,  Latitude,  Longitude,  ShotStatus,  Uphole, TBComment,  Process)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_SOURCELINK_TB_Dynamite_Duplicated(TriggerIndex integer, Unit_ID integer, FileNum integer, EPNumber integer,\
                 SourceLine integer, SourceStation integer,  ShotUtcDateTime,  Latitude,  Longitude,  ShotStatus,  Uphole, TBComment,  Process)")
    
    con.commit()
    cur.close()
    con.close()


def addInvRec(TriggerIndex, Unit_ID, FileNum, EPNumber,SourceLine, SourceStation,  ShotUtcDateTime, Latitude,  Longitude,  ShotStatus,  Uphole, TBComment,  Process):
    con= sqlite3.connect("SourceLink_Dynamite_Log.db")
    cur=con.cursor()    
    cur.execute("INSERT INTO Eagle_SOURCELINK_DYNAMITE_TBMASTER VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",(TriggerIndex, Unit_ID, FileNum,
                EPNumber, SourceLine, SourceStation,  ShotUtcDateTime, Latitude, Longitude,  ShotStatus,  Uphole, TBComment,  Process))
    con.commit()
    con.close()


SOURCELINK_Dynamite_LogFiles()

