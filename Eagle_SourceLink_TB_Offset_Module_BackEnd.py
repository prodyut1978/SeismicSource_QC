import sqlite3
#backend

def SOURCELINK_Microseconds_Offset():
    con= sqlite3.connect("SourceLink_Microseconds_Offset.db")
    cur=con.cursor()
    
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_SOURCELINKTB_Microseconds_Offset_Vib(TriggerIndex integer, Unit_ID integer, FileNum integer, EPNumber integer,\
                 SourceLine integer, SourceStation integer,  ShotUtcDateTime,  Latitude,  Longitude,  ShotStatus,  TBComment)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_SOURCELINK_TB_TEMP_Vib(TriggerIndex integer, Unit_ID integer, FileNum integer, EPNumber integer,\
                 SourceLine integer, SourceStation integer,  ShotUtcDateTime,  Latitude,  Longitude,  ShotStatus,  TBComment)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_SOURCELINK_TB_INVALID_Vib(TriggerIndex integer, Unit_ID integer, FileNum integer, EPNumber integer,\
                 SourceLine integer, SourceStation integer,  ShotUtcDateTime,  Latitude,  Longitude,  ShotStatus,  TBComment)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_SOURCELINK_TB_Duplicated_Vib(TriggerIndex integer, Unit_ID integer, FileNum integer, EPNumber integer,\
                 SourceLine integer, SourceStation integer,  ShotUtcDateTime,  Latitude,  Longitude,  ShotStatus,  TBComment)")
    


    
    con.commit()
    cur.close()
    con.close()



SOURCELINK_Microseconds_Offset()

