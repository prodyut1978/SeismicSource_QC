import sqlite3
#backend

def Global_Mapper_SnailTrails_LogFiles():
    con= sqlite3.connect("Global_Mapper_SnailTrails_Log.db")
    cur=con.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS Global_Mapper_SnailTrails_OBLog(UnitID INTEGER, Lat, Lon, Elevation,\
                Status, GPSQuality_Number,  GPSQuality_Verbal, Satellites, UTCDate,  UTCTime, LocalDate,  LocalTime)")

    cur.execute("CREATE TABLE IF NOT EXISTS Global_Mapper_SnailTrails_Vib(UnitID INTEGER, Lat, Lon, Elevation,\
                Status, GPSQuality_Number,  GPSQuality_Verbal, Satellites, UTCDate,  UTCTime, LocalDate,  LocalTime )")

    cur.execute("CREATE TABLE IF NOT EXISTS Global_Mapper_Merged_SnailTrails(UnitID INTEGER, Lat, Lon, Elevation,\
                Status, GPSQuality_Number,  GPSQuality_Verbal, Satellites, UTCDate,  UTCTime, LocalDate,  LocalTime, MergeFlags )")
   

        
    con.commit()
    cur.close()
    con.close()



Global_Mapper_SnailTrails_LogFiles()

