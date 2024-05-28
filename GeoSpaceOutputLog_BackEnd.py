import sqlite3
#backend

def GeoSpaceOutputLog():
    con= sqlite3.connect("GeoSpaceOutputLogTraceYield.db")
    cur=con.cursor()
    

    cur.execute("CREATE TABLE IF NOT EXISTS GeoSpaceOutputLog(FileNumber INTEGER,  ShotOverrides TEXT,  ProcessType TEXT,\
                  ShotLine BLOB,  ShotStation BLOB, Comment TEXT, NumSeisChannels INTEGER, NumMissChannels INTEGER, NumZeroChannels INTEGER, PercentMissing REAL)")
    cur.execute("CREATE TABLE IF NOT EXISTS GeoSpaceOutputLog_Duplicated(FileNumber INTEGER,  ShotOverrides TEXT,  ProcessType TEXT,\
                  ShotLine BLOB,  ShotStation BLOB, Comment TEXT, NumSeisChannels INTEGER, NumMissChannels INTEGER, NumZeroChannels INTEGER, PercentMissing REAL)")

    
        
    con.commit()
    cur.close()
    con.close()






GeoSpaceOutputLog()

