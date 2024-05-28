import sqlite3
#backend
def SOURCE_Files():
    con= sqlite3.connect("SourceSPS.db")
    cur=con.cursor()    
    cur.execute("CREATE TABLE IF NOT EXISTS SourceFileSPS(SourceLine INTEGER, SourceStation REAL,  SourceLineStationCombined REAL)")    
    con.commit()
    cur.close()
    con.close()



SOURCE_Files()

