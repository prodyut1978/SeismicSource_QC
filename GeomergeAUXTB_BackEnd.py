import sqlite3
#backend

def GeomergeAUX_TB():
    con= sqlite3.connect("GeomergeAUXTB.db")
    cur=con.cursor()
    

    cur.execute("CREATE TABLE IF NOT EXISTS Geomerge_AUXTB_MERGED(TriggerIndex integer,  ProfileId integer,  ShotNumber integer, EpNumber integer,\
                  ShotLine integer,  ShotStation integer,   ShotUtcDateTime,  Latitude,  Longitude,  ShotStatus,   Comment, DeviceType integer, AUXUnitNumber integer)")

    cur.execute("CREATE TABLE IF NOT EXISTS Geomerge_AUXTB_TEMP(TriggerIndex integer,  ProfileId integer,  ShotNumber integer, EpNumber integer,\
                  ShotLine integer,  ShotStation integer,   ShotUtcDateTime,  Latitude,  Longitude,  ShotStatus,   Comment)")

    cur.execute("CREATE TABLE IF NOT EXISTS AUXBOX_Profile(ProfileId integer, AUXUnitNumber integer, DeviceType)")
        
    con.commit()
    cur.close()
    con.close()


def addInvRec(ProfileId, AUXUnitNumber, DeviceType):
    con= sqlite3.connect("GeomergeAUXTB.db")
    cur=con.cursor()    
    cur.execute("INSERT INTO AUXBOX_Profile VALUES (?,?,?)",(ProfileId, AUXUnitNumber, DeviceType))
    con.commit()
    con.close()



GeomergeAUX_TB()

