import sqlite3
#backend

def HAWKOBlogImport():
    con= sqlite3.connect("HAWK_OBLog.db")
    cur=con.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_HAWK_OBLog_MASTER(MasterSystemFieldRecordID integer, EPNumber integer, ShotID integer, Omit boolean, FileType text, File_CorrBeforeStack, File_CorrAfterStack,\
                File_UncorrStack, File_CorrEP, File_UncorrEP, Timebreak_SecondUnixTimeStamp , Timebreak_mSecs, Timebreak_uSecs ,RecordLength_mSecs, Acquisition_Time_mSecs, SourceLine integer,\
                SourceStation integer, SourceType_DynamiteorVibroseis text, Vibes64_bit_mask, SampleRateuSecs, SourceX, SourceY, SourceZ, GridUnits, SweepFile, SweepID, SweepType,\
                SweepStartFrequency, SweepEndFrequency, SweepLength, TaperType,\
                StartTaperDuration, EndTaperDuration, Comment)")
    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_HAWK_OBLog_TEMP(MasterSystemFieldRecordID integer, EPNumber integer, ShotID integer, Omit, FileType, File_CorrBeforeStack, File_CorrAfterStack,\
                File_UncorrStack, File_CorrEP, File_UncorrEP, Timebreak_SecondUnixTimeStamp , Timebreak_mSecs, Timebreak_uSecs ,RecordLength_mSecs, Acquisition_Time_mSecs, SourceLine integer,\
                SourceStation integer, SourceType_DynamiteorVibroseis, Vibes64_bit_mask, SampleRateuSecs, SourceX, SourceY, SourceZ, GridUnits, SweepFile, SweepID, SweepType,\
                SweepStartFrequency, SweepEndFrequency, SweepLength, TaperType,\
                StartTaperDuration, EndTaperDuration, Comment)")

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_HAWK_OBLog_INVALID(MasterSystemFieldRecordID integer, EPNumber integer, ShotID integer, Omit, FileType, File_CorrBeforeStack, File_CorrAfterStack,\
                File_UncorrStack, File_CorrEP, File_UncorrEP, Timebreak_SecondUnixTimeStamp , Timebreak_mSecs, Timebreak_uSecs ,RecordLength_mSecs, Acquisition_Time_mSecs, SourceLine integer,\
                SourceStation integer, SourceType_DynamiteorVibroseis, Vibes64_bit_mask, SampleRateuSecs, SourceX, SourceY, SourceZ, GridUnits, SweepFile, SweepID, SweepType,\
                SweepStartFrequency, SweepEndFrequency, SweepLength, TaperType,\
                StartTaperDuration, EndTaperDuration, Comment)")

    cur.execute("CREATE TABLE IF NOT EXISTS Eagle_HAWK_OBLog_DUPLICATESHOTID(MasterSystemFieldRecordID integer, EPNumber integer, ShotID integer, Omit, FileType, File_CorrBeforeStack, File_CorrAfterStack,\
                File_UncorrStack, File_CorrEP, File_UncorrEP, Timebreak_SecondUnixTimeStamp , Timebreak_mSecs, Timebreak_uSecs ,RecordLength_mSecs, Acquisition_Time_mSecs, SourceLine integer,\
                SourceStation integer, SourceType_DynamiteorVibroseis, Vibes64_bit_mask, SampleRateuSecs, SourceX, SourceY, SourceZ, GridUnits, SweepFile, SweepID, SweepType,\
                SweepStartFrequency, SweepEndFrequency, SweepLength, TaperType,\
                StartTaperDuration, EndTaperDuration, Comment)")



   

    
    con.commit()
    con.close()



HAWKOBlogImport()
