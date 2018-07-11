Attribute VB_Name = "mGlobals"
Option Explicit

Global Const CurrentDBase = "c:\osiris\projects\Health Log\HealthLog.mdb"

Global dbase As Database
Global UserAction As String
Global SupLoaded As Boolean
Global NeedToClearSup As Boolean
Global num_rows As Long

' Find the next available number in a given field of a given table.
' REQUIRES: table (table to search)
'           field (field to search)
' RETURNS:  next available number, or -1 if none are available below maxlimit
Public Function FindFreeID(db As Database, table As String, field As String) As Long
    Dim record As Recordset
    Dim record_count As Long

    Set record = db.OpenRecordset("SELECT COUNT (*) AS [Count] FROM " & _
        table, dbOpenDynaset)
    record_count = record!Count
    record.Close
    
    If record_count = 0 Then
        FindFreeID = 1
    Else
        Set record = db.OpenRecordset("SELECT Max(" & field & ") AS [MaxID] FROM " _
            & table, dbOpenDynaset)
        FindFreeID = record!MaxID + 1
        record.Close
    End If
End Function


