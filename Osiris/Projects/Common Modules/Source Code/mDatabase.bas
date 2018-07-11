Attribute VB_Name = "mDatabase"
'This module requires the following references to exist in the project:
'   Microsoft DAO Library (tested with version 3.51)
'   Microsoft Access Objects Library (tested with version 8.0)

'the following constant needs to be defined globally in your project.
'#Const ProgressBar = False  'Allows use of progress bars for lengthy operations.
                            'if this is set to True, the Progress Bar module
                            'must be included in the project.

Option Explicit

Const BLOCKSIZE = 32768 'used for reading/writing memo fields and Large Binary
                        'Objects.  Value is in bytes.


Public Function OpenDBase(databasefile As String) As Database
    Dim db As Database
    
    If GetAttr(databasefile) And vbReadOnly Then
        SetAttr databasefile, vbNormal
        MsgBox "The database '" & databasefile & _
            "' was marked 'Read Only'.  It was reset to 'Normal' to allow " _
            & "normal functionality.", vbInformation
    End If
    
    On Error GoTo Err_OpenDB
    Set db = OpenDatabase(databasefile)
    On Error GoTo 0
    Set OpenDBase = db
    GoTo Done

Err_OpenDB:
    MsgBox "OpenDBase:  OpenDatabase unsuccessful!"
    Set OpenDBase = Nothing
    
Done:
End Function

' Find the next available number in a given field of a given table.
' REQUIRES: table (table to search)
'           field (field to search)
Public Function FindFreeID(db As Database, table As String, field As String) As Long
    Dim record As Recordset

    FindFreeID = 1
    Set record = db.OpenRecordset("SELECT " & field & " FROM " & table & _
            " ORDER BY " & field, dbOpenDynaset)
    If Not record.EOF Then
        If record(field) <> 1 Then
            GoTo Done
        End If
        FindFreeID = record(field) + 1
        record.MoveNext
        While Not record.EOF
            If record(field) <> FindFreeID Then
                GoTo Done
            End If
            FindFreeID = FindFreeID + 1
            record.MoveNext
        Wend
    End If
Done:
    record.Close
End Function


Public Sub CopyTable(db As Database, source_table As String, target_table As String)
    Dim tdfSource As TableDef
    Dim tdfTarget As TableDef
    Dim idxSource As Index
    Dim idxTarget As Index
    Dim fldTarget As field
    Dim fldSource As field
    Dim j As Long
    
    Set tdfSource = db.TableDefs(source_table)
    Set tdfTarget = db.CreateTableDef(target_table)
                
    For j = 0 To tdfSource.Fields.count - 1
        Set fldSource = tdfSource.Fields(j)
        Set fldTarget = tdfTarget.CreateField(fldSource.Name)
        fldTarget.Attributes = fldSource.Attributes
        fldTarget.DefaultValue = fldSource.DefaultValue
        fldTarget.Name = fldSource.Name
        fldTarget.OrdinalPosition = fldSource.OrdinalPosition
        fldTarget.Required = fldSource.Required
        fldTarget.Size = fldSource.Size
        fldTarget.Type = fldSource.Type
        If fldSource.Type = dbText Or fldSource.Type = dbMemo Then
            fldTarget.AllowZeroLength = fldSource.AllowZeroLength
        End If
        tdfTarget.Fields.Append fldTarget
    Next j
    For j = 0 To tdfSource.Indexes.count - 1
        Set idxSource = tdfSource.Indexes(j)
        Set idxTarget = tdfTarget.CreateIndex(idxSource.Name)
        idxTarget.Clustered = idxSource.Clustered
        idxTarget.Fields = idxSource.Fields
        idxTarget.IgnoreNulls = idxSource.IgnoreNulls
        idxTarget.Primary = idxSource.Primary
        idxTarget.Required = idxSource.Required
        idxTarget.Unique = idxSource.Unique
        tdfTarget.Indexes.Append idxTarget
    Next j
                
    db.TableDefs.Append tdfTarget
                        
    Set tdfSource = Nothing
    Set tdfTarget = Nothing
    Set idxSource = Nothing
    Set idxTarget = Nothing
    Set fldSource = Nothing
    Set fldTarget = Nothing
End Sub

Public Function CloseDBase(db As Database) As Boolean
    On Error GoTo CloseErr
    db.Close
    CloseDBase = True
    GoTo Done
CloseErr:
    CloseDBase = False
Done:
End Function

Public Function CompactDB(db_name As String, Optional db As Database = Nothing, _
                Optional UseProgressBar As Boolean = False, _
                Optional ProgForm As Form, _
                Optional Icon As String, _
                Optional AbortButton As Boolean = False) As Database
    Dim TempDatabase As String
    Dim errloop As Error

    #If ProgressBar Then
        If UseProgressBar Then
            InitProgressBar ProgForm, "Compacting the Database . . .", 0, 100, _
                Icon, False, , AbortButton
        End If
    #End If
    
    If Not db Is Nothing Then
        If Not CloseDBase(db) Then
            MsgBox "CompactDB:  Couldn't close the database!", vbCritical
            Set CompactDB = Nothing
            GoTo Done
        End If
        Set db = Nothing
    End If
    
    TempDatabase = Left$(db_name, Len(db_name) - 4) & "1.mdb"
    
    On Error GoTo Err_Compact
    DBEngine.CompactDatabase db_name, TempDatabase
    On Error GoTo 0

    FileCopy TempDatabase, db_name
    Kill TempDatabase
    If db Is Nothing Then
        Set db = OpenDBase(db_name)
    End If
    
    Set CompactDB = db
    GoTo Done


Err_Compact:
    DisplayDBEngineErrors

Err_General:
    Set CompactDB = Nothing

Done:
    #If ProgressBar Then
        If UseProgressBar Then
            fProgForm.Hide
        End If
    #End If
End Function

Public Function RepairDB(db_name As String, Optional db As Database = Nothing, _
                Optional UseProgressBar As Boolean = False, _
                Optional ProgForm As Form, _
                Optional Icon As String = "", _
                Optional AbortButton As Boolean = False) As Database
    Dim errloop As Error

    #If ProgressBar Then
        If UseProgressBar Then
            InitProgressBar ProgForm, "Repairing the Database . . .", 0, 100, _
                Icon, False, , AbortButton
        End If
    #End If
    
    If Not db Is Nothing Then
        If Not CloseDBase(db) Then
            MsgBox "RepairDB:  Couldn't close the database!", vbCritical
            Set RepairDB = Nothing
            GoTo Done
        End If
        Set db = Nothing
    End If

    On Error GoTo Err_Repair
    DBEngine.RepairDatabase (db_name)
    On Error GoTo 0
    
    If db Is Nothing Then
        Set db = OpenDBase(db_name)
    End If
    
    Set RepairDB = db
    GoTo Done

Err_Repair:
    DisplayDBEngineErrors

Err_General:
    Set RepairDB = Nothing

Done:
    #If ProgressBar Then
        If UseProgressBar Then
            fProgForm.Hide
        End If
    #End If
End Function

' Reads a BLOB from a disk file and stores the contents in the ' specified table and field.
' PREREQUISITES: The specified table with the OLE object field to contain the binary data
' must be opened in Visual Basic code (Access Basic
' code in Microsoft Access 2.0 and earlier) and the correct record
' navigated to prior to calling the ReadBLOB() function. '
' ARGUMENTS:
' Source - The path and filename of the binary information
' to be read and stored.
' T - The table object to store the data in.
' Field - The OLE object field in table T to store the data in.
'
' RETURN:
' The number of bytes read from the Source file.
'**************************************************************
Public Function ReadBLOB(Source As String, t As Recordset, sField As String)
    Dim NumBlocks As Integer, _
        SourceFile As Integer, _
        i As Integer
    Dim FileLength As Long, _
        LeftOver As Long
    Dim FileData As String
    Dim RetVal As Variant

    On Error GoTo Err_ReadBLOB
            
    ' Open the source file.
    SourceFile = FreeFile
    Open Source For Binary Access Read As SourceFile
    ' Get the length of the file.
    FileLength = LOF(SourceFile)
    If FileLength = 0 Then
        ReadBLOB = 0
        GoTo Done
    End If
    ' Calculate the number of blocks to read and leftover bytes.
    NumBlocks = FileLength \ BLOCKSIZE
    LeftOver = FileLength Mod BLOCKSIZE
    ' SysCmd is used to manipulate status bar meter.

    RetVal = SysCmd(acSysCmdInitMeter, "Reading BLOB", FileLength \ 1000)

    ' Put first record in edit mode.
    t.MoveFirst
    t.Edit
    ' Read the leftover data, writing it to the table.
    FileData = String$(LeftOver, 32)
    Get SourceFile, , FileData
    t(sField).AppendChunk (FileData)
    RetVal = SysCmd(acSysCmdUpdateMeter, LeftOver / 1000)
    'Read the remaining blocks of data, writing them to the table.
    FileData = String$(BLOCKSIZE, 32)

    For i = 1 To NumBlocks
        Get SourceFile, , FileData
        t(sField).AppendChunk (FileData)
        RetVal = SysCmd(acSysCmdUpdateMeter, BLOCKSIZE * i / 1000)
    Next i
    ' Update the record and terminates function.
    t("data_value") = "<BINARY> " & DB_GetFileNameFromPath(Source)
    t.Update
    RetVal = SysCmd(acSysCmdRemoveMeter)
    ReadBLOB = FileLength
    GoTo Done
    
Err_ReadBLOB:
    ReadBLOB = -err

Done:
    DoEvents
    Close SourceFile
End Function

Public Function WriteBLOB(t As Recordset, sField As String, Destination As String)
    Dim NumBlocks As Integer
    Dim DestFile As Integer
    Dim i As Integer
    Dim FileLength As Long
    Dim LeftOver As Long
    Dim FileData As String
    Dim RetVal As Variant
    
    On Error GoTo Err_WriteBLOB ' Get the size of the field.
    
    
    DestFile = FreeFile
    
    ' Remove any existing destination file.
    Open Destination For Output As DestFile
    Close DestFile

    Open Destination For Binary As DestFile
    
    FileLength = t(sField).FieldSize()
    
    If FileLength = 0 Then
        WriteBLOB = 0
        GoTo Done
    End If
    

    NumBlocks = FileLength \ BLOCKSIZE
    LeftOver = FileLength Mod BLOCKSIZE

    
    FileData = t(sField).GetChunk(0, LeftOver)
    Put DestFile, , FileData
    
    For i = 1 To NumBlocks
        FileData = t(sField).GetChunk((i - 1) * BLOCKSIZE + LeftOver, BLOCKSIZE)
        Put DestFile, , FileData
    Next i

    WriteBLOB = FileLength
    GoTo Done
    
Err_WriteBLOB:
    WriteBLOB = -err
    
Done:
    DoEvents
    Close DestFile
End Function

Public Function CopyBLOB(srcRec As Recordset, srcField As String, dstRec As Recordset, _
        dstField As String)
    
    Dim NumBlocks As Integer, _
        DestFile As Integer, _
        i As Integer
    Dim FileLength As Long, _
        LeftOver As Long
    Dim FileData As String
    Dim RetVal As Variant
    
    On Error GoTo Err_copyBLOB ' Get the size of the field.
    
    FileLength = srcRec(srcField).FieldSize()
    
    If FileLength = 0 Then
        CopyBLOB = 0
        GoTo Done
    End If
    

    NumBlocks = FileLength \ BLOCKSIZE
    LeftOver = FileLength Mod BLOCKSIZE ' Remove any existing destination file.
    

    dstRec.MoveFirst
    dstRec.Edit

    FileData = String$(LeftOver, 32)
    FileData = srcRec(srcField).GetChunk(0, LeftOver)
    dstRec(dstField).AppendChunk (FileData)
    
    For i = 1 To NumBlocks
        FileData = srcRec(srcField).GetChunk((i - 1) * BLOCKSIZE + LeftOver, BLOCKSIZE)
        dstRec(dstField).AppendChunk (FileData)
    Next i
    dstRec.Update
    CopyBLOB = FileLength
    GoTo Done
    
Err_copyBLOB:
    CopyBLOB = -err

Done:
    DoEvents
End Function

Public Function ReadMemo(Source As String, t As Recordset, sField As String) As Integer
    
    Dim SourceFile As Integer
    Dim FileLength As Long
    Dim FileData As String

    SourceFile = FreeFile
    Open Source For Input As SourceFile

    FileLength = LOF(SourceFile)
    If FileLength = 0 Then
        ReadMemo = 0
        GoTo Done
    End If
    
    t.MoveFirst
    t.Edit
    Do Until EOF(SourceFile)
        Line Input #SourceFile, FileData
        t(sField).AppendChunk (FileData)
    Loop
    t.Update
    ReadMemo = 1

Done:
    DoEvents
    Close SourceFile
End Function


Public Function WriteMemo(t As Recordset, sField As String, Target As String) As Integer
    
    Dim TargetFile As Integer
    Dim FileLength As Long
    Dim FileData As String
    Dim NumBlocks As Long
    Dim LeftOver As Long
    Dim i As Long

    TargetFile = FreeFile
    Open Target For Output As TargetFile

    FileLength = t(sField).FieldSize()
    If FileLength = 0 Then
        WriteMemo = 0
        GoTo Done
    End If
    
    NumBlocks = FileLength \ BLOCKSIZE
    LeftOver = FileLength Mod BLOCKSIZE
    
    FileData = t(sField).GetChunk(0, LeftOver)
    Print #TargetFile, FileData
    
    For i = 1 To NumBlocks
        FileData = t(sField).GetChunk((i - 1) * BLOCKSIZE + LeftOver, BLOCKSIZE)
        Print #TargetFile, FileData
    Next i
    
    WriteMemo = 1
Done:
    Close TargetFile
    DoEvents
End Function

Public Function CopyMemo(srcRec As Recordset, srcField As String, dstRec As Recordset, _
        dstField As String, TempFile As String) As Integer
        
    Dim result As Integer
    Dim ValidFile As Boolean
    
    result = 1
    
    result = WriteMemo(srcRec, srcField, TempFile) <> 0
    
    On Error GoTo FileError
    If Dir(TempFile) = "" Then
        ValidFile = False
    Else
        ValidFile = True
    End If
    GoTo NoFileError
FileError:
    ValidFile = False
NoFileError:
    On Error GoTo 0
    
    If Not ValidFile Then
        MsgBox "The necessary temp file '" & TempFile _
            & "' was not created successfully." & Chr$(13) _
            & "The CopyMemo function did not complete successfully.", vbCritical
        result = 0
    End If
    
    If result <> 0 Then
        result = ReadMemo(TempFile, dstRec, dstField)
    End If
    DoEvents
    
    Kill TempFile
Done:
    CopyMemo = result
End Function

Public Sub DisplayDBEngineErrors()

Dim errloop As Error

    ' Notify user of any errors that result from
    ' executing the query.
    If DBEngine.Errors.count > 0 Then
        For Each errloop In DBEngine.Errors
            MsgBox "Error number: " & errloop.Number & vbCr & _
                errloop.description
        Next errloop
    End If

End Sub

Private Function DB_GetFileNameFromPath(filename As String) As String
    Dim i As Long
    Dim lastpos As Long
    
    If Left$(filename, 1) = "\" Then
        lastpos = 1
    Else
        lastpos = 0
    End If
    
    i = InStr(1, filename, "\", vbTextCompare)
    While i <> 0
        lastpos = i
        i = InStr(i + 1, filename, "\", vbTextCompare)
    Wend
    
    DB_GetFileNameFromPath = mID$(filename, lastpos + 1, Len(filename) - lastpos)

End Function

