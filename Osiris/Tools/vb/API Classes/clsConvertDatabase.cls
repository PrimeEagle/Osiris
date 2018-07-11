VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConvertDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Enum TConvertTo
    ctVersion10 = 1
    ctVersion11 = 8
    ctVersion20 = 16
    ctVersion30 = 32
End Enum

'local variable(s) to hold property value(s)
Private mvarConvertTo                     As TConvertTo 'local copy
Private mvarSourceDatabaseName            As String 'local copy
Private mvarDestinationDatabaseName       As String 'local copy
Private mvarDeleteOriginalAfterConversion As Boolean 'local copy
Private mvarEnableMessage                 As Boolean 'local copy

Public Property Let DeleteOriginalAfterConversion(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DeleteOriginalAfterConversion = 5
    mvarDeleteOriginalAfterConversion = vData
End Property

Public Property Get DeleteOriginalAfterConversion() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DeleteOriginalAfterConversion
    DeleteOriginalAfterConversion = mvarDeleteOriginalAfterConversion
End Property

Public Property Let EnableMessage(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DeleteOriginalAfterConversion = 5
    mvarEnableMessage = vData
End Property


Public Property Get EnableMessage() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DeleteOriginalAfterConversion
    EnableMessage = mvarEnableMessage
End Property


Public Function Convert() As Boolean
    Dim WS                          As Workspace
    Dim szDBVersion                 As String
    Dim szOldDBName                 As String
    Dim ConvertedDB, DB             As Database
    Dim DriveUsed                   As New clsDriveInformation
    Dim ConvertTo                   As Variant
    Dim Messageform                 As Form
    
    On Error GoTo ErrHandler
    Convert = False
    
    If Trim$(mvarSourceDatabaseName) <> "" Then
        If Trim$(mvarDestinationDatabaseName) <> "" Then
            Set WS = DBEngine.Workspaces(0)
            Set DB = WS.OpenDatabase(mvarSourceDatabaseName, True)
            
            'If current DB version < ConvertTo, then it should be converted
            If Not NeedsConverted(DB, ConvertTo) Then
                Convert = True
            Else
                DB.Close
                
                If MsgBox("Database " & mvarSourceDatabaseName & " should be converted to Access 3.5. Convert database?", vbYesNo, "Convert Database?") = vbYes Then
                    'Get the drive name and check to see if it exists.
                    If InStr(mvarSourceDatabaseName, ":") <> 0 Then
                        DriveUsed.Drive = Left$(mvarSourceDatabaseName, 2)
                    Else
                        DriveUsed.Drive = "C:"
                    End If
                    
                    If (DriveUsed.Exists) Then
                        'Check for enough space to convert the database
                        'DBFile.FileName = mvarSourceDatabaseName
                        If (FileLen(mvarSourceDatabaseName) * 2.5) < DriveUsed.FreeBytes Then
                        
                            ' If the Destination Database already exists, kill it
                            If Dir$(mvarDestinationDatabaseName, vbNormal) <> "" Then
                                Kill mvarDestinationDatabaseName
                            End If
                            
'                            'If the coder wants a message while converting
'                            If mvarEnableMessage Then
'                                Messageform =
'                            End If
                            
                            'Compact and compress the database
                            DBEngine.CompactDatabase mvarSourceDatabaseName, mvarDestinationDatabaseName, dbLangGeneral, ConvertTo
                            
                            'Close the databases, rename/kill the old to a backup name and the new
                            'database to the applications main DB name.
                            If mvarDeleteOriginalAfterConversion Then
                                Kill mvarSourceDatabaseName
                            Else
                                If Dir$(mvarSourceDatabaseName & ".OLD", vbNormal) <> "" Then
                                    Kill mvarSourceDatabaseName & ".OLD"
                                End If
                                Name mvarSourceDatabaseName As mvarSourceDatabaseName & ".OLD"
                            End If
                            
                            'Rename the new database to the name of the old database
                            Name mvarDestinationDatabaseName As mvarSourceDatabaseName
                            
                            Convert = True
                            
                            MsgBox "Database was successfully converted", vbInformation
                        Else
                            MsgBox "Drive could not be found.", vbCritical, "Error!"
                        End If
                    Else
                        MsgBox "Not enough drive space to convert file", vbCritical, "Error!"
                    End If
                End If
            End If
        Else
            MsgBox "DestinationDatabaseName property is required", vbOKOnly + vbCritical, "Required Property"
        End If
    Else
        MsgBox "SourceDatabaseName property is required", vbOKOnly + vbCritical, "Required Property"
    End If
    
    'Free the memory
    Set WS = Nothing
    Set DriveUsed = Nothing

    Exit Function
    
'Inform the user of an error and close everything.
ErrHandler:
    MsgBox "An error occurred while converting the database"
    
    Set WS = Nothing
    Set DriveUsed = Nothing

    Err.Clear
End Function

Public Property Let DestinationDatabaseName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DestinationDatabaseName = 5
    mvarDestinationDatabaseName = vData
End Property


Public Property Get DestinationDatabaseName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DestinationDatabaseName
    DestinationDatabaseName = mvarDestinationDatabaseName
End Property



Public Property Let SourceDatabaseName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SourceDatabaseName = 5
    mvarSourceDatabaseName = vData
End Property


Public Property Get SourceDatabaseName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SourceDatabaseName
    SourceDatabaseName = mvarSourceDatabaseName
End Property



Public Property Let ConvertTo(ByVal vData As TConvertTo)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ConvertTo = 5
    mvarConvertTo = vData
End Property


Public Property Get ConvertTo() As TConvertTo
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ConvertTo
    ConvertTo = mvarConvertTo
End Property



Private Sub Class_Initialize()
    mvarConvertTo = ctVersion30
    mvarSourceDatabaseName = ""
    mvarDestinationDatabaseName = ""
    mvarDeleteOriginalAfterConversion = False
    mvarEnableMessage = True
End Sub

Private Function NeedsConverted(DB As Database, ByRef ConvertTo As Variant) As Boolean
    NeedsConverted = False
    Select Case Left$(DB.Version, 3)
        Case Is = "2.0"
            NeedsConverted = (mvarConvertTo > ctVersion20)
            
        Case Is = "3.0"
            NeedsConverted = (mvarConvertTo > ctVersion30)
    End Select
    
    Select Case mvarConvertTo
        Case Is = ctVersion20
            ConvertTo = dbVersion20
            
        Case Is = ctVersion30
            ConvertTo = dbVersion30
    End Select
End Function
