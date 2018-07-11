VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tcyTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarExists    As Boolean  'local copy
Private mvarTableName As String   'local copy
Private mvarDatabase  As Database 'local copy
Private colFldNames   As New Collection
Private colFldTypes   As New Collection

Public Property Let Database(ByVal vData As Database)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Database = Form1
    Set mvarDatabase = vData
End Property


Public Property Get Database() As Database
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Database
    Set Database = mvarDatabase
End Property

Public Property Let TableName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TableName = 5
    mvarTableName = vData
End Property


Public Property Get TableName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TableName
    TableName = mvarTableName
End Property



Public Property Let Exists(ByVal vData As Boolean)
    mvarExists = vData
End Property


Public Property Get Exists() As Boolean
    Dim Tbl As TableDef
    
    mvarExists = False
    
    If Not (mvarDatabase Is Nothing) Then
        For Each Tbl In mvarDatabase.TableDefs
            If UCase$(Tbl.Name) = UCase$(Trim$(mvarTableName)) Then
                mvarExists = True
            End If
        Next
    End If
    
    Set Tbl = Nothing
    
    Exists = mvarExists
End Property

Public Function Create() As Boolean
    
End Function

Public Sub AddField(FldName As String, FldType As Long)
    colFldNames.Add FldName, FldName
    colFldTypes.Add FldType, FldName
End Sub

Public Function RemoveField(FldName As String) As Boolean
    On Error Resume Next
    colFldNames.Remove FldName
    colFldTypes.Remove FldName
    
    RemoveField = (Err.Number = 0)
End Function

Public Sub ClearFields()
    Set colFldNames = Nothing
    Set colFldTypes = Nothing
End Sub
