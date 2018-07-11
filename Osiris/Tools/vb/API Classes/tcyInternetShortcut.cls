VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsInternetShortcut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarURL As String 'local copy
Private mvarFileName As String 'local copy

Public Function Create() As Boolean
    Dim F As Long
    
    On Error GoTo C
    
    Create = False
    If mvarFileName <> "" And mvarURL <> "" Then
        F = FreeFile
            
        If UCase$(Right$(mvarFileName, 4)) <> ".URL" Then
            mvarFileName = mvarFileName & ".URL"
        End If
        
        Open mvarFileName For Output As #F
            Print #F, "[InternetShortcut]"
            Print #F, "URL=" & mvarURL
        Close #F
        
        Create = True
    Else
        MsgBox "FileName and/or mvarURL property are required.", vbOKOnly + vbCritical, "Property Error"
    End If
    
    Exit Function
    
C:
    Err.Clear
End Function

Public Property Let FileName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileName = 5
    mvarFileName = vData
End Property


Public Property Get FileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileName
    FileName = mvarFileName
End Property



Public Property Let URL(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.URL = 5
    mvarURL = vData
End Property


Public Property Get URL() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.URL
    URL = mvarURL
End Property



