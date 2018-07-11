VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWaitShellExecute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


'local variable(s) to hold property value(s)
Private mvarCommandLine As String 'local copy
Private mvarClass       As String 'local copy
Private mvarCaption     As String 'local copy

Public Function Execute() As Boolean
    Const HINSTANCE_ERROR% = 0
    
    Dim hInstChild As Double
    Dim shname     As String

    If mvarCommandLine <> "" Then
        hInstChild = Shell(mvarCommandLine, vbNormalFocus)
        
        If mvarClass = "" Then
            hInstChild = FindWindow(vbNullString, mvarCaption)
        Else
            hInstChild = FindWindow(mvarClass, vbNullString)
        End If
        
        If hInstChild >= HINSTANCE_ERROR Then
            If mvarClass = "" Then
                While FindWindow(vbNullString, mvarCaption) <> 0
                    DoEvents
                Wend
            Else
                While FindWindow(mvarClass, vbNullString) <> 0
                    DoEvents
                Wend
            End If
        End If
    End If
    
    Execute = (hInstChild > HINSTANCE_ERROR)
End Function

Public Property Let Caption(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Caption = 5
    mvarCaption = vData
End Property


Public Property Get Caption() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Caption
    Caption = mvarCaption
End Property



Public Property Let Class(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Class = 5
    mvarClass = vData
End Property


Public Property Get Class() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Class
    Class = mvarClass
End Property



Public Property Let CommandLine(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CommandLine = 5
    mvarCommandLine = vData
End Property


Public Property Get CommandLine() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CommandLine
    CommandLine = mvarCommandLine
End Property



Private Sub Class_Initialize()
    mvarClass = ""
    mvarCaption = ""
    mvarCommandLine = ""
End Sub
