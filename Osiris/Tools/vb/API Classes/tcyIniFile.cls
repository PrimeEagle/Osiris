VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tcyIniFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'local variable(s) to hold property value(s)
Private mvarFileName  As String 'local copy
Private mvarParameter As String 'local copy
Private mvarSection   As String 'local copy
Private mvarValue     As String 'local copy
Private mvarDefault   As String 'local copy

Public Property Let Default(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Default = 5
    mvarDefault = vData
End Property

Public Property Get Default() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Default
    Default = mvarDefault
End Property

Public Function DeleteSection() As Boolean

End Function

Public Property Let Value(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Value = 5
    mvarValue = vData
    
    If Not SetValue() Then
        mvarValue = ""
    End If
End Property


Public Property Get Value() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Value
    GetValue
    
    Value = mvarValue
End Property

Public Property Let Section(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Section = 5
    mvarSection = vData

    GetValue
End Property

Public Property Get Section() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Section
    Section = mvarSection
End Property

Public Property Let Parameter(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Parameter = 5
    mvarParameter = vData
    
    GetValue
End Property

Public Property Get Parameter() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Parameter
    Parameter = mvarParameter
End Property

Public Property Let FileName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileName = 5
    mvarFileName = vData
    
    GetValue
End Property

Public Property Get FileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileName
    FileName = mvarFileName
End Property

Private Sub GetValue()
    Dim Bytes  As Integer
    Dim Temp   As String * 255
    Dim Result As String
    
    Result = mvarDefault
    Bytes = 0
    
    Bytes = GetPrivateProfileString(mvarSection, mvarParameter, mvarDefault, Temp, 255, mvarFileName)
                
    If (Bytes <> 0) Then
        Result = Left$(Temp, Bytes)
    End If
    
    mvarValue = Result
End Sub

Private Function SetValue() As Boolean
    Dim Result As Long
    
    Result = WritePrivateProfileString(mvarSection, mvarParameter, mvarValue, mvarFileName)
    
    SetValue = (Result <> 0)
End Function



Private Sub Class_Initialize()
    mvarDefault = ""
    mvarFileName = ""
    mvarSection = ""
    mvarParameter = ""
    mvarValue = ""
End Sub
