VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tcyNCDefaults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarProfile      As String 'local copy
Private mvarCurrentDate  As Date 'local copy
Private mvarCurrentMonth As Byte 'local copy
Private mvarFileName     As String 'local copy

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

Public Function Save() As Boolean
    Dim I As New tcyIniFile
    
    If mvarFileName = "" Then
        mvarFileName = App.Path & "\Default Provider.ncp"
    End If
    
    With I
        .FileName = mvarFileName
        .Section = "DefaultValues"
        
        If mvarProfile = "" Then mvarProfile = App.Path & "\Default Provider.ncp"
        .Parameter = "Profile"
        .Value = mvarProfile
        
        .Parameter = "CurrentDate"
        .Value = Format$(Date, "Short Date")
        
        .Parameter = "CurrentMonth"
        .Value = CStr(mvarCurrentMonth)
        
    End With
    
    Save = True
    
    Set I = Nothing
End Function

Public Function Load() As Boolean
    Dim I As New tcyIniFile
    
    If mvarFileName = "" Then
        mvarFileName = App.Path & "\NetClock.ini"
    End If
    
    With I
        .FileName = mvarFileName
        .Section = "DefaultValues"
        
        .Parameter = "Profile"
        mvarProfile = .Value
        
        .Default = Format$(Date, "Short Date")
        .Parameter = "CurrentDate"
        mvarCurrentDate = CDate(.Value)
        
        .Default = CStr(Month(Date))
        .Parameter = "CurrentMonth"
        mvarCurrentMonth = CByte(.Value)
        
    End With
    
    Load = True
    
    Set I = Nothing
End Function

Public Property Let CurrentMonth(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CurrentMonth = 5
    mvarCurrentMonth = vData
End Property


Public Property Get CurrentMonth() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CurrentMonth
    CurrentMonth = mvarCurrentMonth
End Property



Public Property Let CurrentDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CurrentDate = 5
    mvarCurrentDate = vData
End Property


Public Property Get CurrentDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CurrentDate
    CurrentDate = mvarCurrentDate
End Property



Public Property Let Profile(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Profile = 5
    mvarProfile = vData
End Property


Public Property Get Profile() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Profile
    Profile = mvarProfile
End Property



Private Sub Class_Initialize()
    mvarProfile = "Default Profile.ncp"
    mvarCurrentDate = Date
    mvarCurrentMonth = Month(Date)
    mvarFileName = App.Path & "\NetClock.ini"
End Sub

Private Sub Class_Terminate()
    Save
End Sub
