VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormatDrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function SHFormatDrive Lib "shell32" (ByVal hwndOwner As Long, ByVal iDrive As Long, ByVal iCapacity As Long, ByVal iFormatType As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
                                     
Private Type OSVERSIONINFO
      dwOSVersionInfoSize As Long
      dwMajorVersion As Long
      dwMinorVersion As Long
      dwBuildNumber As Long
      dwPlatformId As Long
      szCSDVersion As String * 128      ' Maintenance string for PSS usage
End Type
   
Enum TPlatform
    vbWin32s = 0
    vbWin95 = 1
    vbWinNT = 2
End Enum

Enum TCapacity
    vbDefault = 0
    vb360k = 3
    vb720k = 5
End Enum

Enum TAction
    vb95Quick = 0
    vbNTQuick = 1
    vb95Full = 1
    vbNTFull = 0
    vbSystemOnly = 2
End Enum
 
      
'Enum TDriveTypes
'    vbRemovable = 2
'    vbFixed = 3
'    vbRemote = 4
'    vbCDROM = 5
'    vbRAM = 6
'End Enum

'local variable(s) to hold property value(s)
Private mvarAction  As TAction 'local copy
Private mvarCapacity As TCapacity 'local copy
Private mvarDrive   As String 'local copy

Public Property Let Drive(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Drive = 5
    mvarDrive = vData
End Property


Public Property Get Drive() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Drive
    Drive = mvarDrive
End Property



Public Property Let Capacity(ByVal vData As TCapacity)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Capaciy = 5
    mvarCapacity = vData
End Property


Public Property Get Capacity() As TCapacity
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Capaciy
    Capaciy = mvarCapacity
End Property



Public Property Let Action(ByVal vData As TAction)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Action = 5
    mvarAction = vData
End Property


Public Property Get Action() As TAction
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Action
    Action = mvarAction
End Property


Private Function IsWinNT() As Boolean
   Dim osvi As OSVERSIONINFO
   
   osvi.dwOSVersionInfoSize = Len(osvi)
   
   GetVersionEx osvi
   
   IsWinNT = (osvi.dwPlatformId = vbWinNT)
End Function

Public Function Format() As Boolean
    Dim Result As Long
    
    Result = -1
    If IsWinNT() Then
        If mvarAction = vb95Full Or mvarAction = vb95Quick Then
            MsgBox "Property Action is set to Win95 settings rather than NT.", vbOKOnly + vbCritical, "Property Error"
            
            Exit Function
        End If
    End If
    
    If DriveExists() Then
        Result = SHFormatDrive(GetDesktopWindow(), CDriveToInt(), CLng(mvarCapacity), CLng(mvarAction))
    
        Format = (Result = 0)
    End If
End Function

Private Sub Class_Initialize()
    mvarCapacity = vbDefault
    
    If IsWinNT() Then
        mvarAction = vbNTFull
    Else
        mvarAction = vb95Full
    End If
End Sub

Private Function DriveExists() As Boolean
    Dim X As Variant
    Dim LD As New clsLogicalDrives
    Dim Drives As Collection
    
    DriveExists = False
    LD.GetLogicalDrives
    Set Drives = LD.Drives
    For Each X In Drives
        If UCase$(X) = mvarDrive Then
            DriveExists = True
            Exit For
        End If
    Next
    
    Set X = Nothing
    Set LD = Nothing
End Function

Private Function CDriveToInt() As Integer
    Dim X       As Variant
    Dim LD      As New clsLogicalDrives
    Dim Drives  As Collection
    Dim Counter As Integer
    Dim Found    As Boolean
    
    LD.GetLogicalDrives
    Set Drives = LD.Drives
    Counter = 0
    Found = False
    
    For Each X In Drives
        If UCase$(X) = mvarDrive Then
            Found = True
            Exit For
        End If
        Counter = Counter + 1
    Next
    
    If Found Then
        CDriveToInt = Counter
    Else
        CDriveToInt = -1
    End If
    
    Set X = Nothing
    Set LD = Nothing
End Function
