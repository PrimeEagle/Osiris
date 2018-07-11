VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogicalDrives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
      
Enum TDriveTypes
    vbRemovable = 2
    vbFixed = 3
    vbRemote = 4
    vbCDROM = 5
    vbRAM = 6
End Enum

'local variable(s) to hold property value(s)
Private mvarDrives As New Collection 'local copy

Public Function DriveType(ByVal Drive As String) As Long
    DriveType = GetDriveType(Drive)
End Function

Public Property Get Drives() As Collection
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Drives
    Set Drives = mvarDrives
End Property


Public Sub GetLogicalDrives()
  'get the available drives, determine their type,
  'and if CD, get the CD volume label
   Dim R As Long
   Dim DriveType As Long
   Dim allDrives As String
   Dim JustOneDrive As String
   Dim CDLabel As String
   Dim pos As Integer
   Dim CDfound As Boolean
   
  Set mvarDrives = Nothing
  
  'pad the string with spaces
   allDrives$ = Space$(64)

  'call the API to get the string containing all drives
   R& = GetLogicalDriveStrings(Len(allDrives$), allDrives$)

  'trim off any trailing spaces.  AllDrives$
  'now contains all the drive letters.
   allDrives$ = Left$(allDrives$, R&)
   
  'begin a loop
   Do
    
     'first check that there is a chr$(0) in the string
      pos% = InStr(allDrives$, Chr$(0))
    
     'if there's one, then...
      If pos% Then
       
       'extract the drive up to the chr$(0)
       JustOneDrive$ = Left$(allDrives$, pos% - 1)
      
       mvarDrives.Add JustOneDrive$, JustOneDrive$
       
       'and remove that from the Alldrives string,
       'so it won't be checked again
       allDrives$ = Mid$(allDrives$, pos% + 1, Len(allDrives$))
      End If
   
  Loop Until allDrives$ = ""
End Sub

