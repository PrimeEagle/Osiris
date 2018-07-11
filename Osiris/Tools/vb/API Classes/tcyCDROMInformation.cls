VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCDROMInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Private Const DRIVE_CDROM = 5

'local variable(s) to hold property value(s)
Private mvarDrive As String 'local copy
Private mvarExists As Boolean 'local copy
Private mvarVolumeName As String 'local copy

Public Property Get VolumeName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VolumeName
    VolumeName = mvarVolumeName
End Property

Public Property Get Exists() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Exists
    FindCDROM
    Exists = mvarExists
End Property



Public Property Get Drive() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Drive
    Drive = mvarDrive
End Property



Private Sub FindCDROM()
  'get the available drives, determine their type,
  'and if CD, get the CD volume label
   Dim r As Long
   Dim DriveType As Long
   Dim allDrives As String
   Dim JustOneDrive As String
   Dim CDLabel As String
   Dim pos As Integer
   Dim CDfound As Boolean
   
  'pad the string with spaces
   allDrives$ = Space$(64)

  'call the API to get the string containing all drives
   r& = GetLogicalDriveStrings(Len(allDrives$), allDrives$)

  'trim off any trailing spaces.  AllDrives$
  'now contains all the drive letters.
   allDrives$ = Left$(allDrives$, r&)
   
  'begin a loop
   Do
    
     'first check that there is a chr$(0) in the string
      pos% = InStr(allDrives$, Chr$(0))
    
     'if there's one, then...
      If pos% Then
       
       'extract the drive up to the chr$(0)
        JustOneDrive$ = Left$(allDrives$, pos% - 1)
      
       'and remove that from the Alldrives string,
       'so it won't be checked again
        allDrives$ = Mid$(allDrives$, pos% + 1, Len(allDrives$))
      
       'with the one drive, call the API to
       'determine the drive type
        DriveType& = GetDriveType(JustOneDrive$)
        
       'check if it's what we want
        If DriveType& = DRIVE_CDROM Then
         
          'got it (or at least the first one,
          'anyway, if more than one), so set
          'the found flag... this part can be modified
          'to continue searching remaining drives for
          'those systems that might have more than
          'one CD installed.
           CDfound = True
           CDLabel$ = rgbGetVolumeLabel(JustOneDrive$)
          
          'we're done for now, so get out
           Exit Do
        
        End If
      End If
   
  Loop Until allDrives$ = "" Or DriveType& = DRIVE_CDROM

 'display the appropriate message
  mvarExists = CDfound
  If CDfound Then
        mvarDrive = UCase$(JustOneDrive$)
        mvarVolumeName = CDLabel$
  Else
        mvarDrive = ""
        mvarVolumname = ""
  End If
End Sub


Private Function rgbGetVolumeLabel(CDPath$) As String

 'create working variables
 'to keep it simple, use dummy variables for info
 'we're not interested in right now
  Dim r As Long
  Dim DrvVolumeName As String
  Dim pos As Integer

  Dim UnusedVal1 As Long
  Dim UnusedVal2 As Long
  Dim UnusedVal3 As Long
  Dim UnusedStr As String
  
  DrvVolumeName$ = Space$(14)
  UnusedStr$ = Space$(32)

 'do what it says
  r& = GetVolumeInformation(CDPath$, _
                            DrvVolumeName$, _
                            Len(DrvVolumeName$), _
                            UnusedVal1&, UnusedVal2&, _
                            UnusedVal3&, _
                            UnusedStr$, Len(UnusedStr$))

 'error check
  If r& = 0 Then Exit Function

 'the volume label
  pos% = InStr(DrvVolumeName$, Chr$(0))
  If pos% Then DrvVolumeName$ = Left$(DrvVolumeName$, pos% - 1)
  If Len(Trim$(DrvVolumeName$)) = 0 Then DrvVolumeName$ = "(no label)"
  
  rgbGetVolumeLabel = DrvVolumeName$

End Function

