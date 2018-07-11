VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDriveInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long

Private Const SEM_FAILCRITICALERRORS = &H1

'local variable(s) to hold property value(s)
Private lpRootPathName          As String
Private lpSectorsPerCluster     As Long
Private lpBytesPerSector        As Long
Private lpNumberOfFreeClusters  As Long
Private lpTotalNumberOfClusters As Long
Private mvarTotalBytes          As Currency 'local copy
Private mvarFreeBytes           As Currency 'local copy
Private mvarDrive               As String 'local copy
Private mvarUsedBytes           As Currency
Private mvarSerialID            As String 'local copy
Private mvarVolumeName          As String 'local copy
Private mvarExists              As Boolean 'local copy
Private mvarAvailableDrives     As String 'local copy

Public Property Get AvailableDrives() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AvailableDrives
    GetAvailableDrives
    AvailableDrives = mvarAvailableDrives
End Property

Public Property Get Exists() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Exists
    Exists = mvarExists
End Property

Public Property Let VolumeName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.VolumeName = 5
    mvarVolumeName = vData
End Property

Public Property Get VolumeName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.VolumeName
    VolumeName = mvarVolumeName
End Property

Public Property Get SerialID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SerialID
    SerialID = mvarSerialID
End Property

Public Property Let Drive(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Drive = 5
    Dim Result As Long
    
    If vData <> "" Then
        mvarExists = DriveExists(Left$(vData, 1))
        
        If mvarExists Then
            mvarDrive = vData
            
            If Right$(mvarDrive, 1) <> "\" Then
                 mvarDrive = mvarDrive & "\"
            End If
            
            Result = GetDiskFreeSpace(mvarDrive, lpSectorsPerCluster, lpBytesPerSector, lpNumberOfFreeClusters, lpTotalNumberOfClusters)
            
            mvarTotalBytes = CCur(lpSectorsPerCluster) * CCur(lpBytesPerSector) * CCur(lpTotalNumberOfClusters)
            mvarFreeBytes = CCur(lpSectorsPerCluster) * CCur(lpBytesPerSector) * CCur(lpNumberOfFreeClusters)
            mvarUsedBytes = CCur(lpSectorsPerCluster) * CCur(lpBytesPerSector) * (CCur(lpTotalNumberOfClusters) - CCur(lpNumberOfFreeClusters))
            
            rgbGetVolumeInformationRDI mvarDrive, mvarVolumeName, mvarSerialID
        Else
            'mvarFreeBytesAvailable = 0
            mvarTotalBytes = 0
            mvarFreeBytes = 0
            mvarUsedBytes = 0
            mvarSerialID = ""
            mvarVolumeName = ""
        End If
    End If
End Property


Public Property Get Drive() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Drive
    Drive = mvarDrive
End Property

Public Property Get UsedBytes() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalUsed
    UsedBytes = (mvarTotalBytes - mvarFreeBytes)
End Property

Public Property Get FreeBytes() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FreeBytes
    FreeBytes = mvarFreeBytes
End Property

Public Property Get TotalBytes() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TotalBytes
    TotalBytes = mvarTotalBytes
End Property

Private Sub rgbGetVolumeInformationRDI(PathName As String, DrvVolumeName As String, DrvSerialNo As String)
 'create working variables
 'to keep it simple, use dummy variables for info
 'we're not interested in right now
  Dim r             As Long
  Dim pos           As Integer
  Dim HiWord        As Long
  Dim HiHexStr      As String
  Dim LoWord        As Long
  Dim LoHexStr      As String
  Dim VolumeSN      As Long
  Dim MaxFNLen      As Long

  Dim UnusedStr     As String
  Dim UnusedVal1    As Long
  Dim UnusedVal2    As Long

 'pad the strings
  DrvVolumeName = Space$(14)
  UnusedStr = Space$(32)

 'do what it says
  r& = GetVolumeInformation(PathName, DrvVolumeName, Len(DrvVolumeName), VolumeSN&, UnusedVal1&, UnusedVal2&, UnusedStr, Len(UnusedStr))

 'error check
  If r& = 0 Then Exit Sub

 'determine the volume label
  pos% = InStr(DrvVolumeName, Chr$(0))
  If pos% Then DrvVolumeName = Left$(DrvVolumeName, pos% - 1)
  If Len(Trim$(DrvVolumeName)) = 0 Then DrvVolumeName = "(no label)"
  
 'determine the drive volume id
  HiWord& = GetHiWord(VolumeSN&) And &HFFFF&
  LoWord& = GetLoWord(VolumeSN&) And &HFFFF&
 
  HiHexStr = Format$(Hex(HiWord&), "0000")
  LoHexStr = Format$(Hex(LoWord&), "0000")
 
  DrvSerialNo = HiHexStr & "-" & LoHexStr
End Sub

Private Function GetHiWord(dw As Long) As Integer
    If dw& And &H80000000 Then
        GetHiWord% = (dw& \ 65535) - 1
    Else
        GetHiWord% = dw& \ 65535
    End If
End Function

Private Function GetLoWord(dw As Long) As Integer
    If dw& And &H8000& Then
        GetLoWord% = &H8000 Or (dw& And &H7FFF&)
    Else
        GetLoWord% = dw& And &HFFFF&
    End If
End Function

Private Function DriveExists(drvName As String) As Boolean
  'retrieve the available drives, and check each against the passed drive name
  'working variables
   Dim r As Long
   Dim allDrives As String
 
  'pad the string with spaces
   allDrives = Space$(64)
 
  'call the API to get the string containing all drives
   r& = GetLogicalDriveStrings(Len(allDrives), allDrives)
   
  'do a instr search, using text compare (not case sensitive)
   DriveExists = InStr(1, allDrives, drvName, 1)
End Function

Public Function DiskInDrive() As Boolean
    Dim ErrorMode As Long
    Dim F As Long
    
    F = FreeFile
    ErrorMode = SetErrorMode(SEM_FAILCRITICALERRORS)
    
    On Error Resume Next
    Open Left$(mvarDrive, 1) & ":\tya.tya" For Output As #F
    Close #F
    
    DiskInDrive = Not (Err.Number = 71)
    
    Kill Left$(mvarDrive, 1) & ":\tya.tya"
    
    On Error GoTo 0
    
    SetErrorMode ErrorMode
End Function

Private Sub GetAvailableDrives()
  'retrieve the available drives, and check each against the passed drive name
  'working variables
   Dim r As Long
   Dim allDrives As String
 
  'pad the string with spaces
   allDrives = Space$(64)
 
  'call the API to get the string containing all drives
   r& = GetLogicalDriveStrings(Len(allDrives), allDrives)
   
    mvarAvailableDrives = allDrives
End Sub

Private Sub Class_Initialize()
    GetAvailableDrives
End Sub
