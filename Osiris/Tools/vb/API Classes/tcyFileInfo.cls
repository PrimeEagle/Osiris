VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIMEREC, lpSystemTime As SYSTEMTIMEREC) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIMEREC, lpLastAccessTime As FILETIMEREC, lpLastWriteTime As FILETIMEREC) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCTREC, ByVal wStyle As Long) As Long
Private Declare Function hread Lib "kernel32" Alias "_hread" (ByVal hFile As Long, lpBuffer As Any, ByVal lBytes As Long) As Long
Private Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, ByVal lpvSource As Any, ByVal cbcopy As Long)
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
   
Private Const VS_FFI_SIGNATURE = &HFEEF04BD
Private Const VS_FFI_STRUCVERSION = &H10000
Private Const VS_FFI_FILEFLAGSMASK = &H3F&
Private Const OF_READ = &H0
Private Const OF_SHARE_DENY_NONE = &H40
Private Const OFS_MAXPATHNAME = 128
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260

Enum EBuildType
    VS_FF_DEBUG = &H1
    VS_FF_PRERELEASE = &H2
    VS_FF_PATCHED = &H4
    VS_FF_PRIVATEBUILD = &H8
    VS_FF_INFOINFERRED = &H10
    VS_FF_SPECIALBUILD = &H20
End Enum

Enum EOperatingSystem
    VOS_UNKNOWN = &H0
    VOS_DOS = &H10000
    VOS_OS216 = &H20000
    VOS_OS232 = &H30000
    VOS_NT = &H40000
    VOS__BASE = &H0
    VOS__WINDOWS16 = &H1
    VOS__PM16 = &H2
    VOS__PM32 = &H3
    VOS__WINDOWS32 = &H4
    VOS_DOS_WINDOWS16 = &H10001
    VOS_DOS_WINDOWS32 = &H10004
    VOS_OS216_PM16 = &H20002
    VOS_OS232_PM32 = &H30003
    VOS_NT_WINDOWS32 = &H40004
End Enum

Enum ESubType
    VFT2_UNKNOWN = &H0
    VFT2_DRV_PRINTER = &H1
    VFT2_DRV_KEYBOARD = &H2
    VFT2_DRV_LANGUAGE = &H3
    VFT2_DRV_DISPLAY = &H4
    VFT2_DRV_MOUSE = &H5
    VFT2_DRV_NETWORK = &H6
    VFT2_DRV_SYSTEM = &H7
    VFT2_DRV_INSTALLABLE = &H8
    VFT2_DRV_SOUND = &H9
    VFT2_DRV_COMM = &HA
End Enum

Enum TFileType
    VFT_UNKNOWN = &H0
    VFT_APP = &H1
    VFT_DLL = &H2
    VFT_DRV = &H3
    VFT_FONT = &H4
    VFT_VXD = &H5
    VFT_STATIC_LIB = &H7
End Enum

Enum TFontType
    VFT_FONT_RASTER = &H1&
    VFT_FONT_VECTOR = &H2&
    VFT_FONT_TRUETYPE = &H3&
End Enum

Private Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
   dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
   dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
   dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
   dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
   dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
   dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
   dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
   dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
   dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
   dwFileFlagsMask As Long        '  = &h3F for version "0.42"
   dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
   dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
   dwFileType As Long             '  e.g. VFT_DRIVER
   dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
   dwFileDateMS As Long           '  e.g. 0
   dwFileDateLS As Long           '  e.g. 0
End Type

Private Type FILETIME
    dwLowDateTime  As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type


Private Type OFSTRUCTREC
     cBytes As Byte
     fFixedDisk As Byte
     nErrCode As Integer
     Reserved1 As Integer
     Reserved2 As Integer
     szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Type FILETIMEREC
     dwLowDateTime As Long
     dwHighDateTime As Long
End Type

Private Type SYSTEMTIMEREC
     wYear As Integer
     wMonth As Integer
     wDayOfWeek As Integer
     wDay As Integer
     wHour As Integer
     wMinute As Integer
     wSecond As Integer
     wMilliseconds As Integer
End Type

'local variable(s) to hold property value(s)
Private mvarFileName As String 'local copy
Private mvarBuildType As String 'local copy
Private mvarLastAccessed As Variant 'local copy
Private mvarCreated As Variant 'local copy
Private mvarFileType As String 'local copy
Private mvarFileSubType As String 'local copy
Private mvarFileVersion As String 'local copy
Private mvarComments As String 'local copy
Private mvarCompanyName As String 'local copy
Private mvarInternalName As String 'local copy
Private mvarLegalCopyright As String 'local copy
Private mvarFileDescription As String 'local copy
Private mvarLegalTrademark As String 'local copy
Private mvarOLESelfRegister As Boolean 'local copy
Private mvarOriginalFileName As String 'local copy
Private mvarProductName As String 'local copy
Private mvarFileSize As Currency 'local copy
Private mvarLastModified As Variant 'local copy
Private mvarStructureVersion As String 'local copy
Private mvarProductVersion As String 'local copy
Private mvarOperatingSystem As String 'local copy
Private mvarExists As Boolean 'local copy

Public MajorVersion As Integer
Public MinorVersion As Integer
Public RevisionVersion As Integer
Public BuildVersion As Integer

Public Property Get Exists() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Exists
    
    Exists = mvarExists
End Property

Public Property Get OperatingSystem() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OperatingSystem
    OperatingSystem = mvarOperatingSystem
End Property

Public Property Get ProductVersion() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProductVersion
    ProductVersion = mvarProductVersion
End Property

Public Property Get StructureVersion() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StructureVersion
    StructureVersion = mvarStructureVersion
End Property

Public Property Get LastModified() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastModified
    LastModified = mvarLastModified
End Property

Public Property Get FileSize() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileSize
    FileSize = mvarFileSize
End Property

Public Property Get ProductName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProductName
    ProductName = mvarProductName
End Property


Public Property Get OriginalFileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OriginalFileName
    OriginalFileName = mvarOriginalFileName
End Property

Public Property Get OLESelfRegister() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OLESelfRegister
    OLESelfRegister = mvarOLESelfRegister
End Property

Public Property Get LegalTrademark() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LegalTrademark
    LegalTrademark = mvarLegalTrademark
End Property

Public Property Get FileDescription() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileDescription
    FileDescription = mvarFileDescription
End Property

Public Property Get LegalCopyright() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LegalCopyright
    LegalCopyright = mvarLegalCopyright
End Property

Public Property Get InternalName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.InternalName
    InternalName = mvarInternalName
End Property

Public Property Get CompanyName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CompanyName
    CompanyName = mvarCompanyName
End Property

Public Property Get Comments() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Comments
    Comments = mvarComments
End Property

Public Property Get FileVersion() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileVersion
    FileVersion = mvarFileVersion
End Property

Public Property Get FileSubType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileSubType
    FileSubType = mvarFileSubType
End Property

Public Property Get FileType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileType
    FileType = mvarFileType
End Property

Public Property Get Created() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Created
    Created = mvarCreated
End Property

Public Property Get LastAccessed() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LastAccessed
    LastAccessed = mvarLastAccessed
End Property

Public Property Get BuildType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BuildType
    BuildType = mvarBuildType
End Property

Public Property Let FileName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileName = 5
    mvarFileName = vData
    
    'Is file property set?
    If (mvarFileName <> "") Then
        If FileExists(mvarFileName) Then
            AssignValues
        End If
    Else
        Class_Initialize
    End If
End Property

Public Property Get FileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileName
    FileName = mvarFileName
End Property

Private Sub AssignValues()
    Dim rc                As Long
    Dim lDummy            As Long
    Dim sBuffer()         As Byte
    Dim lBufferLen        As Long
    Dim lVerPointer       As Long
    Dim udtVerBuffer      As VS_FIXEDFILEINFO
    Dim lVerbufferLen     As Long
    
    '*** Get size ****
    lBufferLen = GetFileVersionInfoSize(mvarFileName, lDummy)
    If (lBufferLen < 1) Then
        Class_Initialize
        Exit Sub
    End If
    
    '**** Store info to udtVerBuffer struct ****
    ReDim sBuffer(lBufferLen)
    rc = GetFileVersionInfo(mvarFileName, 0&, lBufferLen, sBuffer(0))
    rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
    MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)
    
    '**** Determine Structure Version number - NOT USED ****
    mvarStructureVersion = Format$(udtVerBuffer.dwStrucVersionh) & "." & Format$(udtVerBuffer.dwStrucVersionl)
    
    '**** Determine File Version number ****
    mvarFileVersion = Format$(udtVerBuffer.dwFileVersionMSh) & "." & Format$(udtVerBuffer.dwFileVersionMSl) & "." & Format$(udtVerBuffer.dwFileVersionLSh) & "." & Format$(udtVerBuffer.dwFileVersionLSl)
    MajorVersion = udtVerBuffer.dwFileVersionMSh
    MinorVersion = udtVerBuffer.dwFileVersionMSl
    RevisionVersion = udtVerBuffer.dwFileVersionLSh
    BuildVersion = udtVerBuffer.dwFileVersionLSl
    
    '**** Determine Product Version number ****
    mvarProductVersion = Format$(udtVerBuffer.dwProductVersionMSh) & "." & Format$(udtVerBuffer.dwProductVersionMSl) & "." & Format$(udtVerBuffer.dwProductVersionLSh) & "." & Format$(udtVerBuffer.dwProductVersionLSl)
   
    '**** Determine Boolean attributes of File ****
    mvarBuildType = ""
    If udtVerBuffer.dwFileFlags And VS_FF_DEBUG Then mvarBuildType = "Debug "
    If udtVerBuffer.dwFileFlags And VS_FF_PRERELEASE Then mvarBuildType = mvarBuildType & "PreRelease "
    If udtVerBuffer.dwFileFlags And VS_FF_PATCHED Then mvarBuildType = mvarBuildType & "Patched "
    If udtVerBuffer.dwFileFlags And VS_FF_PRIVATEBUILD Then mvarBuildType = mvarBuildType & "Private "
    If udtVerBuffer.dwFileFlags And VS_FF_INFOINFERRED Then mvarBuildType = mvarBuildType & "Information "
    If udtVerBuffer.dwFileFlags And VS_FF_SPECIALBUILD Then mvarBuildType = mvarBuildType & "Special "
       
    If mvarBuildType = "" Then mvarBuildType = "Unknown"
    
    '**** Determine OS for which file was designed ****
    mvarOperatingSystem = ""
    Select Case udtVerBuffer.dwFileOS
        Case VOS_DOS_WINDOWS16
            mvarOperatingSystem = "DOS-Win16"
        
        Case VOS_DOS_WINDOWS32
            mvarOperatingSystem = "DOS-Win32"
        
        Case VOS_OS216_PM16
            mvarOperatingSystem = "OS/2-16 PM-16"
        
        Case VOS_OS232_PM32
            mvarOperatingSystem = "OS/2-16 PM-32"
        
        Case VOS_NT_WINDOWS32
            mvarOperatingSystem = "NT-Win32"
        
        Case 4 'Windows 95 (original)
            mvarOperatingSystem = "Win 95"
            
        Case Else
            mvarOperatingSystem = "Unknown"
    End Select
    
    mvarFileSubType = "N/A"
    Select Case udtVerBuffer.dwFileType
        Case VFT_APP
            mvarFileType = "App"
            
        Case VFT_DLL
            mvarFileType = "DLL"
            
        Case VFT_DRV
            mvarFileType = "Driver"
            
            Select Case udtVerBuffer.dwFileSubtype
                Case VFT2_DRV_PRINTER
                    mvarFileSubType = "Printer drv"
                
                Case VFT2_DRV_KEYBOARD
                    mvarFileSubType = "Keyboard drv"
                
                Case VFT2_DRV_LANGUAGE
                    mvarFileSubType = "Language drv"
                
                Case VFT2_DRV_DISPLAY
                    mvarFileSubType = "Display drv"
                
                Case VFT2_DRV_MOUSE
                    mvarFileSubType = "Mouse drv"
                    
                Case VFT2_DRV_NETWORK
                    mvarFileSubType = "Network drv"
                    
                Case VFT2_DRV_SYSTEM
                    mvarFileSubType = "System drv"
                
                Case VFT2_DRV_INSTALLABLE
                    mvarFileSubType = "Installable"
                    
                Case VFT2_DRV_SOUND
                    mvarFileSubType = "Sound drv"
                    
                Case VFT2_DRV_COMM
                    mvarFileSubType = "Comm drv"
                    
                Case VFT2_UNKNOWN
                    mvarFileSubType = "Unknown driver"
            End Select
        
        Case VFT_FONT
            mvarFileType = "Font"
            
            Select Case udtVerBuffer.dwFileSubtype
                Case VFT_FONT_RASTER
                    mvarFileSubType = "Raster Font"
            
                Case VFT_FONT_VECTOR
                    mvarFileSubType = "Vector Font"
            
                Case VFT_FONT_TRUETYPE
                    mvarFileSubType = "TrueType Font"
                    
                Case Else
                    mvarFileSubType = "Unknown font"
            End Select
        
        Case VFT_VXD
            mvarFileType = "VxD"
        
        Case VFT_STATIC_LIB
            mvarFileType = "Lib"
        
        Case Else
            mvarFileType = "Unknown type"
    End Select
    
    'Filesize
    mvarFileSize = FileLen(mvarFileName)
    
    'Other product Information
    mvarComments = Version_GetString("Comments")
    mvarCompanyName = Version_GetString("CompanyName")
    mvarFileDescription = Version_GetString("FileDescription")
    mvarInternalName = Version_GetString("InternalName")
    mvarLegalCopyright = Version_GetString("LegalCopyright")
    mvarLegalTrademark = Version_GetString("LegalTrademark")
    
    If (Version_GetString("OLESelfRegister") <> "Error") Then
        mvarOLESelfRegister = True
    Else
        mvarOLESelfRegister = False
    End If
    
    mvarOriginalFileName = Version_GetString("OriginalFileName")
    mvarProductName = Version_GetString("ProductName")
    
    Get_FileDateTimeInfo
End Sub

Private Function Version_GetString(ByVal SeekValue As String) As String
    Dim tBuf        As String
    Dim NullPos     As Integer
    Dim xLateLang   As Integer
    Dim xLateCode   As Integer
    Dim NumEntries  As Integer
    Dim fiiAddr     As Long
    Dim fiiAddr2    As Long
    Dim xLateString As String
    Dim xLateval    As Long
    Dim fiiLen      As Long
    Dim di          As Long
    Dim Lp          As Integer
    Dim Result      As String
    
    Dim VerBuf()    As Byte
    Dim fResSize    As Long
    Dim fResHnd     As Long
    
    'Default values
    Result = "Error"
    
    fResSize = GetFileVersionInfoSize(mvarFileName, fResHnd)
    
    If (fResSize <> 0) Then
        ReDim VerBuf(fResSize + 1)
        ReDim InfoBuf(fResSize + 1)
        
        di = GetFileVersionInfo(mvarFileName, fResHnd, fResSize, VerBuf(0))
        
        If (di <> 0) Then
            di = VerQueryValue(VerBuf(0), "\VarFileInfo\Translation", fiiAddr, fiiLen)
            
            If (di <> 0) Then
                NumEntries = fiiLen / 4
                xLateval = 0
                
                For Lp = 1 To NumEntries
                    MoveMemory xLateLang, ByVal fiiAddr, 2
                    MoveMemory xLateCode, ByVal (fiiAddr + 2), 2
                    
                    If xLateLang = &H409 Then Exit For
                    
                    fiiAddr = fiiAddr + 4
                Next Lp
            Else
                xLateLang = &H409
                xLateCode = 0
            End If
            
            xLateString = Hex(xLateCode)
            
            While Len(xLateString) < 4
                xLateString = "0" & xLateString
            Wend
            
            xLateString = Hex(xLateLang) & xLateString
            
            While Len(xLateString) < 8
                xLateString = "0" & xLateString
            Wend
            
            di = VerQueryValue(VerBuf(0), "\StringFileInfo\" & xLateString & "\" & SeekValue, fiiAddr2, fiiLen)
            
            If (di <> 0) Then
                tBuf = String(fiiLen + 1, Chr(0))
                
                MoveMemory ByVal tBuf, ByVal fiiAddr2, fiiLen
                
                NullPos = InStr(tBuf, Chr(0))
                
                If (NullPos > 1) Then
                    Result = Left(tBuf, NullPos - 1)
                Else
                    Result = "None"
                End If
            Else
                Result = "Unavailable"
            End If
        End If
    End If
    
    Version_GetString = Result
End Function

Private Sub Get_FileDateTimeInfo()
    Dim sInpFile       As String
    Dim hFile          As Integer
    Dim FileStruct     As OFSTRUCTREC
    Dim iRC            As Integer
    Dim CreationTime   As FILETIMEREC
    Dim LastAccessTime As FILETIMEREC
    Dim LastWriteTime  As FILETIMEREC
    Dim SystemTime     As SYSTEMTIMEREC
    
    hFile = OpenFile(mvarFileName, FileStruct, OF_READ Or OF_SHARE_DENY_NONE)
    If hFile = 0 Then
        MsgBox "Can't open the file", vbExclamation
        Exit Sub
    End If
    
    If GetFileTime(hFile, CreationTime, LastAccessTime, LastWriteTime) Then
        With SystemTime
            FileTimeToSystemTime LastAccessTime, SystemTime
            mvarLastAccessed = CDate(Format(CStr(.wMonth) & "/" & CStr(.wDay) & "/" & CStr(.wYear) & " " & CStr(.wHour) & ":" & CStr(.wMinute) & ":" & CStr(.wSecond), "General Date"))
            
            FileTimeToSystemTime CreationTime, SystemTime
            mvarCreated = CDate(Format(CStr(.wMonth) & "/" & CStr(.wDay) & "/" & CStr(.wYear) & " " & CStr(.wHour) & ":" & CStr(.wMinute) & ":" & CStr(.wSecond), "General Date"))
            
            FileTimeToSystemTime LastWriteTime, SystemTime
            mvarLastModified = CDate(Format(CStr(.wMonth) & "/" & CStr(.wDay) & "/" & CStr(.wYear) & " " & CStr(.wHour) & ":" & CStr(.wMinute) & ":" & CStr(.wSecond), "General Date"))
        End With
    Else
        Class_Initialize
    End If
    
    iRC = lclose(hFile)
End Sub

Private Sub Class_Initialize()
    mvarFileName = ""
    mvarBuildType = ""
    mvarFileType = ""
    mvarFileSubType = ""
    mvarFileVersion = ""
    mvarComments = ""
    mvarCompanyName = ""
    mvarInternalName = ""
    mvarLegalCopyright = ""
    mvarFileDescription = ""
    mvarLegalTrademark = ""
    mvarOLESelfRegister = False
    mvarOriginalFileName = ""
    mvarProductName = ""
    mvarFileSize = 0#
    mvarStructureVersion = ""
    mvarProductVersion = ""
    mvarOperatingSystem = ""
End Sub

Private Function FileExists(sSource As String) As Boolean
    Dim WFD As WIN32_FIND_DATA
    Dim hFile As Long
    
    hFile = FindFirstFile(sSource, WFD)
    
    FileExists = (hFile <> INVALID_HANDLE_VALUE)
    
    Call FindClose(hFile)
    
    mvarExists = FileExists
End Function

Public Function Unqiue(Optional Path As String, Optional Prefix As String) As String
    Dim r As Long     'get the user's windows\temp folder  'pad the passed string
    Dim sWinTmpDir As String
    Dim sTmpFile As String
    
    If IsMissing(Path) Then
        sWinTmpDir = Space$(MAX_PATH)
        
        'get the folder
        r = GetTempPath(MAX_PATH, sWinTmpDir)
        
        'r contains the number of chrs up to the
        'terminating null, so a simple left$ can
        'be used. Its also conveniently terminated
        'with a slash.
        sWinTmpDir = Left$(sWinTmpDir, r)
    Else
        sWinTmpDir = Path
    End If
    
    If IsMissing(Prefix) Then
        Prefix = "VBN"
    End If
    
    'pad a working string
    sTmpFile = Space$(MAX_PATH)
    
    r = GetTempFileName(sWinTmpDir, Prefix, 0, sTmpFile)
    If r <> 0 Then
        'strip the trailing null
        sTmpFile = Left$(sTmpFile, InStr(sTmpFile, Chr$(0)) - 1)
    End If
    
    Unqiue = sTmpFile
End Function
