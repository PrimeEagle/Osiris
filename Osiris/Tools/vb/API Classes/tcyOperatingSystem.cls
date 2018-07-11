VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOperatingSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Declarations
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long

'Constants
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    szCSDVersion As String * 128
End Type

Enum EPlatform
    IsWin32 = 0
    isWin95 = 1
    isNT = 2
End Enum

'Variables
Dim myVer As OSVERSIONINFO

'Default Property Values:
Const m_def_Version = ""
Const m_def_Platform = ""
Const m_def_OEM = ""
Const m_def_Major = 0
Const m_def_Minor = 0
Const m_def_Build = 0
Const m_def_CSDVersion = ""
Const m_def_Organization = ""
Const m_def_Owner = ""

Dim m_Version As String
Dim m_Platform As String
Dim m_OEM As String
Dim m_Major As Long
Dim m_Minor As Long
Dim m_Build As Long
Dim m_CSDVersion As String * 128
Dim m_Organization As String
Dim m_Owner As String

Private Sub Class_Initialize()
    Dim Result As Long
    Dim clsRegistry As New clsRegistry
    
    Result = GetVersionEx(myVer)
    
    myVer.dwOSVersionInfoSize = 148
    
    Result = GetVersionEx(myVer)
    
    With myVer
        Select Case .dwPlatformID
            Case Is = IsWin32
                m_Platform = "Win 32"
                
            Case Is = isWin95
                m_Platform = "Windows 95"
            
            Case Is = isNT
                m_Platform = "Windows NT"
                
            Case Else
                m_Platform = "<Unknown>"
        End Select
            
        m_Version = CStr(.dwMajorVersion) & "." & (.dwMinorVersion) & " Build " & CStr((.dwBuildNumber And &HFFFF&)) & CStr(.szCSDVersion)
        m_Version = Left$(m_Version, InStr(m_Version, Chr$(0)) - 1)
        m_Major = .dwMajorVersion
        m_Minor = .dwMinorVersion
        m_Build = (.dwBuildNumber And &HFFFF&)
        m_CSDVersion = .szCSDVersion
    End With
    
    With clsRegistry
        .Root = LocalMachine
        .Path = "Software\Microsoft\Windows\Currentversion"
        .Key = "ProductID"
        
        m_OEM = .Value
        
        .Key = "RegisteredOrganization"
        m_Organization = .Value
        
        .Key = "RegisteredOwner"
        m_Owner = .Value
    End With
    
    Set clsRegistry = Nothing
End Sub

Public Property Get Version() As String
    Version = m_Version
End Property

Public Property Get Platform() As String
    Platform = m_Platform
End Property

Public Property Get OEM() As String
    OEM = m_OEM
End Property

Public Property Get Major() As Long
    Major = m_Major
End Property

Public Property Get Minor() As Long
    Minor = m_Minor
End Property

Public Property Get Build() As Long
    Build = m_Build
End Property

Public Property Get CSDVersion() As String
    CSDVersion = m_CSDVersion
End Property

Public Property Get Organization() As String
    Organization = m_Organization
End Property

Public Property Get Owner() As String
    Owner = m_Owner
End Property

