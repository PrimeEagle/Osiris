VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long

Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long
Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long

Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

'Default Property Values:
Const m_def_Defaults = 0        'None
Const m_def_Root = &H80000002   'HKEY_LOCAL_MACHINE
Const m_def_Path = ""
Const m_def_Key = ""
Const m_def_Value = ""

'Registry security values
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20

Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Const KEY_EXECUTE = KEY_READ
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

'Property Variables:
Dim m_Defaults As Variant
Dim m_Root     As ERoot
Dim m_Path     As String
Dim m_Key      As String
Dim m_Value    As Variant

'Enumerated property types
Enum ERoot
    Classes = &H80000000
    CurrentUser = &H80000001
    LocalMachine = &H80000002
    Users = &H80000003
    PerformanceData = &H80000004
End Enum

Enum EDefaults
    None = 0
    UserName = 1
    ComputerName = 2
    AppData = 3
    Cache = 4
    Cookies = 5
    Desktop = 6
    Favorites = 7
    Fonts = 8
    History = 9
    NetHood = 10
    Personal = 11
    Programs = 12
    Recent = 13
    SendTo = 14
    StartMenu = 15
    StartUp = 16
    Templates = 17
End Enum

Enum EType
    asNone = 0                          ' No value type
    asString = 1                        ' Unicode nul terminated string
    asStringNull = 2                    ' Unicode nul terminated string
    asBinary = 3                        ' Free form binary
    asDWord = 4                         ' 32-bit number
    asSmallDWord = 4                    ' 32-bit number
    asLargeDWord = 5                    ' 32-bit number
    asLink = 6                          ' Symbolic Link (unicode)
    asMultipleNull = 7                  ' Multiple Unicode strings
    asResourceList = 8                  ' Resource list in the resource map
    asResourceDescriptor = 9            ' Resource list in the hardware description
    asResourceRequirementsList = 10
    asBoolean = 91
End Enum

Public Property Get Root() As ERoot
    Root = m_Root
End Property

Public Property Let Root(ByVal New_Root As ERoot)
    m_Root = New_Root
    
    Update_Value
End Property

Public Property Get Path() As String
    Path = m_Path
End Property

Public Property Let Path(ByVal New_Path As String)
    m_Path = Remove_BackSlash(New_Path)
    
    Update_Value
End Property

Public Property Get Key() As String
    Key = m_Key
End Property

Public Property Let Key(ByVal New_Key As String)
    m_Key = New_Key
    
    Update_Value
End Property

Public Property Get Value() As Variant
    Value = Update_Value
End Property

Public Property Let Value(ByVal New_Value As Variant)
    Dim lRetVal As Long      'result of the SetValueEx function
    Dim hKey    As Long      'handle of open key
    Dim hTemp   As Long

    lRetVal = RegOpenKeyEx(m_Root, m_Path, 0, KEY_ALL_ACCESS, hKey)
        
    If hKey = 0 Then
          lRetVal = RegCreateKeyEx(m_Root, m_Path, 0&, vbNullString, KEY_ALL_ACCESS, 0&, &H3F, hKey, lRetVal)
    End If
    
    lRetVal = Value_Set(hKey, m_Key, New_Value)
    
    If (lRetVal = 0) Then
        m_Value = New_Value
        'PropertyChanged "Value"
    End If
    
    RegCloseKey (hKey)
End Property

'Initialize Properties for User Control
Private Sub Class_Initialize()
    m_Root = m_def_Root
    m_Path = m_def_Path
    m_Key = m_def_Key
    m_Value = m_def_Value
    m_Defaults = m_def_Defaults
End Sub

Private Function Value_Get(ByVal lhKey As Long, vValue As Variant) As Long
    Dim cch    As Long
    Dim lrc    As Long
    Dim lType  As Long
    Dim lValue As Long
    Dim sValue As String
    Dim Result As Variant

    On Error GoTo QueryValueExError
    Result = Empty

    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhKey, m_Key, 0&, lType, 0&, cch)
    
    If lrc <> 0 Then Error 5

    Select Case lType
        Case asString, asStringNull
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, m_Key, 0&, lType, sValue, cch)
            
            If (lrc = 0) Then
                Result = Left$(sValue, cch - 1)
            End If
        
        Case asDWord
            lrc = RegQueryValueExLong(lhKey, m_Key, 0&, lType, lValue, cch)

            If (lrc = 0) Then
                Result = lValue
            End If
    End Select

QueryValueExExit:

    If IsEmpty(Result) Then
        Value_Get = -1
        vValue = Empty
    Else
        vValue = Result
        Value_Get = 0
    End If
    
    Exit Function

QueryValueExError:

    Resume QueryValueExExit
End Function

Private Function Value_Set(ByVal hKey As Long, ByVal sValueName As String, ByVal vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String
    Dim lType  As Long
    Dim Result As Long
    
    lType = Value_Type(vValue)
    Result = Empty
    
    Select Case lType
        Case Is = asString, asStringNull
            sValue = vValue & Chr$(0)
            Result = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        
        Case asDWord
            lValue = vValue
            Result = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)

    End Select
    
    Value_Set = Result
End Function

Public Property Get Defaults() As EDefaults
    Defaults = m_Defaults
End Property

Public Property Let Defaults(ByVal New_Defaults As EDefaults)
    m_Defaults = New_Defaults
    
    Defaults_AssignToProperties
End Property

Private Sub Defaults_AssignToProperties()
    Dim BasePath As String
    
    BasePath = "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"

    Select Case m_Defaults
        Case Is = UserName
            m_Root = LocalMachine
            m_Path = "System\CurrentControlSet\Control"
            m_Key = "Current User"
            
        Case Is = ComputerName
            m_Root = LocalMachine
            m_Path = "System\CurrentControlSet\Control\ComputerName\ComputerName"
            m_Key = "ComputerName"
        
        Case Is = AppData
            m_Root = CurrentUser
            m_Path = BasePath
            m_Key = "AppData"
            
        Case Is = Cache
            m_Root = CurrentUser
            m_Path = BasePath
            m_Key = "Cache"
        
        Case Is = Cookies
            m_Root = CurrentUser
            m_Path = BasePath
            m_Key = "Cookies"
        
        Case Is = Desktop
            m_Root = CurrentUser
            m_Path = BasePath
            m_Key = "Desktop"
        
        Case Is = Favorites
            m_Root = CurrentUser
            m_Path = BasePath
            m_Key = "Favorites"
        
        Case Is = Fonts
            m_Root = CurrentUser
            m_Path = BasePath
            m_Key = "Fonts"
        
        Case Is = History
            m_Root = CurrentUser
            m_Path = BasePath
            m_Key = "History"
        
        Case Is = NetHood
            m_Root = CurrentUser
            m_Path = BasePath
            m_Key = "NetHood"
        
        Case Is = Personal
            m_Root = CurrentUser
            m_Path = BasePath
            m_Key = "Personal"
        
        Case Is = Programs
            m_Root = CurrentUser
            m_Path = BasePath
            m_Key = "Programs"
        
        Case Is = Recent
            m_Root = CurrentUser
            m_Path = BasePath
            m_Key = "Recent"
        
        Case Is = SendTo
            m_Root = CurrentUser
            m_Path = BasePath
            m_Key = "SendTo"
        
        Case Is = StartMenu
            m_Root = CurrentUser
            m_Path = BasePath
            m_Key = "Start Menu"
        
        Case Is = StartUp
            m_Root = CurrentUser
            m_Path = BasePath
            m_Key = "StartUp"
        
        Case Is = Templates
            m_Root = CurrentUser
            m_Path = BasePath
            m_Key = "Templates"
    End Select
End Sub

Private Function Value_Type(ByVal vValue As Variant) As Long
    Dim Result As Long
    
    Result = Empty
    
    Select Case VarType(vValue)
        Case Is = vbString
            Result = asString
            
        Case Is = vbLong
            Result = asDWord
            
        Case Is = vbBoolean
            Result = asBoolean
    End Select
    
    Value_Type = Result
End Function

Private Function Remove_BackSlash(ByVal NewPath As String) As String
    Dim Result As String
    
    Result = NewPath
    
    If Left$(Result, 1) = "\" Then
        Result = Mid$(Result, 2, Len(Result))
    End If
    
    If Right$(Result, 1) = "\" Then
        Result = Left$(Result, Len(Result) - 1)
    End If
    
    Remove_BackSlash = Result
End Function

Private Function Update_Value() As Variant
    Dim lRetVal As Long          'result of the API functions
    Dim hKey    As Long          'handle of opened key
    Dim Result  As Variant       'setting of queried value

    lRetVal = RegOpenKeyEx(CLng(m_Root), m_Path, 0, KEY_ALL_ACCESS, hKey)
    
    lRetVal = Value_Get(hKey, Result)
    
    If (lRetVal = 0) Then
        Update_Value = Result
    End If
    
    RegCloseKey (hKey)
End Function

Public Function DeletePath() As Boolean
    DeletePath = CBool(RegDeleteKey(CLng(m_Root), m_Path))
End Function

