VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRegisterFiletype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long

Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003

Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_INVALID_PARAMETER = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259

Private Const KEY_ALL_ACCESS = &H3F

Private Const REG_OPTION_NON_VOLATILE = 0

'local variable(s) to hold property value(s)
Private mvarDescription As String 'local copy
Private mvarExtension As String 'local copy
Private mvarCommandLine As String 'local copy

Public Function Register() As Boolean
    Dim sShellPath As String
    
    Register = False
    If mvarCommandLine <> "" Then
        If mvarExtension <> "" Then
            If mvarDescription <> "" Then

                CreateNewKey mvarExtension, HKEY_CLASSES_ROOT
                SetKeyValue mvarExtension, "", mvarDescription & ".Document", REG_SZ
                CreateNewKey mvarDescription & ".Document\shell\open\command", HKEY_CLASSES_ROOT
                SetKeyValue mvarDescription & ".Document", "", mvarDescription, REG_SZ
                
                ' Double quotes around the path are required if long filenames are used
                sShellPath = Chr$(34) & mvarCommandLine & Chr$(34) & " " & Chr$(34) & "%1" & Chr$(34)
                
                SetKeyValue mvarDescription & ".Document\shell\open\command", "", sShellPath, REG_SZ
                
                Register = True
            Else
                MsgBox "Description property is required", vbOKOnly + vbInformation, "Invalid property"
            End If
        Else
            MsgBox "Extension property is required", vbOKOnly + vbInformation, "Invalid property"
        End If
    Else
        MsgBox "CommandLine property is required", vbOKOnly + vbInformation, "Invalid property"
    End If
End Function

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



Public Property Let Extension(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Extension = 5
    mvarExtension = vData
End Property


Public Property Get Extension() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Extension
    Extension = mvarExtension
End Property



Public Property Let Description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
End Property


Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property



Private Sub Class_Initialize()
    mvarDescription = ""
    mvarCommandLine = ""
    mvarExtension = ""
    
End Sub


Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long

    Dim lValue As Long
    Dim sValue As String
    Select Case lType
        Case REG_SZ
            sValue = vValue & Chr$(0)
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)

    End Select

End Function


Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)

     Dim hNewKey As Long         'handle to the new key
     Dim lRetVal As Long         'result of the RegCreateKeyEx function

     lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hNewKey, lRetVal)
     RegCloseKey (hNewKey)

End Sub

Sub SetKeyValue(sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)

     Dim lRetVal As Long         'result of the SetValueEx function
     Dim hKey As Long         'handle of open key

     'open the specified key
     lRetVal = RegOpenKeyEx(HKEY_CLASSES_ROOT, sKeyName, 0, KEY_ALL_ACCESS, hKey)
     lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
     RegCloseKey (hKey)

End Sub



