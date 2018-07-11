VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsControlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Enum TCPL
    vbControlpanel = -1
    vbAccessabilityGeneral = 0
    vbAccessabilityKeyboard = 1
    vbAccessabilitySound = 2
    vbAccessabilityDisplay = 3
    vbAccessabilityMouse = 4
    vbAddNewHardware = 5
    vbaddnewprinter = 6
    vbInstallUnInstall = 7
    vbWindowsSetup = 8
    vbStartupDisk = 9
    vbDialUpNetworking = 10
    vbBackground = 11
    vbScreenSaver = 12
    vbAppearance = 13
    vbSettings = 14
    vbAudio = 15
    vbVideo = 16
    vbMIDI = 17
    vbCDMusic = 18
    VBMMAdvanced = 19
    vbSound = 20
    vbSystemGeneral = 21
    vbDeviceManager = 22
    vbHardwareProfiles = 23
    vbPerformance = 24
    vbNetworkConfig = 25
    vbOpenWith = 26
    vbSetDateTime = 27
    vbTimeZone = 28
    vbFindfast = 29
    vbInternet = 30
    vbRegionalSettings = 31
    vbJoystick = 32
    vbModem = 33
    vbPassword = 34
    vbTelephony = 35
    vbMouse = 36
End Enum

Public Sub Show(ByVal CPL As TCPL, Optional OpenFile As Variant)
    Dim Buf    As String * 255
    Dim SysDir As String
    Dim Bytes  As Long
    
    Bytes = GetSystemDirectory(Buf, 255)
    SysDir = Left$(Buf, Bytes)
    
    Select Case CPL
        Case Is = vbControlpanel
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL", "", vbNormalFocus
        Case Is = vbAccessabilityGeneral
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL access.cpl,,5", "", vbNormalFocus
        Case Is = vbAccessabilityKeyboard
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL access.cpl,,1", "", vbNormalFocus
        Case Is = vbAccessabilitySound
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL access.cpl,,2", "", vbNormalFocus
        Case Is = vbAccessabilityDisplay
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL access.cpl,,3", "", vbNormalFocus
        Case Is = vbAccessabilityMouse
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL access.cpl,,4", "", vbNormalFocus
        Case Is = vbAddNewHardware
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL sysdm.cpl @1", "", vbNormalFocus
        Case Is = vbaddnewprinter
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,SHHelpShortcuts_RunDLL AddPrinter", "", vbNormalFocus
        Case Is = vbInstallUnInstall
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL appwiz.cpl,,1", "", vbNormalFocus
        Case Is = vbWindowsSetup
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL appwiz.cpl,,2", "", vbNormalFocus
        Case Is = vbStartupDisk
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL appwiz.cpl,,3", "", vbNormalFocus
        'Case Is = vbDialUpNetworking
        '    ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "rnaui.dll,RnaDial internet", "", vbNormalFocus
        Case Is = vbBackground
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL desk.cpl,,0", "", vbNormalFocus
        Case Is = vbScreenSaver
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL desk.cpl,,1", "", vbNormalFocus
        Case Is = vbAppearance
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL desk.cpl,,2", "", vbNormalFocus
        Case Is = vbSettings
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL desk.cpl,,3", "", vbNormalFocus
        Case Is = vbAudio
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL mmsys.cpl,,0", "", vbNormalFocus
        Case Is = vbVideo
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL mmsys.cpl,,1", "", vbNormalFocus
        Case Is = vbMIDI
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL mmsys.cpl,,2", "", vbNormalFocus
        Case Is = vbCDMusic
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL mmsys.cpl,,3", "", vbNormalFocus
        Case Is = VBMMAdvanced
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL mmsys.cpl,,4", "", vbNormalFocus
        Case Is = vbSound
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL mmsys.cpl,,5", "", vbNormalFocus
        Case Is = vbSystemGeneral
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL sysdm.cpl,,0", "", vbNormalFocus
        Case Is = vbDeviceManager
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL sysdm.cpl,,1", "", vbNormalFocus
        Case Is = vbHardwareProfiles
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL sysdm.cpl,,2", "", vbNormalFocus
        Case Is = vbPerformance
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL sysdm.cpl,,3", "", vbNormalFocus
        Case Is = vbNetworkConfig
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL netcpl.cpl", "", vbNormalFocus
        Case Is = vbOpenWith
            If Not IsMissing(OpenFile) Then
                ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,OpenAs_RunDLL " & CStr(OpenFile), "", vbNormalFocus
            Else
                MsgBox "When using the clsControlPanel: vbOpenWith, you must also specify a valid path:\file", vbOKOnly + vbInformation, "Parameters not valid"
            End If
        Case Is = vbSetDateTime
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL timedate.cpl,,0", "", vbNormalFocus
        Case Is = vbTimeZone
            ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL timedate.cpl,,1", "", vbNormalFocus
        Case Is = vbFindfast
            If Dir$(SysDir & "\findfast.cpl", vbNormal) <> "" Then
                ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL findfast.cpl", "", vbNormalFocus
            End If
        Case Is = vbInternet
            If Dir$(SysDir & "\inetcpl.cpl", vbNormal) <> "" Then
                ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL inetcpl.cpl", "", vbNormalFocus
            End If
        Case Is = vbRegionalSettings
            If Dir$(SysDir & "\intl.cpl", vbNormal) <> "" Then
                ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL intl.cpl", "", vbNormalFocus
            End If
        Case Is = vbJoystick
            If Dir$(SysDir & "\joy.cpl", vbNormal) <> "" Then
                ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL joy.cpl", "", vbNormalFocus
            End If
        Case Is = vbModem
            If Dir$(SysDir & "\modem.cpl", vbNormal) <> "" Then
                ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL modem.cpl", "", vbNormalFocus
            End If
        Case Is = vbPassword
            If Dir$(SysDir & "\password.cpl", vbNormal) <> "" Then
                ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL password.cpl", "", vbNormalFocus
            End If
        Case Is = vbTelephony
            If Dir$(SysDir & "\Telephon.cpl", vbNormal) <> "" Then
                ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL telephon.cpl", "", vbNormalFocus
            End If
        Case Is = vbMouse
            If Dir$(SysDir & "\Mouse.cpl", vbNormal) <> "" Then
                ShellExecute GetDesktopWindow(), "open", "rundll32.exe", "shell32.dll,Control_RunDLL Mouse.cpl", "", vbNormalFocus
            End If
    End Select
End Sub

