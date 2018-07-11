VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tcyTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Property Variables:
Dim mvarEnabled     As Boolean
Dim mvarToolTipText As String
Dim mvarIcon        As Object
Dim mvarForm        As Object 'local copy

'Event Declarations:
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event LeftMouseDown()
Event LeftMouseUp()
Event LeftMouseDoubleClick()
Event RightMouseDown()
Event RightMouseUp()
Event RightMouseDoubleClick()

'Declare a user-defined variable to pass to the Shell_NotifyIcon
'function.
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'The following constants are the messages sent to the
'Shell_NotifyIcon function to add, modify, or delete an icon from the
'taskbar status area.
Private Const NImvarADD = &H0
Private Const NImvarMODIFY = &H1
Private Const NImvarDELETE = &H2

'The following constants are the flags that indicate the valid
'members of the NOTIFYICONDATA data type.
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

'The following constant is the message sent when a mouse event occurs
'within the rectangular boundaries of the icon in the taskbar status
'area.
Private Const WmvarMOUSEMOVE = &H200

'The following constants are used to determine the mouse input on the
'the icon in the taskbar status area.
'Left-click constants.
Private Const WmvarLBUTTONDBLCLK = &H203   'Double-click
Private Const WmvarLBUTTONDOWN = &H201     'Button down
Private Const WmvarLBUTTONUP = &H202       'Button up

'Right-click constants.
Private Const WmvarRBUTTONDBLCLK = &H206   'Double-click
Private Const WmvarRBUTTONDOWN = &H204     'Button down
Private Const WmvarRBUTTONUP = &H205       'Button up

'Declare the API function call.
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'Dimension a variable as the user-defined data type.
Dim nid As NOTIFYICONDATA

Public Property Let Form(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Form = Form1
    Set mvarForm = vData
End Property


Public Property Get ToolTipText() As String
    ToolTipText = mvarToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    mvarToolTipText = New_ToolTipText
    PropertyChanged "ToolTipText"
    
    nid.szTip = mvarToolTipText & vbNullChar
    Shell_NotifyIcon NImvarMODIFY, nid
End Property

Public Property Get Icon() As Object
    Set Icon = mvarIcon
End Property

Public Property Let Icon(ByVal New_Icon As Object)
    Set mvarIcon = New_Icon
End Property

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
    'Event occurs when the mouse pointer is within the rectangular
    'boundaries of the icon in the taskbar status area.
    Select Case (X / Screen.TwipsPerPixelX)
        Case WmvarLBUTTONDOWN
            RaiseEvent LeftMouseDown
            
        Case WmvarLBUTTONUP
            RaiseEvent LeftMouseUp
            
        Case WmvarLBUTTONDBLCLK
            RaiseEvent LeftMouseDoubleClick
            
        Case WmvarRBUTTONDOWN
            RaiseEvent RightMouseDown
            
        Case WmvarRBUTTONUP
            RaiseEvent RightMouseUp
        Case WmvarRBUTTONDBLCLK
            RaiseEvent RightMouseDoubleClick
    End Select
End Sub

Public Property Get Enabled() As Boolean

    Enabled = mvarEnabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    mvarEnabled = New_Enabled

    If mvarEnabled Then
        AddToTray
    Else
        DeleteFromTray
    End If
End Property

Private Function AddToTray() As Boolean
    'Click this button to add an icon to the taskbar status area.
    'Set the individual values of the NOTIFYICONDATA data type.
    With nid
        .cbSize = Len(nid)
        .hWnd = mvarForm.hWnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WmvarMOUSEMOVE
        
        If Not (mvarIcon Is Nothing) Then
            .hIcon = mvarForm.Icon
        Else
            .hIcon = 0
        End If
        
        .szTip = mvarToolTipText & vbNullChar
    End With
    
    'Call the Shell_NotifyIcon function to add the icon to the taskbar
    'status area.
    AddToTray = Shell_NotifyIcon(NImvarADD, nid)
End Function

Private Function DeleteFromTray() As Boolean
'Click this button to delete the added icon from the taskbar
    'status area by calling the Shell_NotifyIcon function.
    DeleteFromTray = Shell_NotifyIcon(NImvarDELETE, nid)
End Function


Private Sub Class_Initialize()
    mvarToolTipText = ""
    
    Set mvarIcon = LoadPicture("")
    mvarEnabled = False
End Sub

Private Sub Class_Terminate()
    Shell_NotifyIcon NImvarDELETE, nid
End Sub
