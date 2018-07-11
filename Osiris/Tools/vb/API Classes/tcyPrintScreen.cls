VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPrintScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Enum TSource
    vbScreen = 0
    vbActiveForm = 1
End Enum

'local variable(s) to hold property value(s)
Private mvarSource As Integer 'local copy

Public Sub SendToPrinter()
    Printer.PaintPicture Clipboard.GetData(2), 0, 0
End Sub

Public Sub Capture()
    keybd_event vbKeySnapshot, mvarSource, 0&, 0&
    
    'Needed to wait for system to complete copy
    DoEvents
End Sub

Public Property Let Source(ByVal vData As TSource)
    mvarSource = vData
End Property

Public Property Get Source() As TSource
    Source = mvarSource
End Property

Private Sub Class_Initialize()
    mvarSource = vbScreen
End Sub
