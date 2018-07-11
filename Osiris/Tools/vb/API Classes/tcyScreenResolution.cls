VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsScreenResolution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000
Private Const CCFORMNAME = 32
Private Const CCDEVICENAME = 32
Private Const SM_CXSCREEN = 0       'Screen width
Private Const SM_CYSCREEN = 1       'Screen height

Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer

    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer

    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

'local variable(s) to hold property value(s)
Private mvarHorizontal As Integer 'local copy
Private mvarVertical   As Integer 'local copy
Private mvarCurrentHorizontal As Integer 'local copy
Private mvarCurrentVertical As Integer 'local copy
Private mvarCurrentNumberOfColors As Long 'local copy

Public Property Get CurrentNumberOfColors() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CurrentNumberOfColors
    CurrentNumberOfColors = mvarCurrentNumberOfColors
End Property


Public Property Get CurrentVertical() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CurrentVertical
    CurrentVertical = mvarCurrentVertical
End Property


Public Property Get CurrentHorizontal() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CurrentHorizontal
    CurrentHorizontal = mvarCurrentHorizontal
End Property




Public Sub Change()
    Dim DevM As DEVMODE
    Dim a    As Boolean
    Dim i    As Long
    Dim b    As Long

    If mvarHorizontal <> 0 And mvarVertical <> 0 Then
        i = 0
    
        'Enumerate settings
        Do
            a = EnumDisplaySettings(0&, i&, DevM)
            i = i + 1
        Loop Until (a = False)
    
        'Change settings
        DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
    
        DevM.dmPelsWidth = mvarHorizontal
        DevM.dmPelsHeight = mvarVertical
    
        b = ChangeDisplaySettings(DevM, 0)
    End If
    
    GetCurrenctVideoInformation
End Sub

Public Property Let Vertical(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Vertical = 5
    mvarVertical = vData
End Property


Public Property Get Vertical() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Vertical
    Vertical = mvarVertical
End Property



Public Property Let Horizontal(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Horizontal = 5
    mvarHorizontal = vData
End Property


Public Property Get Horizontal() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Horizontal
    Horizontal = mvarHorizontal
End Property



Private Sub GetCurrenctVideoInformation()
    Const BITSPIXEL = 12
    Const PLANES = 14
    
    Dim R      As Long
    Dim Result As Long
    Dim hSrcDC As Long

    hSrcDC = GetDC(GetDesktopWindow())
    mvarCurrentNumberOfColors = GetDeviceCaps(hSrcDC, PLANES) * 2 ^ GetDeviceCaps(hSrcDC, BITSPIXEL)
    R = ReleaseDC(GetDesktopWindow(), hSrcDC)
    
    mvarCurrentHorizontal = GetSystemMetrics(SM_CXSCREEN)
    mvarCurrentVertical = GetSystemMetrics(SM_CYSCREEN)
End Sub

Private Sub Class_Initialize()
    GetCurrenctVideoInformation
End Sub
