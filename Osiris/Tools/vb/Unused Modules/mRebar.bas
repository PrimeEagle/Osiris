Attribute VB_Name = "mRebar"
Option Explicit

Public Type RBHITTESTINFO
    ptApi As POINTAPI
    flags As Long
    iBand As Long
End Type

Public Type NMREBAR
    NMHDR As Long
    uBand As Long
    wID As Long
    cyChild As Long
    cyBand As Long
End Type

Public Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Type REBARINFO
       cbSize As Integer
       fMask As Integer
       himl As Long
End Type

Type REBARBANDINFO
       cbSize As Integer
       fMask As Integer
       fStyle As Integer
       clrFore As Long
       clrBack As Long
       lpText As String
       cch As Integer
       iImage As Integer
       hWndChild As Long
       cxMinChild As Integer
       cyMinChild As Integer
       cx As Integer
       hbmBack As Long
       wID As Integer
       cyChild As Integer
       cyMaxChild As Integer
       cyIntegral As Integer
       cxIdeal As Integer
       lParam As Long
       cxHeader As Integer
End Type

Type tagInitCommonControlsEx   ' icc
    dwSize As Long    ' size of this structure
    dwICC As Long     ' flags indicating which classes to be initialized
End Type


' Ensures that the common control dynamic-link library (DLL) is loaded.
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
' IE3 & later Returns True (non-zero) if successful, or False otherwise.
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As tagInitCommonControlsEx) As Boolean

Public Sub InitRebar()
    If UseProgressBar Then
        InitProgressBar "Initializing Rebar Controls . . .", 0, 100, _
            LARGE_ICON_PATH & "Interface.bmp", False
    End If
    fMainForm.tbToolBar.Buttons("Open").Image = fMainForm.imlMenu.ListImages.Item("Open").Index
    fMainForm.tbToolBar.Buttons("Copy").Image = fMainForm.imlMenu.ListImages.Item("Copy").Index
    fMainForm.tbToolBar.Buttons("Cut").Image = fMainForm.imlMenu.ListImages.Item("Cut").Index
    fMainForm.tbToolBar.Buttons("Paste").Image = fMainForm.imlMenu.ListImages.Item("Paste").Index
    fMainForm.tbToolBar.Buttons("Print").Image = fMainForm.imlMenu.ListImages.Item("Print").Index
    fMainForm.tbToolBar.Buttons("Help").Image = fMainForm.imlMenu.ListImages.Item("Help").Index
    fMainForm.tbToolBar.Buttons("Delete").Image = fMainForm.imlMenu.ListImages.Item("Delete").Index
    fMainForm.tbToolBar.Buttons("Prop").Image = fMainForm.imlMenu.ListImages.Item("Prop").Index
    fMainForm.tbToolBar.Buttons("Security").Image = fMainForm.imlMenu.ListImages.Item("Security").Index
    fMainForm.tbToolBar.Buttons("Rename").Image = fMainForm.imlMenu.ListImages.Item("Security").Index
    
'    Rebar.TBMakeFlat fMainForm.tbToolBar
   'Create The Rebar
'    With Rebar
'    Set .Parent = fMainForm
'        .Create
'    End With
 
    'Add the bands with the child
'    Rebar.AddBands "", 0, fMainForm.picTBContainer.hWnd, 0, 10
    If UseProgressBar Then
        fProgForm.Hide
    End If
End Sub
