VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clstoolBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Const WM_USER = &H400
Private Const TB_SETSTYLE = WM_USER + 56
Private Const TB_GETSTYLE = WM_USER + 57
Private Const TBSTYLE_FLAT = &H800
 
'local variable(s) to hold property value(s)
Private mvarFlat    As Boolean 'local copy
Private mvarToolbar As Object 'local copy

Public Property Let Toolbar(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Toolbar = Form1
    Set mvarToolbar = vData
End Property


Public Property Get Toolbar() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Toolbar
    Set Toolbar = mvarToolbar
End Property


Public Property Let Flat(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Flat = 5
   
   Dim Style As Long
   Dim hToolbar As Long
   Dim r As Long
   
  If Not (mvarToolbar Is Nothing) Then
    mvarFlat = vData
   
    'get the handle of the toolbar
     hToolbar = FindWindowEx(mvarToolbar.hwnd, 0&, "ToolbarWindow32", vbNullString)
     
    'retrieve the toolbar styles
     'style = SendMessageLong(hToolbar, TB_GETSTYLE, 0&, 0&)
     
    'Set the new style flag
     If mvarFlat Then
        Style = 1442841185
     Else
        Style = 1442843233
     End If
     
     If Style And TBSTYLE_FLAT Then
        Style = Style Xor TBSTYLE_FLAT
     Else
        Style = Style Or TBSTYLE_FLAT
     End If
     
    'apply the new style to the toolbar
     r = SendMessageLong(hToolbar, TB_SETSTYLE, 0, Style)
     mvarToolbar.Refresh
  End If
End Property


Public Property Get Flat() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Flat
    Flat = mvarFlat
End Property



