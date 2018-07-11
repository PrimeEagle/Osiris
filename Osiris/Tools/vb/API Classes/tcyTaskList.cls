VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTaskList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'SAMPLE CODE
'
'Private Sub Command1_Click()
'    Dim x As New clsTaskList
'    Dim y As New Collection
'    Dim Lp As Variant
'
'    Set y = x.TaskList(Form1.hwnd)
'
'    For Each Lp In y
'        If Left$(Lp, 1) <> Chr$(0) Then
'            Debug.Print Lp
'        End If
'    Next
'
'    Set x = Nothing
'End Sub
'

Option Explicit

Private Const GW_CHILD = 5
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDLAST = 1
Private Const GW_HWNDNEXT = 2
Private Const GW_HWNDPREV = 3
Private Const GW_OWNER = 4

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Event CurrentWindow(currwnd As Long, windowtext As String)

'local variable(s) to hold property value(s)
Private mvarTaskList As New Collection 'local copy

Public Property Get TaskList(hwndStart As Long) As Collection
    Dim currwnd As Long
    Dim currhwnd As Long
    'Dim hwndStart As Long
    Dim windowtext As String
    Dim length As Long
    
    Set mvarTaskList = Nothing
    
    'hwndStart = GetDesktopWindow()
    currwnd = GetWindow(hwndStart, GW_HWNDFIRST)
    
    While currwnd <> 0
        'length = GetWindowTextLength(currwnd)
        windowtext = Space$(255)
    
        'windowtext = Space$(length + 1)
        length = GetWindowText(currwnd, windowtext, 255)
    
        If Left$(windowtext, 1) <> Chr$(0) Then
            mvarTaskList.Add windowtext
        End If
    
        currwnd = GetWindow(currwnd, GW_HWNDNEXT)
        DoEvents
    Wend
    
    Set TaskList = mvarTaskList
End Property



