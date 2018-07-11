VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_GETDROPPEDWIDTH = &H15F
Private Const CB_SETDROPPEDWIDTH = &H160

Private mvarComboBox As Object 'local copy
Private mvarListWidth As Long 'local copy

Public Property Let ListWidth(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ListWidth = 5
    Dim R As Long
    
    If Not (mvarComboBox Is Nothing) Then
        mvarListWidth = vData
        R = SendMessageLong(mvarComboBox.hwnd, CB_SETDROPPEDWIDTH, mvarListWidth, 0)
    End If
End Property


Public Property Get ListWidth() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ListWidth
    If Not (mvarComboBox Is Nothing) Then
        mvarListWidth = SendMessageLong(mvarComboBox.hwnd, CB_GETDROPPEDWIDTH, 0, 0)
        ListWidth = mvarListWidth
    End If
End Property




Public Sub HideList()
    Dim R As Long
    
    If Not (mvarComboBox Is Nothing) Then
        R = SendMessageLong(mvarComboBox.hwnd, CB_SHOWDROPDOWN, False, 0)
    End If
End Sub

Public Sub ShowList()
    Dim R As Long
    
    If Not (mvarComboBox Is Nothing) Then
        R = SendMessageLong(mvarComboBox.hwnd, CB_SHOWDROPDOWN, True, 0)
    End If
End Sub

Public Property Let ComboBox(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ComboBox = Form1
    Set mvarComboBox = vData
End Property


Public Property Get ComboBox() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ComboBox
    Set ComboBox = mvarComboBox
End Property



