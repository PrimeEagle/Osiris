VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Enum TProgressColor
    Black = 0
    blue = 1
    Green = 2
    Cyan = 3
    Red = 4
    Magenta = 5
    Yellow = 6
    White = 7
    Gray = 8
    LtBlue = 9
    LtGreen = 10
    LtCyan = 11
    LtRed = 12
    LtMagenta = 13
    LtYellow = 14
    BrWhite = 15
End Enum

'local variable(s) to hold property value(s)
Private mvarPictureBox As Object 'local copy
Private mvarPercent    As Byte 'local copy
Private mvarCaption    As Boolean 'local copy
Private mvarProgressColor      As TProgressColor 'local copy
Private mvarForeColor  As TProgressColor 'local copy
Public Property Let ForeColor(ByVal vData As TProgressColor)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ForeColor = 5
    mvarForeColor = vData
End Property


Public Property Get ForeColor() As TProgressColor
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ForeColor
    ForeColor = mvarForeColor
End Property




Public Property Let Color(ByVal vData As TProgressColor)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Color = 5
    mvarProgressColor = vData
End Property


Public Property Get Color() As TProgressColor
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Color
    Color = mvarProgressColor
End Property


Public Sub Reset()
    If Not (mvarPictureBox Is Nothing) Then
        mvarPictureBox.Cls
        mvarPercent = 0
    End If
End Sub

Public Property Let Caption(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Caption = 5
    mvarCaption = vData
End Property


Public Property Get Caption() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Caption
    Caption = mvarCaption
End Property




Public Property Let Percent(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Percent = 5
    Dim Right As Long
    
    If Not (mvarPictureBox Is Nothing) Then
        If vData <= 100 Then
            mvarPercent = vData
            
            Right = (mvarPictureBox.Width * (mvarPercent / 100))
            
            mvarPictureBox.Cls
            mvarPictureBox.Line (0, 0)-(Right, mvarPictureBox.Height - 75), QBColor(mvarProgressColor), BF
            
            If mvarCaption Then
                mvarPictureBox.CurrentX = (mvarPictureBox.Width - 600) / 2
                mvarPictureBox.CurrentY = (mvarPictureBox.Height - 250) / 2
                
                mvarPictureBox.ForeColor = QBColor(mvarForeColor)
                mvarPictureBox.Print Format$((mvarPercent / 100), "Percent")
            End If
            
            'mvarPictureBox.Parent.Refresh
            'mvarPictureBox.Refresh
        End If
    End If
End Property


Public Property Get Percent() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Percent
    Percent = mvarPercent
    
    
End Property



Public Property Let PictureBox(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.PictureBox = Form1
    Set mvarPictureBox = vData
End Property


Public Property Get PictureBox() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PictureBox
    Set PictureBox = mvarPictureBox
End Property



Private Sub Class_Initialize()
    mvarPercent = 0
    mvarProgressColor = blue
    mvarCaption = False
End Sub
