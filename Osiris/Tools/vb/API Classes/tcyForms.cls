VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsForms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'
'COMMENTS: This cannot be reversed (that I know of).  If you know the
'          code to make a form normal again after setting it to
'          transparent, let me know!
'

Option Explicit
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long

Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Const WS_EX_TRANSPARENT = &H20&
Private Const GWL_EXSTYLE = (-20)

'local variable(s) to hold property value(s)
Private mvarForm        As Object 'local copy
Private mvarTransparent As Boolean 'local copy
Private mvarFlash       As Boolean 'local copy

Public Property Let Flash(ByVal vData As Boolean)
    Dim Result As Long
    
    If Not (mvarForm Is Nothing) Then
        mvarFlash = vData
        
        Result = FlashWindow(mvarForm.hwnd, True)
    Else
        MsgBox "clsForms: Form property not set", vbOKOnly + vbCritical, "Not supported"
    End If
End Property

Public Property Get Flash() As Boolean
    Flash = mvarFlash
End Property

Public Property Let Transparent(ByVal vData As Boolean)
    Dim Result As Long
    
    If Not (mvarForm Is Nothing) Then
        mvarTransparent = vData
        
        If mvarTransparent Then
            Result = SetWindowLong(mvarForm.hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
        Else
            MsgBox "Sorry, once a form has been set to transparent, it cannot be reversed.  Windows does not fully support this feature...", vbOKOnly + vbCritical, "Not supported"
        End If
    Else
        MsgBox "clsForms: Form property not set", vbOKOnly + vbCritical, "Not supported"
    End If
End Property

Public Property Get Transparent() As Boolean
    Transparent = mvarTransparent
End Property

Public Property Let Form(ByVal vData As Object)
    Set mvarForm = vData
End Property

Public Property Get Form() As Object
    Set Form = mvarForm
End Property

Public Sub Reshape(T As Long, L As Long, W As Long, H As Long)
    Dim Region1 As Long
    Dim Region2 As Long
    
    If Not (mvarForm Is Nothing) Then
        W = (mvarForm.Width / Screen.TwipsPerPixelX)
        H = (mvarForm.Height / Screen.TwipsPerPixelY)
        
        Region1 = CreateEllipticRgn(T, L, W, H)
        
        Region2 = SetWindowRgn(mvarForm.hwnd, Region1, True)
    Else
        MsgBox "clsForms: Form property not set", vbOKOnly + vbCritical, "Not supported"
    End If
End Sub


