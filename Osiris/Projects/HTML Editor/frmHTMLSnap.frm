VERSION 5.00
Begin VB.Form frmHTMLSnap 
   Caption         =   "Snap to Grid Settings"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3090
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   3090
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cbCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cbOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CheckBox cbSnap 
      Caption         =   "&Snap to Grid"
      Height          =   255
      Left            =   645
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox tbY 
      Height          =   330
      Left            =   945
      TabIndex        =   3
      Top             =   825
      Width           =   855
   End
   Begin VB.TextBox tbX 
      Height          =   330
      Left            =   945
      TabIndex        =   2
      Top             =   300
      Width           =   855
   End
   Begin VB.Label lblY 
      Caption         =   "Y:"
      Height          =   255
      Left            =   645
      TabIndex        =   1
      Top             =   863
      Width           =   375
   End
   Begin VB.Label lblX 
      Caption         =   "X:"
      Height          =   255
      Left            =   645
      TabIndex        =   0
      Top             =   338
      Width           =   375
   End
End
Attribute VB_Name = "frmHTMLSnap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbCancel_Click()
    Unload Me
End Sub

Private Sub cbOk_Click()
    frmHTMLEd.LastSnap = cbSnap.Value
    If cbSnap.Value = 1 Then
        frmHTMLEd.LastX = val(tbX.text)
        frmHTMLEd.LastY = val(tbY.text)
    End If
    Unload Me
End Sub

Private Sub cbSnap_Click()
    If cbSnap.Value <> 1 Then
        tbX.Enabled = False
        tbY.Enabled = False
        tbX.BackColor = vbButtonFace
        tbY.BackColor = vbButtonFace
    Else
        tbX.Enabled = True
        tbY.Enabled = True
        tbX.BackColor = vbWindowBackground
        tbY.BackColor = vbWindowBackground
    End If
End Sub

Private Sub Form_Load()
    cbSnap.Value = frmHTMLEd.LastSnap
    tbX.text = format$(frmHTMLEd.LastX)
    tbY.text = format$(frmHTMLEd.LastY)
    If cbSnap.Value <> 1 Then
        tbX.Enabled = False
        tbY.Enabled = False
        tbX.BackColor = vbButtonFace
        tbY.BackColor = vbButtonFace
    Else
        tbX.Enabled = True
        tbY.Enabled = True
        tbX.BackColor = vbWindowBackground
        tbY.BackColor = vbWindowBackground
    End If
End Sub

Private Sub tbX_LostFocus()
    If Not IsNumeric(tbX.text) Then
        MsgBox "There must be a numeric value in this cell!", vbExclamation
        tbX.SetFocus
    Else
        tbX.text = format$(CLng(tbX.text))
        If val(tbX.text) < 0 Or val(tbX.text) > 100 Then
            MsgBox "The value must be in the range of 0-100.", vbExclamation
            tbX.SetFocus
        End If
    End If
End Sub

Private Sub tbY_LostFocus()
    If Not IsNumeric(tbY.text) Then
        MsgBox "There must be a numeric value in this cell!", vbExclamation
        tbY.SetFocus
    Else
        tbY.text = format$(CLng(tbY.text))
        If val(tbY.text) < 0 Or val(tbY.text) > 100 Then
            MsgBox "The value must be in the range of 0-100.", vbExclamation
            tbY.SetFocus
        End If
    End If
End Sub
