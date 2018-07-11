VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Login"
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   2100
      TabIndex        =   5
      Tag             =   "Cancel"
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   495
      TabIndex        =   4
      Tag             =   "OK"
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1305
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   525
      Width           =   2325
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1305
      TabIndex        =   3
      Top             =   135
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   248
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Tag             =   "&Password:"
      Top             =   540
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   248
      Index           =   0
      Left            =   105
      TabIndex        =   2
      Tag             =   "&User Name:"
      Top             =   150
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpbuffer As String, nSize As Long) As Long

Public OK As Boolean    'true if login successful, else false

Private Sub Form_Load()
    Dim sBuffer As String
    Dim lSize As Long


    sBuffer = Space$(255)               'get username from system
    lSize = Len(sBuffer)
    Call GetUserName(sBuffer, lSize)
    If lSize > 0 Then                   'if name returned is not zero length
        txtUserName.text = Left$(sBuffer, lSize)  'place name in textbox
    Else
        txtUserName.text = vbNullString 'no username, so set to nullstring
    End If
End Sub

Private Sub cmdCancel_Click()
    OK = False
    Me.Hide   'don't unload here since mMain needs to check the OK boolean
End Sub

Private Sub cmdOK_Click()
    'To Do - create test for correct password
    'check for correct password
    If txtPassword.text = "" Then
        OK = True
        Me.Hide 'don't unload here since mMain needs to check the OK boolean
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
        txtPassword.SelStart = 0
        txtPassword.SelLength = Len(txtPassword.text)
    End If
End Sub

