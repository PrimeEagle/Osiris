VERSION 5.00
Begin VB.Form frmHTMLSourceAbout 
   Caption         =   "About Osiris Source Viewer"
   ClientHeight    =   3072
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   4908
   LinkTopic       =   "Form1"
   ScaleHeight     =   3072
   ScaleWidth      =   4908
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1680
      TabIndex        =   1
      Top             =   2520
      Width           =   1260
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   432
      Left            =   120
      ScaleHeight     =   263.118
      ScaleMode       =   0  'User
      ScaleWidth      =   263.118
      TabIndex        =   0
      Top             =   240
      Width           =   432
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Disclaimer:  yada, yada, yada........"
      ForeColor       =   &H00000000&
      Height          =   348
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   4572
   End
   Begin VB.Label lblVersion 
      Caption         =   "Non-Release Test Version"
      Height          =   228
      Left            =   840
      TabIndex        =   4
      Top             =   780
      Width           =   2436
   End
   Begin VB.Label lblTitle 
      Caption         =   "Osiris Source Viewer"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   840
      TabIndex        =   3
      Top             =   240
      Width           =   2436
   End
   Begin VB.Label lblDescription 
      Caption         =   "Copyright 1998 HDMA Software.  All rights reserved."
      ForeColor       =   &H00000000&
      Height          =   444
      Left            =   840
      TabIndex        =   2
      Top             =   1128
      Width           =   3852
   End
End
Attribute VB_Name = "frmHTMLSourceAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub
