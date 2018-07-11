VERSION 5.00
Begin VB.Form frmHTMLEditAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Osiris HTML Editor"
   ClientHeight    =   3090
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4935
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2131.518
   ScaleMode       =   0  'User
   ScaleWidth      =   4626.559
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      ScaleHeight     =   336.791
      ScaleMode       =   0  'User
      ScaleWidth      =   336.791
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1800
      TabIndex        =   0
      Top             =   2520
      Width           =   1260
   End
   Begin VB.Label lblDescription 
      Caption         =   "Copyright 1998 HDMA Software.  All rights reserved."
      ForeColor       =   &H00000000&
      Height          =   444
      Left            =   960
      TabIndex        =   2
      Top             =   1128
      Width           =   3852
   End
   Begin VB.Label lblTitle 
      Caption         =   "Osiris HTML Editor"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   960
      TabIndex        =   4
      Top             =   240
      Width           =   2436
   End
   Begin VB.Label lblVersion 
      Caption         =   "Non-Release Test Version"
      Height          =   228
      Left            =   960
      TabIndex        =   5
      Top             =   780
      Width           =   2436
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Disclaimer:  yada, yada, yada........"
      ForeColor       =   &H00000000&
      Height          =   348
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   4572
   End
End
Attribute VB_Name = "frmHTMLEditAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  Unload Me
End Sub


