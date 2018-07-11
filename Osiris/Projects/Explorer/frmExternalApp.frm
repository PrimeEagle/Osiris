VERSION 5.00
Begin VB.Form frmExternalApp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Launching External Application"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cbOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   2880
      Width           =   975
   End
   Begin VB.CheckBox chkAlways 
      Caption         =   "Show this dialog next time."
      Height          =   195
      Left            =   1080
      TabIndex        =   0
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblWarning 
      Caption         =   "WARNING:  You are about to launch an external application!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label lblExternal 
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   4095
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmExternalApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbOK_Click()
    SaveSetting App.EXEName, "Options", "Show External App Warning", _
        chkAlways.Value
    Unload Me
End Sub

Private Sub Form_Load()
    chkAlways.Value = GetSetting(App.EXEName, "Options", "Show External App Warning", 1)
End Sub
