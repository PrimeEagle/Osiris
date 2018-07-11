VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgressBar 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5445
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cbAbort 
      Cancel          =   -1  'True
      Caption         =   "ABORT!"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.PictureBox picProgBar 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   120
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   120
      Width           =   540
   End
   Begin MSComctlLib.ProgressBar pbPBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblPBar 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "frmProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbAbort_Click()
    On Error Resume Next
    dbase.Close
    End
End Sub
