VERSION 5.00
Begin VB.Form frmDBProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Properties"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   4215
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Current User:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label lblCurrentUser 
         AutoSize        =   -1  'True
         Caption         =   "CurrentUser"
         Height          =   195
         Left            =   1395
         TabIndex        =   15
         Top             =   240
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4215
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Path:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Size:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   435
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Last Modified:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "Microsoft Access Database"
         Height          =   195
         Left            =   1395
         TabIndex        =   9
         Top             =   240
         Width           =   1950
      End
      Begin VB.Label lblPath 
         Caption         =   "Path"
         Height          =   435
         Left            =   1395
         TabIndex        =   8
         Top             =   600
         Width           =   2730
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         Caption         =   "Size"
         Height          =   195
         Left            =   1395
         TabIndex        =   7
         Top             =   1080
         Width           =   300
      End
      Begin VB.Label lblModified 
         AutoSize        =   -1  'True
         Caption         =   "Modified"
         Height          =   195
         Left            =   1395
         TabIndex        =   6
         Top             =   1440
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Microsoft Jet Version:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   1035
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DAO Version:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   1155
      End
      Begin VB.Label lblJetVer 
         AutoSize        =   -1  'True
         Caption         =   "Jet Ver"
         Height          =   195
         Left            =   1395
         TabIndex        =   3
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblDAOVer 
         AutoSize        =   -1  'True
         Caption         =   "DAO Ver"
         Height          =   195
         Left            =   1395
         TabIndex        =   2
         Top             =   2280
         Width           =   630
      End
   End
   Begin VB.CommandButton cbOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   3360
      Width           =   975
   End
End
Attribute VB_Name = "frmDBProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim length As Long
    Dim newlength As Double
    Dim tempstr As String
    Dim version As String
    
    lblCurrentUser.Caption = CurrentUser
    
    lblPath.Caption = CurrentDatabaseFile
    length = FileLen(CurrentDatabaseFile)
    tempstr = format$(length, "###,###,###") & " bytes"
    newlength = length / 1024
    If newlength > 1024 Then
        newlength = newlength / 1024
        tempstr = tempstr & " (" & format$(newlength, "###0.00") & " MB)"
    Else
        tempstr = tempstr & " (" & format$(newlength, "###0.00") & " KB)"
    End If
    lblSize.Caption = tempstr
    lblModified.Caption = FileDateTime(CurrentDatabaseFile)
    
    version = dbase.version
    Select Case version
        Case "1.0"
            version = version & " (MS Access 1.0)"
        Case "1.1"
            version = version & " (MS Access 1.1)"
        Case "2.0"
            version = version & " (MS Access 2.0)"
        Case "2.5"
            version = version & " (not used in MS Access)"
        Case "3.0"
            version = version & " (MS Access 7.0/95)"
        Case "3.5"
            version = version & " (MS Access 8.0/97)"
    End Select
    lblJetVer.Caption = version
    lblDAOVer.Caption = DBEngine.version
   
End Sub
