VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{2C4F587F-97F9-11D1-B346-444553540000}#19.0#0"; "CONTROLESDOMPP.OCX"
Begin VB.Form frmMainOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Osiris Options"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MPPControls.BrowseForFolders BrowseDir 
      Height          =   615
      Left            =   1635
      TabIndex        =   20
      Top             =   3735
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   960
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cbOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   3840
      Width           =   1095
   End
   Begin TabDlg.SSTab sstOptions 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   6376
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "General"
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "chkRefreshIcons"
      Tab(0).Control(1)=   "chkExternalApp"
      Tab(0).Control(2)=   "chkReopenDatabase"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Startup"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkQuotes"
      Tab(1).Control(1)=   "chkSplash"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Paths"
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label3"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label4"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "tbTempFolder"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cbChangeTemp"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cbChangeLargeIcon"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "tbLargeIconFolder"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "cbChangeSmallIcon"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "tbSmallIconFolder"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cbChangeDatabase"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "tbDatabase"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).ControlCount=   12
      Begin VB.TextBox tbDatabase 
         Height          =   285
         Left            =   1290
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   600
         Width           =   2535
      End
      Begin VB.CommandButton cbChangeDatabase 
         Caption         =   "Change..."
         Height          =   255
         Left            =   3930
         TabIndex        =   17
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox tbSmallIconFolder 
         Height          =   285
         Left            =   1290
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   2040
         Width           =   2535
      End
      Begin VB.CommandButton cbChangeSmallIcon 
         Caption         =   "Change..."
         Height          =   255
         Left            =   3930
         TabIndex        =   14
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox tbLargeIconFolder 
         Height          =   285
         Left            =   1290
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1560
         Width           =   2535
      End
      Begin VB.CommandButton cbChangeLargeIcon 
         Caption         =   "Change..."
         Height          =   255
         Left            =   3930
         TabIndex        =   11
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton cbChangeTemp 
         Caption         =   "Change..."
         Height          =   255
         Left            =   3930
         TabIndex        =   10
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox tbTempFolder 
         Height          =   285
         Left            =   1290
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.CheckBox chkReopenDatabase 
         Caption         =   "Close/Re-open Database during a 'Refresh'."
         Height          =   375
         Left            =   -74760
         TabIndex        =   7
         Top             =   1560
         Width           =   2895
      End
      Begin VB.CheckBox chkExternalApp 
         Caption         =   "Show Warning when launching an external application."
         Height          =   375
         Left            =   -74760
         TabIndex        =   6
         Top             =   600
         Width           =   2895
      End
      Begin VB.CheckBox chkSplash 
         Caption         =   "Show Splash screen on startup."
         Height          =   255
         Left            =   -74760
         TabIndex        =   5
         Top             =   600
         Width           =   2895
      End
      Begin VB.CheckBox chkQuotes 
         Caption         =   "Show Quote of the Day at startup."
         Height          =   255
         Left            =   -74760
         TabIndex        =   4
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CheckBox chkRefreshIcons 
         Caption         =   "Refresh Icons during a 'Refresh'."
         Height          =   375
         Left            =   -74760
         TabIndex        =   3
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "Startup Database:"
         Height          =   375
         Left            =   210
         TabIndex        =   19
         Top             =   525
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Small Icon Folder:"
         Height          =   390
         Left            =   180
         TabIndex        =   16
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "Large Icon Folder:"
         Height          =   390
         Left            =   180
         TabIndex        =   13
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Temp Folder:"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   1080
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMainOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub cbChangeDatabase_Click()
    CommonDialog1.Filter = "Access Database (*.mdb)|*.mdb"
    CommonDialog1.FilterIndex = 0
    On Error GoTo userCancel
        CommonDialog1.ShowOpen
    On Error GoTo 0
    CurrentDatabaseFile = CommonDialog1.filename
userCancel:
End Sub

Private Sub cbChangeLargeIcon_Click()
    BrowseDir.Título = "Choose a New Large Icon Folder..." 'set the title of the folder browser
    BrowseDir.Mostrar 'shows the folder browser
    If BrowseDir.DiretorioRetornado <> "" Then
        tbLargeIconFolder.text = BrowseDir.DiretorioRetornado & "\"
    End If
End Sub

Private Sub cbChangeSmallIcon_Click()
    BrowseDir.Título = "Choose a New Small Icon Folder..." 'set the title of the folder browser
    BrowseDir.Mostrar 'shows the folder browser
    If BrowseDir.DiretorioRetornado <> "" Then
        tbSmallIconFolder.text = BrowseDir.DiretorioRetornado & "\"
    End If
End Sub

Private Sub cbChangeTemp_Click()
    BrowseDir.Título = "Choose a New Temp Folder..." 'set the title of the folder browser
    BrowseDir.Mostrar 'shows the folder browser
    If BrowseDir.DiretorioRetornado <> "" Then
        tbTempFolder.text = BrowseDir.DiretorioRetornado & "\"
    End If
End Sub

Private Sub cbOK_Click()
    
    Me.MousePointer = vbHourglass

    If Not ValidFolder(tbTempFolder.text) Then
        MsgBox "The folder '" & tbTempFolder.text & "' does not exist!" _
            & Chr(13) & "You must choose a valid folder.", vbCritical
        tbTempFolder.SetFocus
        Me.MousePointer = vbArrow
        Exit Sub
    End If
    
    If Not ValidFolder(tbLargeIconFolder.text) Then
        MsgBox "The folder '" & tbLargeIconFolder.text & "' does not exist!" _
            & Chr(13) & "You must choose a valid folder.", vbCritical
        tbLargeIconFolder.SetFocus
        Me.MousePointer = vbArrow
        Exit Sub
    End If
    
    If Not ValidFolder(tbSmallIconFolder.text) Then
        MsgBox "The folder '" & tbSmallIconFolder.text & "' does not exist!" _
            & Chr(13) & "You must choose a valid folder.", vbCritical
        tbSmallIconFolder.SetFocus
        Me.MousePointer = vbArrow
        Exit Sub
    End If
    
    If Not ValidFile(tbDatabase.text) Then
        MsgBox "The file '" & tbDatabase.text & "' does not exist!" _
            & Chr(13) & "You must choose a valid file.", vbCritical
        tbDatabase.SetFocus
        Me.MousePointer = vbArrow
        Exit Sub
    End If
    
    SaveSetting App.EXEName, "Options", "Show External App Warning", _
        chkExternalApp.Value
    SaveSetting App.EXEName, "Options", "Show Splash", _
        chkSplash.Value
    SaveSetting App.EXEName, "Options", "Show Tips at Startup", _
        chkQuotes.Value
    SaveSetting App.EXEName, "Options", "Refresh Icons", _
        chkRefreshIcons.Value
    SaveSetting App.EXEName, "Options", "Reopen Database", _
        chkReopenDatabase.Value
    SaveSetting App.EXEName, "Options", "Temp Folder", _
        tbTempFolder.text
    SaveSetting App.EXEName, "Options", "Large Icon Folder", _
        tbLargeIconFolder.text
    SaveSetting App.EXEName, "Options", "Small Icon Folder", _
        tbSmallIconFolder.text
    SaveSetting App.EXEName, "Options", "Startup Database", _
        tbDatabase.text
        
    'Clean out the old temp folder before changing it to the new one
    ClearTempDir CurTempFolder
    CurTempFolder = tbTempFolder.text
    LargeIconFolder = tbLargeIconFolder.text
    SmallIconFolder = tbSmallIconFolder.text
    CurrentDatabaseFile = tbDatabase.text
    
    Me.MousePointer = vbArrow
    Unload Me
End Sub


Private Sub Form_Load()
    sstOptions.Tab = 0
    chkExternalApp.Value = GetSetting(App.EXEName, "Options", _
            "Show External App Warning", 1)
    chkSplash.Value = GetSetting(App.EXEName, "Options", _
            "Show Splash", 1)
    chkQuotes.Value = GetSetting(App.EXEName, "Options", _
            "Show Tips at Startup", 1)
    chkRefreshIcons.Value = GetSetting(App.EXEName, "Options", _
            "Refresh Icons", 1)
    chkReopenDatabase.Value = GetSetting(App.EXEName, "Options", _
            "Reopen Database", 1)
    tbTempFolder.text = GetSetting(App.EXEName, "Options", _
            "Temp Folder", "UNKNOWN!")
    tbLargeIconFolder.text = GetSetting(App.EXEName, "Options", _
            "Large Icon Folder", "UNKNOWN!")
    tbSmallIconFolder.text = GetSetting(App.EXEName, "Options", _
            "Small Icon Folder", "UNKNOWN!")
    tbDatabase.text = GetSetting(App.EXEName, "Options", _
            "Startup Database", "UNKNOWN!")
End Sub

Private Sub tbDatabase_GotFocus()
    tbDatabase.SelStart = 0
    tbDatabase.SelLength = Len(tbDatabase.text)
End Sub

Private Sub tbLargeIconFolder_GotFocus()
    tbLargeIconFolder.SelStart = 0
    tbLargeIconFolder.SelLength = Len(tbLargeIconFolder.text)
End Sub

Private Sub tbSmallIconFolder_GotFocus()
    tbSmallIconFolder.SelStart = 0
    tbSmallIconFolder.SelLength = Len(tbSmallIconFolder.text)
End Sub
