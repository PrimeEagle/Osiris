VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Health Log"
   ClientHeight    =   8280
   ClientLeft      =   165
   ClientTop       =   405
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   9540
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   14208
      _Version        =   393216
      TabOrientation  =   3
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   706
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Physical Profile"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Supplements"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblSup(0)"
      Tab(1).Control(1)=   "tbSup(0)"
      Tab(1).Control(2)=   "cboSup(0)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Diet"
      TabPicture(2)   =   "frmMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Exercise"
      TabPicture(3)   =   "frmMain.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Journal"
      TabPicture(4)   =   "frmMain.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Illnesses && Injuries"
      TabPicture(5)   =   "frmMain.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      Begin VB.ComboBox cboSup 
         Height          =   315
         Index           =   0
         Left            =   -72840
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   60
         Width           =   975
      End
      Begin VB.TextBox tbSup 
         Height          =   300
         Index           =   0
         Left            =   -73920
         TabIndex        =   2
         Top             =   60
         Width           =   975
      End
      Begin VB.Label lblSup 
         Caption         =   "Vitamin C:"
         Height          =   195
         Index           =   0
         Left            =   -74775
         TabIndex        =   1
         Top             =   60
         Width           =   705
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuGraphs 
      Caption         =   "&Graphs"
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsUsers 
         Caption         =   "&Users"
         Begin VB.Menu mnuToolsUsersAdd 
            Caption         =   "&New User..."
            Shortcut        =   ^N
         End
         Begin VB.Menu mnuToolsUsersModify 
            Caption         =   "&Modify User..."
            Shortcut        =   ^M
         End
         Begin VB.Menu mnuToolsUsersDelete 
            Caption         =   "&Delete User..."
            Shortcut        =   ^D
         End
      End
      Begin VB.Menu mnuToolBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsNutrients 
         Caption         =   "&Nutrients"
         Begin VB.Menu mnuToolsNutrientsAdd 
            Caption         =   "&Add..."
         End
         Begin VB.Menu mnuToolsNutrientsModify 
            Caption         =   "&Modify..."
         End
         Begin VB.Menu mnuToolsNutrientsDelete 
            Caption         =   "&Delete..."
         End
      End
      Begin VB.Menu mnuToolBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsSup 
         Caption         =   "&Supplements"
         Begin VB.Menu mnuToolsSupAdd 
            Caption         =   "&Add..."
         End
         Begin VB.Menu mnuToolsSupModify 
            Caption         =   "&Modify..."
         End
         Begin VB.Menu mnuToolsSupDelete 
            Caption         =   "&Delete..."
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Health Log..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const MAX_PER_COL = 18


Private Sub Form_Load()
    Set dbase = OpenDatabase(CurrentDBase)
End Sub


Private Sub ShowSupplements()
    Dim record As Recordset
    Dim nutr_count As Long
    Dim supp_count As Long
    Dim unit_count As Long
    Dim i As Long
    Dim j As Long
    Dim t As Long
    Dim NewTop As Long
    Dim NewLeft As Long


    Me.MousePointer = vbHourglass
    
    Set record = dbase.OpenRecordset("SELECT COUNT (*) AS [Count] FROM " _
        & "HL_Nutrients", dbOpenDynaset)
    nutr_count = record!Count
    record.Close
    
    Set record = dbase.OpenRecordset("SELECT COUNT (*) AS [Count] FROM " _
        & "HL_Units", dbOpenDynaset)
    unit_count = record!Count
    record.Close
    
    Set record = dbase.OpenRecordset("SELECT * FROM HL_Units ORDER BY [Unit]", dbOpenDynaset)
    While Not record.EOF
        cboSup(0).AddItem record![Unit]
        record.MoveNext
    Wend
    record.Close
    
    Set record = dbase.OpenRecordset("SELECT * FROM HL_Nutrients" _
        & " WHERE [Display] = TRUE ORDER BY [Name]", dbOpenDynaset)
    
    t = 1
    num_rows = 0
    i = 1
    While Not record.EOF
        
        If num_rows Mod MAX_PER_COL = 0 And num_rows <> 0 Then
            NewTop = cboSup(0).Top
            NewLeft = cboSup(i - 1).Left + cboSup(i - 1).Width + 100
        Else
            NewTop = cboSup(i - 1).Top + cboSup(i - 1).Height + 100
            NewLeft = lblSup(i - 1).Left
        End If
        
        Load lblSup(i)
        lblSup(i).Visible = True
        lblSup(i).Height = lblSup(i - 1).Height
        lblSup(i).Width = lblSup(i - 1).Width
        lblSup(i).Top = NewTop
        lblSup(i).Left = NewLeft
        lblSup(i).Caption = record![Name]
        
        Load tbSup(i)
        tbSup(i).Visible = True
        tbSup(i).Height = tbSup(i - 1).Height
        tbSup(i).Width = tbSup(i - 1).Width
        tbSup(i).Top = NewTop
        tbSup(i).Left = lblSup(i).Left + lblSup(i).Width + 50
        tbSup(i).TabIndex = t
        t = t + 1
        
        
        Load cboSup(i)
        cboSup(i).Visible = True
        cboSup(i).Width = cboSup(i - 1).Width
        cboSup(i).Top = NewTop
        cboSup(i).Left = tbSup(i).Left + tbSup(i).Width + 50
        For j = 0 To unit_count - 1
            cboSup(i).AddItem cboSup(0).List(j)
        Next j
        cboSup(i).ListIndex = 0
        cboSup(i).TabIndex = t
        t = t + 1
        num_rows = num_rows + 1
        i = i + 1
        record.MoveNext
    Wend
    
    lblSup(0).Visible = False
    tbSup(0).Visible = False
    cboSup(0).Visible = False
    record.Close
    Me.MousePointer = vbArrow
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuToolsNutrientsAdd_Click()
    UserAction = "Add"
    frmNutrient.Show vbModal
End Sub

Private Sub mnuToolsNutrientsDelete_Click()
    UserAction = "Delete"
    frmNutrient.Show vbModal
End Sub

Private Sub mnuToolsNutrientsModify_Click()
    UserAction = "Modify"
    frmNutrient.Show vbModal
End Sub

Private Sub mnuToolsUsersAdd_Click()
    UserAction = "Add"
    frmUsers.Show vbModal
End Sub

Private Sub mnuToolsUsersDelete_Click()
    UserAction = "Delete"
    frmUsers.Show vbModal
End Sub

Private Sub mnuToolsUsersModify_Click()
    UserAction = "Modify"
    frmUsers.Show vbModal
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Caption = "Supplements" And Not SupLoaded Then
        If NeedToClearSup Then
            ClearSupControls
            NeedToClearSup = False
        End If
        ShowSupplements
        SupLoaded = True
    End If
End Sub


Private Sub tbSup_GotFocus(Index As Integer)
    tbSup(Index).SelStart = 0
    tbSup(Index).SelLength = Len(tbSup(Index).Text)
End Sub

Private Sub tbSup_LostFocus(Index As Integer)
    If IsNumeric(tbSup(Index).Text) Then
        tbSup(Index).Text = CLng(tbSup(Index).Text)
    Else
        If Not tbSup(Index).Text = "" Then
            MsgBox "You must enter a valid number or clear the value!", vbExclamation
            tbSup(Index).SetFocus
        End If
    End If
End Sub

Public Sub ClearSupControls()
    Dim i As Long
    
    For i = 1 To num_rows
        Unload lblSup(i)
        Unload tbSup(i)
        Unload cboSup(i)
    Next i
End Sub
