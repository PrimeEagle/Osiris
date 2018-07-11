VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDatabase 
   Caption         =   "Database Tools"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4230
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   4230
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cbOK 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   3735
      Width           =   975
   End
   Begin VB.CommandButton cbCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   3735
      Width           =   975
   End
   Begin VB.CommandButton cbApply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   3735
      Width           =   975
   End
   Begin VB.CommandButton cbHelp 
      Caption         =   "&Help"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   3735
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3600
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   6350
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Add Table"
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Delete Table"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Maintenance"
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame1(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   3195
         Index           =   2
         Left            =   75
         TabIndex        =   13
         Top             =   315
         Width           =   4005
         Begin VB.CommandButton cbRepair 
            Caption         =   "&Repair Database"
            Height          =   390
            Left            =   240
            TabIndex        =   15
            Top             =   600
            Width           =   1545
         End
         Begin VB.CommandButton cbCompact 
            Caption         =   "&Compact Database"
            Height          =   390
            Left            =   2160
            TabIndex        =   14
            Top             =   600
            Width           =   1545
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3195
         Index           =   1
         Left            =   -74925
         TabIndex        =   6
         Top             =   315
         Width           =   4005
         Begin MSComctlLib.ImageCombo cboDelete 
            Height          =   330
            Left            =   1245
            TabIndex        =   12
            Top             =   1395
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
            Text            =   "ImageCombo1"
         End
         Begin VB.Label lblDelete 
            Caption         =   "Table Name:"
            Height          =   285
            Left            =   225
            TabIndex        =   11
            Top             =   1440
            Width           =   960
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3195
         Index           =   0
         Left            =   -74925
         TabIndex        =   5
         Top             =   315
         Width           =   4005
         Begin VB.TextBox tbAdd 
            Height          =   315
            Left            =   1305
            TabIndex        =   8
            Top             =   1905
            Width           =   2430
         End
         Begin MSComctlLib.ImageCombo cboAdd 
            Height          =   330
            Left            =   1320
            TabIndex        =   7
            Top             =   720
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
            Text            =   "ImageCombo1"
         End
         Begin VB.Label Label2 
            Caption         =   "Table Name:"
            Height          =   285
            Left            =   225
            TabIndex        =   10
            Top             =   1950
            Width           =   990
         End
         Begin VB.Label Label1 
            Caption         =   "Table Type:"
            Height          =   255
            Left            =   255
            TabIndex        =   9
            Top             =   765
            Width           =   960
         End
      End
   End
End
Attribute VB_Name = "frmDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbApply_Click()
    Dim tempstr As String
    Dim tempstrU As String
    Dim response As Integer
    Dim i As Integer
    
    Select Case SSTab1.Tab
        Case 0 'Add Table
            cboAdd.SetFocus
            SendKeys "{END}", True
            SendKeys "+{HOME}", True
            tempstr = tbAdd.text
            tempstrU = UCase(tempstr)
            For i = 0 To dbase.TableDefs.count - 1
                If tempstrU = UCase(dbase.TableDefs(i).Name) Then
                    MsgBox "There is already a table called '" & tempstr & "'!", _
                        vbOKOnly + vbExclamation
                        tbAdd.text = ""
                    GoTo Done
                End If
            Next i
            CopyTable dbase, cboAdd.SelectedItem.text, tempstr
            tbAdd.text = ""
            MsgBox "The table '" & tempstr & "' was created.", _
                    vbInformation + vbOKOnly
            TableComboNeedsRefresh = True
        Case 1 'Delete Table
            tempstr = cboDelete.SelectedItem.text
            DeleteTable dbase, tempstr
            TableComboNeedsRefresh = True
        Case 2 'Database Maintenance
            cbOK.Enabled = False
            cbApply.Enabled = False
    End Select
Done:
    BuildAdd
    BuildDelete
End Sub

Private Sub cbCancel_Click()
    Unload Me
End Sub

Private Sub cbCompact_Click()
    Me.MousePointer = vbHourglass
    Set dbase = CompactDB(CurrentDatabaseFile, dbase, True, fProgForm)
    If Not dbase Is Nothing Then
        MsgBox "The database was compacted successfully.", vbOKOnly + vbInformation
    End If
    Me.MousePointer = vbArrow
End Sub

Private Sub cbOK_Click()
    cbApply_Click
    Unload Me
End Sub

Private Sub cbRepair_Click()
    Me.MousePointer = vbHourglass
    Set dbase = RepairDB(CurrentDatabaseFile, dbase, True, fProgForm)
    If Not dbase Is Nothing Then
        MsgBox "The database was repaired successfully.", vbOKOnly + vbInformation
    End If
    Me.MousePointer = vbArrow
End Sub

Private Sub Form_Load()
    BuildAdd
    BuildDelete
    SSTab1.Tab = 0
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
    BuildAdd
    BuildDelete
    Select Case SSTab1.Tab
        Case 0 'Add
            If Len(tbAdd.text) > 0 Then
                cbOK.Enabled = True
                cbApply.Enabled = True
            Else
                cbOK.Enabled = False
                cbApply.Enabled = False
            End If
            cbApply.Caption = "&Add"
        Case 1 'Delete
            cboDelete.SelectedItem = cboDelete.ComboItems(1)
            If cboDelete.SelectedItem.text = "NO VALID TABLES!" Then
                cbOK.Enabled = False
                cbApply.Enabled = False
            Else
                cbOK.Enabled = True
                cbApply.Enabled = True
            End If
            cbApply.Caption = "&Delete"
        Case 2 'Maintenance
            cbOK.Enabled = False
            cbApply.Enabled = False
    End Select
End Sub


Private Sub tbAdd_Change()
    If Len(tbAdd.text) > 0 Then
        cbOK.Enabled = True
        cbApply.Enabled = True
    Else
        cbOK.Enabled = False
        cbApply.Enabled = False
    End If
End Sub

Private Sub BuildDelete()
    Dim i As Integer
    Dim TestIndex As Long
    
    'Set up the Delete tab
    cboDelete.ComboItems.Clear
    Set cboDelete.ImageList = Nothing
    On Error Resume Next
    Set cboDelete.ImageList = fMainForm.imlMenu
    On Error GoTo 0
    For i = 0 To dbase.TableDefs.count - 1
        If Not (UCase(mID$(dbase.TableDefs(i).Name, 1, 4)) = "MSYS" _
            Or UCase(mID$(dbase.TableDefs(i).Name, 1, 3)) = "TM_" _
            Or UCase(mID$(dbase.TableDefs(i).Name, 1, 3)) = "GS_" _
            Or UCase(mID$(dbase.TableDefs(i).Name, 1, 1)) = "~" _
            Or dbase.TableDefs(i).Name = DB_DefaultDataTable _
            Or dbase.TableDefs(i).Name = DB_InboxTable _
            Or dbase.TableDefs(i).Name = DB_NodeTable) Then
                On Error Resume Next
                TestIndex = -1
                TestIndex = fMainForm.imlMenu.ListImages.item("AccessTable").Index
                If TestIndex = -1 Then
                    TestIndex = 0
                End If
                On Error GoTo 0
                cboDelete.ComboItems.Add , UCase(dbase.TableDefs(i).Name), _
                    dbase.TableDefs(i).Name, TestIndex, TestIndex, 0
        End If
    Next i
    If cboDelete.ComboItems.count = 0 Then
        On Error Resume Next
        TestIndex = -1
        TestIndex = fMainForm.imlMenu.ListImages.item("AccessTable").Index
        If TestIndex = -1 Then
            TestIndex = 0
        End If
        On Error GoTo 0
        cboDelete.ComboItems.Add , "NO VALID TABLES!", "NO VALID TABLES!", _
            TestIndex, TestIndex, 0
    End If
    cboDelete.SelectedItem = cboDelete.ComboItems(1)
    If cboDelete.SelectedItem.text = "NO VALID TABLES!" Then
        cbOK.Enabled = False
        cbApply.Enabled = False
    Else
        cbOK.Enabled = True
        cbApply.Enabled = True
    End If
End Sub

Private Sub BuildAdd()
    Dim i As Long
    Dim TestIndex As Long
    
    ' Set up the Add tab
    On Error Resume Next
    Set cboAdd.ImageList = fMainForm.imlMenu
    On Error GoTo 0
    cboAdd.ComboItems.Clear
    For i = 0 To dbase.TableDefs.count - 1
        If UCase(mID$(dbase.TableDefs(i).Name, 1, 3)) = "TM_" Then
            On Error Resume Next
            TestIndex = -1
            TestIndex = fMainForm.imlMenu.ListImages.item("AccessTable").Index
            If TestIndex = -1 Then
                TestIndex = 0
            End If
            On Error GoTo 0
            cboAdd.ComboItems.Add , UCase(dbase.TableDefs(i).Name), _
                dbase.TableDefs(i).Name, TestIndex, TestIndex, 0
        End If
    Next i
    cboAdd.SelectedItem = cboAdd.ComboItems(1)
End Sub
