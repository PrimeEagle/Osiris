VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Properties for ???"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   5925
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Attach New Binary"
   End
   Begin VB.PictureBox picPicture 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   31
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton cbHelp 
      Caption         =   "&Help"
      Height          =   375
      Left            =   4800
      TabIndex        =   28
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton cbApply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   27
      Top             =   6600
      Width           =   975
   End
   Begin VB.TextBox tbName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1320
      TabIndex        =   26
      Text            =   "Name"
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Frame fraName 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   17
      Top             =   720
      Width           =   4935
   End
   Begin VB.CommandButton cbCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton cbOK 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   6600
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   11245
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabMaxWidth     =   882
      TabCaption(0)   =   "General"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraAttrib"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraHistory"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Data"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraType"
      Tab(1).Control(1)=   "fraValue"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Icons"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraIcons"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "QuickAdd Type"
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraObjects"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Link"
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraLink"
      Tab(4).ControlCount=   1
      Begin VB.Frame fraLink 
         Caption         =   "Available Globals"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   -74640
         TabIndex        =   37
         Top             =   1860
         Width           =   4935
         Begin MSComctlLib.TreeView tvLink 
            Height          =   4005
            Left            =   180
            TabIndex        =   38
            Top             =   300
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   7064
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   353
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Appearance      =   1
         End
         Begin MSComctlLib.ListView lvLink 
            Height          =   4035
            Left            =   2250
            TabIndex        =   39
            Top             =   300
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   7117
            View            =   2
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Frame fraObjects 
         Caption         =   "Available Types"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   -74640
         TabIndex        =   34
         Top             =   1860
         Width           =   4935
         Begin MSComctlLib.TreeView tvQuick 
            Height          =   4005
            Left            =   195
            TabIndex        =   35
            Top             =   300
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   7064
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   353
            LabelEdit       =   1
            Style           =   5
            Appearance      =   1
         End
         Begin MSComctlLib.ListView lvQuick 
            Height          =   4035
            Left            =   2250
            TabIndex        =   36
            Top             =   285
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   7117
            View            =   2
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
      End
      Begin VB.Frame fraIcons 
         Caption         =   "Icons"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   -74640
         TabIndex        =   23
         Top             =   1980
         Width           =   4935
         Begin MSComctlLib.ImageCombo cboIcon 
            Height          =   330
            Index           =   1
            Left            =   1440
            TabIndex        =   24
            Top             =   720
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
            Text            =   "ImageCombo1"
         End
         Begin MSComctlLib.ImageCombo cboIcon 
            Height          =   330
            Index           =   2
            Left            =   1440
            TabIndex        =   29
            Top             =   2160
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Locked          =   -1  'True
            Text            =   "ImageCombo1"
         End
         Begin MSComctlLib.ImageCombo cboIcon 
            Height          =   330
            Index           =   3
            Left            =   1440
            TabIndex        =   32
            Top             =   720
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Text            =   "ImageCombo1"
         End
         Begin VB.Label lblIcon2 
            Caption         =   "Icon 2:"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Label lblIcon1 
            Caption         =   "Icon 1:"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.Frame fraType 
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74640
         TabIndex        =   19
         Top             =   1980
         Width           =   4935
         Begin VB.CommandButton cbDelTable 
            Caption         =   "Delete"
            Height          =   255
            Left            =   3870
            TabIndex        =   33
            Top             =   405
            Width           =   855
         End
         Begin MSComctlLib.ImageCombo cboType 
            Height          =   330
            Left            =   1290
            TabIndex        =   21
            Top             =   345
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Text            =   "ImageCombo1"
         End
         Begin VB.Label lblType 
            Caption         =   "Data Table:"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   390
            Width           =   1110
         End
      End
      Begin VB.Frame fraValue 
         Caption         =   "Value"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   -74640
         TabIndex        =   18
         Top             =   2940
         Width           =   4935
         Begin VB.CommandButton cbSaveAs 
            Caption         =   "&Save As ..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   1560
            TabIndex        =   44
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cbClear 
            Caption         =   "&Clear"
            Enabled         =   0   'False
            Height          =   375
            Left            =   2880
            TabIndex        =   43
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cbEdit 
            Caption         =   "&Edit"
            Height          =   375
            Left            =   240
            TabIndex        =   42
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox tbValue 
            Height          =   2535
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            Top             =   720
            Width           =   4455
         End
         Begin VB.Label lblAttachmentPath 
            Caption         =   "Label1"
            Height          =   975
            Left            =   240
            TabIndex        =   41
            Top             =   960
            Width           =   4455
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame fraHistory 
         Caption         =   "History"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   360
         TabIndex        =   8
         Top             =   1680
         Width           =   4935
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Last Modified:  "
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   720
            Width           =   1080
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Created:  "
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   690
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Created By:  "
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   480
            Width           =   915
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Last Modified By:  "
            Height          =   195
            Left            =   120
            TabIndex        =   50
            Top             =   960
            Width           =   1305
         End
         Begin VB.Label lblModifiedBy 
            AutoSize        =   -1  'True
            Caption         =   "Last Modified By:  "
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1560
            TabIndex        =   12
            Top             =   960
            Width           =   1305
         End
         Begin VB.Label lblCreatedBy 
            AutoSize        =   -1  'True
            Caption         =   "Created By:  "
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1560
            TabIndex        =   11
            Top             =   480
            Width           =   915
         End
         Begin VB.Label lblCreated 
            AutoSize        =   -1  'True
            Caption         =   "Created:  "
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1560
            TabIndex        =   10
            Top             =   240
            Width           =   690
         End
         Begin VB.Label lblLastMod 
            AutoSize        =   -1  'True
            Caption         =   "Last Modified:  "
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1560
            TabIndex        =   9
            Top             =   720
            Width           =   1080
         End
      End
      Begin VB.Frame fraAttrib 
         Caption         =   "Attributes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   360
         TabIndex        =   3
         Top             =   3000
         Width           =   4935
         Begin VB.CheckBox cbVariation 
            Caption         =   "Variation"
            Height          =   255
            Left            =   2520
            TabIndex        =   55
            Top             =   2280
            Width           =   975
         End
         Begin VB.CheckBox cbSublink 
            Caption         =   "Sublink"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2520
            TabIndex        =   47
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CheckBox cbLink 
            Caption         =   "Linked"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   2640
            Width           =   1215
         End
         Begin VB.CheckBox cbCreateNode 
            Caption         =   "Create Node"
            Height          =   255
            Left            =   2520
            TabIndex        =   7
            Top             =   2280
            Width           =   1215
         End
         Begin VB.CheckBox cbCreateItem 
            Caption         =   "Create Item"
            Height          =   255
            Left            =   2520
            TabIndex        =   6
            Top             =   2640
            Width           =   1215
         End
         Begin VB.CheckBox cbSystem 
            Caption         =   "System"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   2280
            Width           =   1695
         End
         Begin VB.CheckBox cbReadOnly 
            Caption         =   "Read Only"
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   1920
            Width           =   1215
         End
         Begin VB.Label lblLinkNode 
            AutoSize        =   -1  'True
            Caption         =   "Link Node:"
            Height          =   195
            Left            =   2520
            TabIndex        =   58
            Top             =   1200
            Width           =   780
         End
         Begin VB.Label lblLinkNodeID 
            AutoSize        =   -1  'True
            Caption         =   "None"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   3600
            TabIndex        =   57
            Top             =   1200
            Width           =   390
         End
         Begin VB.Label lblLinkNodeName 
            AutoSize        =   -1  'True
            Caption         =   "Label7"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   2640
            TabIndex        =   56
            Top             =   1440
            Width           =   480
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblParentName 
            AutoSize        =   -1  'True
            Caption         =   "Label7"
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   360
            TabIndex        =   54
            Top             =   1440
            Width           =   1680
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblGlobalType 
            AutoSize        =   -1  'True
            Caption         =   "None"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1320
            TabIndex        =   49
            Top             =   840
            Width           =   390
         End
         Begin VB.Label lblGlobalLabel 
            AutoSize        =   -1  'True
            Caption         =   "Global Type:"
            Height          =   195
            Left            =   240
            TabIndex        =   48
            Top             =   840
            Width           =   900
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Order:"
            Height          =   195
            Left            =   240
            TabIndex        =   46
            Top             =   600
            Width           =   435
         End
         Begin VB.Label lblOrder 
            AutoSize        =   -1  'True
            Caption         =   "1"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1080
            TabIndex        =   45
            Top             =   600
            Width           =   90
         End
         Begin VB.Label lblParentNodeVal 
            AutoSize        =   -1  'True
            Caption         =   "None"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1320
            TabIndex        =   16
            Top             =   1200
            Width           =   390
         End
         Begin VB.Label lblIDVal 
            AutoSize        =   -1  'True
            Caption         =   "1"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1080
            TabIndex        =   15
            Top             =   360
            Width           =   90
         End
         Begin VB.Label lblParentNode 
            AutoSize        =   -1  'True
            Caption         =   "Parent Node:"
            Height          =   195
            Left            =   240
            TabIndex        =   14
            Top             =   1200
            Width           =   945
         End
         Begin VB.Label lblIDLabel 
            AutoSize        =   -1  'True
            Caption         =   "Node ID:"
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   645
         End
      End
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const NUM_FLAGS = 14
Const tabGeneral = 0
Const tabData = 1
Const tabIcons = 2
Const tabQuick = 3
Const tabLink = 4


Dim nodl As Node
Dim nodq As Node
Dim ChangedFlags(1 To NUM_FLAGS) As Boolean
Dim orig_read_only As Boolean
Dim orig_sublink As Boolean
Dim orig_system_node As Boolean
Dim orig_linked As Boolean
Dim orig_link_node_id As Long
Dim orig_name As String
Dim orig_create_node As Integer
Dim orig_create_item As Integer
Dim orig_icon_normal As String
Dim orig_icon_selected As String
Dim orig_quicktype_id As Long
Dim orig_data_type As String
Dim orig_icon_small As String
Dim orig_icon_large As String
Dim orig_table_name As String

Private Sub cbApply_Click()
    Me.MousePointer = vbHourglass
    
    If FocusFrom = "TreeView" Then
        ApplyTreeView
    ElseIf FocusFrom = "ListView" Then
        ApplyListView
    End If
    
    Me.MousePointer = vbArrow
End Sub

Private Sub cbCancel_Click()
    Me.Hide
    PropertiesActive = False
End Sub

Private Sub cbClear_Click()
    Dim record As Recordset
    
    If vbYes = MsgBox("WARNING:  This will permanently delete the binary attachment from this record!" _
        & Chr(13) & Chr(13) & "Are you sure you want to delete it ?", vbYesNo + vbExclamation) Then
        
        Changed ("cbClear")
        
        Set record = dbase.OpenRecordset("SELECT * FROM " _
            & tvAttribCol(nodx.key).table_name & " WHERE Data_ID = " _
            & RemoveK(CurrentItem.key), dbOpenDynaset)
        record.Edit
        record!data_value = ""
        record!binary_data_value = Null
        record.Update
        record.Close
        MsgBox "Binary attachment for '" & CurrentItem.text & "' was removed.", vbInformation
        LastTab = SSTab1.Tab
        InitPropertiesForm
    End If
End Sub

Private Sub cbCreateItem_Click()
'    Changed ("cbCreateItem")
    If Not cbApply.Enabled Then cbApply.Enabled = True
    If Not cbOK.Enabled Then cbOK.Enabled = True
End Sub

Private Sub cbCreateNode_Click()
'    Changed ("cbCreateNode")
    If Not cbApply.Enabled Then cbApply.Enabled = True
    If Not cbOK.Enabled Then cbOK.Enabled = True
End Sub

Private Sub cbDelTable_Click()
    Dim table_name As String
    Dim deleted As Boolean
    
    cboType.SetFocus
    SendKeys "{END}", True
    SendKeys "+{HOME}", True
    table_name = cboType.SelectedItem.text
    
    selected_table = tvAttribCol(nodx.key).table_name
    
    deleted = DeleteTable(dbase, table_name)
    If deleted Then
        cboType.ComboItems.Remove UCase(table_name)
        cboType.SelectedItem = cboType.ComboItems.item(UCase(selected_table))
    End If
    fMainForm.LoadListView nodx.key
    TableComboNeedsRefresh = True
End Sub

Private Sub cbEdit_Click()
    Dim tempstr As String
    Dim record As Recordset
    Dim attachment_file As String
    
    Changed ("cbEdit")
    
    tempstr = lvAttribCol(CurrentItem.key).data_type
    If UCase(mID$(tempstr, 1, 6)) = "BINARY" Then
        tempstr = "Binary"
    End If
    Select Case tempstr
        Case "Binary"
            On Error GoTo Done
                CommonDialog1.ShowOpen
            On Error GoTo 0
            attachment_file = CommonDialog1.filename
            Set record = dbase.OpenRecordset("SELECT data_value,binary_data_value FROM " _
                & tvAttribCol(nodx.key).table_name & " WHERE Data_ID = " _
                & RemoveK(CurrentItem.key), dbOpenDynaset)
            ReadBLOB attachment_file, record, "binary_data_value"
            record.Close
            MsgBox "'" & attachment_file & "' was added to '" _
                & CurrentItem.text & "' as a binary attachment.", vbInformation
            LastTab = SSTab1.Tab
            InitPropertiesForm
        Case "String", "HTML", "URL"
            MsgBox "TEMP: Load HTML Editor", vbInformation
            'Load frmHTMLEdit
    End Select
Done:
End Sub

Private Sub cbLink_Click()
    If cbLink.Value = 1 Then
        SSTab1.TabEnabled(tabLink) = True
    Else
        SSTab1.TabEnabled(tabLink) = False
    End If
'    Changed ("cbLink")
    If Not cbApply.Enabled Then cbApply.Enabled = True
    If Not cbOK.Enabled Then cbOK.Enabled = True
End Sub

Private Sub cboIcon_Click(Index As Integer)
'    Changed ("cboIcon")
    If Not cbApply.Enabled Then cbApply.Enabled = True
    If Not cbOK.Enabled Then cbOK.Enabled = True
End Sub

Private Sub cbOK_Click()
    cbApply_Click
    PropertiesActive = False
    Me.Hide
End Sub

Private Sub cboType_Change()
'    Changed ("cboType")
    If Not cbApply.Enabled Then cbApply.Enabled = True
    If Not cbOK.Enabled Then cbOK.Enabled = True
End Sub

Private Sub cboType_Click()
'    Changed ("cboType")
    If Not cbApply.Enabled Then cbApply.Enabled = True
    If Not cbOK.Enabled Then cbOK.Enabled = True
    
    If cboType.SelectedItem.text = DB_DefaultDataTable _
            Or cboType.SelectedItem.text = DB_InboxTable Then
        cbDelTable.Enabled = False
    Else
        cbDelTable.Enabled = True
    End If
End Sub

Private Sub cbReadOnly_Click()
'    Changed ("cbReadOnly")
    If Not cbApply.Enabled Then cbApply.Enabled = True
    If Not cbOK.Enabled Then cbOK.Enabled = True
End Sub

Private Sub cbSaveAs_Click()
    Dim attachment_file As String
    Dim record As Recordset
    Dim response As Integer
    
ShowDialog:
    On Error GoTo Done
        CommonDialog1.ShowSave
    On Error GoTo 0
    attachment_file = CommonDialog1.filename
    If Dir(attachment_file) <> "" Then
        response = MsgBox("The file '" & attachment_file & "' already exists." _
            & Chr(13) & "Do you want to overwrite it ?", vbYesNoCancel + vbExclamation)
        Select Case response
            Case vbNo
                GoTo ShowDialog
            Case vbCancel
                GoTo Done
        End Select
    End If
    Set record = dbase.OpenRecordset("SELECT data_value,binary_data_value FROM " _
        & tvAttribCol(nodx.key).table_name & " WHERE Data_ID = " _
        & RemoveK(CurrentItem.key), dbOpenDynaset)
    WriteBLOB record, "binary_data_value", attachment_file
    record.Close
    MsgBox "Attachment saved as '" & attachment_file & "'.", vbInformation
    LastTab = SSTab1.Tab
    InitPropertiesForm
Done:
End Sub





Private Sub cbSystem_Click()
'    Changed ("cbSystem")
    If Not cbApply.Enabled Then cbApply.Enabled = True
    If Not cbOK.Enabled Then cbOK.Enabled = True
End Sub

Private Sub cbVariation_Click()
'    Changed ("cbVariation")
    If Not cbApply.Enabled Then cbApply.Enabled = True
    If Not cbOK.Enabled Then cbOK.Enabled = True
End Sub

Private Sub lvLink_GotFocus()
    tvLink.SetFocus
End Sub

Private Sub lvQuick_GotFocus()
    tvQuick.SetFocus
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    LastTab = SSTab1.Tab
End Sub

Private Sub SSTab1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cbCancel_Click
    End Select
End Sub

Private Sub tbName_Change()
'    Changed ("tbName")
    If Not cbApply.Enabled Then cbApply.Enabled = True
    If Not cbOK.Enabled Then cbOK.Enabled = True
End Sub

Public Sub InitPropertiesForm()
    
    Me.MousePointer = vbHourglass
    If FocusFrom = "TreeView" Then
        UpdateFromTreeView
    ElseIf FocusFrom = "ListView" Then
        UpdateFromListView
    End If
    
    If cbLink.Value = 1 Then
        SSTab1.TabEnabled(tabLink) = True
    Else
        SSTab1.TabEnabled(tabLink) = False
        If SSTab1.Tab = tabLink Then
            LastTab = tabGeneral
        End If
    End If
    
    If tvAttribCol(nodx.key).quicktypeid <> 0 Then
        SSTab1.TabEnabled(tabQuick) = True
    Else
        SSTab1.TabEnabled(tabQuick) = False
        If SSTab1.Tab = tabQuick Then
            LastTab = tabGeneral
        End If
    End If
    
    
    SSTab1.Tab = LastTab

    DoEvents 'this is necessary, so everything is finished updating
             'before we reset all the flags.
    ResetChangedFlags
    
    cbOK.Enabled = False
    cbApply.Enabled = False
    
    Me.MousePointer = vbArrow
End Sub

Private Sub tbName_GotFocus()
    tbName.SelStart = 0
    tbName.SelLength = Len(tbName.text)
End Sub

Private Sub tbName_LostFocus()
    If TypeOf fMainForm.ActiveControl Is TreeView Then
        If fMainForm.ChildNodeHasSameName(nodx.parent.key, tbName.text) And _
            orig_name <> tbName.text Then
                MsgBox "'" & nodx.parent.text & "' already has a node called '" & _
                    tbName.text & "'." & Chr(13) & _
                    "Please give the current node a different name.", vbExclamation
                tbName.SetFocus
        End If
    ElseIf TypeOf fMainForm.ActiveControl Is ListView Then
        If Not fMainForm.lvListView.FindItem(tbName.text) Is Nothing And _
            orig_name <> tbName.text Then
                MsgBox "An item by the name of '" & tbName.text & "' already exists." _
                    & Chr(13) & "Please give the current item a different name.", vbExclamation
                tbName.SetFocus
        End If
    End If
End Sub

Private Sub tbValue_Change()
    Changed ("tbValue")
End Sub


Private Sub BuildQuickTree()
    Dim m As Long
    Dim n As Long
    Dim selected_index As Long
    Dim quicktypeid As Long
    
    selected_index = -1
    
    tvQuick.Nodes.Clear
    tvQuick.ImageList = Nothing
    If fMainForm.imlTempTree.ListImages.count <> 0 Then
        tvQuick.ImageList = fMainForm.imlTempTree
    End If
    
    n = fMainForm.FindANode("QuickAdd Nodes", False, True)
    n = fMainForm.tvTreeView.Nodes(AddK(n)).Index
    CopyNodes_to_TVQuickTree n, 0
    
    If tvQuick.Nodes.count <> 0 Then
        quicktypeid = tvAttribCol(nodx.key).quicktypeid
        If quicktypeid <> 0 Then
            Set tvQuick.SelectedItem = _
                tvQuick.Nodes(AddK(quicktypeid))
            tvQuick.SelectedItem.EnsureVisible
            Set nodq = tvQuick.SelectedItem
            tvQuick_NodeClick nodq
        End If
    End If
    
Done:
End Sub

Private Sub CopyNodes_to_TVQuickTree(ByVal n As Long, ByVal localParent As Long)

    If fMainForm.tvTreeView.Nodes(n).Children > 0 Then
        n = fMainForm.tvTreeView.Nodes(n).Child.Index
        While n <> fMainForm.tvTreeView.Nodes(n).LastSibling.Index
            If localParent = 0 Then
                tvQuick.Nodes.Add , , fMainForm.tvTreeView.Nodes(n).key, _
                    fMainForm.tvTreeView.Nodes(n).text, _
                    fMainForm.tvTreeView.Nodes(n).Image, _
                    fMainForm.tvTreeView.Nodes(n).SelectedImage
            Else
                tvQuick.Nodes.Add localParent, tvwChild, _
                    fMainForm.tvTreeView.Nodes(n).key, _
                    fMainForm.tvTreeView.Nodes(n).text, _
                    fMainForm.tvTreeView.Nodes(n).Image, _
                    fMainForm.tvTreeView.Nodes(n).SelectedImage
            End If
            CopyNodes_to_TVQuickTree n, tvQuick.Nodes.count
            n = fMainForm.tvTreeView.Nodes(n).Next.Index
        Wend
        If localParent = 0 Then
            tvQuick.Nodes.Add , , fMainForm.tvTreeView.Nodes(n).key, _
                fMainForm.tvTreeView.Nodes(n).text, _
                fMainForm.tvTreeView.Nodes(n).Image, _
                fMainForm.tvTreeView.Nodes(n).SelectedImage
        Else
            tvQuick.Nodes.Add localParent, tvwChild, _
                fMainForm.tvTreeView.Nodes(n).key, _
                fMainForm.tvTreeView.Nodes(n).text, _
                fMainForm.tvTreeView.Nodes(n).Image, _
                fMainForm.tvTreeView.Nodes(n).SelectedImage
        End If
        CopyNodes_to_TVQuickTree n, tvQuick.Nodes.count
    End If

End Sub

Private Sub tbValue_GotFocus()
    tbValue.SelStart = 0
    tbValue.SelLength = Len(tbValue.text)
End Sub

Private Sub tvQuick_NodeClick(ByVal Node As MSComctlLib.Node)
    If nodq.key = Node.key Then
        GoTo Done
    End If
    If Not Node.parent Is Nothing Then
        SendKeys "{LEFT}"
        Exit Sub
    End If
    
'    Changed ("tvQuick")
    If Not cbApply.Enabled Then cbApply.Enabled = True
    If Not cbOK.Enabled Then cbOK.Enabled = True
    
    If Not nodq Is Nothing Then
        nodq.Expanded = False
    End If
    Set nodq = Node
    nodq.EnsureVisible
    BuildQuickList
    cbOK.Enabled = True
    cbApply.Enabled = True
Done:
    If nodq.Expanded = False Then
        nodq.Expanded = True
    End If
End Sub

Private Sub BuildQuickList()
    Dim listrecord As Recordset
    Dim dataitem As ListItem
    Dim i As Long
    Dim icon_large As String
    Dim icon_small As String
    Dim tempstr As String
    Dim ParentNodeId As Long
    Dim TestIndex As Long
    Dim TestIndex2 As Long
    
    ParentNodeId = RemoveK(nodq.key)
    
    lvQuick.ListItems.Clear
    Set lvQuick.Icons = fMainForm.imlIconsLarge
    Set lvQuick.SmallIcons = fMainForm.imlIconsSmall
    
    Set listrecord = dbase.OpenRecordset("SELECT * FROM " & DB_QuickAddItemsTable _
        & " WHERE Parent_Node = " & ParentNodeId, dbOpenDynaset)
    i = 1
    While Not listrecord.EOF
    
        icon_large = listrecord!icon_large
        icon_small = listrecord!icon_small
    
        On Error Resume Next
        TestIndex = -1
        TestIndex = fMainForm.imlIconsSmall.ListImages.item("K" & _
                icon_small).Index
        If TestIndex = -1 Then
            TestIndex = 0
        End If
        On Error GoTo 0
        
        On Error Resume Next
        TestIndex2 = -1
        TestIndex2 = fMainForm.imlIconsLarge.ListImages.item("K" _
                & icon_large).Index
        If TestIndex2 = -1 Then
            TestIndex2 = 0
        End If
        On Error GoTo 0
        
        Set dataitem = lvQuick.ListItems.Add(i, AddK(listrecord!data_id), _
                listrecord!data_label, TestIndex2, TestIndex)
        
        i = i + 1
        listrecord.MoveNext
    Wend
    listrecord.Close

    If Not lvQuick.SelectedItem Is Nothing Then
        lvQuick.SelectedItem.Selected = False
    End If
End Sub

Private Sub BuildLinkTree()
    Dim m As Long
    Dim n As Long
    Dim selected_index As Long
    Dim nodekey As String
    
    selected_index = -1
    
    tvLink.Nodes.Clear
    tvLink.ImageList = Nothing
    If fMainForm.imlTempTree.ListImages.count <> 0 Then
        tvLink.ImageList = fMainForm.imlTempTree
    End If
    
    n = fMainForm.FindANode("Link Sources", False, True)
    n = fMainForm.tvTreeView.Nodes(AddK(n)).Index   'get index from key
    CopyNodes_to_TVLinkTree n, 0
         
    If tvLink.Nodes.count <> 0 Then
        If tvAttribCol(nodx.key).Link_NodeID <> 0 Then
            Set tvLink.SelectedItem = tvLink.Nodes(AddK(tvAttribCol(nodx.key).Link_NodeID))
        Else
            Set tvLink.SelectedItem = tvLink.Nodes(1)
        End If
        tvLink.SelectedItem.EnsureVisible
        Set nodl = tvLink.SelectedItem
        tvLink_NodeClick nodl
    End If
Done:
End Sub
Private Sub CopyNodes_to_TVLinkTree(ByVal n As Long, ByVal localParent As Long)

    If fMainForm.tvTreeView.Nodes(n).Children > 0 Then
        n = fMainForm.tvTreeView.Nodes(n).Child.Index
        While n <> fMainForm.tvTreeView.Nodes(n).LastSibling.Index
            If localParent = 0 Then
                tvLink.Nodes.Add , , fMainForm.tvTreeView.Nodes(n).key, _
                    fMainForm.tvTreeView.Nodes(n).text, _
                    fMainForm.tvTreeView.Nodes(n).Image, _
                    fMainForm.tvTreeView.Nodes(n).SelectedImage
            Else
                tvLink.Nodes.Add localParent, tvwChild, _
                    fMainForm.tvTreeView.Nodes(n).key, _
                    fMainForm.tvTreeView.Nodes(n).text, _
                    fMainForm.tvTreeView.Nodes(n).Image, _
                    fMainForm.tvTreeView.Nodes(n).SelectedImage
            End If
            CopyNodes_to_TVLinkTree n, tvLink.Nodes.count
            n = fMainForm.tvTreeView.Nodes(n).Next.Index
        Wend
        If localParent = 0 Then
            tvLink.Nodes.Add , , fMainForm.tvTreeView.Nodes(n).key, _
                fMainForm.tvTreeView.Nodes(n).text, _
                fMainForm.tvTreeView.Nodes(n).Image, _
                fMainForm.tvTreeView.Nodes(n).SelectedImage
        Else
            tvLink.Nodes.Add localParent, tvwChild, _
                fMainForm.tvTreeView.Nodes(n).key, _
                fMainForm.tvTreeView.Nodes(n).text, _
                fMainForm.tvTreeView.Nodes(n).Image, _
                fMainForm.tvTreeView.Nodes(n).SelectedImage
        End If
        CopyNodes_to_TVLinkTree n, tvLink.Nodes.count
    End If

End Sub
Private Sub tvLink_NodeClick(ByVal Node As MSComctlLib.Node)
    Set nodl = Node
    nodl.EnsureVisible
    BuildLinkList
'    Changed ("tvLink")
    If Not cbApply.Enabled Then cbApply.Enabled = True
    If Not cbOK.Enabled Then cbOK.Enabled = True
End Sub

Private Sub BuildLinkList()
    Dim listrecord As Recordset
    Dim dataitem As ListItem
    Dim i As Long
    Dim icon_large As String
    Dim icon_small As String
    Dim tempstr As String
    Dim ParentNodeId As Long
    Dim TestIndex As Long
    Dim TestIndex2 As Long
    
    ParentNodeId = RemoveK(nodl.key)
    
    lvLink.ListItems.Clear
    If fMainForm.imlIconsLarge.ListImages.count <> 0 Then
        Set lvLink.Icons = fMainForm.imlIconsLarge
    End If
    If fMainForm.imlIconsSmall.ListImages.count <> 0 Then
        Set lvLink.SmallIcons = fMainForm.imlIconsSmall
    End If
    
    Set listrecord = dbase.OpenRecordset("SELECT * FROM " & DB_GlobalsTable & " WHERE Parent_Node = " & ParentNodeId, dbOpenDynaset)
    i = 1
    While Not listrecord.EOF
    
        icon_large = listrecord!icon_large
        icon_small = listrecord!icon_small
    
        On Error Resume Next
        TestIndex = -1
        TestIndex = fMainForm.imlIconsSmall.ListImages.item("K" & _
                icon_small).Index
        If TestIndex = -1 Then
            TestIndex = 0
        End If
        On Error GoTo 0
        
        On Error Resume Next
        TestIndex2 = -1
        TestIndex2 = fMainForm.imlIconsSmall.ListImages.item("K" & _
                icon_large).Index
        If TestIndex2 = -1 Then
            TestIndex2 = 0
        End If
        On Error GoTo 0
        
        Set dataitem = lvLink.ListItems.Add(i, AddK(listrecord!data_id), _
            listrecord!data_label, TestIndex2, TestIndex)
        
        i = i + 1
        listrecord.MoveNext
    Wend
    listrecord.Close

    If Not lvLink.SelectedItem Is Nothing Then
        lvLink.SelectedItem.Selected = False
    End If
End Sub

Private Sub UpdateFromTreeView()
        
    Dim NumHiddenTables As Integer
    Dim tempstr As String
    Dim i As Long
    Dim parent As Long
    Dim record As Recordset
    Static LastGSStatus As Boolean
    Static CurrentGSStatus As Boolean
    Dim suffix As String
    Dim TestIndex As Long
    
    ResetChangedFlags
    
    ' Visible/Invisible --------------------------------------------------
    'show global and quicktype tabs
    If Not SSTab1.TabVisible(tabQuick) Then SSTab1.TabVisible(tabQuick) = True
    If Not SSTab1.TabVisible(tabLink) Then SSTab1.TabVisible(tabLink) = True
    
    'show create node/item checkboxes
    If Not cbCreateNode.Visible Then cbCreateNode.Visible = True
    If Not cbCreateItem.Visible Then cbCreateItem.Visible = True
    
    'show normal (1) and selected (2) icons
    If Not cboIcon(1).Visible Then cboIcon(1).Visible = True
    If Not cboIcon(2).Visible Then cboIcon(2).Visible = True
    
    'don't show large icons (3)
    If cboIcon(3).Visible Then cboIcon(3).Visible = False
    
    If fraValue.Visible Then fraValue.Visible = False
    If Not cbDelTable.Visible Then cbDelTable.Visible = True
    
    If Not lblGlobalType.Visible Then lblGlobalType.Visible = True
    If Not lblGlobalLabel.Visible Then lblGlobalLabel.Visible = True
    
    If Not cbDelTable.Visible Then cbDelTable.Visible = True
    
    If cbVariation.Visible Then cbVariation.Visible = False
    
    
    orig_sublink = tvAttribCol(nodx.key).sublink
    orig_read_only = tvAttribCol(nodx.key).read_only
    orig_system_node = tvAttribCol(nodx.key).system_node
    If tvAttribCol(nodx.key).Link_NodeID = 0 Then
        orig_linked = False
        If lblLinkNode.Visible Then lblLinkNode.Visible = False
        If lblLinkNodeID.Visible Then lblLinkNodeID.Visible = False
        If lblLinkNodeName.Visible Then lblLinkNodeName.Visible = False
    Else
        orig_linked = True
        If Not lblLinkNode.Visible Then lblLinkNode.Visible = True
        If Not lblLinkNodeID.Visible Then lblLinkNodeID.Visible = True
        If Not lblLinkNodeName.Visible Then lblLinkNodeName.Visible = True
    End If
    
    ' Enable/Disable --------------------------------------------------
    If orig_system_node Or orig_sublink Or orig_read_only Then
            If tbName.Enabled Then tbName.Enabled = False
            If cbReadOnly.Enabled Then cbReadOnly.Enabled = False
            If cbCreateItem.Enabled Then cbCreateItem.Enabled = False
            If cbCreateNode.Enabled Then cbCreateNode.Enabled = False
            If cbLink.Enabled Then cbLink.Enabled = False
            If cboType.Enabled Then cboType.Enabled = False
            If cboIcon(1).Enabled Then cboIcon(1).Enabled = False
            If cboIcon(2).Enabled Then cboIcon(2).Enabled = False
            If tvQuick.Enabled Then tvQuick.Enabled = False
            If lvQuick.Enabled Then lvQuick.Enabled = False
            If tvLink.Enabled Then tvLink.Enabled = False
            If lvLink.Enabled Then lvLink.Enabled = False
    Else
            If Not tbName.Enabled Then tbName.Enabled = True
            If Not cbReadOnly.Enabled Then cbReadOnly.Enabled = True
            If Not cbCreateItem.Enabled Then cbCreateItem.Enabled = True
            If Not cbCreateNode.Enabled Then cbCreateNode.Enabled = True
            If Not cbLink.Enabled Then cbLink.Enabled = True
            If Not cboType.Enabled Then cboType.Enabled = True
            If Not cboIcon(1).Enabled Then cboIcon(1).Enabled = True
            If Not cboIcon(2).Enabled Then cboIcon(2).Enabled = True
            If Not tvQuick.Enabled Then tvQuick.Enabled = True
            If Not lvQuick.Enabled Then lvQuick.Enabled = True
            If Not tvLink.Enabled Then tvLink.Enabled = True
            If Not lvLink.Enabled Then lvLink.Enabled = True
    End If
    
    'everything is disabled when the node is marked read only.  We need
    'to enable cbReadOnly so the user can change the status.
    If (orig_read_only) And (Not orig_sublink) And (Not orig_system_node) Then
        If Not cbReadOnly.Enabled Then cbReadOnly.Enabled = True
    End If
    
    'allow changing of system flag, but only if in DEBUG mode
    If DEBUGGING And Not orig_sublink Then
        If Not cbSystem.Enabled Then cbSystem.Enabled = True
        cbSystem.Caption = "System (DEBUG)"
    Else
        If cbSystem.Enabled Then cbSystem.Enabled = False
        cbSystem.Caption = "System"
    End If
    

        
    ' Fill in values --------------------------------------------------
        

    'Set the form title
    tempstr = "Node Properties for '" & nodx.text & "'"
    suffix = ""
    If orig_read_only Then
        suffix = " (READ ONLY"
    End If
    
    If orig_sublink Then
        If orig_read_only Then
            suffix = suffix & ", SUBLINK)"
        Else
            suffix = " (SUBLINK)"
        End If
    Else
        If orig_read_only Then
            suffix = suffix & ")"
        End If
    End If
    
    Me.Caption = tempstr & suffix
    
    'load the picture next to the name
    On Error Resume Next
    TestIndex = -1
    TestIndex = fMainForm.imlIconsLarge.ListImages("K" _
            & tvAttribCol(nodx.key).icon_normal).Index
    If TestIndex = -1 Then
        Set picPicture.picture = Nothing
    Else
        Set picPicture.picture = fMainForm.imlIconsLarge.ListImages("K" _
                & tvAttribCol(nodx.key).icon_normal).picture
    End If
    On Error GoTo 0
    
    'the node name
    tbName.text = nodx.text
    orig_name = nodx.text
    
    'fill in the history info
    lblCreated.Caption = tvAttribCol(nodx.key).created
    lblCreatedBy.Caption = tvAttribCol(nodx.key).created_by
    lblLastMod.Caption = tvAttribCol(nodx.key).last_modified
    lblModifiedBy.Caption = tvAttribCol(nodx.key).modified_by

    'the node id
    lblIDVal.Caption = RemoveK(nodx.key)
    
    'load the parent node id & name
    If nodx.Root.key = nodx.key Then
        lblParentNodeVal.Caption = "None"
        lblParentName.Caption = ""
    Else
        lblParentNodeVal.Caption = parent
        lblParentName.Caption = "(" & nodx.parent.text & ")"
    End If
        
    Set record = dbase.OpenRecordset("SELECT [Order],[Global_Type] FROM " & DB_NodeTable _
        & " WHERE [Node_ID] = " & RemoveK(nodx.key), dbOpenDynaset)
    lblOrder.Caption = record!order
    If IsNull(record!global_type) Then
        lblGlobalType.Caption = "None"
    Else
        lblGlobalType.Caption = record!global_type
    End If
    record.Close
    
    'load the linked node info
    If orig_linked Then
        orig_link_node_id = tvAttribCol(nodx.key).Link_NodeID
        lblLinkNodeID.Caption = orig_link_node_id
        lblLinkNodeName.Caption = fMainForm.tvTreeView.Nodes(AddK(orig_link_node_id)).text
    End If
    
    'read only attribute
    If orig_read_only = True Then
        cbReadOnly.Value = 1
    Else
        cbReadOnly.Value = 0
    End If
        
    'read create_node attribute
    If tvAttribCol(nodx.key).create_node = True Then
        cbCreateNode.Value = 1
        orig_create_node = 1
    Else
        cbCreateNode.Value = 0
        orig_create_node = 0
    End If
        
    'read create_item attribute
    If tvAttribCol(nodx.key).create_item = True Then
        cbCreateItem.Value = 1
        orig_create_item = 1
    Else
        cbCreateItem.Value = 0
        orig_create_item = 0
    End If
        
    'read system_node attribute
    If orig_system_node = True Then
        cbSystem.Value = 1
    Else
        cbSystem.Value = 0
    End If
    
    'read sublink attribute
    If tvAttribCol(nodx.key).sublink = True Then
        cbSublink.Value = 1
    Else
        cbSublink.Value = 0
    End If
        
    'read link_nodeid attribute
    If orig_linked Then
        cbLink.Value = 1
    Else
        cbLink.Value = 0
    End If
        
    'set the text for the label and the frame
    lblType.Caption = "Data Table:"
    fraType.Caption = "Location of List Items"
        
    
    LastGSStatus = CurrentGSStatus
    orig_table_name = tvAttribCol(nodx.key).table_name
    If Left$(orig_table_name, 3) = "GS_" Then
        CurrentGSStatus = True
    Else
        CurrentGSStatus = False
    End If
    
    If CurrentGSStatus Or LastGSStatus <> CurrentGSStatus Then
        TableComboNeedsRefresh = True
    End If
    
    'build the data table combo box only if it has changed since the last
    'time it was built (TableComboNeedsRefresh will be TRUE if it has changed)
    If TableComboNeedsRefresh Then
        cboType.ComboItems.Clear
        
        NumHiddenTables = 0
        Set cboType.ImageList = Nothing
        If fMainForm.imlMenu.ListImages.count <> 0 Then
            Set cboType.ImageList = fMainForm.imlMenu
        End If
            
        If Left$(orig_table_name, 3) = "GS_" Then
            If cboType.Enabled Then cboType.Enabled = False
            On Error Resume Next
            TestIndex = -1
            TestIndex = fMainForm.imlMenu.ListImages.item("AccessTable").Index
            If TestIndex = -1 Then
                TestIndex = 0
            End If
            On Error GoTo 0
            cboType.ComboItems.Add , UCase(orig_table_name), _
                orig_table_name, TestIndex, TestIndex, 0
        Else
            If Not cboType.Enabled Then cboType.Enabled = True
            For i = 0 To dbase.TableDefs.count - 1
                If UCase(mID$(dbase.TableDefs(i).Name, 1, 4)) = "MSYS" _
                    Or UCase(mID$(dbase.TableDefs(i).Name, 1, 3)) = "TM_" _
                    Or UCase(mID$(dbase.TableDefs(i).Name, 1, 3)) = "GS_" _
                    Or UCase(mID$(dbase.TableDefs(i).Name, 1, 1)) = "~" _
                    Or dbase.TableDefs(i).Name = DB_NodeTable Then
                    NumHiddenTables = NumHiddenTables + 1
                Else
                    On Error Resume Next
                    TestIndex = -1
                    TestIndex = fMainForm.imlMenu.ListImages.item("AccessTable").Index
                    If TestIndex = -1 Then
                        TestIndex = 0
                    End If
                    On Error GoTo 0
                    cboType.ComboItems.Add , UCase(dbase.TableDefs(i).Name), _
                        dbase.TableDefs(i).Name, TestIndex, TestIndex, 0
                End If
            Next i
        End If
        TableComboNeedsRefresh = False
    End If
    
    'set the value of the table name
    Set cboType.SelectedItem = cboType.ComboItems(UCase(orig_table_name))
    
    'enable/disable the delete button depending on whether
    'this table can be deleted or not.
    If cboType.SelectedItem.text = DB_DefaultDataTable Or _
            cboType.SelectedItem.text = DB_InboxTable Or _
            CurrentGSStatus Then
        If cbDelTable.Enabled Then cbDelTable.Enabled = False
    Else
        If Not cbDelTable.Enabled Then cbDelTable.Enabled = True
    End If
    
    'set the icon labels
    lblIcon1.Caption = "Normal Icon:"
    lblIcon2.Caption = "Selected Icon:"
        
    'assign the initial normal icon
    orig_icon_normal = tvAttribCol(nodx.key).icon_normal
    If orig_icon_normal = "" Then
        Set cboIcon(1).SelectedItem = cboIcon(1).ComboItems(NONE_LABEL)
    Else
        On Error Resume Next
        TestIndex = -1
        TestIndex = cboIcon(1).ComboItems("K" & orig_icon_normal).Index
        If TestIndex = -1 Then
            Set cboIcon(1).SelectedItem = cboIcon(1).ComboItems(NONE_LABEL)
        Else
            Set cboIcon(1).SelectedItem = cboIcon(1).ComboItems("K" & orig_icon_normal)
        End If
        On Error GoTo 0
    End If
    
    'assign the initial selected icon
    orig_icon_selected = tvAttribCol(nodx.key).icon_selected
    If orig_icon_selected = "" Then
        Set cboIcon(2).SelectedItem = cboIcon(2).ComboItems(NONE_LABEL)
    Else
        On Error Resume Next
        TestIndex = -1
        TestIndex = cboIcon(2).ComboItems("K" & orig_icon_selected).Index
        If TestIndex = -1 Then
            Set cboIcon(2).SelectedItem = cboIcon(2).ComboItems(NONE_LABEL)
        Else
            Set cboIcon(2).SelectedItem = cboIcon(2).ComboItems("K" & orig_icon_selected)
        End If
        On Error GoTo 0
    End If

    orig_quicktype_id = tvAttribCol(nodx.key).quicktypeid
    
    If SSTab1.Tab = tabQuick Then
        If tvQuick.Visible Then tvQuick.Visible = False
        If lvQuick.Visible Then lvQuick.Visible = False
    End If
    Set nodq = Nothing
    BuildQuickTree
    If Not tvQuick.Visible Then tvQuick.Visible = True
    If Not lvQuick.Visible Then lvQuick.Visible = True

    If SSTab1.Tab = tabLink Then
        If tvLink.Visible Then tvLink.Visible = False
        If lvLink.Visible Then lvLink.Visible = False
    End If
    Set nodl = Nothing
    BuildLinkTree
    If Not tvLink.Visible Then tvLink.Visible = True
    If Not lvLink.Visible Then lvLink.Visible = True

End Sub

Private Sub UpdateFromListView()
    Dim lvItemPrefix As String
    Dim record As Recordset
    Dim tempstr As String
    Dim templong As Long
    Dim temp_value As String
    Dim tempicon As String
    Dim Linked As Boolean
    Dim suffix As String
    Dim sublink As Boolean
    Dim system_node As Boolean
    Dim TestIndex As Long
    
    TableComboNeedsRefresh = True
    ResetChangedFlags
    
    'if we were previously on tab tabQuick or tab tabLink, change to tab 0,
    'because tabQuick and tabLink are about to become invisible
    If SSTab1.Tab = tabQuick Or SSTab1.Tab = tabLink Then 'QuickAdd or Link tabs
        SSTab1.Tab = tabGeneral
        LastTab = tabGeneral
    End If
    
    If CurrentItem Is Nothing Then
        If cbApply.Enabled Then cbApply.Enabled = False
        If cbOK.Enabled Then cbOK.Enabled = False
        Exit Sub 'They clicked on white space in the listview
    Else
        lvItemPrefix = Left$(CurrentItem.key, 1)
    End If
    
    
    ' Visible/Invisible --------------------------------------------------
    'hide global and quicktype tabs
    If SSTab1.TabVisible(tabQuick) Then SSTab1.TabVisible(tabQuick) = False
    If SSTab1.TabVisible(tabLink) Then SSTab1.TabVisible(tabLink) = False
    
    'hide create node/item checkboxes
    If cbCreateNode.Visible Then cbCreateNode.Visible = False
    If cbCreateItem.Visible Then cbCreateItem.Visible = False
    
    'show large (3) and small (2) icons
    If Not cboIcon(3).Visible Then cboIcon(3).Visible = True
    If Not cboIcon(2).Visible Then cboIcon(2).Visible = True
    
    'don't show normal icons (1)
    If cboIcon(1).Visible Then cboIcon(1).Visible = False
    
    If Not fraValue.Visible Then fraValue.Visible = True
    
    If cbDelTable.Visible Then cbDelTable.Visible = False
    
    
    If lblGlobalType.Visible Then lblGlobalType.Visible = False
    If lblGlobalLabel.Visible Then lblGlobalLabel.Visible = False
    
    If cbDelTable.Visible Then cbDelTable.Visible = False
    
    If Not cbVariation.Visible Then cbVariation.Visible = True
        
    sublink = tvAttribCol(nodx.key).sublink
    system_node = tvAttribCol(nodx.key).system_node
    orig_read_only = lvAttribCol(CurrentItem.key).read_only
    
    If lvItemPrefix = "L" Then
        Linked = True
    Else
        Linked = False
    End If
    
    ' Enable/Disable --------------------------------------------------
    
    If system_node Or sublink Or Linked Or orig_read_only Then
            If tbName.Enabled Then tbName.Enabled = False
            If cbReadOnly.Enabled Then cbReadOnly.Enabled = False
            If cbLink.Enabled Then cbLink.Enabled = False
            If cboType.Enabled Then cboType.Enabled = False
            If cboIcon(3).Enabled Then cboIcon(3).Enabled = False
            If cboIcon(2).Enabled Then cboIcon(2).Enabled = False
            If tbValue.Enabled Then tbValue.Enabled = False
            If cbEdit.Enabled Then cbEdit.Enabled = False
    Else
            If Not tbName.Enabled Then tbName.Enabled = True
            If Not cbReadOnly.Enabled Then cbReadOnly.Enabled = True
            If Not cbLink.Enabled Then cbLink.Enabled = True
            If Not cboType.Enabled Then cboType.Enabled = True
            If Not cboIcon(3).Enabled Then cboIcon(3).Enabled = True
            If Not cboIcon(2).Enabled Then cboIcon(2).Enabled = True
            If Not tbValue.Enabled Then tbValue.Enabled = True
            If Not cbEdit.Enabled Then cbEdit.Enabled = True
    End If
    
    If cbLink.Enabled Then cbLink.Enabled = False
    If cbSystem.Enabled Then cbSystem.Enabled = False
    If cbVariation.Enabled Then cbVariation.Enabled = False
    
    'everything is disabled when the item is marked read only.  We need
    'to enable cbReadOnly so the user can change the status.
    If (orig_read_only) And (Not sublink) And (Not system_node) And (Not Linked) Then
        If Not cbReadOnly.Enabled Then cbReadOnly.Enabled = True
    End If
    
    ' Set the values --------------------------------------------------
    
    'Set the form title
    tempstr = "Item Properties for '" & CurrentItem.text & "'"
    suffix = ""
    If orig_read_only Then
        suffix = " (READ ONLY"
    End If
    
    If Linked Then
        If orig_read_only Then
            suffix = suffix & ", LINKED)"
        Else
            suffix = " (LINKED)"
        End If
    Else
        If orig_read_only Then
            suffix = suffix & ")"
        End If
    End If
    
    Me.Caption = tempstr & suffix
    
    'the node name
    orig_name = CurrentItem.text
    tbName.text = orig_name
    
    'set the picture
    If CurrentItem.icon <> 0 Then
        On Error Resume Next
        TestIndex = -1
        TestIndex = fMainForm.imlTempLarge.ListImages(CurrentItem.icon).Index
        If TestIndex = -1 Then
            Set picPicture.picture = Nothing
        Else
            Set picPicture.picture = _
                fMainForm.imlTempLarge.ListImages(CurrentItem.icon).picture
        End If
        On Error GoTo 0
    Else
        picPicture.Cls
    End If
        
    'fill in the history info
    lblCreated.Caption = lvAttribCol(CurrentItem.key).created
    lblCreatedBy.Caption = lvAttribCol(CurrentItem.key).created_by
    lblLastMod.Caption = lvAttribCol(CurrentItem.key).last_modified
    lblModifiedBy.Caption = lvAttribCol(CurrentItem.key).modified_by

    'the node id
    lblIDVal.Caption = RemoveK(CurrentItem.key)
    
    'load the parent node id & name
    lblParentNodeVal.Caption = lvAttribCol(CurrentItem.key).parent_node
    lblParentName.Caption = "(" & nodx.text & ")"
        
    Set record = dbase.OpenRecordset("SELECT [Order] FROM " & DB_NodeTable _
        & " WHERE [Node_ID] = " & RemoveK(nodx.key), dbOpenDynaset)
    lblOrder.Caption = record!order
    record.Close
        
    'read only attribute
    If orig_read_only Then
        cbReadOnly.Value = 1
    Else
        cbReadOnly.Value = 0
    End If
        
    'read system_node attribute
    If system_node = True Then
        cbSystem.Value = 1
    Else
        cbSystem.Value = 0
    End If
    
    'fill in link status of item
    If Linked Then
        cbLink.Value = 1
    Else
        cbLink.Value = 0
    End If
    
    'read sublink attribute
    If sublink = True Then
        cbSublink.Value = 1
    Else
        cbSublink.Value = 0
    End If
    
    'read variation flag
    If lvAttribCol(CurrentItem.key).variation = True Then
        cbVariation = 1
    Else
        cbVariation = 0
    End If
        
    'fill the data type combo box
    lblType.Caption = "Data Type:"
    fraType.Caption = "Type"
    
    Set cboType.ImageList = Nothing
    If fMainForm.imlIconsSmall.ListImages.count <> 0 Then
        Set cboType.ImageList = fMainForm.imlIconsSmall
    End If
    cboType.ComboItems.Clear
    Set record = dbase.OpenRecordset("SELECT Data_Label,Icon_Small FROM " _
        & DB_DataTypesTable & " ORDER BY Data_Label", dbOpenDynaset)
    While Not record.EOF
        tempstr = record!data_label
        tempicon = record!icon_small
        On Error Resume Next
        TestIndex = -1
        TestIndex = fMainForm.imlIconsSmall.ListImages.item("K" & tempicon).Index
        If TestIndex = -1 Then
            TestIndex = 0
        End If
        On Error GoTo 0
        cboType.ComboItems.Add , UCase(tempstr), tempstr, TestIndex, TestIndex, 0
        record.MoveNext
    Wend
    record.Close
    orig_data_type = lvAttribCol(CurrentItem.key).data_type
    cboType.SelectedItem = cboType.ComboItems(UCase(orig_data_type))
    
       
    If tvAttribCol(nodx.key).quicktypeid = 0 Then
        'QuickTypeID = 0 means that this is a quick type item
        tempstr = DB_QuickAddItemsTable
    ElseIf Linked Then
        'if it is linked take table of the node that the parent
        'node is linked to
        tempstr = tvAttribCol(AddK(tvAttribCol(nodx.key).Link_NodeID)).table_name
    Else
        'otherwise, take the table of the parent node
        tempstr = tvAttribCol(nodx.key).table_name
    End If
        
    'read the value of the item
    Set record = dbase.OpenRecordset("SELECT [Data_Value],[Data_ID] FROM " & _
        tempstr & " WHERE Data_ID = " & RemoveK(CurrentItem.key), dbOpenDynaset)
    tbValue.text = record!data_value
    record.Close
        
    'set the icon labels
    lblIcon1.Caption = "Large Icon:"
    lblIcon2.Caption = "Small Icon:"
        
    'assign the initial large icon
    orig_icon_large = lvAttribCol(CurrentItem.key).icon_large
    If orig_icon_large = "" Then
        Set cboIcon(3).SelectedItem = cboIcon(3).ComboItems(NONE_LABEL)
    Else
        On Error Resume Next
        TestIndex = -1
        TestIndex = cboIcon(3).ComboItems("K" & orig_icon_large).Index
        If TestIndex = -1 Then
            Set cboIcon(3).SelectedItem = cboIcon(3).ComboItems(NONE_LABEL)
        Else
            Set cboIcon(3).SelectedItem = cboIcon(3).ComboItems("K" & orig_icon_large)
        End If
        On Error GoTo 0
    End If
    
    'assign the initial small icon
    orig_icon_small = lvAttribCol(CurrentItem.key).icon_small
    If orig_icon_small = "" Then
        Set cboIcon(2).SelectedItem = cboIcon(2).ComboItems(NONE_LABEL)
    Else
        On Error Resume Next
        TestIndex = -1
        TestIndex = cboIcon(2).ComboItems("K" & lvAttribCol(CurrentItem.key).icon_small).Index
        If TestIndex = -1 Then
            Set cboIcon(2).SelectedItem = cboIcon(2).ComboItems(NONE_LABEL)
        Else
            Set cboIcon(2).SelectedItem = cboIcon(2).ComboItems("K" & lvAttribCol(CurrentItem.key).icon_small)
        End If
        On Error GoTo 0
    End If
        
    If tvAttribCol(nodx.key).Link_NodeID <> 0 Then
        'if it is not a linked item, then read its data value
        'from the parent node's data table
        If lvItemPrefix = "K" Then
            Set record = dbase.OpenRecordset("SELECT Data_Value FROM " _
                & tvAttribCol(nodx.key).table_name & " WHERE Data_ID = " _
                & RemoveK(CurrentItem.key), dbOpenDynaset)
        Else
            'if it is a linked item, read the data id from the
            'source data table
            tempstr = AddK(tvAttribCol(nodx.key).Link_NodeID)
            Set record = dbase.OpenRecordset("SELECT data_ID FROM " _
                & tvAttribCol(tempstr).table_name & " WHERE data_label = '" _
                & CurrentItem.text & "'", dbOpenDynaset)
                templong = record!data_id
                record.Close
            'get the data value from the source table
            Set record = dbase.OpenRecordset("SELECT data_value FROM " _
                & tvAttribCol(tempstr).table_name & " WHERE Data_ID = " _
                & templong, dbOpenDynaset)
        End If
    Else
        'if the node is not linked, get the data value from the
        'parent node's data table
        Set record = dbase.OpenRecordset("SELECT data_value FROM " _
            & tvAttribCol(nodx.key).table_name & " WHERE Data_ID = " _
            & RemoveK(CurrentItem.key), dbOpenDynaset)
    End If
        
    'limit the display of the data value to 255 bytes.
    If record("data_value").FieldSize > 255 Then
        temp_value = "*** This value is too large to display!  " _
            & "Please use the Editor to view/modify. ***"
    Else
        temp_value = record!data_value
    End If
    record.Close
        
    tempstr = lvAttribCol(CurrentItem.key).data_type
    If UCase(mID$(tempstr, 1, 6)) = "BINARY" Then
        tempstr = "Binary"
    End If
    Select Case tempstr
        Case "Binary"
            lblAttachmentPath.Visible = True
            'If Not cbEdit.Enabled Then cbEdit.Enabled = True
            If Not cbEdit.Visible Then cbEdit.Visible = True
            If Not cbClear.Visible Then cbClear.Visible = True
            If Not cbSaveAs.Visible Then cbSaveAs.Visible = True
            If tbValue.Visible Then tbValue.Visible = False
            If temp_value = "" Then
                lblAttachmentPath.Caption = "No attached data."
                If cbClear.Enabled Then cbClear.Enabled = False
                If cbSaveAs.Enabled Then cbSaveAs.Enabled = False
            Else
                lblAttachmentPath.Caption = "Attached Binary Data:  " & _
                    mID$(temp_value, 10, Len(temp_value) - 9)
                If Not cbClear.Enabled Then cbClear.Enabled = True
                If Not cbSaveAs.Enabled Then cbSaveAs.Enabled = True
            End If
            cbEdit.Caption = "&Attach New"
        Case "String", "HTML", "URL"
            cbEdit.Caption = "&Edit"
            If lblAttachmentPath.Visible Then lblAttachmentPath.Visible = False
            If Not cbEdit.Visible Then cbEdit.Visible = True
            If cbClear.Visible Then cbClear.Visible = False
            If cbSaveAs.Visible Then cbSaveAs.Visible = False
            If Not tbValue.Visible Then tbValue.Visible = True
            tbValue.text = temp_value
        Case Else
            If lblAttachmentPath.Visible Then lblAttachmentPath.Visible = False
            If cbEdit.Visible Then cbEdit.Visible = False
            If cbClear.Visible Then cbClear.Visible = False
            If cbSaveAs.Visible Then cbSaveAs.Visible = False
            If Not tbValue.Visible Then tbValue.Visible = True
    End Select
End Sub

Private Sub ApplyTreeView()
    Dim record As Recordset
    Dim TempNodeKey As String
    Dim TempNodeParentKey As String
    Dim TempPrevSiblingKey As String
    Dim TempNextSiblingKey As String
    Dim Match As Boolean
    Dim i As Long
    Dim old_table As String
    Dim response As Integer
    Dim UpdateIcons As Boolean
    Dim Linked As Boolean
    Dim CurrentNode As Node
    Dim read_only As Boolean
    Dim system_node As Boolean
    Dim type_enabled As Boolean
    Dim UpdateListViewIcons As Boolean
    Dim UpdateListView As Boolean
    
    'assume dont need to update icons
    UpdateIcons = False
    UpdateListViewIcons = False
    UpdateListView = False
    
    'Open the record for modification
    Set record = dbase.OpenRecordset("SELECT * FROM " & DB_NodeTable _
        & " WHERE Node_ID = " & RemoveK(nodx.key), dbOpenDynaset)
    record.Edit
    
    'Update the Name, if it changed
    If orig_name <> tbName.text Then
        record!node_desc = tbName.text
        nodx.text = tbName.text
    End If
    
    'Update the modified time, always
    tvAttribCol(nodx.key).last_modified = Now
    record!last_modified = tvAttribCol(nodx.key).last_modified
        
    'update the modified by, always
    record!modified_by = CurrentUser
    tvAttribCol(nodx.key).modified_by = CurrentUser
    
    'update the read_only flag, if it changed
    If cbReadOnly.Value = 1 Then
        read_only = True
    Else
        read_only = False
    End If
    If orig_read_only <> read_only Then
        UpdateIcons = True
        record!read_only = read_only
        tvAttribCol(nodx.key).read_only = read_only
    End If
        
    'update Create_Node, if it changed
    If orig_create_node <> cbCreateNode.Value Then
        If cbCreateNode.Value = 1 Then
            record!create_node = True
            tvAttribCol(nodx.key).create_node = True
        Else
            record!create_node = False
            tvAttribCol(nodx.key).create_node = False
        End If
    End If
        
    'Update Create_Item, if it changed
    If orig_create_item <> cbCreateItem.Value Then
        If cbCreateItem.Value = 1 Then
            record!create_item = True
            tvAttribCol(nodx.key).create_item = True
        Else
            record!create_item = False
            tvAttribCol(nodx.key).create_item = False
        End If
    End If
    
    'Update System Node, if it changed
    If cbSystem.Value = 1 Then
        system_node = True
    Else
        system_node = False
    End If
    If orig_system_node <> system_node Then
        UpdateIcons = True
        UpdateListViewIcons = True
        record!system_node = system_node
        tvAttribCol(nodx.key).system_node = system_node
    End If
    
    'Update the Icons, if they changed
    If orig_icon_normal <> cboIcon(1).SelectedItem.text Then
        UpdateIcons = True
        If cboIcon(1).SelectedItem.text = NONE_LABEL Then
            record!icon_normal = ""
            tvAttribCol(nodx.key).icon_normal = ""
        Else
            record!icon_normal = cboIcon(1).SelectedItem.text
            tvAttribCol(nodx.key).icon_normal = cboIcon(1).SelectedItem.text
        End If
    
        If cboIcon(1).SelectedItem.text = NONE_LABEL Then
            nodx.Image = 0
        End If
    
        'fMainForm.SetTVIcons
        'lvNeedsRefresh = True
        'fMainForm.tvTreeView.SelectedItem = nodx
        'fMainForm.tvTreeView_NodeClick nodx
    End If
    
    If orig_icon_selected <> cboIcon(2).SelectedItem.text Then
        UpdateIcons = True
        If cboIcon(2).SelectedItem.text = NONE_LABEL Then
            record!icon_selected = ""
            tvAttribCol(nodx.key).icon_selected = ""
        Else
            record!icon_selected = cboIcon(2).SelectedItem.text
            tvAttribCol(nodx.key).icon_selected = cboIcon(2).SelectedItem.text
        End If
    
        If cboIcon(2).SelectedItem.text = NONE_LABEL Then
            nodx.SelectedImage = 0
        End If
    
        'fMainForm.SetTVIcons
        'lvNeedsRefresh = True
        'fMainForm.tvTreeView.SelectedItem = nodx
        'fMainForm.tvTreeView_NodeClick nodx
    End If
    
    'Update the QuickTypeID, if it changed
    If Not nodq Is Nothing Then
        If orig_quicktype_id <> RemoveK(nodq.key) Then
            tvAttribCol(nodx.key).quicktypeid = RemoveK(nodq.key)
            record!quicktypeid = RemoveK(nodq.key)
            fMainForm.UpdateAddPopupMenus
        End If
    End If
    
    
    'Update the Table Name, if it changed
    
    'save the enabled status
    If cboType.Enabled Then
        type_enabled = True
    Else
        type_enabled = False
    End If
    
    If Not type_enabled Then cboType.Enabled = True
    
    'select the text
    cboType.SetFocus
    SendKeys "{END}", True
    SendKeys "+{HOME}", True
    
    'reset back to the original enabled status
    If cboType.Enabled <> type_enabled Then
        cboType.Enabled = type_enabled
    End If

    If orig_table_name <> cboType.SelText Then
        Match = False
        For i = 0 To dbase.TableDefs.count - 1
            If cboType.SelText = dbase.TableDefs(i).Name Then
                Match = True
            End If
        Next i
    
        If Match Then
            If cboType.Enabled = True Then
                record!table_name = cboType.SelText
                old_table = tvAttribCol(nodx.key).table_name
                tvAttribCol(nodx.key).table_name = cboType.SelText
                If old_table <> tvAttribCol(nodx.key).table_name Then
                    CopyDataTableItems dbase, old_table, _
                            tvAttribCol(nodx.key).table_name, _
                            RemoveK(nodx.key)
                End If
            End If
        Else
            response = MsgBox("The table '" & cboType.SelText & _
                "' does not exist in the database." & Chr(13) & _
                "Do you wish to create this table?", _
                vbYesNoCancel + vbCritical, "Table Does Not Exist!")
            Select Case response
                Case vbYes
                    record!table_name = cboType.SelText
                    old_table = tvAttribCol(nodx.key).table_name
                    tvAttribCol(nodx.key).table_name = cboType.SelText
                    CopyTable dbase, "TM_DataTableTemplate", cboType.SelText
                    CopyDataTableItems dbase, old_table, _
                            tvAttribCol(nodx.key).table_name, _
                            RemoveK(nodx.key)
                    TableComboNeedsRefresh = True
                Case vbNo
                    record!table_name = cboType.SelText
                    tvAttribCol(nodx.key).table_name = cboType.SelText
                Case vbCancel
                    'record.Close
                    GoTo Done
            End Select
        End If
    End If
            
    
    If cbLink.Value = 1 Then
        Linked = True
    Else
        Linked = False
    End If
    If orig_linked <> Linked Then       'if linked status changed
        UpdateListView = True
        UpdateIcons = True
        If Linked Then
            If Not nodl Is Nothing Then
                tvAttribCol(nodx.key).Link_NodeID = RemoveK(nodl.key)
                record!Link_NodeID = RemoveK(nodl.key)
                record.Update
                record.Close
                'add sublinks
                If fMainForm.tvTreeView.Nodes(nodl.key).Children > 0 Then
                    Set CurrentNode = fMainForm.tvTreeView.Nodes(nodl.key).Child
                    For i = 1 To fMainForm.tvTreeView.Nodes(nodl.key).Children
                        fMainForm.FillTree CurrentNode.Index, _
                            False, False, False, False
                        Set CurrentNode = CurrentNode.Next
                    Next i
                End If
                GoTo DoneNoClose
            End If
        Else
            tvAttribCol(nodx.key).Link_NodeID = 0
            record!Link_NodeID = Null
            'delete sublinks
            For i = 1 To nodx.Children
                fMainForm.DeleteNode nodx.Child.key, False
            Next i
        End If
    Else            'if link status did not change
        If Linked And Not nodl Is Nothing Then
            If orig_link_node_id <> RemoveK(nodl.key) Then
                UpdateListView = True
                'delete sublinks
                For i = 1 To nodx.Children
                    fMainForm.DeleteNode nodx.Child.key, False
                Next i
                tvAttribCol(nodx.key).Link_NodeID = RemoveK(nodl.key)
                record!Link_NodeID = RemoveK(nodl.key)
                record.Update
                record.Close
                'add sublinks
                If fMainForm.tvTreeView.Nodes(nodl.key).Children > 0 Then
                    Set CurrentNode = fMainForm.tvTreeView.Nodes(nodl.key).Child
                    For i = 1 To fMainForm.tvTreeView.Nodes(nodl.key).Children
                        fMainForm.FillTree CurrentNode.Index, _
                            False, False, False, False
                        Set CurrentNode = CurrentNode.Next
                    Next i
                End If
                GoTo DoneNoClose
            End If
        End If
    End If

Done:
    record.Update
    record.Close
DoneNoClose:
    
    If UpdateIcons Then
        If tvAttribCol(nodx.key).Link_NodeID <> 0 Then
            Linked = True
        Else
            Linked = False
        End If
        nodx.Image = fMainForm.CreateOverlayedImage(fMainForm.imlIconsSmall, "K" & _
            tvAttribCol(nodx.key).icon_normal, fMainForm.imlTempTree, Linked, _
            tvAttribCol(nodx.key).read_only, tvAttribCol(nodx.key).system_node)
        nodx.SelectedImage = fMainForm.CreateOverlayedImage(fMainForm.imlIconsSmall, _
            "K" & tvAttribCol(nodx.key).icon_selected, fMainForm.imlTempTree, Linked, _
            tvAttribCol(nodx.key).read_only, tvAttribCol(nodx.key).system_node)
        nodx.ExpandedImage = nodx.SelectedImage
    End If
    If UpdateListViewIcons Then
        For i = 1 To fMainForm.lvListView.ListItems.count
            If Left$(fMainForm.lvListView.ListItems(i).key, 1) = "L" Then
                Linked = True
            Else
                Linked = False
            End If
            fMainForm.lvListView.ListItems(i).icon = fMainForm.CreateOverlayedImage(fMainForm.imlIconsLarge, "K" & _
                lvAttribCol(fMainForm.lvListView.ListItems(i).key).icon_large, fMainForm.imlTempLarge, Linked, _
                lvAttribCol(fMainForm.lvListView.ListItems(i).key).read_only, tvAttribCol(nodx.key).system_node)
            fMainForm.lvListView.ListItems(i).SmallIcon = fMainForm.CreateOverlayedImage(fMainForm.imlIconsSmall, _
                "K" & lvAttribCol(fMainForm.lvListView.ListItems(i).key).icon_small, fMainForm.imlTempSmall, Linked, _
                lvAttribCol(fMainForm.lvListView.ListItems(i).key).read_only, tvAttribCol(nodx.key).system_node)
        Next i
    End If
    If UpdateListView Then
        fMainForm.LoadListView nodx.key
    End If
    InitPropertiesForm
    LastTab = SSTab1.Tab
End Sub

Private Sub ApplyListView()
    Dim tempstr As String
    Dim record As Recordset
    Dim UpdateIcons As Boolean
    Dim read_only As Boolean
    Dim Linked As Boolean
    
    UpdateIcons = False
    
    If tvAttribCol(nodx.key).quicktypeid = 0 Then
        tempstr = DB_QuickAddItemsTable
    Else
        tempstr = tvAttribCol(nodx.key).table_name
    End If
        
    'Open the recordset for updating
    Set record = dbase.OpenRecordset("SELECT * FROM " & tempstr & " WHERE Data_ID = " _
                & RemoveK(CurrentItem.key), dbOpenDynaset)
    record.Edit
    
    'Update the data type, if it has changed

    If cboType.Enabled Then
        cboType.SetFocus
        SendKeys "{END}", True
        SendKeys "+{HOME}", True
        If orig_data_type <> cboType.SelText Then
            record!data_type = cboType.SelText
            lvAttribCol(CurrentItem.key).data_type = cboType.SelText
        End If
    End If
    
    'Update the value, if it has changed
    If HasChanged("tbValue") Then
        tempstr = cboType.SelectedItem.text
        Select Case tempstr
            Case "String", "Number", "Date", "Boolean", "URL"
                If Not DataIsValid(tempstr, tbValue.text) Then
                    If SSTab1.Tab <> tabData Then
                        SSTab1.Tab = tabData
                    End If
                    MsgBox "The data entered is not valid for the type '" _
                        & tempstr & "'", vbExclamation
                    tbValue.SetFocus
                    GoTo DoneWithErrors
                End If
                tbValue.text = ValidateData(tempstr, tbValue.text)
                record!data_value = tbValue.text
            Case Else
        End Select
    End If
    
    'Update the name, if it has changed
    If orig_name <> tbName.text Then
        record!data_label = tbName.text
        CurrentItem.text = tbName.text
    End If
    
    'Update the last modified time, always
    lvAttribCol(CurrentItem.key).last_modified = Now
    record!last_modified = lvAttribCol(CurrentItem.key).last_modified
        
    'Update the last modified by, always
    record!modified_by = CurrentUser
    lvAttribCol(CurrentItem.key).modified_by = CurrentUser
    
    
    'Update the read_only status, if it has changed
    If cbReadOnly.Value = 1 Then
        read_only = True
    Else
        read_only = False
    End If
    If orig_read_only <> read_only Then
        record!read_only = read_only
        lvAttribCol(CurrentItem.key).read_only = read_only
        UpdateIcons = True
    End If
    
    'Update the icons, if they have changed
    If orig_icon_large <> cboIcon(3).SelectedItem.text Then
        If cboIcon(3).SelectedItem.text = NONE_LABEL Then
            record!icon_large = ""
            lvAttribCol(CurrentItem.key).icon_large = ""
        Else
            record!icon_large = cboIcon(3).SelectedItem.text
            lvAttribCol(CurrentItem.key).icon_large = cboIcon(3).SelectedItem.text
        End If
    
        If cboIcon(3).SelectedItem.text = NONE_LABEL Then
            CurrentItem.icon = 0
        Else
            UpdateIcons = True
        End If
    End If
    
    If orig_icon_small <> cboIcon(2).SelectedItem.text Then
        If cboIcon(2).SelectedItem.text = NONE_LABEL Then
            record!icon_small = ""
           lvAttribCol(CurrentItem.key).icon_small = ""
        Else
            record!icon_small = cboIcon(2).SelectedItem.text
            lvAttribCol(CurrentItem.key).icon_small = cboIcon(2).SelectedItem.text
        End If
        
        If cboIcon(2).SelectedItem.text = NONE_LABEL Then
            CurrentItem.SmallIcon = 0
        Else
            UpdateIcons = True
        End If
    End If
    
Done:
    record.Update
    record.Close
    If UpdateIcons Then
        If Left$(CurrentItem.key, 1) = "L" Then
            Linked = True
        Else
            Linked = False
        End If
        CurrentItem.icon = fMainForm.CreateOverlayedImage(fMainForm.imlIconsLarge, "K" & _
            lvAttribCol(CurrentItem.key).icon_large, fMainForm.imlTempLarge, Linked, _
            lvAttribCol(CurrentItem.key).read_only, tvAttribCol(nodx.key).system_node)
        CurrentItem.SmallIcon = fMainForm.CreateOverlayedImage(fMainForm.imlIconsSmall, _
            "K" & lvAttribCol(CurrentItem.key).icon_small, fMainForm.imlTempSmall, Linked, _
            lvAttribCol(CurrentItem.key).read_only, tvAttribCol(nodx.key).system_node)
    End If
    InitPropertiesForm
    LastTab = SSTab1.Tab
    Exit Sub
DoneWithErrors:
    record.Close
End Sub

Private Sub ResetChangedFlags()
    Dim i As Long
    
    For i = 1 To NUM_FLAGS
        ChangedFlags(i) = False
    Next i
End Sub

Private Sub Changed(t As String)
    Select Case t
        Case "tbValue"
            ChangedFlags(9) = True
        Case "cbEdit"
            ChangedFlags(10) = True
        Case "cbClear"
            ChangedFlags(11) = True
        Case Else
            MsgBox "error: couldn't find item to set as changed"
    End Select
    
    If Not cbApply.Enabled Then cbApply.Enabled = True
    If Not cbOK.Enabled Then cbOK.Enabled = True
End Sub

Private Function HasChanged(t As String) As Boolean
    
    HasChanged = False
    Select Case t
        Case "tbValue"
            If ChangedFlags(9) = True Then HasChanged = True
        Case "cbEdit"
            If ChangedFlags(10) = True Then HasChanged = True
        Case "cbClear"
            If ChangedFlags(11) = True Then HasChanged = True
        Case Else
            MsgBox "error: couldn't find item to check if changed"
    End Select
End Function
