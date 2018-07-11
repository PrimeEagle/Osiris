VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{2C4F587F-97F9-11D1-B346-444553540000}#19.0#0"; "CONTROLESDOMPP.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Osiris"
   ClientHeight    =   8310
   ClientLeft      =   2340
   ClientTop       =   1005
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   11415
   Begin MPPControls.BrowseForFolders BrowseDir 
      Height          =   615
      Left            =   9945
      TabIndex        =   12
      Top             =   6300
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   750
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   1323
      BandCount       =   4
      ImageList       =   "imlMenu"
      _CBWidth        =   11415
      _CBHeight       =   750
      _Version        =   "6.0.8169"
      Child1          =   "tbStand"
      MinWidth1       =   3495
      MinHeight1      =   330
      Width1          =   3495
      NewRow1         =   0   'False
      Child2          =   "tbExternal"
      MinHeight2      =   330
      Width2          =   1155
      NewRow2         =   0   'False
      Child3          =   "tbEdit"
      MinWidth3       =   4005
      MinHeight3      =   330
      Width3          =   6000
      NewRow3         =   -1  'True
      Child4          =   "tbUtilities"
      MinHeight4      =   330
      Width4          =   11415
      NewRow4         =   0   'False
      Begin MSComctlLib.Toolbar tbExternal 
         Height          =   330
         Left            =   3885
         TabIndex        =   11
         Top             =   30
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "HTML"
               Object.ToolTipText     =   "Launch HTML Editor"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Source"
               Object.ToolTipText     =   "Launch Source Viewer"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbUtilities 
         Height          =   330
         Left            =   6195
         TabIndex        =   10
         Top             =   390
         Width           =   5130
         _ExtentX        =   9049
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "DB"
               Object.ToolTipText     =   "Database Management"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Spelling"
               Object.ToolTipText     =   "Spell Checker"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Security"
               Object.ToolTipText     =   "Security"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbEdit 
         Height          =   330
         Left            =   165
         TabIndex        =   9
         Top             =   390
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cut"
               Object.ToolTipText     =   "Cut"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Copy"
               Object.ToolTipText     =   "Copy"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Paste"
               Object.ToolTipText     =   "Paste"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Add"
               Object.ToolTipText     =   "Add New"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Delete"
               Object.ToolTipText     =   "Delete"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Rename"
               Object.ToolTipText     =   "Rename"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "MoveUp"
               Object.ToolTipText     =   "Move Node Up"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "MoveDown"
               Object.ToolTipText     =   "Move Node Down"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Execute"
               Object.ToolTipText     =   "Execute Contents"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Variation"
               Object.ToolTipText     =   "Create Variation"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Edit"
               Object.ToolTipText     =   "Launch Editor"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbStand 
         Height          =   330
         Left            =   165
         TabIndex        =   8
         Top             =   30
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imlMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Open"
               Object.ToolTipText     =   "Open Database"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Print"
               Object.ToolTipText     =   "Print"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Prop"
               Object.ToolTipText     =   "Properties"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Find"
               Object.ToolTipText     =   "Find"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Replace"
               Object.ToolTipText     =   "Replace"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Refresh"
               Object.ToolTipText     =   "Refresh Database"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Help"
               Object.ToolTipText     =   "Help"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Timer tmrNodeClick 
      Enabled         =   0   'False
      Left            =   9945
      Top             =   4920
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   8040
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14579
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "10/22/98"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "10:49 PM"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   7320
      Left            =   9720
      ScaleHeight     =   3187.442
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   72
   End
   Begin MSComctlLib.ImageList imlTempLarge 
      Left            =   9855
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView tvTreeView 
      Height          =   7440
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   13123
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imlIconsSmall"
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10560
      Top             =   4230
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open Database File"
      InitDir         =   "\"
   End
   Begin VB.PictureBox picTitles 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   11880
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   390
      Width           =   11880
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ListView:"
         Height          =   270
         Index           =   1
         Left            =   3405
         TabIndex        =   6
         Tag             =   " ListView:"
         Top             =   15
         Width           =   6330
      End
      Begin VB.Label lblTitle 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Content Tree"
         Height          =   270
         Index           =   0
         Left            =   30
         TabIndex        =   5
         Tag             =   " TreeView:"
         Top             =   15
         Width           =   3330
      End
   End
   Begin MSComctlLib.ImageList imlMenu 
      Left            =   9840
      Top             =   2790
      _ExtentX        =   2646
      _ExtentY        =   1323
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483629
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlIconsSmall 
      Left            =   9870
      Top             =   1230
      _ExtentX        =   2646
      _ExtentY        =   1323
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlIconsLarge 
      Left            =   9840
      Top             =   2010
      _ExtentX        =   2646
      _ExtentY        =   1323
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvListView 
      Height          =   7440
      Left            =   4560
      TabIndex        =   3
      Top             =   720
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   13123
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      OLEDropMode     =   1
      AllowReorder    =   -1  'True
      _Version        =   393217
      Icons           =   "imlIconsLarge"
      SmallIcons      =   "imlIconsSmall"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imlTempSmall 
      Left            =   10530
      Top             =   3585
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlTempTree 
      Left            =   9870
      Top             =   4230
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
   End
   Begin VB.Image TempImage 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   9960
      Top             =   5655
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgSplitter 
      Height          =   4785
      Left            =   3240
      MousePointer    =   9  'Size W E
      Top             =   840
      Width           =   150
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open (Ctrl-O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileDBProperties 
         Caption         =   "Database P&roperties..."
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t (Ctrl-X)"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy (Ctrl-C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste (Ctrl-V)"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete (Del)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find (Ctrl-F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditReplace 
         Caption         =   "R&eplace (Ctrl-H)"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu mnuNodes 
      Caption         =   "&Nodes"
      Begin VB.Menu mnuNodesAdd 
         Caption         =   "&Add New... (Ins)"
      End
      Begin VB.Menu mnuNodesBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNodesRename 
         Caption         =   "Rena&me (R)"
      End
      Begin VB.Menu mnuNodesBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNodesMoveUp 
         Caption         =   "Move &Up"
      End
      Begin VB.Menu mnuNodesMoveDown 
         Caption         =   "Move D&own"
      End
      Begin VB.Menu mnuNodesBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNodesProperties 
         Caption         =   "P&roperties..."
      End
   End
   Begin VB.Menu mnuItems 
      Caption         =   "&Items"
      Begin VB.Menu mnuItemsAdd 
         Caption         =   "&Add New...(Ins)"
      End
      Begin VB.Menu mnuItemsBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItemsRename 
         Caption         =   "Rena&me (R)"
      End
      Begin VB.Menu mnuItemsBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItemsExecute 
         Caption         =   "E&xecute"
      End
      Begin VB.Menu mnuItemsVariation 
         Caption         =   "Create &Variation"
      End
      Begin VB.Menu mnuItemsEdit 
         Caption         =   "Launch &Editor"
      End
      Begin VB.Menu mnuItemsBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItemsSelectAll 
         Caption         =   "Select &All (Ctrl-A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuItemsInvertSelection 
         Caption         =   "&Invert Selection"
      End
      Begin VB.Menu mnuItemsBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuItemsProperties 
         Caption         =   "P&roperties..."
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "Lar&ge Icons"
         Index           =   0
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "S&mall Icons"
         Index           =   1
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "&List"
         Index           =   2
      End
      Begin VB.Menu mnuListViewMode 
         Caption         =   "&Details"
         Index           =   3
      End
      Begin VB.Menu mnuViewBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewArrangeIcons 
         Caption         =   "Arrange &Icons"
         Begin VB.Menu mnuVAIByItem 
            Caption         =   "by &Item"
         End
         Begin VB.Menu mnuVAIByDataID 
            Caption         =   "by &Data ID"
         End
         Begin VB.Menu mnuVAIByParentNode 
            Caption         =   "by &Parent Node"
         End
         Begin VB.Menu mnuVAIByType 
            Caption         =   "by &Type"
         End
      End
      Begin VB.Menu mnuViewBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsSpelling 
         Caption         =   "Check &Spelling"
      End
      Begin VB.Menu mnuToolsBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsDatabase 
         Caption         =   "&Database Tools"
      End
      Begin VB.Menu mnuToolsBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsSecurity 
         Caption         =   "&Security..."
      End
   End
   Begin VB.Menu mnuModule 
      Caption         =   "&Modules"
      Begin VB.Menu mnuModuleHTML 
         Caption         =   "&HTML Editor"
      End
      Begin VB.Menu mnuModuleSource 
         Caption         =   "&Source Viewer"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "&Search For Help On..."
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpQuote 
         Caption         =   "&Quote of the Day"
      End
      Begin VB.Menu mnuHelpBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Osiris..."
      End
   End
   Begin VB.Menu mnuDrag 
      Caption         =   "Drag"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuDragMove 
         Caption         =   "&Move Here"
      End
      Begin VB.Menu mnuDragCopy 
         Caption         =   "&Copy Here"
      End
      Begin VB.Menu mnuDragBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDragCancel 
         Caption         =   "Cancel"
      End
   End
   Begin VB.Menu mnuRCPopupTree1 
      Caption         =   "RC_PopupTree1"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuRCPopupTree1Find 
         Caption         =   "&Find..."
      End
      Begin VB.Menu mnuRCPopupTree1Replace 
         Caption         =   "&Replace..."
      End
      Begin VB.Menu mnuRCPopupTree1Bar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCPopupTree1Print 
         Caption         =   "&Print..."
      End
      Begin VB.Menu mnuRCPopupTree1Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCPopupTree1Cut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuRCPopupTree1Copy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuRCPopupTree1Paste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuRCPopupTree1Bar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCPopupTree1Delete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuRCPopupTree1Rename 
         Caption         =   "Rena&me"
      End
      Begin VB.Menu mnuRCPopupTree1Bar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCPopupTree1MoveUp 
         Caption         =   "Move &Up"
      End
      Begin VB.Menu mnuRCPopupTree1MoveDown 
         Caption         =   "Move D&own"
      End
      Begin VB.Menu mnuRCPopupTree1Bar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCPopupTree1Add 
         Caption         =   "&Add"
         Begin VB.Menu mnuRCPopupTree1AddNode 
            Caption         =   "Placeholder"
            Index           =   0
         End
      End
      Begin VB.Menu mnuRCPopupTree1Bar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCPopupTree1Properties 
         Caption         =   "P&roperties"
      End
   End
   Begin VB.Menu mnuRCPopupList1 
      Caption         =   "RC_PopupList1"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuRCPopupList1Edit 
         Caption         =   "Launch &Editor"
      End
      Begin VB.Menu mnuRCPopupList1Execute 
         Caption         =   "E&xecute"
      End
      Begin VB.Menu mnuRCPopupList1Bar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCPopupList1Find 
         Caption         =   "&Find..."
      End
      Begin VB.Menu mnuRCPopupList1Replace 
         Caption         =   "&Replace..."
      End
      Begin VB.Menu mnuRCPopupList1Bar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCPopupList1Print 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuRCPopupList1Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCPopupList1Copy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuRCPopupList1Cut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuRCPopupList1Delete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuRCPopupList1Bar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCPopupList1Rename 
         Caption         =   "Rena&me"
      End
      Begin VB.Menu mnuRCPopupList1Bar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCPopupList1Variation 
         Caption         =   "Create &Variation"
      End
      Begin VB.Menu mnuRCPopupList1Bar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCPopupList1Properties 
         Caption         =   "P&roperties"
      End
   End
   Begin VB.Menu mnuRCPopupList2 
      Caption         =   "RC_PopupList2"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuRCPopupList2View 
         Caption         =   "&View"
         Begin VB.Menu mnuRCPopupList2VLIcons 
            Caption         =   "Lar&ge Icons"
         End
         Begin VB.Menu mnuRCPopupList2VSIcons 
            Caption         =   "S&mall Icons"
         End
         Begin VB.Menu mnuRCPopupList2VList 
            Caption         =   "&List"
         End
         Begin VB.Menu mnuRCPopupList2VDetail 
            Caption         =   "&Detail"
         End
      End
      Begin VB.Menu mnuRCPopupList2Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCPopupList2Arrange 
         Caption         =   "Arrange &Icons"
         Begin VB.Menu mnuRCPopupList2AItem 
            Caption         =   "&Item"
         End
         Begin VB.Menu mnuRCPopupList2ADataID 
            Caption         =   "&Data ID"
         End
         Begin VB.Menu mnuRCPopupList2AParent 
            Caption         =   "&Parent Node"
         End
         Begin VB.Menu mnuRCPopupList2AType 
            Caption         =   "&Type"
         End
      End
      Begin VB.Menu mnuRCPopupList2Bar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCPopupList2Paste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuRCPopupList2Bar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRCPopupList2Add 
         Caption         =   "&Add"
         Begin VB.Menu mnuRCPopupList2AddItem 
            Caption         =   ""
            Index           =   1
         End
      End
   End
   Begin VB.Menu mnuInbox 
      Caption         =   "mnuInbox"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuInboxAutoSort 
         Caption         =   "&Auto Sort"
      End
      Begin VB.Menu mnuInboxBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInboxCopy 
         Caption         =   "&Copy Here"
      End
      Begin VB.Menu mnuInboxBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInboxCancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const sglSplitLimit = 1500   'the minimum width of the tree view or the list view
Const DEFAULT_DBFILE = "C:\osiris\database\osiris.mdb"
Const DEFAULT_LARGE_ICON = "Text-4"
Const DEFAULT_SMALL_ICON = "Text-4"
Const DEFAULT_NORMAL_ICON = "3DFolder"
Const DEFAULT_SELECTED_ICON = "3DFolderOpen"
Const DEFAULT_ITEM_NAME = "Unnamed List Item"

Public LoadBlankHTML As Boolean
Public LoadBlankSourceViewer As Boolean

Dim DoingACut As Boolean
Dim FoundNodeMatch As Boolean
Dim LoadingTreeView As Boolean
Dim CopyLinkedItem As Boolean
Dim UniqueIconKey As String
Dim CutANDInPasteNode As Boolean
Dim DragDropped As Boolean
Dim NodeKeyCutFrom As String
Dim RButton As Boolean
Dim lvSomethingSelected As Boolean
Dim Devil As Boolean
Dim mbMoving As Boolean
Dim nodxExpanded As Boolean
Dim ReadytoPasteItem As Boolean
Dim ReadytoPasteNode As Boolean
Dim ItemWasCut As Boolean
Dim NodeWasCut As Boolean
Dim MouseKey As Integer
Dim DragSource As String
Dim DragTarget As String
Dim DragSubSource As String
Dim DragSubTarget As String
Dim DragItem As Object
Dim CurrentData As MSComctlLib.DataObject
Dim lParam As String
Dim IgnoreItemClick As Boolean
Dim ItemBuffer As New Collection
Dim SelectedItemsBuffer As New Collection

'****************************************************************************
'*********************          TREEVIEW FUNCTIONS          *****************
'****************************************************************************

'----------------------------------------------------------------------------
'Populates the node tree based on a table in a given database.  When it
'searches for data item tables in the database, it will ignore temporary
'and system tables, the DB_NodeTable, and template tables (any tables whose
'names begin with "TM_".
'REQUIRES:  db - the database to load from.
'----------------------------------------------------------------------------
Private Sub LoadTreeView(Optional UseProgressBar As Boolean = False)
    
    Dim record As Recordset
    Dim i As Long
    Dim bad_table_count As Long 'counts number of tables that didn't exist,
                                'but were assigned to nodes.
    Dim new_table As String
    Dim count As Long
    Dim result As Boolean
    Dim Link_NodeID As Long
    Dim parent As Long
    Dim quicktypeid As Long
    
    
    'set this flag, so that all screen updates are not performed
    'for each node that is added while loading the treeview.
    LoadingTreeView = True
    lvListView.Visible = False
    
    Me.MousePointer = vbHourglass
    
    'Loop through the database and get all the valid table names
    'that can store data values for the list items.  Temporarily
    'store the table names in the listview.
    'Filter out unwanted tables, such MSYS* (Access System Tables),
    '~* (Temporary Tables created by Access), the DB_NodeTable, and
    'Template Tables (TM_*)
    For i = 0 To dbase.TableDefs.count - 1
        If Not (UCase(mID$(dbase.TableDefs(i).Name, 1, 4)) = "MSYS" _
            Or UCase(mID$(dbase.TableDefs(i).Name, 1, 3)) = "TM_" _
            Or UCase(mID$(dbase.TableDefs(i).Name, 1, 1)) = "~" _
            Or dbase.TableDefs(i).Name = DB_NodeTable) Then
                lvListView.ListItems.Add , UCase(dbase.TableDefs(i).Name), _
                dbase.TableDefs(i).Name
        End If
    Next i
    
    'If the default data table doesn't exist, display an error and exit
    'the sub.
    If lvListView.FindItem(DB_DefaultDataTable) Is Nothing Then
        MsgBox "LoadTreeView:  The table '" & DB_DefaultDataTable _
                & "' was not found!", vbCritical
        GoTo Done
    End If
    
    'If the progressbar is in use, get the number of nodes from the database
    'and initialize the progress bar.
    If UseProgressBar Then
        Set record = dbase.OpenRecordset("SELECT COUNT (*) AS [Count] FROM " _
            & DB_NodeTable, dbOpenDynaset)
        record_count = record!count
        record.Close
        InitProgressBar fProgForm, "Loading TreeView . . .", 0, record_count, _
            LargeIconFolder & PBICON_LoadTreeView, , InFormLoad
        count = 0   'reset the progress bar counter to 0
    End If
    
    'Select all the nodes from the DB_NodeTable, in the order given by
    'the [Order] field.
    Set record = dbase.OpenRecordset("SELECT * FROM " & DB_NodeTable _
            & " ORDER BY [Order]", dbOpenDynaset)
    bad_table_count = 0
    
    'loop through all the nodes in the table
    While Not record.EOF
        If UseProgressBar Then
            'increment the counter once for each node
            count = count + 1
            fProgForm.pbPBar1.Value = count
        End If
        
        'if records contain a Null in the database, we translate
        'these to a zero in the memory structure.
        If IsNull(record!parent) Then
            parent = 0
        Else
            parent = record!parent
        End If
        
        If IsNull(record!quicktypeid) Then
            quicktypeid = 0
        Else
            quicktypeid = record!quicktypeid
        End If
        
        If IsNull(record!Link_NodeID) Then
            Link_NodeID = 0
        Else
            Link_NodeID = record!Link_NodeID
        End If
        
        'if the assigned table for a node doesn't exist in the database,
        'reassign to the DB_DefaultDataTable.
        new_table = record!table_name
        If lvListView.FindItem(new_table) Is Nothing Then
            bad_table_count = bad_table_count + 1
            new_table = DB_DefaultDataTable
            record.Edit
            record!table_name = new_table
            record.Update
        End If

        'Add the node to the treeview structure
        AddNode parent, record!node_id, record!node_desc, _
                record!icon_normal, record!icon_selected, record!read_only, _
                new_table, quicktypeid, _
                record!create_item, record!create_node, _
                record!system_node, record!created, _
                record!created_by, record!last_modified, _
                record!modified_by, Link_NodeID, , record!sublink
        record.MoveNext
    Wend
    
Done:
    record.Close
    lvListView.ListItems.Clear
    lvListView.Visible = True
    
    'set the title bar of the form to display the path of
    'the currently opened database
    Me.Caption = "Osiris Explorer - (" & CurrentDatabaseFile & ")"
    
    Me.MousePointer = vbArrow
    
    'if we had to replace any table names, inform the user of how
    'many were replaced, and what they were replaced with.
    If bad_table_count > 0 Then
        MsgBox "There were " & Format$(bad_table_count) _
            & " node(s) in the database that were referencing non-existant tables." _
            & Chr(13) & "They were reset to reference the '" & DB_DefaultDataTable _
            & "' table.", vbOKOnly + vbExclamation
    End If
    
    'get rid of the progress bar
    If UseProgressBar Then
        fProgForm.Hide
    End If
    
    'reset the flag, so normal screen updates happen with adding new nodes.
    LoadingTreeView = False
End Sub
'----------------------------------------------------------------------------
'*** Recursive ***
'Populates one branch of the node tree based on a table in a given database.
'REQUIRES:  db - the database to load from.
'----------------------------------------------------------------------------
Public Sub LoadTreeBranch(start_nodeid As Long, relationship As Long, relative As Long)
    Dim record As Recordset
    Dim i As Long
    Dim table_edited As Boolean
    Dim new_table As String
    Dim new_relative As Long
    Dim parent As Long
    Dim quicktypeid As Long
    Dim Link_NodeID As Long
    
    'find the requested node in the database
    Set record = dbase.OpenRecordset("SELECT * FROM " & DB_NodeTable _
        & " WHERE Node_ID = " & start_nodeid, dbOpenDynaset)

    If Not record.EOF Then
        'if a value in the database is "Null", use a zero in memory instead.
        If IsNull(record!parent) Then
            parent = 0
        Else
            parent = record!parent
        End If
        
        If IsNull(record!quicktypeid) Then
            quicktypeid = 0
        Else
            quicktypeid = record!quicktypeid
        End If
        
        If IsNull(record!Link_NodeID) Then
            Link_NodeID = 0
        Else
            Link_NodeID = record!Link_NodeID
        End If
        
        'determine the relationship
        If relationship = tvwChild Then
            new_relative = parent
        Else
            new_relative = relative
        End If
        
        'add the node
        Call AddNode(new_relative, record!node_id, record!node_desc, _
                    record!icon_normal, record!icon_selected, record!read_only, _
                    record!table_name, quicktypeid, _
                    record!create_item, record!create_node, _
                    record!system_node, record!created, _
                    record!created_by, record!last_modified, _
                    record!modified_by, Link_NodeID, relationship, _
                    record!sublink)
        record.Close
        
        'see if it has any children
        Set record = dbase.OpenRecordset("SELECT * FROM " & DB_NodeTable _
            & " WHERE Parent = " & start_nodeid & " ORDER BY [Order]", _
            dbOpenDynaset)
        
        'recurse through all its children, and add them too
        While Not record.EOF
            LoadTreeBranch record!node_id, tvwChild, 0
            record.MoveNext
        Wend
        record.Close
    End If
End Sub
'----------------------------------------------------------------------------
'Deletes a node, all its children, and all list items assosciated with each
'of those nodes.  The deletion of the children and list items are handled
'automatically the relations set up in the database.
'REQUIRE:   NodeKey - the key of the node to be deleted.
'           SelectCloseRelative - (Default=TRUE) Once a node is deleted,
'                   a sibling or parent is selected, depending on what is
'                   available.  Set this to 'FALSE' to disable this action.
'RETURNS:   TRUE if successfull, FALSE if not.
'----------------------------------------------------------------------------
Public Function DeleteNode(ByVal nodekey As String, _
            Optional SelectCloseRelative As Boolean = True) As Boolean
    
    Dim TempNodeKey As String   'stores a temporary node key, as a string
    Dim record As Recordset
    Dim listrecord As Recordset
    Dim record_count As Long
    Dim order As Long
    Dim tempstr As String
    Dim errloop As Error

    DeleteNode = True 'default the return value to no-error
            
    'retrieve the order of the node wished to delete
    Set record = dbase.OpenRecordset("SELECT Node_ID,[Order] FROM " & DB_NodeTable & _
        " WHERE Node_ID = " & RemoveK(nodekey), dbOpenDynaset)
    order = record!order
    record.Close
        
    'count how many nodes there are in the database before the Delete happens.
    Set record = dbase.OpenRecordset("SELECT COUNT (*) AS [Count] " _
        & "FROM " & DB_NodeTable, dbOpenDynaset)
    record_count = record!count
    record.Close
        
    'delete the node, its children, and their list items from the database.
    On Error GoTo Err_Execute
    dbase.Execute "DELETE * FROM " & DB_NodeTable & " WHERE Node_ID = " & RemoveK(nodekey), dbFailOnError
    On Error GoTo 0
        
    'get # deleted from (prev # of nodes - # of nodes)
    Set record = dbase.OpenRecordset("SELECT COUNT (*) AS [Count] " _
        & "FROM " & DB_NodeTable, dbOpenDynaset)
    record_count = record_count - record!count
    record.Close
        
    'reorder the [Order] field for the following nodes
    On Error GoTo Err_Execute
    dbase.Execute "UPDATE " & DB_NodeTable & " SET [Order]=[Order] - " _
        & record_count & " WHERE [Order] > " & order, dbFailOnError
    On Error GoTo 0

    'this is to not select the parent or sibling if the flag is not set
    If SelectCloseRelative Then
        'first try choosing the next sibling
        Set nodx = tvTreeView.Nodes(nodekey).Next
        If nodx Is Nothing Then
            'if there is no next sibling, try the previous sibling
            Set nodx = tvTreeView.Nodes(nodekey).Previous
            If nodx Is Nothing Then
                'if there was no previous sibling either, then
                'choose the parent node.
                Set nodx = tvTreeView.Nodes(nodekey).parent
            End If
        End If
        lvNeedsRefresh = True
        'select the appropriate node
        tvTreeView.SelectedItem = nodx
        tvTreeView_NodeClick nodx
    End If
    ClearNode nodekey   'clear the node branch from the treeview
    GoTo Done

Err_Execute:
    DisplayDBEngineErrors

Done:
End Function
'----------------------------------------------------------------------------
' Adds a node to the node tree (not to the database).  If the parametes
' specify that the node is to be a child of node zero, then it is treated
' as a root node.
' REQUIRES: node id of the relative of the new node,
'           node id of the new node,
'           description of the new node,
'           normal icon of the new node,
'           selected icon of the new node,
'           read_only flag, data table name,
'           quicktype id, create item flag,
'           create node flag, system node flag,
'           creation date, create by user,
'           last modified date, last modified by user,
'           link node id, relationship to the relative,
'           sublink flag
'----------------------------------------------------------------------------
Private Sub AddNode(relative As Long, thenodeid As Long, _
        description As String, icon_normal As String, _
        icon_selected As String, read_only As Boolean, _
        table_name As String, quicktypeid As Long, _
        create_item As Boolean, create_node As Boolean, _
        system_node As Boolean, created As Variant, _
        created_by As String, last_modified As Variant, _
        modified_by As String, Link_NodeID As Long, _
        Optional relationship As Long = tvwChild, _
        Optional sublink As Boolean = False)

    Dim tmptvnode As New tvnode
    Dim Linked As Boolean
    
    If Link_NodeID <> 0 Then
        Linked = True
    Else
        Linked = False
    End If
    
    'Check to see if the node is a root node (parent=0) or not,
    'and add the node appropriately.
    If relative = 0 And relationship = tvwChild Then
        Set nodx = tvTreeView.Nodes.Add(, , AddK(thenodeid), _
            description, CreateOverlayedImage(imlIconsSmall, "K" & _
            icon_normal, imlTempTree, Linked, read_only, system_node), _
            CreateOverlayedImage(imlIconsSmall, "K" & _
            icon_selected, imlTempTree, Linked, read_only, system_node))
    Else
        Set nodx = tvTreeView.Nodes.Add(AddK(relative), _
            relationship, AddK(thenodeid), description, _
            CreateOverlayedImage(imlIconsSmall, "K" & _
            icon_normal, imlTempTree, Linked, read_only, system_node), _
            CreateOverlayedImage(imlIconsSmall, "K" & _
            icon_selected, imlTempTree, Linked, read_only, system_node))
    End If


    'Dim i As Long
    'Dim indexoffset As Long
    'Dim commapos As Integer
    
        
    'store the normal and selected icon indeces in the tag field of the node
    'nodx.tag = format$(CreateOverlayedImage(imlIconsSmall, "K" & _
        icon_normal, imlTempTree, AddL, read_only, system_node))
    'nodx.tag = nodx.tag & "," & format$(CreateOverlayedImage(imlIconsSmall, "K" & _
        icon_selected, imlTempTree, AddL, read_only, system_node))
    
    'copy the icon indeces out of the tag property and into the
    'appropriate icon properties
    'commapos = InStr(1, nodx.tag, ",", vbTextCompare)
    'nodx.Image = val(Left$(nodx.tag, commapos - 1))
    'nodx.SelectedImage = val(Right$(nodx.tag, _
    '        Len(nodx.tag) - commapos))
    'nodx.ExpandedImage = nodx.SelectedImage


    'fill in the tvattribcol memory structure.
    tmptvnode.icon_normal = icon_normal
    tmptvnode.icon_selected = icon_selected
    tmptvnode.read_only = read_only
    tmptvnode.table_name = table_name
    tmptvnode.quicktypeid = quicktypeid
    tmptvnode.create_item = create_item
    tmptvnode.create_node = create_node
    tmptvnode.system_node = system_node
    tmptvnode.created = created
    tmptvnode.created_by = created_by
    tmptvnode.last_modified = last_modified
    tmptvnode.modified_by = modified_by
    tmptvnode.Link_NodeID = Link_NodeID
    tmptvnode.sublink = sublink
    tvAttribCol.Add tmptvnode, nodx.key
End Sub
'----------------------------------------------------------------------------
'Refreshes the entire database, basically by unloading and reloading the all
'the configuration data and the node tree.
'----------------------------------------------------------------------------
Public Sub RefreshNodes(Optional UseProgressBar As Boolean = False)
    
    Dim i As Long
    Dim RefreshIcons As Long
    Dim ReopenDatabase As Long
    Dim OldNodeKey As String
    
    RefreshIcons = GetSetting(App.EXEName, "Options", _
            "Refresh Icons", 1)
    ReopenDatabase = GetSetting(App.EXEName, "Options", _
            "Reopen Database", 1)
   
    RefreshInProgress = True
    
    If UseProgressBar Then
        InitProgressBar fProgForm, "Clearing Node Tree Data . . .", 0, tvTreeView.Nodes.count, _
            LargeIconFolder & "Hostmanager.bmp", , False
    End If
    
    'save the currently selected node, so we can reselect it when
    'we are done with the refresh
    OldNodeKey = nodx.key
    
    'clear out the node tree and the parallel tvAttribCol collection
    For i = 1 To tvTreeView.Nodes.count
        tvAttribCol.Remove (tvTreeView.Nodes(i).key)
        If UseProgressBar Then
            fProgForm.pbPBar1.Value = i
        End If
    Next i
    tvTreeView.Nodes.Clear
    Set nodx = Nothing
    
    If UseProgressBar Then
        fProgForm.Hide
    End If
    
    'if the "Refresh Icons" option is turned on, then unbind all the image
    'lists and clear them out.
    If RefreshIcons Then
        'unbind the image lists
        Set tvTreeView.ImageList = Nothing
        Set lvListView.Icons = Nothing
        Set lvListView.SmallIcons = Nothing
        Set fPropForm.cboIcon(1).ImageList = Nothing
        Set fPropForm.cboIcon(2).ImageList = Nothing
        Set fPropForm.cboIcon(3).ImageList = Nothing
        Set fPropForm.tvQuick.ImageList = Nothing
        Set fPropForm.tvLink.ImageList = Nothing
        Set fPropForm.lvQuick.Icons = Nothing
        Set fPropForm.lvQuick.SmallIcons = Nothing
        Set fPropForm.lvLink.Icons = Nothing
        Set fPropForm.lvLink.SmallIcons = Nothing
    
        'clear out all images from the image list
        imlIconsLarge.ListImages.Clear
        imlIconsSmall.ListImages.Clear
    End If
    
    'clear out the listview and the parallel lvAttribCol collection
    For i = 1 To lvListView.ListItems.count
        lvAttribCol.Remove lvListView.ListItems.item(i).key
    Next i
    lvListView.ListItems.Clear
    
    
    'close the database and reopen it, if the "Reopen Database" option
    'is turned on.  This is only necessary to catch changes made to the
    'database very recently, that the MS Jet Engine hasn't had a chance to
    'update yet, sort of like a "DoEvents" for the database.
    If ReopenDatabase Then
        dbase.Close
        Set dbase = OpenDBase(CurrentDatabaseFile)
        If dbase Is Nothing Then
            MsgBox "RefreshNodes:  Error while attempting to re-open the database!", _
                    vbCritical
            Exit Sub
        End If
    End If
    
    If RefreshIcons Then
        LoadIcons True
        Set tvTreeView.ImageList = imlIconsSmall
        Set lvListView.Icons = imlIconsLarge
        Set lvListView.SmallIcons = imlIconsSmall
        Set fPropForm.cboIcon(1).ImageList = imlIconsSmall
        Set fPropForm.cboIcon(2).ImageList = imlIconsSmall
        Set fPropForm.cboIcon(3).ImageList = imlIconsLarge
        Set fPropForm.lvQuick.Icons = imlIconsLarge
        Set fPropForm.lvQuick.SmallIcons = imlIconsSmall
        Set fPropForm.lvLink.Icons = imlIconsLarge
        Set fPropForm.lvLink.SmallIcons = imlIconsSmall
    End If
    
    'reload the treeview and set up its icons
    LoadTreeView True
    'SetTVIcons
    
    'choose the node they were on when they hit refresh
    Set nodx = tvTreeView.Nodes(OldNodeKey)
    nodx.EnsureVisible
    Set tvTreeView.SelectedItem = nodx
    nodx.Selected = True
    lvNeedsRefresh = True
    tvTreeView_NodeClick nodx
    
    RefreshInProgress = False
End Sub
'----------------------------------------------------------------------------
'If the user types the asterisk key (*) and the current node is expanded,
'it will be collapsed.  If the current node is collapsed, it will be
'expanded.
'----------------------------------------------------------------------------
Public Sub AsteriskKey(ByVal wParam As Long)
    
    If wParam = 42 Then
        Me.MousePointer = vbHourglass
        If nodxExpanded Then
            SendMessageLong tvTreeView.hwnd, WM_SETREDRAW, 0, ByVal 0&
            FullCollapse nodx.Index
            nodx.EnsureVisible
            SendMessageLong tvTreeView.hwnd, WM_SETREDRAW, 1, ByVal 0&
        Else
            SendMessageLong tvTreeView.hwnd, WM_SETREDRAW, 0, ByVal 0&
            FullExpand nodx.Index
            nodx.EnsureVisible
            SendMessageLong tvTreeView.hwnd, WM_SETREDRAW, 1, ByVal 0&
        End If
        Me.MousePointer = vbArrow
    End If
End Sub










'****************************************************************************
'*********************          LISTVIEW FUNCTIONS          *****************
'****************************************************************************


'----------------------------------------------------------------------------
'Displays the list view for a give node id.
'REQUIRES:  the key of the node to load the list view for.
'----------------------------------------------------------------------------
Public Sub LoadListView(ByVal ParentNodeKey As String, _
        Optional SelectSelectedItemBufferItems As Boolean = False)

    Dim i As Long
    Dim tempstr As String
    Dim Link_NodeID As Long
    Dim TestIndex As Long
    
    Me.MousePointer = vbHourglass
    
    'turn off redraw while loading the listview
    LockWindowUpdate lvListView.hwnd
    
    'clear out existing list items in the lvAttribCol collection
    For i = 1 To lvListView.ListItems.count
        lvAttribCol.Remove lvListView.ListItems.item(i).key
    Next i
    
    'clear out the current listview
    lvListView.ListItems.Clear
    
    'get the correct table name
    If tvAttribCol(ParentNodeKey).quicktypeid = 0 Then
        tempstr = DB_QuickAddItemsTable
    Else
        tempstr = tvAttribCol(ParentNodeKey).table_name
    End If
    
    Link_NodeID = tvAttribCol(ParentNodeKey).Link_NodeID
    
        
    'load the listview without clearing whats in it first
    LoadLVWithoutClear tempstr, RemoveK(ParentNodeKey), False
    
    'if the node is linked, get the table name from the master copy
    'and load the listview for it on top of this one.  If there are
    'items with the same names, they will not be loaded during the second
    'pass (so the real copies win out over linked copies).
    If Link_NodeID <> 0 Then
        tempstr = tvAttribCol(AddK(Link_NodeID)).table_name
        LoadLVWithoutClear tempstr, Link_NodeID, True
    End If
    
    'set the overlay icons and add them to the listview
'    SetLargeOverlays
'    SetSmallOverlays
    
Done:
    'deselect whatever is is selected.  If we don't do this, when we
    'try to select the first one, we could end up with more than one
    'item selected (since multiselect is set to TRUE for the listview)
    If Not lvListView.SelectedItem Is Nothing Then
        lvListView.SelectedItem.Selected = False
    End If
    
    If SelectSelectedItemBufferItems Then
        For i = 1 To SelectedItemsBuffer.count
            Set CurrentItem = lvListView.FindItem(SelectedItemsBuffer("K" & i))
            If Not CurrentItem Is Nothing Then
                CurrentItem.Selected = True
            End If
        Next i
    Else
        'if there are any items in the listview, then select the first one
        'by default.
        If lvListView.ListItems.count > 0 Then
            Set CurrentItem = lvListView.ListItems(1)
            Set lvListView.SelectedItem = CurrentItem
        End If
    End If

    UpdateStatusBar 0
    
    'turn redraw back on
    LockWindowUpdate (0&)
    Me.MousePointer = vbArrow
    
    lvLoaded = True
End Sub
'----------------------------------------------------------------------------
'Loads in list items and adds them to the listview, without clearing out
'already existing list items.
'REQUIRES:  db - the database to look in
'           table_name - the table that the data items are in
'           ParentNodeID - the parent node of the list items to load
'           AddL - flag to decide whether the key of the items starts
'                  with "K"  or "L"
'----------------------------------------------------------------------------
Private Sub LoadLVWithoutClear(table_name As String, ParentNodeId As Long, _
                                    Optional AddL As Boolean = False)
    Dim i As Long
    Dim listrecord As Recordset
    Dim dataitem As ListItem
    Dim Linked As Boolean
    Dim CurrentItemKey As String
    
    'select from the database all listitems that under the ParentNodeID
    Set listrecord = dbase.OpenRecordset("SELECT * FROM " & table_name _
            & " WHERE Parent_Node = " & ParentNodeId, dbOpenDynaset)
    i = 1
    While Not listrecord.EOF
        'if it already exists in the listview, skip to the next one
        'this is what prevents local and linked copies with the same
        'name from conflicting.
        If Not lvListView.FindItem(listrecord!data_label) Is Nothing Then
            GoTo NextListItem
        End If
    
        'if it didn't exist in the listview, create and add an lvAttribCol entry
        'for it
        LoadLVAttribCol nodx.key, listrecord!icon_large, listrecord!icon_small, listrecord!read_only, _
                listrecord!data_type, listrecord!created, listrecord!created_by, _
                listrecord!last_modified, listrecord!modified_by, listrecord!variation, _
                listrecord!data_id, AddL
    
        'add the item to the list view, and set it as CurrentItem
        CurrentItemKey = AddK(listrecord!data_id, AddL)
        Set CurrentItem = lvListView.ListItems.Add(, CurrentItemKey, _
            listrecord!data_label, CreateOverlayedImage(imlIconsLarge, "K" & _
            lvAttribCol(CurrentItemKey).icon_large, imlTempLarge, AddL, _
            lvAttribCol(CurrentItemKey).read_only, tvAttribCol(nodx.key).system_node), _
            CreateOverlayedImage(imlIconsSmall, _
            "K" & lvAttribCol(CurrentItemKey).icon_small, imlTempSmall, AddL, _
            lvAttribCol(CurrentItemKey).read_only, tvAttribCol(nodx.key).system_node))
    
        'if we are in report mode, fill in the subitems as well
        If lvListView.View = lvwReport Then
            CurrentItem.SubItems(1) = listrecord!data_id
            CurrentItem.SubItems(2) = listrecord!parent_node
            CurrentItem.SubItems(3) = listrecord!icon_large
            CurrentItem.SubItems(4) = listrecord!icon_small
            CurrentItem.SubItems(5) = listrecord!data_type
        End If
        
        i = i + 1
NextListItem:
        listrecord.MoveNext
    Wend
    listrecord.Close
End Sub
'----------------------------------------------------------------------------
'Deletes a list item from the database and the listview,
'then selects the next item in line.
'REQUIRES:  DataID - the id of the item to be deleted,
'           ParentKey - the key of the parent node
'           ResetSelection - if TRUE, selects the next item after deletion.
'RETURNS:   TRUE if successfull, FALSE if not
'----------------------------------------------------------------------------
Private Function DeleteListItem(DataKey As String, ParentKey As String, _
                ResetSelection As Boolean) As Boolean
    
    Dim listrecord As Recordset
    Dim TableName As String
    'Dim tempstr As String
    
    'assume the operation will be successfull
    DeleteListItem = True
    
    'get the correct table name for this item.
    If tvAttribCol(ParentKey).quicktypeid = 0 Then
        TableName = DB_QuickAddItemsTable
    Else
        TableName = tvAttribCol(ParentKey).table_name
    End If
    
    'Open the table, and locate the correct list item in it
    Set listrecord = dbase.OpenRecordset("SELECT read_only,data_label FROM " _
        & TableName & " WHERE Data_ID = " & RemoveK(DataKey), dbOpenDynaset)
                    
    'if list item not found, display error msg, and return
    If listrecord.EOF Then
        MsgBox "DeleteListItem:  Data_ID " & Format$(RemoveK(DataKey)) & " was NOT found!"
        'if list item is read-only, display error msg, and exit function
        listrecord.Close
        DeleteListItem = False
        GoTo Done
    End If
    
    'if item is found and not read-only, then delete the record
    Me.MousePointer = vbHourglass
    listrecord.Delete   'delete from database
    listrecord.Close
    
    'if we were displaying the item at the time it was deleted,
    'the remove it from memory as well.
    If nodx.key = ParentKey Then
        lvAttribCol.Remove DataKey
        lvListView.ListItems.Remove DataKey
        lvNumListItemsSelected = lvNumListItemsSelected - 1
        
        'if this flag is true, select the next item in line
        If ResetSelection Then
            SendKeys " "    'spacebar selects the next item in line
            Set CurrentItem = lvListView.SelectedItem
        End If
    End If

Done:
    UpdateStatusBar lvNumListItemsSelected
    Me.MousePointer = vbArrow
End Function
'----------------------------------------------------------------------------
'Creates a temporary lvitem object, fills in its values, and adds it to
'the lvAttribCol collection.
'REQUIRES:  large icon, small icon,
'           read only flag, data type,
'           creation date, created by user,
'           last modified date, last modified by user,
'           data id, AddL (flag to determine whether key starts with "K" or
'           "L"
'----------------------------------------------------------------------------
Private Sub LoadLVAttribCol(ParentNodeKey As String, icon_large As String, _
        icon_small As String, read_only As Boolean, data_type As String, _
        created As Variant, created_by As String, last_modified As Variant, _
        modified_by As String, variation As Boolean, data_id As Long, _
        Optional AddL As Boolean = False)
                    
    Dim tmplvitem As New lvitem
        
    tmplvitem.parent_node = RemoveK(ParentNodeKey)
    tmplvitem.icon_large = icon_large
    tmplvitem.icon_small = icon_small
    tmplvitem.read_only = read_only
    tmplvitem.data_type = data_type
    tmplvitem.created = created
    tmplvitem.created_by = created_by
    tmplvitem.last_modified = last_modified
    tmplvitem.modified_by = modified_by
    tmplvitem.variation = variation
        
    lvAttribCol.Add tmplvitem, AddK(data_id, AddL)
End Sub
Public Sub lvListView_AfterLabelEdit(Cancel As Integer, NewString As String)
    
    Dim listrecord As Recordset
    Dim tempstr As String
    
    'make sure they entered something, if not cancel the rename
    If Len(NewString) < 1 Then
        Cancel = True
        GoTo Done
    End If
    
    'check for invalid characters in the name (' or ").  If the user entered one,
    'start over with the rename
    If InStr(1, NewString, "'", vbTextCompare) <> 0 Or _
            InStr(1, NewString, Chr(34), vbTextCompare) <> 0 Then
        Cancel = True
        MsgBox "Quotes and Apostrophes are not allowed; please reenter.", _
                vbExclamation
        lvListView.StartLabelEdit
        GoTo Done
    End If
    
    'by now, we have a potentially valid name, but we need to check if it is already
    'in use.  If it is, cancel the rename.
    If Not lvListView.FindItem(NewString) Is Nothing Then
        MsgBox "An item by the name of '" & NewString & "' already exists." _
            & Chr(13) & "Please give the current item a different name.", vbExclamation
        Cancel = True
        lvListView.StartLabelEdit
        GoTo Done
    End If
    
    'pick the correct table for the item and change the name.
    If tvAttribCol(nodx.key).quicktypeid = 0 Then
        tempstr = DB_QuickAddItemsTable
    Else
        tempstr = tvAttribCol(nodx.key).table_name
    End If
    
    Set listrecord = dbase.OpenRecordset("SELECT * FROM " & tempstr _
        & " WHERE Data_ID = " _
        & RemoveK(CurrentItem.key), dbOpenDynaset)
    If listrecord.EOF Then
        listrecord.Close
        MsgBox "lvListView_AfterLabelEdit:  Item not in database!", vbCritical
        GoTo Done
    End If

    listrecord.Edit
    listrecord!data_label = NewString
    listrecord!last_modified = Now
    listrecord!modified_by = CurrentUser
    listrecord.Update
    listrecord.Close
    
    'update the listview
    CurrentItem.text = NewString

Done:
    If Cancel Then
        lvSaveEraseRect = False
        lvEraseBkGnd1 = True
        lvEraseBkGnd2 = True
        lvNextNoErase = False
    End If
End Sub
Private Sub lvListView_BeforeLabelEdit(Cancel As Integer)
    'these variables are used by the subclassing to handle redraws
    If Cancel Then
        lvSaveEraseRect = False
        lvEraseBkGnd1 = True
        lvEraseBkGnd2 = True
        lvNextNoErase = False
    End If
End Sub

Private Sub lvListView_Click()
    
    'if we're dragging over the listview with the right button, we don't want
    'to process an item click so we ignore it with this flag.
    If IgnoreItemClick Then
        IgnoreItemClick = False
        RButton = False
    Else
        'check to see how many list items are currently selected
        lvNumListItemsSelected = SendMessage(lvListView.hwnd, _
            LVM_GETSELECTEDCOUNT, 0&, 0&)
        
        If lvNumListItemsSelected > 0 Then
            lvSomethingSelected = True
        Else
            lvSomethingSelected = False
        End If
        
        If lvSomethingSelected Then
            If lvNumListItemsSelected = 1 Then
                'enable the rename and execute options
                tbEdit.Buttons("Rename").Enabled = True
                mnuItemsRename.Enabled = True
                mnuRCPopupList1Rename.Enabled = True
                
                tbEdit.Buttons("Execute").Enabled = True
                mnuItemsExecute.Enabled = True
                mnuRCPopupList1Execute.Enabled = True
                
                tbStand.Buttons("Prop").Enabled = True
                mnuItemsProperties.Enabled = True
                mnuRCPopupList1Properties.Enabled = True
                
                mnuRCPopupList1Edit.Enabled = True
                tbEdit.Buttons("Edit").Enabled = True
                mnuItemsEdit.Enabled = True
                
            ElseIf lvNumListItemsSelected > 1 Then
                'disable the rename and execute options
                tbEdit.Buttons("Rename").Enabled = False
                mnuItemsRename.Enabled = False
                mnuRCPopupList1Rename.Enabled = False
                
                tbEdit.Buttons("Execute").Enabled = False
                mnuItemsExecute.Enabled = False
                mnuRCPopupList1Execute.Enabled = False
                
                tbStand.Buttons("Prop").Enabled = False
                mnuItemsProperties.Enabled = False
                mnuRCPopupList1Properties.Enabled = False
                
                mnuRCPopupList1Edit.Enabled = False
                tbEdit.Buttons("Edit").Enabled = False
                mnuItemsEdit.Enabled = False
            End If
            
            'enable the delete options
            tbEdit.Buttons("Delete").Enabled = True
            mnuRCPopupList1Delete.Enabled = True
            mnuEditDelete.Enabled = True
            
            'disable paste options if something selected
            mnuRCPopupList2Paste.Enabled = False
            mnuEditPaste.Enabled = False
            tbEdit.Buttons("Paste").Enabled = False
            
            'if they clicked with the right button, then popup the menu
            If RButton = True Then
                PopupMenu mnuRCPopupList1
                RButton = False
            End If
        Else  'nothing is selected in the listview
            'since nothing is selected, we can disable all rename and delete
            'options in toolbars/menus
            tbStand.Buttons("Prop").Enabled = False
            mnuItemsProperties.Enabled = False
            mnuRCPopupList1Properties.Enabled = False
            
            mnuItemsRename.Enabled = False
            tbEdit.Buttons("Rename").Enabled = False
            mnuRCPopupList1Rename.Enabled = False
            
            tbEdit.Buttons("Execute").Enabled = False
            mnuItemsExecute.Enabled = False
            mnuRCPopupList1Execute.Enabled = False
            
            mnuEditDelete.Enabled = False
            tbEdit.Buttons("Delete").Enabled = False
            mnuRCPopupList1Delete.Enabled = False
            
            tbEdit.Buttons("Variation").Enabled = False
            
            'if they clicked with the right button, enable/disable the paste
            'option based on the flag.  This flag is set when an item is cut
            'or copied, and is ready to be pasted.
            If ReadytoPasteItem Then
                mnuRCPopupList2Paste.Enabled = True
                mnuEditPaste.Enabled = True
                tbEdit.Buttons("Paste").Enabled = True
            Else
                mnuRCPopupList2Paste.Enabled = False
                mnuEditPaste.Enabled = False
                tbEdit.Buttons("Paste").Enabled = False
            End If
            If RButton = True Then
                'display the popup menu after the paste has been set appropriately
                PopupMenu mnuRCPopupList2
                RButton = False
            End If
        End If
    End If
End Sub
Private Sub lvListView_DblClick()
    'if a list item is double clicked, attempt to execute it.
    If Not CurrentItem Is Nothing Then
        mnuRCPopupList1Execute_Click
    End If
End Sub

Private Sub lvListView_GotFocus()
    
    Dim itemsSelected As Integer
    
    If LoadingTreeView Then
        Exit Sub
    End If
    
    're-enable the select all/invert selection options
    mnuItemsSelectAll.Enabled = True
    mnuItemsInvertSelection.Enabled = True
    
    
    'this is because of a bug in the listview control that doesn't
    'allow renames to work properly.  This is the workaround.
    If Not ListMouseClick Then
        If Not Devil Then
            SendKeys " "
            Set CurrentItem = lvListView.SelectedItem
            lvSomethingSelected = True
        Else
            Devil = False
            lvListView.StartLabelEdit
        End If
    End If
    
    'enable/disable delete on the toolbar, edit menu, and popup menu
    If lvListView.ListItems.count = 0 Then
        tbEdit.Buttons("Delete").Enabled = False
        mnuEditDelete.Enabled = False
        mnuRCPopupList1Delete.Enabled = False
    Else
        tbEdit.Buttons("Delete").Enabled = True
        mnuEditDelete.Enabled = True
        mnuRCPopupList1Delete.Enabled = True
    End If
    
    
    'enable/disable the rename options
    Select Case lvNumListItemsSelected
        Case 1
            If Not CurrentItem Is Nothing Then
                tbEdit.Buttons("Rename").Enabled = True
                mnuItemsRename.Enabled = True
                mnuRCPopupList1Rename = True
                
                tbEdit.Buttons("Execute").Enabled = True
                mnuItemsExecute.Enabled = True
                mnuRCPopupList1Execute.Enabled = True
                
                mnuRCPopupList1Edit.Enabled = True
                tbEdit.Buttons("Edit").Enabled = True
                mnuItemsEdit.Enabled = True
                
            Else
                tbEdit.Buttons("Rename").Enabled = False
                mnuItemsRename.Enabled = False
                mnuRCPopupList1Rename = False
                
                tbEdit.Buttons("Execute").Enabled = False
                mnuItemsExecute.Enabled = False
                mnuRCPopupList1Execute.Enabled = False
                
                mnuRCPopupList1Edit.Enabled = False
                tbEdit.Buttons("Edit").Enabled = False
                mnuItemsEdit.Enabled = False
                
            End If
        Case Else
            tbEdit.Buttons("Rename").Enabled = False
            mnuItemsRename.Enabled = False
            mnuRCPopupList1Rename = False
            
            tbEdit.Buttons("Execute").Enabled = False
            mnuItemsExecute.Enabled = False
            mnuRCPopupList1Execute.Enabled = False
            
            mnuRCPopupList1Edit.Enabled = False
            tbEdit.Buttons("Edit").Enabled = False
            mnuItemsEdit.Enabled = False
            
    End Select
    
    If Not mnuItems.Enabled Then mnuItems.Enabled = True
    If mnuNodes.Enabled Then mnuNodes.Enabled = False
    
    UpdateStatusBar lvNumListItemsSelected
End Sub
Private Sub lvListView_ItemClick(ByVal item As ListItem)
    
    Dim i As Long
    Dim OneIsVariation As Boolean
    Dim OneIsNotVariation As Boolean

    
    FocusFrom = "ListView"
    lvSomethingSelected = True
    SkipCnt = 0
    TrapMultiSelectDrag = True
    
    Set CurrentItem = item
    
    lvNumListItemsSelected = SendMessage(lvListView.hwnd, LVM_GETSELECTEDCOUNT, 0&, 0&)
    UpdateStatusBar lvNumListItemsSelected
    
    'enable/disable the variaton menu item, and set the data type of the item
    'at the same time.
    If tvAttribCol(nodx.key).Link_NodeID <> 0 Then
        For i = 1 To lvListView.ListItems.count
            If lvListView.ListItems(i).Selected = True Then
                If lvAttribCol(lvListView.ListItems(i).key).variation Then
                    OneIsVariation = True
                Else
                    OneIsNotVariation = True
                End If
                'assume rest of selection is variation or not variation; so exit for
                Exit For
            End If
        Next i
        If OneIsVariation Then
            mnuRCPopupList1Variation.Caption = "Remove Variation"
            tbEdit.Buttons("Variation").ToolTipText = "Remove Variation"
            mnuRCPopupList1Variation.Enabled = True
            tbEdit.Buttons("Variation").Enabled = True
            mnuItemsVariation.Enabled = True
            mnuItemsVariation.Caption = "Remove Variation"
        Else
            mnuRCPopupList1Variation.Caption = "Create Variation"
            tbEdit.Buttons("Variation").ToolTipText = "Create Variation"
            mnuRCPopupList1Variation.Enabled = True
            tbEdit.Buttons("Variation").Enabled = True
            mnuItemsVariation.Enabled = True
            mnuItemsVariation.Caption = "Create Variation"
        End If
    Else
        mnuRCPopupList1Variation.Enabled = False
        tbEdit.Buttons("Variation").Enabled = False
        mnuItemsVariation.Enabled = False
    End If
    
    'disable the properties option if the Properties dialog is already open.
    'otherwise, enable it.
    If PropertiesActive Then
        mnuRCPopupList1Properties.Enabled = False
        mnuItemsProperties.Enabled = False
        tbStand.Buttons("Prop").Enabled = False
        LastTab = fPropForm.SSTab1.Tab
        fPropForm.InitPropertiesForm
        DoEvents
    Else
        LastTab = 0
    End If
End Sub

Private Sub lvListView_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim ShiftDown, AltDown, CtrlDown
    Dim record As Recordset
    
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    
    Select Case KeyCode
        
        'execute the item, if exactly one is selected
        Case vbKeyReturn
            If (lvSomethingSelected) And (lvListView.ListItems.count > 0) Then
                mnuRCPopupList1Execute_Click
            End If
        
        'delete the item, if exactly one is selected
        Case vbKeyDelete
            If (lvSomethingSelected) And (lvListView.ListItems.count > 0) Then
                mnuRCPopupList1Delete_Click
            End If
        
        ' if no items are selected, and we are allowed to add items,
        ' then add a new item
        Case vbKeyInsert
            mnuItemsAdd_Click
        'rename the item, if exactly one is selected
        Case vbKeyR And CtrlDown
            If (lvSomethingSelected) And (lvNumListItemsSelected = 1) Then
                If Not CurrentItem Is Nothing Then
                    If Not lvAttribCol(CurrentItem.key).read_only Then    'if not readonly
                        mnuRCPopupList1Rename_Click
                    End If
                End If
            End If
        
        '''''Brian added this line on 8/21/98''''''
        Case vbKeyF3
        '''''Brian added this line on 8/21/98''''''
            If StartFindItemIndex <> 0 Or StartFindNodeIndex <> 0 Then
        '''''Brian added this line on 8/21/98''''''
                fReplaceForm.cmdFindNext_Click
        '''''Brian added this line on 8/21/98''''''
            End If

        'refresh the node tree
        Case vbKeyF5
            mnuViewRefresh_Click
        
        'select all
        Case vbKeyA And CtrlDown
            lvSomethingSelected = True
        
        'same as performing a right-click on the item
        Case vbKeyF10 And Shift = 1, 93
            If (lvSomethingSelected) And (lvListView.ListItems.count > 0) Then
                'make sure the menu pops up near the selected item
                Call PopupMenu(mnuRCPopupList1, , lvListView.Left + _
                        lvListView.SelectedItem.Left + 200, _
                        lvListView.Top + lvListView.SelectedItem.Top + 200)
            Else
                'enable/disable the paste options depending on whether an item
                'is currently in the buffer and ready to be pasted.
                If ReadytoPasteItem Then
                    tbEdit.Buttons("Paste").Enabled = True
                    mnuEditPaste.Enabled = True
                    mnuRCPopupList2Paste.Enabled = True
                Else
                    tbEdit.Buttons("Paste").Enabled = False
                    mnuEditPaste.Enabled = False
                    mnuRCPopupList2Paste.Enabled = False
                End If
                PopupMenu mnuRCPopupList2
            End If
        Case Else
    End Select
    
    UpdateStatusBar lvNumListItemsSelected
End Sub

Private Sub lvListView_LostFocus()
    ListMouseClick = False
    
    'since we can't do a select all in any other control besides
    'the listview, disable it.
    mnuItemsSelectAll.Enabled = False
    mnuItemsInvertSelection.Enabled = False
End Sub
Private Sub lvListView_MouseDown(Button As Integer, _
                Shift As Integer, X As Single, Y As Single)
    
    Dim i As Long
    
    ListMouseClick = True
    MouseKey = Shift

    'the listview control doesn't right-button drags, and has bugs with
    'right button selects, so we do it manually with the right button
    '(Button=2)
    If Button = 2 Then
        RButton = True
        lParam = Hex(X / Screen.TwipsPerPixelX)
        While Len(lParam) < 4
            lParam = "0" & lParam
        Wend
        lParam = "&H" & Hex(Y / Screen.TwipsPerPixelX) & lParam
    End If

    Set CurrentItem = lvListView.HitTest(X, Y)
    
    If CurrentItem Is Nothing Then
        lvNumListItemsSelected = 0
        TrapMultiSelectDrag = False
    End If
    
    UpdateStatusBar lvNumListItemsSelected
    Set DragItem = lvListView.HitTest(X, Y)
End Sub


Private Sub lvListView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'if the right button is down while the mouse is moving (right dragging),
    'then we also tell the listview control that the left button is down as well.
    'This is necessary, because the listview doesn't support right-dragging, so
    'we fake a left-drag, and check the RButton flag to tell it apart from a real
    'left drag.
    If Button = 2 Then
        RButton = True
        SendMessageLong lvListView.hwnd, WM_LBUTTONDOWN, 1, Val(lParam)
    End If
End Sub

Private Sub lvListView_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    TrapMultiSelectDrag = False
    MouseKey = -1
    
    If CurrentItem Is Nothing Then
       lvListView.SelectedItem = Nothing
    End If

    UpdateStatusBar lvNumListItemsSelected
End Sub
Private Sub lvListView_OLEDragDrop(data As MSComctlLib.DataObject, _
                Effect As Long, Button As Integer, _
                Shift As Integer, X As Single, Y As Single)
    
    Dim i As Long
    
    'skip this function if they clicked on blank space
    If DragItem Is Nothing Then
        GoTo Done
    End If
    
    DragTarget = "LIST"
    Effect = vbDropEffectNone
    
    Set lvListView.DropHighlight = Nothing
    Set CurrentItem = lvListView.HitTest(X, Y)
   
    'if the right button popup menu is to be displayed, enable/disable
    'the move option (can't from a list view into the same list view).
    If RButton = True Then
        If DragSource = "LIST" Then
            mnuDragMove.Enabled = False
        Else
            mnuDragMove.Enabled = True
        End If
        PopupMenu mnuDrag
        RButton = False
    Else
    'if they dragged and dropped with the left button, do a paste
    'since we already did the copy when they started to drag.
        If DragSource = "LIST" Then
            mnuRCPopupList2Paste_Click
        End If
    End If
    
Done:
    'reset the variable since the drag/drop operation is now complete.
    DragSource = ""
    Set DragItem = Nothing
End Sub
Private Sub lvListView_OLEDragOver(data As MSComctlLib.DataObject, _
                Effect As Long, Button As Integer, _
                Shift As Integer, X As Single, Y As Single, _
                state As Integer)
                
    Dim TempItem As ListItem
    Dim i As Long
    
    'check for right button, and don't count it as a click if they
    'are just dragging over the listview.
    If Button = 2 Then
        RButton = True
        IgnoreItemClick = True
    End If
                
    On Error Resume Next
    Set TempItem = lvListView.HitTest(X, Y)
                
    'if we are leaving the listview control, do nothing
    If state = 1 Then 'state=1 means leaving the listview control
        Effect = vbDropEffectNone
        Set lvListView.DropHighlight = Nothing
    Else    'dragging within the listview control
      Select Case DragSource
        Case "TREE"     'do nothing, because we don't allow drags from
                        'the treeview into the listview
            Effect = vbDropEffectNone
            Set lvListView.DropHighlight = Nothing
        Case "LIST"
            If RButton Then 'a right drag within the listview will result
                            'in either a copy or a move
                If TempItem Is Nothing Then 'dragging onto blank space works fine
                   Effect = vbDropEffectCopy
                   Set lvListView.DropHighlight = Nothing
                Else    'dragging onto another item is not allowed
                   Effect = vbDropEffectNone
                   Set lvListView.DropHighlight = TempItem
                End If
            Else    'a left drag within the listview is a copy
                'dragging onto blank space is fine
                If MouseKey = vbCtrlMask And TempItem Is Nothing Then
                    Effect = vbDropEffectCopy
                    Set lvListView.DropHighlight = Nothing
                Else    'dragging onto another item is not allowed
                    Effect = vbDropEffectNone
                    Set lvListView.DropHighlight = Nothing
                End If
            End If
        Case Else
            Effect = vbDropEffectNone
            Set lvListView.DropHighlight = Nothing
      End Select
    End If
End Sub
Private Sub lvListView_OLEStartDrag(data As MSComctlLib.DataObject, _
                AllowedEffects As Long)
    
    DragSource = "LIST"
    
    'if they started to drag blank space, do nothing
    If DragItem Is Nothing Then
        AllowedEffects = vbDropEffectNone 'nothing
    Else 'they are dragging a list item
        'We're going to check to see if the parent of the current node is of
        'Global Type "Inbox".
        'However, we first have to rule out that it is the root node, because
        'the root has no parent and will generate an error if we try to
        'access its parent.
        If nodx.key <> nodx.Root.key Then
            If RemoveK(nodx.parent.key) = FindANode("Inbox", False, True) Then
                DragSubSource = nodx.text
            End If
        End If
        'if the drag was started with the left button with the Ctrl key pressed down
        'OR
        'it was a right drag, then we are going to do a copy.
        If (RButton = False And MouseKey = vbCtrlMask) Or _
                (RButton = True) Then
            AllowedEffects = vbDropEffectCopy 'copy Ctrl+Drag
            mnuRCPopupList1Copy_Click
        'if it was a plain left drag, we are going to do a move
        ElseIf (RButton = False And MouseKey = 0) Then
            AllowedEffects = vbDropEffectMove 'move Drag
            mnuRCPopupList1Cut_Click
        End If
    End If
End Sub



'****************************************************************************
'*********************          OVERLAY FUNCTIONS          ******************
'****************************************************************************

'----------------------------------------------------------------------------
'Creates and overlay icon, if necessary, and returns the index of the newly
'created icon in the DstIml.
'REQUIRES:  SrcIml - the source image list where the original icon is stored.
'           SrcImageKey - the key of the source icon in SrcIml
'           DstIml - the destination image list where the new icon should be
'                    stored.
'           linked - is it a linked node/item?
'           read_only - is it a read_only node/item?
'           system - is it a system node/item?
'----------------------------------------------------------------------------
Public Function CreateOverlayedImage(SrcIml As ImageList, SrcImageKey As String, _
            DstIml As ImageList, Linked As Boolean, _
            read_only As Boolean, System As Boolean) As Long

    Dim DstImageKey As String
    Dim TestIndex As Long
    
    'if there is no icon, the SrcImageKey = "K", so check for that case,
    'and exit if its true.
    If SrcImageKey = "K" Then
        GoTo Done
    End If

    'if source image does not exist then can't make overlay,
    'so return 0 for the icon index
    On Error Resume Next    'reset err variable
    CreateOverlayedImage = -1
    CreateOverlayedImage = SrcIml.ListImages(SrcImageKey).Index
    If CreateOverlayedImage = -1 Then
        CreateOverlayedImage = 0
        On Error GoTo 0
        GoTo Done
    End If
    On Error GoTo 0
    
    'if there is an icon, append a dash to the key, followed by tags denoting
    'different states that would necessitate an overlay.
    DstImageKey = SrcImageKey & "-"
    If Linked Then
        DstImageKey = DstImageKey & "L"
    End If
    If read_only Then
        DstImageKey = DstImageKey & "R"
    End If
    If System Then
        DstImageKey = DstImageKey & "S"
    End If

    'check to see if this item is already in the destination image list
    'by trying to access it.  If its there, it means we have a duplicate,
    'so return its index and exit.  If its not there, it generates an
    'error, so we trap it and continue.
    On Error Resume Next
    CreateOverlayedImage = -1
    CreateOverlayedImage = DstIml.ListImages(DstImageKey).Index
    If CreateOverlayedImage <> -1 Then
        On Error GoTo 0
        GoTo Done
    End If
    On Error GoTo 0
    
    'if we got to here, it means the icon is not a duplicate, so we
    'should add it to the image list.
    
    'reset the key to get base icon to add overlays to
    DstImageKey = SrcImageKey & "-"
    'add the base icon
    On Error Resume Next
    TestIndex = -1
    TestIndex = DstIml.ListImages(DstImageKey).Index
    If TestIndex = -1 Then
        On Error GoTo 0
        DstIml.ListImages.Add , DstImageKey, _
            SrcIml.ListImages(SrcImageKey).picture
    End If
    On Error GoTo 0
        
    'add the appropriate overlays
    If Linked Then
        OverlayNextImage DstIml, DstImageKey, DstImageKey & "L", "K" & Link_Icon
        DstImageKey = DstImageKey & "L"
'        OverlayNextImage DstIml, DstImageKey, "K" & Link_Icon
    End If
    
    If read_only Then
        OverlayNextImage DstIml, DstImageKey, DstImageKey & "R", "K" & Read_Icon
        DstImageKey = DstImageKey & "R"
'        OverlayNextImage DstIml, DstImageKey, "K" & Read_Icon
    End If
    
    If System Then
        OverlayNextImage DstIml, DstImageKey, DstImageKey & "S", "K" & System_Icon
        DstImageKey = DstImageKey & "S"
'        OverlayNextImage DstIml, DstImageKey, "K" & System_Icon
    End If
    
    'return the index of the newly created icon
    CreateOverlayedImage = DstIml.ListImages(DstImageKey).Index
Done:
End Function
'----------------------------------------------------------------------------
'Replaces an image in the imagelist with an overlayed image.
'REQUIRES:  ImageList - the image list containing the image to be replaced.
'           ImageKey - the key of the image to be replaced.
'           OverlayImageKey - the key of the overlayed image which will replace
'                             the original.
'----------------------------------------------------------------------------
Private Sub OverlayNextImage(Iml As ImageList, SrcImageKey As String, _
        DstImageKey As String, OverlayImageKey As String)
    
    Dim TestIndex As Long
    
    On Error Resume Next
    TestIndex = -1
    TestIndex = Iml.ListImages(DstImageKey).Index
    If TestIndex = -1 Then
        On Error GoTo 0
        'overlay the image and temporarily store it in the image control
        Set TempImage.picture = Iml.Overlay(SrcImageKey, OverlayImageKey)
        'add the overlayed image from the image control to the image list
        Iml.ListImages.Add , DstImageKey, TempImage.picture
    End If
    On Error GoTo 0

End Sub
'----------------------------------------------------------------------------
'Sets the icons for the items in the listview when the view mode is
'lage icons.
'----------------------------------------------------------------------------
Public Sub SetLargeOverlays()
    Dim i As Long
    Dim AddL As Boolean
    Dim indexoffset As Long
    
    'loop through all current list items and store the icon index in the
    'tag field
    For i = 1 To lvListView.ListItems.count
        If Left$(lvListView.ListItems(i).key, 1) = "L" Then
            AddL = True
        Else
            AddL = False
        End If
        lvListView.ListItems(i).tag = CreateOverlayedImage(imlIconsLarge, "K" & _
            lvAttribCol(lvListView.ListItems(i).key).icon_large, imlTempLarge, _
            AddL, lvAttribCol(lvListView.ListItems(i).key).read_only, _
            tvAttribCol(nodx.key).system_node)
    Next i
    
    'bind the listview to the image list
    If imlTempLarge.ListImages.count <> 0 Then
        lvListView.Icons = imlTempLarge
    End If
    
    'copy the icon index from the tag field to the icon
    For i = 1 To lvListView.ListItems.count
        lvListView.ListItems(i).Icon = lvListView.ListItems(i).tag
    Next i
End Sub
'----------------------------------------------------------------------------
'Sets the icons for the items in the listview when the view mode is
'small icons, details, or list.
'----------------------------------------------------------------------------
Public Sub SetSmallOverlays()
    Dim i As Long
    Dim AddL As Boolean
    Dim indexoffset As Long
    
    'loop through all current list items and store the icon index in the
    'tag field
    For i = 1 To lvListView.ListItems.count
        If Left$(lvListView.ListItems(i).key, 1) = "L" Then
            AddL = True
        Else
            AddL = False
        End If
        lvListView.ListItems(i).tag = CreateOverlayedImage(imlIconsSmall, "K" & _
            lvAttribCol(lvListView.ListItems(i).key).icon_small, imlTempSmall, _
            AddL, lvAttribCol(lvListView.ListItems(i).key).read_only, _
            tvAttribCol(nodx.key).system_node)
    Next i
    
    'bind the listview to the image list
    If imlTempSmall.ListImages.count <> 0 Then
        lvListView.SmallIcons = imlTempSmall
    End If
    
    'copy the icon index from the tag field to the icon
    For i = 1 To lvListView.ListItems.count
        lvListView.ListItems(i).SmallIcon = lvListView.ListItems(i).tag
    Next i
    
End Sub
'----------------------------------------------------------------------------
'Sets the icons for the treeview
'----------------------------------------------------------------------------
Public Sub SetTreeOverlays()
    Dim i As Long
    Dim AddL As Boolean
    Dim indexoffset As Long
    Dim commapos As Integer
    
    'loop through all of the tree nodes
    For i = 1 To tvTreeView.Nodes.count
        If tvAttribCol(tvTreeView.Nodes(i).key).Link_NodeID <> 0 Then
            AddL = True
        Else
            AddL = False
        End If
        
        'store the normal and selected icon indeces in the tag field of the node
        tvTreeView.Nodes(i).tag = Format$(CreateOverlayedImage(imlIconsSmall, "K" & _
            tvAttribCol(tvTreeView.Nodes(i).key).icon_normal, imlTempTree, _
            AddL, tvAttribCol(tvTreeView.Nodes(i).key).read_only, _
            tvAttribCol(tvTreeView.Nodes(i).key).system_node))
        tvTreeView.Nodes(i).tag = tvTreeView.Nodes(i).tag & "," & _
            Format$(CreateOverlayedImage(imlIconsSmall, "K" & _
            tvAttribCol(tvTreeView.Nodes(i).key).icon_selected, imlTempTree, _
            AddL, tvAttribCol(tvTreeView.Nodes(i).key).read_only, _
            tvAttribCol(tvTreeView.Nodes(i).key).system_node))
    Next i
    
    'bind the image list to the treeview
    tvTreeView.ImageList = imlTempTree
    
    'copy the icon indeces out of the tag property and into the
    'appropriate icon properties
    For i = 1 To tvTreeView.Nodes.count
        commapos = InStr(1, tvTreeView.Nodes(i).tag, ",", vbTextCompare)
        tvTreeView.Nodes(i).Image = Val(Left$(tvTreeView.Nodes(i).tag, commapos - 1))
        tvTreeView.Nodes(i).SelectedImage = Val(Right$(tvTreeView.Nodes(i).tag, _
                Len(tvTreeView.Nodes(i).tag) - commapos))
        tvTreeView.Nodes(i).ExpandedImage = tvTreeView.Nodes(i).SelectedImage
    Next i
End Sub











'****************************************************************************
'*********************           MENU FUNCTIONS            ******************
'****************************************************************************

Private Sub mnuDragCancel_Click()
    DragDropped = False
    Set DragItem = Nothing
End Sub
Private Sub mnuDragCopy_Click()
    
    Select Case DragSource
        Case "TREE"
            'if we're doing a tree-to-tree copy, then it's okay
            'do a paste
            If DragTarget = "TREE" Then
                mnuRCPopupTree1Paste_Click
            End If
        Case "LIST"
            'if we're doing a list-to-tree copy, then it's okay
            'do a paste, and select the target node
            If DragTarget = "TREE" Then
                mnuRCPopupList2Paste_Click
                Set nodx = tvTreeView.SelectedItem
                lvNeedsRefresh = True
                tvTreeView_NodeClick nodx
                lblTitle(1).Caption = "Contents of '" & nodx.text & "'"
            'if we're doing a list-to-list copy, then it's okay
            'do a paste
            ElseIf DragTarget = "LIST" Then
                mnuRCPopupList2Paste_Click
            End If
    'all other cases are ignored (i.e. a tree-to-list copy)
    End Select
End Sub

Private Sub mnuDragMove_Click()
    Dim i As Long
    
    Select Case DragSource
        Case "TREE"
            'if we're doing a tree-to-tree move, then it's okay
            'change the implied copy to a cut and do a paste
            NodeWasCut = True
            If DragTarget = "TREE" Then
            
                mnuRCPopupTree1Paste_Click
            End If
        Case "LIST"
            'if we're doing a list-to-tree move, then it's okay
            'change the implied copy to a cut, do a paste,
            'and select the target node
            ItemWasCut = True
            If DragTarget = "TREE" Then
                mnuRCPopupList2Paste_Click
                Set nodx = tvTreeView.SelectedItem
                lvNeedsRefresh = True
                tvTreeView_NodeClick nodx
                lblTitle(1).Caption = "Contents of '" & nodx.text & "'"
            'if we're doing a list-to-list move, then it's okay
            'change the implied copy to a cut, and do a paste
            ElseIf DragTarget = "LIST" Then
                mnuRCPopupList2Paste_Click
            End If
    'all other cases are ignored (i.e. a tree-to-list move)
    End Select
End Sub

Private Sub mnuEditDelete_Click()
    If TypeOf Me.ActiveControl Is TreeView Then
        mnuRCPopupTree1Delete_Click
    ElseIf TypeOf Me.ActiveControl Is ListView Then
        mnuRCPopupList1Delete_Click
    End If
End Sub
Private Sub mnuEditFind_Click()
    'if FindEnabled = True, we're doing a regular find.  Otherwise,
    'we're doing a find & replace.
    FindEnabled = True
    fReplaceForm.Show
    fReplaceForm.InitReplaceForm
End Sub


Private Sub mnuFileDBProperties_Click()
    frmDBProperties.Show vbModal
End Sub

Private Sub mnuItemsAdd_Click()
    isInsertKey = True
    PopupMenu mnuRCPopupList2
End Sub

Private Sub mnuItemsExecute_Click()
    mnuRCPopupList1Execute_Click
End Sub

Private Sub mnuItemsProperties_Click()
    If (Not CurrentItem Is Nothing) Then
        Me.MousePointer = vbHourglass
        PropertiesActive = True
        LastTab = 0
        FocusFrom = "ListView"
        fPropForm.Show vbModeless, Me
        fPropForm.InitPropertiesForm
        Me.MousePointer = vbArrow
    End If
End Sub

Private Sub mnuItemsRename_Click()
    mnuRCPopupList1Rename_Click
End Sub

Private Sub mnuItemsVariation_Click()
    mnuRCPopupList1Variation_Click
End Sub

Private Sub mnuNodesProperties_Click()
    If (Not nodx Is Nothing) Then
        Me.MousePointer = vbHourglass
        PropertiesActive = True
        LastTab = 0
        FocusFrom = "TreeView"
        fPropForm.Show vbModeless, Me
        fPropForm.InitPropertiesForm
        Me.MousePointer = vbArrow
    End If
End Sub

Private Sub mnuNodesAdd_Click()
    Dim i As Long
    Dim j As Long
    Dim n As Long
    
    isInsertKey = True
    i = 0
    n = nodx.Index
    While tvTreeView.Nodes(n).Visible And n <> nodx.Root.Index
        n = tvTreeView.Nodes(n).parent.Index
        i = i + 1
    Wend
        
    j = 0
    n = nodx.Index
    While tvTreeView.Nodes(n).Visible And n <> nodx.Root.Index
        j = j + FindNodeHeight(n)
        n = tvTreeView.Nodes(n).parent.Index
    Wend
    If Not tvTreeView.Nodes(n).Visible Then
        j = j - 1
    End If
    Call PopupMenu(mnuRCPopupTree1, , i * (tvTreeView.Indentation + 30) + _
        tvTreeView.Left + 500, j * 240 + tvTreeView.Top + 200)
End Sub

Private Sub mnuNodesMoveDown_Click()
    mnuRCPopupTree1MoveDown_Click
End Sub

Private Sub mnuNodesMoveUp_Click()
    mnuRCPopupTree1MoveUp_Click
End Sub

Private Sub mnuEditReplace_Click()
    'if FindEnabled = True, we're doing a regular find.  Otherwise,
    'we're doing a find & replace.
    FindEnabled = False
    fReplaceForm.Show
    fReplaceForm.InitReplaceForm
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePrint_Click()
    CommonDialog1.Flags = 0
    CommonDialog1.CancelError = True
    On Error GoTo Done
    CommonDialog1.ShowPrinter
    MsgBox "TEMP:  Code for printing will go here.", vbInformation
Done:
End Sub

Private Sub mnuHelpAbout_Click()
    frmSplash.Show vbModal, Me
End Sub

Private Sub mnuHelpQuote_Click()
    frmQuotes.Show vbModal, Me
End Sub
Private Sub mnuInboxAutoSort_Click()
        
    Dim i As Long
    Dim tempcategory As String
    Dim ext As String
    Dim record As Recordset
    Dim icon_large As String
    Dim icon_small As String
    Dim rsp As Integer
    Dim NodeNum As Long
    Dim data_type As String
    Dim InboxNode As Node
    Dim AddedFiles As Long
    Dim TotalFiles As Long
    Dim TempItem As ListItem
    Dim tempstr As String
    Dim comma_pos As Integer
    
    Me.MousePointer = vbHourglass
    
    TotalFiles = CurrentData.Files.count
    AddedFiles = 0
    
    'expand the Inbox and make sure it is visible
    nodx.Expanded = True
    nodx.EnsureVisible
    Set InboxNode = nodx
    
    'loop through for each file that was dropped onto the Inbox
    For i = 1 To CurrentData.Files.count
        tempcategory = "Binary (Misc.)"
        'extract the extension for each file (without the '.').  Will return
        'URL for http:// types and FTP for ftp:// types.
        ext = FileExtension(CurrentData.Files(i))
        'if the extension is invalid, ask the user whether they
        'want to continue with the next one, or abort the operation.
        If ext = "INVALID" Then
            rsp = MsgBox("The file '" & CurrentData.Files(i) & "' cannot be processed " _
                    & "because it does not contain a valid extension." _
                    & Chr(13) & "Click 'OK' to try the next file, if there is one." _
                    & Chr(13) & "Click 'Cancel' to abort the operation.", _
                    vbExclamation + vbOKCancel)
            If rsp = vbOK Then
                GoTo TryNext
            Else
                GoTo Done
            End If
        Else
            'if the extension is a valid one, query the database for its type and icons
            Set record = dbase.OpenRecordset("SELECT Data_Value, data_type, Icon_Large, Icon_Small " _
                & "FROM " & DB_ExtensionsTable _
                & " WHERE Data_Label = '" & ext & "'", dbOpenDynaset)
            If Not record.EOF Then
                'if the extension is defined in the database, get its type, large icon,
                'and small icon
                tempstr = record!data_value
                comma_pos = InStr(1, tempstr, ",")
            
                If comma_pos = 0 Then
                    rsp = MsgBox("There is a possible error in the database.  Please check the " _
                        & "entry for the extension '" & ext & "' in the '" _
                        & DB_ExtensionsTable & "' table and make sure that " _
                        & "the [Data_Value] field is in the form 'category,data type' " _
                        & "(no spaces before or after the comma)." _
                        & Chr(13) & "Click 'OK' to try the next file, if there is one." _
                        & Chr(13) & "Click 'Cancel' to abort the operation.", _
                        vbCritical + vbOKCancel)
                    If rsp = vbOK Then
                        On Error GoTo 0
                        Set nodx = InboxNode
                        GoTo TryNext
                    Else
                        GoTo Done
                    End If
                End If
                
                data_type = mID$(tempstr, comma_pos + 1, Len(tempstr))
                tempcategory = Left$(tempstr, comma_pos - 1)
                icon_large = record!icon_large
                icon_small = record!icon_small
            Else
                'if the extension is not in the database, ask the user whether they
                'want to continue with the next one, or abort the operation.
                data_type = "String"
                tempcategory = "Miscellaneous Files"
            End If
            record.Close
        
            'scan children of Inbox until we find the one matching the
            'category we want.
            Set nodx = nodx.Child
            On Error GoTo NoNext
            While RemoveK(nodx.key) <> FindANode(tempcategory, False, True)
            'While nodx.text <> tempcategory
                'if the desired category doesn't exist jump to 'NoNext'
                On Error GoTo NoNext
                Set nodx = nodx.Next
                If nodx Is Nothing Then
                    GoTo NoNext
                End If
            Wend
            On Error GoTo 0
            GoTo NoError
NoNext:
            'if we searched all the children of 'Inbox' and couldn't find
            'the category, ask the user if they want to skip to the next
            'one or abort the operation.
            rsp = MsgBox("The file '" & CurrentData.Files(i) & "' cannot be processed " _
                    & "because the required node, '" & tempcategory _
                    & "', was not found under the Inbox." _
                    & Chr(13) & "Click 'OK' to try the next file, if there is one." _
                    & Chr(13) & "Click 'Cancel' to abort the operation.", _
                    vbCritical + vbOKCancel)
            If rsp = vbOK Then
                On Error GoTo 0
                Set nodx = InboxNode
                GoTo TryNext
            Else
                GoTo Done
            End If
NoError:
            'if we found the correct category with no problems, then select
            'that node and add the file as a new list item.
            nodx.Selected = True
            
            'check for duplicates
            Set record = dbase.OpenRecordset("SELECT [Data_ID] FROM " _
                & tvAttribCol(nodx.key).table_name & " WHERE [Data_Label] = '" _
                & CurrentData.Files(i) & "' AND [Parent_Node] = " _
                & RemoveK(nodx.key), dbOpenDynaset)

            If Not record.EOF Then
                rsp = MsgBox("The item '" & CurrentData.Files(i) & "' already exists " _
                    & "in the Inbox node '" & nodx.text & "'.  Do you wish to overwrite " _
                    & "it?  NOTE:  If you choose 'No', this file will be skipped.", vbYesNo + vbExclamation)
                If rsp = vbYes Then
                    'delete existing item
                    record.Delete
                Else
                    Set nodx = InboxNode
                    record.Close
                    GoTo TryNext
                End If
            End If
            record.Close
            
            
            NodeNum = AddItem(icon_large, icon_small, CurrentData.Files(i), _
                    False, data_type, Now, CurrentUser, Now, CurrentUser)
            
            'add the file into the database, too
            Set record = dbase.OpenRecordset("SELECT data_value,binary_data_value FROM " _
                & tvAttribCol(nodx.key).table_name & " WHERE Data_ID = " _
                & NodeNum, dbOpenDynaset)
            ReadBLOB CurrentData.Files(i), record, "binary_data_value"
            record.Close
            AddedFiles = AddedFiles + 1
            
            Set nodx = nodx.parent
        End If
TryNext:
    
    Next i
    MsgBox AddedFiles & " (out of " & TotalFiles & " requested) files were " _
        & "added to the Inbox.", vbInformation
Done:
    Me.MousePointer = vbArrow
    lvNeedsRefresh = True
    Set tvTreeView.SelectedItem = nodx
    tvTreeView.Nodes(nodx.key).EnsureVisible
    tvTreeView_NodeClick nodx
End Sub

Private Sub mnuInboxCancel_Click()
    DragDropped = False
End Sub

Private Sub mnuInboxCopy_Click()
    Dim i As Long
    Dim ItemFileName As String
    Dim icon_large As String
    Dim icon_small As String
    Dim data_type As String
    Dim AddedFiles As Long
    Dim TotalFiles As Long
    Dim ext As String
    Dim record As Recordset
    Dim rsp As Integer
    Dim NodeNum As Long
    Dim data_label As String
    Dim tempstr As String
    Dim comma_pos As Integer
    
    LoadListView nodx.key
   
    Me.MousePointer = vbHourglass
    
    'if the dragged items are files
    If CurrentData.GetFormat(ccCFFiles) Then
        TotalFiles = CurrentData.Files.count
        AddedFiles = 0
    
        For i = 1 To CurrentData.Files.count
            ext = FileExtension(CurrentData.Files(i))
        
            If ext <> "INVALID" Then
                Set record = dbase.OpenRecordset("SELECT data_value, Icon_Large, Icon_Small " _
                    & "FROM " & DB_ExtensionsTable _
                    & " WHERE Data_Label = '" & ext & "'", dbOpenDynaset)
                If Not record.EOF Then
                    'if the extension is defined in the database, get its type, large icon,
                    'and small icon
                    tempstr = record!data_value
                    comma_pos = InStr(1, tempstr, ",")
            
                    If comma_pos = 0 Then
                        rsp = MsgBox("There is a possible error in the database.  Please check the " _
                            & "entry for the extension '" & ext & "' in the '" _
                            & DB_ExtensionsTable & "' table and make sure that " _
                            & "the [Data_Value] field is in the form 'category,data type' " _
                            & "(no spaces before or after the comma)." _
                            & Chr(13) & "Click 'OK' to try the next file, if there is one." _
                            & Chr(13) & "Click 'Cancel' to abort the operation.", _
                            vbCritical + vbOKCancel)
                        If rsp = vbOK Then
                            On Error GoTo 0
                            GoTo TryNext
                        Else
                            GoTo Done
                        End If
                    End If
                    data_type = mID$(tempstr, comma_pos + 1, Len(tempstr))
                    icon_large = record!icon_large
                    icon_small = record!icon_small
                Else
                    'if the extension is not in the database, ask the user whether they
                    'want to continue with the next one, or abort the operation.
                    data_type = "String"
                End If
                record.Close
            End If
                
            Set record = dbase.OpenRecordset("SELECT [Data_ID] FROM " _
                & tvAttribCol(nodx.key).table_name & " WHERE [Data_Label] = '" _
                & CurrentData.Files(i) & "' AND [Parent_Node] = " _
                & RemoveK(nodx.key), dbOpenDynaset)

            If Not record.EOF Then
                rsp = MsgBox("The item '" & CurrentData.Files(i) & "' already exists " _
                    & "in the Inbox node '" & nodx.text & "'.  Do you wish to overwrite " _
                    & "it?  NOTE:  If you choose 'No', this file will be skipped.", vbYesNo + vbExclamation)
                If rsp = vbYes Then
                    'delete existing item
                    DeleteListItem AddK(record!data_id), nodx.key, False
                Else
                    record.Close
                    GoTo TryNext
                End If
            End If
            record.Close
           
            NodeNum = AddItem(icon_large, icon_small, CurrentData.Files(i), False, _
                data_type, Now, CurrentUser, Now, CurrentUser)
            
            'add the file into the database, too
            Set record = dbase.OpenRecordset("SELECT data_value,binary_data_value FROM " _
                & tvAttribCol(nodx.key).table_name & " WHERE Data_ID = " _
                & NodeNum, dbOpenDynaset)
            ReadBLOB CurrentData.Files(i), record, "binary_data_value"
            record.Close
            AddedFiles = AddedFiles + 1
            
TryNext:
        Next i
   
    
    'if the dragged items are type text (this is also used for
    'dragging URLs from Netscape or Internet Explorer.
    ElseIf CurrentData.GetFormat(ccCFText) Then
        TotalFiles = 1
        AddedFiles = 0
        'HTTP url
        If UCase(Left$(CurrentData.GetData(ccCFText), 7)) = "HTTP://" Then
            
            Set record = dbase.OpenRecordset("SELECT Data_Type, Icon_Large, Icon_Small " _
                & "FROM " & DB_ExtensionsTable _
                & " WHERE Data_Label = 'HTTP'", dbOpenDynaset)
            If Not record.EOF Then
                'if the extension is defined in the database, get its type, large icon,
                'and small icon
                data_type = record!data_type
                icon_large = record!icon_large
                icon_small = record!icon_small
            Else
                data_type = "String"
            End If
        'FTP url
        ElseIf UCase(Left$(CurrentData.GetData(ccCFText), 7)) = "FTP://" Then
            Set record = dbase.OpenRecordset("SELECT Data_Type, Icon_Large, Icon_Small " _
                & "FROM " & DB_ExtensionsTable _
                & " WHERE Data_Label = 'FTP'", dbOpenDynaset)
            If Not record.EOF Then
                'if the extension is defined in the database, get its type, large icon,
                'and small icon
                data_type = record!data_type
                icon_large = record!icon_large
                icon_small = record!icon_small
            Else
                data_type = "String"
            End If
        'some other text
        Else
            data_type = "String"
        End If
        
        If Len(CurrentData.GetData(ccCFText)) < 64 Then
            data_label = CurrentData.GetData(ccCFText)
        Else
            data_label = Left$(CurrentData.GetData(ccCFText), 64)
        End If
        
        Set record = dbase.OpenRecordset("SELECT [Data_ID] FROM " _
            & tvAttribCol(nodx.key).table_name & " WHERE [Data_Label] = '" _
            & CurrentData.GetData(ccCFText) & "' AND [Parent_Node] = " _
            & RemoveK(nodx.key), dbOpenDynaset)

        If Not record.EOF Then
            rsp = MsgBox("The item '" & CurrentData.GetData(ccCFText) & "' already exists " _
                & "in the Inbox node '" & nodx.text & "'.  Do you wish to overwrite " _
                & "it?  NOTE:  If you choose 'No', this file will be skipped.", vbYesNo + vbExclamation)
            If rsp = vbYes Then
                'delete existing item
                lvListView.ListItems.Remove AddK(record!data_id)
                lvAttribCol.Remove AddK(record!data_id)
                record.Delete
            Else
                record.Close
                GoTo Done
            End If
        End If
        record.Close
        
        NodeNum = AddItem(icon_large, icon_small, data_label, _
            False, data_type, Now, CurrentUser, Now, CurrentUser)
        
        Set record = dbase.OpenRecordset("SELECT data_value FROM " _
                & tvAttribCol(nodx.key).table_name & " WHERE Data_ID = " _
                & NodeNum, dbOpenDynaset)
        
        If Not record.EOF Then
            record.Edit
            record!data_value = CurrentData.GetData(ccCFText)
            record.Update
            record.Close
            AddedFiles = AddedFiles + 1
        Else
            MsgBox "mnuInboxCopy:  Could not add item to the database!", vbExclamation
        End If
    End If
    
Done:
    Me.MousePointer = vbArrow
    lvNeedsRefresh = True
    Set tvTreeView.SelectedItem = nodx
    tvTreeView.Nodes(nodx.key).EnsureVisible
    tvTreeView_NodeClick nodx
    MsgBox AddedFiles & " (out of " & TotalFiles & " requested) items were " _
        & "added to the Inbox.", vbInformation
    
End Sub

Private Sub mnuModuleHTML_Click()
    LoadBlankHTML = True    'so the HTML Editor doesn't automatically try
                            'to load data from the current item.
    
    MsgBox "TEMP:  Load HTML Editor"
End Sub

Private Sub mnuModuleSource_Click()
    LoadBlankSourceViewer = True    'so the Source Editor doesn't automatically try
                                    'to load data from the current item.
    MsgBox "TEMP: Load Source Viewer"
    'frmHTMLSource.Show vbModal, Me
    LoadBlankSourceViewer = False
End Sub







'****************************************************************************
'*********************           MISC FUNCTIONS            ******************
'****************************************************************************

'----------------------------------------------------------------------------
'Loads all the icons from the icon paths and stores them into the desired
'tables.  Also fill 3 combo boxes with the icons, for use in the Properties
'dialog.
'REQUIRES:  names of the large and small icon tables.
'----------------------------------------------------------------------------
Private Sub LoadIcons(Optional UseProgressBar As Boolean = False)

    Dim tempstr As String
    Dim largeIconRecord As Recordset
    Dim smallIconRecord As Recordset
    Dim record As Recordset
    Dim count As Long
    Dim LastCount As Long
    Dim LargeNodeNum As Long
    Dim SmallNodeNum As Long

    '''''Brian added this line on 8/21/98''''''
    Dim LargeOrder As Long
    '''''Brian added this line on 8/21/98''''''
    Dim SmallOrder As Long
    
    Dim imgx As ListImage
    Dim i As Long
    Dim TempKey As String
    Dim Icon As String

    If SHOW_TIMING_INFO Then
        Dim t As Variant
        Dim t_total As Variant
        t = Timer
        t_total = t
    End If
    Me.MousePointer = vbHourglass
    
    'clear out the icon image lists
    imlIconsLarge.ListImages.Clear
    imlIconsSmall.ListImages.Clear
    
    'unbind the imagelists from the image combo controls
    fPropForm.cboIcon(1).ImageList = Nothing
    fPropForm.cboIcon(2).ImageList = Nothing
    fPropForm.cboIcon(3).ImageList = Nothing
    
    'clear out the image combo controls
    fPropForm.cboIcon(1).ComboItems.Clear
    fPropForm.cboIcon(2).ComboItems.Clear
    fPropForm.cboIcon(3).ComboItems.Clear
    

    ' set size of large icons
    imlIconsLarge.ImageHeight = LARGE_ICON_SIZE
    imlIconsLarge.ImageWidth = LARGE_ICON_SIZE
        
    ' set size of small icons
    imlIconsSmall.ImageHeight = SMALL_ICON_SIZE
    imlIconsSmall.ImageWidth = SMALL_ICON_SIZE

    'save the parent nodes for both tables
    LargeNodeNum = FindANode("Large Icons", False, True)
    SmallNodeNum = FindANode("Small Icons", False, True)
        
    Set record = dbase.OpenRecordset("SELECT Node_ID,[Order] FROM " _
        & DB_NodeTable & " WHERE Node_ID IN (" & LargeNodeNum & "," & SmallNodeNum _
        & ")", dbOpenDynaset)
    LargeOrder = record!order
    record.MoveNext
    SmallOrder = record!order
    record.Close
    
    ' clear out both tables
    On Error GoTo Err_Execute
    dbase.Execute "DELETE * FROM " & DB_LargeIconsTable, dbFailOnError
    dbase.Execute "DELETE * FROM " & DB_SmallIconsTable, dbFailOnError
    On Error GoTo 0
    
    Set largeIconRecord = dbase.OpenRecordset(DB_LargeIconsTable, dbOpenTable)
    Set smallIconRecord = dbase.OpenRecordset(DB_SmallIconsTable, dbOpenTable)
    
    If UseProgressBar Then
        count = 1
        tempstr = Dir(LargeIconFolder & "*.bmp", vbNormal)
        While tempstr <> ""
            tempstr = Dir
            count = count + 1
        Wend
        If RefreshInProgress Then
            tempstr = "Refreshing Icons . . ."
        Else
            tempstr = "Loading Icons . . ."
        End If
        InitProgressBar fProgForm, tempstr, 0, count * 4, _
            LargeIconFolder & PBICON_LoadIcon, , InFormLoad
    End If
    
    If SHOW_TIMING_INFO Then
        Debug.Print "LoadIcons:  Overhead = " & Timer - t
        t = Timer
    End If
    
    ' fill the large icons table and the image list at the same time
    count = 1
    tempstr = Dir(LargeIconFolder & "*.bmp", vbNormal)
    While tempstr <> ""
        tempstr = Left$(tempstr, Len(tempstr) - 4)
        largeIconRecord.AddNew
        largeIconRecord!data_id = count
        largeIconRecord!data_label = tempstr
        largeIconRecord!icon_large = tempstr
        largeIconRecord!read_only = True
        largeIconRecord!parent_node = LargeNodeNum
        largeIconRecord!created = Now
        largeIconRecord!created_by = "Osiris"
        largeIconRecord!last_modified = Now
        largeIconRecord!modified_by = "Osiris"
        largeIconRecord!order = LargeOrder
        largeIconRecord!variation = False
        largeIconRecord.Update
        tempstr = Dir
        count = count + 1
        If UseProgressBar Then
            fProgForm.pbPBar1.Value = count
        End If
    Wend
    LastCount = count
    largeIconRecord.Close
    
    If SHOW_TIMING_INFO Then
        Debug.Print "LoadIcons:  Create Large Icons table = " & Timer - t
        t = Timer
    End If
    
    Set largeIconRecord = dbase.OpenRecordset("SELECT data_label FROM " _
        & DB_LargeIconsTable & " ORDER BY Data_Label", dbOpenDynaset)
    While Not largeIconRecord.EOF
        tempstr = largeIconRecord!data_label
        Set imgx = imlIconsLarge.ListImages.Add(, "K" & tempstr, _
            LoadPicture(LargeIconFolder & tempstr & ".bmp"))
        largeIconRecord.MoveNext
    Wend
    
    If SHOW_TIMING_INFO Then
        Debug.Print "LoadIcons:  Build Large Icons Image List = " & Timer - t
        t = Timer
    End If
    
    ' fill the small icons table and the image list
    count = 1
    tempstr = Dir(SmallIconFolder & "*.bmp", vbNormal)
    While tempstr <> ""
        tempstr = Left$(tempstr, Len(tempstr) - 4)
        smallIconRecord.AddNew
        smallIconRecord!data_id = count
        smallIconRecord!data_label = tempstr
        smallIconRecord!icon_small = tempstr
        smallIconRecord!read_only = True
        smallIconRecord!parent_node = SmallNodeNum
        smallIconRecord!created = Now
        smallIconRecord!created_by = "Osiris"
        smallIconRecord!last_modified = Now
        smallIconRecord!modified_by = "Osiris"
        smallIconRecord!order = SmallOrder
        smallIconRecord!variation = False
        smallIconRecord.Update
        tempstr = Dir
        count = count + 1
        If UseProgressBar Then
            fProgForm.pbPBar1.Value = LastCount + count
        End If
    Wend
    LastCount = LastCount + count
    smallIconRecord.Close
    
    If SHOW_TIMING_INFO Then
        Debug.Print "LoadIcons:  Create Small Icons Table = " & Timer - t
        t = Timer
    End If
    
    Set smallIconRecord = dbase.OpenRecordset("SELECT data_label FROM " _
        & DB_SmallIconsTable & " ORDER BY Data_Label", dbOpenDynaset)
    While Not smallIconRecord.EOF
        tempstr = smallIconRecord!data_label
        Set imgx = imlIconsSmall.ListImages.Add(, "K" & tempstr, _
            LoadPicture(SmallIconFolder & tempstr & ".bmp"))
        smallIconRecord.MoveNext
    Wend
    
    If SHOW_TIMING_INFO Then
        Debug.Print "LoadIcons:  Build Small Icons Image List = " & Timer - t
    End If
    
    largeIconRecord.Close
    smallIconRecord.Close
    
    'assign image lists to the icon combo boxes
    If imlIconsSmall.ListImages.count <> 0 Then
        Set fPropForm.cboIcon(1).ImageList = imlIconsSmall
        Set fPropForm.cboIcon(2).ImageList = imlIconsSmall
    End If
    If imlIconsLarge.ListImages.count <> 0 Then
        Set fPropForm.cboIcon(3).ImageList = imlIconsLarge
    End If
    
    
    'open the small icons image list
    
    'add the None items to each image combo control
    fPropForm.cboIcon(1).ComboItems.Add , NONE_LABEL, NONE_LABEL, , , 0
    fPropForm.cboIcon(2).ComboItems.Add , NONE_LABEL, NONE_LABEL, , , 0
    fPropForm.cboIcon(3).ComboItems.Add , NONE_LABEL, NONE_LABEL, , , 0
  
    t = Timer
  
    'build the icon combo boxes
    For i = 1 To imlIconsSmall.ListImages.count
        TempKey = imlIconsSmall.ListImages.item(i).key
        Icon = mID$(TempKey, 2, Len(TempKey) - 1)
        fPropForm.cboIcon(1).ComboItems.Add , TempKey, Icon, _
            imlIconsSmall.ListImages.item(i).Index, _
            imlIconsSmall.ListImages.item(i).Index, 0
        fPropForm.cboIcon(2).ComboItems.Add , TempKey, Icon, _
            imlIconsSmall.ListImages.item(i).Index, _
            imlIconsSmall.ListImages.item(i).Index, 0
        If UseProgressBar Then
            fProgForm.pbPBar1.Value = LastCount + count
        End If
    Next i
    LastCount = LastCount + count
    
    If SHOW_TIMING_INFO Then
        Debug.Print "LoadIcons:  Build 2 small icon combos = " & Timer - t
        t = Timer
    End If
    
    For i = 1 To imlIconsLarge.ListImages.count
        TempKey = imlIconsLarge.ListImages.item(i).key
        Icon = mID$(TempKey, 2, Len(TempKey) - 1)
        fPropForm.cboIcon(3).ComboItems.Add , TempKey, Icon, _
            imlIconsLarge.ListImages.item(i).Index, _
            imlIconsLarge.ListImages.item(i).Index, 0
        If UseProgressBar Then
            fProgForm.pbPBar1.Value = LastCount + count
        End If
    Next i
    
    If SHOW_TIMING_INFO Then
        Debug.Print "LoadIcons:  Build large icon combo = " & Timer - t
    End If
    
    If UseProgressBar Then
        fProgForm.Hide
    End If
    
    Me.MousePointer = vbArrow
    If SHOW_TIMING_INFO Then
        Debug.Print "LoadIcons:  Total Execution Time was " & Timer - t_total _
            & " seconds to execute."
    End If
    Exit Sub

Err_Execute:
    DisplayDBEngineErrors

End Sub
Private Sub Form_Load()
    Dim n As Long
    Dim TempBoolean As Boolean
    Dim TestIndex As Long
    
    InFormLoad = True
    
    'Initialize Globals
    CurrentUser = "Brian and John"
    LoadingTreeView = False
    DragSource = ""
    Set DragItem = Nothing
    RButton = False
    'CurrentDatabaseFile = DEFAULT_DBFILE
    Devil = False
    ListMouseClick = False
    ReadytoPasteItem = False
    ReadytoPasteNode = False
    FindEnabled = False
    IgnoreItemClick = False
    lvEraseBkGnd1 = True
    lvEraseBkGnd2 = True
    lvNextNoErase = False
    lvSaveEraseRect = False
    lvCancel = True                 'assume canceled label edit
    lvCountPaints = False
    SkipCnt = 0
    TrapMultiSelectDrag = False
    isInsertKey = False
    RefreshInProgress = False
    imgSplitter.Left = 3000
    CutANDInPasteNode = False
    TempFileCounter = 0
    UniqueIconKey = ""
    CopyLinkedItem = False
    LoadBlankHTML = False
    LoadBlankSourceViewer = False
    TableComboNeedsRefresh = True
'-----------------------end of global initializations--------------------
    
    'set up the initial size and position of the form
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    
    If Me.Width < 3000 Then
        Me.Width = 3000
    End If
    SizeControls imgSplitter.Left, True
    Select Case lvListView.View
        Case lvwIcon, lvwSmallIcon
            lvListView.Arrange = lvwAutoTop
        Case Else
            lvListView.Arrange = lvwAutoLeft
    End Select
    
    
    ' Initialize the load listview timer
    tmrNodeClick.Enabled = False
    tmrNodeClick.Interval = NODE_CLICK_TIMER_DELAY
    
    
    TempBoolean = True

ChooseTempFolder:
    If GetTempFolder(TempBoolean) = False Then
        If vbYes = MsgBox("You can not run this application without choosing a Temp Folder!  " _
            & "Do you want to select a Temp Folder now ?" & Chr(13) _
            & "NOTE:  Choosing 'NO' will exit the program.", vbCritical + vbYesNo) Then
            TempBoolean = False
            GoTo ChooseTempFolder
        Else
            End
        End If
    End If
    
    TempBoolean = True
    
ChooseLargeIconFolder:
    If GetLargeIconFolder(TempBoolean) = False Then
        If vbYes = MsgBox("You can not run this application without choosing a Large Icon Folder!  " _
            & "Do you want to select a Large Icon Folder now ?" & Chr(13) _
            & "NOTE:  Choosing 'NO' will exit the program.", vbCritical + vbYesNo) Then
            TempBoolean = False
            GoTo ChooseLargeIconFolder
        Else
            End
        End If
    End If
    
    TempBoolean = True
    
ChooseSmallIconFolder:
    If GetsmallIconFolder(TempBoolean) = False Then
        If vbYes = MsgBox("You can not run this application without choosing a Small Icon Folder!  " _
            & "Do you want to select a Small Icon Folder now ?" & Chr(13) _
            & "NOTE:  Choosing 'NO' will exit the program.", vbCritical + vbYesNo) Then
            TempBoolean = False
            GoTo ChooseSmallIconFolder
        Else
            End
        End If
    End If
    
    TempBoolean = True
    
ChooseDatabase:
    If GetCurrentDatabase(TempBoolean) = False Then
        If vbYes = MsgBox("You can not run this application without choosing a Startup Database!  " _
            & "Do you want to select a Startup Database now ?" & Chr(13) _
            & "NOTE:  Choosing 'NO' will exit the program.", vbCritical + vbYesNo) Then
            TempBoolean = False
            GoTo ChooseDatabase
        Else
            End
        End If
    End If
    
    'if we're in debug mode, don't repair or compact the database during
    'load, because then we can't have Access open at the same time.  If
    'this is the case, just open the database.
    If Not DEBUGGING Then
        Set dbase = RepairDB(CurrentDatabaseFile, dbase, True, fProgForm)
        If dbase Is Nothing Then
            Exit Sub
        End If
        Set dbase = CompactDB(CurrentDatabaseFile, dbase, True, fProgForm)
        If dbase Is Nothing Then
            Exit Sub
        End If
    Else
        Set dbase = OpenDBase(CurrentDatabaseFile)
        If dbase Is Nothing Then
            MsgBox "Form_Load:  Couldn't open the database!", vbCritical
            Exit Sub
        End If
    End If

    'sets up the global variable for table names used by the database functions
    InitDBase "Nodes", "Data", "GS_QuickAdd_Items", "GS_LargeIcons", _
                "GS_SmallIcons", "GS_Globals", "GS_Extensions", "GS_Data_Types", _
                "Inbox"
    
   
    'empty out temp files, if necessary
    ClearTempDir CurTempFolder
    
    'load all icons in the appropriate directory
    LoadIcons True
    
    'add custom add item & separator
    Load mnuRCPopupTree1AddNode(1)
    mnuRCPopupTree1AddNode(1).Caption = "Custom"
    mnuRCPopupTree1AddNode(1).tag = FindANode("QuickAdd Custom", False, True)
    mnuRCPopupTree1AddNode(0).Visible = False
    Load mnuRCPopupTree1AddNode(2)
    mnuRCPopupTree1AddNode(2).Caption = "-"
    
    
    'load the node tree and assign icons to it
    tvTreeView.ImageList = Nothing
    'setup temp tree icon image list
    imlTempTree.ListImages.Clear
    imlTempTree.ImageHeight = SMALL_ICON_SIZE
    imlTempTree.ImageWidth = SMALL_ICON_SIZE
    On Error Resume Next
    TestIndex = -1
    TestIndex = imlIconsSmall.ListImages("K" & Link_Icon).Index
    If TestIndex <> -1 Then
        imlTempTree.ListImages.Add , "K" & Link_Icon, _
            imlIconsSmall.ListImages("K" & Link_Icon).picture
    End If
    On Error GoTo 0
    
    On Error Resume Next
    TestIndex = -1
    TestIndex = imlIconsSmall.ListImages("K" & Read_Icon).Index
    If TestIndex <> -1 Then
        imlTempTree.ListImages.Add , "K" & Read_Icon, _
            imlIconsSmall.ListImages("K" & Read_Icon).picture
    End If
    On Error GoTo 0
    
    On Error Resume Next
    TestIndex = -1
    TestIndex = imlIconsSmall.ListImages("K" & System_Icon).Index
    If TestIndex <> -1 Then
        imlTempTree.ListImages.Add , "K" & System_Icon, _
            imlIconsSmall.ListImages("K" & System_Icon).picture
    End If
    On Error GoTo 0
    
    If imlTempTree.ListImages.count <> 0 Then
        tvTreeView.ImageList = imlTempTree
    End If
    
    LoadTreeView True
    'SetTVIcons
    
    
    'remove image list bindings for the listview
    lvListView.Icons = Nothing
    lvListView.SmallIcons = Nothing
    'clear out the temporary image list for large icons
    imlTempLarge.ListImages.Clear
    'clear out the temporary image list for small icons
    imlTempSmall.ListImages.Clear
    
    'define the small icon size
    'if these lines are missing, the image list defaults to 32x32,
    'regardless of what the properties settings are.
    imlTempSmall.ImageHeight = SMALL_ICON_SIZE
    imlTempSmall.ImageWidth = SMALL_ICON_SIZE
    
    'add the overlay images for the different attributes to the image list
    On Error Resume Next
    TestIndex = -1
    TestIndex = imlIconsLarge.ListImages("K" & Link_Icon).Index
    If TestIndex <> -1 Then
        imlTempLarge.ListImages.Add , "K" & Link_Icon, _
            imlIconsLarge.ListImages("K" & Link_Icon).picture
    End If
    On Error GoTo 0
    
    On Error Resume Next
    TestIndex = -1
    TestIndex = imlIconsLarge.ListImages("K" & Read_Icon).Index
    If TestIndex <> -1 Then
        imlTempLarge.ListImages.Add , "K" & Read_Icon, _
            imlIconsLarge.ListImages("K" & Read_Icon).picture
    End If
    On Error GoTo 0
    
    On Error Resume Next
    TestIndex = -1
    TestIndex = imlIconsLarge.ListImages("K" & System_Icon).Index
    If TestIndex <> -1 Then
        imlTempLarge.ListImages.Add , "K" & System_Icon, _
            imlIconsLarge.ListImages("K" & System_Icon).picture
    End If
    On Error GoTo 0
    
    'add the overlay images for the different attributes to the image list
    On Error Resume Next
    TestIndex = -1
    TestIndex = imlIconsSmall.ListImages("K" & Link_Icon).Index
    If TestIndex <> -1 Then
        imlTempSmall.ListImages.Add , "K" & Link_Icon, _
            imlIconsSmall.ListImages("K" & Link_Icon).picture
    End If
    On Error GoTo 0
    
    On Error Resume Next
    TestIndex = -1
    TestIndex = imlIconsSmall.ListImages("K" & Read_Icon).Index
    If TestIndex <> -1 Then
        imlTempSmall.ListImages.Add , "K" & Read_Icon, _
            imlIconsSmall.ListImages("K" & Read_Icon).picture
    End If
    On Error GoTo 0
    
    On Error Resume Next
    TestIndex = -1
    TestIndex = imlIconsSmall.ListImages("K" & System_Icon).Index
    If TestIndex <> -1 Then
        imlTempSmall.ListImages.Add , "K" & System_Icon, _
            imlIconsSmall.ListImages("K" & System_Icon).picture
    End If
    On Error GoTo 0
    
    If imlTempLarge.ListImages.count <> 0 Then
        lvListView.Icons = imlTempLarge
    End If
    
    If imlTempSmall.ListImages.count <> 0 Then
        lvListView.SmallIcons = imlTempSmall
    End If
    
    
    'build and create the owner-drawn menus
    BuildimlMenu True
    InitOwnerDrawMenus True
    
    'setup the toolbar
    InitToolbar True
    
    'initialize subclassing (this function checks to see if debugging is True
    'or not.  If it is, then subclassing is not initialized.
    InitSubclassing True
    
    Me.Show

    'initialize both views, load the list view, and update the status bar
    InitTreeView
    InitListView
    LoadListView nodx.key
    UpdateStatusBar 0, True
    
    QuickAddNodesNodeID = FindANode("QuickAdd Nodes", False, True)
    
    tbEdit.Buttons("Paste").Enabled = False
    mnuEditPaste.Enabled = False
    mnuRCPopupList2Paste.Enabled = False
    mnuRCPopupTree1Paste.Enabled = False
        
    'check to see if we should show the quote of the day
    If GetSetting(App.EXEName, "Options", "Show Tips at Startup", 1) Then
        frmQuotes.Show vbModal, Me
    End If
    
    InFormLoad = False
End Sub
Private Sub Form_Terminate()
    Dim i As Integer
    
    'if debugging mode is not one (which means subclassing is active) then
    'unhook all instances.
    If Not DEBUGGING Then
        For i = MIN_INSTANCES To MAX_INSTANCES
            If Instances(i).in_use Then
                UnHookWindow (i)
                Instances(i).in_use = False
            End If
        Next i
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Dim i As Integer
    
    'clean out any files in the temp directory
    ClearTempDir CurTempFolder
        
    'unload the forms
    Unload fProgForm
    Unload fPropForm
    Unload fReplaceForm
    For i = Forms.count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    
    'save window position and size
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    'save the last view mode for the listview
    SaveSetting App.Title, "Settings", "ViewMode", lvListView.View
    
    'close the database
    dbase.Close
End Sub








Private Sub mnuNodesRename_Click()
    mnuRCPopupTree1Rename_Click
End Sub

'****************************************************************************
'*********************         UNSORTED FUNCTIONS          ******************
'****************************************************************************



Private Sub mnuRCPopupList1Copy_Click()
    Dim i As Integer
    Dim buffer_index As Long
    Dim TempBufferItem As lvBufferItem
    Dim record As Recordset
    Dim tempstr As String
    
    ReadytoPasteItem = False
    ReadytoPasteNode = False
    ItemWasCut = False
    NodeWasCut = False

    For i = 1 To ItemBuffer.count
        ItemBuffer.Remove "K" & i
    Next i
    
    buffer_index = 0
    
    For i = 1 To lvListView.ListItems.count
        Set TempBufferItem = Nothing
        Set TempBufferItem = New lvBufferItem
        If lvListView.ListItems.item(i).Selected = True Then
            
            buffer_index = buffer_index + 1
            TempBufferItem.text = lvListView.ListItems.item(i).text
            TempBufferItem.key = lvListView.ListItems.item(i).key
            TempBufferItem.parent_node = lvAttribCol(lvListView.ListItems.item(i).key).parent_node
            TempBufferItem.icon_large = lvAttribCol(lvListView.ListItems.item(i).key).icon_large
            TempBufferItem.icon_small = lvAttribCol(lvListView.ListItems.item(i).key).icon_small
            TempBufferItem.read_only = lvAttribCol(lvListView.ListItems.item(i).key).read_only
            TempBufferItem.data_type = lvAttribCol(lvListView.ListItems.item(i).key).data_type
            TempBufferItem.created = lvAttribCol(lvListView.ListItems.item(i).key).created
            TempBufferItem.created_by = lvAttribCol(lvListView.ListItems.item(i).key).created_by
            TempBufferItem.last_modified = Now
            TempBufferItem.modified_by = CurrentUser
            TempBufferItem.variation = lvAttribCol(lvListView.ListItems.item(i).key).variation
            ItemBuffer.Add TempBufferItem, "K" & Format$(buffer_index)
                
            If CopyLinkedItem Then
                Set record = dbase.OpenRecordset("SELECT data_value, " _
                    & "binary_data_value" & " FROM " & _
                    tvAttribCol(AddK(tvAttribCol(nodx.key).Link_NodeID)).table_name _
                    & " WHERE Data_ID = " _
                    & RemoveK(lvListView.ListItems.item(i).key), dbOpenDynaset)
            Else
                Set record = dbase.OpenRecordset("SELECT data_value, " _
                    & "binary_data_value" & " FROM " & _
                    tvAttribCol(nodx.key).table_name _
                    & " WHERE Data_ID = " _
                    & RemoveK(lvListView.ListItems.item(i).key), dbOpenDynaset)
            End If
            
            WriteMemo record, "data_value", CurTempFolder & "CopyMemo" & Format$(buffer_index)
            If Not ValidFile(CurTempFolder & "CopyMemo" & Format$(buffer_index)) Then
                MsgBox "Couldn't create necessary temp file, '" _
                    & CurTempFolder & "CopyMemo" & Format$(buffer_index) _
                    & "' while attempting to copy [Data_Value] field of item '" _
                    & lvListView.ListItems.item(i).text & "'" _
                    & Chr$(13) & "This item was not copied successfully.", vbCritical
            End If
            
            WriteBLOB record, "binary_data_value", CurTempFolder & "CopyBLOB" & Format$(buffer_index)
            If Not ValidFile(CurTempFolder & "CopyBLOB" & Format$(buffer_index)) Then
                MsgBox "Couldn't create necessary temp file, '" _
                    & CurTempFolder & "CopyBLOB" & Format$(buffer_index) _
                    & "' while attempting to copy [Binary_Data_Value] field of item '" _
                    & lvListView.ListItems.item(i).text & "'" _
                    & Chr$(13) & "This item was not copied successfully.", vbCritical
            End If
            record.Close
TryNextOne:
        End If
    Next i
    
    NumListItemsCopied = ItemBuffer.count
    
    If NumListItemsCopied > 0 Then
        ReadytoPasteItem = True
    End If
End Sub

Private Sub mnuRCPopupList1Cut_Click()
    Dim i As Long
    Dim tempstr As String
    
    tempstr = ""
    If tvAttribCol(nodx.key).Link_NodeID <> 0 And (Not ClickedVariationMenu) Then
        tempstr = "Linked (TRUE)"
    End If
    
    If tvAttribCol(nodx.key).read_only Then
        If tempstr = "" Then
            tempstr = "Read Only (TRUE)"
        Else
            tempstr = tempstr & ", Read Only (TRUE)"
        End If
    End If
    
    If tempstr <> "" Then
        MsgBox "Items cannot be cut from '" & nodx.text & "' because it has the " _
            & "following attribute(s) set:  " & tempstr & ".", vbExclamation
        Exit Sub
    End If
    
    If lvNumListItemsSelected = 1 Then
        tempstr = ""
        If lvAttribCol(CurrentItem.key).read_only Then
            tempstr = "Read Only (TRUE)"
        End If
        
        If tempstr <> "" Then
            MsgBox "'" & CurrentItem.text & "' cannot be cut because the following " _
                & "attribute(s) are set:  " & tempstr & ".", vbExclamation
            Exit Sub
        End If
        
        If Not CurrentItem.Ghosted Then CurrentItem.Ghosted = True
    Else
        tempstr = ""
        For i = 1 To lvListView.ListItems.count
            If lvListView.ListItems.item(i).Selected = True Then
                If lvAttribCol(lvListView.ListItems.item(i).key).read_only Then
                    tempstr = "Read Only (TRUE)"
                End If
                
                If tempstr <> "" Then
                    MsgBox "'" & lvListView.ListItems.item(i).text & "' cannot be cut because the following " _
                        & "attribute(s) are set:  " & tempstr & ".", vbExclamation
                    Exit Sub
                End If
            End If
        Next i
        
        For i = 1 To lvListView.ListItems.count
            If lvListView.ListItems.item(i).Selected = True Then
                
                If Not lvListView.ListItems.item(i).Ghosted Then
                    lvListView.ListItems.item(i).Ghosted = True
                End If
                
            End If
        Next i
    End If
    
    mnuRCPopupList1Copy_Click
    NodeKeyCutFrom = nodx.key
    ItemWasCut = True
End Sub

Private Sub mnuRCPopupList1Delete_Click()
    DeleteListItems True
End Sub

Private Sub DeleteListItems(ConfirmMsgBoxes As Boolean)

    Dim i As Integer
    Dim number_items As Long
    Dim MsgReturnCode As Long
    Dim tempstr As String
    
    tempstr = ""
    If tvAttribCol(nodx.key).Link_NodeID <> 0 And (Not ClickedVariationMenu) Then
        tempstr = "Linked (TRUE)"
    End If
    
    If tvAttribCol(nodx.key).read_only Then
        If tempstr = "" Then
            tempstr = "Read Only (TRUE)"
        Else
            tempstr = tempstr & ", Read Only (TRUE)"
        End If
    End If
    
    If tempstr <> "" Then
        MsgBox "Items cannot be deleted from '" & nodx.text & "' because it has the " _
            & "following attribute(s) set:  " & tempstr & ".", vbExclamation
        Exit Sub
    End If
    
    number_items = lvListView.ListItems.count
   
    NodeKeyCutFrom = nodx.key
    If lvNumListItemsSelected > 1 Then
        If ConfirmMsgBoxes Then
            MsgReturnCode = MsgBox("Are you sure you want to delete " _
                    & Format$(lvNumListItemsSelected) & _
                    " list items?", vbYesNo + vbExclamation, "Delete Nodes")
            If MsgReturnCode = vbNo Then
                GoTo Done
            End If
        End If
        For i = number_items To 1 Step -1
            If lvListView.ListItems.item(i).Selected = True Then
                DeleteListItem lvListView.ListItems.item(i).key, _
                            nodx.key, False
            End If
        Next i
    ElseIf lvNumListItemsSelected = 1 Then
            tempstr = ""
            If lvAttribCol(CurrentItem.key).read_only Then
                tempstr = "Read Only (TRUE)"
            End If
            
            If tempstr <> "" Then
                MsgBox "'" & CurrentItem.text & "' cannot be deleted because the following " _
                    & "attribute(s) are set:  " & tempstr & ".", vbExclamation
                GoTo Done
            End If
        If ConfirmMsgBoxes Then
            MsgReturnCode = MsgBox("Are you sure you want to delete " & Chr$(34) & _
                CurrentItem.text & Chr$(34) & "?", vbYesNo + vbExclamation, _
                "Delete Node")
            If MsgReturnCode = vbNo Then
                GoTo Done
            End If
        End If
        DeleteListItem CurrentItem.key, nodx.key, True
    End If
    
Done:
End Sub

Private Sub mnuRCPopupList1Edit_Click()
    Select Case lvAttribCol(CurrentItem.key).data_type
        Case "HTML", "String"
            MsgBox "TEMP: Load HTML Editor", vbInformation
            'Load frmHTMLEdit
        Case Else
            MsgBox "The editor does not support items of type '" _
                & lvAttribCol(CurrentItem.key).data_type & "'!", vbExclamation
    End Select
End Sub

Private Sub mnuRCPopupList1Execute_Click()
    Dim record As Recordset
    Dim data_size As Long
    Dim TempFileName As String
    Dim hProcess As Long
    Dim tempstr As String
    Dim ShowWarning As Integer
    
    Me.MousePointer = vbHourglass
    
    If Left$(CurrentItem.key, 1) = "L" Then
        tempstr = "NOT ALLOWED"
    Else
        tempstr = lvAttribCol(CurrentItem.key).data_type
        If UCase(Left$(tempstr, 6)) = "BINARY" Then
            tempstr = "Binary"
        End If
    End If
   
    Select Case tempstr
        Case "String", "HTML"
            Set record = dbase.OpenRecordset("SELECT data_value,binary_data_value FROM " _
                & tvAttribCol(nodx.key).table_name & " WHERE Data_ID = " _
                & RemoveK(CurrentItem.key), dbOpenDynaset)
            data_size = record("data_value").FieldSize
            If data_size <= 0 Then
                MsgBox "There is no data attached to this item.", vbExclamation
                Exit Sub
            End If
            
            ShowWarning = GetSetting(App.EXEName, "Options", _
                    "Show External App Warning", 1)
            If ShowWarning <> 0 Then
                Me.MousePointer = vbArrow
                frmExternalApp.Show vbModal, Me
                Me.MousePointer = vbHourglass
            End If
        
            TempFileName = CurTempFolder & tempstr & "Exec." & tempstr
            WriteMemo record, "data_value", TempFileName

            If Not ValidFile(TempFileName) Then
                MsgBox "The necessary temp file '" & TempFileName _
                    & "' was not created successfully.  The item '" _
                    & CurrentItem.text & "' could not be executed.", vbCritical
                Exit Sub
            End If
            
            Me.MousePointer = vbArrow
            hProcess = ShellExecute(0&, "Open", TempFileName, "", 0&, SW_NORMAL)
            record.Close
        Case "Binary"
            Set record = dbase.OpenRecordset("SELECT data_value,binary_data_value FROM " _
                & tvAttribCol(nodx.key).table_name & " WHERE Data_ID = " _
                & RemoveK(CurrentItem.key), dbOpenDynaset)
            data_size = record("binary_data_value").FieldSize
            If data_size <= 0 Then
                MsgBox "There is no data attached to this item.", vbExclamation
                Me.MousePointer = vbArrow
                Exit Sub
            End If
            
            ShowWarning = GetSetting(App.EXEName, "Options", _
                    "Show External App Warning", 1)
            If ShowWarning <> 0 Then
                Me.MousePointer = vbArrow
                frmExternalApp.Show vbModal, Me
                Me.MousePointer = vbHourglass
            End If
            
            'strip out "<BINARY> " at start of field
            TempFileName = mID$(record!data_value, 10, Len(record!data_value) - 9)
            
            'strip out path info, in case its still there
            TempFileName = GetFileNameFromPath(TempFileName)
            WriteBLOB record, "binary_data_value", CurTempFolder & TempFileName

            If Not ValidFile(TempFileName) Then
                MsgBox "The necessary temp file '" & TempFileName _
                    & "' was not created successfully.  The item '" _
                    & CurrentItem.text & "' could not be executed.", vbCritical
                Exit Sub
            End If
            
            Me.MousePointer = vbArrow
            hProcess = ShellExecute(0&, "Open", TempFileName, "", 0&, SW_NORMAL)
            record.Close
        Case "URL"
            Set record = dbase.OpenRecordset("SELECT data_value FROM " _
                & tvAttribCol(nodx.key).table_name & " WHERE Data_ID = " _
                & RemoveK(CurrentItem.key), dbOpenDynaset)
            data_size = record("data_value").FieldSize
            If data_size <= 0 Then
                MsgBox "There is no data attached to this item.", vbExclamation
                Exit Sub
            End If
            
            ShowWarning = GetSetting(App.EXEName, "Options", _
                    "Show External App Warning", 1)
            If ShowWarning <> 0 Then
                Me.MousePointer = vbArrow
                frmExternalApp.Show vbModal, Me
                Me.MousePointer = vbHourglass
            End If
        
            Me.MousePointer = vbArrow
            hProcess = ShellExecute(0&, "Open", record!data_value, "", 0&, SW_NORMAL)
            record.Close
        Case Else
            Me.MousePointer = vbArrow
            MsgBox "Items of type '" & tempstr & "' may not be executed!", _
                vbExclamation
    End Select
End Sub

Private Sub mnuRCPopupList1Find_Click()
    mnuEditFind_Click
End Sub

Private Sub mnuRCPopupList1Print_Click()
    mnuFilePrint_Click
End Sub

Private Sub mnuRCPopupList1Properties_Click()
    mnuItemsProperties_Click
End Sub

Private Sub mnuRCPopupList1Rename_Click()
    Dim tempstr As String
    
    tempstr = ""
    If lvAttribCol(CurrentItem.key).read_only Then
        tempstr = "Read Only (TRUE)"
    End If
    
    If tvAttribCol(nodx.key).read_only Then
        If tempstr = "" Then
            tempstr = "Read Only Parent (TRUE)"
        Else
            tempstr = tempstr & ", Read Only Parent (TRUE)"
        End If
    End If
    
   If tvAttribCol(nodx.key).Link_NodeID <> 0 Then
        If tempstr = "" Then
            tempstr = "Linked (TRUE)"
        Else
            tempstr = tempstr & ", Linked (TRUE)"
        End If
    End If
    
    If tempstr <> "" Then
        MsgBox "'" & CurrentItem.text & "' cannot be renamed because the following " _
            & "attribute(s) are set:  " & tempstr & ".", vbExclamation
        Exit Sub
    End If

    lvListView.StartLabelEdit
End Sub

Private Sub mnuRCPopupList1Replace_Click()
    mnuEditReplace_Click
End Sub

Private Sub mnuRCPopupList1Variation_Click()
    Dim OneIsVariation As Boolean
    Dim OneIsNotVariation As Boolean
    Dim i As Integer
    Dim NameofSelectedItem As String
    Dim tempstr As String
    
    tempstr = ""
    If tvAttribCol(nodx.key).read_only Then
        tempstr = "Read Only"
    End If
    
    If tempstr <> "" Then
        MsgBox "The variation attribute for the item(s) cannot be altered " _
            & "because the parent node '" & nodx.text & "' has the following " _
            & "attribute(s) set:  " & tempstr & ".", vbExclamation
        GoTo Done
    End If
    
    For i = 1 To lvListView.ListItems.count
        If lvListView.ListItems(i).Selected = True Then
            If lvAttribCol(lvListView.ListItems(i).key).variation Then
                OneIsVariation = True
            Else
                OneIsNotVariation = True
            End If
        End If
    Next i
    If OneIsVariation And OneIsNotVariation Then
        MsgBox "The variation attribute for the item(s) cannot be altered " _
            & "because the selection includes both variations and non-variations.", _
            vbExclamation
        GoTo Done
    End If
    
    If mnuRCPopupList1Variation.Caption = "Create Variation" Then
        LockWindowUpdate lvListView.hwnd
        CopyLinkedItem = True
        mnuRCPopupList1Copy_Click
        
        FillSelectedItemsBuffer
        
        ClickedVariationMenu = True
        mnuRCPopupList2Paste_Click
        ClickedVariationMenu = False
        
        For i = 1 To NumListItemsCopied
            lvListView.ListItems.Remove ItemBuffer("K" & i).key
            lvAttribCol.Remove ItemBuffer("K" & i).key
        Next i
        
        For i = 1 To SelectedItemsBuffer.count
            Set CurrentItem = lvListView.FindItem(SelectedItemsBuffer("K" & i))
            If Not CurrentItem Is Nothing Then
                CurrentItem.Selected = True
            End If
        Next i
        
        CopyLinkedItem = False
        mnuRCPopupList1Variation.Caption = "Remove Variation"
        mnuItemsVariation.Caption = "Remove Variation"
        tbEdit.Buttons("Variation").ToolTipText = "Remove Variation"
        LockWindowUpdate (0&)
    Else        'this is the remove variation case
        If lvNumListItemsSelected > 1 Then
            NameofSelectedItem = "these " & Format$(lvNumListItemsSelected) _
                & " items"
        Else
            NameofSelectedItem = "'" & CurrentItem.text & "'"
        End If
        If vbYes = MsgBox("Are you sure you want to remove the variation made on " _
                & NameofSelectedItem & "?" & Chr(13) _
                & "Choosing Yes will revert the item(s) back to the global value.", _
                vbYesNo + vbExclamation, "Remove Variation") Then
            LockWindowUpdate lvListView.hwnd
            
            FillSelectedItemsBuffer
            
            ClickedVariationMenu = True
            DeleteListItems False
            ClickedVariationMenu = False
            
            LoadListView nodx.key, True
            
            mnuRCPopupList1Variation.Caption = "Create Variation"
            mnuItemsVariation.Caption = "Create Variation"
            tbEdit.Buttons("Variation").ToolTipText = "Create Variation"
            LockWindowUpdate (0&)
        End If
    End If
    
Done:
End Sub

Private Sub mnuRCPopupList2ADataID_Click()
    mnuVAIByDataID_Click
End Sub

Private Function AddItem(Optional icon_large As String = DEFAULT_LARGE_ICON, _
        Optional icon_small As String = DEFAULT_SMALL_ICON, _
        Optional ItemName As String = DEFAULT_ITEM_NAME, _
        Optional RenameItem As Boolean = True, _
        Optional DataType As String = "String", _
        Optional created As Variant = 0, _
        Optional created_by As String = "Osiris", _
        Optional last_modified As Variant = 0, _
        Optional modified_by As String = "Osiris", _
        Optional data_value As String = "", _
        Optional AddL As Boolean = False, _
        Optional TempImageList As Boolean = False, _
        Optional read_only As Boolean = False)
    Dim record As Recordset
    Dim listrecord As Recordset
    Dim FreeNumber As Long
    Dim tempstr As String
    Dim tmplvitem As New lvitem
    Dim variation As Boolean
    Dim order As Long
    Dim CurrentItemKey As String
    
    AddItem = 0
    
    If ClickedVariationMenu Then
        variation = True
    Else
        variation = False
    End If
    
    Set record = dbase.OpenRecordset("SELECT [Order] FROM " & DB_NodeTable _
        & " WHERE Node_ID = " & RemoveK(nodx.key), dbOpenDynaset)
    If record.EOF Then
        MsgBox "Error: parent node not found in database.", vbCritical
        Exit Function
    End If
    order = record!order
    record.Close
        
    If tvAttribCol(nodx.key).quicktypeid = 0 Then
        tempstr = DB_QuickAddItemsTable
    Else
        tempstr = tvAttribCol(nodx.key).table_name
    End If
    
    FreeNumber = FindFreeID(dbase, tempstr, "Data_ID")
    
    RectHeight = 0
    RectWidth = 0
    If RenameItem Then
        lvNoEraseCnt = 0
        lvEraseBkGnd1 = False
        lvEraseBkGnd2 = False
        lvSaveEraseRect = True
        lvCancel = True                 'assume canceled label edit
    End If

    CurrentItemKey = AddK(FreeNumber, AddL)
    
    tmplvitem.parent_node = RemoveK(nodx.key)
    tmplvitem.read_only = read_only
    tmplvitem.data_type = DataType
    tmplvitem.icon_large = icon_large
    tmplvitem.icon_small = icon_small
    tmplvitem.created = created
    tmplvitem.created_by = created_by
    tmplvitem.last_modified = last_modified
    tmplvitem.modified_by = modified_by
    tmplvitem.variation = variation
    
    lvAttribCol.Add tmplvitem, CurrentItemKey
    
    Set CurrentItem = lvListView.ListItems.Add(, CurrentItemKey, _
            ItemName, CreateOverlayedImage(imlIconsLarge, "K" & _
            icon_large, imlTempLarge, AddL, _
            read_only, tvAttribCol(nodx.key).system_node), _
            CreateOverlayedImage(imlIconsSmall, _
            "K" & icon_small, imlTempSmall, AddL, _
            read_only, tvAttribCol(nodx.key).system_node))

    lvNumListItemsSelected = lvNumListItemsSelected + 1

    UpdateStatusBar lvNumListItemsSelected

    If lvListView.View = lvwReport Then
        CurrentItem.SubItems(1) = FreeNumber
        CurrentItem.SubItems(2) = RemoveK(nodx.key)
        CurrentItem.SubItems(3) = icon_large
        CurrentItem.SubItems(4) = icon_small
        CurrentItem.SubItems(5) = DataType
    End If
        
    Set listrecord = dbase.OpenRecordset(tempstr, dbOpenTable)
        
    listrecord.AddNew
    listrecord!data_id = FreeNumber
    listrecord!data_label = ItemName
    listrecord!parent_node = RemoveK(nodx.key)
    listrecord!read_only = read_only
    listrecord!icon_large = icon_large
    listrecord!icon_small = icon_small
    listrecord!data_type = DataType
    listrecord!data_value = data_value
    listrecord!binary_data_value = Nothing
    listrecord!last_modified = last_modified
    listrecord!modified_by = modified_by
    listrecord!created = created
    listrecord!created_by = created_by
    listrecord!variation = variation
    listrecord!order = order
    listrecord.Update
    listrecord.Close
    
    lvListView.MultiSelect = False
    lvListView.SelectedItem = CurrentItem
    lvListView.MultiSelect = True
    lvSomethingSelected = True
    If RenameItem Then
        tvTreeView.SetFocus      'this is here because microsoft is the devil
        Devil = True
        lvListView.SetFocus      'remove at your own risk
    End If
    
    AddItem = FreeNumber
End Function

Private Sub mnuRCPopupList2AddItem_Click(Index As Integer)
    Dim tag As String
    Dim lastpos As Long
    Dim TempPos As Long
    Dim tempstr As String
    Dim TempCaption As String
    Dim icon_large As String
    Dim icon_small As String
    Dim data_type As String
    Dim i As Long
    Dim rename_item As Boolean
    
    ' if the custom node
    If RemoveK(nodx.key) = mnuRCPopupTree1AddNode(1).tag Then
        MsgBox "Items cannot be added under the Custom QuickAdd node."
        GoTo Done
    End If
    
    'if a child node of a quickadd node
    If Not nodx.parent Is Nothing Then
        If Not nodx.parent.parent Is Nothing Then
            If RemoveK(nodx.parent.parent.key) = QuickAddNodesNodeID Then
                MsgBox "Items for this QuickAdd node must be added " _
                    & "under the master copy: " & _
                    tvTreeView.Nodes(AddK( _
                    FindQuickTypeRootNode(nodx.Index))).FullPath, vbExclamation
                GoTo Done
            End If
        End If
    End If
    
    tempstr = ""
    If tvAttribCol(nodx.key).Link_NodeID <> 0 Then
        tempstr = "Linked (TRUE)"
    End If
    
    If tvAttribCol(nodx.key).read_only Then
        If tempstr = "" Then
            tempstr = "Read Only (TRUE)"
        Else
            tempstr = tempstr & ", Read Only (TRUE)"
        End If
    End If
    
    If Not tvAttribCol(nodx.key).create_item Then
        If tempstr = "" Then
            tempstr = "Create Item (FALSE)"
        Else
            tempstr = tempstr & ", Create Item (FALSE)"
        End If
    End If
    
    If tempstr <> "" Then
        MsgBox "'" & nodx.text & "' cannot have items added under it because the " _
            & "following attribute(s) are set:  " & tempstr & ".", vbExclamation
        Exit Sub
    End If
    
    tag = mnuRCPopupList2AddItem(Index).tag
    TempPos = InStr(1, tag, ",", 1)
    If TempPos = 1 Then
        icon_large = ""
    Else
        icon_large = mID$(tag, 1, TempPos - 1)
    End If
    lastpos = TempPos
        
    TempPos = InStr(lastpos + 1, tag, ",", 1)
    If TempPos = (lastpos + 1) Then
        icon_small = ""
    Else
        icon_small = mID$(tag, lastpos + 1, TempPos - lastpos - 1)
    End If
    lastpos = TempPos
        
    tempstr = mnuRCPopupList2AddItem(Index).Caption
    TempCaption = tempstr
    data_type = mID$(TempCaption, 8, Len(TempCaption) - 7)
    
    i = 1
    While Not lvListView.FindItem(tempstr) Is Nothing
        Select Case i
            Case Is >= 10000
                MsgBox "You have too many numbered items!  " _
                    & "Please rename some before adding this item.", vbExclamation
                Exit Sub
            Case Is >= 1000     ' 4 digit
                tempstr = Left$(tempstr, Len(tempstr) - 7)
            Case Is >= 100      ' 3 digit
                i = i + 1
                tempstr = Left$(tempstr, Len(tempstr) - 6)

            Case Is >= 10       '2 digit
                i = i + 1
                tempstr = Left$(tempstr, Len(tempstr) - 5)
            Case Else           '1 digit
                If i = 1 Then
                    i = 2
                Else
                    i = i + 1
                    tempstr = Left$(tempstr, Len(tempstr) - 4)
                End If
        End Select
        tempstr = tempstr & " (" & Format$(i) & ")"
    Wend
    
    If Index <= number_of_custom_items Then 'custom item
        ' if they're adding a new data type, it will always be a string.
        If RemoveK(nodx.key) = FindANode("Data Types", False, True) Then
            data_type = "String"
        End If
        rename_item = True
    Else    'Quickadd item
        rename_item = False
    End If
    Call AddItem(icon_large, icon_small, tempstr, _
            rename_item, data_type, Now, CurrentUser, Now, CurrentUser)
            
Done:
End Sub

Private Sub mnuRCPopupList2AItem_Click()
    mnuVAIByItem_Click
End Sub

Private Sub mnuRCPopupList2AParent_Click()
    mnuVAIByParentNode_Click
End Sub

Private Sub mnuRCPopupList2AType_Click()
    mnuVAIByType_Click
End Sub

Private Sub mnuRCPopupList2Paste_Click()
    Dim i As Long
    Dim j As Long
    Dim tempstr As String
    Dim tempstr2 As String
    Dim listrecord As Recordset
    Dim FreeNumber As Long
    Dim lvSrcRecord As Recordset
    Dim created As Variant
    Dim created_by As String
    Dim data_id As Long
    Dim record As Recordset
    
    tempstr = ""
    If ItemWasCut And NodeKeyCutFrom = nodx.key Then
        tempstr = "Source and Target Node are Same"
    End If
    
    If tvAttribCol(nodx.key).Link_NodeID <> 0 And (Not ClickedVariationMenu) Then
        If tempstr = "" Then
            tempstr = "Linked (TRUE)"
        Else
            tempstr = tempstr & ", Linked (TRUE)"
        End If
    End If
    
    If tvAttribCol(nodx.key).read_only Then
        If tempstr = "" Then
            tempstr = "Read Only (TRUE)"
        Else
            tempstr = tempstr & ", Read Only (TRUE)"
        End If
    End If
    
    If Not tvAttribCol(nodx.key).create_item And (Not ClickedVariationMenu) Then
        If tempstr = "" Then
            tempstr = "Create Item (False)"
        Else
            tempstr = tempstr & ", Create Item (FALSE)"
        End If
    End If
    
    If tempstr <> "" Then
        MsgBox "'" & nodx.text & "' cannot have anything pasted into it because" _
            & " the following attribute(s) are set:  " & tempstr & ".", vbExclamation
        Exit Sub
    End If
    
    For i = 1 To NumListItemsCopied
        tempstr = ItemBuffer("K" & Format$(i)).text
        If Not CopyLinkedItem And Not lvListView.FindItem(tempstr) Is Nothing Then
            j = 1
            While Not lvListView.FindItem(tempstr) Is Nothing
                Select Case j
                    Case Is >= 10000
                        MsgBox "You have too many numbered items!  " _
                            & "Please rename some before pasting this item.", vbExclamation
                        Exit Sub
                    Case Is >= 1000     ' 4 digit
                        tempstr = Left$(tempstr, Len(tempstr) - 7)
                    Case Is >= 100      ' 3 digit
                        j = j + 1
                        tempstr = Left$(tempstr, Len(tempstr) - 6)
    
                    Case Is >= 10       '2 digit
                        j = j + 1
                        tempstr = Left$(tempstr, Len(tempstr) - 5)
                    Case Else           '1 digit
                        If j = 1 Then
                            j = 2
                        Else
                            j = j + 1
                            tempstr = Left$(tempstr, Len(tempstr) - 4)
                        End If
                End Select
                tempstr = tempstr & " (" & Format$(j) & ")"
            Wend
            ItemBuffer("K" & Format$(i)).text = tempstr
        End If
        
        tempstr = tvAttribCol(nodx.key).table_name
        FreeNumber = FindFreeID(dbase, tempstr, "Data_ID")
        If Not DragDropped Then
            If ItemWasCut Then
                created = ItemBuffer("K" & i).created
                created_by = ItemBuffer("K" & i).created_by
            Else
                created = Now
                created_by = CurrentUser
            End If
            If tvAttribCol(nodx.key).Link_NodeID <> 0 Then
                data_id = AddItem(ItemBuffer("K" & i).icon_large, _
                    ItemBuffer("K" & i).icon_small, ItemBuffer("K" & i).text, _
                    False, ItemBuffer("K" & i).data_type, created, created_by, _
                    Now, CurrentUser, , False, True, ItemBuffer("K" & i).read_only)
            Else
                data_id = AddItem(ItemBuffer("K" & i).icon_large, _
                    ItemBuffer("K" & i).icon_small, ItemBuffer("K" & i).text, _
                    False, ItemBuffer("K" & i).data_type, created, created_by, _
                    Now, CurrentUser, , , , ItemBuffer("K" & i).read_only)
            End If
            Set record = dbase.OpenRecordset("SELECT data_value, " _
                & "binary_data_value" & " FROM " & tvAttribCol(nodx.key).table_name _
                & " WHERE Data_ID = " _
                & data_id, dbOpenDynaset)
            
            'the ReadBLOB function sets the [Binary_Data_Value] field,
            'along with the filename in the [Data_Value] field, which in this
            'case is just the name of the temp file.  Since we don't want that,
            'but want the original filename instead, we must call this BEFORE
            'ReadMemo.  ReadMemo will copy the original filename into the
            '[Data_Value] field.  If ReadMemo is called before ReadBLOB the
            'original filename will be lost!!!
            
            If Not ValidFile(CurTempFolder & "CopyBLOB" & Format$(i)) Then
                MsgBox "The necessary temp file '" _
                    & CurTempFolder & "CopyBLOB" & Format$(i) _
                    & "' was not found." _
                    & Chr$(13) & "'" & ItemBuffer("K" & i).text _
                    & "' was not pasted successfully.", vbCritical
                GoTo NoBLOB
            End If
            ReadBLOB CurTempFolder & "CopyBLOB" & Format$(i), record, "binary_data_value"
            
NoBLOB:
            
            'see comment above for the ReadBLOB function!!!!!!!!!!
            If Not ValidFile(CurTempFolder & "CopyMemo" & Format$(i)) Then
                MsgBox "The necessary temp file '" _
                    & CurTempFolder & "CopyMemo" & Format$(i) _
                    & "' was not found." _
                    & Chr$(13) & "'" & ItemBuffer("K" & i).text _
                    & "' was not pasted successfully.", vbCritical
                GoTo NoMemo
            End If
            ReadMemo CurTempFolder & "CopyMemo" & Format$(i), record, "data_value"
NoMemo:
            record.Close
        Else
            If i = lvNumListItemsSelected Then
                DragDropped = False
            End If
        End If
        If lvListView.View = lvwReport Then
            CurrentItem.SubItems(1) = FreeNumber
            CurrentItem.SubItems(2) = RemoveK(nodx.key)
            CurrentItem.SubItems(3) = ItemBuffer("K" & i).icon_large
            CurrentItem.SubItems(4) = ItemBuffer("K" & i).icon_small
            CurrentItem.SubItems(5) = ItemBuffer("K" & i).data_type
        End If
NextItem:
    Next i
    lvListView.MultiSelect = False
    If Not CurrentItem Is Nothing Then
        lvListView.SelectedItem = CurrentItem
    End If
    lvListView.MultiSelect = True
    lvSomethingSelected = True
    
    If ItemWasCut Then
        For i = 1 To NumListItemsCopied
            DeleteListItem ItemBuffer("K" & i).key, _
                    AddK(ItemBuffer("K" & i).parent_node), False
        Next i
        ItemWasCut = False
        ReadytoPasteItem = False
        
        For i = 1 To lvListView.ListItems.count
            If lvListView.ListItems.item(i).Selected = True Then
                lvListView.ListItems.item(i).Ghosted = False
            End If
        Next i
    End If
    
Done:

End Sub

Private Sub mnuRCPopupList2VDetail_Click()
    mnuListViewMode_Click (3)
End Sub

Private Sub mnuRCPopupList2VLIcons_Click()
    mnuListViewMode_Click (0)
End Sub

Private Sub mnuRCPopupList2VList_Click()
    mnuListViewMode_Click (2)
End Sub

Private Sub mnuRCPopupList2VSIcons_Click()
    mnuListViewMode_Click (1)
End Sub



Private Sub mnuRCPopupTree1AddNode_Click(Index As Integer)
    Dim NodeNumber As Long  'node of the parent to the new addition
    Dim FreeNumber As Long
    Dim icon_normal As String
    Dim icon_selected As String
    Dim record As Recordset
    Dim i As Long
    Dim lastpos As Long
    Dim TempPos As Long
    Dim table_name As String
    Dim node_desc As String
    Dim read_only As Boolean
    Dim create_node As Boolean
    Dim create_item As Boolean
    Dim m As Integer
    Dim n As Integer
    Dim tag As Long
    Dim SrcNodeID As Long
    Dim tempstr As String
    
    RButton = False
   
    ' if the custom node
    If RemoveK(nodx.key) = mnuRCPopupTree1AddNode(1).tag Then
        MsgBox "Nodes cannot be added under the Custom QuickAdd node."
        GoTo Done
    End If
    
    'if a child node of a quickadd type node
    If Not nodx.parent Is Nothing Then
        If Not nodx.parent.parent Is Nothing Then
            If RemoveK(nodx.parent.parent.key) = QuickAddNodesNodeID Then
                MsgBox "Nodes cannot be added under a QuickAdd Subnode." & _
                    Chr(13) & "The QuickAdd structure is restricted to a depth of two levels."
                GoTo Done
            End If
        End If
    End If
    
    tempstr = ""
    If tvAttribCol(nodx.key).read_only Then
        tempstr = "Read Only (TRUE)"
    End If
    
    If Not tvAttribCol(nodx.key).create_node Then
        If tempstr = "" Then
            tempstr = "Create Node (FALSE)"
        Else
            tempstr = tempstr & ", Create Node (FALSE)"
        End If
    End If
    
    If tvAttribCol(nodx.key).Link_NodeID <> 0 Then
        If tempstr = "" Then
            tempstr = "Linked (TRUE)"
        Else
            tempstr = tempstr & ", Linked (TRUE)"
        End If
    End If
    
    If tempstr <> "" Then
        MsgBox "You cannot add a node under '" & nodx.text & "' because it has the following " _
            & "attribute(s) set:  " & tempstr & ".", vbExclamation
        GoTo Done
    End If
    
    Me.MousePointer = vbHourglass
    
    tvTreeView.Nodes(nodx.key).Expanded = True
    tag = mnuRCPopupTree1AddNode(Index).tag
    
    tempstr = tvTreeView.Nodes(AddK(Abs(tag))).text
    i = 1
    While ChildNodeHasSameName(nodx.key, tempstr)
        Select Case i
            Case Is >= 10000
                MsgBox "You have too many numbered nodes!  " _
                    & "Please rename some before adding this node.", vbExclamation
                Exit Sub
            Case Is >= 1000     ' 4 digit
                tempstr = Left$(tempstr, Len(tempstr) - 7)
            Case Is >= 100      ' 3 digit
                i = i + 1
                tempstr = Left$(tempstr, Len(tempstr) - 6)
            Case Is >= 10       '2 digit
                i = i + 1
                tempstr = Left$(tempstr, Len(tempstr) - 5)
            Case Else           '1 digit
                If i = 1 Then
                    i = 2
                Else
                    i = i + 1
                    tempstr = Left$(tempstr, Len(tempstr) - 4)
                End If
        End Select
        tempstr = tempstr & " (" & Format$(i) & ")"
    Wend
    
    If Index = 1 Then    'custom
        If tempstr <> "Custom" Then
            MsgBox "The QuickAdd Node 'Custom' was either deleted or renamed!"
        Else
            n = tvTreeView.Nodes(AddK(tag)).Index
            FillTree n, False, True, True, , , tempstr
            'SetTVIcons
            Set nodx = nodx.Child.LastSibling     'set as newly created child
            tvTreeView.SelectedItem = nodx
            lvNeedsRefresh = True
            lvLoaded = False
            tmrNodeClick.Interval = 1
            tvTreeView_NodeClick nodx
            While Not lvLoaded
                DoEvents
            Wend
            tmrNodeClick.Interval = NODE_CLICK_TIMER_DELAY
            nodx.text = tempstr
            tvTreeView.StartLabelEdit
        End If
    Else
        If tag < 0 Then    'if add a Warrior Global
            tag = Abs(tag)
            FillTree tvTreeView.Nodes(AddK(tag)).Index, False, False, False, _
                    True, , tempstr
            'SetTVIcons
            Set nodx = nodx.Child.LastSibling     'set as newly created child
            tvTreeView.SelectedItem = nodx
            lvNeedsRefresh = True
            tvTreeView_NodeClick nodx
            nodx.text = tempstr
        Else    ' if add a QuickAdd
            SrcNodeID = FindQuickTypeRootNode(tvTreeView.Nodes(AddK(tag)).Index)
            FillTree tvTreeView.Nodes(AddK(SrcNodeID)).Index, False, True, True, _
                , , tempstr
            'SetTVIcons
            Set nodx = nodx.Child.LastSibling     'set as newly created child
            tvTreeView.SelectedItem = nodx
            lvNeedsRefresh = True
            tvTreeView_NodeClick nodx
            nodx.text = tempstr
            tvTreeView.StartLabelEdit
        End If
    End If
    lblTitle(1).Caption = "Contents of '" & nodx.text & "'"
Done:
    Me.MousePointer = vbArrow
End Sub

Private Sub mnuRCPopupTree1Copy_Click()
    Dim i As Integer
    Dim tempstr As String
    
    tempstr = ""
    If tvAttribCol(nodx.key).sublink Then
        tempstr = "Sublink"
    End If
    
    If tempstr <> "" Then
        MsgBox "'" & nodx.text & "' cannot be copied because the following " _
            & "attribute(s) are set:  " & tempstr & ".", vbExclamation
        Exit Sub
    End If
        
    'reset status to copy, instead of cut
    ItemWasCut = False
    NodeWasCut = False
    
    NodeBuffer.text = nodx.text
    NodeBuffer.key = nodx.key
    NodeBuffer.read_only = tvAttribCol(nodx.key).read_only
    NodeBuffer.table_name = tvAttribCol(nodx.key).table_name
    NodeBuffer.quicktypeid = tvAttribCol(nodx.key).quicktypeid
    NodeBuffer.icon_normal = tvAttribCol(nodx.key).icon_normal
    NodeBuffer.icon_selected = tvAttribCol(nodx.key).icon_selected
    NodeBuffer.create_node = tvAttribCol(nodx.key).create_node
    NodeBuffer.create_item = tvAttribCol(nodx.key).create_item
    NodeBuffer.system_node = tvAttribCol(nodx.key).system_node
    NodeBuffer.created = tvAttribCol(nodx.key).created
    NodeBuffer.created_by = tvAttribCol(nodx.key).created_by
    NodeBuffer.last_modified = tvAttribCol(nodx.key).last_modified
    NodeBuffer.modified_by = tvAttribCol(nodx.key).modified_by
    NodeBuffer.Link_NodeID = tvAttribCol(nodx.key).Link_NodeID
    NodeBuffer.sublink = tvAttribCol(nodx.key).sublink
                
    'deselect the current item
    lvListView.SelectedItem = Nothing
    Set CurrentItem = Nothing
    
    'set flags since we're ready to paste node(s) and item(s)
    ReadytoPasteItem = True
    ReadytoPasteNode = True
End Sub

Private Sub mnuRCPopupTree1Cut_Click()
    Dim i As Long
    Dim tempstr As String
    
    tempstr = ""
    If tvAttribCol(nodx.key).sublink Then
        tempstr = "Sublink (TRUE)"
    End If
    
    If ScanBranchForAttrib(nodx.Index, tempstr, "cut") Then
        Exit Sub
    End If
    
    'reset status to cut instead of copy
    ItemWasCut = True
    NodeWasCut = True
    
    NodeBuffer.text = nodx.text
    NodeBuffer.key = nodx.key
    NodeBuffer.read_only = tvAttribCol(nodx.key).read_only
    NodeBuffer.table_name = tvAttribCol(nodx.key).table_name
    NodeBuffer.quicktypeid = tvAttribCol(nodx.key).quicktypeid
    NodeBuffer.icon_normal = tvAttribCol(nodx.key).icon_normal
    NodeBuffer.icon_selected = tvAttribCol(nodx.key).icon_selected
    NodeBuffer.create_node = tvAttribCol(nodx.key).create_node
    NodeBuffer.create_item = tvAttribCol(nodx.key).create_item
    NodeBuffer.system_node = tvAttribCol(nodx.key).system_node
    NodeBuffer.created = tvAttribCol(nodx.key).created
    NodeBuffer.created_by = tvAttribCol(nodx.key).created_by
    NodeBuffer.last_modified = tvAttribCol(nodx.key).last_modified
    NodeBuffer.modified_by = tvAttribCol(nodx.key).modified_by
    NodeBuffer.Link_NodeID = tvAttribCol(nodx.key).Link_NodeID
    NodeBuffer.sublink = tvAttribCol(nodx.key).sublink
    
    'deselect the current item
    lvListView.SelectedItem = Nothing
    Set CurrentItem = Nothing
    
    'set flags since we're ready to paste node(s) and item(s)
    ReadytoPasteNode = True
    ReadytoPasteItem = True
End Sub

Private Sub mnuRCPopupTree1Delete_Click()
    Dim record As Recordset
    Dim currentNodesParent As Long
    Dim response As Long
    Dim tempstr As String
    Dim nodekey As String
    Dim Lpos As Integer
    Dim ResetPaste As Boolean
    
    Set record = dbase.OpenRecordset("SELECT COUNT (*) AS [Count] " _
        & "FROM " & DB_NodeTable & " WHERE Link_NodeID = " & RemoveK(nodx.key), _
        dbOpenDynaset)
    record_count = record!count
    record.Close
    
    If tvAttribCol(nodx.key).sublink Then
        tempstr = "Sublink (TRUE)"
    End If
    
    If ScanBranchForAttrib(nodx.Index, tempstr, "delete") Then
        Exit Sub
    End If
    
    If tvAttribCol(nodx.key).table_name = DB_GlobalsTable And record_count > 0 Then
        response = MsgBox("There are " & Format$(record_count) & _
            " nodes currently linked to '" & nodx.text & "'." & _
            Chr(13) & "Deleting this node could have drastic " & _
            "consequences for the database.  Are you sure you " & _
            "want to delete it?", vbCritical + vbYesNo)
        If response = vbYes Then
            response = MsgBox("Do you wish to convert all the linked references " _
                    & "to local copies?" & Chr(13) & Chr(13) & _
                    "NOTE:  If you do not convert the references to local copies, " & _
                    "they will be deleted!  Existing locals will not be affected.", _
                    vbYesNoCancel + vbCritical)
            Select Case response
                Case vbYes
                    Set record = dbase.OpenRecordset("SELECT Node_ID, Table_Name " & _
                        "FROM " & DB_NodeTable & " WHERE Link_NodeID = " & RemoveK(nodx.key), _
                        dbOpenDynaset)
                    Dim oldnodx As Node
                    Set oldnodx = nodx
                    While Not record.EOF
                        Set nodx = tvTreeView.Nodes(AddK(record!node_id))
                        FillTree tvTreeView.Nodes(oldnodx.key).Index, _
                                True, True, False, False, True
                        record.MoveNext
                    Wend
                    Set nodx = oldnodx
                    record.Close
                Case vbNo
                    GoTo BreakLinks
                Case vbCancel
                    Exit Sub
            End Select
BreakLinks:
            Set record = dbase.OpenRecordset("SELECT * FROM " & DB_NodeTable & _
                " WHERE Link_NodeID = " & RemoveK(nodx.key), dbOpenDynaset)
            While Not record.EOF
                record.Edit
                record!Link_NodeID = Null
                record.Update
                nodekey = AddK(record!node_id)
                tvAttribCol(nodekey).Link_NodeID = 0
                'this removes the link overlay from the treeview nodes
                tvTreeView.Nodes(nodekey).Image = _
                        CreateOverlayedImage(imlIconsSmall, "K" & _
                        tvAttribCol(nodekey).icon_normal, _
                        imlTempTree, False, tvAttribCol(nodekey).read_only, _
                        tvAttribCol(nodekey).system_node)
                tvTreeView.Nodes(nodekey).SelectedImage = _
                        CreateOverlayedImage(imlIconsSmall, "K" & _
                        tvAttribCol(nodekey).icon_selected, imlTempTree, _
                        False, tvAttribCol(nodekey).read_only, _
                        tvAttribCol(nodekey).system_node)
                tvTreeView.Nodes(nodekey).ExpandedImage = _
                        tvTreeView.Nodes(nodekey).SelectedImage
                record.MoveNext
            Wend
            record.Close
            GoTo Delete
            Exit Sub
        Else
            Exit Sub
        End If
    End If

    tempstr = ""
    ResetPaste = False
    If ReadytoPasteNode Then
        If (IsDescendantOf(nodx, tvTreeView.Nodes(NodeBuffer.key)) Or _
            IsDescendantOf(tvTreeView.Nodes(NodeBuffer.key), nodx)) Then
                'if deleting a node you just did copy or cut on
                tempstr = "You have copied/cut this node." _
                    & " Deleting it will prevent you from pasting it." & Chr$(13)
                ResetPaste = True
        End If
    End If
    
    If Not CutANDInPasteNode Then
        tempstr = tempstr & "Are you sure you want to delete " & Chr$(34) & _
                nodx.text & Chr$(34) & "?"
        If vbYes = MsgBox(tempstr, vbYesNo + vbExclamation, "Delete Node") Then
            If ResetPaste Then
                ReadytoPasteNode = False
                ReadytoPasteItem = False
                mnuRCPopupTree1Paste.Enabled = False
                mnuEditPaste.Enabled = False
                tbEdit.Buttons("Paste").Enabled = False
            End If
            GoTo Delete
        End If
        Exit Sub
    End If
Delete:
        Me.MousePointer = vbHourglass
        If CutANDInPasteNode Then
            DeleteNode nodx.key, False
        Else
            DeleteNode nodx.key
        End If
        Me.MousePointer = vbArrow
End Sub

Public Function FillTree(ByVal m As Integer, Optional ByVal pastelvItems _
            As Boolean = True, Optional ByVal pasteOneNode As Boolean _
            = False, Optional ByVal SrcIsQuickAdd As Boolean = False, _
            Optional ByVal Linked As Boolean = False, _
            Optional ByVal JustListItems As Boolean = False, _
            Optional DstNodeName As String = "") As Long
    Dim n As Integer
    Dim NodeNumber As Long
    Dim FreeNumber As Long
    Dim record As Recordset
    Dim i As Integer
    Dim SrcStr As String
    Dim DstStr As String
    Dim lvSrcRecord As Recordset
    Dim lvDstRecord As Recordset
    Dim number_of_children As Integer
    Dim last_sib_index As Integer
    Dim X As Integer
    Dim PrevNodeID As Long
    Dim TempNode As Node
    Dim order As Long
    Dim Link_NodeID As Long
    Dim quicktypeid As Long
    
    number_of_children = tvTreeView.Nodes(m).Children
    If number_of_children > 0 Then
        n = tvTreeView.Nodes(m).Child.Index
        last_sib_index = tvTreeView.Nodes(n).LastSibling.Index
    End If
    
    NodeNumber = RemoveK(nodx.key)
    
    If Not JustListItems Then
        If nodx.Children > 0 Then
            Set TempNode = nodx.Child.LastSibling
            While TempNode.Children > 0
                Set TempNode = TempNode.Child.LastSibling
            Wend
            PrevNodeID = RemoveK(TempNode.key)
        Else
            PrevNodeID = RemoveK(nodx.key)
        End If
        Set record = dbase.OpenRecordset("SELECT [Order] FROM " & DB_NodeTable _
            & " WHERE Node_ID = " & PrevNodeID, dbOpenDynaset)
        If Not record.EOF Then
            order = record!order + 1
            record.Close    'must close table before execute
        
            Set record = dbase.OpenRecordset("SELECT [Order] FROM " & DB_NodeTable _
                & " WHERE [Order] >= " & order & " ORDER BY [Order] DESC", dbOpenDynaset)
            While Not record.EOF
                record.Edit
                record!order = record!order + 1
                record.Update
                record.MoveNext
            Wend
            '    dbase.Execute "UPDATE " & DB_NodeTable & " SET [Order]=[Order]+1 " _
            '    & "WHERE [Order] >= " & Order & ";"
        Else
            MsgBox "FillTree:  This node had no previous node!", vbCritical
            FillTree = -1
            record.Close
            Exit Function
        End If
        
        Set record = dbase.OpenRecordset(DB_NodeTable, dbOpenTable)
        record.AddNew
        record!parent = NodeNumber
        If Len(DstNodeName) = 0 Then
            record!node_desc = tvTreeView.Nodes(m).text
        Else
            record!node_desc = DstNodeName
        End If
        record!icon_normal = tvAttribCol(tvTreeView.Nodes(m).key).icon_normal
        record!icon_selected = tvAttribCol(tvTreeView.Nodes(m).key).icon_selected
        record!read_only = tvAttribCol(tvTreeView.Nodes(m).key).read_only
        record!table_name = tvAttribCol(nodx.key).table_name
        record!system_node = False
        record!created = Now
        record!created_by = CurrentUser
        record!last_modified = Now
        record!modified_by = CurrentUser
        record!global_type = Null
        If SrcIsQuickAdd Then
            Link_NodeID = 0
            record!sublink = False
            'here nodx is the dest parent
            If nodx.text = "QuickAdd Nodes" Then
                quicktypeid = 0
                record!create_node = True
                record!create_item = True
            Else
                'here nodx.parent is the dest parent's parent
                If Not nodx.parent Is Nothing Then
                    If nodx.parent.text = "QuickAdd Nodes" Then
                        quicktypeid = 0
                        record!create_node = False
                        record!create_item = False
                    Else
                        quicktypeid = RemoveK(tvTreeView.Nodes(m).key)
                        record!create_node = True
                        record!create_item = True
                    End If
                Else
                    quicktypeid = RemoveK(tvTreeView.Nodes(m).key)
                    record!create_node = True
                    record!create_item = True
                End If
            End If
        Else
            'if dest parent is linked then dest is sublink
            If (tvAttribCol(nodx.key).Link_NodeID <> 0) Or tvAttribCol(nodx.key).sublink Then
                Link_NodeID = RemoveK(tvTreeView.Nodes(m).key)
                record!sublink = True
            Else
                record!sublink = False
                If Linked Then
                    Link_NodeID = RemoveK(tvTreeView.Nodes(m).key)
                Else
                    Link_NodeID = 0
                End If
            End If
            quicktypeid = tvAttribCol(tvTreeView.Nodes(m).key).quicktypeid
            record!create_node = True
            record!create_item = True
        End If
        FreeNumber = FindFreeID(dbase, DB_NodeTable, "Node_ID")
        FillTree = FreeNumber
        record!node_id = FreeNumber
        record!order = order
        AddNode NodeNumber, FreeNumber, tvTreeView.Nodes(m).text, _
                tvAttribCol(tvTreeView.Nodes(m).key).icon_normal, _
                tvAttribCol(tvTreeView.Nodes(m).key).icon_selected, _
                record!read_only, _
                record!table_name, _
                quicktypeid, _
                record!create_item, _
                record!create_node, _
                record!system_node, record!created, _
                record!created_by, record!last_modified, _
                record!modified_by, Link_NodeID, , record!sublink
        If quicktypeid = 0 Then
            record!quicktypeid = Null
        Else
            record!quicktypeid = quicktypeid
        End If
        If Link_NodeID = 0 Then
            record!Link_NodeID = Null
        Else
            record!Link_NodeID = Link_NodeID
        End If
        record.Update
        record.Close
    'after the above call, nodx is now the newly added node
    End If
    
    If pastelvItems Then
        SrcStr = tvAttribCol(tvTreeView.Nodes(m).key).table_name
        DstStr = tvAttribCol(nodx.key).table_name
        Set lvSrcRecord = dbase.OpenRecordset("SELECT * FROM " & SrcStr & " WHERE Parent_Node = " & RemoveK(tvTreeView.Nodes(m).key), dbOpenDynaset)
        Set lvDstRecord = dbase.OpenRecordset(DstStr, dbOpenTable)
        While Not lvSrcRecord.EOF
            FreeNumber = FindFreeID(dbase, DstStr, "Data_ID")
            lvDstRecord.AddNew
            lvDstRecord!data_id = FreeNumber
            lvDstRecord!data_label = lvSrcRecord!data_label
            lvDstRecord!parent_node = RemoveK(nodx.key)
            lvDstRecord!icon_large = lvSrcRecord!icon_large
            lvDstRecord!icon_small = lvSrcRecord!icon_small
            lvDstRecord!data_type = lvSrcRecord!data_type
            lvDstRecord!read_only = lvSrcRecord!read_only
            lvDstRecord!created = lvSrcRecord!created
            lvDstRecord!created_by = lvSrcRecord!created_by
            lvDstRecord!last_modified = Now
            lvDstRecord!modified_by = CurrentUser
            lvDstRecord!variation = lvSrcRecord!variation
            lvDstRecord.Update
            CopyMemo lvSrcRecord, "data_value", lvDstRecord, _
                "data_value", CurTempFolder & "CopyMemo.tmp"
            CopyBLOB lvSrcRecord, "binary_data_value", lvDstRecord, "binary_data_value"
            lvSrcRecord.MoveNext
        Wend
        lvSrcRecord.Close
        lvDstRecord.Close
    End If
    
    If number_of_children > 0 And Not pasteOneNode Then
        While n <> last_sib_index
            '********** added by John 9/11/98 *********************
            If tvAttribCol(tvTreeView.Nodes(n).key).Link_NodeID <> 0 Then
                Linked = True
            Else
                Linked = False
            End If
            '******************************************************
            FillTree n, pastelvItems, , SrcIsQuickAdd, Linked
            ' Set n to the next node's index.
            n = tvTreeView.Nodes(n).Next.Index
        Wend
        '********** added by John 9/11/98 *********************
        If tvAttribCol(tvTreeView.Nodes(n).key).Link_NodeID <> 0 Then
            Linked = True
        Else
            Linked = False
        End If
        '******************************************************
        FillTree n, pastelvItems, False, SrcIsQuickAdd, Linked, JustListItems
    End If
    
    Set nodx = nodx.parent
    
    Me.MousePointer = vbArrow
               
End Function
Private Sub mnuRCPopupTree1Find_Click()
    mnuEditFind_Click
End Sub


Private Sub mnuRCPopupTree1MoveDown_Click()
    Dim TempNodeKey As String
    
    Me.MousePointer = vbHourglass
    
    'remember the node the user clicked on, so it
    'can be reselected after the swap
    TempNodeKey = nodx.key
    
    'swap the nodes in the databse table
    DBMoveNodeDown dbase, nodx
    
    'swap nodes in memory with the next node (since we move down)
    tvSwapNodes nodx.key, nodx.Next.key

    'reselect the original node, now in a new position
    Set nodx = tvTreeView.Nodes(TempNodeKey)
    tvTreeView.SelectedItem = nodx
    nodx.Selected = True
    tvTreeView_NodeClick nodx
    
    Me.MousePointer = vbArrow
End Sub

Private Sub mnuRCPopupTree1MoveUp_Click()
    Dim TempNodeKey As String
    
    Me.MousePointer = vbHourglass
    
    'remember the node the user clicked on, so it
    'can be reselected after the swap
    TempNodeKey = nodx.key
    
    'instead of moving this node up, select the previous
    'node and move it down.  If there was no previous node, the menu option
    'would have been disabled, so there is no need to check for that here.
    DBMoveNodeDown dbase, nodx.Previous
    
    'swap nodes in memory with the previous node (since we move up)
    tvSwapNodes nodx.Previous.key, nodx.key

    'reselect the original node, now in a new position
    Set nodx = tvTreeView.Nodes(TempNodeKey)
    tvTreeView.SelectedItem = nodx
    nodx.Selected = True
    tvTreeView_NodeClick nodx
    
    Me.MousePointer = vbArrow
End Sub

Private Sub mnuRCPopupTree1Paste_Click()
    Dim NewNodeID As Long
    Dim record As Recordset
    Dim Linked As Boolean
    Dim record_count As Long
    Dim oldnodx As Node
    Dim tempstr As String
    Dim SrcNodeOrder As Long
    Dim SrcNodeNextOrder As Long
    Dim DstNodeOrder As Long
    Dim order As Long
    Dim NumNodesToMove As Long
    Dim CurrentNode As Node
    Dim NewOrder As Long
    Dim i As Long
    
    tempstr = ""
    If tvAttribCol(nodx.key).Link_NodeID <> 0 Then
        tempstr = "Linked (TRUE)"
    End If
    
    If tvAttribCol(nodx.key).read_only Then
        If tempstr = "" Then
            tempstr = "Read Only (TRUE)"
        Else
            tempstr = tempstr & ", Read Only (TRUE)"
        End If
    End If
    
    If Not tvAttribCol(nodx.key).create_node Then
        If tempstr = "" Then
            tempstr = "Create Node (FALSE)"
        Else
            tempstr = tempstr & ", Create Node (FALSE)"
        End If
    End If
    
    'see if the source node has any list items under it
    Set record = dbase.OpenRecordset("SELECT COUNT (*) AS [Count] FROM " _
        & NodeBuffer.table_name _
        & " WHERE Parent_Node = " & RemoveK(NodeBuffer.key), dbOpenDynaset)
    record_count = record!count
    record.Close
    
    'if the source node has list items AND the target node has
    'Create_Item to FALSE, don't allow the paste.
    If Not tvAttribCol(nodx.key).create_item And record_count > 0 Then
        If tempstr = "" Then
            tempstr = "Create Item (FALSE)"
        Else
            tempstr = tempstr & ", Create Item (FALSE)"
        End If
    End If
    
    If tempstr <> "" Then
        MsgBox "'" & nodx.text & "' cannot have anything pasted into it because" _
            & " the following attribute(s) are set:  " & tempstr & ".", vbExclamation
        Exit Sub
    End If
    
    If Not tvTreeView.Nodes(NodeBuffer.key) Is Nothing Then
        If IsDescendantOf(nodx, tvTreeView.Nodes(NodeBuffer.key)) Then
            MsgBox "'" & nodx.text & "' cannot have anything pasted into it because" _
                & " the destination node is contained in the source node.", vbExclamation
            Exit Sub
        End If
    End If
    
    'check for duplicate name
    tempstr = NodeBuffer.text
    i = 1
    While ChildNodeHasSameName(nodx.key, tempstr)
        Select Case i
            Case Is >= 10000
                MsgBox "You have too many numbered nodes!  " _
                    & "Please rename some before pasting this node.", vbExclamation
                Exit Sub
            Case Is >= 1000     ' 4 digit
                tempstr = Left$(tempstr, Len(tempstr) - 7)
            Case Is >= 100      ' 3 digit
                i = i + 1
                tempstr = Left$(tempstr, Len(tempstr) - 6)
            Case Is >= 10       '2 digit
                i = i + 1
                tempstr = Left$(tempstr, Len(tempstr) - 5)
            Case Else           '1 digit
                If i = 1 Then
                    i = 2
                Else
                    i = i + 1
                    tempstr = Left$(tempstr, Len(tempstr) - 4)
                End If
        End Select
        tempstr = tempstr & " (" & Format$(i) & ")"
    Wend
        
    Me.MousePointer = vbHourglass
        
    tvTreeView.Nodes(nodx.key).Expanded = True
    
    If NodeWasCut Then
        'set new parent in database
        On Error GoTo Err_Execute
        dbase.Execute "UPDATE " & DB_NodeTable & " SET [Parent] = " & _
                RemoveK(nodx.key) & " WHERE Node_ID = " & RemoveK(NodeBuffer.key), _
                dbFailOnError
        On Error GoTo 0
        
        'get old start order
        Set record = dbase.OpenRecordset("SELECT Node_ID,Order FROM " & _
                DB_NodeTable & " WHERE Node_ID = " & RemoveK(NodeBuffer.key), _
                dbOpenDynaset)
        If Not record.EOF Then
            order = record!order
        Else
            MsgBox "Error: node to move not found!"
        End If
        record.Close
        
        'find next of old & get number of nodes to move
        Set CurrentNode = tvTreeView.Nodes(NodeBuffer.key)
        If Not CurrentNode.Next Is Nothing Then
            Set CurrentNode = CurrentNode.Next
        Else
            Set CurrentNode = NextofParent(CurrentNode, nodx.Root.Index)
            If CurrentNode Is Nothing Then
                NumNodesToMove = FindFreeID(dbase, DB_NodeTable, "Order") - order
                GoTo GotNumNodes
            End If
        End If
        
        Set record = dbase.OpenRecordset("SELECT Node_ID,Order FROM " & _
                DB_NodeTable & " WHERE Node_ID = " & RemoveK(CurrentNode.key), _
                dbOpenDynaset)
        If Not record.EOF Then
            NumNodesToMove = record!order - order
        Else
            MsgBox "Error: next of node to move not found!"
        End If
        record.Close
        
GotNumNodes:
        'find next of dest parent & get new start order
        Set CurrentNode = nodx
        If Not CurrentNode.Next Is Nothing Then
            Set CurrentNode = CurrentNode.Next
        Else
            Set CurrentNode = NextofParent(CurrentNode, nodx.Root.Index)
            If CurrentNode Is Nothing Then
                NewOrder = FindFreeID(dbase, DB_NodeTable, "Order")
                GoTo GotNewOrder
            End If
        End If
        
        Set record = dbase.OpenRecordset("SELECT Node_ID,Order FROM " & _
                DB_NodeTable & " WHERE Node_ID = " & RemoveK(CurrentNode.key), _
                dbOpenDynaset)
        If Not record.EOF Then
            NewOrder = record!order
        Else
            MsgBox "Error: next of dest parent node not found!"
        End If
        record.Close
        
GotNewOrder:
        
        If NewOrder = order Then
            GoTo DoneOrdering
        End If
        'set move selection orders to negatives (make hole)
        On Error GoTo Err_Execute
        dbase.Execute "UPDATE " & DB_NodeTable & " SET [Order] = -1 * [Order]" _
            & " WHERE [Order] BETWEEN " & order & " AND " & _
            order + NumNodesToMove - 1, dbFailOnError
        On Error GoTo 0
        'move hole to new location
        If NewOrder < order Then      'move up
            Set record = dbase.OpenRecordset("SELECT [Order] FROM " & DB_NodeTable _
                & " WHERE [Order] BETWEEN " & NewOrder & " And " _
                & order - 1 & " ORDER BY [Order] DESC", dbOpenDynaset)
            While Not record.EOF
                record.Edit
                record!order = record!order + NumNodesToMove
                record!node_desc = tempstr
                record.Update
                record.MoveNext
            Wend
            record.Close
            'put selection back in new hole
            On Error GoTo Err_Execute
            dbase.Execute "UPDATE " & DB_NodeTable & " SET [Order] = " _
                    & "-1 * [Order] - " & order & " + " & NewOrder _
                    & " WHERE [Order] < 0", dbFailOnError
            On Error GoTo 0
        Else                          'move down
            Set record = dbase.OpenRecordset("SELECT [Order] FROM " & DB_NodeTable _
                & " WHERE [Order] BETWEEN " & order + NumNodesToMove & " And " _
                & NewOrder - 1 & " ORDER BY [Order] ASC", dbOpenDynaset)
            While Not record.EOF
                record.Edit
                record!order = record!order - NumNodesToMove
                record!node_desc = tempstr
                record.Update
                record.MoveNext
            Wend
            record.Close
            'put selection back in new hole
            On Error GoTo Err_Execute
            dbase.Execute "UPDATE " & DB_NodeTable & " SET [Order] = " _
                    & "-1 * [Order] - " & order & " + " & NewOrder - NumNodesToMove _
                    & " WHERE [Order] < 0", dbFailOnError
            On Error GoTo 0
        End If
        
DoneOrdering:
        NewNodeID = RemoveK(NodeBuffer.key)
        Set CurrentNode = nodx
        'clear node from memory
        ClearNode NodeBuffer.key
        'load the node and children
        LoadTreeBranch NewNodeID, tvwChild, RemoveK(nodx.key)
        Set nodx = CurrentNode
        'SetTVIcons
        
        'reset flags
        ReadytoPasteItem = False
        ReadytoPasteNode = False
        ItemWasCut = False
        NodeWasCut = False
    
    Else    ' this is for copying
        If NodeBuffer.Link_NodeID <> 0 Then
            Linked = True
        Else
            Linked = False
        End If
        NewNodeID = FillTree(tvTreeView.Nodes(NodeBuffer.key).Index, , _
                , , Linked, , tempstr)
        'SetTVIcons
    End If
    
    'choose the newly pasted node
    Set nodx = nodx.Child.LastSibling     'set as newly created child
    tvTreeView.SelectedItem = nodx
    lvNeedsRefresh = True
    tvTreeView_NodeClick nodx
    nodx.text = tempstr
    GoTo Done
    
Err_Execute:
    DisplayDBEngineErrors

Done:
    Me.MousePointer = vbArrow
    
End Sub

Private Sub mnuRCPopupTree1Print_Click()
    mnuFilePrint_Click
End Sub

Private Sub mnuRCPopupTree1Properties_Click()
    mnuNodesProperties_Click
End Sub

Private Sub mnuRCPopupTree1Rename_Click()
    Dim tempstr As String
    
    tempstr = ""
    If tvAttribCol(nodx.key).sublink Then
        tempstr = "Sublink (TRUE)"
    End If
    
    If tvAttribCol(nodx.key).read_only Then
        If tempstr = "" Then
            tempstr = "Read Only (TRUE)"
        Else
            tempstr = tempstr & ", Read Only (TRUE)"
        End If
    End If
    
    If tempstr <> "" Then
        MsgBox "'" & nodx.text & "' cannot be renamed because the following " _
            & "attribute(s) are set:  " & tempstr & ".", vbExclamation
        Exit Sub
    End If
    
    tvTreeView.StartLabelEdit
End Sub

Private Sub mnuRCPopupTree1Replace_Click()
    mnuEditReplace_Click
End Sub

Private Sub mnuToolsDatabase_Click()
    frmDatabase.Show vbModeless, Me
End Sub

Private Sub mnuToolsSecurity_Click()
    MsgBox "TEMP:  Security goes here.", vbInformation
End Sub

Private Sub mnuToolsSpelling_Click()
    MsgBox "TEMP:  Spell checker goes here.", vbInformation
End Sub

Private Sub mnuVAIByDataID_Click()
    lvListView.SortOrder = lvwAscending
    lvListView.SortKey = DATA_ID_COLUMN
End Sub

Private Sub mnuVAIByItem_Click()
    lvListView.SortOrder = lvwAscending
    lvListView.SortKey = ITEM_COLUMN
End Sub

Private Sub mnuVAIByParentNode_Click()
    lvListView.SortOrder = lvwAscending
    lvListView.SortKey = PARENT_NODE_COLUMN
End Sub

Private Sub mnuVAIByType_Click()
    lvListView.SortOrder = lvwAscending
    lvListView.SortKey = TYPE_COLUMN
End Sub

Private Sub mnuViewOptions_Click()
    frmMainOptions.Show vbModal, Me
End Sub

Private Sub mnuViewStatusBar_Click()
    If mnuViewStatusBar.Checked Then
        sbStatusBar.Visible = False
        mnuViewStatusBar.Checked = False
    Else
        sbStatusBar.Visible = True
        mnuViewStatusBar.Checked = True
    End If
    SizeControls imgSplitter.Left
End Sub


Private Sub mnuViewToolbar_Click()
    If mnuViewToolbar.Checked Then
        CoolBar1.Visible = False
        mnuViewToolbar.Checked = False
    Else
        CoolBar1.Visible = True
        mnuViewToolbar.Checked = True
    End If
    SizeControls imgSplitter.Left
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub


Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim sglPos As Single
    

    If mbMoving Then
        sglPos = X + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub


Public Sub SizeControls(ByVal SplitPos As Integer, _
        Optional first_time As Boolean = False)
    
    SendMessageLong lvListView.hwnd, WM_SETREDRAW, 0, ByVal 0&
    'set the width
    If SplitPos < sglSplitLimit Then
        SplitPos = sglSplitLimit 'minimum width of tree view
    End If
    If SplitPos > (Me.Width - sglSplitLimit) Then
        SplitPos = Me.Width - sglSplitLimit 'max width of tree view
    End If
    tvTreeView.Width = SplitPos
    imgSplitter.Left = SplitPos
    LastSplitterLeft = SplitPos
    lvListView.Left = SplitPos + 23
    lvListView.Width = Me.Width - (tvTreeView.Width + 125)
    lblTitle(0).Width = tvTreeView.Width - 7
    lblTitle(1).Left = lvListView.Left + 15
    lblTitle(1).Width = lvListView.Width - 30
    'set the top
    If CoolBar1.Visible Or first_time Then
        'picTitles.Top = picTBContainer.Top + picTBContainer.Height + 100
        picTitles.Top = CoolBar1.Top + CoolBar1.Height + 100
        lvListView.Top = picTitles.Top + picTitles.Height + 1
        tvTreeView.Top = lvListView.Top + 20
    Else
        'picTitles.Top = picTBContainer.Top + 100
        picTitles.Top = CoolBar1.Top + 100
        lvListView.Top = picTitles.Top + picTitles.Height + 1
        tvTreeView.Top = lvListView.Top + 20
    End If
    

    'set the height
    If sbStatusBar.Visible Or first_time Then
        tvTreeView.Height = Me.ScaleHeight - lvListView.Top - sbStatusBar.Height - 35
    Else
        tvTreeView.Height = Me.ScaleHeight - lvListView.Top - 35
    End If
    

    lvListView.Height = tvTreeView.Height + 30
    imgSplitter.Top = tvTreeView.Top
    imgSplitter.Height = tvTreeView.Height
    
    SendMessageLong lvListView.hwnd, WM_SETREDRAW, 1, ByVal 0&
        
End Sub

Public Sub UpdateStatusBar(itemsSelected As Long, Optional initial_load As Boolean = False)
    If initial_load Then
        On Error Resume Next    'if the picture does not exist, then blank picture
        sbStatusBar.Panels.item(1).picture = LoadPicture(SmallIconFolder & SBICON_Logo)
        On Error GoTo 0
    Else
        If itemsSelected = 0 Or TypeOf Me.ActiveControl Is TreeView Then
            sbStatusBar.Panels.item(1).text = Format$(lvListView.ListItems.count) & " Item(s)"
        Else
            sbStatusBar.Panels.item(1).text = Format$(itemsSelected) & " Item(s) selected"
        End If
    End If
End Sub


Private Sub tbEdit_ButtonClick(ByVal Button As Button)

    Select Case Button.key
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "Add"
            If TypeOf Me.ActiveControl Is TreeView Then
                mnuNodesAdd_Click
            ElseIf TypeOf Me.ActiveControl Is ListView Then
                mnuItemsAdd_Click
            End If
        Case "Rename"
            If TypeOf Me.ActiveControl Is TreeView Then
                mnuNodesRename_Click
            ElseIf TypeOf Me.ActiveControl Is ListView Then
                mnuItemsRename_Click
            End If
        Case "MoveUp"
            mnuRCPopupTree1MoveUp_Click
        Case "MoveDown"
            mnuRCPopupTree1MoveDown_Click
        Case "Execute"
            mnuRCPopupList1Execute_Click
        Case "Variation"
            mnuRCPopupList1Variation_Click
    End Select
End Sub

Private Sub tbExternal_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.key
        Case "HTML"
            mnuModuleHTML_Click
        Case "Source"
            mnuModuleSource_Click
    End Select
End Sub

Private Sub tbstand_ButtonClick(ByVal Button As Button)

    Select Case Button.key
        Case "Prop"
            If TypeOf Me.ActiveControl Is TreeView Then
                mnuNodesProperties_Click
            ElseIf TypeOf Me.ActiveControl Is ListView Then
                mnuItemsProperties_Click
            End If
        Case "Open"
            mnuFileOpen_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Refresh"
            RefreshNodes True
        Case "Find"
            mnuEditFind_Click
        Case "Replace"
            mnuEditReplace_Click
        Case "Help"
            mnuHelpContents_Click
    End Select
End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer
    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If err Then
            MsgBox err.description
        End If
        On Error GoTo 0
    End If
End Sub


Private Sub mnuHelpSearch_Click()
    

    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        err.Clear
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If err Then
            MsgBox err.description
        End If
        On Error GoTo 0
    End If
End Sub


Private Sub mnuListViewMode_Click(Index As Integer)
    'uncheck the current type
    mnuListViewMode(lvListView.View).Checked = False
    'set the listview mode
    lvListView.View = Index
    'check the new type
    mnuListViewMode(Index).Checked = True

    If Index = lvwReport Then
        LoadListView nodx.key
    End If
    Select Case lvListView.View
        Case lvwIcon, lvwSmallIcon
            lvListView.Arrange = lvwAutoTop
        Case Else
            lvListView.Arrange = lvwAutoLeft
    End Select
End Sub

Private Sub mnuViewRefresh_Click()
    RefreshNodes True
End Sub

Private Sub mnuEditCopy_Click()
    Dim i As Long

    If TypeOf Me.ActiveControl Is ListView Then
        For i = 1 To lvListView.ListItems.count
            If lvListView.ListItems.item(i).Selected = True Then
                lvListView.ListItems.item(i).Ghosted = False
            End If
        Next i
        mnuRCPopupList1Copy_Click
    Else
        mnuRCPopupTree1Copy_Click
    End If
End Sub


Private Sub mnuEditCut_Click()
    Dim i As Long
    
    If TypeOf Me.ActiveControl Is ListView Then
        mnuRCPopupList1Cut_Click
    Else
        mnuRCPopupTree1Cut_Click
    End If
End Sub


Private Sub mnuitemsSelectAll_Click()
    UpdateStatusBar 0
End Sub


Private Sub mnuitemsInvertSelection_Click()
    Dim i As Integer
    Dim itemsSelected As Long
    
        itemsSelected = 0
        For i = 1 To lvListView.ListItems.count
            If lvListView.ListItems.item(i).Selected = True Then
                lvListView.ListItems.item(i).Selected = False
            Else
                lvListView.ListItems.item(i).Selected = True
                itemsSelected = itemsSelected + 1
            End If
        Next i
        UpdateStatusBar itemsSelected
    
End Sub


Private Sub mnuEditPaste_Click()
    If TypeOf Me.ActiveControl Is ListView Then
        If ReadytoPasteItem Then
            mnuRCPopupList2Paste_Click
        End If
    Else
        If ReadytoPasteNode Then
            mnuRCPopupTree1Paste_Click
        End If
    End If
End Sub

Private Sub mnuFileOpen_Click()
    CommonDialog1.Filter = "Access Database (*.mdb)|*.mdb"
    CommonDialog1.FilterIndex = 0
    On Error GoTo userCancel
        CommonDialog1.ShowOpen
    On Error GoTo 0
    CurrentDatabaseFile = CommonDialog1.filename
    CompactDB CurrentDatabaseFile, dbase, True, fProgForm
    RefreshNodes True
userCancel:
    On Error GoTo 0
End Sub

Private Sub tbUtilities_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.key
        Case "DB"
            mnuToolsDatabase_Click
        Case "Spelling"
            mnuToolsSpelling_Click
        Case "Security"
            mnuToolsSecurity_Click
    End Select
End Sub

'This event triggers loading of the listview 1/2 second after the node is clicked.
'This prevents slowdown when cycling through all the nodes with the keyboard.
Private Sub tmrNodeClick_Timer()
    LoadListView nodx.key
    DoEvents
    tmrNodeClick.Enabled = False
End Sub

Public Sub tvTreeView_AfterLabelEdit(Cancel As Integer, NewString As String)
    
    Dim record As Recordset

    Set record = dbase.OpenRecordset("SELECT * FROM " & DB_NodeTable _
            & " WHERE Node_ID = " & RemoveK(nodx.key), dbOpenDynaset)
            
    If record.EOF Then
        GoTo Done
    End If
    If Len(NewString) < 1 Then
        Cancel = True
        GoTo Done
    End If
    
    'by now, we have a potentially valid name, but we need to check if it is already
    'in use.  If it is, cancel the rename.
    If Not nodx.parent Is Nothing Then
        If ChildNodeHasSameName(nodx.parent.key, NewString) Then
            Cancel = True
            MsgBox "'" & nodx.parent.text & "' already has a node called '" & _
                    NewString & "'." & Chr(13) & _
                    "Please give the current node a different name.", vbExclamation
            tvTreeView.StartLabelEdit
            GoTo Done
        End If
    End If
    
    
    If InStr(1, NewString, "'", vbTextCompare) <> 0 Or _
            InStr(1, NewString, Chr(34), vbTextCompare) <> 0 Then
        Cancel = True
        MsgBox "Quotes and Apostrophes are not allowed; please reenter.", _
                vbExclamation
        tvTreeView.StartLabelEdit
        GoTo Done
    End If
    
    record.Edit
    record!node_desc = NewString
    record!last_modified = Now
    record!modified_by = CurrentUser
    record.Update
    nodx.text = NewString
    lblTitle(1).Caption = "Contents of '" & nodx.text & "'"

Done:
    record.Close
End Sub

Private Sub tvTreeView_GotFocus()
    
    If tvTreeView.Nodes.count = 0 Then            'enable/disable the delete button on the toolbar
        tbEdit.Buttons("Delete").Enabled = False
    Else
        tbEdit.Buttons("Delete").Enabled = True
    End If
    
    If nodx Is Nothing Then
        tbEdit.Buttons("Rename").Enabled = False
    Else
        tbEdit.Buttons("Rename").Enabled = True
    End If
    
    tbEdit.Buttons("Execute").Enabled = False
    tbEdit.Buttons("Variation").Enabled = False
    tbEdit.Buttons("Edit").Enabled = False
    
    If mnuItems.Enabled Then mnuItems.Enabled = False
    If Not mnuNodes.Enabled Then mnuNodes.Enabled = True
    UpdateStatusBar 0
End Sub

Private Sub tvTreeView_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim n As Integer
    Dim i As Integer
    Dim j As Integer
    Dim ShiftDown, AltDown, CtrlDown
    
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    
    Select Case KeyCode
    Case vbKeyDelete
        mnuRCPopupTree1Delete_Click
    Case vbKeyInsert
        mnuNodesAdd_Click
    Case vbKeyR And CtrlDown
        If Not tvAttribCol(nodx.key).read_only Then   'if read_only
            mnuRCPopupTree1Rename_Click
        End If
    Case vbKeyS And CtrlDown     'hotkey to ensure nodx visible
        nodx.EnsureVisible

    '''''Brian added this line on 8/21/98''''''
    Case vbKeyF3
    '''''Brian added this line on 8/21/98''''''
        If StartFindNodeIndex <> 0 Then
    '''''Brian added this line on 8/21/98''''''
            fReplaceForm.cmdFindNext_Click
    '''''Brian added this line on 8/21/98''''''
        End If

    Case vbKeyF5
        mnuViewRefresh_Click
    Case vbKeyF10 And Shift = 1, 93
        If ReadytoPasteNode Or ReadytoPasteItem Then
            mnuRCPopupTree1Paste.Enabled = True
            mnuEditPaste.Enabled = True
            tbEdit.Buttons("Paste").Enabled = True
        Else
            mnuRCPopupTree1Paste.Enabled = False
            mnuEditPaste.Enabled = False
            tbEdit.Buttons("Paste").Enabled = False
        End If
        
        i = 0
        n = nodx.Index
        While tvTreeView.Nodes(n).Visible And n <> nodx.Root.Index
            n = tvTreeView.Nodes(n).parent.Index
            i = i + 1
        Wend
        
        j = 0
        n = nodx.Index
        While tvTreeView.Nodes(n).Visible And n <> nodx.Root.Index
            j = j + FindNodeHeight(n)
            n = tvTreeView.Nodes(n).parent.Index
        Wend
        If Not tvTreeView.Nodes(n).Visible Then
            j = j - 1
        End If
        Call PopupMenu(mnuRCPopupTree1, , i * (tvTreeView.Indentation + 30) + _
            tvTreeView.Left + 500, j * 240 + tvTreeView.Top + 200)
    Case vbKey8 And Shift = 1, vbKeyMultiply
        If nodx.Expanded Then
            nodxExpanded = True
        Else
            nodxExpanded = False
        End If
    Case Else
    End Select
    UpdateStatusBar 0
    
End Sub

Private Function FindNodeHeight(ByVal n As Integer) As Integer
        
        FindNodeHeight = 0
        While tvTreeView.Nodes(n).Visible And n <> _
                tvTreeView.Nodes(n).FirstSibling.Index
            n = tvTreeView.Nodes(n).Previous.Index
            If tvTreeView.Nodes(n).Expanded Then
                FindNodeHeight = FindNodeHeight + _
                FindNodeHeight( _
                    tvTreeView.Nodes(n).Child.LastSibling.Index) + 1
            Else
                FindNodeHeight = FindNodeHeight + 1
            End If
        Wend
        If tvTreeView.Nodes(n).Visible Then
            FindNodeHeight = FindNodeHeight + 1
        End If
        
End Function

Private Sub FullCollapse(ByVal m As Integer)

    Dim n As Integer
    
    If tvTreeView.Nodes(m).Children > 0 Then
        n = tvTreeView.Nodes(m).Child.Index
        While n <> tvTreeView.Nodes(n).LastSibling.Index
            FullCollapse n
            ' Set n to the next node's index.
            n = tvTreeView.Nodes(n).Next.Index
        Wend
        FullCollapse n
        tvTreeView.Nodes(m).Expanded = False
    End If
            
End Sub

Private Sub FullExpand(ByVal m As Integer)

    Dim n As Integer
    
    If tvTreeView.Nodes(m).Children > 0 Then
        n = tvTreeView.Nodes(m).Child.Index
        While n <> tvTreeView.Nodes(n).LastSibling.Index
            FullExpand n
            ' Set n to the next node's index.
            n = tvTreeView.Nodes(n).Next.Index
        Wend
        FullExpand n
        tvTreeView.Nodes(m).Expanded = True
    End If
            
End Sub

Private Sub tvTreeView_MouseDown(Button As Integer, _
                Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        RButton = True
    End If
    Set DragItem = tvTreeView.HitTest(X, Y)
    MouseKey = Shift
End Sub

Private Sub tvTreeView_MouseUp(Button As Integer, Shift As Integer, _
        X As Single, Y As Single)
    MouseKey = -1
    Set DragItem = Nothing
End Sub

Public Sub tvTreeView_NodeClick(ByVal Node As Node)
    Dim tempstr As String
    
    'disable the load listview timer with every node click
    'it is enabled below, if necessary
    tmrNodeClick.Enabled = False
    
    FocusFrom = "TreeView" 'global used in InitProperties
    
    ' if the user is not clicking on the current node (repeat click) OR
    ' the listview needs to be refreshed OR
    ' the Properties Dialog is currently open
    If Node.key <> nodx.key Or lvNeedsRefresh Or PropertiesActive Then
        Set nodx = Node
        UpdateAddPopupMenus  'recalculate the popupmenu items
        If RButton = False Then 'clicked with L Button
            'enable the countdown timer to load the listview
            'when this timer finishes counting, it will load the listview
            tmrNodeClick.Enabled = True
            
            'update the title for the listview
            lblTitle(1).Caption = "Contents of '" & nodx.text & "'"
        End If
        'listview has been refreshed, so clear the flag
        lvNeedsRefresh = False
    End If
    
    'Is the Properties Dialog currently open ?
    If PropertiesActive Then
        mnuRCPopupTree1Properties.Enabled = False 'disable the Properties Menu Item
        LastTab = fPropForm.SSTab1.Tab 'remember the last tab position
        fPropForm.InitPropertiesForm 're-load all values for Properties
        DoEvents  'is this necessary ????
    Else
        mnuRCPopupTree1Properties.Enabled = True 'enable the Properties Menu Item
        LastTab = 0 'reset to the first tab
    End If
    
    '*** Enable/Disable Move Up & Down Menu Items ***
    If tvAttribCol(nodx.key).sublink Then
        'Can't move sublinks up & down in the treeview
        mnuRCPopupTree1MoveUp.Enabled = False
        mnuRCPopupTree1MoveDown.Enabled = False
        mnuNodesMoveUp.Enabled = False
        mnuNodesMoveDown.Enabled = False
        tbEdit.Buttons("MoveUp").Enabled = False
        tbEdit.Buttons("MoveDown").Enabled = False
    Else
        
        'Can't move the first sibling up
        If nodx.key = nodx.FirstSibling.key Then
            mnuRCPopupTree1MoveUp.Enabled = False
            mnuNodesMoveUp.Enabled = False
            tbEdit.Buttons("MoveUp").Enabled = False
        Else
            mnuRCPopupTree1MoveUp.Enabled = True
            mnuNodesMoveUp.Enabled = True
            tbEdit.Buttons("MoveUp").Enabled = True
        End If
    
        'Can't move last sibling down
        If nodx.key = nodx.LastSibling.key Then
            mnuRCPopupTree1MoveDown.Enabled = False
            mnuNodesMoveDown.Enabled = False
            tbEdit.Buttons("MoveDown").Enabled = False
        Else
            mnuRCPopupTree1MoveDown.Enabled = True
            mnuNodesMoveDown.Enabled = True
            tbEdit.Buttons("MoveDown").Enabled = True
        End If
    End If
        
    If ReadytoPasteItem Or ReadytoPasteNode Then
        mnuRCPopupTree1Paste.Enabled = True
        mnuEditPaste.Enabled = True
        tbEdit.Buttons("Paste").Enabled = True
    Else
        mnuRCPopupTree1Paste.Enabled = False
        mnuEditPaste.Enabled = False
        tbEdit.Buttons("Paste").Enabled = False
    End If
    If RButton = True Then
        PopupMenu mnuRCPopupTree1
        RButton = False
    End If

End Sub
Private Sub lvListView_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    lvListView.SortKey = ColumnHeader.Index - 1
    ' Set Sorted to True to sort the list.
    lvListView.Sorted = True
    If lvListView.SortOrder = lvwAscending Then
        lvListView.SortOrder = lvwDescending
    Else
        lvListView.SortOrder = lvwAscending
    End If
    
End Sub

Private Sub tvTreeView_OLEDragDrop(data As MSComctlLib.DataObject, _
                Effect As Long, Button As Integer, _
                Shift As Integer, X As Single, Y As Single)
    
    Set CurrentData = data
    DragTarget = "TREE"
    Effect = vbDropEffectNone
    Set nodx = tvTreeView.HitTest(X, Y)
    tvTreeView.DropHighlight = Nothing

    If nodx Is Nothing Then
        GoTo Done
    End If
    If nodx.key <> nodx.Root.key Then
        If FindANode("Inbox", False, True) = RemoveK(nodx.parent.key) Then
            DragSubTarget = nodx.text
        End If
    End If
    
    'if dragging with the Right Button
    If RButton = True Then
        RButton = False
        Select Case DragSource
            Case "LIST"
                If Not nodx Is Nothing Then
                    'if dropping onto the Inbox node
                    If FindANode("Inbox", False, True) = RemoveK(nodx.key) Then
                        mnuInboxAutoSort.Enabled = True
                        mnuInboxCopy.Enabled = False
                        PopupMenu mnuInbox
                    'if dropping onto a child of Inbox node
                    ElseIf Not nodx.parent Is Nothing Then
                        If FindANode("Inbox", False, True) = RemoveK(nodx.parent.key) Then
                            DragDropped = True
                            mnuInboxAutoSort.Enabled = False
                            mnuInboxCopy.Enabled = True
                            PopupMenu mnuInbox
                        Else
                            DragDropped = True
                            PopupMenu mnuDrag
                        End If
                    'if dropping on any other node
                    Else
                        DragDropped = True
                        PopupMenu mnuDrag
                    End If
                Else
                    PopupMenu mnuDrag
                End If
            Case "TREE"
                PopupMenu mnuDrag
            Case Else
                If (Not nodx Is Nothing) And (data.GetFormat(ccCFFiles) Or _
                        data.GetFormat(ccCFText)) Then
                        If (FindANode("Inbox", False, True) = RemoveK(nodx.key)) Then
                            mnuInboxAutoSort.Enabled = True
                            mnuInboxCopy.Enabled = False
                            PopupMenu mnuInbox
                        ElseIf (FindANode("Inbox", False, True) = _
                                RemoveK(nodx.parent.key)) Then
                            mnuInboxAutoSort.Enabled = False
                            mnuInboxCopy.Enabled = True
                            PopupMenu mnuInbox
                        End If
                End If
        End Select
    'if dragging with the Left Button
    Else
        Select Case DragSource
            Case "TREE"
                mnuRCPopupTree1Paste_Click
            Case "LIST"
                If MouseKey = vbCtrlMask Then
                    mnuDragCopy_Click
                Else
                    mnuDragMove_Click
                End If
            Case Else
                If (Not nodx Is Nothing) And (FindANode("Inbox", False, True) = _
                    RemoveK(nodx.parent.key)) And (data.GetFormat(ccCFFiles) Or _
                    data.GetFormat(ccCFText)) Then
                        mnuInboxCopy_Click
                End If
        End Select
    End If
Done:
    DragSource = ""
    Set DragItem = Nothing
End Sub

Private Sub tvTreeView_OLEDragOver(data As MSComctlLib.DataObject, _
                Effect As Long, Button As Integer, _
                Shift As Integer, X As Single, Y As Single, _
                state As Integer)
    
    Dim TempNode As Node
                
    If Button = 2 Then
        RButton = True
    Else
        RButton = False
    End If
    
    
    Set TempNode = tvTreeView.HitTest(X, Y)
    
    Select Case DragSource
        Case "TREE"
            If RButton Then
                If TempNode Is Nothing Then
                    Effect = vbDropEffectNone
                    Set tvTreeView.DropHighlight = Nothing
                Else
                    Effect = vbDropEffectCopy
                    Set tvTreeView.DropHighlight = TempNode
                End If
            Else
                If Shift = vbCtrlMask Then
                    If TempNode Is Nothing Then
                        Effect = vbDropEffectNone
                        Set tvTreeView.DropHighlight = Nothing
                    Else
                        Effect = vbDropEffectCopy
                        Set tvTreeView.DropHighlight = TempNode
                    End If
                Else
                    If TempNode Is Nothing Then
                        Effect = vbDropEffectNone
                        Set tvTreeView.DropHighlight = Nothing
                    Else
                        Effect = vbDropEffectMove
                        Set tvTreeView.DropHighlight = TempNode
                    End If
                End If
            End If
            
        Case "LIST"
            If RButton Then
                If TempNode Is Nothing Then
                   Effect = vbDropEffectNone
                   Set tvTreeView.DropHighlight = Nothing
                Else
                   Effect = vbDropEffectCopy
                   Set tvTreeView.DropHighlight = TempNode
                End If
            Else
                If Shift = vbCtrlMask Then
                    If TempNode Is Nothing Then
                        Effect = vbDropEffectNone
                        Set tvTreeView.DropHighlight = Nothing
                    Else
                        Effect = vbDropEffectCopy
                        Set tvTreeView.DropHighlight = TempNode
                    End If
                Else
                    If TempNode Is Nothing Then
                        Effect = vbDropEffectNone
                        Set tvTreeView.DropHighlight = Nothing
                    Else
                        Effect = vbDropEffectMove
                        Set tvTreeView.DropHighlight = TempNode
                    End If
                End If
            End If
        Case Else
            If RButton Then
                If (Not TempNode Is Nothing) Then
                    If ((FindANode("Inbox", False, True) = RemoveK(TempNode.key)) _
                        Or (FindANode("Inbox", False, True) = RemoveK(TempNode.parent.key))) Then
                        Effect = vbDropEffectCopy
                        Set tvTreeView.DropHighlight = TempNode
                    Else
                        Effect = vbDropEffectNone
                        Set tvTreeView.DropHighlight = Nothing
                    End If
                Else
                   Effect = vbDropEffectNone
                   Set tvTreeView.DropHighlight = Nothing
                End If
            Else
                ' in this case, both L drag and L-Ctrl drag
                ' will result in a copy.
                If (Not TempNode Is Nothing) Then
                    If FindANode("Inbox", False, True) = _
                        RemoveK(TempNode.parent.key) Then
                            Effect = vbDropEffectCopy
                            Set tvTreeView.DropHighlight = TempNode
                    Else
                        Effect = vbDropEffectNone
                        Set tvTreeView.DropHighlight = Nothing
                    End If
                Else
                    Effect = vbDropEffectNone
                    Set tvTreeView.DropHighlight = Nothing
                End If
            End If

    End Select
End Sub

Private Sub tvTreeView_OLEStartDrag(data As MSComctlLib.DataObject, _
                    AllowedEffects As Long)
                    
    DragSource = "TREE"
    If DragItem Is Nothing Then
        AllowedEffects = vbDropEffectNone 'nothing
    Else
        Set nodx = DragItem
        If MouseKey = vbCtrlMask Or RButton = True Then
            AllowedEffects = vbDropEffectCopy 'copy Ctrl+Drag
            mnuRCPopupTree1Copy_Click
        Else
            AllowedEffects = vbDropEffectMove 'move Drag
            mnuRCPopupTree1Cut_Click
        End If
    End If
End Sub

Public Sub UpdateAddPopupMenus()
    Dim tempstr As String
    Dim record As Recordset
    Dim currentnumend As Integer
    Dim currentnumstart As Integer
    Dim itemnumber As Long
    Dim strlength As Long
    Dim menuindex As Integer
    Dim node_id As Long
    Dim record_count As Long
    Dim m As Long
    Dim n As Long
    Dim MenuCount As Integer
    
    'make custom item visible
    mnuRCPopupTree1AddNode(1).Visible = True
    'clear treeview add popup menu
    MenuCount = mnuRCPopupTree1AddNode.count
    For menuindex = MenuCount - 1 To 2 Step -1
        Unload mnuRCPopupTree1AddNode(menuindex)
    Next menuindex
    
    'clear listview add popup menu
    
    MenuCount = mnuRCPopupList2AddItem.count
    For menuindex = MenuCount To 2 Step -1
        Unload mnuRCPopupList2AddItem(menuindex)
    Next menuindex
    
    ' Disable the Add Menu Option for the listview if the Create Item flag is FALSE
    ' or if the node is a sublink
'    If (Not tvAttribCol(nodx.key).create_item) Or tvAttribCol(nodx.key).sublink Then
'        mnuRCPopupList2Add.Enabled = False
'    Else
'        mnuRCPopupList2Add.Enabled = True
'    End If
    
    If tvAttribCol(nodx.key).table_name = DB_DataTypesTable Then
        mnuRCPopupList2AddItem(1).Caption = "Custom Data Type"
        mnuRCPopupList2AddItem(1).tag = DEFAULT_LARGE_ICON & _
                "," & DEFAULT_SMALL_ICON & "," & "String"
        number_of_custom_items = 1
    Else
        record_count = 2
        Set record = dbase.OpenRecordset("SELECT * FROM " & DB_DataTypesTable _
            & " ORDER BY Data_Label", dbOpenDynaset)
        If Not record.EOF Then
            mnuRCPopupList2AddItem(1).Caption = "Custom " & record!data_label
            mnuRCPopupList2AddItem(1).tag = record!icon_large & _
                "," & record!icon_small & "," & record!data_type
            record.MoveNext
        End If
        While Not record.EOF
            Load mnuRCPopupList2AddItem(record_count)
            mnuRCPopupList2AddItem(record_count).Caption = "Custom " & record!data_label
            mnuRCPopupList2AddItem(record_count).tag = record!icon_large & _
                "," & record!icon_small & "," & record!data_type
            record.MoveNext
            record_count = record_count + 1
        Wend
        record.Close
        number_of_custom_items = record_count - 1
    End If
    
    
    node_id = tvAttribCol(nodx.key).quicktypeid
    If node_id = 0 Then
        ' if a root quick add node
        If Not nodx.parent Is Nothing Then
            If RemoveK(nodx.parent.key) = QuickAddNodesNodeID Then
                Set record = dbase.OpenRecordset("SELECT * FROM " & DB_NodeTable & " WHERE parent = " _
                        & QuickAddNodesNodeID & " ORDER BY [Order]", dbOpenDynaset)
                If Not record.EOF Then
                    'set placeholder visible so can make custom item invisible
                    mnuRCPopupTree1AddNode(0).Visible = True
                    'make custom item invisible
                    mnuRCPopupTree1AddNode(1).Visible = False
                    menuindex = 2
                    If Not record.EOF Then
                        'loop until get first valid quicktype root
                        While menuindex = 2 And Not record.EOF
                            ' if not the current node and not the custom node
                            If record!node_id <> RemoveK(nodx.key) And _
                                    record!node_id <> mnuRCPopupTree1AddNode(1).tag Then
                                Load mnuRCPopupTree1AddNode(menuindex)
                                'now can make placeholder invisible
                                mnuRCPopupTree1AddNode(0).Visible = False
                                mnuRCPopupTree1AddNode(menuindex).Caption = record!node_desc
                                mnuRCPopupTree1AddNode(menuindex).tag = record!node_id
                                ' this required since it loads initially as invisible
                                mnuRCPopupTree1AddNode(menuindex).Visible = True
                                menuindex = menuindex + 1
                            End If
                            record.MoveNext
                            'if couldn't find any valid quicktypes
                            If menuindex = 2 And record.EOF Then
                                'just load the separator and make placeholder invisible
                                Load mnuRCPopupTree1AddNode(2)
                                mnuRCPopupTree1AddNode(2).Caption = "-"
                                ' this required since it loads initially as invisible
                                mnuRCPopupTree1AddNode(2).Visible = True
                                mnuRCPopupTree1AddNode(0).Visible = False
                            End If
                        Wend
                    Else
                        'if not valid quicktypes then load separator and make placeholder
                        'invisible
                        Load mnuRCPopupTree1AddNode(2)
                        ' this required since it loads initially as invisible
                        mnuRCPopupTree1AddNode(2).Visible = True
                        mnuRCPopupTree1AddNode(2).Caption = "-"
                        mnuRCPopupTree1AddNode(0).Visible = False
                    End If
                    'load the rest of the quicktypes into the menu
                    While Not record.EOF
                        ' if not the current node and not the custom node
                        If record!node_id <> RemoveK(nodx.key) And _
                                record!node_id <> mnuRCPopupTree1AddNode(1).tag Then
                            Load mnuRCPopupTree1AddNode(menuindex)
                            mnuRCPopupTree1AddNode(menuindex).Caption = record!node_desc
                            mnuRCPopupTree1AddNode(menuindex).tag = record!node_id
                            mnuRCPopupTree1AddNode(menuindex).Visible = True
                            menuindex = menuindex + 1
                        End If
                        record.MoveNext
                    Wend
                End If
                record.Close
            End If
        End If
        GoTo Done
    End If

    'load treeview add popup menu
    Set record = dbase.OpenRecordset("SELECT * FROM " & DB_NodeTable & " WHERE parent = " _
            & node_id & " ORDER BY Node_Desc", dbOpenDynaset)
    If Not record.EOF Then
        Load mnuRCPopupTree1AddNode(2)
        mnuRCPopupTree1AddNode(2).Caption = "-"
        ' this required since it loads initially as invisible
        mnuRCPopupTree1AddNode(2).Visible = True
        menuindex = 3
        While Not record.EOF
            Load mnuRCPopupTree1AddNode(menuindex)
            mnuRCPopupTree1AddNode(menuindex).Caption = record!node_desc
            mnuRCPopupTree1AddNode(menuindex).tag = record!node_id
            ' this required since it loads initially as invisible
            mnuRCPopupTree1AddNode(menuindex).Visible = True
            menuindex = menuindex + 1
            record.MoveNext
        Wend
    End If
    record.Close
        
    'load listview add popup menu
    Set record = dbase.OpenRecordset("SELECT * FROM " & DB_QuickAddItemsTable & " WHERE Parent_Node = " _
            & node_id & " ORDER BY Data_Label", dbOpenDynaset)
    If Not record.EOF Then
        Load mnuRCPopupList2AddItem(record_count)
        mnuRCPopupList2AddItem(record_count).Caption = "-"
        menuindex = record_count + 1
        While Not record.EOF
            Load mnuRCPopupList2AddItem(menuindex)
            mnuRCPopupList2AddItem(menuindex).Caption = record!data_label
            mnuRCPopupList2AddItem(menuindex).tag = record!icon_large & _
                    "," & record!icon_small & "," & record!data_type
            menuindex = menuindex + 1
            record.MoveNext
        Wend
    End If
    record.Close
    
    'load warrior global types
    n = FindANode("Link Sources", False, True)
    If n = -1 Then
        GoTo doneGlobalAdd
    End If
        
    node_id = FindANode(tvTreeView.Nodes(AddK(node_id)).text, True, False, n)
    If node_id = -1 Then
        GoTo doneGlobalAdd
    End If
    'load globals to treeview add popup menu
    Set record = dbase.OpenRecordset("SELECT * FROM " & DB_NodeTable & " WHERE parent = " _
            & node_id & " ORDER BY Node_Desc", dbOpenDynaset)
    While Not record.EOF
        Load mnuRCPopupTree1AddNode(menuindex)
        mnuRCPopupTree1AddNode(menuindex).Caption = "*" & record!node_desc
        mnuRCPopupTree1AddNode(menuindex).tag = -record!node_id
        ' this required since it loads initially as invisible
        mnuRCPopupTree1AddNode(menuindex).Visible = True
        menuindex = menuindex + 1
        record.MoveNext
    Wend
    record.Close
doneGlobalAdd:

Done:
End Sub

Private Sub InitTreeView()
    tvTreeView.SetFocus
    Set nodx = tvTreeView.Nodes(1).Root
    tvTreeView.SelectedItem = nodx
    nodx.Expanded = True
    lvNeedsRefresh = True
    tvTreeView_NodeClick nodx
End Sub

Private Sub InitListView()
    lvListView.View = Val(GetSetting(App.Title, "Settings", "ViewMode", "0"))
    mnuListViewMode(lvListView.View).Checked = True
    lblTitle(1).Caption = "Contents of '" & nodx.text & "'"
    
    Call lvListView.ColumnHeaders.Add(1, "C1", "Item", lvListView.Width / 7, lvwColumnLeft)
    Call lvListView.ColumnHeaders.Add(2, "C2", "Data ID", lvListView.Width / 7, lvwColumnLeft)
    Call lvListView.ColumnHeaders.Add(3, "C3", "Parent Node", lvListView.Width / 7, lvwColumnLeft)
    Call lvListView.ColumnHeaders.Add(4, "C4", "Large Icon", lvListView.Width / 7, lvwColumnLeft)
    Call lvListView.ColumnHeaders.Add(5, "C5", "Small Icon", lvListView.Width / 7, lvwColumnLeft)
    Call lvListView.ColumnHeaders.Add(6, "C6", "Data_Type", lvListView.Width / 7, lvwColumnLeft)
End Sub

Private Sub tvSwapNodes(A_Key As String, B_Key As String)
    Dim parent_node As Node
    Dim TempKey As String
    Dim AFirst As Boolean
    Dim record As Recordset
    
    If Not tvTreeView.Nodes(A_Key).Next Is Nothing Then
        If tvTreeView.Nodes(A_Key).Next.key = B_Key Then
            AFirst = True
        Else
            AFirst = False
        End If
    Else
        AFirst = False
    End If
    
    Set parent_node = tvTreeView.Nodes(A_Key).parent
    If parent_node.Children = 2 Then  'these are the only two siblings
        ClearNode A_Key
        ClearNode B_Key
        If AFirst Then  ' A is the first one
            LoadTreeBranch RemoveK(B_Key), tvwChild, 0
            LoadTreeBranch RemoveK(A_Key), tvwChild, 0
        Else  ' B is the first one
            LoadTreeBranch RemoveK(A_Key), tvwChild, 0
            LoadTreeBranch RemoveK(B_Key), tvwChild, 0
        End If
    Else
        If AFirst Then ' A is the first one
            'there is at least one more node before these two
            If Not tvTreeView.Nodes(A_Key).Previous Is Nothing Then
                TempKey = tvTreeView.Nodes(A_Key).Previous.key
                ClearNode A_Key
                ClearNode B_Key
                LoadTreeBranch RemoveK(B_Key), tvwNext, RemoveK(TempKey)
                LoadTreeBranch RemoveK(A_Key), tvwNext, RemoveK(B_Key)
            'there is at least one more node after these two
            Else
                TempKey = tvTreeView.Nodes(B_Key).Next.key
                ClearNode A_Key
                ClearNode B_Key
                LoadTreeBranch RemoveK(A_Key), tvwPrevious, RemoveK(TempKey)
                LoadTreeBranch RemoveK(B_Key), tvwPrevious, RemoveK(A_Key)
            End If
        Else
            'there is at least one more node before these two
            If Not tvTreeView.Nodes(B_Key).Previous Is Nothing Then
                TempKey = tvTreeView.Nodes(B_Key).Previous.key
                ClearNode A_Key
                ClearNode B_Key
                LoadTreeBranch RemoveK(A_Key), tvwNext, RemoveK(TempKey)
                LoadTreeBranch RemoveK(B_Key), tvwNext, RemoveK(A_Key)
            'there is at least one more node after these two
            Else
                TempKey = tvTreeView.Nodes(A_Key).Next.key
                ClearNode A_Key
                ClearNode B_Key
                LoadTreeBranch RemoveK(B_Key), tvwPrevious, RemoveK(TempKey)
                LoadTreeBranch RemoveK(A_Key), tvwPrevious, RemoveK(B_Key)
            End If
        End If
    End If
    
    'SetTVIcons
End Sub

Public Sub ClearNode(ByVal nodekey As String)
    Dim m As String
    Dim n As String

    If tvTreeView.Nodes(nodekey).Children > 0 Then
        n = tvTreeView.Nodes(nodekey).Child.key
        While n <> tvTreeView.Nodes(n).LastSibling.key
            m = tvTreeView.Nodes(n).Next.key
            ClearNode n
            n = m
        Wend
        ClearNode n
    End If
    
    tvAttribCol.Remove (nodekey)
    tvTreeView.Nodes.Remove (nodekey)   'remove node requested
End Sub
Public Function FindNodeByName(nodekey As String, SearchStr As String) As Long
    
    Dim n As String
    Dim result As Long
    
    If tvTreeView.Nodes(nodekey).text <> SearchStr Then
        If tvTreeView.Nodes(nodekey).Children > 0 Then
            n = tvTreeView.Nodes(nodekey).Child.key
            While n <> tvTreeView.Nodes(n).LastSibling.key
                result = FindNodeByName(n, SearchStr)
                If result <> -1 Then
                    FindNodeByName = result
                    Exit Function
                End If
                n = tvTreeView.Nodes(n).Next.key
            Wend
            result = FindNodeByName(n, SearchStr)
            If result <> -1 Then
                FindNodeByName = result
                Exit Function
            End If
        End If
        FindNodeByName = -1
    Else
        FindNodeByName = RemoveK(tvTreeView.Nodes(nodekey).key)
    End If
End Function


Public Function FindANode(SearchStr As String, SearchName As Boolean, _
        SearchGlobalType As Boolean, Optional ParentNodeId As Long = -1, _
        Optional Err_MsgBox As Boolean = True) As Long

    Dim record As Recordset
    
    'if both search flags are FALSE, or both are TRUE display error and quit
    If Not (SearchName Xor SearchGlobalType) Then
        If Err_MsgBox Then
            MsgBox "FindANode:  You must search for only one of 'Name' or 'Global Type'", _
                vbExclamation
        End If
        FindANode = -1
        Exit Function
    Else
        If SearchName Then 'searching by Name
            If ParentNodeId = -1 Then
                If Err_MsgBox Then
                    MsgBox "FindANode:  ParentNodeID is NOT optional if you select Search by Name!", _
                        vbExclamation
                End If
                FindANode = -1
                Exit Function
            End If
            FindANode = FindNodeByName(AddK(ParentNodeId), SearchStr)
        Else 'searching by Global_Type
            Set record = dbase.OpenRecordset("SELECT Node_ID,Global_Type" _
                & " FROM " & DB_NodeTable & " WHERE Global_Type = '" _
                & SearchStr & "'", dbOpenDynaset)
            If record.EOF Then
                If Err_MsgBox Then
                    MsgBox "The Node_Desc '" & SearchStr & _
                        "' was not found in the '" & DB_NodeTable & "' table.", _
                        vbCritical
                End If
                FindANode = -1
                record.Close
                GoTo Done
            End If
            FindANode = record!node_id
            record.Close
        End If
    End If
Done:
End Function

Private Function FindQuickTypeRootNode(n As Integer) As Long
    Dim X As Long
        
        X = tvTreeView.Nodes(AddK(QuickAddNodesNodeID)).Child.Index
        While X <> tvTreeView.Nodes(X).LastSibling.Index
            If tvTreeView.Nodes(X).text = tvTreeView.Nodes(n).text Then
                GoTo doneNodeSearch
            End If
            ' Set x to the next node's index.
            X = tvTreeView.Nodes(X).Next.Index
        Wend
doneNodeSearch:
        If tvTreeView.Nodes(X).text <> tvTreeView.Nodes(n).text Then
            Dim oldnodx As Node
            Set oldnodx = nodx
            Set nodx = tvTreeView.Nodes(AddK(QuickAddNodesNodeID))
            FindQuickTypeRootNode = FillTree(n, False, True, True)
            Set nodx = oldnodx
            MsgBox "The QuickAdd Type, " & tvTreeView.Nodes(n).text & _
                ", did not exist and was added.", vbInformation
        Else
            FindQuickTypeRootNode = RemoveK(tvTreeView.Nodes(X).key)
        End If
        
End Function

Public Sub InitToolbar(Optional UseProgressBar As Boolean = False)
    If UseProgressBar Then
        InitProgressBar fProgForm, "Initializing Toolbars . . .", 0, 100, _
            LargeIconFolder & "Interface.bmp", False, InFormLoad
    End If
    On Error Resume Next
    fMainForm.tbStand.Buttons("Open").Image = fMainForm.imlMenu.ListImages.item("Open").Index
    fMainForm.tbStand.Buttons("Print").Image = fMainForm.imlMenu.ListImages.item("Print").Index
    fMainForm.tbStand.Buttons("Help").Image = fMainForm.imlMenu.ListImages.item("Help").Index
    fMainForm.tbStand.Buttons("Prop").Image = fMainForm.imlMenu.ListImages.item("Prop").Index
    fMainForm.tbStand.Buttons("Find").Image = fMainForm.imlMenu.ListImages.item("Prop").Index
    fMainForm.tbStand.Buttons("Replace").Image = fMainForm.imlMenu.ListImages.item("Prop").Index
    fMainForm.tbStand.Buttons("Refresh").Image = fMainForm.imlMenu.ListImages.item("Prop").Index
    fMainForm.tbStand.Buttons("Help").Image = fMainForm.imlMenu.ListImages.item("Prop").Index
    
    fMainForm.tbEdit.Buttons("Copy").Image = fMainForm.imlMenu.ListImages.item("Copy").Index
    fMainForm.tbEdit.Buttons("Cut").Image = fMainForm.imlMenu.ListImages.item("Cut").Index
    fMainForm.tbEdit.Buttons("Paste").Image = fMainForm.imlMenu.ListImages.item("Paste").Index
    fMainForm.tbEdit.Buttons("Delete").Image = fMainForm.imlMenu.ListImages.item("Delete").Index
    fMainForm.tbEdit.Buttons("Add").Image = fMainForm.imlMenu.ListImages.item("Security").Index
    fMainForm.tbEdit.Buttons("Rename").Image = fMainForm.imlMenu.ListImages.item("Security").Index
    fMainForm.tbEdit.Buttons("MoveUp").Image = fMainForm.imlMenu.ListImages.item("Security").Index
    fMainForm.tbEdit.Buttons("MoveDown").Image = fMainForm.imlMenu.ListImages.item("Security").Index
    fMainForm.tbEdit.Buttons("Variation").Image = fMainForm.imlMenu.ListImages.item("Security").Index
    fMainForm.tbEdit.Buttons("Execute").Image = fMainForm.imlMenu.ListImages.item("Security").Index
    
    fMainForm.tbExternal.Buttons("HTML").Image = fMainForm.imlMenu.ListImages.item("Security").Index
    fMainForm.tbExternal.Buttons("Source").Image = fMainForm.imlMenu.ListImages.item("Security").Index
    
    fMainForm.tbUtilities.Buttons("DB").Image = fMainForm.imlMenu.ListImages.item("Security").Index
    fMainForm.tbUtilities.Buttons("Spelling").Image = fMainForm.imlMenu.ListImages.item("Security").Index
    fMainForm.tbUtilities.Buttons("Security").Image = fMainForm.imlMenu.ListImages.item("Security").Index
    On Error GoTo 0
    
    If UseProgressBar Then
        fProgForm.Hide
    End If
End Sub


Public Sub SetTVIcons()
    
    tvTreeView.ImageList = Nothing
    'setup temp tree icon image list
    imlTempTree.ListImages.Clear
    imlTempTree.ImageHeight = SMALL_ICON_SIZE
    imlTempTree.ImageWidth = SMALL_ICON_SIZE
    On Error Resume Next
    imlTempTree.ListImages.Add , "K" & Link_Icon, _
        imlIconsSmall.ListImages("K" & Link_Icon).picture
    imlTempTree.ListImages.Add , "K" & Read_Icon, _
        imlIconsSmall.ListImages("K" & Read_Icon).picture
    imlTempTree.ListImages.Add , "K" & System_Icon, _
        imlIconsSmall.ListImages("K" & System_Icon).picture
    On Error GoTo 0
    SetTreeOverlays

End Sub

Private Function GetTempFolder(ShowMsg As Boolean) As Boolean
    CurTempFolder = GetSetting(App.EXEName, "Options", "Temp Folder", "UNKNOWN!")
    If Not ValidFolder(CurTempFolder) Then
        If ShowMsg Then
            MsgBox "The Temp Folder '" & CurTempFolder & "' does not exist!  " _
                & "You must choose a new one.", vbCritical
        End If
        BrowseDir.Título = "Choose a New Temp Folder..." 'set the title of the folder browser
        BrowseDir.Mostrar 'shows the folder browser
        If BrowseDir.DiretorioRetornado = "" Then 'wil be "" if user hit cancel
            GetTempFolder = False
            Exit Function
        Else
            CurTempFolder = BrowseDir.DiretorioRetornado & "\"
            SaveSetting App.EXEName, "Options", "Temp Folder", _
                CurTempFolder
        End If
    End If
    GetTempFolder = True
End Function
Private Function GetLargeIconFolder(ShowMsg As Boolean) As Boolean
    LargeIconFolder = GetSetting(App.EXEName, "Options", "Large Icon Folder", "UNKNOWN!")
    If Not ValidFolder(LargeIconFolder) Then
        If ShowMsg Then
            MsgBox "The Large Icon Folder '" & LargeIconFolder & "' does not exist!  " _
                & "You must choose a new one.", vbCritical
        End If
        BrowseDir.Título = "Choose a New Large Icon Folder..." 'set the title of the folder browser
        BrowseDir.Mostrar 'shows the folder browser
        If BrowseDir.DiretorioRetornado = "" Then 'wil be "" if user hit cancel
            GetLargeIconFolder = False
            Exit Function
        Else
            LargeIconFolder = BrowseDir.DiretorioRetornado & "\"
            SaveSetting App.EXEName, "Options", "Large Icon Folder", _
                LargeIconFolder
        End If
    End If
    GetLargeIconFolder = True
End Function

Private Function GetsmallIconFolder(ShowMsg As Boolean) As Boolean
    SmallIconFolder = GetSetting(App.EXEName, "Options", "small Icon Folder", "UNKNOWN!")
    If Not ValidFolder(SmallIconFolder) Then
        If ShowMsg Then
            MsgBox "The small Icon Folder '" & SmallIconFolder & "' does not exist!  " _
                & "You must choose a new one.", vbCritical
        End If
        BrowseDir.Título = "Choose a New small Icon Folder..." 'set the title of the folder browser
        BrowseDir.Mostrar 'shows the folder browser
        If BrowseDir.DiretorioRetornado = "" Then 'wil be "" if user hit cancel
            GetsmallIconFolder = False
            Exit Function
        Else
            SmallIconFolder = BrowseDir.DiretorioRetornado & "\"
            SaveSetting App.EXEName, "Options", "small Icon Folder", _
                SmallIconFolder
        End If
    End If
    GetsmallIconFolder = True
End Function
Private Function GetCurrentDatabase(ShowMsg As Boolean) As Boolean
    CurrentDatabaseFile = GetSetting(App.EXEName, "Options", "Startup Database", "UNKNOWN!")
    If Not ValidFile(CurrentDatabaseFile) Then
        If ShowMsg Then
            MsgBox "The Startup Database '" & CurrentDatabaseFile & "' does not exist!  " _
                & "You must choose a new one.", vbCritical
        End If
        CommonDialog1.Filter = "Access Database (*.mdb)|*.mdb"
        CommonDialog1.FilterIndex = 0
        On Error GoTo userCancel
            CommonDialog1.ShowOpen
        On Error GoTo 0
        CurrentDatabaseFile = CommonDialog1.filename
        SaveSetting App.EXEName, "Options", "Startup Database", _
                CurrentDatabaseFile
        GoTo Done
userCancel:
        On Error GoTo 0
        CurrentDatabaseFile = False
        Exit Function
    End If
Done:
    GetCurrentDatabase = True
End Function

'----------------------------------------------------------------------------
'*** Recursive ***
'Recursively scans a node branch for requested attributes.  Stops as soon as it
'finds a single node/list item that matches.
'REQUIRES:  StartNodeIndex - the index of the parent node of the branch.
'           tempstr - hold the message displaying what was found.  Can be
'                       passed as a parameter, so the calling function can
'                       add its own message, too.
'           read_only - whether to stop on nodes/items that have read_only=TRUE
'           system_node - whether to stop on nodes/items that have system_node=TRUE
'           list_items - whether to check if list items have read_only=TRUE as well
'RETURNS:   TRUE - if at least one match was found
'           FALSE - if no matches were found.
'----------------------------------------------------------------------------
Private Function ScanBranchForAttrib(StartNodeIndex As Long, tempstr As String, _
        KeyWord As String, Optional read_only As Boolean = True, _
        Optional system_node As Boolean = True, _
        Optional list_items As Boolean = True) As Boolean

    Dim n As Long
    Dim listrecord As Recordset
    
    'initialize the start node
    n = StartNodeIndex
    
    If read_only Then
        'Check if the current node is Read Only
        
        If tvAttribCol(tvTreeView.Nodes(n).key).read_only = True Then
            If tempstr = "" Then
                tempstr = "Read Only (TRUE)"
            Else
                tempstr = tempstr & ", Read Only (TRUE)"
            End If
        End If
    End If
    
    If list_items Then
        'See if any list items are Read Only
        
        'SQL to pull out all the first list item that has the current node as a parent
        'AND is marked Read Only.  If there are none, then there were no matches.
        'If one is found, its name is displayed in the error message.
        Set listrecord = dbase.OpenRecordset("SELECT TOP 1 Data_Label,Parent_Node," _
                & "Read_Only FROM " _
                & tvAttribCol(tvTreeView.Nodes(n).key).table_name _
                & " WHERE Parent_Node = " & RemoveK(tvTreeView.Nodes(n).key) _
                & " AND Read_Only = TRUE", dbOpenDynaset)
        If Not listrecord.EOF Then
            If tempstr = "" Then
                MsgBox "You cannot " & KeyWord & " '" & tvTreeView.Nodes(n).text _
                        & "' because it has the following list item marked as " _
                        & "Read Only: '" & listrecord!data_label & "'.", vbExclamation
                listrecord.Close
                ScanBranchForAttrib = True
                GoTo Done
            Else
                tempstr = tempstr & "." & Chr(13) & "Also it has the following list item marked as " _
                        & "Read Only: '" & listrecord!data_label & "'."
            End If
        End If
        listrecord.Close
    End If
    
    'If there is an error message to display, show it here.
    If tempstr <> "" Then
        MsgBox "You cannot " & KeyWord & " '" & tvTreeView.Nodes(n).text & "' because it has the following " _
                & "attribute(s) set:  " & tempstr & ".", vbExclamation
        ScanBranchForAttrib = True
        GoTo Done
    End If
    
    'start the recursion to scan all nodes below this one.
    If tvTreeView.Nodes(StartNodeIndex).Children > 0 Then
        n = tvTreeView.Nodes(StartNodeIndex).Child.Index
        While n <> tvTreeView.Nodes(n).LastSibling.Index
            If ScanBranchForAttrib(n, "", KeyWord) Then
                ScanBranchForAttrib = True
                GoTo Done
            End If
            n = tvTreeView.Nodes(n).Next.Index
        Wend
        If ScanBranchForAttrib(n, "", KeyWord) Then
            ScanBranchForAttrib = True
        End If
    End If
Done:
End Function

Private Sub FillSelectedItemsBuffer()
    Dim i As Integer
    Dim buffer_index As Long
    Dim BufferItemLabel As String
    
    For i = 1 To SelectedItemsBuffer.count
        SelectedItemsBuffer.Remove "K" & i
    Next i
    
    buffer_index = 0
    
    For i = 1 To lvListView.ListItems.count
        If lvListView.ListItems.item(i).Selected = True Then
            buffer_index = buffer_index + 1
            BufferItemLabel = lvListView.ListItems.item(i).text
            SelectedItemsBuffer.Add BufferItemLabel, "K" & Format$(buffer_index)
        End If
    Next i
    
End Sub

Public Function ChildNodeHasSameName(ParentKey As String, SearchName As String) _
        As Boolean
        
    Dim num_children As Long
    Dim i As Long
    
    ChildNodeHasSameName = False
    num_children = tvTreeView.Nodes(ParentKey).Children
    If num_children > 0 Then
        i = tvTreeView.Nodes(ParentKey).Child.Index
        
        While i <> tvTreeView.Nodes(i).LastSibling.Index
            If tvTreeView.Nodes(i).text = SearchName Then
                ChildNodeHasSameName = True
                GoTo Done
            End If
            i = tvTreeView.Nodes(i).Next.Index
        Wend
        
        If tvTreeView.Nodes(i).text = SearchName Then
            ChildNodeHasSameName = True
        End If
    End If
    
Done:

End Function


