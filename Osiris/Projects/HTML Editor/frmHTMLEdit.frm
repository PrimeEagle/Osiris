VERSION 5.00
Object = "{683364A1-B37D-11D1-ADC5-006008A5848C}#1.0#0"; "DHTMLEd.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmHTMLEdit 
   Caption         =   "Osiris HTML Editor"
   ClientHeight    =   6345
   ClientLeft      =   165
   ClientTop       =   645
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList imlMenu 
      Left            =   8640
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   1296
      BandCount       =   7
      _CBWidth        =   9375
      _CBHeight       =   735
      _Version        =   "6.0.8169"
      Child1          =   "cboFontName"
      MinHeight1      =   315
      Width1          =   1815
      NewRow1         =   0   'False
      Child2          =   "cboFontSize"
      MinHeight2      =   315
      Width2          =   810
      NewRow2         =   -1  'True
      Child3          =   "cboStyle"
      MinHeight3      =   315
      Width3          =   1815
      NewRow3         =   0   'False
      Child4          =   "tbEdit"
      MinHeight4      =   330
      Width4          =   4170
      NewRow4         =   0   'False
      Child5          =   "tbFormat"
      MinHeight5      =   330
      Width5          =   4530
      NewRow5         =   0   'False
      Child6          =   "tbMisc"
      MinHeight6      =   330
      Width6          =   5010
      NewRow6         =   0   'False
      Child7          =   "tbTable"
      MinHeight7      =   330
      Width7          =   3210
      NewRow7         =   0   'False
      Begin VB.ComboBox cboFontSize 
         Height          =   315
         Left            =   165
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   375
         Width           =   615
      End
      Begin VB.ComboBox cboFontName 
         Height          =   315
         Left            =   165
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   30
         Width           =   9120
      End
      Begin VB.ComboBox cboStyle 
         Height          =   315
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   375
         Width           =   1620
      End
      Begin MSComctlLib.Toolbar tbTable 
         Height          =   330
         Left            =   9345
         TabIndex        =   6
         Top             =   375
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Insert Table"
               Description     =   "Insert Table"
               Object.ToolTipText     =   "Insert Table"
               Object.Tag             =   "Insert Table"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Insert Row"
               Description     =   "Insert Row"
               Object.ToolTipText     =   "Insert Row"
               Object.Tag             =   "Insert Row"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Insert Column"
               Description     =   "Insert Column"
               Object.ToolTipText     =   "Insert Column"
               Object.Tag             =   "Insert Column"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Insert Cell"
               Description     =   "Insert Cell"
               Object.ToolTipText     =   "Insert Cell"
               Object.Tag             =   "Insert Cell"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Delete Row"
               Description     =   "Delete Row"
               Object.ToolTipText     =   "Delete Row"
               Object.Tag             =   "Delete Row"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Delete Column"
               Description     =   "Delete Column"
               Object.ToolTipText     =   "Delete Column"
               Object.Tag             =   "Delete Column"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Delete Cell"
               Description     =   "Delete Cell"
               Object.ToolTipText     =   "Delete Cell"
               Object.Tag             =   "Delete Cell"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Merge Cells"
               Description     =   "Merge Cells"
               Object.ToolTipText     =   "Merge Cells"
               Object.Tag             =   "Merge Cells"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split Cells"
               Description     =   "Split Cells"
               Object.ToolTipText     =   "Split Cells"
               Object.Tag             =   "Split Cells"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbFormat 
         Height          =   330
         Left            =   7050
         TabIndex        =   5
         Top             =   375
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   16
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bold"
               Description     =   "Bold"
               Object.ToolTipText     =   "Bold"
               Object.Tag             =   "Bold"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Italic"
               Description     =   "Italic"
               Object.ToolTipText     =   "Italic"
               Object.Tag             =   "Italic"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Underline"
               Description     =   "Underline"
               Object.ToolTipText     =   "Underline"
               Object.Tag             =   "Underline"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FGColor"
               Description     =   "FGColor"
               Object.ToolTipText     =   "Foreground Color"
               Object.Tag             =   "FGColor"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "BGColor"
               Description     =   "BGColor"
               Object.ToolTipText     =   "Background Color"
               Object.Tag             =   "BGColor"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Numbers"
               Description     =   "Numbers"
               Object.ToolTipText     =   "Numbers"
               Object.Tag             =   "Numbers"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bullets"
               Description     =   "Bullets"
               Object.ToolTipText     =   "Bullets"
               Object.Tag             =   "Bullets"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Decrease Indent"
               Description     =   "Decrease Indent"
               Object.ToolTipText     =   "Decrease Indent"
               Object.Tag             =   "Decrease Indent"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Increase Indent"
               Description     =   "Increase Indent"
               Object.ToolTipText     =   "Increase Indent"
               Object.Tag             =   "Increase Indent"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "LeftJustify"
               Description     =   "LeftJustify"
               Object.ToolTipText     =   "LeftJustify"
               Object.Tag             =   "LeftJustify"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Center"
               Description     =   "Center"
               Object.ToolTipText     =   "Center"
               Object.Tag             =   "Center"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "RightJustify"
               Description     =   "RightJustify"
               Object.ToolTipText     =   "RightJustify"
               Object.Tag             =   "RightJustify"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Find"
               Description     =   "Find"
               Object.ToolTipText     =   "Find"
               Object.Tag             =   "Find"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbMisc 
         Height          =   330
         Left            =   9180
         TabIndex        =   4
         Top             =   375
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "View Source"
               Description     =   "View Source"
               Object.ToolTipText     =   "View Source"
               Object.Tag             =   "View Source"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Insert Image"
               Description     =   "Insert Image"
               Object.ToolTipText     =   "Insert Image"
               Object.Tag             =   "Insert Image"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Insert Link"
               Description     =   "Insert Link"
               Object.ToolTipText     =   "Insert Link"
               Object.Tag             =   "Insert Link"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Grid Options"
               Description     =   "Grid Options"
               Object.ToolTipText     =   "Grid Options"
               Object.Tag             =   "Grid Options"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Show Borders"
               Description     =   "Show Borders"
               Object.ToolTipText     =   "Show Borders"
               Object.Tag             =   "Show Borders"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Show Glyphs"
               Description     =   "Show Glyphs"
               Object.ToolTipText     =   "Show Glyphs"
               Object.Tag             =   "Show Glyphs"
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Make Absolute"
               Description     =   "Make Absolute"
               Object.ToolTipText     =   "Make Absolute"
               Object.Tag             =   "Make Absolute"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Lock"
               Description     =   "Lock"
               Object.ToolTipText     =   "Lock"
               Object.Tag             =   "Lock"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbEdit 
         Height          =   330
         Left            =   2850
         TabIndex        =   3
         Top             =   375
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "New"
               Description     =   "New"
               Object.ToolTipText     =   "New File"
               Object.Tag             =   "New"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Open"
               Description     =   "Open"
               Object.ToolTipText     =   "Open File"
               Object.Tag             =   "Open"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Save"
               Description     =   "Save"
               Object.ToolTipText     =   "Save File"
               Object.Tag             =   "Save"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Print"
               Description     =   "Print"
               Object.ToolTipText     =   "Print"
               Object.Tag             =   "Print"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cut"
               Description     =   "Cut"
               Object.ToolTipText     =   "Cut"
               Object.Tag             =   "Cut"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Copy"
               Description     =   "Copy"
               Object.ToolTipText     =   "Copy"
               Object.Tag             =   "Copy"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Paste"
               Description     =   "Paste"
               Object.ToolTipText     =   "Paste"
               Object.Tag             =   "Paste"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Undo"
               Description     =   "Undo"
               Object.ToolTipText     =   "Undo"
               Object.Tag             =   "Undo"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Redo"
               Description     =   "Redo"
               Object.ToolTipText     =   "Redo"
               Object.Tag             =   "Redo"
            EndProperty
         EndProperty
      End
   End
   Begin DHTMLEDLibCtl.DHTMLEdit HTMLEdit 
      Height          =   5175
      Left            =   60
      TabIndex        =   1
      Top             =   840
      Width           =   9225
      ActivateApplets =   0   'False
      ActivateActiveXControls=   0   'False
      ActivateDTCs    =   -1  'True
      ShowDetails     =   0   'False
      ShowBorders     =   0   'False
      Appearance      =   1
      Scrollbars      =   -1  'True
      ScrollbarAppearance=   1
      SourceCodePreservation=   -1  'True
      AbsoluteDropMode=   0   'False
      SnapToGrid      =   0   'False
      SnapToGridX     =   50
      SnapToGridY     =   50
      BrowseMode      =   0   'False
      UseDivOnCarriageReturn=   0   'False
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6090
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7805
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   2
            AutoSize        =   2
            TextSave        =   "10/22/98"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            TextSave        =   "11:16 AM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   4935
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu FileNewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
      End
      Begin VB.Menu FileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditSub 
         Caption         =   "Undo"
         Index           =   0
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditSub 
         Caption         =   "Redo"
         Index           =   1
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEditSub 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuEditSub 
         Caption         =   "Cut"
         Index           =   3
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditSub 
         Caption         =   "Copy"
         Index           =   4
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditSub 
         Caption         =   "Paste"
         Index           =   5
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditSub 
         Caption         =   "Select All"
         Index           =   6
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditSub 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuEditSub 
         Caption         =   "Find Text"
         Index           =   8
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewSub 
         Caption         =   "Show &Borders"
         Index           =   0
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "Show Gl&yphs"
         Index           =   1
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "Snap To &Grid..."
         Index           =   3
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "Bro&wse Mode"
         Index           =   5
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "&Edit Mode"
         Index           =   6
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuViewSub 
         Caption         =   "Document &Source"
         Index           =   8
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuInsertSub 
         Caption         =   "&Image..."
         Index           =   0
      End
      Begin VB.Menu mnuInsertSub 
         Caption         =   "&Anchor..."
         Index           =   1
      End
      Begin VB.Menu mnuInsertButton 
         Caption         =   "&Button"
      End
      Begin VB.Menu mnuInsertHTML 
         Caption         =   "&HTML Tag..."
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "&Format"
      Begin VB.Menu mnuFormatSub 
         Caption         =   "Fo&nt..."
         Index           =   0
      End
      Begin VB.Menu mnuFormatSub 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFormatSub 
         Caption         =   "D&elimiters"
         Index           =   2
         Begin VB.Menu mnuDelimSub 
            Caption         =   "&None"
            Index           =   0
         End
         Begin VB.Menu mnuDelimSub 
            Caption         =   "N&umbered List"
            Index           =   1
         End
         Begin VB.Menu mnuDelimSub 
            Caption         =   "&Bulleted List"
            Index           =   2
         End
      End
      Begin VB.Menu mnuFormatSub 
         Caption         =   "&Alignment"
         Index           =   3
         Begin VB.Menu mnuFormatAlignSub 
            Caption         =   "&Left Justify"
            Index           =   0
         End
         Begin VB.Menu mnuFormatAlignSub 
            Caption         =   "&Center"
            Index           =   1
         End
         Begin VB.Menu mnuFormatAlignSub 
            Caption         =   "&Right Justify"
            Index           =   2
         End
      End
      Begin VB.Menu mnuFormatSub 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuFormatSub 
         Caption         =   "&Decrease Indent"
         Index           =   5
      End
      Begin VB.Menu mnuFormatSub 
         Caption         =   "&Increase Indent"
         Index           =   6
      End
      Begin VB.Menu mnuFormatSub 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuFormatSub 
         Caption         =   "&Foreground Color..."
         Index           =   8
      End
      Begin VB.Menu mnuFormatSub 
         Caption         =   "&Background Color..."
         Index           =   9
      End
   End
   Begin VB.Menu mnuOrder 
      Caption         =   "&Order"
      Begin VB.Menu mnuOrderSub 
         Caption         =   "Make Absolute"
         Index           =   0
      End
      Begin VB.Menu mnuOrderSub 
         Caption         =   "Bring To Front"
         Index           =   1
      End
      Begin VB.Menu mnuOrderSub 
         Caption         =   "Send To Back"
         Index           =   2
      End
      Begin VB.Menu mnuOrderSub 
         Caption         =   "Bring Forward"
         Index           =   3
      End
      Begin VB.Menu mnuOrderSub 
         Caption         =   "Send Back"
         Index           =   4
      End
      Begin VB.Menu mnuOrderSub 
         Caption         =   "Bring Above Text"
         Index           =   5
      End
      Begin VB.Menu mnuOrderSub 
         Caption         =   "Send Below Text"
         Index           =   6
      End
      Begin VB.Menu mnuOrderSub 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu mnuOrderSub 
         Caption         =   "Lock Element"
         Index           =   8
      End
   End
   Begin VB.Menu mnuTable 
      Caption         =   "T&able"
      Begin VB.Menu mnuTableSub 
         Caption         =   "Insert Table..."
         Index           =   0
      End
      Begin VB.Menu mnuTableSub 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuTableSub 
         Caption         =   "Insert Row"
         Index           =   2
      End
      Begin VB.Menu mnuTableSub 
         Caption         =   "Insert Column"
         Index           =   3
      End
      Begin VB.Menu mnuTableSub 
         Caption         =   "Insert Cell"
         Index           =   4
      End
      Begin VB.Menu mnuTableSub 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuTableSub 
         Caption         =   "Delete Rows"
         Index           =   6
      End
      Begin VB.Menu mnuTableSub 
         Caption         =   "Delete Columns"
         Index           =   7
      End
      Begin VB.Menu mnuTableSub 
         Caption         =   "Delete Cells"
         Index           =   8
      End
      Begin VB.Menu mnuTableSub 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuTableSub 
         Caption         =   "Merge Cells"
         Index           =   10
      End
      Begin VB.Menu mnuTableSub 
         Caption         =   "Split Cell"
         Index           =   11
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About Osiris HTML Editor..."
      End
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "&Help"
      End
   End
End
Attribute VB_Name = "frmHTMLEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Const MAX_ARGUMENTS = 10

Public LastSnap As Integer
Public LastX As Integer
Public LastY As Integer


Dim dbase As Database
Dim FromEvent As Boolean
Dim HTMLEditorInitialized As Boolean
Dim CurrentFile As String
Dim CanSave As Boolean
Dim FilePath
Dim CapableOfOrder As Boolean
Dim IsAbsPos As Boolean
Dim IsTable As Boolean
Dim ItemCountNormal As Long
Dim ItemCountOrder As Long
Dim ItemCountTable As Long
Dim DataItem As Recordset
Dim CmdType As String
Dim CurrentFileName As String
Dim CurTempFolder As String

Private Sub cboStyle_Click()
    Dim state As DHTMLEDITCMDF
    Dim format As String
    
    If Not FromEvent Then
        state = HTMLEdit.QueryStatus(DECMD_SETBLOCKFMT)
    
        If state >= DECMDF_ENABLED Then
            HTMLEdit.ExecCommand DECMD_SETBLOCKFMT, OLECMDEXECOPT_DONTPROMPTUSER, _
                cboStyle.List(cboStyle.ListIndex)
        End If
    End If
End Sub
    'Command Line Syntax:
    'HTMLEd [FILENAME] | -db [DBNAME] -table [TABLENAME]
    '       -searchfield [SEARCHFIELDNAME] -searchvalue [SEARCHVALUE]
    '       -loadfield [LOADFIELDNAME]
    '
    'nothing - load a blank form
    '
    'filename - load that filename
    '
    '-db..... - load requested field from the database
    '
    'anything else, including if both filename and db fields are given,
    'gives an error msgbox is shown, and a blank form is loaded.
    '
    'for a correctly passed command line, the arguments
    'in the CmdArray are as follows:
    '(1) "-db"
    '(2) database path
    '(3) "-table"
    '(4) table name
    '(5) "-searchfield"
    '(6) Field to search
    '(7) "-searchvalue"
    '(8) Value to search for in that field
    '(9) "-loadfield"
    '(10) Field to load editor data from
Private Sub Form_Load()
    Dim record As Recordset
    Dim filenum As Long
    Dim TempFileName As String
    Dim data_size As Long
    Dim tempstr As String
    Dim CmdArray() As String
   
    Me.Show
    
    Me.MousePointer = vbHourglass
    
    CmdArray() = GetCommandLine()
    CmdType = ParseCommandLine(CmdArray())
    

    LastSnap = 0
    LastX = 0
    LastY = 0
    FromEvent = False
    
    InitToolbars
    
    HTMLEditorInitialized = False
    Select Case CmdType
        Case "FILE"
            mnuFileSave.Enabled = False
            CurrentFileName = CmdArray(1)
            mnuFileSave.Caption = "&Save"
        Case "DATABASE"
            Set DataItem = dbase.OpenRecordset("SELECT * " _
                & " FROM " & CmdArray(4) & " WHERE " & CmdArray(6) _
                & " = " & val(CmdArray(8)), dbOpenDynaset)
            CurrentFile = DataItem!data_label
            If DataItem!read_only Then
                CanSave = False
            Else
                CanSave = True
            End If
            mnuFileSave.Caption = "&Save to '" & DataItem!data_label & "'"
        Case "BLANK"
            CurrentFile = "Unnamed Document"
            CanSave = True
            mnuFileSave.Caption = "&Save"
        Case Else
            CurrentFile = "Unnamed Document"
            CanSave = True
    End Select

    SetFormCaption
    
    InitFonts

    If CmdType = "DATABASE" Then
        data_size = DataItem("data_value").FieldSize
        If data_size <= 0 Then
            GoTo Done
        End If

        tempstr = DataItem!data_type
        
        MsgBox "Chang CurTempFolder here when Registry is cleaned up, " _
            & "so it isn't hard-coded.", vbInformation
        
        CurTempFolder = "C:\Osiris\Temp"
        
        TempFileName = CurTempFolder & tempstr & "Load." & tempstr
        WriteMemo record, "data_value", TempFileName

        If Not ValidFile(TempFileName) Then
            MsgBox "The necessary temp file '" & TempFileName _
                & "' was not created successfully." _
                & Chr$(13) & "Loading of application failed.", vbCritical
            record.Close
            dbase.Close
            Me.MousePointer = vbArrow
            End
        End If

        On Error GoTo HTMLError
        HTMLEdit.LoadDocument TempFileName, False
        On Error GoTo 0

        GoTo Done
HTMLError:
        MsgBox "Form_Load:  Load unsuccessful!" & vbCr & _
            "Error number: " & Err.Number & _
            vbCr & Err.Description
        record.Close
        dbase.Close
        Me.MousePointer = vbArrow
        End
    End If
Done:
    tbEdit.Buttons("Save").Enabled = False
    mnuFileSave.Enabled = False
    mnuFileSaveAs.Enabled = False
    mnuViewSub(5).Checked = False 'browse mode
    mnuViewSub(6).Checked = True 'edit mode
    Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckSaveStatus
End Sub

Private Sub HTMLEdit_ShowContextMenu(ByVal xPos As Long, ByVal yPos As Long)
    
    Dim cmdState As DHTMLEDITCMDF
    Dim strings() As String
    Dim states() As OLE_TRISTATE
    
    CapableOfOrder = False
    IsAbsPos = False
    IsTable = False
        
    ' Determine if the selected element is Order capable
    cmdState = HTMLEdit.QueryStatus(DECMD_MAKE_ABSOLUTE)
    If cmdState >= DECMDF_ENABLED Then
        CapableOfOrder = True
    End If
    
    'Use DECMD_SEND_TO_BACK to determine if this element is abs positioned
    cmdState = HTMLEdit.QueryStatus(DECMD_SEND_TO_BACK)
    If cmdState >= DECMDF_ENABLED Then
        IsAbsPos = True
    End If
    
    'Use DECMD_INSERTROW to determine if this element is a table
    cmdState = HTMLEdit.QueryStatus(DECMD_INSERTROW)
    If cmdState >= DECMDF_ENABLED Then
        IsTable = True
    End If
    
    ItemCountNormal = 6
    
    If CapableOfOrder Then
        ItemCountOrder = 2 '1 Item + 1 Separator
    Else
        ItemCountOrder = 0
    End If
    
    
    If IsTable Then
        ItemCountTable = 6 '4 Items + 2 Separators
    Else
        ItemCountTable = 0
    End If
    
    
    ReDim strings(0 To ItemCountNormal + ItemCountOrder + ItemCountTable)
    ReDim states(0 To ItemCountNormal + ItemCountOrder + ItemCountTable)
    
    strings(0) = "Cut"
    strings(1) = "Copy"
    strings(2) = "Paste"
    strings(3) = ""
    strings(4) = "Select All"
    strings(5) = ""
    strings(6) = "Font..."
        
    cmdState = HTMLEdit.QueryStatus(DECMD_CUT)
    If cmdState >= DECMDF_ENABLED Then
         states(0) = Unchecked
     Else
         states(0) = Gray
    End If
    
    cmdState = HTMLEdit.QueryStatus(DECMD_COPY)
    If cmdState >= DECMDF_ENABLED Then
         states(1) = Unchecked
     Else
         states(1) = Gray
    End If
    
    cmdState = HTMLEdit.QueryStatus(DECMD_PASTE)
    If cmdState >= DECMDF_ENABLED Then
         states(2) = Unchecked
     Else
         states(2) = Gray
    End If
        
    states(3) = Unchecked
    
    cmdState = HTMLEdit.QueryStatus(DECMD_SELECTALL)
    If cmdState >= DECMDF_ENABLED Then
         states(4) = Unchecked
     Else
         states(4) = Gray
    End If
    
    states(5) = Unchecked
    
    cmdState = HTMLEdit.QueryStatus(DECMD_FONT)
    If cmdState >= DECMDF_ENABLED Then
         states(6) = Unchecked
     Else
         states(6) = Gray
    End If
    
    If CapableOfOrder Then
        strings(ItemCountNormal + 1) = ""
        states(ItemCountNormal + 1) = Unchecked
        If IsAbsPos Then
            strings(ItemCountNormal + 2) = "Make 1D"
        Else
            strings(ItemCountNormal + 2) = "Make 2D"
        End If
        states(ItemCountNormal + 2) = Unchecked
    End If
    
    If IsTable Then
        strings(ItemCountNormal + ItemCountOrder + 1) = ""
        states(ItemCountNormal + ItemCountOrder + 1) = Unchecked
        strings(ItemCountNormal + ItemCountOrder + 2) = "Insert Row"
        states(ItemCountNormal + ItemCountOrder + 2) = Unchecked
        strings(ItemCountNormal + ItemCountOrder + 3) = "Insert Column"
        states(ItemCountNormal + ItemCountOrder + 3) = Unchecked
        strings(ItemCountNormal + ItemCountOrder + 4) = ""
        states(ItemCountNormal + ItemCountOrder + 4) = Unchecked
        strings(ItemCountNormal + ItemCountOrder + 5) = "Delete Row"
        states(ItemCountNormal + ItemCountOrder + 5) = Unchecked
        strings(ItemCountNormal + ItemCountOrder + 6) = "Delete Column"
        states(ItemCountNormal + ItemCountOrder + 6) = Unchecked
        
    End If
    HTMLEdit.SetContextMenu strings, states

End Sub

Private Sub mnuDelimSub_Click(Index As Integer)
    Select Case Index
        Case 0
            HTMLEdit.ExecCommand DECMD_SETBLOCKFMT, OLECMDEXECOPT_DONTPROMPTUSER, "Normal"
        Case 1
            HTMLEdit.ExecCommand DECMD_ORDERLIST, OLECMDEXECOPT_DONTPROMPTUSER
        Case 2
            HTMLEdit.ExecCommand DECMD_UNORDERLIST, OLECMDEXECOPT_DONTPROMPTUSER
    End Select
End Sub

Private Sub mnuFilePrint_Click()
    CommonDialog1.Flags = 0
    CommonDialog1.CancelError = True
    On Error GoTo Done
    CommonDialog1.ShowPrinter
    HTMLEdit.PrintDocument True
    MsgBox "TEMP:  Code for printing will go here.", vbInformation
Done:
End Sub

Private Sub mnuFormatAlignSub_Click(Index As Integer)
    Select Case Index
        Case 0
            HTMLEdit.ExecCommand DECMD_JUSTIFYLEFT, OLECMDEXECOPT_DONTPROMPTUSER
        Case 1
            HTMLEdit.ExecCommand DECMD_JUSTIFYCENTER, OLECMDEXECOPT_DONTPROMPTUSER
        Case 2
            HTMLEdit.ExecCommand DECMD_JUSTIFYRIGHT, OLECMDEXECOPT_DONTPROMPTUSER
    End Select
End Sub

Private Sub mnuFormatSub_Click(Index As Integer)
    Dim foreColor As String
    Dim bgColor As String
    
    Select Case Index
        Case 0
            HTMLEdit.ExecCommand DECMD_FONT, OLECMDEXECOPT_PROMPTUSER
        Case 5
            HTMLEdit.ExecCommand DECMD_OUTDENT, OLECMDEXECOPT_DONTPROMPTUSER
        Case 6
            HTMLEdit.ExecCommand DECMD_INDENT, OLECMDEXECOPT_DONTPROMPTUSER
        Case 8
            CommonDialog1.color = 0
            CommonDialog1.CancelError = True
            On Error GoTo Done
            CommonDialog1.ShowColor
            On Error GoTo 0
            foreColor = ""
            foreColor = FormatRGBString(CommonDialog1.color)
            HTMLEdit.ExecCommand DECMD_SETFORECOLOR, OLECMDEXECOPT_DONTPROMPTUSER, _
                foreColor
        Case 9
            CommonDialog1.color = 0
            CommonDialog1.CancelError = True
            On Error GoTo Done
            CommonDialog1.ShowColor
            On Error GoTo 0
            bgColor = ""
            bgColor = FormatRGBString(CommonDialog1.color)
            HTMLEdit.ExecCommand DECMD_SETBACKCOLOR, OLECMDEXECOPT_DONTPROMPTUSER, _
                bgColor
    End Select
Done:
End Sub

Private Sub mnuOrderSub_Click(Index As Integer)
    Dim command As DHTMLEDITCMDID
    
    Select Case Index
        Case 0
            command = DECMD_MAKE_ABSOLUTE
        Case 1
            command = DECMD_BRING_TO_FRONT
        Case 2
            command = DECMD_SEND_TO_BACK
        Case 3
            command = DECMD_BRING_FORWARD
        Case 4
            command = DECMD_SEND_BACKWARD
        Case 5
            command = DECMD_BRING_ABOVE_TEXT
        Case 6
            command = DECMD_SEND_BELOW_TEXT
        Case 7
            command = 0
        Case 8
            command = DECMD_LOCK_ELEMENT
    End Select
    
    If Not command = 0 Then 'seperator bars
        HTMLEdit.ExecCommand command, OLECMDEXECOPT_DODEFAULT
    End If

End Sub

Private Sub HTMLEdit_DocumentComplete()
    
    If Not HTMLEditorInitialized Then
        Dim format As DEGetBlockFmtNamesParam
        Dim i As Long
        Dim fontSize As Long
        Dim formatName As Variant
        
        ' Create the block format names holder
        Set format = CreateObject("DEGetBlockFmtNamesParam.DEGetBlockFmtNamesParam.1")
        
        ' Get the localized strings for the DECMD_SETBLOCKFMT command
        HTMLEdit.ExecCommand DECMD_GETBLOCKFMTNAMES, OLECMDEXECOPT_DONTPROMPTUSER, format
        
        ' Put the strings into the Format menu
        i = 0
        For Each formatName In format.Names
            cboStyle.AddItem formatName
            i = i + 1
        Next
        
        ' Get the current font size and update the combo
        HTMLEdit.ExecCommand DECMD_GETFONTSIZE, OLECMDEXECOPT_DONTPROMPTUSER, fontSize
        
        cboFontSize.ListIndex = fontSize - 1
    End If
    HTMLEditorInitialized = True
End Sub

Private Sub mnuEditSub_Click(Index As Integer)
    Dim command As DHTMLEDITCMDID
    
    Select Case Index
        Case 0
            command = DECMD_UNDO
        Case 1
            command = DECMD_REDO
        Case 2
            command = 0
        Case 3
            command = DECMD_CUT
        Case 4
            command = DECMD_COPY
        Case 5
            command = DECMD_PASTE
        Case 6
            command = DECMD_SELECTALL
        Case 7
            command = 0
        Case 8
            command = DECMD_FINDTEXT
    End Select
          
    If Not command = 0 Then
        HTMLEdit.ExecCommand command, OLECMDEXECOPT_DODEFAULT
    End If
End Sub

Private Sub mnuFileExit_Click()
    CheckSaveStatus
    Unload Me
End Sub

Private Sub mnuFileNew_Click()
    CheckSaveStatus
    FilePath = ""
    HTMLEdit.LoadDocument ""
    CanSave = True
    CurrentFile = "Unnamed Document"
    If CmdType <> "DATABASE" Then
        CmdType = "BLANK"
    End If
    SetFormCaption
End Sub

Private Sub mnuFileOpen_Click()
    CheckSaveStatus
    FilePath = ""
    HTMLEdit.LoadDocument "", True
    CurrentFile = HTMLEdit.CurrentDocumentPath
    If CmdType <> "DATABASE" Then
        CmdType = "FILE"
    End If
    SetFormCaption
End Sub

Private Sub mnuFileSave_Click()
    Dim record As Recordset
    Dim filenum As Long
    Dim tempstr As String
    Dim TempFileName As String

    Me.MousePointer = vbHourglass
    
    Select Case CmdType
        Case "FILE"
            SaveFile CurrentFileName, False, True
        Case "DATABASE"
            If HTMLEdit.DocumentHTML = "" Then
                record.Edit
                record!data_value = ""
                record.Update
            Else
                filenum = GetAvailableTempFile
                TempFileName = CurTempFolder & format$(filenum) & ".HTML"
                SaveFile TempFileName, False, False
                DoEvents
        
                If Not ValidFile(TempFileName) Then
                    MsgBox "The necessary temp file '" & TempFileName _
                        & "' was not created successfully!", vbCritical
                    GoTo SaveError
                End If
                ReadMemo TempFileName, record, "data_value"
                DeleteTempFile filenum
            End If
        Case Else
            MsgBox "ERROR:  Save menu should be when cmdType is not " _
                & "FILE or DATABASE!", vbCritical
    End Select
    
    MsgBox "HTML Data for '" & DataItem!data_label & "' was saved.", vbInformation
    GoTo Done
SaveError:
    MsgBox "HTML Data for '" & DataItem!data_label & "' was not successfully saved.", vbExclamation
Done:
    Me.MousePointer = vbArrow
End Sub

Private Sub mnuFileSaveAs_Click()
    SaveFile "", True, True
End Sub

Private Sub cboFontName_Click()
    Dim tempstr As String
    
    If Not FromEvent Then
        tempstr = cboFontName.List(cboFontName.ListIndex)
        If (HTMLEditorInitialized) Then
            HTMLEdit.ExecCommand DECMD_SETFONTNAME, OLECMDEXECOPT_DONTPROMPTUSER, tempstr
        End If
    End If
End Sub


Private Sub cboFontSize_Click()
    Dim templong As Long
    
    templong = cboFontSize.ListIndex + 1
    If (HTMLEditorInitialized) Then
        HTMLEdit.ExecCommand DECMD_SETFONTSIZE, OLECMDEXECOPT_DONTPROMPTUSER, templong
    End If
End Sub

Private Sub mnuInsert_Click()
    ScanInsertMenu
End Sub

Private Sub mnuInsertButton_Click()
    Dim doc As Object
    Dim selection As Object
    Dim tr As Object
    ' This routine inserts a button at the current selection
    
    ' Get the DHTML Document object
    Set doc = HTMLEdit.Document
    ' Get the DHTML Selection object
    Set selection = doc.selection
    ' Create a TextRange on the current selection
    Set tr = selection.createRange
    
    tr.pasteHTML ("<BUTTON TITLE=Button>Button!</B>")
End Sub

Private Sub mnuInsertHTML_Click()
    frmHTMLInsertHTML.Show vbModal, Me
End Sub

Private Sub mnuInsertSub_Click(Index As Integer)
    Dim command As DHTMLEDITCMDID
    
    Select Case Index
        Case 0
            command = DECMD_IMAGE
        Case 1
            command = DECMD_HYPERLINK
    End Select
    
    If Not command = 0 Then
        HTMLEdit.ExecCommand command, OLECMDEXECOPT_DODEFAULT
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    frmHTMLEditAbout.Show vbModal, Me
End Sub

Private Sub mnuTable_Click()
    ScanTableMenu
End Sub

Private Sub mnuOrder_Click()
    ScanOrderMenu
End Sub

Private Sub mnuTableSub_Click(Index As Integer)
    Dim command As DHTMLEDITCMDID
    
    If Index = 0 Then ' Insert Table
        frmHTMLInsertTable.Show vbModal, Me
        Exit Sub
    End If
    
    Select Case Index
        Case 2
            command = DECMD_INSERTROW
        Case 3
            command = DECMD_INSERTCOL
        Case 4
            command = DECMD_INSERTCELL
        Case 6
            command = DECMD_DELETEROWS
        Case 7
            command = DECMD_DELETECOLS
        Case 8
            command = DECMD_DELETECELLS
        Case 10
            command = DECMD_MERGECELLS
        Case 11
            command = DECMD_SPLITCELL
    End Select
    
    If Not command = 0 Then
        HTMLEdit.ExecCommand command, OLECMDEXECOPT_DODEFAULT
    End If
End Sub

Private Sub tbEdit_ButtonClick(ByVal Button As MSComctlLib.Button)
    ScanToolbars
    Select Case Button.Key
        Case "New"
            mnuFileNew_Click
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Cut"
            mnuEditSub_Click (3)
        Case "Copy"
            mnuEditSub_Click (4)
        Case "Paste"
            mnuEditSub_Click (5)
        Case "Undo"
            mnuEditSub_Click (0)
        Case "Redo"
            mnuEditSub_Click (1)
    End Select
End Sub

Private Sub tbFormat_ButtonClick(ByVal Button As Button)
    ScanToolbars
    Select Case Button.Key
        Case "Bold"
            HTMLEdit.ExecCommand DECMD_BOLD, OLECMDEXECOPT_DONTPROMPTUSER
        Case "Italic"
            HTMLEdit.ExecCommand DECMD_ITALIC, OLECMDEXECOPT_DONTPROMPTUSER
        Case "Underline"
            HTMLEdit.ExecCommand DECMD_UNDERLINE, OLECMDEXECOPT_DONTPROMPTUSER
        Case "Numbers"
            mnuDelimSub_Click (1)
        Case "Bullets"
            mnuDelimSub_Click (2)
        Case "Outdent"
            mnuFormatSub_Click (6)
        Case "Indent"
            mnuFormatSub_Click (5)
        Case "LeftJustify"
            mnuFormatAlignSub_Click (0)
        Case "Center"
            mnuFormatAlignSub_Click (1)
        Case "RightJustify"
            mnuFormatAlignSub_Click (2)
        Case "FGColor"
            mnuFormatSub_Click (8)
        Case "BGColor"
            mnuFormatSub_Click (9)
        End Select
End Sub

Private Sub HTMLEdit_ContextMenuAction(ByVal itemIndex As Long)

    ' Handle user selection on the custom context menu
   Select Case itemIndex
    Case 0
        HTMLEdit.ExecCommand DECMD_CUT, OLECMDEXECOPT_DODEFAULT
    Case 1
        HTMLEdit.ExecCommand DECMD_COPY, OLECMDEXECOPT_DODEFAULT
    Case 2
        HTMLEdit.ExecCommand DECMD_PASTE, OLECMDEXECOPT_DODEFAULT
    Case 4
        HTMLEdit.ExecCommand DECMD_SELECTALL, OLECMDEXECOPT_DODEFAULT
    Case 6
        HTMLEdit.ExecCommand DECMD_FONT, OLECMDEXECOPT_PROMPTUSER
    End Select
    
    If CapableOfOrder Then
        Select Case itemIndex
        Case ItemCountNormal + 2
            HTMLEdit.ExecCommand DECMD_MAKE_ABSOLUTE, OLECMDEXECOPT_DODEFAULT
        End Select
    End If
    
    If IsTable Then
        Select Case itemIndex
        Case ItemCountNormal + ItemCountOrder + 2
            HTMLEdit.ExecCommand DECMD_INSERTROW, OLECMDEXECOPT_DODEFAULT
        Case ItemCountNormal + ItemCountOrder + 3
            HTMLEdit.ExecCommand DECMD_INSERTCOL, OLECMDEXECOPT_DODEFAULT
        Case ItemCountNormal + ItemCountOrder + 5
            HTMLEdit.ExecCommand DECMD_DELETEROWS, OLECMDEXECOPT_DODEFAULT
        Case ItemCountNormal + ItemCountOrder + 6
            HTMLEdit.ExecCommand DECMD_DELETECOLS, OLECMDEXECOPT_DODEFAULT
        End Select
    End If
End Sub

Private Sub HTMLEdit_DisplayChanged()
    Dim Button As String
    Dim cmds As Long
    Dim blockFmt As String
    Dim fontSize As Long
    Dim state As DHTMLEDITCMDF
    
    FromEvent = True
    
    If CanSave And (mnuFileSave.Enabled = False Or _
            tbEdit.Buttons("Save").Enabled = False) Then
        If CmdType = "DATABASE" Then
            mnuFileSave.Enabled = True
            tbEdit.Buttons("Save").Enabled = True
        End If
    End If
    
    state = HTMLEdit.QueryStatus(DECMD_GETBLOCKFMT)
    If state >= DECMDF_ENABLED Then
        HTMLEdit.ExecCommand DECMD_GETBLOCKFMT, OLECMDEXECOPT_DONTPROMPTUSER, blockFmt
        If blockFmt <> "" Then
            cboStyle.text = blockFmt
        End If
    End If
    
    state = HTMLEdit.QueryStatus(DECMD_GETFONTNAME)
    If state >= DECMDF_ENABLED Then
        Dim fontName As String
        HTMLEdit.ExecCommand DECMD_GETFONTNAME, OLECMDEXECOPT_DONTPROMPTUSER, fontName
        If fontName <> "" Then
            cboFontName.text = fontName
        End If
    End If
    
    state = HTMLEdit.QueryStatus(DECMD_GETFONTNAME)
    If state >= DECMDF_ENABLED Then
        HTMLEdit.ExecCommand DECMD_GETFONTSIZE, OLECMDEXECOPT_DONTPROMPTUSER, fontSize
        If fontSize >= 1 Then
            cboFontSize.text = fontSize
        Else
            cboFontSize.text = ""
        End If
    End If
    
    ScanToolbars

    FromEvent = False
    
End Sub

Private Sub mnuEdit_Click()
    ScanEditMenu
'    Dim cmdIndex As Integer
    
'    For cmdIndex = 0 To NUM_EDIT_MENU_ITEMS
'        UpdateMenu mnuEditSub(cmdIndex), CmdEditMenu(cmdIndex)
'    Next cmdIndex
End Sub

Private Sub UpdateMenu(menu As Control, command As DHTMLEDITCMDID)
    Dim state As DHTMLEDITCMDF

    If Not command = 0 Then
        state = HTMLEdit.QueryStatus(command)
        
        If (state >= DECMDF_ENABLED) Then
            menu.Enabled = True
        Else
            menu.Enabled = False
        End If
    End If
End Sub

Private Sub mnuViewSub_Click(Index As Integer)
    Dim state As Boolean
    
    Select Case Index
        Case 0 ' Borders
            state = HTMLEdit.ShowBorders
            state = Not state
            HTMLEdit.ShowBorders = state
            mnuViewSub(Index).Checked = state
        Case 1 ' Show Details
            state = HTMLEdit.ShowDetails
            state = Not state
            HTMLEdit.ShowDetails = state
            mnuViewSub(Index).Checked = state
        Case 3 'Snap to Grid
            frmHTMLSnap.Show vbModal, Me
            HTMLEdit.SnapToGrid = LastSnap
            If LastSnap Then
                HTMLEdit.SnapToGridX = LastX
                HTMLEdit.SnapToGridY = LastY
            End If
        Case 5 'Browse Mode
            MsgBox "TEMP: Browse Mode", vbInformation
            mnuViewSub(5).Checked = True
            mnuViewSub(6).Checked = False
            HTMLEdit.BrowseMode = True
        Case 6 'Edit Mode
            MsgBox "TEMP: Edit Mode", vbInformation
            mnuViewSub(6).Checked = True
            mnuViewSub(5).Checked = False
            HTMLEdit.BrowseMode = False
        Case 8
            MsgBox "TEMP: Launch Source Viewer Here.", vbInformation
    End Select
End Sub

Private Sub SetFormCaption()
    If Len(CurrentFile) > 0 Then
        Me.Caption = "Osiris HTML Editor:      Editing '" & CurrentFile & "'"
    Else
        Me.Caption = "Osiris HTML Editor"
    End If
End Sub

Private Function FormatRGBString(val As Long) As String
    Dim color As String
    Dim pad As Long
    Dim r As String
    Dim g As String
    Dim b As String
    
    ' This function formats a long consisting of rgb values
    ' taken from the CommonDialog color dialog
    ' to a string in the form of "#RRGGBB" where RRGGBB are
    ' hex values
    
    ' convert to hex
    color = Hex(val)
    'determine how many zeros to pad in front of converted value
    pad = 6 - Len(color)
    
    If pad Then
        color = String(pad, "0") & color
    End If
        
    'Extract the rgb components
    r = Right(color, 2)
    g = Mid(color, 3, 2)
    b = Left(color, 2)
    
    ' Swab r and b position, color dialog returns
    ' bgr instead of rgb
    color = "#" & r & g & b
    
    FormatRGBString = color
End Function


Private Sub InitFonts()
    Dim i As Long
    
    For i = 0 To Screen.FontCount - 1
        cboFontName.AddItem Screen.Fonts(i)
    Next i
    cboFontName.ListIndex = 0
    
    For i = 0 To 6
        cboFontSize.AddItem format$(i + 1)
    Next i
    cboFontSize.ListIndex = 0
End Sub

Private Sub SaveFile(filename As String, ShowDialog As Boolean, _
        UpdateTitleBar As Boolean)
    
    Me.MousePointer = vbHourglass
    
    HTMLEdit.SaveDocument filename, ShowDialog
    
    If UpdateTitleBar Then
        CurrentFile = HTMLEdit.CurrentDocumentPath
        SetFormCaption
    End If
    
    Me.MousePointer = vbArrow
End Sub

Private Sub tbMisc_ButtonClick(ByVal Button As MSComctlLib.Button)
    ScanToolbars
    Select Case Button.Key
        Case "View Source"
            mnuViewSub_Click (8)
        Case "Insert Image"
            mnuInsertSub_Click (0)
        Case "Insert Link"
            mnuInsertSub_Click (1)
        Case "Grid Options"
            mnuViewSub_Click (3)
        Case "Show Borders"
            mnuViewSub_Click (0)
        Case "Show Glyphs"
            mnuViewSub_Click (1)
        Case "Make Absolute"
            mnuOrderSub_Click (0)
        Case "Lock"
            mnuOrderSub_Click (8)
    End Select
End Sub

Private Sub tbTable_ButtonClick(ByVal Button As MSComctlLib.Button)
    ScanToolbars
    Select Case Button.Key
        Case "Insert Table"
            mnuTableSub_Click (0)
        Case "Insert Row"
            mnuTableSub_Click (2)
        Case "Insert Column"
            mnuTableSub_Click (3)
        Case "Insert Cell"
            mnuTableSub_Click (4)
        Case "Delete Row"
            mnuTableSub_Click (6)
        Case "Delete Column"
            mnuTableSub_Click (7)
        Case "Delete Cell"
            mnuTableSub_Click (8)
        Case "Merge Cells"
            mnuTableSub_Click (10)
        Case "Split Cells"
            mnuTableSub_Click (11)
    End Select
End Sub


Private Sub SetEnableLatchTB(command As DHTMLEDITCMDID, Button As Button)
    Dim state As DHTMLEDITCMDF
    
    state = HTMLEdit.QueryStatus(command)
    
    If (state >= DECMDF_ENABLED) Then
        Button.Enabled = True
    Else
        Button.Enabled = False
    End If
            
    If (state = DECMDF_LATCHED) Then
        Button.Value = tbrPressed
    Else
        Button.Value = tbrUnpressed
    End If
End Sub

Private Sub ScanToolbars()
    SetEnableLatchTB DECMD_BOLD, tbFormat.Buttons("Bold")
    SetEnableLatchTB DECMD_ITALIC, tbFormat.Buttons("Italic")
    SetEnableLatchTB DECMD_UNDERLINE, tbFormat.Buttons("Underline")
    SetEnableLatchTB DECMD_SETBACKCOLOR, tbFormat.Buttons("BGColor")
    SetEnableLatchTB DECMD_SETFORECOLOR, tbFormat.Buttons("FGColor")
    SetEnableLatchTB DECMD_INDENT, tbFormat.Buttons("Increase Indent")
    SetEnableLatchTB DECMD_OUTDENT, tbFormat.Buttons("Decrease Indent")
    SetEnableLatchTB DECMD_ORDERLIST, tbFormat.Buttons("Numbers")
    SetEnableLatchTB DECMD_UNORDERLIST, tbFormat.Buttons("Bullets")
    SetEnableLatchTB DECMD_JUSTIFYLEFT, tbFormat.Buttons("LeftJustify")
    SetEnableLatchTB DECMD_JUSTIFYRIGHT, tbFormat.Buttons("RightJustify")
    SetEnableLatchTB DECMD_JUSTIFYCENTER, tbFormat.Buttons("Center")
    SetEnableLatchTB DECMD_HYPERLINK, tbMisc.Buttons("Insert Link")
    SetEnableLatchTB DECMD_IMAGE, tbMisc.Buttons("Insert Image")
    SetEnableLatchTB DECMD_INSERTTABLE, tbTable.Buttons("Insert Table")
    SetEnableLatchTB DECMD_INSERTROW, tbTable.Buttons("Insert Row")
    SetEnableLatchTB DECMD_INSERTCOL, tbTable.Buttons("Insert Column")
    SetEnableLatchTB DECMD_INSERTCELL, tbTable.Buttons("Insert Cell")
    SetEnableLatchTB DECMD_DELETEROWS, tbTable.Buttons("Delete Row")
    SetEnableLatchTB DECMD_DELETECOLS, tbTable.Buttons("Delete Column")
    SetEnableLatchTB DECMD_DELETECELLS, tbTable.Buttons("Delete Cell")
    SetEnableLatchTB DECMD_MERGECELLS, tbTable.Buttons("Merge Cells")
    SetEnableLatchTB DECMD_SPLITCELL, tbTable.Buttons("Split Cells")
End Sub

Private Sub SetEnableEditMenu(command As DHTMLEDITCMDID, item As Integer)
    Dim state As DHTMLEDITCMDF
    
    state = HTMLEdit.QueryStatus(command)
    
    If (state >= DECMDF_ENABLED) Then
        mnuEditSub(item).Enabled = True
    Else
        mnuEditSub(item).Enabled = False
    End If
End Sub
Private Sub SetEnableInsertMenu(command As DHTMLEDITCMDID, item As Integer)
    Dim state As DHTMLEDITCMDF
    
    state = HTMLEdit.QueryStatus(command)
    
    If (state >= DECMDF_ENABLED) Then
        mnuInsertSub(item).Enabled = True
    Else
        mnuInsertSub(item).Enabled = False
    End If
End Sub

Private Sub SetEnableTableMenu(command As DHTMLEDITCMDID, item As Integer)
    Dim state As DHTMLEDITCMDF
    
    state = HTMLEdit.QueryStatus(command)
    
    If (state >= DECMDF_ENABLED) Then
        mnuTableSub(item).Enabled = True
    Else
        mnuTableSub(item).Enabled = False
    End If
End Sub
Private Sub SetEnableOrderMenu(command As DHTMLEDITCMDID, item As Integer)
    Dim state As DHTMLEDITCMDF
    
    state = HTMLEdit.QueryStatus(command)
    
    If (state >= DECMDF_ENABLED) Then
        mnuOrderSub(item).Enabled = True
    Else
        mnuOrderSub(item).Enabled = False
    End If
End Sub

Private Sub ScanEditMenu()
    SetEnableEditMenu DECMD_UNDO, 0
    SetEnableEditMenu DECMD_REDO, 1
    SetEnableEditMenu DECMD_CUT, 3
    SetEnableEditMenu DECMD_COPY, 4
    SetEnableEditMenu DECMD_PASTE, 5
    SetEnableEditMenu DECMD_SELECTALL, 6
    SetEnableEditMenu DECMD_FINDTEXT, 8
End Sub

Private Sub ScanInsertMenu()
    SetEnableInsertMenu DECMD_IMAGE, 0
    SetEnableInsertMenu DECMD_HYPERLINK, 1
End Sub

Private Sub ScanTableMenu()
    SetEnableTableMenu DECMD_INSERTTABLE, 0
    SetEnableTableMenu DECMD_INSERTROW, 2
    SetEnableTableMenu DECMD_INSERTCOL, 3
    SetEnableTableMenu DECMD_INSERTCELL, 4
    SetEnableTableMenu DECMD_DELETEROWS, 6
    SetEnableTableMenu DECMD_DELETECOLS, 7
    SetEnableTableMenu DECMD_DELETECELLS, 8
    SetEnableTableMenu DECMD_MERGECELLS, 10
    SetEnableTableMenu DECMD_SPLITCELL, 11
End Sub

Private Sub ScanOrderMenu()
    SetEnableOrderMenu DECMD_MAKE_ABSOLUTE, 0
    SetEnableOrderMenu DECMD_BRING_TO_FRONT, 1
    SetEnableOrderMenu DECMD_SEND_TO_BACK, 2
    SetEnableOrderMenu DECMD_BRING_FORWARD, 3
    SetEnableOrderMenu DECMD_SEND_BACKWARD, 4
    SetEnableOrderMenu DECMD_BRING_ABOVE_TEXT, 5
    SetEnableOrderMenu DECMD_SEND_BELOW_TEXT, 6
    SetEnableOrderMenu DECMD_LOCK_ELEMENT, 8
End Sub

Private Sub InitToolbars()
    On Error Resume Next
    Set tbFormat.ImageList = imlMenu
    Set tbEdit.ImageList = imlMenu
    Set tbMisc.ImageList = imlMenu
    Set tbTable.ImageList = imlMenu
    
    tbEdit.Buttons("New").Image = imlMenu.ListImages.item("New").Index
    tbEdit.Buttons("Open").Image = imlMenu.ListImages.item("Open").Index
    tbEdit.Buttons("Save").Image = imlMenu.ListImages.item("Save").Index
    tbEdit.Buttons("Print").Image = imlMenu.ListImages.item("Print").Index
    tbEdit.Buttons("Cut").Image = imlMenu.ListImages.item("Cut").Index
    tbEdit.Buttons("Copy").Image = imlMenu.ListImages.item("Copy").Index
    tbEdit.Buttons("Paste").Image = imlMenu.ListImages.item("Paste").Index
    tbEdit.Buttons("Undo").Image = imlMenu.ListImages.item("Undo").Index
    tbEdit.Buttons("Redo").Image = imlMenu.ListImages.item("Redo").Index
    
    tbMisc.Buttons("View Source").Image = imlMenu.ListImages.item("Preview").Index
    tbMisc.Buttons("Insert Image").Image = imlMenu.ListImages.item("Image").Index
    tbMisc.Buttons("Insert Link").Image = imlMenu.ListImages.item("Link").Index
    tbMisc.Buttons("Grid Options").Image = imlMenu.ListImages.item("Snap Grid 16").Index
    tbMisc.Buttons("Show Borders").Image = imlMenu.ListImages.item("Borders").Index
    tbMisc.Buttons("Show Glyphs").Image = imlMenu.ListImages.item("Details").Index
    tbMisc.Buttons("Make Absolute").Image = imlMenu.ListImages.item("Absolute Pos").Index
    tbMisc.Buttons("Lock").Image = imlMenu.ListImages.item("Security").Index
    
    tbFormat.Buttons("Bold").Image = imlMenu.ListImages.item("Bold").Index
    tbFormat.Buttons("Italic").Image = imlMenu.ListImages.item("Italic").Index
    tbFormat.Buttons("Underline").Image = imlMenu.ListImages.item("Underline").Index
    tbFormat.Buttons("Numbers").Image = imlMenu.ListImages.item("Numbers").Index
    tbFormat.Buttons("Decrease Indent").Image = imlMenu.ListImages.item("Decrease Indent").Index
    tbFormat.Buttons("Increase Indent").Image = imlMenu.ListImages.item("Increase Indent").Index
    tbFormat.Buttons("LeftJustify").Image = imlMenu.ListImages.item("Left Just").Index
    tbFormat.Buttons("Center").Image = imlMenu.ListImages.item("Center Just").Index
    tbFormat.Buttons("RightJustify").Image = imlMenu.ListImages.item("Right Just").Index
    tbFormat.Buttons("Bullets").Image = imlMenu.ListImages.item("Bullets").Index
    tbFormat.Buttons("FGColor").Image = imlMenu.ListImages.item("FGColor").Index
    tbFormat.Buttons("BGColor").Image = imlMenu.ListImages.item("BGColor").Index
    tbFormat.Buttons("Find").Image = imlMenu.ListImages.item("Find").Index
    
    tbTable.Buttons("Insert Table").Image = imlMenu.ListImages.item("Insert Table").Index
    tbTable.Buttons("Insert Row").Image = imlMenu.ListImages.item("Insert Row").Index
    tbTable.Buttons("Insert Column").Image = imlMenu.ListImages.item("Insert Column").Index
    tbTable.Buttons("Insert Cell").Image = imlMenu.ListImages.item("Insert Cell").Index
    tbTable.Buttons("Delete Row").Image = imlMenu.ListImages.item("Delete Row").Index
    tbTable.Buttons("Delete Column").Image = imlMenu.ListImages.item("Delete Column").Index
    tbTable.Buttons("Delete Cell").Image = imlMenu.ListImages.item("Delete Cell").Index
    tbTable.Buttons("Merge Cells").Image = imlMenu.ListImages.item("Merge Cell").Index
    tbTable.Buttons("Split Cells").Image = imlMenu.ListImages.item("Split Cell 16").Index
    On Error GoTo 0
    
End Sub

Function GetCommandLine()   'Declare variables.
   Dim InArg As Boolean
   Dim NumArgs As Integer
   Dim i As Integer
   Dim c As String
   Dim CmdLine As String
   Dim CmdLineLen As Long
   
   ReDim ArgArray(MAX_ARGUMENTS) As String
   
   NumArgs = 0
   InArg = False
   
   'Get command line arguments.
   CmdLine = command()
   CmdLineLen = Len(CmdLine)
   
   'Go thru command line one character at a time.
   MsgBox "TEMP:  Need to check for parameters that have spaces in them, " _
    & "set off by quotes.", vbInformation
   
   For i = 1 To CmdLineLen
      c = Mid$(CmdLine, i, 1)      'Test for space or tab.
      If (c <> " " And c <> vbTab) Then         'Neither space nor tab.
         'Test if already in argument.
         If Not InArg Then
            'New argument begins.
            'Test for too many arguments.
            If NumArgs = MAX_ARGUMENTS Then
                MsgBox "Too many arguments were passed on the command line!  " _
                    & "The maximum allowed is " & MAX_ARGUMENTS & ".", vbExclamation
                    Exit For
            End If
            NumArgs = NumArgs + 1
            InArg = True
         End If
         'Concatenate character to current argument.
         ArgArray(NumArgs) = ArgArray(NumArgs) & c
      Else
         'Found a space or tab.
         'Set InArg flag to False.
         InArg = False
      End If
   Next i
   
   'Resize array just enough to hold arguments.
   ReDim Preserve ArgArray(NumArgs) As String
   'Return Array in Function name.
   GetCommandLine = ArgArray()
End Function

Private Function ParseCommandLine(cmdparam() As String) As String
    Dim cmd As String
    Dim test_flag As Boolean
    Dim table_index As Integer
    Dim test_index As Boolean
    Dim i As Integer
    Dim record As Recordset
    
    'get the command line arguments
    cmd = command()
    
    'no arguments
    If Len(cmd) = 0 Then
        ParseCommandLine = "BLANK"
        Exit Function
    End If
    
    'if the argument is a valid file then use it
    If ValidFile(cmd) Then
        ParseCommandLine = "FILE"
        Exit Function
    End If
    
    'otherwise, parse the arguments into an array
    cmdparam() = GetCommandLine
    
    If UCase(cmdparam(1)) <> "-DB" Then
        ParseCommandLine = "BLANK"
        Exit Function
    End If
    
    If Not ValidFile(cmdparam(2)) Then
        MsgBox "'" & cmdparam(2) _
            & "' is not a valid path or filename!", vbExclamation
        ParseCommandLine = "BLANK"
        Exit Function
    End If
    
    'we have a database, so open it.
    Set dbase = OpenDBase(cmdparam(2))
    If dbase Is Nothing Then
        'if databse couldn't be opened, open a blank document
        On Error Resume Next
        dbase.Close
        On Error GoTo 0
        ParseCommandLine = "BLANK"
        Exit Function
    End If
    
    'see if the next parameter is "-table"
    If UCase(cmdparam(3)) <> "-TABLE" Then
        ParseCommandLine = "BLANK"
        Exit Function
    End If
    
    'check for the table name existing in the database
    test_flag = False
    For i = 0 To dbase.TableDefs.Count - 1
        If UCase(dbase.TableDefs(i).Name) = UCase(cmdparam(4)) Then
            test_flag = True
            table_index = i
            Exit For
        End If
    Next i
    
    If Not test_flag Then
        MsgBox "'" & cmdparam(4) & "' is not a valid table name!", vbExclamation
        dbase.Close
        ParseCommandLine = "BLANK"
        Exit Function
    End If
    
    'see if the next parameter is "-searchfield"
    If UCase(cmdparam(5)) <> "-SEARCHFIELD" Then
        ParseCommandLine = "BLANK"
        Exit Function
    End If
    
    'check for the field name existing in the table
    test_flag = False
    For i = 0 To dbase.TableDefs(table_index).Fields.Count - 1
        If UCase(dbase.TableDefs(table_index).Fields(i).Name) = UCase(cmdparam(6)) Then
            test_flag = True
            test_index = i
            Exit For
        End If
    Next i
    
    If Not test_flag Then
        MsgBox "'" & cmdparam(6) & "' is not a valid field!", vbExclamation
        dbase.Close
        ParseCommandLine = "BLANK"
        Exit Function
    End If
    
    'see if the next parameter is "-searchvalue"
    If UCase(cmdparam(7)) <> "-SEARCHVALUE" Then
        ParseCommandLine = "BLANK"
        Exit Function
    End If
    
    Set record = dbase.OpenRecordset("SELECT " & cmdparam(6) _
        & " FROM " & cmdparam(4) & " WHERE " & cmdparam(6) _
        & " = " & val(cmdparam(8)), dbOpenDynaset)
    If record.EOF Then
        MsgBox "'" & cmdparam(6) & " = " & cmdparam(8) _
            & "' is not a valid field value!", vbExclamation
        record.Close
        dbase.Close
        ParseCommandLine = "BLANK"
        Exit Function
    End If
    record.Close
   
    'see if the next parameter is "-loadfield"
    If UCase(cmdparam(9)) <> "-LOADFIELD" Then
        ParseCommandLine = "BLANK"
        Exit Function
    End If
    
    'check for the field name existing in the table
    test_flag = False
    For i = 0 To dbase.TableDefs(table_index).Fields.Count - 1
        If UCase(dbase.TableDefs(table_index).Fields(i).Name) = UCase(cmdparam(10)) Then
            test_flag = True
            test_index = i
            Exit For
        End If
    Next i
    
    If Not test_flag Then
        MsgBox "'" & cmdparam(10) & "' is not a valid field!", vbExclamation
        dbase.Close
        ParseCommandLine = "BLANK"
        Exit Function
    End If
    
    ParseCommandLine = "DATABASE"
End Function


Private Sub CheckSaveStatus()
    Dim rsp As Integer
        
    If HTMLEdit.IsDirty And CanSave Then
        Select Case CmdType
            Case "DATABASE"
                If vbYes = MsgBox("You have unsaved changes.  Do you want to " _
                    & "save the changes before quitting ?", vbQuestion + vbYesNo) Then
                        mnuFileSave_Click
                    MsgBox "Changes to '" & DataItem!data_label & "' were saved.", vbInformation
                End If
                DataItem.Close
                dbase.Close
            Case "FILE"
                If vbYes = MsgBox("You have unsaved changes.  Do you want to " _
                    & "save the changes before quitting ?", vbQuestion + vbYesNo) Then
                        mnuFileSave_Click
                    MsgBox "Changes to '" & CurrentFileName & "' were saved.", vbInformation
                End If
            Case "BLANK"
                rsp = MsgBox("This document is not assosciated with an item.  " _
                    & "Do you wish to save it to a file instead?" _
                    & Chr(13) & "NOTE:  If you choose 'No', " _
                    & "your changes will be lost.", vbExclamation + vbYesNo)
                If rsp = vbYes Then
                    mnuFileSaveAs_Click
                End If
        End Select
    End If
End Sub
