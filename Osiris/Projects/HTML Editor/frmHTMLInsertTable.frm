VERSION 5.00
Begin VB.Form frmHTMLInsertTable 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Table"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Table Attributes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   6015
      Begin VB.TextBox tbHeight 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4440
         TabIndex        =   25
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox tbAlign 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4440
         TabIndex        =   21
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox tbBGColor 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4440
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox tbBorderColor 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4440
         TabIndex        =   19
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox tbBorder 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Text            =   "1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox tbCellPadding 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Text            =   "1"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox tbCellSpacing 
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Text            =   "1"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox tbWidth 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Text            =   "75"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblHeight 
         Caption         =   "Height:"
         Height          =   255
         Left            =   3000
         TabIndex        =   26
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Align:"
         Height          =   255
         Left            =   3000
         TabIndex        =   24
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Background Color:"
         Height          =   255
         Left            =   3000
         TabIndex        =   23
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblBorderColor 
         Caption         =   "Border Color:"
         Height          =   255
         Left            =   3000
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "%"
         Height          =   255
         Left            =   2760
         TabIndex        =   16
         Top             =   1440
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Width:"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Cell Spacing:"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Cell Padding:"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Border:"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Properties"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6015
      Begin VB.TextBox tbCaption 
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   480
         Width           =   4455
      End
      Begin VB.TextBox tbRows 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox tbColumns 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label CaptionLabel 
         Caption         =   "Caption:"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label RowLabel 
         Caption         =   "Rows:"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label ColLabel 
         Caption         =   "Columns:"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1440
         Width           =   1455
      End
   End
   Begin VB.CommandButton cbCancel 
      Cancel          =   -1  'True
      Caption         =   "Cacel"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cbOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   4440
      Width           =   1095
   End
End
Attribute VB_Name = "frmHTMLInsertTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TableParameters As DEInsertTableParam

Private Sub cbCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set TableParameters = _
        CreateObject("DEInsertTableParam.DEInsertTableParam.1")
    tbRows = TableParameters.NumRows
    tbColumns = TableParameters.NumCols
    tbCaption = TableParameters.Caption
End Sub

Private Sub cbOk_Click()
    TableParameters.NumRows = tbRows
    TableParameters.NumCols = tbColumns
    TableParameters.TableAttrs = "border = " & format$(tbBorder.text) _
        & " cellpadding = " & format$(tbCellPadding.text) _
        & " cellspacing = " & format$(tbCellSpacing.text) _
        & " width = " & tbWidth.text & "%"
    TableParameters.Caption = tbCaption.text
    
    frmHTMLEd.HTMLEdit.ExecCommand DECMD_INSERTTABLE, OLECMDEXECOPT_DONTPROMPTUSER, TableParameters
    Unload Me
End Sub

Private Sub tbBorder_GotFocus()
    tbBorder.SelStart = 0
    tbBorder.SelLength = Len(tbBorder.text)
End Sub

Private Sub tbBorder_LostFocus()
    If Not IsNumeric(tbBorder.text) Then
        MsgBox "There must be a numeric value in this cell!", vbExclamation
        tbBorder.SetFocus
    End If
End Sub

Private Sub tbCellPadding_GotFocus()
    tbCellPadding.SelStart = 0
    tbCellPadding.SelLength = Len(tbCellPadding.text)
End Sub

Private Sub tbCellPadding_LostFocus()
    If Not IsNumeric(tbCellPadding.text) Then
        MsgBox "There must be a numeric value in this cell!", vbExclamation
        tbCellPadding.SetFocus
    End If
End Sub

Private Sub tbCellSpacing_GotFocus()
    tbCellSpacing.SelStart = 0
    tbCellSpacing.SelLength = Len(tbCellSpacing.text)
End Sub

Private Sub tbCellSpacing_LostFocus()
    If Not IsNumeric(tbCellSpacing.text) Then
        MsgBox "There must be a numeric value in this cell!", vbExclamation
        tbCellSpacing.SetFocus
    End If
End Sub

Private Sub tbWidth_GotFocus()
    tbWidth.SelStart = 0
    tbWidth.SelLength = Len(tbWidth.text)
End Sub

Private Sub tbWidth_LostFocus()
    If Not IsNumeric(tbWidth.text) Then
        MsgBox "There must be a numeric value in this cell!", vbExclamation
        tbWidth.SetFocus
    Else
        tbWidth.text = format$(CLng(tbWidth.text))
        If val(tbWidth.text) <= 0 Or val(tbWidth.text) > 100 Then
            MsgBox "The width must be in the range of 1-100% !", vbExclamation
            tbWidth.SetFocus
        End If
    End If
End Sub

Private Sub tbRows_GotFocus()
    tbRows.SelStart = 0
    tbRows.SelLength = Len(tbRows.text)
End Sub

Private Sub tbRows_LostFocus()
    If Not IsNumeric(tbRows.text) Then
        MsgBox "There must be a numeric value in this cell!", vbExclamation
        tbRows.SetFocus
    End If
End Sub

Private Sub tbColumns_GotFocus()
    tbColumns.SelStart = 0
    tbColumns.SelLength = Len(tbColumns.text)
End Sub

Private Sub tbColumns_LostFocus()
    If Not IsNumeric(tbColumns.text) Then
        MsgBox "There must be a numeric value in this cell!", vbExclamation
        tbColumns.SetFocus
    End If
End Sub
