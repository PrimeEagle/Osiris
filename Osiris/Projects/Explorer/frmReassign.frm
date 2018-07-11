VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReassign 
   Caption         =   "Reassign Table Names"
   ClientHeight    =   1515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1515
   ScaleWidth      =   6210
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageCombo cboTable 
      Height          =   330
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
      Text            =   "ImageCombo1"
   End
   Begin VB.CommandButton cbCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cbOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblMessage 
      Caption         =   "There are "
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   5055
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmReassign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbCancel_Click()
    selected_table = "-1"
    Unload Me
End Sub

Private Sub cbOK_Click()
    cboTable.SetFocus
    SendKeys "{END}", True
    SendKeys "+{HOME}", True
    
    selected_table = cboTable.SelectedItem.text
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim tempstr As String
    Dim TestIndex As Long
    
    lblMessage.Caption = "There are " & format$(record_count) & _
            " records that reference the table " & fPropForm.cboType.SelectedItem.text & _
            ".  These records will be reassigned to one of the following tables:"
            
    Set cboTable.ImageList = fMainForm.imlMenu
        
    For i = 1 To fPropForm.cboType.ComboItems.count
        tempstr = fPropForm.cboType.ComboItems(i).text
        If tempstr <> fPropForm.cboType.SelectedItem.text Then
            On Error Resume Next
            TestIndex = -1
            TestIndex = fMainForm.imlMenu.ListImages.item("AccessTable").Index
            If TestIndex = -1 Then
                TestIndex = 0
            End If
            On Error GoTo 0
            cboTable.ComboItems.Add , UCase(tempstr), tempstr, TestIndex, _
                    TestIndex, 0
        End If
    Next i
    cboTable.SelectedItem = cboTable.ComboItems(1)
End Sub

