VERSION 5.00
Begin VB.Form frmNutrient 
   Caption         =   "Nutrient"
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   ScaleHeight     =   2355
   ScaleWidth      =   4140
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox tbNut 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.CheckBox cDisplay 
      Caption         =   "Display ?"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cbCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2783
      TabIndex        =   8
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cbApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   1583
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cbOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   383
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
   Begin VB.ComboBox cboRDA 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox tbRDA 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   840
      Width           =   975
   End
   Begin VB.ComboBox cboNut 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "RDA:"
      Height          =   195
      Left            =   480
      TabIndex        =   9
      Top             =   840
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nutrient:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   600
   End
End
Attribute VB_Name = "frmNutrient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CanExit As Boolean

Private Sub cbApply_Click()
    Dim record As Recordset
    
    Me.MousePointer = vbHourglass
    
    If UserAction = "Modify" Or UserAction = "Add" Then
        If cboRDA.Text = "" Then
            MsgBox "You must choose units for the RDA!", vbExclamation
            cboRDA.SetFocus
            GoTo Done
        End If
        
        If tbRDA.Text = "" Then
            MsgBox "You must enter an RDA!", vbExclamation
            tbRDA.SetFocus
            GoTo Done
        End If
    End If
    
    If UserAction = "Modify" Then
        If cboNut.Text = "" Then
            MsgBox "You must enter a name for the nutrient!", vbExclamation
            cboNut.SetFocus
            GoTo Done
        End If
    End If
    
    If UserAction = "Add" Then
        If tbNut.Text = "" Then
            MsgBox "You must enter a name for the nutrient!", vbExclamation
            tbNut.SetFocus
            GoTo Done
        End If
    End If
    
    Select Case UserAction
        Case "Modify"
            Set record = dbase.OpenRecordset("SELECT * FROM HL_Nutrients WHERE [Name] = '" _
                & cboNut.Text & "'", dbOpenDynaset)
            If Not record.EOF Then
                record.Edit
                record![RDA] = tbRDA.Text
                record![RDA Units] = cboRDA.Text
                If cDisplay.Value = 1 Then
                    record![Display] = True
                Else
                    record![Display] = False
                End If

                record.Update
                CanExit = True
                NeedToClearSup = True
                SupLoaded = False
            Else
                MsgBox "This nutrient was not found in the database!", vbExclamation
            End If
            cbApply.Enabled = False
            cbOK.Enabled = False
        Case "Add"
            Set record = dbase.OpenRecordset("SELECT * FROM HL_Nutrients WHERE [Name] = '" _
                & tbNut.Text & "'", dbOpenDynaset)
            If record.EOF Then
                record.AddNew
                record![Nutrient_ID] = FindFreeID(dbase, "HL_Nutrients", "Nutrient_ID")
                record![Name] = tbNut.Text
                record![RDA] = tbRDA.Text
                record![RDA Units] = cboRDA.Text
                If cDisplay.Value = 1 Then
                    record![Display] = True
                Else
                    record![Display] = False
                End If
                record.Update
                CanExit = True
                cbApply.Enabled = False
                cbOK.Enabled = False
                NeedToClearSup = True
                SupLoaded = False
            Else
                MsgBox "The nutrient '" & tbNut.Text & "' already exists!", _
                    vbExclamation
                GoTo Done
            End If
        Case "Delete"
            If vbNo = MsgBox("Are you sure you want to delete the nutrient '" _
                & cboNut.Text & "' ?", vbYesNo + vbExclamation) Then
                GoTo Done
            Else
                dbase.Execute ("DELETE * FROM HL_Nutrients WHERE [Name] = '" _
                & cboNut.Text & "';")
                AutoFillName
                CanExit = True
                NeedToClearSup = True
                SupLoaded = False
            End If
    End Select
Done:
    Me.MousePointer = vbArrow
End Sub

Private Sub cbCancel_Click()
    Unload Me
End Sub

Private Sub cborda_Click()
    If Not cbApply.Enabled Then
        cbApply.Enabled = True
        cbOK.Enabled = True
    End If
End Sub

Private Sub cbOK_Click()
    cbApply_Click
    If CanExit Then
        Unload Me
    End If
End Sub

Private Sub cbonut_Click()
    Select Case UserAction
        Case "Modify", "Delete"
            FillRestFromName
    End Select
End Sub


Private Sub cDisplay_Click()
    If Not cbApply.Enabled Then
        cbApply.Enabled = True
        cbOK.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    
    Dim record As Recordset
    
    Set record = dbase.OpenRecordset("SELECT * FROM HL_Units ORDER BY [Unit]", _
        dbOpenDynaset)
    While Not record.EOF
        cboRDA.AddItem record![Unit]
        record.MoveNext
    Wend
    record.Close
    cboRDA.ListIndex = 0
    
   
    Select Case UserAction
        Case "Modify"
            Me.Caption = "Modify Nutrient"
            tbNut.Visible = False
            cboNut.Visible = True
            cbApply.Caption = "Apply"
            cbOK.Enabled = False
            cbApply.Enabled = False
            
            AutoFillName
            
        Case "Add"
            Me.Caption = "Add New Nutrient"
            tbNut.Visible = True
            cboNut.Visible = False
            tbRDA.Text = ""
            cbApply.Caption = "Add"
            cbOK.Enabled = False
            cbApply.Enabled = False
            
        Case "Delete"
            Me.Caption = "Delete Nutrient"
            tbNut.Visible = False
            cboNut.Visible = True
            cbApply.Caption = "Delete"
            
            AutoFillName
    End Select
    cbOK.Enabled = False
    cbApply.Enabled = False
End Sub

Private Sub FillRestFromName()
    Dim record As Recordset
    
    Set record = dbase.OpenRecordset("SELECT * FROM HL_Nutrients WHERE [Name] = '" _
        & cboNut.Text & "'", dbOpenDynaset)
        
    While Not record.EOF
        cboRDA.Text = record![RDA Units]
        tbRDA.Text = record![RDA]
        If record![Display] = True Then
            cDisplay.Value = 1
        Else
            cDisplay.Value = 0
        End If
        record.MoveNext
    Wend
    record.Close
End Sub

Private Sub AutoFillName()
    Dim record As Recordset
    Dim i As Long
    
    For i = 0 To cboNut.ListCount - 1
        cboNut.RemoveItem 0
    Next i
    
    Set record = dbase.OpenRecordset("SELECT DISTINCT [Name]" _
        & " FROM HL_Nutrients ORDER BY [Name]", dbOpenDynaset)
    While Not record.EOF
        cboNut.AddItem record![Name]
        record.MoveNext
    Wend
    record.Close
    If cboNut.ListCount > 0 Then
        cboNut.Text = cboNut.List(0)
    Else
        cboNut.ListIndex = -1
        cboRDA.ListIndex = -1
        tbRDA.Text = ""
        cDisplay.Value = 0
        cbApply.Enabled = False
        cbOK.Enabled = False
    End If
End Sub

Private Sub tbrda_Change()
    If Not cbApply.Enabled Then
        cbApply.Enabled = True
        cbOK.Enabled = True
    End If
End Sub

Private Sub tbrda_GotFocus()
    tbRDA.SelStart = 0
    tbRDA.SelLength = Len(tbRDA.Text)
End Sub

Private Sub tbnut_Change()
    If Not cbApply.Enabled Then
        cbApply.Enabled = True
        cbOK.Enabled = True
    End If
End Sub

Private Sub tbnut_GotFocus()
    tbNut.SelStart = 0
    tbNut.SelLength = Len(tbNut.Text)
End Sub


Private Sub tbRDA_LostFocus()
    If Not IsNumeric(tbRDA.Text) Then
        MsgBox "You must enter a valid number!", vbExclamation
        tbRDA.SetFocus
    End If
End Sub
