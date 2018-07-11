VERSION 5.00
Begin VB.Form frmUsers 
   Caption         =   "Users"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   7755
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboPreg 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox tbFirstName 
      Height          =   285
      Left            =   5400
      TabIndex        =   4
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox tbLastName 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   2055
   End
   Begin VB.ComboBox cboSex 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cbApply 
      Caption         =   "Apply"
      Default         =   -1  'True
      Height          =   495
      Left            =   3210
      TabIndex        =   10
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cbCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4650
      TabIndex        =   11
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cbOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   1770
      TabIndex        =   9
      Top             =   2880
      Width           =   1095
   End
   Begin VB.ComboBox cboFrame 
      Height          =   315
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox tbDOB 
      Height          =   315
      Left            =   5400
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.ComboBox cboFirstName 
      Height          =   315
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.ComboBox cboLastName 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lblPreg 
      Caption         =   "Are you pregnant?"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Frame Size:"
      Height          =   255
      Left            =   4320
      TabIndex        =   15
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Date of Birth:"
      Height          =   255
      Left            =   4320
      TabIndex        =   14
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Sex:"
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "First Name:"
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Last Name:"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmUsers"
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
        If cboSex.Text = "" Then
            MsgBox "You must choose a Sex!", vbExclamation
            cboSex.SetFocus
            GoTo Done
        End If
        
        If cboPreg.Text = "" And cboPreg.Visible Then
            MsgBox "You must enter whether you are pregnant or not!", vbExclamation
            cboPreg.SetFocus
            GoTo Done
        End If
        
        If tbDOB.Text = "" Then
            MsgBox "You must enter a Date of Birth!", vbExclamation
            tbDOB.SetFocus
            GoTo Done
        End If
    
        If Not IsDate(tbDOB.Text) Then
            MsgBox "You must enter a valid Date of Birth!", vbExclamation
            tbDOB.SetFocus
            GoTo Done
        End If
    
        If cboFrame.Text = "" Then
            MsgBox "You must choose a Frame Size!", vbExclamation
            cboFrame.SetFocus
            GoTo Done
        End If
    End If
    
    If UserAction = "Add" Then
        If Len(tbLastName.Text) <= 0 Then
            MsgBox "You must enter a Last Name!", vbExclamation
            tbLastName.SetFocus
            GoTo Done
        End If
        
        If Len(tbFirstName.Text) <= 0 Then
            MsgBox "You must enter a First Name!", vbExclamation
            tbFirstName.SetFocus
            GoTo Done
        End If
    End If
    
    Select Case UserAction
        Case "Modify"
            Set record = dbase.OpenRecordset("SELECT * FROM HL_Users WHERE [Last Name] = '" _
                & cboLastName.Text & "' AND [First Name] = '" & cboFirstName.Text & "'", _
                dbOpenDynaset)
            If Not record.EOF Then
                record.Edit
                record![Sex] = cboSex.Text
                If cboPreg.Visible = False Then
                    record![Pregnant] = False
                Else
                    If cboPreg.Text = "Yes" Then
                        record![Pregnant] = True
                    Else
                        record![Pregnant] = False
                    End If
                End If
                record![Birthday] = tbDOB.Text
                record![Frame Size] = cboFrame.Text
                record.Update
                CanExit = True
                If cboPreg.Text = "Yes" Then
                    GetPregnancyInfo
                End If
            Else
                MsgBox "This user was not found in the database!", vbExclamation
            End If
            cbApply.Enabled = False
            cbOK.Enabled = False
        Case "Add"
            Set record = dbase.OpenRecordset("SELECT * FROM HL_Users WHERE [Last Name] = '" _
                & tbLastName.Text & "' AND [First Name] = '" & tbFirstName.Text & "'", _
                dbOpenDynaset)
            If record.EOF Then
                record.AddNew
                record![User_ID] = FindFreeID(dbase, "HL_Users", "User_ID")
                record![Last Name] = tbLastName.Text
                record![First Name] = tbFirstName.Text
                record![Sex] = cboSex.Text
                If cboPreg.Visible = False Then
                    record![Pregnant] = False
                Else
                    If cboPreg.Text = "Yes" Then
                        record![Pregnant] = True
                    Else
                        record![Pregnant] = False
                    End If
                End If
                record![Birthday] = tbDOB.Text
                record![Frame Size] = cboFrame.Text
                record.Update
                CanExit = True
                cbApply.Enabled = False
                cbOK.Enabled = False
                If cboPreg.Text = "Yes" Then
                    GetPregnancyInfo
                End If
            Else
                MsgBox "The user '" & tbFirstName.Text & " " & tbLastName.Text _
                    & "' already exists!", vbExclamation
                GoTo Done
            End If
        Case "Delete"
            If vbNo = MsgBox("Are you sure you want to delete the user '" _
                & cboFirstName.Text & " " & cboLastName.Text & "' ?", _
                vbYesNo + vbExclamation) Then
                GoTo Done
            Else
                dbase.Execute ("DELETE * FROM HL_Users WHERE [Last Name] = '" _
                & cboLastName.Text & "' AND [First Name] = '" & cboFirstName.Text & "';")
                AutoFillLastName
                CanExit = True
            End If
    End Select
Done:
    Me.MousePointer = vbArrow
End Sub

Private Sub cbCancel_Click()
    Unload Me
End Sub

Private Sub cboFirstName_Click()
    Select Case UserAction
        Case "Modify", "Delete"
            FillRestFromFirst
    End Select
End Sub


Private Sub cboFrame_Click()
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

Private Sub cboLastName_Click()
    Select Case UserAction
        Case "Modify", "Delete"
            FillFirstFromLast
    End Select
End Sub



Private Sub cboSex_Click()
    If Not cbApply.Enabled Then
        cbApply.Enabled = True
        cbOK.Enabled = True
    End If
    If cboSex.Text = "Male" Then
        cboPreg.Visible = False
        lblPreg.Visible = False
    Else
        cboPreg.Visible = True
        lblPreg.Visible = True
    End If
End Sub

Private Sub Form_Load()
    
    cboSex.AddItem "Male"
    cboSex.AddItem "Female"
            
    cboFrame.AddItem "Small"
    cboFrame.AddItem "Medium"
    cboFrame.AddItem "Large"
    
    cboPreg.AddItem "Yes"
    cboPreg.AddItem "No"
    
    Select Case UserAction
        Case "Modify"
            Me.Caption = "Modify User"
            tbLastName.Visible = False
            tbFirstName.Visible = False
            cboLastName.Visible = True
            cboFirstName.Visible = True
            cbApply.Caption = "Apply"
            cbOK.Enabled = False
            cbApply.Enabled = False
            
            AutoFillLastName
            
        Case "Add"
            Me.Caption = "Add New User"
            tbLastName.Visible = True
            tbFirstName.Visible = True
            cboLastName.Visible = False
            cboFirstName.Visible = False
            tbDOB.Text = ""
            cbApply.Caption = "Add User"
            cbOK.Enabled = False
            cbApply.Enabled = False
            
        Case "Delete"
            Me.Caption = "Delete User"
            tbLastName.Visible = False
            tbFirstName.Visible = False
            cboLastName.Visible = True
            cboFirstName.Visible = True
            cbApply.Caption = "Delete"
            
            AutoFillLastName
    End Select
End Sub

Private Sub FillFirstFromLast()
    Dim i As Long
    Dim record As Recordset
    
    For i = 0 To cboFirstName.ListCount - 1
        cboFirstName.RemoveItem 0
    Next i
    
    Set record = dbase.OpenRecordset("SELECT [Last Name],[First Name]" _
        & " FROM HL_Users WHERE [Last Name] = '" & cboLastName.Text _
        & "' ORDER BY [First Name]", dbOpenDynaset)
    While Not record.EOF
        cboFirstName.AddItem record![First Name]
        record.MoveNext
    Wend
    record.Close
    If cboFirstName.ListCount > 0 Then
        cboFirstName.Text = cboFirstName.List(0)
    End If
End Sub

Private Sub FillRestFromFirst()
    Dim record As Recordset
    
    Set record = dbase.OpenRecordset("SELECT * FROM HL_Users WHERE [Last Name] = '" _
        & cboLastName.Text & "' AND [First Name] = '" & cboFirstName.Text & "'", _
        dbOpenDynaset)
        
    While Not record.EOF
        cboSex.Text = record![Sex]
        cboFrame.Text = record![Frame Size]
        tbDOB.Text = CDate(record![Birthday])
        If record![Pregnant] = True Then
            cboPreg.Text = "Yes"
        Else
            cboPreg.Text = "No"
        End If
        record.MoveNext
    Wend
    record.Close
End Sub

Private Sub AutoFillLastName()
    Dim record As Recordset
    Dim i As Long
    
    For i = 0 To cboLastName.ListCount - 1
        cboLastName.RemoveItem 0
    Next i
    
    Set record = dbase.OpenRecordset("SELECT DISTINCT [Last Name]" _
        & " FROM HL_Users ORDER BY [Last Name]", dbOpenDynaset)
    While Not record.EOF
        cboLastName.AddItem record![Last Name]
        record.MoveNext
    Wend
    record.Close
    If cboLastName.ListCount > 0 Then
        cboLastName.Text = cboLastName.List(0)
    Else
        cboFirstName.ListIndex = -1
        cboFirstName.ListIndex = -1
        cboSex.ListIndex = -1
        cboPreg.ListIndex = -1
        tbDOB.Text = ""
        cboFrame.ListIndex = -1
        cbApply.Enabled = False
        cbOK.Enabled = False
    End If
End Sub

Private Sub tbDOB_Change()
    If Not cbApply.Enabled Then
        cbApply.Enabled = True
        cbOK.Enabled = True
    End If
End Sub

Private Sub tbDOB_GotFocus()
    tbDOB.SelStart = 0
    tbDOB.SelLength = Len(tbDOB.Text)
End Sub

Private Sub tbFirstName_Change()
    If Not cbApply.Enabled Then
        cbApply.Enabled = True
        cbOK.Enabled = True
    End If
End Sub

Private Sub tbFirstName_GotFocus()
    tbFirstName.SelStart = 0
    tbFirstName.SelLength = Len(tbFirstName.Text)
End Sub

Private Sub tbLastName_Change()
    If Not cbApply.Enabled Then
        cbApply.Enabled = True
        cbOK.Enabled = True
    End If
End Sub

Private Sub tbLastName_GotFocus()
    tbLastName.SelStart = 0
    tbLastName.SelLength = Len(tbLastName.Text)
End Sub

Private Sub GetPregnancyInfo()
    MsgBox "TEMP:  Get more info about the pregnancy.", vbInformation
End Sub
