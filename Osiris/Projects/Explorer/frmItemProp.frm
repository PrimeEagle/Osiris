VERSION 5.00
Begin VB.Form frmItemProp 
   Caption         =   "Form1"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   8175
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboSmallIcon 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1440
      Width           =   2415
   End
   Begin VB.ComboBox cboLargeIcon 
      Height          =   315
      Left            =   1560
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   960
      Width           =   2415
   End
   Begin VB.ComboBox cboDataType 
      Height          =   315
      Left            =   5760
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   240
      Width           =   1815
   End
   Begin VB.CheckBox cbReadOnly 
      Caption         =   "Read Only"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox tbDataLabel 
      Height          =   375
      Left            =   1440
      MaxLength       =   255
      TabIndex        =   5
      Text            =   "Text1"
      ToolTipText     =   "Type in the label of your data (max 64 characters)."
      Top             =   2760
      Width           =   5415
   End
   Begin VB.CommandButton cbCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      ToolTipText     =   "Do not save changes."
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cbOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      ToolTipText     =   "Save changes."
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox tbDataValue 
      Height          =   1845
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   "Type in the value of your data."
      Top             =   3360
      Width           =   5415
   End
   Begin VB.Label lblSmallIcon 
      Caption         =   "Small Icon:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblParentNodeVal 
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5760
      TabIndex        =   14
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblParentNode 
      Caption         =   "Parent Node:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   10
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lblLargeIcon 
      Caption         =   "Large Icon:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblDataType 
      Caption         =   "Data Type:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblDataIDVal 
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblDataID 
      Caption         =   "Data ID:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lblDataLabel 
      Caption         =   "Data Label:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblDataValue 
      Caption         =   "Data Value:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3360
      Width           =   1095
   End
End
Attribute VB_Name = "frmItemProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lastlargeicon As String
Dim lastsmallicon As String
Dim lastdatatype As String


Private Sub cbCancel_Click()
    Unload Me
    fMainForm.Show
End Sub

Private Sub cboDataType_Click()
    If cboDataType.text <> lastdatatype Then
        cbOK.Enabled = True
    End If
End Sub

Private Sub cboLargeIcon_Change()
    If cboLargeIcon.text <> lastlargeicon Then
        cbOK.Enabled = True
    End If
End Sub

Private Sub cbolargeicon_Click()
    If cboLargeIcon.text <> lastlargeicon Then
        cbOK.Enabled = True
    End If
End Sub

Private Sub cbOK_Click()
    Dim record As Recordset
    Dim tempstr2 As String
    Dim cMyUC As Object

    tempstr2 = tvAttribCol.Item(AddK(lvAttribCol(currentItem.Key).parent_node)).table_name
    Set record = dbase.OpenRecordset("SELECT * FROM " & tempstr2 & _
                " WHERE Data_ID = " & RemoveK(currentItem.Key), dbOpenDynaset)
    record.Edit
    record!data_type = cboDataType.text
    
    If cboLargeIcon.text = NONE_LABEL Then
        record!icon_large = ""
    Else
        record!icon_large = cboLargeIcon.text
    End If
    
    If cboSmallIcon.text = NONE_LABEL Then
        record!icon_small = ""
    Else
        record!icon_small = cboSmallIcon.text
    End If
    
    
    
    If cbReadOnly.Value = 1 Then
        record!read_only = True
    Else
        record!read_only = False
    End If
    
    record!data_value = tbDataValue.text
    record!data_label = tbDataLabel.text
    record.Update
    record.Close
    
    
    currentItem.text = tbDataLabel.text
            
    CopyMemory cMyUC, Instances(1).ClassAddr, 4
    
    If cboLargeIcon.text = NONE_LABEL Then
        currentItem.Icon = 0
    Else
        currentItem.Icon = cMyUC.imlIconsLarge.ListImages.Item("K" & cboLargeIcon.text).Index
    End If
    
    If cboSmallIcon.text = NONE_LABEL Then
        currentItem.SmallIcon = 0
    Else
        currentItem.SmallIcon = cMyUC.imlIconsSmall.ListImages.Item("K" & cboSmallIcon.text).Index
    End If
    
    CopyMemory cMyUC, 0&, 4
    
    Unload Me
    fMainForm.Show
End Sub

Private Sub cboSmallIcon_Change()
    If cboSmallIcon.text <> lastsmallicon Then
        cbOK.Enabled = True
    End If
End Sub

Private Sub cboSmallIcon_Click()
    If cboSmallIcon.text <> lastsmallicon Then
        cbOK.Enabled = True
    End If
End Sub

Private Sub cbReadOnly_Click()
    If cbReadOnly.Value = 1 Then
        cbReadOnly.FontSize = 12
    Else
        cbReadOnly.FontSize = 8
    End If
    cbOK.Enabled = True
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ShiftDown, AltDown, CtrlDown
    
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    Select Case KeyCode
        Case vbKeyEscape
            cbCancel_Click
    End Select
End Sub

Private Sub Form_Load()
    
    Dim record As Recordset
    Dim iconrecord As Recordset
    Dim tempstr As String
    Dim tempstr2 As String
   
    tempstr2 = tvAttribCol.Item(AddK(lvAttribCol(currentItem.Key).parent_node)).table_name
    Set record = dbase.OpenRecordset("SELECT * FROM " & tempstr2 & _
                " WHERE Data_ID = " & RemoveK(currentItem.Key), dbOpenDynaset)

    
    lblDataIDVal.ForeColor = RGB(0, 0, 255)
    lblDataIDVal.Caption = record!data_id
    
    
    cboDataType.AddItem "String"
    cboDataType.AddItem "Number"
    cboDataType.AddItem "Date"
    cboDataType.AddItem "Boolean"
    cboDataType.AddItem "Image"
    cboDataType.text = record!data_type
    lastdatatype = cboDataType.text
    
    tbDataLabel.text = record!data_label
    Me.Caption = "Properties for '" & tbDataLabel.text & "'"
    
    tbDataValue.text = record!data_value
    
    Set iconrecord = dbase.OpenRecordset("LargeIcons", dbOpenTable)
    cboLargeIcon.AddItem NONE_LABEL
    While Not iconrecord.EOF
        tempstr = iconrecord!data_label
        cboLargeIcon.AddItem tempstr
        iconrecord.MoveNext
    Wend
    iconrecord.Close
    
    tempstr = record!icon_large
    If tempstr = "" Then
        cboLargeIcon.text = NONE_LABEL
    Else
        cboLargeIcon.text = tempstr
    End If
    
    lastlargeicon = cboLargeIcon.text
    
    Set iconrecord = dbase.OpenRecordset("SmallIcons", dbOpenTable)
    cboSmallIcon.AddItem NONE_LABEL
    While Not iconrecord.EOF
        tempstr = iconrecord!data_label
        cboSmallIcon.AddItem tempstr
        iconrecord.MoveNext
    Wend
    iconrecord.Close
    
    tempstr = record!icon_small
    If tempstr = "" Then
        cboSmallIcon.text = NONE_LABEL
    Else
        cboSmallIcon.text = tempstr
    End If
    
    lastsmallicon = cboSmallIcon.text
    
   
    lblParentNodeVal.ForeColor = RGB(0, 0, 255)
    lblParentNodeVal.Caption = record!parent_node
   
    cbReadOnly.FontBold = True
    If record!read_only = True Then
        cbReadOnly.Value = 1
    Else
        cbReadOnly.Value = 0
    End If
    
        
    
    
    record.Close
        
    tbDataLabel.FontBold = True
    tbDataLabel.FontSize = 10
    
    Me.FontBold = tbDataLabel.FontBold
    Me.FontSize = tbDataLabel.FontSize
    
    
    tbDataLabel.Left = 1400
    tbDataValue.Left = 1400
    
        
    tbDataLabel.Width = TextWidth(tbDataLabel) + 400
    If tbDataLabel.Width > 5000 Then
        tbDataValue.Width = tbDataLabel.Width
    Else
        tbDataValue.Width = 5000
    End If
    
    cbOK.Left = tbDataValue.Left
    cbCancel.Left = tbDataValue.Left + tbDataValue.Width - cbCancel.Width
                
    cbOK.Enabled = False
    Me.Width = tbDataValue.Left + tbDataValue.Width + 1400
    
    Select Case tvAttribCol(nodx.Key).table_name
        Case "LargeIcons", "SmallIcons"
            cboDataType.Enabled = False
            cbReadOnly.Enabled = False
            tbDataLabel.Enabled = False
            tbDataValue.Enabled = False
            cboLargeIcon.Enabled = False
            cboSmallIcon.Enabled = False
    End Select
End Sub


Private Sub tbDataLabel_Change()
    Static oldLabelLength As Integer
    Dim i As Integer
    
    i = Len(tbDataLabel.text)
    If i <> oldLabelLength Then
        tbDataLabel.Width = TextWidth(tbDataLabel) + 400
        If tbDataLabel.Width > 5000 Then
            tbDataValue.Width = tbDataLabel.Width
        Else
            tbDataValue.Width = 5000
        End If
        cbCancel.Left = tbDataValue.Left + tbDataValue.Width - cbCancel.Width
        Me.Width = tbDataValue.Left + tbDataValue.Width + 1400
    End If
    oldLabelLength = i
    If i > 0 And Not lvAttribCol(currentItem.Key).read_only Then   'if not read only
        cbOK.Enabled = True
    Else
        cbOK.Enabled = False
    End If
End Sub

Private Sub tbDataValue_Change()
    If Not lvAttribCol(currentItem.Key).read_only Then   'if not read only
        cbOK.Enabled = True
    End If
End Sub


