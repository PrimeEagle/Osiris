VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHTMLSource 
   Caption         =   "Osiris HTML Source Viewer"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   510
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   6435
   ScaleWidth      =   6615
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6120
      Top             =   0
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtbSource 
      Height          =   5784
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   6612
      _ExtentX        =   11668
      _ExtentY        =   10213
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      AutoVerbMenu    =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6180
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   2937
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
            TextSave        =   "10/19/98"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            AutoSize        =   2
            TextSave        =   "3:11 PM"
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
   Begin MSComctlLib.Toolbar tbEdit 
      Height          =   264
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   5772
      _ExtentX        =   10186
      _ExtentY        =   476
      ButtonWidth     =   487
      ButtonHeight    =   466
      Wrappable       =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
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
            Key             =   "Save As"
            Description     =   "Save As"
            Object.ToolTipText     =   "Save As..."
            Object.Tag             =   "Save As"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Description     =   "Print"
            Object.ToolTipText     =   "Print"
            Object.Tag             =   "Print"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Description     =   "Cut"
            Object.ToolTipText     =   "Cut"
            Object.Tag             =   "Cut"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Description     =   "Copy"
            Object.ToolTipText     =   "Copy"
            Object.Tag             =   "Copy"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Description     =   "Paste"
            Object.ToolTipText     =   "Paste"
            Object.Tag             =   "Paste"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Description     =   "Undo"
            Object.ToolTipText     =   "Undo"
            Object.Tag             =   "Undo"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Refresh"
            Description     =   "Refresh"
            Object.ToolTipText     =   "Refresh from HTML Editor"
            Object.Tag             =   "Refresh"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Update"
            Description     =   "Update"
            Object.ToolTipText     =   "Update HTML Editor"
            Object.Tag             =   "Update"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open File..."
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &as..."
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
      Begin VB.Menu mnuFileUpdate 
         Caption         =   "&Update HTML Editor"
      End
      Begin VB.Menu mnuFileRefresh 
         Caption         =   "&Refresh from HTML Editor"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
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
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About Source Viewer..."
      End
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "&Help..."
      End
   End
End
Attribute VB_Name = "frmHTMLSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ExitConfirmed As Boolean
Dim TryingToExit As Boolean
Dim CancelExit As Boolean
Dim CurrentFile As String

Private Sub Form_Load()
    InitToolbars
    If Not fMainForm.LoadBlankSourceViewer Then
        rtbSource.text = frmHTMLEdit.HTMLEdit.DocumentHTML
    End If
    mnuFileUpdate.Enabled = False
    tbEdit.Buttons("Update").Enabled = False
    If Not HTMLEditRunning Then
        mnuFileRefresh.Enabled = False
        tbEdit.Buttons("Refresh").Enabled = False
    End If
    ExitConfirmed = True
    TryingToExit = False
    CancelExit = False
    CurrentFile = ""
    mnuFileSave.Enabled = False
    tbEdit.Buttons("Save").Enabled = False
    UpdateCaption
End Sub

Private Sub mnuEditCopy_Click()
    SendMessage rtbSource.hwnd, WM_COPY, 0, 0
End Sub

Private Sub mnuEditCut_Click()
   SendMessage rtbSource.hwnd, WM_CUT, 0, 0
End Sub

Private Sub mnuEditDelete_Click()
    SendKeys "{DEL}"
End Sub

Private Sub mnuEditPaste_Click()
   SendMessage rtbSource.hwnd, WM_PASTE, 0, 0
End Sub

Private Sub mnuEditUndo_Click()
    SendMessage rtbSource.hwnd, EM_UNDO, 0, 0
End Sub

Private Sub mnuFileExit_Click()
    Dim response As Integer
    
    TryingToExit = True
    
    If HTMLEditRunning Then
        If Not ExitConfirmed Then
            response = MsgBox("Do you want to update your changes before quitting?", _
                vbExclamation + vbYesNoCancel)
            Select Case response
                Case vbYes
                    mnuFileUpdate_Click
                Case vbCancel
                    TryingToExit = False
                    Exit Sub
            End Select
        End If
    Else
        If Not ExitConfirmed Then
            response = MsgBox("Do you want to save your changes before quitting?", _
                vbExclamation + vbYesNoCancel)
            Select Case response
                Case vbYes
                    If CurrentFile <> "" Then
                        mnuFileSave_Click
                    Else
                        mnuFileSaveAs_Click
                    End If
                Case vbCancel
                    TryingToExit = False
                    Exit Sub
            End Select
        End If
    End If
Quit:
    If Not CancelExit Then
        Unload Me
    End If
    CancelExit = False
End Sub

Private Sub mnuFileNew_Click()
    If Not ExitConfirmed Then
        If vbNo = MsgBox("You have unsaved changes.  Do you wish to continue?", _
                vbYesNo + vbExclamation) Then
            Exit Sub
        End If
    End If
NewFile:
    rtbSource.text = ""
    CurrentFile = "Unnamed Document"
    UpdateCaption
    ExitConfirmed = True
End Sub

Private Sub mnuFileOpen_Click()
    CommonDialog1.Filter = "HTML Files (*.htm;*.html)|*.htm;*.html|" _
        & "ASCII Text Files (*.txt)|*.txt|" _
        & "All Files (*.*)|*.*"
    CommonDialog1.FilterIndex = 1 'set default to HTML
    On Error GoTo userCancel
        CommonDialog1.ShowOpen
    On Error GoTo 0
    Select Case CommonDialog1.FilterIndex
        Case 1, 2, 3
            rtbSource.LoadFile CommonDialog1.filename, rtfText
            CurrentFile = CommonDialog1.filename
            UpdateCaption
        Case Else
            MsgBox "This file type is not supported by the Source Viewer.", vbExclamation
    End Select
userCancel:
End Sub

Private Sub mnuFilePrint_Click()
    On Local Error GoTo Error_Handler:
    With CommonDialog1
        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        If rtbSource.SelLength = 0 Then
            .Flags = .Flags + cdlPDAllPages
        Else
            .Flags = .Flags + cdlPDSelection
        End If
        .ShowPrinter
        On Local Error Resume Next
        Printer.Print ""
        rtbSource.SelPrint Printer.hdc
        Printer.EndDoc
    End With
    Exit Sub

Error_Handler:
    If err <> cdlCancel Then
        MsgBox "mnuFilePrint_Click:  Error " & err & "; " & Error
    End If
End Sub

Private Sub mnuFileRefresh_Click()
    If Not ExitConfirmed Then
        If vbNo = MsgBox("You have unsaved modifications.  If you refresh, your changes will be lost." _
            & Chr(13) & "Are you sure you want to refresh?", _
            vbExclamation + vbYesNo) Then
            Exit Sub
        End If
    End If
DoRefresh:
    rtbSource.text = frmHTMLEdit.HTMLEdit.DocumentHTML
    ExitConfirmed = True
End Sub

Private Sub mnuFileSave_Click()
    If CurrentFile <> "" Then
        rtbSource.SaveFile CurrentFile, rtfText
    End If
End Sub

Private Sub mnuFileSaveAs_Click()
    CommonDialog1.Filter = "HTML Files (*.htm;*.html)|*.htm;*.html|" _
        & "ASCII Text Files (*.txt)|*.txt|"
    CommonDialog1.FilterIndex = 1 'set default to HTML
    On Error GoTo userCancel
        CommonDialog1.ShowSave
    On Error GoTo 0
    Select Case CommonDialog1.FilterIndex
        Case 1, 2
            rtbSource.SaveFile CommonDialog1.filename, rtfText
            CurrentFile = CommonDialog1.filename
            mnuFileSave.Enabled = True
            tbEdit.Buttons("Save").Enabled = True
            UpdateCaption
        Case Else
            MsgBox "This file type is not supported by the Source Viewer.", vbExclamation
    End Select
    ExitConfirmed = True
    Exit Sub
userCancel:
    CancelExit = True
End Sub

Private Sub mnuFileUpdate_Click()
    Dim TempFileName As String
    
    If frmHTMLEdit.HTMLEdit.IsDirty Then
        If vbNo = MsgBox("You have unsaved changes in the HTML Editor.  " _
                & "Are you sure you want to continue with the update?", _
                vbExclamation + vbYesNo) Then
            If TryingToExit Then
                CancelExit = True
            End If
            Exit Sub
        End If
    End If
DoUpdate:
    TempFileName = CurTempFolder & "Update" & GetAvailableTempFile & ".html"
    rtbSource.SaveFile TempFileName, rtfText
    
    If Not ValidFile(TempFileName) Then
        MsgBox "The necessary temp file '" & TempFileName & "' was not " _
            & "created successfully!" & Chr$(13) _
            & "Update failed.", vbCritical
        Exit Sub
    End If
    
    frmHTMLEdit.HTMLEdit.LoadDocument TempFileName, False
    DoEvents
    Kill TempFileName
    ExitConfirmed = True
End Sub

Private Sub mnuHelpAbout_Click()
    frmHTMLSourceAbout.Show vbModal, Me
End Sub

Private Sub rtbSource_Change()
    If ExitConfirmed Then
        ExitConfirmed = False
    End If
    If mnuFileUpdate.Enabled = False And HTMLEditRunning Then
        mnuFileUpdate.Enabled = True
        tbEdit.Buttons("Update").Enabled = True
    End If
End Sub

Private Sub UpdateCaption()
    If CurrentFile = "" Then
        Me.Caption = "Osiris Source Viewer  (Unnamed Document)"
    Else
        Me.Caption = "Osiris Source Viewer  (" & CurrentFile & ")"
    End If
End Sub

Private Sub InitToolbars()
    On Error Resume Next
    Set tbEdit.ImageList = fMainForm.imlMenu

    tbEdit.Buttons("New").Image = fMainForm.imlMenu.ListImages.item("New").Index
    tbEdit.Buttons("Open").Image = fMainForm.imlMenu.ListImages.item("Open").Index
    tbEdit.Buttons("Save").Image = fMainForm.imlMenu.ListImages.item("Save").Index
    tbEdit.Buttons("Save As").Image = fMainForm.imlMenu.ListImages.item("Save").Index
    tbEdit.Buttons("Print").Image = fMainForm.imlMenu.ListImages.item("Print").Index
    tbEdit.Buttons("Cut").Image = fMainForm.imlMenu.ListImages.item("Cut").Index
    tbEdit.Buttons("Copy").Image = fMainForm.imlMenu.ListImages.item("Copy").Index
    tbEdit.Buttons("Paste").Image = fMainForm.imlMenu.ListImages.item("Paste").Index
    tbEdit.Buttons("Undo").Image = fMainForm.imlMenu.ListImages.item("Undo").Index
    tbEdit.Buttons("Refresh").Image = fMainForm.imlMenu.ListImages.item("Undo").Index
    tbEdit.Buttons("Update").Image = fMainForm.imlMenu.ListImages.item("Redo").Index
    On Error GoTo 0
    
End Sub

Private Sub tbEdit_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.key
        Case "New"
            mnuFileNew_Click
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Save As"
            mnuFileSaveAs_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Undo"
            mnuEditUndo_Click
        Case "Refresh"
            mnuFileRefresh_Click
        Case "Update"
            mnuFileUpdate_Click
    End Select
End Sub
