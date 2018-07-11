VERSION 5.00
Begin VB.Form frmCMsgBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custom Message Box"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5415
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cb3 
      Caption         =   "3"
      Height          =   345
      Left            =   3960
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cb2 
      Caption         =   "2"
      Height          =   345
      Left            =   2280
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cb1 
      Caption         =   "1"
      Height          =   345
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Image imgCritical 
      Height          =   480
      Left            =   1560
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgExclamation 
      Height          =   480
      Left            =   1080
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgQuestion 
      Height          =   480
      Left            =   600
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgInformation 
      Height          =   480
      Left            =   120
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgCMsgBox 
      Height          =   735
      Left            =   240
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblCMsgBox 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmCMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const BUTTON_GAP = 250      'twips, distance between two buttons

Const BUTTON_BORDER = 200   ' twips, distance between bottom
                            ' of form to bottom of button

Const MAX_FORM_WIDTH = 5500 'twips

Const MAX_MSG_WIDTH = 4000  'twips

Dim ButtonPressed As Integer
Dim HasIcon As Boolean


Private Sub cb1_Click()
    ButtonPressed = 1
    CalculateResponse
    Unload Me
End Sub

Private Sub cb2_Click()
    ButtonPressed = 2
    CalculateResponse
    Unload Me
End Sub

Private Sub cb3_Click()
    ButtonPressed = 3
    CalculateResponse
    Unload Me
End Sub

Private Sub Form_Load()
    'Set up custom buttons
    'Set up custom icons
    Select Case CMsgBox_WhichIcon
        Case vbInformation
            'set icon here
            imgCMsgBox.picture = imgInformation.picture
            HasIcon = True
        Case vbCritical
            'set icon here
            imgCMsgBox.picture = imgCritical.picture
            HasIcon = True
        Case vbExclamation
            'set icon here
            imgCMsgBox.picture = imgExclamation.picture
            HasIcon = True
        Case vbQuestion
            'set icon here
            imgCMsgBox.picture = imgQuestion.picture
            HasIcon = True
        Case Else
            If CMsgBox_IconIndex < 1 Then
                imgCMsgBox.picture = Nothing
                HasIcon = False
            Else
                imgCMsgBox.picture = CMsgBox_ImgList.ListImages(CMsgBox_IconIndex).picture
                HasIcon = True
            End If
    End Select
    
    Select Case CMsgBox_WhichButtons
        Case vbRetryCancel
            cb1.Caption = "Retry"
            cb2.Caption = "Cancel"
            cb2.Visible = True
            cb2.Cancel = True
            SetSizes (2)
        Case vbAbortRetryIgnore
            cb1.Caption = "Abort"
            cb2.Caption = "Retry"
            cb2.Visible = True
            cb3.Caption = "Ignore"
            cb3.Visible = True
            SetSizes (3)
        Case vbYesNoCancel
            cb1.Caption = "Yes"
            cb2.Caption = "No"
            cb2.Visible = True
            cb3.Caption = "Cancel"
            cb3.Visible = True
            cb3.Cancel = True
            SetSizes (3)
        Case vbYesNo
            cb1.Caption = "Yes"
            cb2.Caption = "No"
            cb2.Visible = True
            SetSizes (2)
        Case vbCancel
            cb2.Caption = "Cancel"
            cb2.Visible = True
            cb2.Cancel = True
            SetSizes (2)
        Case vbOKOnly
            cb1.Caption = "OK"
            cb2.Visible = False
            cb3.Visible = False
            SetSizes (1)
    End Select
    
    'Set up custom alignment
    Select Case CMsgBox_WhichAlignment
        Case vbMsgBoxRight
            lblCMsgBox.Alignment = 1 ' Right
        Case Else
            lblCMsgBox.Alignment = 0 'Left
    End Select
    
    lblCMsgBox.Caption = CMsgBox_Text
    Me.Caption = CMsgBox_Title
    
    'Set up custom for default buttons
    Select Case CMsgBox_WhichDefault
        Case vbDefaultButton2
            cb2.TabIndex = 1
            cb3.TabIndex = 2
            cb1.TabIndex = 3
        Case vbDefaultButton3
            cb3.TabIndex = 1
            cb1.TabIndex = 2
            cb2.TabIndex = 3
        Case Else
            cb1.TabIndex = 1
            cb2.TabIndex = 2
            cb3.TabIndex = 3
    End Select
End Sub

Private Sub CalculateResponse()
    Select Case CMsgBox_WhichButtons
        Case vbRetryCancel
            Select Case ButtonPressed
                Case 1
                    CMsgBox_Response = vbRetry
                Case 2
                    CMsgBox_Response = vbCancel
            End Select
        Case vbAbortRetryIgnore
            Select Case ButtonPressed
                Case 1
                    CMsgBox_Response = vbAbort
                Case 2
                    CMsgBox_Response = vbRetry
                Case 3
                    CMsgBox_Response = vbIgnore
            End Select
        Case vbYesNoCancel
            Select Case ButtonPressed
                Case 1
                    CMsgBox_Response = vbYes
                Case 2
                    CMsgBox_Response = vbNo
                Case 3
                    CMsgBox_Response = vbCancel
            End Select
        Case vbYesNo
            Select Case ButtonPressed
                Case 1
                    CMsgBox_Response = vbYes
                Case 2
                    CMsgBox_Response = vbNo
            End Select
        Case vbCancel
            Select Case ButtonPressed
                Case 1
                    CMsgBox_Response = vbOK
                Case 2
                    CMsgBox_Response = vbCancel
            End Select
    End Select
End Sub

Private Sub AlignButtons(NumButtons As Integer)
    cb1.Top = lblCMsgBox.Top + lblCMsgBox.Height + BUTTON_BORDER
    cb2.Top = cb1.Top
    cb3.Top = cb1.Top
    
    Select Case NumButtons
        Case 1
            cb1.Left = Me.Width / 2 - cb1.Width / 2
        Case 2
            cb1.Left = Me.Width / 2 - cb1.Width - BUTTON_GAP / 2
            cb2.Left = Me.Width / 2 + BUTTON_GAP / 2
        Case 3
            cb2.Left = Me.Width / 2 - cb2.Width / 2
            cb1.Left = cb2.Left - cb1.Width - BUTTON_GAP
            cb3.Left = cb2.Left + cb2.Width + BUTTON_GAP
    End Select
End Sub

Private Sub SetSizes(NumButtons As Integer)
    Dim FormWidth As Long
    Dim FormHeight As Long
    Dim ButtonWidth As Long
    Dim MessageWidth As Long
    Dim MessageHeight As Long
    Dim MessageNumLines As Long
    Dim IconWidth As Long
    Dim LimitingWidth As Long
    Dim NumberCarriageReturns As Long
    Dim TempChar As String
    Dim TempMsg As String
    Dim i As Long
    
    ' the width of all the buttons, including gaps in between and
    ' on the outside edges
    ButtonWidth = (NumButtons) * cb1.Width _
                    + (NumButtons + 1) * BUTTON_GAP
                    
    ' the width of the entire text message, in twips, if it
    ' were written on a single line
    MessageWidth = TextWidth(CMsgBox_Text)
    
    ' the height of the entire text message, in twips, if it
    ' were written on a single line
    MessageNumLines = 1
    
    TempMsg = ""
    For i = 1 To Len(CMsgBox_Text)
        TempChar = mID$(CMsgBox_Text, i, 1)
        If TempChar = Chr(13) Then
            NumberCarriageReturns = NumberCarriageReturns + 1
        Else
            TempMsg = TempMsg & TempChar
        End If
    Next i
    MessageHeight = TextHeight(TempMsg)
   
    'the limiting width is whichever is going to take up the most space,
    'the width of the buttons, or the width of the message & icon.
    If ButtonWidth > MessageWidth Then
        LimitingWidth = ButtonWidth
    Else
        LimitingWidth = MessageWidth
    End If
    
    'calculate the necessary width of the form, taking into account
    'the message width, the width of an icon, if present, and the
    'necessary gaps between them.
    If HasIcon Then
        IconWidth = imgCMsgBox.Width
        FormWidth = 3 * BUTTON_GAP + IconWidth + LimitingWidth
    Else
        IconWidth = 0
        FormWidth = 2 * BUTTON_GAP + LimitingWidth
    End If
    
    If FormWidth > MAX_FORM_WIDTH Or NumberCarriageReturns <> 0 Then
        'this means that the message is going to take up more than one line
        If FormWidth > MAX_FORM_WIDTH Then
            FormWidth = MAX_FORM_WIDTH
        End If
        
        'how many lines will the msg take up?
        MessageWidth = TextWidth(TempMsg)
        MessageNumLines = (MessageWidth \ MAX_MSG_WIDTH) + NumberCarriageReturns
        
        'the previous division truncates the remainder,
        'so if there was a remainder, we need add one more line
        'to accomodate it.
        If MessageWidth Mod MAX_MSG_WIDTH <> 0 Then
            MessageNumLines = MessageNumLines + 1
        End If
        
        'the total height of the message is the number of lines
        'times the height of each line.
        MessageHeight = MessageNumLines * MessageHeight
        
        'the width is always the max width in this case,
        'since it takes up more than one line.
        MessageWidth = MAX_MSG_WIDTH
    End If
    
    lblCMsgBox.Width = MessageWidth
    lblCMsgBox.Height = MessageHeight
    Me.Width = FormWidth
    
    'align the buttons in the newly sized form
    AlignButtons (NumButtons)
    Me.Height = cb1.Top + cb1.Height + BUTTON_BORDER + 300
End Sub
