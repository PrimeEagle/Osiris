VERSION 5.00
Begin VB.Form frmQuotes 
   Caption         =   "Quote of the Day"
   ClientHeight    =   3285
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5520
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "&Show Quotes at Startup"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   2940
      Width           =   2055
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "Random &Quote"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   120
      ScaleHeight     =   2655
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sun Tzu said ..."
         Height          =   255
         Left            =   540
         TabIndex        =   5
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   1635
         Left            =   180
         TabIndex        =   4
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmQuotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' The in-memory database of tips.
Dim Tips As New Collection

' Name of tips file
Const TIP_FILE = "QUOTES.TXT"

' Index in collection of tip currently being displayed.
Dim CurrentTip As Long
'Dim ShowAtStartup As Long



Private Sub DoNextTip()

    ' Select a tip at random.
    CurrentTip = Int((Tips.count * Rnd) + 1)
    
    ' Or, you could cycle through the Tips in order

'    CurrentTip = CurrentTip + 1
'    If Tips.Count < CurrentTip Then
'        CurrentTip = 1
'    End If
    
    ' Show it.
    frmQuotes.DisplayCurrentTip
    
End Sub

Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' Each tip read in from file.
    Dim InFile As Integer   ' Descriptor for file.
    
    ' Obtain the next free file descriptor.
    InFile = FreeFile
    
    ' Make sure a file is specified.
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Make sure the file exists before trying to open it.
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Read the collection from a text file.
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' Display a tip at random.
    DoNextTip
    
    LoadTips = True
    
End Function


Private Sub cmdNextTip_Click()
    DoNextTip
End Sub

Private Sub cmdOK_Click()
    SaveSetting App.EXEName, "Options", "Show Tips at Startup", _
        chkLoadTipsAtStartup.Value
    Unload Me
End Sub

Private Sub Form_Load()
        
    ' Set the checkbox, this will force the value to be written back out to the registry
    
    chkLoadTipsAtStartup.Value = GetSetting(App.EXEName, "Options", _
            "Show Tips at Startup")
    
    ' Seed Rnd
    Randomize
    
    ' Read in the tips file and display a tip at random.
    If LoadTips(App.Path & "\" & TIP_FILE) = False Then
        lblTipText.Caption = "That the " & TIP_FILE & " file was not found? " & vbCrLf & vbCrLf & _
           "Create a text file named " & TIP_FILE & " using NotePad with 1 tip per line. " & _
           "Then place it in the same directory as the application. "
    End If
End Sub

Public Sub DisplayCurrentTip()
    If Tips.count > 0 Then
        lblTipText.Caption = Tips.item(CurrentTip)
    End If
End Sub
