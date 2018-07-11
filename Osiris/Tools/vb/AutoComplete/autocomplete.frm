VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   ClientHeight    =   7515
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   7335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   7815
   End
   Begin VB.Menu SaveAs 
      Caption         =   "Save As"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dbase As Database
Dim currentdatabasefile As String

Dim Inserted As Boolean
Dim DontChange As Boolean
Dim DontChange2 As Boolean
Dim lastcurrentWordLen As Integer
Dim lastLeftPos As Integer
Dim lastPos As Integer
Dim InsertedString As String
Dim InsertedEndPos As Integer

Private Sub Form_Load()
    Dim tempstr As String
    
    currentdatabasefile = "db1.mdb"
    OpenDBase dbase, currentdatabasefile
    
    Inserted = False
    DontChange = False
    DontChange2 = False
    lastcurrentWordLen = 0
    lastLeftPos = 1
    lastPos = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    dbase.Close
End Sub

Private Sub SaveAs_Click()
    CommonDialog1.ShowSave
    If CommonDialog1.filename <> "" Then
        Open CommonDialog1.filename For Output As #1    ' Open file for output.
        Print #1, Text1.Text  ' Print text to file.
        Close #1
    End If
End Sub

Private Sub Text1_Change()
    Dim currentPos As Integer
    Dim currentChar As String
    Dim currentWord As String
    Dim currentText As String
    Dim i As Integer
    Dim record As Recordset
    
    If DontChange Then
        DontChange = False
        DontChange2 = True
        Exit Sub
    End If
    
    If DontChange2 Then
        DontChange2 = False
        Exit Sub
    End If
    
    currentPos = Text1.SelStart
    
    If currentPos > 0 Then
        If currentPos > lastPos Then
            currentChar = Mid$(Text1.Text, currentPos, 1)
            If currentChar = " " Or _
                    currentChar = Chr(10) Or _
                    currentChar = Chr(9) Then
                    'for right going space, carriage return, or tab char
                lastLeftPos = currentPos + 1
                currentWord = ""
            Else
                currentWord = Mid$(Text1.Text, lastLeftPos, currentPos - lastLeftPos + 1)
            End If
            If Inserted Then
                If Len(currentWord) > 0 And InStr(1, UCase$(InsertedString), _
                        UCase$(currentWord), vbTextCompare) = 1 Then
                    Text1.SelStart = currentPos     'current char is a match also
                    Text1.SelLength = 1
                    DontChange2 = True
                    Text1.SelText = ""
                    If Len(currentWord) = Len(InsertedString) Then
                        Inserted = False
                    End If
                Else        'current char is not a match so delete inserted text
                    Text1.SelStart = currentPos
                    Text1.SelLength = Len(InsertedString) - Len(currentWord) + 1
                    DontChange2 = True
                    Text1.SelText = ""
                    Inserted = False
                End If
            End If
            If Not Inserted Then
                If Len(currentWord) > 1 Then    'chk if need to insert text
                    Set record = dbase.OpenRecordset("SELECT TOP 1 Field1 FROM " & _
                            "Table1 WHERE Field1 Like '" & currentWord & _
                            "*' ORDER BY Field1;", dbOpenDynaset)
                    If Not record.EOF Then
                        InsertedString = record!Field1
                        Text1.SelStart = lastLeftPos - 1 ' set selection start
                        Text1.SelLength = Len(currentWord)
                        InsertedEndPos = lastLeftPos - 1 + Len(InsertedString)
                        DontChange = True
                        Text1.SelText = InsertedString
                        Text1.SelStart = currentPos ' set selection start
                        Inserted = True
                    End If
                    record.Close
                End If
            End If
        Else        'if backspacing
            Inserted = False
            If currentPos < lastLeftPos - 1 Then    'if lost left space pos
                lastLeftPos = currentPos
                currentChar = Mid$(Text1.Text, lastLeftPos, 1)
                lastLeftPos = lastLeftPos - 1
                While currentChar <> " " And currentChar <> Chr(10) _
                        And currentChar <> Chr(9)
                    If lastLeftPos < 1 Then
                        GoTo LeftFloor
                    End If
                    currentChar = Mid$(Text1.Text, lastLeftPos, 1)
                    lastLeftPos = lastLeftPos - 1
                Wend
LeftOK:
                lastLeftPos = lastLeftPos + 2
                GoTo LeftDone
LeftFloor:
                lastLeftPos = 1
            End If
LeftDone:
            currentWord = Mid$(Text1.Text, lastLeftPos, currentPos - lastLeftPos + 1)
        End If
    Else
        lastLeftPos = 1
        currentWord = ""
    End If
    
    lastPos = currentPos
    lastcurrentWordLen = Len(currentWord)
    'Debug.Print currentWord

End Sub

Public Function OpenDBase(db As Database, databasefile As String) As Boolean
    currentdatabasefile = databasefile ' uses default, or the one returned by the open file dialog box
    On Error GoTo Err_OpenDB
    Set db = OpenDatabase(databasefile)
    On Error GoTo 0
    OpenDBase = True
    GoTo Done

Err_OpenDB:
    MsgBox "OpenDatabase unsuccessful!"
    OpenDBase = False
    
Done:

End Function

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab And Inserted Then
        Text1.SelStart = InsertedEndPos
        Inserted = False
        DontChange = True
        SendKeys "{BACKSPACE}"
    End If
End Sub
