VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tcyConvertTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarDigital As String 'local copy
Private mvarHours As Long 'local copy
Private mvarMinutes As Long 'local copy
Private mvarSeconds As Long 'local copy
Private mvarElapsed As Long 'local copy

Public Property Let Elapsed(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Elapsed = 5
    mvarElapsed = vData
    
    BreakDown
    CDigitalFromElapsed mvarElapsed
End Property


Public Property Get Elapsed() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Elapsed
    Elapsed = mvarElapsed
End Property



Public Property Get Seconds() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Seconds
    Seconds = mvarSeconds
End Property



Public Property Get Minutes() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Minutes
    Minutes = mvarMinutes
End Property



Public Property Get Hours() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Hours
    Hours = mvarHours
End Property



Public Property Let Digital(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Digital = 5
    mvarDigital = vData
    
    CElapsedFromDigital mvarDigital
End Property


Public Property Get Digital() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Digital
    Digital = mvarDigital
End Property


Private Sub BreakDown()
    mvarHours = Fix(mvarElapsed \ 3600)
    mvarMinutes = Fix((mvarElapsed Mod 3600) \ 60)
    mvarSeconds = ((mvarElapsed Mod 3600) Mod 60)
End Sub

Private Sub CDigitalFromElapsed(Elapsed As Long)
    mvarDigital = Format$(mvarHours, "0000") & ":" & Format$(mvarMinutes, "00") & ":" & Format$(mvarSeconds, "00")
End Sub

Private Sub CElapsedFromDigital(Digital As String)
    Dim H As Long
    Dim M As Long
    Dim S As Long
    Dim Pos As Byte
    
    mvarElapsed = 0
    If IsValidDigital(Digital) Then
        Pos = InStr(Digital, ":")
        
        If Pos <> 0 Then
            H = CLng(Left$(Digital, Pos - 1)) * 3600
            M = CLng(Mid$(Digital, Pos + 1, 2)) * 60
            S = CLng(Right$(Digital, 2))
            
            mvarElapsed = H + M + S
        End If
    End If
    
    'Break elapsed into Hours, Minutes, and Seconds
    BreakDown
    
    'Convert everything into digital format
    CDigitalFromElapsed mvarElapsed
End Sub

Private Function IsValidDigital(ByVal Source As String) As Boolean
    Dim Hours     As Long
    Dim Minutes   As Long
    Dim Seconds   As Long
    Dim ColonPos1 As Byte
    Dim ColonPos2 As Byte
    Dim Result    As Boolean
    
    Hours = 0
    Minutes = 0
    Seconds = 0
    Result = False

    'Ok, let's separate the time values
    ColonPos1 = InStr(Source, ":")
    ColonPos2 = InStr(Mid(Source, ColonPos1 + 1, Len(Source)), ":")

    'Check for first and second colon
    If Not ((ColonPos1 = 0) Or (ColonPos2 = 0)) Then
        'Check for invalid minutes}
        If (ColonPos2 <> 1) Then
            ColonPos2 = ColonPos1 + ColonPos2

            Hours = CLng(Mid(Source, 1, ColonPos1 - 1))
            Minutes = CLng(Mid(Source, ColonPos1 + 1, ColonPos2 - (ColonPos1 + 1)))
            Seconds = CLng(Mid(Source, ColonPos2 + 1, Len(Source)))

            If (Minutes >= 0) And (Minutes <= 59) Then
                If (Seconds >= 0) And (Seconds <= 59) Then
                    Result = True
                End If
            End If
        End If
    End If
    
    IsValidDigital = Result
End Function

