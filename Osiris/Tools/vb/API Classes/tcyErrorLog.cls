VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tcyErrorLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Const vbMaxSize = 225

'Default Property Values:
Const m_def_FileName = ""
Const m_def_DateTimeStamp = True
Const m_def_Message = ""

'Property Variables:
Dim m_FileName As String
Dim m_DateTimeStamp As Boolean
Dim m_Message As String

Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Property Let FileName(ByVal New_FileName As String)
    m_FileName = New_FileName
End Property

Public Property Get DateTimeStamp() As Boolean
    DateTimeStamp = m_DateTimeStamp
End Property

Public Property Let DateTimeStamp(ByVal New_DateTimeStamp As Boolean)
    m_DateTimeStamp = New_DateTimeStamp
End Property

Public Function ClearLog() As Boolean
    Dim nFyle As Long
    
    ClearLog = False
    If Trim(m_FileName) <> "" Then
        On Error GoTo ClearLog_Handler
        nFyle = FreeFile
        
        Open m_FileName For Output As #nFyle
            If m_DateTimeStamp Then
                Print #nFyle, strDateTimeStamp() & ",Log Cleared"
            Else
                Print #nFyle, "Log Cleared"
            End If
        Close #nFyle
        
        ClearLog = True
    Else
        Err.Raise 52  'Bad file name or number
    End If
    
    Exit Function
    
ClearLog_Handler:
    m_Message = "Error#" & CStr(Err.Number) & ": " & Err.Description
    Err.Clear
    
End Function

Public Property Get Message() As String
    Message = m_Message
End Property

Public Property Let Message(ByVal New_Message As String)
    Dim nFyle As Long
    
    If Trim(m_FileName) <> "" Then
        m_Message = New_Message
        
        On Error GoTo Message_Handler
        
        nFyle = FreeFile
        Open m_FileName For Append As #nFyle
            If m_DateTimeStamp Then
                Print #nFyle, strDateTimeStamp() & "," & m_Message
            Else
                Print #nFyle, m_Message
            End If
        Close #nFyle
    Else
        Err.Raise 52  'Bad file name or number
    End If
    
    Exit Property
    
Message_Handler:
    m_Message = "Error#" & CStr(Err.Number) & ": " & Err.Description
    Err.Clear
End Property

Private Function strDateTimeStamp() As String
    strDateTimeStamp = Format(Date, "dd-mmm-yyyy") & "," & Format(Time, "hh:nn:ss")
End Function

Private Sub Class_Initialize()
    m_FileName = m_def_FileName
    m_DateTimeStamp = m_def_DateTimeStamp
    m_Message = m_def_Message
End Sub
