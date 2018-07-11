Attribute VB_Name = "mFiles"
Option Explicit

Public Function FileExtension(filename As String) As String
    Dim pos As Long
    
    'if string doesn't end in .XXX, then treat as a folder
    pos = Len(filename)
    While InStr(pos, filename, ".", vbTextCompare) = 0 And pos > 1
        pos = pos - 1
    Wend
    If pos = 1 Then
        If UCase(Mid$(filename, 1, 4)) = "FTP:" Then
            FileExtension = "FTP"
        ElseIf UCase(Mid$(filename, 1, 5)) = "HTTP:" Then
            FileExtension = "HTTP"
        Else
            FileExtension = "INVALID"
        End If
    Else
        FileExtension = UCase(Mid$(filename, pos + 1, Len(filename) - pos))
    End If
End Function

Public Sub ClearTempDir(TempFolder As String)
    Dim CurrentFile As String
    
    CurrentFile = Dir(TempFolder)
    
    While CurrentFile <> ""
        Kill TempFolder & CurrentFile
        CurrentFile = Dir
    Wend
End Sub

Public Function GetAvailableTempFile(TempFolder As String) As Long
    Dim CurrentFile As String
    Dim i As Long
    
    i = 1
    CurrentFile = Dir(TempFolder & Format$(i))
    
    While CurrentFile <> ""
        i = i + 1
        CurrentFile = Dir(TempFolder & Format$(i))
    Wend
    GetAvailableTempFile = i
End Function
Public Function DeleteTempFile(TempFolder As String, filenum As Long)
    Dim tempstr As String
    
    tempstr = Dir(TempFolder & Format$(filenum))
    If tempstr <> "" Then
        Kill TempFolder & Format$(filenum)
    End If
End Function

Public Function GetFileNameFromPath(filename As String) As String
    Dim i As Long
    Dim lastpos As Long
    
    If Left$(filename, 1) = "\" Then
        lastpos = 1
    Else
        lastpos = 0
    End If
    
    i = InStr(1, filename, "\", vbTextCompare)
    While i <> 0
        lastpos = i
        i = InStr(i + 1, filename, "\", vbTextCompare)
    Wend
    
    GetFileNameFromPath = Mid$(filename, lastpos + 1, Len(filename) - lastpos)

End Function

Public Function ValidFile(file As String) As Boolean
    On Error GoTo FileError
    If Dir(file) = "" Then
        ValidFile = False
    Else
        ValidFile = True
    End If
    Exit Function
FileError:
    ValidFile = False
End Function

Public Function ValidFolder(folder As String) As Boolean
    If Dir(folder, vbDirectory) = "" Then
        ValidFolder = False
    Else
        ValidFolder = True
    End If
End Function
