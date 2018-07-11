Attribute VB_Name = "mDataType"
Option Explicit

Public Function DataIsValid(data_type As String, data As String) As Boolean
    Dim tempnum As Long
    
    Select Case data_type
        Case "Boolean"
            If data = "" Or _
                    Trim(UCase(data)) = "TRUE" Or _
                    Trim(UCase(data)) = "FALSE" Or _
                    Trim(UCase(data)) = "YES" Or _
                    Trim(UCase(data)) = "NO" Or _
                    Trim(UCase(data)) = "T" Or _
                    Trim(UCase(data)) = "F" Or _
                    Trim(UCase(data)) = "1" Or _
                    Trim(UCase(data)) = "0" Then
                DataIsValid = True
            Else
                DataIsValid = False
            End If
        Case "Number"
            If data = "" Or IsNumeric(data) Then
                DataIsValid = True
            Else
                DataIsValid = False
            End If
        Case "String"
            DataIsValid = True
        Case "Date"
            If data = "" Or IsDate(data) Then
                DataIsValid = True
            Else
                DataIsValid = False
            End If
        Case "URL"
            If UCase(Left$(data, 7)) = "HTTP://" Or _
                UCase(Left$(data, 6)) = "FTP://" Then
                DataIsValid = True
            Else
                DataIsValid = False
            End If
        Case Else
            MsgBox "DataIsValid:  Data type not recognized."
            DataIsValid = False
    End Select
End Function

Public Function ValidateData(data_type As String, data As String) As String
    Select Case data_type
        Case "Boolean"
            If data = "" Then
                ValidateData = ""
            Else
                If Trim(UCase(data)) = "TRUE" Or _
                        Trim(UCase(data)) = "T" Or _
                        Trim(UCase(data)) = "YES" Or _
                        Trim(UCase(data)) = "1" Then
                    ValidateData = "True"
                Else
                    ValidateData = "False"
                End If
            End If
        Case "Number"
            If data = "" Then
                ValidateData = ""
            Else
                ValidateData = CDbl(data)
            End If
        Case "String", "URL"
            ValidateData = data
        Case "Date"
            If data = "" Then
                ValidateData = ""
            Else
                ValidateData = CDate(data)
            End If
        Case Else
            MsgBox "ValidateData:  Data type not recognized."
    End Select
End Function


