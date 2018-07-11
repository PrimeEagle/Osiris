Attribute VB_Name = "mProgress_Bar"
Option Explicit

Public Sub InitProgressBar(F As Form, label As String, min As Long, max As Long, _
            picture As String, Optional bar_visible As Boolean = True, _
            Optional AbortButton As Boolean = False)
        
        F.pbPBar1.Visible = bar_visible
        F.lblPBar.Caption = label
        F.pbPBar1.min = min
        F.pbPBar1.max = max
        On Error Resume Next    'if picture does not exist, blank picture
        F.picProgBar.picture = LoadPicture(picture)
        On Error GoTo 0
        F.pbPBar1.Value = F.pbPBar1.min
        
        F.Show vbModeless
        
        If AbortButton Then
            If Not F.cbAbort.Visible Then F.cbAbort.Visible = True
        Else
            If F.cbAbort.Visible Then F.cbAbort.Visible = False
        End If
        
        DoEvents
End Sub

