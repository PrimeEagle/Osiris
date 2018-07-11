Attribute VB_Name = "mCMsgBox"
Option Explicit
    
Global CMsgBox_WhichButtons As Long
Global CMsgBox_WhichIcon As Long
Global CMsgBox_WhichDefault As Long
Global CMsgBox_WhichAlignment As Long
Global CMsgBox_Response As VbMsgBoxResult
Global CMsgBox_Text As String
Global CMsgBox_Title As String
Global CMsgBox_ImgList As ImageList
Global CMsgBox_IconIndex As Long
    

' Custom Message Box Function.
' The vbApplicationModal / vbSystemModal parameter is ignored.
Public Function CMsgBox(Prompt As String, _
                        Optional Buttons As VbMsgBoxStyle = vbOKOnly, _
                        Optional Title As String, _
                        Optional ParentForm As Form, _
                        Optional ImgList As ImageList, _
                        Optional IconIndex As Long = -1) As VbMsgBoxResult
    
    If Title = "" Then
        Title = App.Title
    End If
    
    Set CMsgBox_ImgList = ImgList
    CMsgBox_IconIndex = IconIndex
    
    CMsgBox_Title = Title
    CMsgBox_Text = Prompt
    
    ' Which buttons are used is stored in the first 3 bits, so a 7 = '111'
    ' will isolate all of them and remove the other data.
    CMsgBox_WhichButtons = 7 And Buttons
    
    'Icons are stored in bits 5-7, so a 112 = '1110000'
    'will isolate them and remove other data.
    CMsgBox_WhichIcon = 112 And Buttons
    
    'Default Buttons are stored in bits 9-10, so a 768 = '1100000000'
    'will isolate them and remove other data.
    CMsgBox_WhichDefault = 768 And Buttons
    
    'Alignment is stored in bit 20, so a 524,288 = '10000000000000000000'
    'will isolate them and remove other data.
    CMsgBox_WhichAlignment = 524288 And Buttons

    frmCMsgBox.Show vbModal, ParentForm
    
    CMsgBox = CMsgBox_Response
End Function
