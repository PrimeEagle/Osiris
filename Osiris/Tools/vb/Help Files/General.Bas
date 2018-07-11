Attribute VB_Name = "General"
Option Explicit

Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Any) As Long

Global Const HELP_QUIT = 2
Global Const HELP_INDEX = 3
Global Const HELP_HELPONHELP = 4
Global Const HELP_PARTIALKEY = &H105



Sub HelpFunction(lhWnd As Long, HelpCmd As Integer, HelpKey As String)
   
Dim lRtn As Long 'declare the needed variables
   
If HelpCmd = HELP_PARTIALKEY Then
   lRtn = WinHelp(lhWnd, Help.Filename.Text, HelpCmd, HelpKey)
Else
   lRtn = WinHelp(lhWnd, Help.Filename.Text, HelpCmd, 0&)
End If
   
End Sub

