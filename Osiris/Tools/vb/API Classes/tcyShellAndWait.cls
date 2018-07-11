VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tcyShellAndWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const SYNCHRONIZE = &H10000
Private Const INFINITE = &HFFFFFFFF

'local variable(s) to hold property value(s)
Private mvarCommandLine As String 'local copy

Public Property Let CommandLine(ByVal vData As String)
    Dim Res As Long
    
    If (vData <> "") Then
        mvarCommandLine = vData
        
        Res = Shell(mvarCommandLine, vbNormalFocus)
        
        If Res <> 0 Then
            Wait Res
        End If
    End If
End Property


Public Property Get CommandLine() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CommandLine
    CommandLine = mvarCommandLine
End Property

Private Sub Wait(pid As Long)
    Dim Res As Long
    
    Res = OpenProcess(SYNCHRONIZE, 0, pid)
    
    If Res <> 0 Then
        WaitForSingleObject Res, INFINITE
        CloseHandle Res
    End If
End Sub
