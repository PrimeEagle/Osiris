VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCopyMoveEntireFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nSize As Long, ByVal lpBuffer As String) As Long

Enum TAction
    vbMove = 1
    vbCopy = 2
End Enum

Private Type SHFILEOPSTRUCT
    hWnd               As Long
    wFunc              As Long
    pFrom              As String
    pTo                As String
    fFlags             As Integer
    fAborted           As Boolean
    hNameMaps          As Long
    sProgress          As String
End Type
  
Const FO_MOVE = &H1
Const FO_COPY = &H2
Const FOF_SILENT = &H4
Const FOF_RENAMEONCOLLISION = &H8
Const FOF_NOCONFIRMATION = &H10
Const FOF_SIMPLEPROGRESS = &H100
Const FOF_ALLOWUNDO = &H40
  
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

'local variable(s) to hold property value(s)
Private mvarAction As Variant 'local copy
Private mvarSourcePath As String 'local copy
Private mvarDestinationPath As String 'local copy
Private mvarShowProgress As Boolean 'local copy
Private mvarShowFileNames As Boolean 'local copy
Private mvarConfirmation As Boolean 'local copy
Private mvarRenameExisting As Boolean 'local copy

Public Function Execute() As Boolean
    If mvarSourcePath <> "" And mvarDestinationPath <> "" Then
        Execute = (PerformShellAction = 0)
    Else
        MsgBox "SourcePath and/or DestinationPath property(s) have not been specified.", vbOKOnly + vbCritical, "Invalid property"
        Execute = False
    End If
End Function

Public Property Let RenameExisting(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RenameExisting = 5
    mvarRenameExisting = vData
End Property


Public Property Get RenameExisting() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RenameExisting
    RenameExisting = mvarRenameExisting
End Property



Public Property Let Confirmation(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Confirmation = 5
    mvarConfirmation = vData
End Property


Public Property Get Confirmation() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Confirmation
    Confirmation = mvarConfirmation
End Property



Public Property Let ShowFileNames(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ShowFileNames = 5
    mvarShowFileNames = vData
End Property


Public Property Get ShowFileNames() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ShowFileNames
    ShowFileNames = mvarShowFileNames
End Property



Public Property Let ShowProgress(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ShowProgress = 5
    mvarShowProgress = vData
End Property


Public Property Get ShowProgress() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ShowProgress
    ShowProgress = mvarShowProgress
End Property



Public Property Let DestinationPath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DestinationPath = 5
    mvarDestinationPath = vData
End Property


Public Property Get DestinationPath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DestinationPath
    DestinationPath = mvarDestinationPath
End Property



Public Property Let SourcePath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SourcePath = 5
    mvarSourcePath = vData
End Property


Public Property Get SourcePath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SourcePath
    SourcePath = mvarSourcePath
End Property



Public Property Let Action(ByVal vData As TAction)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Action = 5
    mvarAction = vData
End Property


Public Property Get Action() As TAction
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Action
    Action = mvarAction
End Property

Private Sub Class_Initialize()
    mvarAction = vbCopy
    mvarSourcePath = ""
    mvarDestinationPath = ""
    mvarShowProgress = True
    mvarShowFileNames = True
    mvarConfirmation = True
    mvarRenameExisting = False
End Sub

Private Function PerformShellAction() As Long

   Dim FOF_FLAGS As Long
   Dim SHFileOp  As SHFILEOPSTRUCT
   
  'terminate the folder string with a pair of nulls
   mvarSourcePath = mvarSourcePath & Chr$(0) & Chr$(0)
  
  'determine the user's options selected
   FOF_FLAGS = BuildBrowseFlags()
  
  'set up the options
   With SHFileOp
      .wFunc = mvarAction
      .pFrom = mvarSourcePath
      .pTo = mvarDestinationPath
      .fFlags = FOF_FLAGS
   End With
  
  'and perform the chosen copy or move operation
   PerformShellAction = SHFileOperation(SHFileOp)
End Function

Private Function BuildBrowseFlags() As Long
  Dim flag As Long
 
  flag = 0&
   
  If mvarShowProgress Then flag = flag Or FOF_SILENT
  If mvarShowFileNames Then flag = flag Or FOF_SIMPLEPROGRESS
  If mvarConfirmation Then flag = flag Or FOF_NOCONFIRMATION
  If mvarRenameExisting Then flag = flag Or FOF_RENAMEONCOLLISION
  
  BuildBrowseFlags = flag
End Function
