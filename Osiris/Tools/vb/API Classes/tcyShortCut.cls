VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsShortCut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const SHARD_PATH = &H2&
Private Const FO_MOVE = &H1
Private Const FO_RENAME = &H4
Private Const FOF_SILENT = &H4
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_FILESONLY = &H80
Private Const FOF_NOCONFIRMMKDIR = &H200

Private Type SHFILEOPSTRUCT
      hWnd         As Long
      wFunc        As Long
      pFrom        As String
      pTo          As String
      fFlags       As Integer
      fAborted     As Boolean
      hNameMaps    As Long
      sProgress    As String
End Type
     
Private Type SHITEMID
       cb As Long
       abID As Byte
End Type
   
Private Type ITEMIDLIST
       mkid As SHITEMID
End Type
   
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function SHAddToRecentDocs Lib "shell32.dll" (ByVal dwFlags As Long, ByVal dwData As String) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
     
'local variable(s) to hold property value(s)
Private mvarFileName As String 'local copy
Private mvarCommandLine As String 'local copy
Private mvarDestinationPath As String 'local copy
Private mvarCaption As String 'local copy

Public Property Let Caption(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Caption = 5
    mvarCaption = vData
End Property


Public Property Get Caption() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Caption
    Caption = mvarCaption
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




Public Property Let CommandLine(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CommandLine = 5
    mvarCommandLine = vData
End Property


Public Property Get CommandLine() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CommandLine
    CommandLine = mvarCommandLine
End Property



Public Function Create() As Boolean
   Dim r As Long
   Dim i As Integer
   
   Dim FolderPath As String
   Dim StartMenuPath As String
   Dim fNameOld As String
   Dim fNameNew As String
   
   On Error GoTo Create_Handler
   
   Const CSIDL_RECENT = &H8
   mvarFileName = strFileName(mvarCommandLine)
   Create = False
   
   FolderPath = GetSpecialFolder(CSIDL_RECENT)
   
   If FolderPath = "" Then
      Create = False
      Exit Function
   End If
  
   r = SHAddToRecentDocs(SHARD_PATH, mvarCommandLine)
   DoEvents
  
   ShellMoveFiles FolderPath & mvarFileName & ".lnk", mvarDestinationPath
   DoEvents
  
   fNameOld = mvarDestinationPath & mvarFileName & ".lnk"
   fNameNew = mvarDestinationPath & mvarCaption & ".lnk"
   ShellRenameFile fNameOld, fNameNew
   DoEvents
   
   Create = True
   
   Exit Function

Create_Handler:
    Err.Clear
    Create = False
End Function



Private Sub ShellMoveFiles(sFile As String, sDestination As String)
  'set some working variables
   Dim r As Long
   Dim i As Integer
   Dim sFiles As String
   Dim SHFileOp As SHFILEOPSTRUCT
  
  
  'add a final terminating null
   sFile = sFile & Chr$(0)
  
  'for debugging - print the resulting strings
  'Print sFiles
  
  'set up the options
   With SHFileOp
      .wFunc = FO_MOVE
      .pFrom = sFile
      .pTo = sDestination
      .fFlags = FOF_SILENT Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR
   End With
  
  'and perform the move.  Because the folder specified
  'doesn't exist, SHFileOperation will create it. FOF_SILENT
  'above instructs the API to suppress displaying the "flying
  'folders" dialog during the move.  FOF_NOCONFIRMATION suppresses
  'prompting to move the files ... the "Are you sure you want to
  'move etc..." dialog. FOF_NOCONFIRMMKDIR instructs it to
  'create the folder without prompting if it's OK.
   r = SHFileOperation(SHFileOp)
End Sub



Private Function GetSpecialFolder(CSIDL As Long) As String
  
  'a few local variables needed
   Dim r As Long
   Dim sPath As String
   Dim IDL As ITEMIDLIST
   
   Const NOERROR = 0
   Const MAX_LENGTH = 260
  
  'fill the idl structure with the specified folder item
   r = SHGetSpecialFolderLocation(GetDesktopWindow(), CSIDL, IDL)
   
   If r = NOERROR Then
     
     'Of the structure is filled, initialize and
     'retrieve the path from the id list, and return
     'the folder with a trailing slash appended.
      sPath = Space$(MAX_LENGTH)
      r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
       
      If r Then
         GetSpecialFolder = Left$(sPath, _
                            InStr(sPath, Chr$(0)) - 1) & "\"
      End If
      
   End If

End Function



Private Sub ShellRenameFile(sOldName As String, sNewName As String)
  
  'set some working variables
   Dim SHFileOp As SHFILEOPSTRUCT
   Dim r As Long
  
  'add a pair of terminating nulls to each string
   sOldName = sOldName & Chr$(0) & Chr$(0)
   sNewName = sNewName & Chr$(0) & Chr$(0)
  

  'for debugging - print the resulting strings
   'Print sOldName
   'Print sNewName
  
  'set up the options
   With SHFileOp
      .wFunc = FO_RENAME
      .pFrom = sOldName
      .pTo = sNewName
      .fFlags = FOF_SILENT Or FOF_NOCONFIRMATION
   End With
  
  'and rename the file
   r = SHFileOperation(SHFileOp)

End Sub

Private Function strFileName(sPathFile As String) As String
    Dim Lp As Integer
    
    For Lp = Len(sPathFile) To 1 Step -1
        If (Mid$(sPathFile, Lp, 1) = "\") Then
            strFileName = Mid$(sPathFile, Lp + 1, Len(sPathFile))
            Exit For
        End If
    Next Lp
End Function
