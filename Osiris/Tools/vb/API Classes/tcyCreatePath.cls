VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tcyCreatePath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long

Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260

Private Type SECURITY_ATTRIBUTES
   nLength As Long
   lpSecurityDescriptor As Long
   bInheritHandle As Long
End Type

Public Function Create(ByVal CompleteDirectory As String) As Collection
Attribute Create.VB_UserMemId = 0
  'creates nested directories on the drive
  'included in the path by parsing the final
  'directory string into a directory array,
  'and looping through each to create the final path.
  
  'The path could be passed to this method as a
  'pre-filled array, reducing the code.
  
   Dim r As Long
   Dim SA As SECURITY_ATTRIBUTES
   Dim drivePart As String
   Dim newDirectory  As String
   Dim item As String
   Dim sfolders As New Collection
   Dim pos As Integer
   Dim x As Integer
   
  'must have a trailing slash for
  'the GetPart routine below
   If Right$(CompleteDirectory, 1) <> "\" Then
      CompleteDirectory = CompleteDirectory & "\"
   End If
  
  'if there is a drive in the string, get it
  'else, just use nothing - assumes current drive
   pos = InStr(CompleteDirectory, ":")

   If pos Then
         drivePart = GetPart(CompleteDirectory, "\")
   Else
         drivePart = ""
   End If

  'now get the rest of the items that
  'make up the string
   Do Until CompleteDirectory = ""
    'strip off one item (i.e. "Files\")
     item = GetPart(CompleteDirectory, "\")

    'add it to an array for later use, and
    'if this is the first item (x=0),
    'append the drivepart

     If x = 0 Then item = drivePart & item
     sfolders.Add item

    'debug only
     Debug.Print item

    'increment the array counter
     x = x + 1
   Loop

  'Now create the directories.
  'Because the first directory is
  '0 in the array, reinitialise x to -1
   x = 0
   
   Do
      x = x + 1
     'just keep appending the folders in the
     'array to newDirectory.  When x=0 ,
     'newDirectory is "", so the
     'newDirectory gets assigned drive:\firstfolder.
     
     'Subsequent loops adds the next member of the
     'array to the path, forming a fully qualified
     'path to the new directory.
      newDirectory = newDirectory & sfolders(x)
      
     'the only member of the SA type needed (on
     'a win95/98 system at least)
      SA.nLength = LenB(SA)
      
      Call CreateDirectory(newDirectory, SA)
      
     'debug only
      Debug.Print newDirectory
   Loop Until x = (sfolders.Count - 1)
   
  'done. Return x, but add 1 for the 0-based array.
   Set CreateNestedFoldersByPath = sfolders
   
   Set sfolders = Nothing
End Function

Private Function GetPart(startStrg As String, delimiter As String) As String
'takes a string separated by "delimiter",
'splits off 1 item, and shortens the string
'so that the next item is ready for removal.

  Dim c As Integer
  Dim item As String
  
  c = 1
  
  Do
    If Mid$(startStrg, c, 1) = delimiter Then
      item = Mid$(startStrg, 1, c)
      startStrg = Mid$(startStrg, c + 1, Len(startStrg))
      GetPart = item
      Exit Function
    End If

    c = c + 1
  Loop
End Function

