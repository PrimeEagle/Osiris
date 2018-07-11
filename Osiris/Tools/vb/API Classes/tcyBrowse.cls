VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function SHSimpleIDListFromPath Lib "shell32" Alias "#162" (ByVal szPath As String) As Long
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrcpyA Lib "kernel32" (lpString1 As Any, lpString2 As Any) As Long
Private Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long

Private Type SHITEMID 'mkid
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST 'idl
    mkid As SHITEMID
End Type

Private Type BROWSEINFO 'bi
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const MAX_PATH = 260
Private Const WM_USER = &H400
Private Const BFFM_INITIALIZED = 1
Private Const BFFM_SETSTATUSTEXTA As Long = (WM_USER + 100)
Private Const BFFM_SETSTATUSTEXTW As Long = (WM_USER + 104)
Private Const BFFM_ENABLEOK As Long = (WM_USER + 101)
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)
Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

'local variable(s) to hold property value(s)
Private mvarDescription As String 'local copy
Private mvarPath As String 'local copy

Public Sub Show()
    If mvarPath = "" Then
        mvarPath = BrowseDefault
    Else
        mvarPath = BrowseFrom
    End If
End Sub

Public Property Let Path(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Path = 5
    mvarPath = vData
End Property

Public Property Get Path() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Path
    Path = mvarPath
End Property



Public Property Let Description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
End Property


Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property



Private Sub Class_Initialize()
    mvarPath = ""
    mvarDescription = ""
End Sub

Private Function BrowseDefault() As String
    Dim BI    As BROWSEINFO
    Dim IDL   As ITEMIDLIST
    Dim pidl  As Long
    Dim r     As Long
    Dim pos   As Integer
    Dim sPath As String

    BI.hOwner = App.hInstance

    BI.pidlRoot = 0&

    BI.lpszTitle = mvarDescription

    BI.ulFlags = BIF_RETURNONLYFSDIRS

    pidl& = SHBrowseForFolder(BI)

    sPath$ = Space$(512)
    r = SHGetPathFromIDList(ByVal pidl&, ByVal sPath$)

    If r Then
        pos = InStr(sPath$, Chr$(0))
        mvarPath = Left(sPath$, pos - 1)
    Else
        mvarPath = ""
    End If
    
    BrowseDefault = mvarPath
End Function

Private Function BrowseFrom() As String
    Dim BI        As BROWSEINFO
    Dim pidl      As Long
    Dim lpSelPath As Long
    Dim sPath     As String * MAX_PATH
    
    With BI
      .hOwner = App.hInstance
      .pidlRoot = 0
      .lpszTitle = mvarDescription
      .lpfn = FARPROC(AddressOf BrowseCallbackProcStr)
      
      lpSelPath = LocalAlloc(LPTR, Len(mvarPath))
      MoveMemory ByVal lpSelPath, ByVal mvarPath, Len(mvarPath)
      .lParam = lpSelPath
    End With
      
     pidl = SHBrowseForFolder(BI)
     
     If pidl Then
       
        If SHGetPathFromIDList(pidl, sPath) Then
           BrowseFrom = Left$(sPath, InStr(sPath, vbNullChar) - 1)
        Else
          BrowseFrom = ""
        End If
        
        Call CoTaskMemFree(pidl)
     
     End If
     
    Call LocalFree(lpSelPath)
End Function

Private Function BrowseCallbackProcStr(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
   Select Case uMsg
      Case BFFM_INITIALIZED
      
         Call SendMessage(hWnd, BFFM_SETSELECTIONA, True, ByVal StrFromPtrA(lpData))
                          
         Case Else
         
   End Select
End Function

Private Function FARPROC(pfn As Long) As Long
  
  FARPROC = pfn

End Function

