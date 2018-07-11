VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOpenDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Const OFN_ALLOWMULTISELECT = &H200
Const OFN_CREATEPROMPT = &H2000
Const OFN_ENABLEHOOK = &H20
Const OFN_ENABLETEMPLATE = &H40
Const OFN_ENABLETEMPLATEHANDLE = &H80
Const OFN_EXPLORER = &H80000
Const OFN_EXTENSIONDIFFERENT = &H400
Const OFN_FILEMUSTEXIST = &H1000
Const OFN_HIDEREADONLY = &H4
Const OFN_LONGNAMES = &H200000
Const OFN_NOCHANGEDIR = &H8
Const OFN_NODEREFERENCELINKS = &H100000
Const OFN_NOLONGNAMES = &H40000
Const OFN_NONETWORKBUTTON = &H20000
Const OFN_NOREADONLYRETURN = &H8000
Const OFN_NOTESTFILECREATE = &H10000
Const OFN_NOVALIDATE = &H100
Const OFN_OVERWRITEPROMPT = &H2
Const OFN_PATHMUSTEXIST = &H800
Const OFN_READONLY = &H1
Const OFN_SHAREAWARE = &H4000
Const OFN_SHAREFALLTHROUGH = 2
Const OFN_SHAREWARN = 0
Const OFN_SHARENOWARN = 1
Const OFN_SHOWHELP = &H10
Const OFS_MAXPATHNAME = 128

'OFS_FILE_OPEN_FLAGS and OFS_FILE_SAVE_FLAGS below
'are mine to save long statements; they're not a standard Win95 type.
Private Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_CREATEPROMPT Or OFN_NODEREFERENCELINKS Or OFN_HIDEREADONLY
Private Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY

Private Type OPENFILENAME
    nStructSize As Long
    hwndOwner As Long
    hInstance As Long
    sFilter As String
    sCustomFilter As String
    nCustFilterSize As Long
    nFilterIndex As Long
    sFile As String
    nFileSize As Long
    sFileTitle As String
    nTitleSize As Long
    sInitDir As String
    sDlgTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExt As Integer
    sDefFileExt As String
    nCustDataSize As Long
    fnHook As Long
    sTemplateName As String
 End Type

Private OFN As OPENFILENAME

'local variable(s) to hold property value(s)
Private mvarDefaultFileName  As String 'local copy
Private mvarDefaultExtension As String 'local copy
Private mvarFileTitle        As String 'local copy
Private mvarInitialFolder    As String 'local copy
Private mvarDialogTitle      As String 'local copy
Private Filters              As New Collection
Private mvarPathFile         As String 'local copy
Private mvarPath             As String 'local copy
Private mvarFileName         As String 'local copy
Private mvarExtension        As String 'local copy
Private mvarShortFileTitle   As String 'local copy
Private mvarShortFileName    As String 'local copy
Private mvarReadOnly         As Boolean 'local copy

Public Property Get ReadOnly() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ReadOnly
    ReadOnly = mvarReadOnly
End Property


Public Property Get ShortFileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ShortFileName
    ShortFileName = mvarShortFileName
End Property

Public Property Get ShortFileTitle() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ShortFileTitle
    ShortFileTitle = mvarShortFileTitle
End Property

Public Property Get Extension() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Extension
    Extension = mvarExtension
End Property

Public Property Get FileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileName
    FileName = mvarFileName
End Property

Public Property Get Path() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Path
    Path = mvarPath
End Property

Public Property Get PathFile() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PathFile
    PathFile = mvarPathFile
End Property

Public Sub ShowOpen()
  Dim r         As Long
  Dim sp        As Long
  Dim LongName  As String
  Dim shortName As String
  Dim ShortSize As Long

 'to keep lines short(er), I've abbreviated a
 'Null$ to n and n2, and the filter$ to f.
  Dim n As String
  Dim n2 As String
  Dim f As String
  n = Chr$(0)
  n2 = n & n

 '------------------------------------------------
 'INITIALIZATION
 '------------------------------------------------
 'fill in the size of the OFN structure
  OFN.nStructSize = Len(OFN)

 'assign the owner of the dialog; this can be null if no owner.
  OFN.hwndOwner = App.hInstance

 '------------------------------------------------
 'FILTERS
 '------------------------------------------------
  OFN.sFilter = CreateFilterString()
  OFN.nFilterIndex = FilterNumber()

 '------------------------------------------------
 'FILENAME
 '------------------------------------------------
 'pass a default filename and initialize for return value
  OFN.sFile = mvarDefaultFileName & Space$(1024) & n
  OFN.nFileSize = Len(OFN.sFile)

 'default extension applied to a selected file if it has no extension.
  OFN.sDefFileExt = mvarDefaultExtension

 'sFileTitle points to a buffer that receives the title of the
 'selected file. The application should use this string
 'to display the file title. If this member is NULL, the
 'function does not copy the file title.
  OFN.sFileTitle = Space$(512)
  OFN.nTitleSize = Len(OFN.sFileTitle)

 'sInitDir is the string that specifies the initial
 'file directory. If this member is NULL, the system
 'uses the current directory as the initial directory.
  OFN.sInitDir = mvarInitialFolder

 '------------------------------------------------
 'MISC
 '------------------------------------------------
 'sDlgTitle is the title to display in the dialog. If null
 'the default title for the dialog is used.
  OFN.sDlgTitle = mvarDialogTitle

 'flags are the actions and options for the dialog.
  OFN.flags = OFS_FILE_OPEN_FLAGS

 'Finally, show the File Open Dialog
  r = GetOpenFileName(OFN)

 '------------------------------------------------
 'RESULTS
 '------------------------------------------------
  If r Then
    'Path & File Returned (OFN.sFile):
     mvarPathFile = OFN.sFile

    'File Path (from OFN.nFileOffset):
     mvarPath = Left$(OFN.sFile, OFN.nFileOffset)

    'File Name (from OFN.nFileOffset):
     mvarFileName = Mid$(OFN.sFile, OFN.nFileOffset + 1, Len(OFN.sFile) - OFN.nFileOffset - 1)

    'Extension (from OFN.nFileExt):
     mvarExtension = Mid$(OFN.sFile, OFN.nFileExt + 1, Len(OFN.sFile) - OFN.nFileExt)

    'File Name (OFN.sFileTitle):
     mvarFileTitle = OFN.sFileTitle

    'Short 8.3 File Name (using (OFN.sFileTitle):
     LongName = OFN.sFileTitle
     shortName = Space$(128)
     ShortSize = Len(shortName)

     sp = GetShortPathName(LongName, shortName, ShortSize)
     mvarShortFileTitle = Left$(shortName, sp)

    'Short 8.3 File Name (using OFN.sFile):
     LongName = OFN.sFile
     shortName = Space$(128)
     ShortSize = Len(shortName)

     sp = GetShortPathName(LongName, shortName, ShortSize)
     mvarShortFileName = Left$(shortName, sp)

    'User Requested this file be opened as Read Only:
     mvarReadOnly = Abs((OFN.flags And OFN_READONLY) = OFN_READONLY)
  End If
End Sub

Public Sub ClearAllFilters()
    Set Filters = Nothing
End Sub

Public Sub RemoveFilter(ByVal Extension As String)
    On Error Resume Next
    Filters.Remove Extension
    Err.Clear
End Sub

Public Sub AddFilter(ByVal Description As String, ByVal Extension As String)
    On Error Resume Next
    Filters.Add Description & " (*." & Extension & ")", Extension
    Err.Clear
End Sub

Public Property Let DialogTitle(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DialogTitle = 5
    mvarDialogTitle = vData
End Property


Public Property Get DialogTitle() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DialogTitle
    DialogTitle = mvarDialogTitle
End Property



Public Property Let InitialFolder(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.InitialFolder = 5
    mvarInitialFolder = vData
End Property


Public Property Get InitialFolder() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.InitialFolder
    InitialFolder = mvarInitialFolder
End Property



Public Property Let FileTitle(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileTitle = 5
    mvarFileTitle = vData
End Property


Public Property Get FileTitle() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileTitle
    FileTitle = mvarFileTitle
End Property



Public Property Let DefaultExtension(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DefaultExtension = 5
    mvarDefaultExtension = vData
End Property


Public Property Get DefaultExtension() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DefaultExtension
    DefaultExtension = mvarDefaultExtension
End Property



Public Property Let DefaultFileName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DefaultFileName = 5
    mvarDefaultFileName = vData
End Property


Public Property Get DefaultFileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DefaultFileName
    DefaultFileName = mvarDefaultFileName
End Property


Private Sub Class_Initialize()
    mvarDefaultFileName = ""
    mvarDefaultExtension = ""
    mvarFileTitle = ""
    mvarInitialFolder = App.Path
    mvarDialogTitle = "Open File..."
End Sub

Private Sub Class_Terminate()
    Set Filters = Nothing
End Sub

Private Function CreateFilterString() As String
    Dim X      As Variant
    Dim Result As String
    Dim Pos    As Long
    Dim Filter As String
    
    If Not (Filters Is Nothing) Then
        Result = ""
        
        For Each X In Filters
            Pos = InStr(X, "*")
            Filter = Mid$(X, Pos, Len(X))
            Filter = Left$(Filter, Len(Filter) - 1)
            Result = Result & X & Chr$(0) & Filter & Chr$(0)
        Next
                
        Result = Result & Chr$(0)
        
        CreateFilterString = Result
    Else
        CreateFilterString = Chr$(0)
    End If
End Function

Private Function FilterNumber() As Long
    Dim X As Variant
    
    FilterNumber = 1
    For Each X In Filters
        FilterNumber = FilterNumber + 1
        If UCase$(mvarDefaultExtension) = UCase$(X) Then
            Exit For
        End If
    Next
End Function
