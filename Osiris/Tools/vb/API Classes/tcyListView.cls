VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'This is what the ColumnClick routine needs to be to impliment the
'graphic display for ascending/descrnding sorting
'
'Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
'    Dim x As New clsListView
'    Dim I As Long
'    Static sOrder
'
'    sOrder = Not sOrder
'
'    ListView1.SmallIcons = ImageList1 (Up as 1st image, Down as 2nd)
'    ListView1.SortKey = ColumnHeader.Index - 1
'    ListView1.SortOrder = Abs(sOrder)
'    ListView1.Sorted = True
'
'    x.ListView = ListView1
'    For I = 0 To 1
'      If I = ListView1.SortKey Then
'            x.HeaderIcon ListView1.SortKey, ListView1.SortOrder, vbJustifyRight, True
'      Else
'            x.HeaderIcon I, 0, vbNone, False
'      End If
'   Next
'
'   Set x = Nothing
'End Sub

Option Explicit

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
   
Private Const LVM_FIRST = &H1000
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55

Private Const LVIF_STATE = &H8
Private Const LVM_SETITEMSTATE = (LVM_FIRST + 43)
Private Const LVM_GETITEMSTATE As Long = (LVM_FIRST + 44)
Private Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)
Private Const LVIS_STATEIMAGEMASK As Long = &HF000

Private Const LVS_EX_FULLROWSELECT = &H20
Private Const LVS_EX_GRIDLINES = &H1
Private Const LVS_EX_CHECKBOXES As Long = &H4
Private Const LVS_EX_HEADERDRAGDROP = &H10
Private Const LVS_EX_TRACKSELECT = &H8


Private Const GWL_STYLE = (-16)
Private Const LVM_GETHEADER = (LVM_FIRST + 31)
Private Const HDS_HOTTRACK = &H4

Enum TPosition
    vbNone = 0
    vbJustifyLeft = &H2000
    vbJustifyRight = &H1000
End Enum

Private Type LVITEM
   mask         As Long
   iItem        As Long
   iSubItem     As Long
   State        As Long
   stateMask    As Long
   pszText      As String
   cchTextMax   As Long
   iImage       As Long
   lParam       As Long
   iIndent      As Long
End Type

Private Const HDI_IMAGE = &H20
Private Const HDI_FORMAT = &H4
Private Const HDF_LEFT = 0
Private Const HDF_STRING = &H4000
Private Const HDF_IMAGE = &H800
Private Const HDF_BITMAP_ON_RIGHT = &H1000
Private Const HDF_BITMAP = &H2000

Private Const HDM_FIRST = &H1200
Private Const HDM_SETITEM = (HDM_FIRST + 4)

Private Type HD_ITEM
   mask        As Long
   cxy         As Long
   pszText     As String
   hbm         As Long
   cchTextMax  As Long
   fmt         As Long
   lParam      As Long
   iImage      As Long
   iOrder      As Long
End Type

Private Const LVM_GETCOLUMN = (LVM_FIRST + 25)
Private Const LVM_GETCOLUMNORDERARRAY = (LVM_FIRST + 59)
Private Const LVCF_TEXT = &H4

Private Type LVCOLUMN
    mask As Long
    fmt As Long
    cx As Long
    pszText  As String
    cchTextMax As Long
    iSubItem As Long
    iImage As Long
    iOrder As Long
End Type

'font weight vars
Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
   
'SendMessage vars
Private Const WM_SETFONT = &H30
Private Const WM_GETFONT = &H31

Private Const LF_FACESIZE = 32

Private Type LOGFONT
   lfHeight As Long
   lfWidth As Long
   lfEscapement As Long
   lfOrientation As Long
   lfWeight As Long
   lfItalic As Byte
   lfUnderline As Byte
   lfStrikeOut As Byte
   lfCharSet As Byte
   lfOutPrecision As Byte
   lfClipPrecision As Byte
   lfQuality As Byte
   lfPitchAndFamily As Byte
   lfFaceName(LF_FACESIZE) As Byte
End Type

'local variable(s) to hold property value(s)
Private mvarFullRowSelect As Boolean 'local copy
Private mvarListView      As Object 'local copy
Private mvarGridLines     As Boolean 'local copy
Private mvarCheckBoxes    As Boolean 'local copy
Private mvarMoveColumns   As Boolean 'local copy
Private mvarHotTrack      As Boolean 'local copy
Private mvarMouseOverSelect As Boolean 'local copy

Public Property Let MouseOverSelect(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MouseOverSelect = 5
    Dim Result As Long
    Dim Style  As Long
    
    If Not (mvarListView Is Nothing) Then
        Style = SendMessageLong(mvarListView.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
        
        mvarMouseOverSelect = vData
        
        If mvarMouseOverSelect Then
            Style = Style Or LVS_EX_TRACKSELECT
        Else
            Style = Style Xor LVS_EX_TRACKSELECT
        End If
        
        Result = SendMessageLong(mvarListView.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, Style)
    End If
    
End Property


Public Property Get MouseOverSelect() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MouseOverSelect
    MouseOverSelect = mvarMouseOverSelect
End Property




Public Sub HeaderFont(Bold As Boolean, Italic As Boolean, Underline As Boolean, StrikeOut As Boolean)
   Dim LF        As LOGFONT
   Dim r         As Long
   Dim hCurrFont As Long
   Dim hOldFont  As Long
   Dim hHeader   As Long
   Dim hHeaderFont As Long
  If Not (mvarListView Is Nothing) Then
      'get the windows handle to the header
      'portion of the listview
       hHeader = SendMessageLong(mvarListView.hwnd, LVM_GETHEADER, 0, 0)
       
      'get the handle to the font used in the header
       hCurrFont = SendMessageLong(hHeader, WM_GETFONT, 0, 0)
       
      'get the LOGFONT details of the
      'font currently used in the header
       r = GetObject(hCurrFont, Len(LF), LF)
       
      'if GetObject was sucessful...
       If r > 0 Then
         
         'set the font attributes according to the selected check boxes
          If Bold Then
                LF.lfWeight = FW_BOLD
          Else
                LF.lfWeight = FW_NORMAL
          End If
    
          LF.lfItalic = Italic
          LF.lfUnderline = Underline
          LF.lfStrikeOut = StrikeOut
         
         'clean up by deleting any previous font
          If hHeaderFont > 0 Then
                r = DeleteObject(hHeaderFont)
          End If
          
         'create a new font for the header control to use.
         'This font must NOT be deleted until it is no
         'longer required by the control, typically when
         'the application ends (see the Unload sub), or
         'above as a new font is to be created.
          hHeaderFont = CreateFontIndirect(LF)
          
         'select the new font as the header font
          hOldFont = SelectObject(hHeader, hHeaderFont)
          
         'and inform the listview header of the change
          r = SendMessageLong(hHeader, WM_SETFONT, hHeaderFont, True)
       End If
    End If
End Sub
Public Sub HeaderIcon(SortKey As Long, SortOrder As Long, Position As TPosition, ShowImage As Boolean)
   Dim r       As Long
   Dim hHeader As Long
   Dim HD      As HD_ITEM
   Dim ShwImg  As Long
   
  If Not (mvarListView Is Nothing) Then
        'get a handle to the listview header component
         hHeader = SendMessageLong(mvarListView.hwnd, LVM_GETHEADER, 0, 0)
         
         If ShowImage Then
              ShwImg = HDF_IMAGE
         Else
              ShwImg = 0
         End If
         
        'set up the required structure members
         With HD
            .mask = HDI_IMAGE Or HDI_FORMAT
            .fmt = HDF_LEFT Or HDF_STRING Or Position Or ShwImg
            .pszText = mvarListView.ColumnHeaders(SortKey + 1).Text
            
            If ShowImage Then .iImage = SortOrder
         End With
         
        'modify the header
         r = SendMessageAny(hHeader, HDM_SETITEM, SortKey, HD)
    End If
End Sub

Public Property Let HotTrack(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HotTrack = 5
   mvarHotTrack = vData

   Dim r As Long
   Dim hHeader As Long
   Dim rstyle As Long
   
  If Not (mvarListView Is Nothing) Then
      'get a handle to the listview header component
       hHeader = SendMessageLong(mvarListView.hwnd, LVM_GETHEADER, 0, 0)
       
      'set/toggle the hottrack style attribute
       If mvarHotTrack Then
            rstyle = 1342177474 'On
       Else
            rstyle = 1342177478 'Off
       End If
       
       rstyle = rstyle Xor HDS_HOTTRACK
    
      'set the header style
      r = SetWindowLong(hHeader, GWL_STYLE, rstyle)
  End If
End Property


Public Property Get HotTrack() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HotTrack
    HotTrack = mvarHotTrack
End Property




Public Function ColumnOrder() As Collection
'working variables
   Dim I            As Long
   Dim r            As Long
   Dim firstCol     As Long
   Dim lastCol      As Long
   Dim totalCols    As Long
   Dim LVC          As LVCOLUMN
   Dim col          As New Collection
   
   If Not (mvarListView Is Nothing) Then
      'initialize the variables needed. totalCols is the 1-based
      'total required for the API. lastCol is the 0-based
      'number of columns in the listview.
       totalCols = mvarListView.ColumnHeaders.Count
       firstCol = 0
       lastCol = totalCols - 1
          
      'to get the column order, we have to pass an array to the API.
      'On return, it will be filled with the index of the column in
      'incrementing positions.  For example, if column 2 was moved
      'to position 0, the return array would hold 2, 1, 0, 3.
      'And because this will be used directly in the API calls, it's
      'declared As Long.
       ReDim posArray(firstCol To lastCol) As Long
    
       r = SendMessageAny(mvarListView.hwnd, LVM_GETCOLUMNORDERARRAY, totalCols, posArray(firstCol))
         
      'with the array filled, its now a matter of looping through the
      'array, and passing each item as the position (wParam).  The
      'LVCOLUMN type (LVC) will be filled with the data for the
      'passed index (LVCF_TEXT in this example).
       For I = firstCol To lastCol
            col.Add posArray(I)
       Next
       
       Set ColumnOrder = col
   End If
End Function


Public Property Let MoveColumns(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MoveColumns = 5
    Dim Result As Long
    Dim Style  As Long
    
    If Not (mvarListView Is Nothing) Then
        Style = SendMessageLong(mvarListView.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
        
        mvarMoveColumns = vData
        
        If mvarMoveColumns Then
            Style = Style Or LVS_EX_HEADERDRAGDROP
        Else
            Style = Style Xor LVS_EX_HEADERDRAGDROP
        End If
        
        Result = SendMessageLong(mvarListView.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, Style)
    End If
End Property


Public Property Get MoveColumns() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MoveColumns
    MoveColumns = mvarMoveColumns
End Property

Public Function IsChecked(Index As Long) As Boolean
    Dim Result As Long
    
    IsChecked = False
    If Not (mvarListView Is Nothing) Then
        Result = SendMessageAny(mvarListView.hwnd, LVM_GETITEMSTATE, Index, LVIS_STATEIMAGEMASK)
        IsChecked = (Result And &H2000&)
    End If
End Function

Public Sub Checked(Index As Long, State As Boolean)
    Dim LV As LVITEM
    
    If Not (mvarListView Is Nothing) Then
        With LV
          .mask = LVIF_STATE
          .State = IIf(State, &H2000, &H1000)
          .stateMask = LVIS_STATEIMAGEMASK
        End With
        
        SendMessageAny mvarListView.hwnd, LVM_SETITEMSTATE, Index, LV
    End If
End Sub

Public Function SelectedItems() As Collection
   Dim I   As Long
   Dim r   As Long
   Dim LV  As LVITEM
   Dim col As New Collection
   
  'iterate through each item, checking its item state
   For I = 0 To mvarListView.ListItems.Count - 1
      r = SendMessageAny(mvarListView.hwnd, LVM_GETITEMSTATE, I, LVIS_STATEIMAGEMASK)
   
     'when an item is checked, the LVM_GETITEMSTATE call
     'returns 8192 (&H2000&).
      If r And &H2000& Then
        col.Add I
      End If
   Next
   
    Set SelectedItems = col
    Set col = Nothing
End Function

Public Property Let CheckBoxes(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CheckBoxes = 5
    Dim Result As Long
    Dim Style  As Long
    
    If Not (mvarListView Is Nothing) Then
        Style = SendMessageLong(mvarListView.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
        
        mvarCheckBoxes = vData
        
        If mvarCheckBoxes Then
            Style = Style Or LVS_EX_CHECKBOXES
        Else
            Style = Style Xor LVS_EX_CHECKBOXES
        End If
        
        Result = SendMessageLong(mvarListView.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, Style)
    End If
End Property


Public Property Get CheckBoxes() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CheckBoxes
    CheckBoxes = mvarCheckBoxes
End Property


Public Property Let GridLines(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GridLines = 5
    Dim Result As Long
    Dim Style  As Long
    
    If Not (mvarListView Is Nothing) Then
        Style = SendMessageLong(mvarListView.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
        
        mvarGridLines = vData
        
        If mvarGridLines Then
            Style = Style Or LVS_EX_GRIDLINES
        Else
            Style = Style Xor LVS_EX_GRIDLINES
        End If
        
        Result = SendMessageLong(mvarListView.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, Style)
    End If
End Property


Public Property Get GridLines() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GridLines
    GridLines = mvarGridLines
End Property




Public Property Let ListView(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ListView = Form1
    Set mvarListView = vData
End Property


Public Property Get ListView() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ListView
    Set ListView = mvarListView
End Property



Public Property Let FullRowSelect(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FullRowSelect = 5
    Dim Result As Long
    Dim Style  As Long
    
    If Not (mvarListView Is Nothing) Then
        Style = SendMessageLong(mvarListView.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
        
        mvarFullRowSelect = vData
        
        If mvarFullRowSelect Then
            Style = Style Or LVS_EX_FULLROWSELECT
        Else
            Style = Style Xor LVS_EX_FULLROWSELECT
        End If
        
        Result = SendMessageLong(mvarListView.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, Style)
    End If
End Property


Public Property Get FullRowSelect() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FullRowSelect
    FullRowSelect = mvarFullRowSelect
End Property



