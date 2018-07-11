Attribute VB_Name = "mMain_Explorer"
#Const ProgressBar = True

Option Explicit

Private Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

Private Type ExtensionType
    Extension As String
    Category As String
End Type

'-----------------------------------CONSTANTS----------------------------------
Global Const DEBUGGING = True
Global Const SHOW_TIMING_INFO = False
Global Const SMALL_ICON_SIZE = 16
Global Const LARGE_ICON_SIZE = 32
Global Const MENU_ICON_PATH = "c:\osiris\resources\icons\Menu\"
Global Const NONE_LABEL = "<None>"
Global Const MAX_LABEL_LENGTH = 64
Global Const NODE_CLICK_TIMER_DELAY = 250 'in milliseconds, the time between clicking
                                          'a node and loading the listview
Global Const ITEM_COLUMN = 0
Global Const DATA_ID_COLUMN = 1
Global Const PARENT_NODE_COLUMN = 2
Global Const TYPE_COLUMN = 3

Global Const Link_Icon = "link"
Global Const Read_Icon = "read"
Global Const System_Icon = "system"
Global Const SBICON_Logo = "hdmasmall.bmp"
Global Const PBICON_BuildMenu = "Interface.bmp"
Global Const PBICON_LoadIcon = "ICONMAKE.bmp"
Global Const PBICON_LoadTreeView = "Hostmanager.bmp"

Global DB_NodeTable As String
Global DB_DefaultDataTable As String
Global DB_QuickAddItemsTable As String
Global DB_LargeIconsTable As String
Global DB_SmallIconsTable As String
Global DB_GlobalsTable As String
Global DB_ExtensionsTable As String
Global DB_DataTypesTable As String
Global DB_InboxTable As String

'-----------------------------------VARIABLES----------------------------------

Global InFormLoad As Boolean
Global NodeBuffer As New tvBufferNode
Global TableComboNeedsRefresh As Boolean
Global LargeIconFolder As String
Global SmallIconFolder As String
Global ClickedVariationMenu As Boolean
Global StartFindNodeIndex As Integer
Global StartFindItemIndex As Integer
Global lvNumListItemsSelected As Long
Global HTMLEditRunning As Boolean
Global lvLoaded As Boolean
Global CurTempFolder As String
Global LastFind As String
Global LastReplace As String
Global LastCase As Integer
Global LastWhole As Integer
Global lvNeedsRefresh As Boolean
Global ListMouseClick As Boolean
Global WasMinimized As Boolean
Global number_of_custom_items As Long
Global dbase As Database
Global CurrentUser As String
Global TempFileCounter As Long
Global Instance_Scan As Integer
Global nodx As Node    'the current node
Global nodq As Node
Global BufferNode As Node
Global CurrentItem As ListItem
Global tvAttribCol As New Collection
Global lvAttribCol As New Collection
Global NumberExtensions As Long
Global RefreshInProgress As Boolean
Global record_count As Long
Global selected_table As String
Global Qdf As QueryDef
Global QuickAddNodesNodeID As Long
Global FocusFrom As String
Global PropertiesActive As Boolean
Global isInsertKey As Boolean
Global SkipCnt As Integer
Global TrapMultiSelectDrag As Boolean
Global NumListItemsCopied As Long
Global LastSplitterLeft As Integer
Global LastTab As Integer
Global CurrentDatabaseFile As String
Global lvPaintCnt As Integer
Global lvCountPaints As Boolean
Global lvCancel As Boolean
Global lvEraseBkGnd1 As Boolean
Global lvEraseBkGnd2 As Boolean
Global lvNoEraseCnt As Integer
Global lvNextNoErase As Boolean
Global lvSaveEraseRect As Boolean
Global RectHeight As Integer
Global RectWidth As Integer
Global FindEnabled As Boolean
Dim MyFont As Long
Dim OldFont As Long
Dim pnt As New PaintEffects
Public fMainForm As frmMain
Public fPropForm As frmProperties
Public fProgForm As frmProgressBar
Public fReplaceForm As frmReplace


Sub Main()
    
    If Not DEBUGGING Then
        Dim fLogin As New frmLogin  'new login form
        fLogin.Show vbModal
        If Not fLogin.OK Then       'check if they hit cancel
            End                     'Login Failed so exit app
        End If
        Unload fLogin               'unload here after check fLogin.OK
    End If
    
    If GetSetting(App.EXEName, "Options", _
            "Show Splash", 1) Then
        frmSplash.Show vbModal
    End If
    
    Set fProgForm = New frmProgressBar
    Set fPropForm = New frmProperties
    Set fReplaceForm = New frmReplace
    Set fMainForm = New frmMain
    Load fProgForm
    Load fPropForm
    Load fReplaceForm
    Load fMainForm
End Sub


' this adds a 'K' to the front of a numeric key and returns the result as a string
' this is needed since a key must start with a letter due to microsoft's fetishes
'REQUIRES: nothing
Public Function AddK(num As Long, Optional AddL As Boolean = False) As String
    If AddL Then
        AddK = "L" & Format$(num)
    Else
        AddK = "K" & Format$(num)
    End If
End Function

' this removes a 'K' from the front of a string key and returns the result as a long integer
' this is needed since we wanted to store the key as long in the database file
'REQUIRES: nothing
Public Function RemoveK(text As String) As Long
    RemoveK = Val(Right$(text, Len(text) - 1))
End Function

Public Function AddKtoStr(text As String, Optional AddL As Boolean = False) As String
    If AddL Then
        AddKtoStr = "L" & text
    Else
        AddKtoStr = "K" & text
    End If
End Function

Public Sub InitDBase(NodeTable As String, DefaultDataTable As String, _
            QuickAddItemsTable As String, LargeIconsTable As String, _
            SmallIconsTable As String, GlobalsTable As String, _
            ExtensionsTable As String, DataTypesTable As String, _
            InboxTable As String)
    
    'check here to see if these tables actually exist
    
    DB_NodeTable = NodeTable
    DB_DefaultDataTable = DefaultDataTable
    DB_QuickAddItemsTable = QuickAddItemsTable
    DB_LargeIconsTable = LargeIconsTable
    DB_SmallIconsTable = SmallIconsTable
    DB_GlobalsTable = GlobalsTable
    DB_ExtensionsTable = ExtensionsTable
    DB_DataTypesTable = DataTypesTable
    DB_InboxTable = InboxTable
End Sub


Public Sub BuildimlMenu(Optional UseProgressBar As Boolean = False, _
        Optional AbortButton As Boolean = False)
    
    Dim LargeIconFolder As String
    Dim PBICON_BuildMenu As String
    Dim MENU_ICON_PATH As String
    
    LargeIconFolder = "c:\osiris\resources\icons\32"
    MENU_ICON_PATH = "c:\osiris\resources\icons\menu"
    
    MsgBox "TEMP:  LargeIconFolder,PBICON_BuildMenu,MENU_ICON_PATH should be stored in the registry!", vbInformation
    
    If UseProgressBar Then
        InitProgressBar fProgForm, "Building Image List for menus . . .", 0, 100, _
            LargeIconFolder & PBICON_BuildMenu, False, AbortButton
    End If
    
    On Error Resume Next    'if the pictures do not exist, then blank pictures
    fMainForm.imlMenu.ListImages.Add , "Copy", LoadPicture(MENU_ICON_PATH & "Edit Copy 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "Cut", LoadPicture(MENU_ICON_PATH & "Edit Cut 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "Paste", LoadPicture(MENU_ICON_PATH & "Edit Paste 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "New", LoadPicture(MENU_ICON_PATH & "File New 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "Open", LoadPicture(MENU_ICON_PATH & "File Open 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "Print", LoadPicture(MENU_ICON_PATH & "File Print 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "Preview", LoadPicture(MENU_ICON_PATH & "File Print Preview 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "Save", LoadPicture(MENU_ICON_PATH & "File Save 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "Paint", LoadPicture(MENU_ICON_PATH & "Format Color 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "Help", LoadPicture(MENU_ICON_PATH & "Help 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "Delete", LoadPicture(MENU_ICON_PATH & "Edit Delete 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "Prop", LoadPicture(MENU_ICON_PATH & "Properties 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "Security", LoadPicture(MENU_ICON_PATH & "Security 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "AccessTable", LoadPicture(MENU_ICON_PATH & "Access Table 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "Bold", LoadPicture(MENU_ICON_PATH & "Bold 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "Italic", LoadPicture(MENU_ICON_PATH & "Italic 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "Underline", LoadPicture(MENU_ICON_PATH & "Underline 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "Right Just", LoadPicture(MENU_ICON_PATH & "Right Just 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "Left Just", LoadPicture(MENU_ICON_PATH & "Left Just 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "Center Just", LoadPicture(MENU_ICON_PATH & "Center Just 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "Bullets", LoadPicture(MENU_ICON_PATH & "Bullets 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "Font Color", LoadPicture(MENU_ICON_PATH & "Font Color 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "Absolute Mode", LoadPicture(MENU_ICON_PATH & "Absolute Mode 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "Absolute Pos", LoadPicture(MENU_ICON_PATH & "Absolute Pos.bmp")
    fMainForm.imlMenu.ListImages.Add , "BGColor", LoadPicture(MENU_ICON_PATH & "BGColor.bmp")
    fMainForm.imlMenu.ListImages.Add , "FGColor", LoadPicture(MENU_ICON_PATH & "FGColor.bmp")
    fMainForm.imlMenu.ListImages.Add , "Borders", LoadPicture(MENU_ICON_PATH & "Borders.bmp")
    fMainForm.imlMenu.ListImages.Add , "Delete Cell", LoadPicture(MENU_ICON_PATH & "Delete Cell.bmp")
    fMainForm.imlMenu.ListImages.Add , "Delete Column", LoadPicture(MENU_ICON_PATH & "Delete Column.bmp")
    fMainForm.imlMenu.ListImages.Add , "Delete Row", LoadPicture(MENU_ICON_PATH & "Delete Row.bmp")
    fMainForm.imlMenu.ListImages.Add , "Decrease Indent", LoadPicture(MENU_ICON_PATH & "Decrease Indent.bmp")
    fMainForm.imlMenu.ListImages.Add , "Details", LoadPicture(MENU_ICON_PATH & "Details.bmp")
    fMainForm.imlMenu.ListImages.Add , "Find", LoadPicture(MENU_ICON_PATH & "Find.bmp")
    fMainForm.imlMenu.ListImages.Add , "Image", LoadPicture(MENU_ICON_PATH & "Image.bmp")
    fMainForm.imlMenu.ListImages.Add , "Increase Indent", LoadPicture(MENU_ICON_PATH & "Increase Indent.bmp")
    fMainForm.imlMenu.ListImages.Add , "Insert Cell", LoadPicture(MENU_ICON_PATH & "Insert Cell.bmp")
    fMainForm.imlMenu.ListImages.Add , "Insert Column", LoadPicture(MENU_ICON_PATH & "Insert Column.bmp")
    fMainForm.imlMenu.ListImages.Add , "Insert Row", LoadPicture(MENU_ICON_PATH & "Insert Row.bmp")
    fMainForm.imlMenu.ListImages.Add , "Insert Table", LoadPicture(MENU_ICON_PATH & "Insert Table.bmp")
    fMainForm.imlMenu.ListImages.Add , "Link", LoadPicture(MENU_ICON_PATH & "Link.bmp")
    fMainForm.imlMenu.ListImages.Add , "Merge Cell", LoadPicture(MENU_ICON_PATH & "Merge Cell.bmp")
    fMainForm.imlMenu.ListImages.Add , "Redo", LoadPicture(MENU_ICON_PATH & "Redo.bmp")
    fMainForm.imlMenu.ListImages.Add , "Snap Grid 16", LoadPicture(MENU_ICON_PATH & "Snap Grid 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "Split Cell 16", LoadPicture(MENU_ICON_PATH & "Split Cell 16.bmp")
    fMainForm.imlMenu.ListImages.Add , "Undo", LoadPicture(MENU_ICON_PATH & "Undo.bmp")
    fMainForm.imlMenu.ListImages.Add , "Numbers", LoadPicture(MENU_ICON_PATH & "Numbers.bmp")
    On Error GoTo 0
    
    If UseProgressBar Then
        fProgForm.Hide
    End If
End Sub


Public Sub CopyDataTableItems(db As Database, source_table As String, _
            target_table As String, Optional parent_nodeid As Long = 0)
    Dim srcrecord As Recordset
    Dim dstRecord As Recordset
    Dim FreeNumber As Long
    
    If parent_nodeid = 0 Then
        Set srcrecord = db.OpenRecordset(source_table, dbOpenTable)
    Else
        Set srcrecord = db.OpenRecordset("SELECT * FROM " & source_table & _
                " WHERE Parent_Node = " & parent_nodeid, dbOpenDynaset)
    End If
    Set dstRecord = db.OpenRecordset(target_table, dbOpenTable)
    While Not srcrecord.EOF
        FreeNumber = FindFreeID(db, tvAttribCol(nodx.key).table_name, _
            "Data_ID")
        dstRecord.AddNew
        dstRecord!data_id = FreeNumber
        dstRecord!data_label = srcrecord!data_label
        dstRecord!parent_node = srcrecord!parent_node
        dstRecord!icon_large = srcrecord!icon_large
        dstRecord!icon_small = srcrecord!icon_small
        dstRecord!data_type = srcrecord!data_type
        dstRecord!read_only = srcrecord!read_only
        dstRecord!created = srcrecord!created
        dstRecord!created_by = srcrecord!created_by
        dstRecord!last_modified = Now
        dstRecord!modified_by = CurrentUser
        dstRecord!variation = srcrecord!variation
        dstRecord.Update
        CopyMemo srcrecord, "data_value", dstRecord, _
            "data_value", CurTempFolder & "CopyMemo.tmp"
        CopyBLOB srcrecord, "binary_data_value", dstRecord, "binary_data_value"
        srcrecord.Delete
        srcrecord.MoveNext
    Wend
    srcrecord.Close
    dstRecord.Close
End Sub

Public Function DeleteTable(db As Database, table_name As String) As Boolean
    Dim response As Integer
    Dim record As Recordset
    
    response = MsgBox("Are you sure you want to delete the table " & _
            table_name & " from the database ?", vbYesNo + vbQuestion)
    If response = vbYes Then
        Set record = db.OpenRecordset("SELECT COUNT (*) AS [Count] FROM " _
                & DB_NodeTable & " WHERE Table_Name = '" & table_name _
                & "'", dbOpenDynaset)
        record_count = record!count
        record.Close
        
        Set record = db.OpenRecordset("SELECT * FROM " & DB_NodeTable _
                & " WHERE Table_Name = '" & table_name & "'", dbOpenDynaset)
        If record_count > 0 Then
            frmReassign.Show vbModal
            If selected_table <> "-1" Then
                record.MoveFirst
                While Not record.EOF
                    record.Edit
                    record!table_name = selected_table
                    tvAttribCol(AddK(record!node_id)).table_name = selected_table
                    record.Update
                    record.MoveNext
                Wend
                CopyDataTableItems db, table_name, selected_table
            Else
                GoTo Done
            End If
        End If
        db.TableDefs.Delete table_name
        DeleteTable = True
    Else
        DeleteTable = False
    End If
Done:
    record.Close
End Function


' This function will swap the the two node_id's passed to it, in the current database.
' It will also swap references to these nodes in their respective data tables, which
' is stored in the tvattribcol(key).table_name format.  It will NOT swap the visual
' appearance of the nodes in the treeview.  That needs to be handled by the calling
' procedure.
Public Sub DBMoveNodeDown(db As Database, curNode As Node)
    Dim record As Recordset
    Dim Order1 As Long
    Dim Order2 As Long
    Dim Order3 As Long
    Dim NewStartPos As Long
    Dim NextofNextNode As Node
    Dim NextNode As Node
    
    'get the next sibling after the current node
    Set NextNode = curNode.Next
    
    'this check is not really necessary, since the Move Down menu is disabled
    'if the next node is nothing.
    If NextNode Is Nothing Then
        GoTo Done
    End If
    
    'get the order of the current node
    Set record = db.OpenRecordset("SELECT [Order] FROM " & DB_NodeTable _
        & " WHERE Node_ID = " & RemoveK(curNode.key), dbOpenDynaset)
    If record.EOF Then
        record.Close
        MsgBox "DBMoveNodeDown:  Current node not found in database!", vbCritical
        GoTo Done
    End If
    Order1 = record!order   'current node's order #
    record.Close
    
    'get the order of the next node
    Set record = db.OpenRecordset("SELECT [Order] FROM " & DB_NodeTable _
        & " WHERE Node_ID = " & RemoveK(NextNode.key), dbOpenDynaset)
    If record.EOF Then
        record.Close
        MsgBox "DBMoveNodeDown:  Next node not found in database!", vbCritical
        GoTo Done
    End If
    Order2 = record!order   'next node's order #
    record.Close
    
    Set NextofNextNode = NextNode.Next  'get the node after the next node
    
    'if there are no more siblings, get the next one of their parent
    'which would be the next one in line according to [Order]
    If NextofNextNode Is Nothing Then
        Set NextofNextNode = GeneralNextofParent(NextNode)
        If NextofNextNode Is Nothing Then
            'if there is no next node for the parent either, then we are at the
            'very last node, so select the maximum [Order] from the database.
            Set record = db.OpenRecordset("SELECT Max([Order]) AS [MaxOrder] FROM " _
                                & DB_NodeTable, dbOpenDynaset)
            If record.EOF Then
                record.Close
                MsgBox "DBMoveNodeDown:  There was no MAXIMUM Order in the Nodes table!", vbCritical
                GoTo Done
            End If
            Order3 = record!MaxOrder + 1    'the order # of the one
                                            'after the next node
            GoTo UpdateOrder
        End If
    End If
    
    'get the order of the next of next node
    Set record = db.OpenRecordset("SELECT [Order] FROM " & DB_NodeTable _
        & " WHERE Node_ID = " & RemoveK(NextofNextNode.key), dbOpenDynaset)
    If record.EOF Then
        record.Close
        MsgBox "DBMoveNodeDown:  Next of next node not found in database!", vbCritical
        GoTo Done
    End If
    Order3 = record!order   'the order # of the one after the next node
    record.Close
    
UpdateOrder:
    'mark current node and its children for later update
    'by making them the negatives of their current values
    
    On Error GoTo Err_Execute
    db.Execute "UPDATE Nodes SET [Order] = (-1)*[Order] " _
                    & "WHERE [Order] Between " _
                    & Order1 & " AND " & Order2 - 1, dbFailOnError
    
    'update next node and its children
    db.Execute "UPDATE Nodes SET [Order] = " & Order1 & " + ([Order]-" _
                    & Order2 & ") WHERE [Order] Between " _
                    & Order2 & " AND " & Order3 - 1, dbFailOnError
    
    
    NewStartPos = Order1 + (Order3 - Order2)
    
    'update currentnode and its children
    db.Execute "UPDATE Nodes SET [Order] = " & NewStartPos _
                & "+((-2*[Order]+[Order])-" & Order1 & ")" _
                & " WHERE [Order] < 0", dbFailOnError
    On Error GoTo 0
Done:
Exit Sub

Err_Execute:
    DisplayDBEngineErrors

End Sub

Public Sub InitOwnerDrawMenus(Optional UseProgressBar As Boolean = False, _
        Optional AbortButton As Boolean = False)
    Dim i As Long
    
    If DEBUGGING Then
        Exit Sub
    End If
    
    If UseProgressBar Then
        InitProgressBar fProgForm, "Initializing Owner Draw Menus . . .", 0, 100, _
            LargeIconFolder & "Interface.bmp", False, AbortButton
    End If
    Caps(2) = "Open (Ctrl-O)"
    Caps(3) = ""
    Caps(4) = "Properties"
    Caps(5) = ""
    Caps(6) = "Print"
    Caps(7) = ""
    Caps(8) = "Exit"
    Caps(9) = ""
    Caps(10) = "Cut (Ctrl-X)"
    Caps(11) = "Copy (Ctrl-C)"
    Caps(12) = "Paste (Ctrl-V)"
    Caps(13) = "Delete (Del)"
    Caps(14) = ""
    Caps(15) = "Add (Ins)"
    Caps(16) = "Rename (Ctrl-R)"
    Caps(17) = ""
    Caps(18) = "Select All (Ctrl-A)"
    Caps(19) = "Invert Selection"
    Caps(20) = ""
    Caps(21) = "Find (Ctrl-F)"
    Caps(22) = "Replace (Ctrl-H)"
    HasIcon(2) = 5
    HasIcon(3) = 0
    HasIcon(4) = 12
    HasIcon(5) = 0
    HasIcon(6) = 5
    HasIcon(7) = 0
    HasIcon(8) = 2
    HasIcon(9) = 1
    HasIcon(10) = 3
    HasIcon(11) = 11
    HasIcon(12) = 0
    HasIcon(13) = 0
    HasIcon(14) = 0
    HasIcon(15) = 0
    HasIcon(16) = 0
    HasIcon(17) = 0
    HasIcon(18) = 0
    HasIcon(19) = 0
    HasIcon(20) = 0
    HasIcon(21) = 0
    HasIcon(22) = 0
    mMenuItems.MenuForm = fMainForm
    mMenuItems.subMenu = 0
    For i = 0 To 4
        mMenuItems.menuId = i
        OwnerDrawMenu (i + 2)
    Next i
    mMenuItems.subMenu = 1
    For i = 0 To 12
        mMenuItems.menuId = i
        OwnerDrawMenu (i + 2)
    Next i
    If UseProgressBar Then
        fProgForm.Hide
    End If
End Sub


Public Sub InitSubclassing(Optional UseProgressBar As Boolean = False, _
    Optional AbortButton As Boolean = False)
    
    If Not DEBUGGING Then
        If UseProgressBar Then
            InitProgressBar fProgForm, "Initializing Subclassing . . .", 0, 100, _
                LargeIconFolder & "Programs.bmp", False, AbortButton
        End If
        For Instance_Scan = MIN_INSTANCES To MAX_INSTANCES
            If Instances(Instance_Scan).in_use = False Then
                m_MyInstance = Instance_Scan
                Instances(Instance_Scan).in_use = True
                Instances(Instance_Scan).ClassAddr = ObjPtr(fMainForm)
                Exit For
            End If
        Next Instance_Scan
        Call Hook_Window(fMainForm.hwnd, m_MyInstance)
    
        Instances(m_MyInstance + 1).in_use = True
        Instances(m_MyInstance + 1).ClassAddr = ObjPtr(fMainForm.tvTreeView)
        Call Hook_Window(fMainForm.tvTreeView.hwnd, m_MyInstance + 1)
        Instances(m_MyInstance + 2).in_use = True
        Instances(m_MyInstance + 2).ClassAddr = ObjPtr(fMainForm.lvListView)
        Call Hook_Window(fMainForm.lvListView.hwnd, m_MyInstance + 2)
        Instances(m_MyInstance + 3).in_use = True
        Instances(m_MyInstance + 3).ClassAddr = ObjPtr(fMainForm.sbStatusBar)
        Call Hook_Window(fMainForm.sbStatusBar.hwnd, m_MyInstance + 3)
        Instances(m_MyInstance + 4).in_use = True
        Instances(m_MyInstance + 4).ClassAddr = ObjPtr(fPropForm.tvQuick)
        Call Hook_Window(fPropForm.tvQuick.hwnd, m_MyInstance + 4)
        If UseProgressBar Then
            fProgForm.Hide
        End If
    End If
End Sub


Public Function SwitchBoard(ByVal hwnd As Long, ByVal Msg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long
        
    Dim PrevWndProc As Long
    Dim instance_check As Integer
    Static m_UpdateRect As RECT
    Static oldHeight As Long
    Static oldWidth As Long
    Static Resizing As Boolean
    Dim tempRect As RECT
    Dim tdc&
        Dim usedc&
        Dim oldbm&
        Dim bm As BITMAP
        Dim rc As RECT
        Dim offsx&, offsy&
    Dim hBrush%, RetVal%
    Dim tmpRectHeight As Integer
    Dim tmpRectWidth As Integer
    Static erase_wParam As Long
    Static sberase_wParam As Long
    Dim bkcolor As Integer
    Dim RectTooBig As Boolean
    Dim tMinMax As MINMAXINFO
    'Various structs we'll need
    Dim MeasureInfo As MEASUREITEMSTRUCT
    Dim DrawInfo As DRAWITEMSTRUCT
    Dim mii As MENUITEMINFO
    'Set later for separator flag:
    Dim IsSep As Boolean
    'Our custom brush and the old one
    Dim hBr As Long, hOldBr As Long
    'Our custom pen and the old one
    Dim hPen As Long, hOldPen As Long
    'The text color of the menu items
    Dim lTextColor As Long
    'Now much to bump the menu's selection
    'rectangle over
    Dim iRectOffset As Integer
    Dim ptApi As POINTAPI
    Dim xReturn As Long
    Dim BandId As Integer
    Dim tmpitemsSelected As Long
    Static oldsbPanelwidth As Long
    Dim currentwidth As Long
    Dim currentheight As Long
    Static oldlvwidth As Long
    Static oldlvheight As Long
    Dim edit_text As String
    Dim tempstr As String
    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Dim TestIndex As Long
    

    'Do this early as we may unhook
    PrevWndProc = Is_Hooked(hwnd)

    
    Select Case hwnd
    Case Instances(5).hwnd  ' properties form
        Select Case Msg
        Case WM_KEYDOWN    'this makes it so the user can only select a root node
            If wParam = &H28 Then
                'SendKeys "{LEFT}"
                If Not nodq.Next Is Nothing Then
                    nodq.Expanded = False
                    nodq.Next.Selected = True
                    Set nodq = nodq.Next
                    nodq.Expanded = True
                End If
'                Dim tempnode As Node
'               Dim tempx As Long
'                Dim tempy As Long
             
'                tempstr = Hex(lParam)

'                X = Val("&H" & right$(tempstr, 4))
'                Y = Val("&H" & mid$(tempstr, 1, Len(tempstr) - 4))
                'Debug.Print "X in pixels: " & X
                'Debug.Print "Y in pixels: " & Y
                'Debug.Print "X in twips: " & X
                'Debug.Print "Y in twips: " & Y
                'Debug.Print "tvquick left twips: " & (fPropForm.tvQuick.Left + fPropForm.fraObjects.Left) / Screen.TwipsPerPixelX
                'Debug.Print "tvquick top twips: " & (fPropForm.tvQuick.Top + fPropForm.fraObjects.Top) / Screen.TwipsPerPixelY
                
'                tempx = (Screen.TwipsPerPixelX * X) - (fPropForm.tvQuick.Left + fPropForm.fraObjects.Left + fPropForm.SSTab1.Left)
'                tempy = (Screen.TwipsPerPixelY * Y) - (fPropForm.tvQuick.Top + fPropForm.fraObjects.Top + fPropForm.SSTab1.Top)
                'Debug.Print "tempx:  " & tempx
                'Debug.Print "tempy:  " & tempy
'                Set tempnode = fPropForm.tvQuick.HitTest(tempx, tempy)
            
'                If Not tempnode Is Nothing Then
'                    If tempnode.Parent Is Nothing Then
'                        Debug.Print "root node:  " & tempnode.text
'                        SwitchBoard = CallWindowProc(PrevWndProc, hWnd, Msg, wParam, lParam)
'                    Else
'                        Debug.Print "child node:  " & tempnode.text
 '                   End If
'                Else
'                    Debug.Print "white space"
'                    SwitchBoard = CallWindowProc(PrevWndProc, hWnd, Msg, wParam, lParam)
'                End If
            Else
                SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
            End If
        Case Else
            SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
        End Select
    Case Instances(4).hwnd  'Status Bar
        Select Case Msg
        Case WM_PAINT   'prevent paint of left pane on horizontal resize
            If Resizing Then
                If fMainForm.ScaleHeight <> oldHeight Then
                    oldHeight = fMainForm.ScaleHeight
                    SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
                Else
                    Call GetUpdateRect(hwnd, tempRect, False)
                    currentwidth = _
                        fMainForm.sbStatusBar.Panels(1).Width
                    If currentwidth > oldsbPanelwidth Then
                        tempRect.Right = oldsbPanelwidth _
                            \ Screen.TwipsPerPixelX
                    Else
                        tempRect.Right = currentwidth _
                            \ Screen.TwipsPerPixelX
                    End If
                    oldsbPanelwidth = currentwidth
                    
'**************************************************************
'the next line causes an overflow during a horizontal or diagonal resize
'under Windows NT
'**************************************************************
                    ValidateRect hwnd, tempRect
                    SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
                End If
            Else
                SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
            End If
        Case WM_ERASEBKGND 'prevent erase of left pane on horizontal resize
            If Resizing Then
                If fMainForm.ScaleHeight <> oldHeight Then
                    SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
                Else
                    Call GetUpdateRect(hwnd, tempRect, False)
                    tempRect.Right = fMainForm.sbStatusBar.Panels(1).Width _
                        \ Screen.TwipsPerPixelX
                    ValidateRect hwnd, tempRect
                    SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
' this commented code painted the grey border above the statusbar
'                    tempRect.Top = 0
'                    tempRect.Bottom = 40 \ Screen.TwipsPerPixelY
'                    tempRect.Left = 0
'                    tempRect.Right = fMainForm.sbStatusBar.Width \ Screen.TwipsPerPixelX
'                    hBrush% = CreateSolidBrush(SysColor2RGB(vbActiveBorder))
'                    RetVal% = FillRect(sberase_wParam, tempRect, hBrush%)
'                    RetVal% = DeleteObject(hBrush%)
                End If
            Else
                SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
            End If
        Case Else
            SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
        End Select
    Case Instances(3).hwnd   'listview
        Select Case Msg
        Case WM_LBUTTONDOWN  'new mscomctl 6.0 (J++ 6.0 preview 2)
                                'doesn't deselect on single _
                                'click of white space; this does it
            Dim tempobject As ListItem
             
            tempstr = Hex(lParam)

            X = Val("&H" & Right$(tempstr, 4))
            Y = Val("&H" & mID$(tempstr, 1, Len(tempstr) - 4))
           
            Set tempobject = fMainForm.lvListView.HitTest(X * _
                    Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY)
            
            If Not fMainForm.lvListView.SelectedItem Is Nothing And _
                    tempobject Is Nothing And lvNumListItemsSelected = 1 Then
                ListMouseClick = True
                fMainForm.lvListView.SetFocus
                fMainForm.lvListView.SelectedItem.Selected = False
                lvNumListItemsSelected = 0
            Else
                SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
            End If
'        Case WM_PAINT
'            If lvSaveEraseRect Then
'                Call GetUpdateRect(hWnd, tempRect, False)
'                tmpRectHeight = tempRect.Bottom - tempRect.Top
'                tmpRectWidth = tempRect.Right - tempRect.Left
'                RectTooBig = True
'                Select Case fMainForm.lvListView.View
'                Case lvwIcon
'                    If tmpRectWidth < 100 Then
'                        RectTooBig = False
'                    End If
'                Case lvwSmallIcon
'                    If tmpRectHeight < 50 Then
'                        RectTooBig = False
'                    End If
'                Case lvwReport
'                    If tmpRectHeight < 50 Then
'                        RectTooBig = False
'                    End If
'                Case lvwList
'                    If tmpRectHeight < 50 Then
'                        RectTooBig = False
'                    End If
'                End Select
'                If Not RectTooBig Then
'                    If (tmpRectHeight > RectHeight) Or (tmpRectWidth > RectWidth) Then
'                        RectHeight = tmpRectHeight
'                        RectWidth = tmpRectWidth
'                        m_UpdateRect.Bottom = tempRect.Bottom + (RectHeight * 0.25)
'                        m_UpdateRect.Right = tempRect.Right + (RectWidth * 0.25)
'                        m_UpdateRect.Top = tempRect.Top - 2     '2 pixels up to accomidate label edit
'                        If m_UpdateRect.Top < 0 Then
'                            m_UpdateRect.Top = 0
'                        End If
'                        m_UpdateRect.Left = tempRect.Left - 2   '2 pixels left to accomidate label edit
'                        If m_UpdateRect.Left < 0 Then
'                            m_UpdateRect.Left = 0
'                        End If
'                    End If
'                End If
'            End If
'            If lvNextNoErase Then
'                If lvNoEraseCnt > 1 Then
'                    lvNextNoErase = False
'                    lvSaveEraseRect = False
''                    Select Case fmainform.lvlistview.View
''                    Case lvwIcon, lvwSmallIcon
''                        fmainform.lvlistview.Arrange = lvwAutoTop
''                    Case Else
''                        fmainform.lvlistview.Arrange = lvwAutoLeft
''                    End Select
'                End If
'                '                    Debug.Print "top: " & m_UpdateRect.Top
' '                   Debug.Print "left: " & m_UpdateRect.Left
'  '                  Debug.Print "right: " & m_UpdateRect.Right
'   '                 Debug.Print "bottom: " & m_UpdateRect.Bottom
'                    'bkcolor = GetSysColor(vbWindowBackground And 255)
'                    'If bkcolor = 0 Then
'                    '    Debug.Print "color not defined"
'                    'End If
'                    hBrush% = CreateSolidBrush(SysColor2RGB(vbWindowBackground))
'                    RetVal% = FillRect(erase_wParam, m_UpdateRect, hBrush%)
'                    RetVal% = DeleteObject(hBrush%)
'            End If
'            If lvCountPaints Then
'                lvPaintCnt = lvPaintCnt + 1
'                Select Case lvPaintCnt
'                Case 2
'                '    lvCountPaints = True
'                'Case 3
'                    lvCountPaints = False
'                Case Else
'                    SwitchBoard = CallWindowProc(PrevWndProc, hWnd, Msg, wParam, lParam)
'                End Select
'            Else
'                SwitchBoard = CallWindowProc(PrevWndProc, hWnd, Msg, wParam, lParam)
'            End If
'        Case WM_ERASEBKGND
'            If lvEraseBkGnd1 Then            ' this prevents flicker in beginning of add list item
'                SwitchBoard = CallWindowProc(PrevWndProc, hWnd, Msg, wParam, lParam)
'            Else
'                erase_wParam = wParam
'                lvNoEraseCnt = lvNoEraseCnt + 1
'                If lvNoEraseCnt > 2 And Not lvSaveEraseRect Then
'                    lvEraseBkGnd1 = True
'                End If
'                If lvNoEraseCnt = 4 And Not fMainForm.lvListView.View = lvwIcon Then
'                    SwitchBoard = CallWindowProc(PrevWndProc, hWnd, Msg, wParam, lParam)
'                End If
'                If fMainForm.lvListView.View = lvwIcon Then
'                    If lvNoEraseCnt = 3 Or lvNoEraseCnt = 6 Then
'                        SwitchBoard = CallWindowProc(PrevWndProc, hWnd, Msg, wParam, lParam)
'                    End If
'                End If
'            End If
'        Case WM_PARENTNOTIFY
'            If wParam = &H20002 Then       'fwEvent=WM_DESTROY;IDchild:2
'               If Not lvEraseBkGnd2 Then
'                    If lvCancel Then
'                        lvEraseBkGnd1 = True
'                    Else
'                        lvNoEraseCnt = 0
'                        lvEraseBkGnd1 = False
'                        lvNextNoErase = True
'                    End If
'                    lvEraseBkGnd2 = True
'                End If
'            End If
'            SwitchBoard = CallWindowProc(PrevWndProc, hWnd, Msg, wParam, lParam)
        Case WM_COMMAND 'sets the listview label text rename limit to 64 char
            If wParam = &H4000001 Then       'wNotifyCode=EN_UPDATE;wID:1
                SendMessage lParam, EM_SETLIMITTEXT, 64, 0
                lvCancel = False
                SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
            Else
                SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
            End If
'        Case WM_KEYDOWN
'            If wParam = &H2D Then           'nVertKey:VK_INSERT,cRepeat:1
'                If fMainForm.lvListView.View = lvwIcon Then
'                    lvPaintCnt = 0
'                    lvCountPaints = True
'                End If
'            End If
'            SwitchBoard = CallWindowProc(PrevWndProc, hWnd, Msg, wParam, lParam)
'        Case WM_KEYUP
'            If wParam = &H1B Then           'nVertKey:VK_ESCAPE,cRepeat:1
'                lvSaveEraseRect = False
'                lvEraseBkGnd1 = True
'                lvEraseBkGnd2 = True
'                lvNextNoErase = False
'            End If
'            SwitchBoard = CallWindowProc(PrevWndProc, hWnd, Msg, wParam, lParam)
        Case &H204E     ' this sets lvNumListItemsSelected
            If TrapMultiSelectDrag Then
                SkipCnt = SkipCnt + 1
                If SkipCnt > 10 Then
                    tmpitemsSelected = SendMessage(fMainForm.lvListView.hwnd, _
                            LVM_GETSELECTEDCOUNT, 0&, 0&)
                    If tmpitemsSelected <> lvNumListItemsSelected Then
                        lvNumListItemsSelected = tmpitemsSelected
                        fMainForm.UpdateStatusBar lvNumListItemsSelected
                    End If
                    SkipCnt = 0
                End If
            End If
            SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
        Case Else
            SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
        End Select
    Case Instances(2).hwnd   ' treeview
        Select Case Msg
        Case WM_CHAR        'this does the expand/collapse '*' button in the treeview
            On Error Resume Next
            fMainForm.AsteriskKey wParam
            If wParam <> 42 Then
                SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
            End If
            On Error GoTo 0
        'Case WM_COMMAND
        '    If wParam = &H3000000 Then       'EN_UPDATE=&H40000000
        '        Dim txtlen As Integer
        '        txtlen = SendMessage(lParam, WM_GETTEXTLENGTH, 0, 0)
        '        Debug.Print "txtlen: " & txtlen
        '        'SendMessage lParam, WM_GETTEXT, txtlen, edit_text
                'Debug.Print "edit_text:  " & edit_text
        '    End If
        '    SwitchBoard = CallWindowProc(PrevWndProc, hWnd, Msg, wParam, lParam)
        'Case WM_PARENTNOTIFY
        '    If wParam = 1 Then
        '        old_lparam = lParam
        '    End If
        '    SwitchBoard = CallWindowProc(PrevWndProc, hWnd, Msg, wParam, lParam)
        Case Else
            SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
        End Select
        'Select Case MSG
'        Case WM_ENTERSIZEMOVE
'       '     MsgBox "entersizemove"
'            moving = True
'            count = 0
        '    LockWindowUpdate (fMainForm.hwnd)
            'count = count + 1
            'If count = 5 Then
'            SwitchBoard = 0
            '    count = 0
           'End If
        'Case WM_ERASEBKGND
        '    SwitchBoard = 1
'        Case WM_PAINT
'            If moving Then
'                count = count + 1
'                ValidateRect hwnd, Null
'                SwitchBoard = 0
'            Else
 '               SwitchBoard = CallWindowProc(PrevWndProc, hwnd, MSG, wParam, lParam)
'            End If
'        Case WM_EXITSIZEMOVE
'       '     MsgBox "exitsizemove"
'            If moving And count > 0 Then
'                moving = False
'                count = 0
'                InvalidateRect hwnd, Null, False
'        '    LockWindowUpdate (0&)
'            End If
'            SwitchBoard = CallWindowProc(PrevWndProc, hwnd, MSG, wParam, lParam)
        'Case WM_SIZE
        '    LockWindowUpdate (fMainForm.hwnd)
        '    SwitchBoard = CallWindowProc(PrevWndProc, hwnd, MSG, wParam, lParam)
        'Case Else
        '    SwitchBoard = CallWindowProc(PrevWndProc, hWnd, MSG, wParam, lParam)
        'End Select
    Case Instances(1).hwnd   'Form
        Select Case Msg
        Case WM_INITMENUPOPUP  'this popups up the add submenu
            If lParam = 0 And isInsertKey Then
                SendKeys "a"
                isInsertKey = False
            End If
            SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
        Case WM_ENTERSIZEMOVE
            Resizing = True
            SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
        Case WM_EXITSIZEMOVE
            Resizing = False
            'fMainForm.lvListView.Top = fMainForm.picTitles.Top + fMainForm.picTitles.Height + 1
            'fMainForm.lvListView.Width = fMainForm.Width - (fMainForm.tvTreeView.Width + 125)
            'fMainForm.lvListView.Height = fMainForm.tvTreeView.Height + 30
            SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
            'fmainform.lblTitle(1).Width = fmainform.lvListView.Width
            'fmainform.sbStatusBar.Align = vbAlignBottom
'            fmainform.lvListView.Width = fmainform.ScaleWidth - fmainform.lvListView.Left
'            fmainform.lvListView.Height = fmainform.ScaleHeight - fmainform.lvListView.Top _
'                - fmainform.sbStatusBar.Height
'            fmainform.lblTitle(1).Width = fmainform.lvListView.Width
'            fmainform.tvTreeView.Height = fmainform.lvListView.Height
'            fmainform.sbStatusBar.Align = vbAlignBottom
'            SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
        Case WM_GETMINMAXINFO
            CopyMemory tMinMax, ByVal lParam, Len(tMinMax)
            ' Set your values for MinX, MinY, MaxX and MaxY in  pixels
            With tMinMax
                .ptMinTrackSize.X = (fMainForm.lvListView.Left + 1000) \ Screen.TwipsPerPixelX
                .ptMinTrackSize.Y = (fMainForm.lblTitle(1).Top + fMainForm.lblTitle(1).Height + 2000) _
                                        \ Screen.TwipsPerPixelY
                .ptMaxTrackSize.X = Screen.Width \ Screen.TwipsPerPixelX
                .ptMaxTrackSize.Y = Screen.Height \ Screen.TwipsPerPixelY
            End With
            CopyMemory ByVal lParam, tMinMax, Len(tMinMax)
        Case WM_SIZE
            If fMainForm.WindowState = vbMinimized Then
                WasMinimized = True
            Else
                If WasMinimized Then
                    fMainForm.SizeControls LastSplitterLeft
                    WasMinimized = False
                Else
                    'fMainForm.picTitles.Top = fMainForm.picTBContainer.Top + fMainForm.picTBContainer.Height + 100
                    fMainForm.picTitles.Top = fMainForm.CoolBar1.Top + fMainForm.CoolBar1.Height + 100
                    fMainForm.lvListView.Top = fMainForm.picTitles.Top + fMainForm.picTitles.Height + 1
                    fMainForm.tvTreeView.Top = fMainForm.lvListView.Top + 20
                    fMainForm.lvListView.Width = fMainForm.Width - (fMainForm.tvTreeView.Width + 125)
                    fMainForm.tvTreeView.Height = fMainForm.ScaleHeight - fMainForm.lvListView.Top _
                                                - fMainForm.sbStatusBar.Height - 35
                    fMainForm.lvListView.Height = fMainForm.tvTreeView.Height + 30
                    'fMainForm.tbToolBar.Width = fMainForm.ScaleWidth - 1
                    'fMainForm.picTBContainer.Width = fMainForm.ScaleWidth
                    'fMainForm.tbToolBar.Width = fMainForm.ScaleWidth
                    'fMainForm.tbToolBar.Width = fMainForm.tbToolBar.Buttons(fMainForm.tbToolBar.Buttons.count).Left
                    fMainForm.picTitles.Width = fMainForm.ScaleWidth
                    fMainForm.lblTitle(1).Width = fMainForm.lvListView.Width - 30
                    fMainForm.sbStatusBar.Align = vbAlignBottom
                End If
            End If
        Case WM_SHOWWINDOW      ' for case when user presses stop button in IDE
            If lParam = 0 And wParam = 0 Then
                For i = MIN_INSTANCES To MAX_INSTANCES
                    If Instances(i).in_use Then
                        UnHookWindow (i)
                    Instances(i).in_use = False
                    End If
                Next i
            End If
            SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
        Case WM_ENTERIDLE      ' for case when error dialog pops up and user presses debug
            If wParam = 0 Then
                For i = MIN_INSTANCES To MAX_INSTANCES
                    If Instances(i).in_use Then
                        UnHookWindow (i)
                    Instances(i).in_use = False
                    End If
                Next i
            End If
            SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
        Case WM_SETCURSOR       ' for case when user exits using x button in top right corner
            If lParam = &H2010014 Then      'hex for hittest:htclose & lbutton
                For i = MIN_INSTANCES To MAX_INSTANCES
                    If Instances(i).in_use Then
                        UnHookWindow (i)
                    Instances(i).in_use = False
                    End If
                Next i
            End If
            SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
        
        'This procedure is called because we've subclassed
        'this form. We will catch DRAWITEM and MEASUREITEM
        'messages and pass the rest of them on.
        Case WM_DRAWITEM
            CopyMemory DrawInfo, ByVal lParam, LenB(DrawInfo)
            If wParam = 0 And DrawInfo.CtlType = 1 Then     'It was sent by the menu
                
                If Caps(DrawInfo.itemID) = "" Then
                    IsSep = True
                Else
                    IsSep = False
                    MyFont = SendMessage(hwnd, WM_GETFONT, 0&, 0&)
                    OldFont = SelectObject(DrawInfo.hdc, MyFont)
               End If
                
                
                If DrawInfo.itemState = ODS_SELECTED Then
                    hBr = CreateSolidBrush(SysColor(HIGHLIGHT))
                    hPen = GetPen(1, SysColor(HIGHLIGHT))
                    lTextColor = SysColor(HIGHLIGHTTEXT)
                Else
                    hBr = CreateSolidBrush(SysColor(Menu))
                    hPen = GetPen(1, SysColor(Menu))
                    lTextColor = SysColor(MENUTEXT)
                End If
                
                'We're going to draw on the menu
                mQuickGDI.TargethDC = DrawInfo.hdc
                
                'Select our new, correctly colored objects:
                hOldBr = SelectObject(DrawInfo.hdc, hBr)
                hOldPen = SelectObject(DrawInfo.hdc, hPen)
                
                With DrawInfo.rcItem
                    If DrawInfo.itemState <> ODS_SELECTED Then
                        'Clear the space where the image is
                        mQuickGDI.DrawRect .Left, .Top, 22, .Bottom
                    End If
                    

                    
                    If Not IsSep Then
                        'Check to see if the menu item is one of the ones
                        'with a picture. If so, then we need to move the
                        'edge of the drawing rectangle a little to the
                        'left to make room for the image.
                        iRectOffset = IIf(HasIcon(DrawInfo.itemID) <> 0, 23, 0)
                        
                        'Draw the rectangle onto the item's space
                        mQuickGDI.DrawRect .Left + iRectOffset, .Top, .Right, .Bottom
                        
                        'Print the item's text (held in the Caps() array)
                        hPrint .Left + 25, .Top + 3, Caps(DrawInfo.itemID), lTextColor
                    End If
                End With
                
                'Select the old objects into the menu's DC
                Call SelectObject(DrawInfo.hdc, hOldBr)
                Call SelectObject(DrawInfo.hdc, hOldPen)
                
                'Delete the ones we created
                Call DeleteObject(hBr)
                Call DeleteObject(hPen)
                With DrawInfo

                    If HasIcon(DrawInfo.itemID) <> 0 Then
                        On Error Resume Next
                        TestIndex = -1
                        TestIndex = fMainForm.imlMenu.ListImages.item(HasIcon(DrawInfo.itemID)).Index
                        If TestIndex <> -1 Then
                            pnt.PaintTransparentStdPic .hdc, 4, .rcItem.Top + 2, 16, 16, _
                                fMainForm.imlMenu.ListImages.item(HasIcon(DrawInfo.itemID)) _
                                .picture, 0, 0, &HC0C0C0
                        End If
                        On Error GoTo 0
                        'If this item is selected, draw a raised
                        'box around the image
                        If DrawInfo.itemState = ODS_SELECTED Then
                            ThreedBox 1, .rcItem.Top, 21, .rcItem.Bottom - 1
                        End If
                    End If

                    If IsSep Then
                        'Draw the special separator bar
                        ThreedBoxInvert .rcItem.Left, .rcItem.Top, .rcItem.Right - 1, _
                                .rcItem.Bottom - 4
                    End If
                End With

            Else
                SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
            End If
            CopyMemory ByVal lParam, DrawInfo, Len(DrawInfo)

        Case WM_MEASUREITEM
            CopyMemory MeasureInfo, ByVal lParam, Len(MeasureInfo)
            If wParam = 0 And MeasureInfo.CtlType = 1 Then 'It was sent by the menu
                'Get the MEASUREITEM struct from the pointer
                
                If Caps(MeasureInfo.itemID) = "" Then
                    IsSep = True
                Else
                    IsSep = False
                End If
                
                'Tell Windows how big our items are.
                If MeasureInfo.itemID > 6 Then ' Edit menu
                    MeasureInfo.itemWidth = 100
                Else ' File Menu
                    MeasureInfo.itemWidth = 80
                End If
                'If the item being measured is the separator
                'bar, the height should be 5 pixels, 18 if
                'otherwise...
                MeasureInfo.ItemHeight = IIf(IsSep, 5, 20)
                'Return the information back to Windows
                
                'Don't pass this message on:
            Else
                SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
            End If
            CopyMemory ByVal lParam, MeasureInfo, Len(MeasureInfo)
'        Case WM_NOTIFY 'Needed to let us know when mouse has anything to do with Rebar
            'Copy hdr info so we can determine if uMsg is coming from Rebar
            'CopyMemory hdr, ByVal lParam, Len(hdr)
'            CopyMemory RebarHdr, ByVal lParam, Len(RebarHdr)
            'Check hwndFrom (handle of window sending message)
            'If hdr.hwndFrom = Rebar.GetRebarWindow Then
'            If RebarHdr.NMHDR = Rebar.GetRebarWindow Then
'                Call GetCursorPos(ptApi)
'                Call ScreenToClient(hWnd, ptApi)
'                BandInfo.ptApi = ptApi
'                BandInfo.flags = RBHT_CAPTION Or RBHT_GRABBER Or RBHT_CLIENT
'                Call SendMessage(Rebar.GetRebarWindow, RB_HITTEST, 0, BandInfo)
                'Yes it's ours
                '8386744 = Being Sized
                '8387324 = ClickUp anywhere on rebar or gripper
                'If you don't do this when using the toolbar control, then
                'whenever you touch the Rebar or size the bands then
                'toolbars will dissappear.
                'Alot of Flicker
'                If lParam = 8386744 Then
'                    fmainform.tbToolBar.Refresh
'                End If
'                If lParam = 8387324 Then    'older Version
'                    fmainform.tbToolBar.Refresh
'                End If
'               If lParam = 8386792 Then
'                    fmainform.tbToolBar.Refresh
'                End If
'            Else
'                SwitchBoard = CallWindowProc(PrevWndProc, hWnd, Msg, wParam, lParam)
'            End If
'        Case WM_SYSCOLORCHANGE      'Case User changes colors while this is running
'            Rebar.SetBandColors
'            SwitchBoard = CallWindowProc(PrevWndProc, hWnd, Msg, wParam, lParam)
        Case Else
            SwitchBoard = CallWindowProc(PrevWndProc, hwnd, Msg, wParam, lParam)
        End Select
    End Select
        
End Function

