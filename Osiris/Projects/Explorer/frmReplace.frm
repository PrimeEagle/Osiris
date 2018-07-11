VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmReplace 
   Caption         =   "Form1"
   ClientHeight    =   3750
   ClientLeft      =   6930
   ClientTop       =   4830
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   6585
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   6165
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Name && Location"
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cboLookIn"
      Tab(0).Control(1)=   "cbMatchCase"
      Tab(0).Control(2)=   "cbWhole"
      Tab(0).Control(3)=   "cboFind"
      Tab(0).Control(4)=   "cboReplace"
      Tab(0).Control(5)=   "lblLookIn"
      Tab(0).Control(6)=   "lblFind"
      Tab(0).Control(7)=   "lblReplace"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Date"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboCreatedBy"
      Tab(1).Control(1)=   "cboCreatedDate"
      Tab(1).Control(2)=   "cboCreatedTime"
      Tab(1).Control(3)=   "cboLMTime"
      Tab(1).Control(4)=   "cboLMDate"
      Tab(1).Control(5)=   "cboModifiedBy"
      Tab(1).Control(6)=   "lblCreatedDate"
      Tab(1).Control(7)=   "lblCreatedTime"
      Tab(1).Control(8)=   "lblCreated"
      Tab(1).Control(9)=   "lblLastModified"
      Tab(1).Control(10)=   "lblLMTime"
      Tab(1).Control(11)=   "lblLMDate"
      Tab(1).Control(12)=   "lblCreatedBy"
      Tab(1).Control(13)=   "lblModifiedBy"
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Advanced"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cboCreateNode"
      Tab(2).Control(1)=   "cboCreateItem"
      Tab(2).Control(2)=   "cboSystem"
      Tab(2).Control(3)=   "cboReadOnly"
      Tab(2).Control(4)=   "cboTableName"
      Tab(2).Control(5)=   "tboCntTxt"
      Tab(2).Control(6)=   "lblCreateNode"
      Tab(2).Control(7)=   "lblCreateItem"
      Tab(2).Control(8)=   "lblSystem"
      Tab(2).Control(9)=   "lblReadOnly"
      Tab(2).Control(10)=   "lblTableName"
      Tab(2).Control(11)=   "lblCntTxt"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Linking"
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lblLinkNode"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblQuickType"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lblLinked"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cboLinkNode"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cboQuickType"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "cboLinked"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).ControlCount=   6
      Begin VB.ComboBox cboLinked 
         Height          =   315
         Left            =   1680
         TabIndex        =   44
         Top             =   1155
         Width           =   885
      End
      Begin VB.ComboBox cboCreatedBy 
         Height          =   315
         Left            =   -73710
         TabIndex        =   43
         Top             =   1380
         Width           =   2685
      End
      Begin VB.ComboBox cboCreatedDate 
         Height          =   315
         Left            =   -73740
         TabIndex        =   39
         Top             =   945
         Width           =   1215
      End
      Begin VB.ComboBox cboCreatedTime 
         Height          =   315
         Left            =   -71700
         TabIndex        =   38
         Top             =   945
         Width           =   1215
      End
      Begin VB.ComboBox cboLMTime 
         Height          =   315
         Left            =   -71670
         TabIndex        =   34
         Top             =   2220
         Width           =   1215
      End
      Begin VB.ComboBox cboCreateNode 
         Height          =   315
         Left            =   -71520
         TabIndex        =   32
         Top             =   2400
         Width           =   885
      End
      Begin VB.ComboBox cboCreateItem 
         Height          =   315
         Left            =   -71520
         TabIndex        =   30
         Top             =   2880
         Width           =   885
      End
      Begin VB.ComboBox cboSystem 
         Height          =   315
         Left            =   -73800
         TabIndex        =   28
         Top             =   2880
         Width           =   885
      End
      Begin VB.ComboBox cboReadOnly 
         Height          =   315
         Left            =   -73800
         TabIndex        =   26
         Top             =   2400
         Width           =   885
      End
      Begin VB.ComboBox cboQuickType 
         Height          =   315
         Left            =   1680
         TabIndex        =   24
         Top             =   600
         Width           =   2595
      End
      Begin VB.ComboBox cboLinkNode 
         Height          =   315
         Left            =   1680
         TabIndex        =   22
         Top             =   1560
         Width           =   2595
      End
      Begin VB.ComboBox cboTableName 
         Height          =   315
         Left            =   -73560
         TabIndex        =   20
         Top             =   660
         Width           =   2565
      End
      Begin VB.ComboBox cboLMDate 
         Height          =   315
         Left            =   -73710
         TabIndex        =   17
         Top             =   2220
         Width           =   1215
      End
      Begin VB.ComboBox cboModifiedBy 
         Height          =   315
         Left            =   -73710
         TabIndex        =   16
         Top             =   2670
         Width           =   2685
      End
      Begin VB.ComboBox cboLookIn 
         Height          =   315
         Left            =   -73800
         TabIndex        =   14
         Top             =   1845
         Width           =   2685
      End
      Begin VB.TextBox tboCntTxt 
         Height          =   345
         Left            =   -73560
         TabIndex        =   12
         Top             =   1845
         Width           =   2535
      End
      Begin VB.CheckBox cbMatchCase 
         Caption         =   "Matc&h case"
         Height          =   255
         Left            =   -74820
         TabIndex        =   9
         Top             =   2340
         Width           =   1335
      End
      Begin VB.CheckBox cbWhole 
         Caption         =   "Find &whole words only"
         Height          =   255
         Left            =   -74820
         TabIndex        =   8
         Top             =   2700
         Width           =   1935
      End
      Begin VB.ComboBox cboFind 
         Height          =   315
         Left            =   -73800
         TabIndex        =   7
         Top             =   885
         Width           =   2685
      End
      Begin VB.ComboBox cboReplace 
         Height          =   315
         Left            =   -73800
         TabIndex        =   6
         Top             =   1365
         Width           =   2685
      End
      Begin VB.Label lblLinked 
         Caption         =   "Linked:"
         Height          =   255
         Left            =   945
         TabIndex        =   45
         Top             =   1185
         Width           =   615
      End
      Begin VB.Label lblCreatedDate 
         Caption         =   "Date:"
         Height          =   255
         Left            =   -74310
         TabIndex        =   42
         Top             =   975
         Width           =   495
      End
      Begin VB.Label lblCreatedTime 
         Caption         =   "Time:"
         Height          =   255
         Left            =   -72270
         TabIndex        =   41
         Top             =   975
         Width           =   495
      End
      Begin VB.Label lblCreated 
         Caption         =   "Created:"
         Height          =   255
         Left            =   -74790
         TabIndex        =   40
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label lblLastModified 
         Caption         =   "Last Modified:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   37
         Top             =   1935
         Width           =   1215
      End
      Begin VB.Label lblLMTime 
         Caption         =   "Time:"
         Height          =   255
         Left            =   -72240
         TabIndex        =   36
         Top             =   2250
         Width           =   495
      End
      Begin VB.Label lblLMDate 
         Caption         =   "Date:"
         Height          =   255
         Left            =   -74280
         TabIndex        =   35
         Top             =   2250
         Width           =   495
      End
      Begin VB.Label lblCreateNode 
         Caption         =   "Create Node:"
         Height          =   255
         Left            =   -72600
         TabIndex        =   33
         Top             =   2430
         Width           =   975
      End
      Begin VB.Label lblCreateItem 
         Caption         =   "Create Item:"
         Height          =   255
         Left            =   -72600
         TabIndex        =   31
         Top             =   2910
         Width           =   855
      End
      Begin VB.Label lblSystem 
         Caption         =   "System:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   29
         Top             =   2910
         Width           =   615
      End
      Begin VB.Label lblReadOnly 
         Caption         =   "Read Only:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   27
         Top             =   2430
         Width           =   1095
      End
      Begin VB.Label lblQuickType 
         Caption         =   "Quick Type:"
         Height          =   255
         Left            =   480
         TabIndex        =   25
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label lblLinkNode 
         Caption         =   "Link Node:"
         Height          =   255
         Left            =   750
         TabIndex        =   23
         Top             =   1605
         Width           =   855
      End
      Begin VB.Label lblTableName 
         Caption         =   "Table Name:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   21
         Top             =   690
         Width           =   1095
      End
      Begin VB.Label lblCreatedBy 
         Caption         =   "Created By:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   19
         Top             =   1395
         Width           =   1095
      End
      Begin VB.Label lblModifiedBy 
         Caption         =   "Modified by:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   18
         Top             =   2700
         Width           =   975
      End
      Begin VB.Label lblLookIn 
         Caption         =   "Look in:"
         Height          =   255
         Left            =   -74820
         TabIndex        =   15
         Top             =   1860
         Width           =   750
      End
      Begin VB.Label lblCntTxt 
         Caption         =   "Containing Text:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   13
         Top             =   1890
         Width           =   1200
      End
      Begin VB.Label lblFind 
         Caption         =   "Find what:"
         Height          =   255
         Left            =   -74820
         TabIndex        =   11
         Top             =   915
         Width           =   1095
      End
      Begin VB.Label lblReplace 
         Caption         =   "Replace with:"
         Height          =   255
         Left            =   -74820
         TabIndex        =   10
         Top             =   1395
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   5520
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2280
      Width           =   540
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Replace &All"
      Height          =   375
      Left            =   5040
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "&Replace"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "&Find Next"
      Default         =   -1  'True
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAX_COMBO_ITEMS = 10

Dim FoundFirstNode As Boolean
Dim LVwrapped As Boolean
Dim StartFindNodeID As Integer
Dim StartFindItemID As Long
Dim FindLookinNodeIndex As Long
Dim TempPos As Long
Dim EndOfFind As Boolean
Dim ReplacingAll As Boolean
Dim OrderByStr As String
Dim WhereData_LabelStr As String
Dim WhereData_ValueStr As String

Dim FindRecords As Recordset
Dim LookinTableNames As New Collection
Dim Table_Names() As String
Dim LastCntTxt As String

Private Sub TestFindLV(SearchTxt As String, CompareType As Integer)

Dim CurrentNode As Node
Dim table_name As String
Dim QueryStr As String
Dim i As Integer
Dim j As Integer
Dim WhereData_IDStr As String
Dim FROMStr As String
Dim ParentNodeStr As String
Dim listrecord As Recordset
Static StartItemIsMatch As Boolean
Static StartItemPosition As Long

        If StartFindItemIndex <> 0 Or StartFindNodeIndex <> 0 Then
            GoTo NextItem
        End If
        If CurrentItem Is Nothing Then
            If fMainForm.lvListView.ListItems.count > 0 Then
                StartFindItemIndex = fMainForm.lvListView.ListItems(1).Index
                StartFindItemID = RemoveK(fMainForm.lvListView.ListItems(1).key)
            Else
                StartFindItemIndex = 0  'start item is when parent_node passes startfindnode
                StartFindItemID = 0
            End If
        Else
            StartFindItemIndex = CurrentItem.Index
            StartFindItemID = RemoveK(CurrentItem.key)
        End If
        If nodx.FullPath = cboLookIn.text Then
            Set CurrentNode = nodx
            FindLookinNodeIndex = nodx.Index
            StartFindNodeIndex = nodx.Index
            StartFindNodeID = RemoveK(nodx.key)
        Else
            Set CurrentNode = GetNodeFromFullPath(nodx.Root, cboLookIn.text)
            If CurrentNode Is Nothing Then
                GoTo Done
            Else
                FindLookinNodeIndex = CurrentNode.Index
                If IsDescendantOf(nodx, CurrentNode) Then
                    StartFindNodeIndex = nodx.Index
                    StartFindNodeID = RemoveK(nodx.key)
                Else
                    StartFindNodeIndex = CurrentNode.Index
                    StartFindNodeID = RemoveK(CurrentNode.key)
                    StartFindItemIndex = -1 'start item is first record
                    StartFindItemID = -1
                End If
            End If
        End If
        While 1
            table_name = tvAttribCol(CurrentNode.key).table_name
            AddTableNameIfNotInCol table_name
            LookinTableNames(table_name).Add RemoveK(CurrentNode.key)
            If Not CurrentNode.Child Is Nothing Then
                Set CurrentNode = CurrentNode.Child
            Else
                If CurrentNode.Index = FindLookinNodeIndex Then
                    GoTo DoneFillCol
                End If
                If Not CurrentNode.Next Is Nothing Then
                    Set CurrentNode = CurrentNode.Next
                Else
                    Set CurrentNode = NextofParent(CurrentNode, FindLookinNodeIndex)
                    If CurrentNode Is Nothing Then
                        GoTo DoneFillCol
                    End If
                End If
            End If
        Wend
DoneFillCol:
        For i = 1 To LookinTableNames.count
            table_name = Table_Names(i)
            ParentNodeStr = "Parent_Node In ("
            For j = 1 To LookinTableNames(table_name).count
                ParentNodeStr = ParentNodeStr & _
                        LookinTableNames(table_name)(j) & ","
            Next j
            ParentNodeStr = Left$(ParentNodeStr, Len(ParentNodeStr) - 1)
            ParentNodeStr = ParentNodeStr & ")"
            If table_name = tvAttribCol(nodx.key).table_name And _
                    StartFindItemIndex > 0 Then
                WhereData_IDStr = "Data_Id = " & StartFindItemID
            Else
                WhereData_IDStr = "FALSE"
            End If
            QueryStr = QueryStr & "SELECT ALL" _
                & " Data_Label,Data_ID,Parent_Node,[Order]" _
                & " FROM " & table_name _
                & " WHERE (" & WhereData_IDStr _
                & " OR (" & WhereData_LabelStr & " AND " _
                & WhereData_ValueStr _
                & ")) AND (" & ParentNodeStr & ")" _
                & " UNION "
        Next i
        QueryStr = Left$(QueryStr, Len(QueryStr) - 7)
        QueryStr = QueryStr & " ORDER BY [Order], " & OrderByStr
        'Clipboard.SetText QueryStr
        Set FindRecords = dbase.OpenRecordset(QueryStr, dbOpenDynaset)
        If FindRecords.EOF Then
            GoTo AtStart
        End If
        If StartFindItemIndex = -1 Then
            StartItemIsMatch = True
            StartItemPosition = FindRecords.AbsolutePosition
            GoTo FirstMatches
        End If
        While Not FindRecords.EOF     'move to start item in recordset
            If StartFindItemIndex = 0 Then
                If FindRecords!parent_node > StartFindNodeID Then
                    StartItemIsMatch = True
                    StartItemPosition = FindRecords.AbsolutePosition
                    GoTo FirstMatches
                End If
            Else
                If FindRecords!data_id = StartFindItemID And _
                        FindRecords!parent_node = RemoveK(nodx.key) Then
                    GoTo ChkifStartMatches
                End If
            End If
            FindRecords.MoveNext
        Wend
        GoTo AtStart
ChkifStartMatches:
        StartItemPosition = FindRecords.AbsolutePosition
        WhereData_IDStr = "Data_Id = " & StartFindItemID
        ParentNodeStr = "Parent_Node = " & FindRecords!parent_node
        QueryStr = "SELECT ALL Data_ID, Data_Label" _
                & " FROM " & tvAttribCol(AddK(FindRecords!parent_node)).table_name _
                & " WHERE (" & WhereData_IDStr _
                & " AND (" & WhereData_LabelStr & " AND " _
                & WhereData_ValueStr _
                & ")) AND (" & ParentNodeStr & ")"
        Set listrecord = dbase.OpenRecordset(QueryStr, dbOpenDynaset)
        If Not listrecord.EOF Then
            If CompareType = vbBinaryCompare Then
                If InStr(1, listrecord!data_label, SearchTxt, _
                        CompareType) > 0 Then
                    GoTo StartIsMatch
                Else
                    GoTo StartNotMatch
                End If
            End If
StartIsMatch:
            StartItemIsMatch = True
            If RemoveK(nodx.key) = FindRecords!parent_node Then
                CurrentItem.Selected = False
                DoEvents
            End If
            GoTo FoundItem1
        End If
StartNotMatch:
        StartItemIsMatch = False
NextItem:
        FindRecords.MoveNext
FirstMatches:
        If FindRecords.EOF Then
            FindRecords.MoveFirst
            LVwrapped = True
        End If
        If FindRecords.AbsolutePosition = StartItemPosition _
                And LVwrapped Then
            GoTo AtStart
        End If
        If CompareType = vbBinaryCompare Then
            If InStr(1, FindRecords!data_label, SearchTxt, _
                    CompareType) <= 0 Then
                GoTo NextItem
            End If
        End If
FoundItem:
        If RemoveK(nodx.key) <> FindRecords!parent_node Then
            Set nodx = fMainForm.tvTreeView.Nodes(AddK(FindRecords!parent_node))
            Set fMainForm.tvTreeView.SelectedItem = nodx
            nodx.EnsureVisible
            lvNeedsRefresh = True
            lvLoaded = False
            fMainForm.tmrNodeClick.Interval = 1
            fMainForm.tvTreeView_NodeClick nodx
            While Not lvLoaded
                DoEvents
            Wend
            fMainForm.tmrNodeClick.Interval = NODE_CLICK_TIMER_DELAY
        Else
            If Not CurrentItem Is Nothing Then
                If RemoveK(CurrentItem.key) = FindRecords!data_id Then
                    CurrentItem.Selected = False
                    DoEvents
                End If
            End If
        End If
FoundItem1:
        Set CurrentItem = fMainForm.lvListView.ListItems(AddK(FindRecords!data_id))
        Set fMainForm.lvListView.SelectedItem = CurrentItem
        CurrentItem.EnsureVisible
        GoTo Done
AtStart:
        If Not ReplacingAll Then
            MsgBox "The specified region has been searched.", vbInformation
        End If
        EndOfFind = True
        LVwrapped = False
        If StartItemIsMatch Then
            FindRecords.MovePrevious
        End If
        
Done:
End Sub

Private Sub AddTableNameIfNotInCol(table_name As String)

    Dim tempobj As Variant
    Dim tempcol As Collection
    
        On Error GoTo TableNotInCol
        Set tempobj = LookinTableNames(table_name)
        Exit Sub

TableNotInCol:
        Set tempcol = New Collection
        LookinTableNames.Add tempcol, table_name
        ReDim Preserve Table_Names(1 To LookinTableNames.count)
        Table_Names(LookinTableNames.count) = table_name

End Sub


Public Function GetNodeFromFullPath(RootNode As Node, FullPath As String) _
        As Node
    
    Dim CurrentNode As Node
    Dim i As Long
    Dim lastpos As Long
    Dim NodeName As String
    
    
    Set CurrentNode = RootNode
    If Left$(FullPath, 1) = "\" Then
        lastpos = 1
    Else
        lastpos = 0
    End If
    
    While 1
        i = InStr(lastpos + 1, FullPath, "\", vbTextCompare)
        If i = 0 Then
            i = Len(FullPath) + 1
        End If
        If i - lastpos - 1 < 1 Then
            Set GetNodeFromFullPath = CurrentNode
            GoTo Done
        End If
        NodeName = mID$(FullPath, lastpos + 1, i - lastpos - 1)
        Set CurrentNode = FindNode(CurrentNode, NodeName)
        If CurrentNode Is Nothing Then
            Set GetNodeFromFullPath = Nothing
            GoTo Done
        End If
        If i = Len(FullPath) + 1 Then
            Set GetNodeFromFullPath = CurrentNode
            GoTo Done
        End If
        Set CurrentNode = CurrentNode.Child
        lastpos = i
    Wend
    
Done:
End Function

Private Function FindNode(StartNode As Node, NodeName As String) As Node
Dim CurrentNode As Node
    
    Set CurrentNode = StartNode
    While CurrentNode.text <> NodeName
        If CurrentNode.Index = CurrentNode.LastSibling.Index Then
            GoTo NodeNotFound
        End If
        Set CurrentNode = CurrentNode.Next
    Wend
    Set FindNode = CurrentNode
    GoTo Done

NodeNotFound:
    MsgBox "FindNode:  The Look-in node, " & NodeName & ", does not exist", _
            vbInformation
    Set FindNode = Nothing
Done:
End Function

Private Sub cbLink_Click()
    ResetFind
End Sub

Private Sub cboLinked_Validate(Cancel As Boolean)
    If cboLinked.text = "True" Then
        cboLinkNode.Visible = True
    Else
        cboLinkNode.Visible = False
    End If

End Sub

Private Sub cboLookIn_Change()
    ResetFind
End Sub

Private Sub cboLookIn_Click()
    ResetFind
End Sub

Private Sub CmdCancel_Click()
    FindEnabled = False
    Me.Hide
End Sub

Public Sub cmdFindNext_Click()
    Dim SearchTxt As String
    Dim CompareType As Integer
        
    LastFind = cboFind.text
    LastReplace = cboReplace.text
    LastCase = cbMatchCase.Value
    LastWhole = cbWhole.Value
    
    LastCntTxt = tboCntTxt.text
    
    UpdateComboItems
    
    If cbMatchCase.Value = 1 Then
        CompareType = vbBinaryCompare
    Else
        CompareType = vbTextCompare
    End If
    SearchTxt = cboFind.text
    
    If TypeOf fMainForm.ActiveControl Is TreeView Then
        FindTV SearchTxt, CompareType
    ElseIf TypeOf fMainForm.ActiveControl Is ListView Then
        fMainForm.lvListView.MultiSelect = False
        If SearchTxt = "" Then
            WhereData_LabelStr = "TRUE"
        Else
            If cbWhole.Value = 1 Then
                WhereData_LabelStr = _
                "Data_Label LIKE '*[!0-z]" & SearchTxt & "[!0-z]*'" _
                & " OR Data_Label LIKE '" & SearchTxt & "[!0-z]*'" _
                & " OR Data_Label LIKE '*[!0-z]" & SearchTxt & "'" _
                & " OR Data_Label LIKE '" & SearchTxt & "'"
            Else
                WhereData_LabelStr = "Data_Label LIKE '*" & SearchTxt & "*'"
            End If
        End If
        If LastCntTxt = "" Then
            WhereData_ValueStr = "TRUE"
        Else
            WhereData_ValueStr = "Data_Value LIKE '*" & LastCntTxt & "*'"
        End If
        Select Case fMainForm.lvListView.SortKey
        Case ITEM_COLUMN
            OrderByStr = "data_label"
        Case DATA_ID_COLUMN
            OrderByStr = "data_id"
        Case PARENT_NODE_COLUMN
            OrderByStr = "parent_node"
        Case TYPE_COLUMN
            OrderByStr = "data_type"
        End Select
        If fMainForm.lvListView.SortOrder = lvwAscending Then
            OrderByStr = OrderByStr & " ASC"
        Else
            OrderByStr = OrderByStr & " DESC"
        End If
        'FindLV SearchTxt, CompareType
        TestFindLV SearchTxt, CompareType
        fMainForm.lvListView.MultiSelect = True
    Else
        MsgBox "You cannot do a find/replace from this control.", vbExclamation
    End If
Done:
End Sub

Private Sub FindTV(SearchTxt As String, CompareType As Integer)
    
Dim CurrentNode As Node
Dim err As Integer

        If StartFindNodeIndex = 0 Then
            If nodx.FullPath = cboLookIn.text Then
                If nodx.Child Is Nothing Then
                    If Not ReplacingAll Then
                        MsgBox "The entire node tree was searched.", vbInformation
                    End If
                    EndOfFind = True
                    FoundFirstNode = False
                    ResetFind
                    GoTo Done
                Else
                    StartFindNodeIndex = nodx.Child.Index
                    StartFindNodeID = RemoveK(nodx.Child.key)
                    FindLookinNodeIndex = nodx.Index
                    Set CurrentNode = nodx.Child
                End If
            Else
                Set CurrentNode = GetNodeFromFullPath(nodx.Root, cboLookIn.text)
                If CurrentNode Is Nothing Then
                    GoTo Done
                Else
                    FindLookinNodeIndex = CurrentNode.Index
                    If IsDescendantOf(nodx, CurrentNode) Then
                        StartFindNodeIndex = nodx.Index
                        StartFindNodeID = RemoveK(nodx.key)
                    Else
                        StartFindNodeIndex = CurrentNode.Index
                        StartFindNodeID = RemoveK(CurrentNode.key)
                        GoTo NextNode1
                    End If
                End If
            End If
            TempPos = InStr(1, CurrentNode.text, SearchTxt, CompareType)
            If TempPos > 0 Then
                err = ChkForNodeAttribMatch(CurrentNode, SearchTxt)
                If err = -1 Then
                    GoTo Done
                End If
                If err = 1 Then
                    GoTo FoundNode
                End If
            End If
        End If
        Set CurrentNode = nodx
        If CurrentNode.Index = StartFindNodeIndex And FoundFirstNode Then
            GoTo AtStart
        End If
NextNode1:
        While 1
            If Not CurrentNode.Child Is Nothing Then
                Set CurrentNode = CurrentNode.Child
            Else
                If Not CurrentNode.Next Is Nothing Then
                    Set CurrentNode = CurrentNode.Next
                Else
                    Set CurrentNode = NextofParent(CurrentNode, FindLookinNodeIndex)
                    If CurrentNode Is Nothing Then
                        GoTo WraptoBeginning
                    End If
                End If
            End If
            TempPos = InStr(1, CurrentNode.text, _
                    SearchTxt, CompareType)
            If TempPos > 0 Then
                err = ChkForNodeAttribMatch(CurrentNode, SearchTxt)
                If err = -1 Then
                    GoTo Done
                End If
                If err = 1 Then
                    FoundFirstNode = True
                    GoTo FoundNode
                End If
            End If
            If CurrentNode.Index = StartFindNodeIndex And FoundFirstNode Then
                GoTo AtStart
            End If
        Wend
WraptoBeginning:
        Set CurrentNode = _
                fMainForm.tvTreeView.Nodes(FindLookinNodeIndex).Child
        While 1
            TempPos = InStr(1, CurrentNode.text, _
                    SearchTxt, CompareType)
            If TempPos > 0 Then
                err = ChkForNodeAttribMatch(CurrentNode, SearchTxt)
                If err = -1 Then
                    GoTo Done
                End If
                If err = 1 Then
                    FoundFirstNode = True
                    GoTo FoundNode
                End If
            End If
            If CurrentNode.Index = StartFindNodeIndex And FoundFirstNode Then
                GoTo AtStart
            End If
            If Not CurrentNode.Child Is Nothing Then
                Set CurrentNode = CurrentNode.Child
            Else
                If Not CurrentNode.Next Is Nothing Then
                    Set CurrentNode = CurrentNode.Next
                Else
                    Set CurrentNode = NextofParent(CurrentNode, FindLookinNodeIndex)
                    If CurrentNode Is Nothing Then
                        GoTo AtStart
                    End If
                End If
            End If
        Wend
        GoTo Done
FoundNode:
        Set fMainForm.tvTreeView.SelectedItem = CurrentNode
        Set nodx = CurrentNode
        nodx.EnsureVisible
        lvNeedsRefresh = True
        fMainForm.tvTreeView_NodeClick nodx
        GoTo Done
AtStart:
        If Not ReplacingAll Then
            MsgBox "The entire node tree was searched.", vbInformation
        End If
        EndOfFind = True
        FoundFirstNode = False
        If nodx.Index = FindLookinNodeIndex Or _
                IsDescendantOf(nodx, _
                fMainForm.tvTreeView.Nodes(FindLookinNodeIndex)) Then
            StartFindNodeIndex = nodx.Index
            StartFindNodeID = RemoveK(nodx.key)
        Else
            ResetFind
        End If
Done:
End Sub

Private Sub FindLV(SearchTxt As String, CompareType As Integer)

Dim listrecord As Recordset
Dim TempItem As ListItem
Dim CurrentNode As Node
Dim tempint As Integer
Dim Link_NodeID As Long
Dim WhereData_IDStr As String

        If StartFindItemIndex = 0 And StartFindNodeIndex = 0 Then
            If nodx.FullPath = cboLookIn.text Then
                StartFindNodeIndex = nodx.Index
                StartFindNodeID = RemoveK(nodx.key)
                FindLookinNodeIndex = nodx.Index
                Set CurrentNode = nodx
            Else
                Set CurrentNode = GetNodeFromFullPath(nodx.Root, cboLookIn.text)
                If CurrentNode Is Nothing Then
                    GoTo Done
                Else
                    FindLookinNodeIndex = CurrentNode.Index
                    If IsDescendantOf(nodx, CurrentNode) Then
                        StartFindNodeIndex = nodx.Index
                        StartFindNodeID = RemoveK(nodx.key)
                    Else
                        StartFindNodeIndex = CurrentNode.Index
                        StartFindNodeID = RemoveK(CurrentNode.key)
                        GoTo NextNode1
                    End If
                End If
            End If
            If CurrentItem Is Nothing Then
                If fMainForm.lvListView.ListItems.count > 0 Then
                    Set CurrentItem = fMainForm.lvListView.ListItems(1)
                Else
                    GoTo NextNode1
                End If
            Else
                StartFindItemIndex = CurrentItem.Index
                StartFindItemID = RemoveK(CurrentItem.key)
            End If
            Set TempItem = CurrentItem
            TempPos = InStr(1, TempItem.text, SearchTxt, CompareType)
            If TempPos > 0 Then
                If WholeWordChk(TempItem.text, SearchTxt) Then
                    If Left$(TempItem.key, 1) = "L" Then
                        Link_NodeID = tvAttribCol(nodx.key).Link_NodeID
                        Set listrecord = dbase.OpenRecordset( _
                            "SELECT Data_ID" _
                            & " FROM " & tvAttribCol(AddK(Link_NodeID)).table_name _
                            & " WHERE Data_ID = " & RemoveK(TempItem.key) & " AND " _
                            & WhereData_ValueStr, _
                            dbOpenDynaset)
                    Else
                        Set listrecord = dbase.OpenRecordset( _
                            "SELECT Data_ID" _
                            & " FROM " & tvAttribCol(nodx.key).table_name _
                            & " WHERE Data_ID = " & RemoveK(TempItem.key) & " AND " _
                            & WhereData_ValueStr, _
                            dbOpenDynaset)
                    End If
                    If Not listrecord.EOF Then
                        TempItem.Selected = False
                        DoEvents
                        GoTo FoundItem1
                    End If
                    listrecord.Close
                End If
            End If
        End If
        Set CurrentNode = nodx
        Set TempItem = CurrentItem
        If LVwrapped And CurrentNode.Index = StartFindNodeIndex Then
            If StartFindItemIndex = 0 Then
                GoTo AtStart
            ElseIf TempItem.Index = StartFindItemIndex Then
                GoTo AtStart
            End If
        End If
        While TempItem.Index < fMainForm.lvListView.ListItems.count
            Set TempItem = fMainForm.lvListView.ListItems(TempItem.Index + 1)
            TempPos = InStr(1, TempItem.text, SearchTxt, CompareType)
            If TempPos > 0 Then
                If WholeWordChk(TempItem.text, SearchTxt) Then
                    LVwrapped = True
                    If Left$(TempItem.key, 1) = "L" Then
                        Link_NodeID = tvAttribCol(CurrentNode.key).Link_NodeID
                        Set listrecord = dbase.OpenRecordset( _
                            "SELECT Data_ID" _
                            & " FROM " & tvAttribCol(AddK(Link_NodeID)).table_name _
                            & " WHERE Data_ID = " & RemoveK(TempItem.key) & " AND " _
                            & WhereData_ValueStr, _
                            dbOpenDynaset)
                    Else
                        Set listrecord = dbase.OpenRecordset( _
                            "SELECT Data_ID" _
                            & " FROM " & tvAttribCol(CurrentNode.key).table_name _
                            & " WHERE Data_ID = " & RemoveK(TempItem.key) & " AND " _
                            & WhereData_ValueStr, _
                            dbOpenDynaset)
                    End If
                    If Not listrecord.EOF Then
                        GoTo FoundItem1
                    End If
                    listrecord.Close
                End If
            End If
            If CurrentNode.Index = StartFindNodeIndex Then
                If StartFindItemIndex = 0 Then
                    GoTo AtStart
                ElseIf TempItem.Index = StartFindItemIndex Then
                    GoTo AtStart
                End If
            End If
        Wend
NextNode1:
        Set listrecord = GotoNextNode(CurrentNode, SearchTxt)
        On Error GoTo WraptoBeginning
        tempint = listrecord.Type
        On Error GoTo 0
LoopThruItems1:
        If CurrentNode.Index = StartFindNodeIndex And _
                StartFindItemIndex = 0 Then
            GoTo AtStart
        End If
        While Not listrecord.EOF
            If CurrentNode.Index = StartFindNodeIndex And _
                    listrecord!data_id = StartFindItemID Then
                GoTo AtStart
            End If
            TempPos = InStr(1, listrecord!data_label, SearchTxt, CompareType)
            If TempPos > 0 Then
'               If WholeWordChk(listrecord!data_label, SearchTxt) Then
                    GoTo FoundItem2
'                End If
            End If
            listrecord.MoveNext
        Wend
        listrecord.Close
        GoTo NextNode1
WraptoBeginning:
        On Error GoTo 0
        Set CurrentNode = fMainForm.tvTreeView.Nodes(FindLookinNodeIndex)
        If RemoveK(CurrentNode.key) = StartFindNodeID Then
            WhereData_IDStr = "Data_Id = " & StartFindItemID
        Else
            WhereData_IDStr = "FALSE"
        End If
        Set listrecord = dbase.OpenRecordset( _
                    "SELECT Data_Label,Data_ID" _
                    & " FROM " & tvAttribCol(CurrentNode.key).table_name _
                    & " WHERE (" & WhereData_IDStr _
                    & " OR (" & WhereData_LabelStr & " AND " _
                    & WhereData_ValueStr _
                    & " )) AND Parent_Node = " & RemoveK(CurrentNode.key) _
                    & " ORDER BY " & OrderByStr, _
                    dbOpenDynaset)
LoopThruItems2:
        If CurrentNode.Index = StartFindNodeIndex And _
                StartFindItemIndex = 0 Then
            GoTo AtStart
        End If
        While Not listrecord.EOF
            If CurrentNode.Index = StartFindNodeIndex And _
                    listrecord!data_id = StartFindItemID Then
                GoTo AtStart
            End If
            TempPos = InStr(1, listrecord!data_label, SearchTxt, CompareType)
            If TempPos > 0 Then
'                If WholeWordChk(listrecord!data_label, SearchTxt) Then
                    GoTo FoundItem2
'                End If
            End If
            listrecord.MoveNext
        Wend
NextNode2:
        listrecord.Close
        Set listrecord = GotoNextNode(CurrentNode, SearchTxt)
        On Error GoTo AtStart
        tempint = listrecord.Type
        On Error GoTo 0
        GoTo LoopThruItems2
FoundItem2:
        Set fMainForm.tvTreeView.SelectedItem = CurrentNode
        Set nodx = CurrentNode
        nodx.EnsureVisible
        lvNeedsRefresh = True
        lvLoaded = False
        fMainForm.tmrNodeClick.Interval = 1
        fMainForm.tvTreeView_NodeClick nodx
        While Not lvLoaded
            DoEvents
        Wend
        fMainForm.tmrNodeClick.Interval = NODE_CLICK_TIMER_DELAY
        Set TempItem = _
            fMainForm.lvListView.FindItem(listrecord!data_label, _
                    lvwText)
        LVwrapped = True
FoundItem1:
        Set fMainForm.lvListView.SelectedItem = TempItem
        Set CurrentItem = TempItem
        CurrentItem.EnsureVisible
        GoTo Done
AtStart:
        On Error GoTo 0
        If Not ReplacingAll Then
            MsgBox "The entire database was searched.", vbInformation
        End If
        EndOfFind = True
        LVwrapped = False
        If nodx.Index = FindLookinNodeIndex Or _
                IsDescendantOf(nodx, _
                fMainForm.tvTreeView.Nodes(FindLookinNodeIndex)) Then
            StartFindNodeIndex = nodx.Index
            StartFindNodeID = RemoveK(nodx.key)
            StartFindItemIndex = CurrentItem.Index
            StartFindItemID = RemoveK(CurrentItem.key)
        Else
            ResetFind
        End If
Done:
        On Error Resume Next
        listrecord.Close
        On Error GoTo 0

End Sub

Private Function GotoNextNode(CurrentNode As Node, SearchTxt As String) As Recordset
    Dim Link_NodeID As Long
    Dim QueryStr As String
    Dim WhereData_IDStr As String
    
Top:
            If Not CurrentNode.Child Is Nothing Then
                Set CurrentNode = CurrentNode.Child
            Else
                If CurrentNode.Index = FindLookinNodeIndex Then
                    GoTo ReachedEnd
                End If
                If Not CurrentNode.Next Is Nothing Then
                    Set CurrentNode = CurrentNode.Next
                Else
                    Set CurrentNode = NextofParent(CurrentNode, FindLookinNodeIndex)
                    If CurrentNode Is Nothing Then
                        GoTo ReachedEnd
                    End If
                End If
            End If
            Link_NodeID = tvAttribCol(CurrentNode.key).Link_NodeID
            If RemoveK(CurrentNode.key) = StartFindNodeID Then
                WhereData_IDStr = "Data_Id = " & StartFindItemID
            Else
                WhereData_IDStr = "FALSE"
            End If
            If cboLinked.text <> "False" And Link_NodeID <> 0 Then
                QueryStr = "SELECT Data_Label,Data_ID" _
                    & " FROM " & tvAttribCol(CurrentNode.key).table_name _
                    & " WHERE (" & WhereData_IDStr _
                    & " OR (" & WhereData_LabelStr & " AND " _
                    & WhereData_ValueStr _
                    & " )) AND Parent_Node = " & RemoveK(CurrentNode.key) _
                    & " UNION SELECT Data_Label,Data_ID FROM " _
                    & tvAttribCol(AddK(Link_NodeID)).table_name _
                    & " WHERE (" & WhereData_IDStr _
                    & " OR (" & WhereData_LabelStr & " AND " _
                    & WhereData_ValueStr _
                    & " )) AND Parent_Node = " & Link_NodeID _
                    & " ORDER BY " & OrderByStr
            Else
                QueryStr = "SELECT Data_Label,Data_ID" _
                    & " FROM " & tvAttribCol(CurrentNode.key).table_name _
                    & " WHERE (" & WhereData_IDStr _
                    & " OR (" & WhereData_LabelStr & " AND " _
                    & WhereData_ValueStr _
                    & " )) AND Parent_Node = " & RemoveK(CurrentNode.key) _
                    & " ORDER BY " & OrderByStr
            End If
            Set GotoNextNode = dbase.OpenRecordset(QueryStr, dbOpenDynaset)
            If CurrentNode.Index <> StartFindNodeIndex And _
                    GotoNextNode.EOF Then
                GotoNextNode.Close
                GoTo Top
            End If
            GoTo Done
ReachedEnd:
    On Error Resume Next
    GotoNextNode.Close
    On Error GoTo 0
Done:
End Function

Private Sub cmdReplace_Click()
    Dim LenSearchTxt As Long
    Dim SearchTxt As String
    Dim ReplaceTxt As String
    Dim tempstr As String
    
    UpdateComboItems
    
    LenSearchTxt = Len(cboFind.text)
    SearchTxt = cboFind.text
    ReplaceTxt = cboReplace.text
    
    If TypeOf fMainForm.ActiveControl Is TreeView Then
        tempstr = Left$(nodx.text, TempPos - 1) & _
              ReplaceTxt & _
              Right$(nodx.text, Len(nodx.text) - LenSearchTxt - TempPos + 1)
        fMainForm.tvTreeView_AfterLabelEdit False, tempstr
    ElseIf TypeOf fMainForm.ActiveControl Is ListView Then
        tempstr = Left$(CurrentItem.text, TempPos - 1) & _
              ReplaceTxt & _
              Right$(CurrentItem.text, Len(CurrentItem.text) - LenSearchTxt - TempPos + 1)
        fMainForm.lvListView_AfterLabelEdit False, tempstr
    Else
        MsgBox "You cannot do a find/replace from this control.", vbExclamation
    End If

    cmdFindNext_Click
End Sub

Private Sub cmdReplaceAll_Click()
    Dim count As Long
    
    EndOfFind = False
    ReplacingAll = True
    ResetFind
    count = 0
    
    cmdFindNext_Click
    While Not EndOfFind
        cmdReplace_Click
        count = count + 1
    Wend
    MsgBox format$(count) & " references were replaced.", vbInformation
    ReplacingAll = False
End Sub

Private Sub cboFind_Change()
    ResetFind
End Sub

Private Sub cboFind_GotFocus()
    cboFind.SelStart = 0
    cboFind.SelLength = Len(cboFind.text)
End Sub

Private Sub cboReplace_Change()
    cmdReplace.Enabled = True
    cmdReplaceAll.Enabled = True
End Sub

Private Sub cboReplace_GotFocus()
    cboReplace.SelStart = 0
    cboReplace.SelLength = Len(cboFind.text)
End Sub

Private Sub UpdateComboItems()
    Dim i As Long
    Dim Match As Boolean
    
    Match = False
    For i = 0 To cboFind.ListCount - 1
        If cboFind.List(i) = cboFind.text Then
            Match = True
            Exit For
        End If
    Next i
    
    If Not Match Then
        If cboFind.ListCount >= MAX_COMBO_ITEMS Then
            cboFind.RemoveItem 0
        End If
        cboFind.AddItem cboFind.text
    End If
    
    If cboReplace.Visible Then
        Match = False
        For i = 0 To cboReplace.ListCount - 1
            If cboReplace.List(i) = cboReplace.text Then
                Match = True
                Exit For
            End If
        Next i
    End If
    
    If Not Match Then
        If cboReplace.ListCount >= MAX_COMBO_ITEMS Then
            cboReplace.RemoveItem 0
        End If
        cboReplace.AddItem cboReplace.text
    End If
End Sub

Public Sub InitReplaceForm()
    Dim TempNode As Node
    Dim Match As Boolean
    
    EndOfFind = False
    ReplacingAll = False
    
    If FindEnabled Then
        lblReplace.Visible = False
        cboReplace.Visible = False
        cmdReplace.Visible = False
        cmdReplaceAll.Visible = False
        Me.Caption = "Find"
        cboFind.text = LastFind
        cbMatchCase.Value = LastCase
        cbWhole.Value = LastWhole
    Else
        lblReplace.Visible = True
        cboReplace.Visible = True
        cmdReplace.Visible = True
        cmdReplaceAll.Visible = True
        Me.Caption = "Replace"
        cboFind.text = LastFind
        cbMatchCase.Value = LastCase
        cbWhole.Value = LastWhole
        cboReplace.text = LastReplace
    End If
    
    cboLookIn.Clear
    Set TempNode = nodx.Root
    cboLookIn.AddItem TempNode.FullPath
    Set TempNode = TempNode.Child
    Match = False
    While TempNode.Index <> TempNode.LastSibling.Index
        cboLookIn.AddItem TempNode.FullPath
        If TempNode.FullPath = nodx.FullPath Then
            Match = True
        End If
        Set TempNode = TempNode.Next
    Wend
    cboLookIn.AddItem TempNode.FullPath
    If Not Match Then
        cboLookIn.AddItem nodx.FullPath
    End If
    cboLookIn.text = nodx.FullPath
    
    cboFind.SelStart = 0
    cboFind.SelLength = Len(cboFind.text)
    If cboReplace.text = "" Then
        cmdReplace.Enabled = False
        cmdReplaceAll.Enabled = False
    Else
        cmdReplace.Enabled = True
        cmdReplaceAll.Enabled = True
    End If
    ResetFind
    If cboReadOnly.ListCount = 0 Then
        cboReadOnly.AddItem "", 0
        cboReadOnly.AddItem "True", 1
        cboReadOnly.AddItem "False", 2
        cboSystem.AddItem "", 0
        cboSystem.AddItem "True", 1
        cboSystem.AddItem "False", 2
        cboCreateNode.AddItem "", 0
        cboCreateNode.AddItem "True", 1
        cboCreateNode.AddItem "False", 2
        cboCreateItem.AddItem "", 0
        cboCreateItem.AddItem "True", 1
        cboCreateItem.AddItem "False", 2
        cboLinked.AddItem "", 0
        cboLinked.AddItem "True", 1
        cboLinked.AddItem "False", 2
    End If
    cboFind.SetFocus

End Sub

Private Sub ResetFind()
    FoundFirstNode = False
    LVwrapped = False
    StartFindNodeIndex = 0
    StartFindNodeID = 0
    StartFindItemIndex = 0
    StartFindItemID = 0
End Sub

Private Sub tboCntTxt_Change()
    ResetFind
End Sub

Private Function ChkForNodeAttribMatch(CurrentObj As Object, _
        SearchTxt As String) As Integer
    
    ChkForNodeAttribMatch = WholeWordChk(CurrentObj.text, SearchTxt)
    If ChkForNodeAttribMatch < 1 Then
        Exit Function
    End If
        
    ChkForNodeAttribMatch = CreatedDateChk(CurrentObj.key)
    If ChkForNodeAttribMatch < 1 Then
        Exit Function
    End If
    
    ChkForNodeAttribMatch = CreatedTimeChk(CurrentObj.key)
    If ChkForNodeAttribMatch < 1 Then
        Exit Function
    End If
    
    ChkForNodeAttribMatch = CreatedByChk(CurrentObj.key)
    If ChkForNodeAttribMatch < 1 Then
        Exit Function
    End If
    
    ChkForNodeAttribMatch = LMDateChk(CurrentObj.key)
    If ChkForNodeAttribMatch < 1 Then
        Exit Function
    End If
    
    ChkForNodeAttribMatch = LMTimeChk(CurrentObj.key)
    If ChkForNodeAttribMatch < 1 Then
        Exit Function
    End If
    
    ChkForNodeAttribMatch = LastModifiedByChk(CurrentObj.key)
    If ChkForNodeAttribMatch < 1 Then
        Exit Function
    End If
    
    ChkForNodeAttribMatch = TableNameChk(CurrentObj.key)
    If ChkForNodeAttribMatch < 1 Then
        Exit Function
    End If
    
    ChkForNodeAttribMatch = ReadOnlyChk(CurrentObj.key)
    If ChkForNodeAttribMatch < 1 Then
        Exit Function
    End If
    
    ChkForNodeAttribMatch = SystemChk(CurrentObj.key)
    If ChkForNodeAttribMatch < 1 Then
        Exit Function
    End If
    
    ChkForNodeAttribMatch = CreateNodeChk(CurrentObj.key)
    If ChkForNodeAttribMatch < 1 Then
        Exit Function
    End If
    
    ChkForNodeAttribMatch = CreateItemChk(CurrentObj.key)
    If ChkForNodeAttribMatch < 1 Then
        Exit Function
    End If
    
    ChkForNodeAttribMatch = QuickTypeChk(CurrentObj.key)
    If ChkForNodeAttribMatch < 1 Then
        Exit Function
    End If
    
    ChkForNodeAttribMatch = LinkedChk(CurrentObj.key)
    If ChkForNodeAttribMatch < 1 Then
        Exit Function
    End If
    
    'if got this far then, match; so just return
    
End Function

Private Function WholeWordChk(CurrentObjTxt As String, SearchTxt As String) _
            As Integer

                If cbWhole.Value = 1 Then
chkLeftChar:
                    If TempPos - 1 > 0 Then
                        If mID$(CurrentObjTxt, TempPos - 1, 1) = " " Or _
                           mID$(CurrentObjTxt, TempPos - 1, 1) = "," Or _
                           mID$(CurrentObjTxt, TempPos - 1, 1) = "\" Or _
                           mID$(CurrentObjTxt, TempPos - 1, 1) = "/" Or _
                           mID$(CurrentObjTxt, TempPos - 1, 1) = "(" _
                        Then
                            GoTo chkRightChar
                        Else
                            GoTo NotFound
                        End If
                    Else
                        GoTo chkRightChar
                    End If
chkRightChar:
                    If TempPos + Len(SearchTxt) <= Len(CurrentObjTxt) Then
                        If mID$(CurrentObjTxt, TempPos + Len(SearchTxt), 1) = " " Or _
                           mID$(CurrentObjTxt, TempPos + Len(SearchTxt), 1) = "," Or _
                           mID$(CurrentObjTxt, TempPos + Len(SearchTxt), 1) = "/" Or _
                           mID$(CurrentObjTxt, TempPos + Len(SearchTxt), 1) = "\" Or _
                           mID$(CurrentObjTxt, TempPos + Len(SearchTxt), 1) = ")" _
                        Then
                            GoTo FoundNode
                        End If
                    Else
                        GoTo FoundNode
                    End If
                Else
                    GoTo FoundNode
                End If
NotFound:
            WholeWordChk = 0
            Exit Function
FoundNode:
            WholeWordChk = 1
            Exit Function
HaltError:
            WholeWordChk = -1
            
End Function

Private Function CreatedDateChk(CurrentObjKey As String) As Integer

Dim SearchCreatedDateStr As String
Dim SearchCreatedDate As Variant
Dim NodeCreatedDate As Variant

    SearchCreatedDateStr = cboCreatedDate.text
    If Len(SearchCreatedDateStr) = 0 Then
        GoTo FoundNode
    End If
    If IsDate(SearchCreatedDateStr) Then
        SearchCreatedDate = DateValue(SearchCreatedDateStr)
        NodeCreatedDate = DateValue(tvAttribCol(CurrentObjKey).Created)
        If SearchCreatedDate = NodeCreatedDate Then
            GoTo FoundNode
        Else
            GoTo NotFound
        End If
    Else
        MsgBox "The 'Last Modified Date' search value is not a valid date." & Chr(13) & _
                "Please re-enter.", vbInformation, "Notice"
        GoTo HaltError
    End If

NotFound:
        CreatedDateChk = 0
        Exit Function
FoundNode:
        CreatedDateChk = 1
        Exit Function
HaltError:
        CreatedDateChk = -1

End Function

Private Function CreatedTimeChk(CurrentObjKey As String) As Integer

Dim SearchCreatedTimeStr As String
Dim SearchCreatedTime As Variant
Dim NodeCreatedTime As Variant

    SearchCreatedTimeStr = cboCreatedTime.text
    If Len(SearchCreatedTimeStr) = 0 Then
        GoTo FoundNode
    End If
    If IsDate(SearchCreatedTimeStr) Then
        SearchCreatedTime = TimeValue(SearchCreatedTimeStr)
        NodeCreatedTime = TimeValue(tvAttribCol(CurrentObjKey).Created)
        If SearchCreatedTime = NodeCreatedTime Then
            GoTo FoundNode
        Else
            GoTo NotFound
        End If
    Else
        MsgBox "The 'Last Modified Time' search value is not a valid time." & Chr(13) & _
                "Please re-enter.", vbInformation, "Notice"
        GoTo HaltError
    End If

NotFound:
        CreatedTimeChk = 0
        Exit Function
FoundNode:
        CreatedTimeChk = 1
        Exit Function
HaltError:
        CreatedTimeChk = -1

End Function

Private Function CreatedByChk(CurrentObjKey As String) As Integer

Dim CreatedByStr As String

    CreatedByStr = cboCreatedBy.text
    If Len(CreatedByStr) = 0 Then
        GoTo FoundNode
    End If
    If InStr(1, tvAttribCol(CurrentObjKey).created_by, CreatedByStr, _
            vbTextCompare) > 0 Then
        GoTo FoundNode
    Else
        GoTo NotFound
    End If
    
NotFound:
        CreatedByChk = 0
        Exit Function
FoundNode:
        CreatedByChk = 1
        Exit Function
HaltError:
        CreatedByChk = -1

End Function

Private Function LMDateChk(CurrentObjKey As String) As Integer

Dim SearchLMDateStr As String
Dim SearchLMDate As Variant
Dim NodeLMDate As Variant

    SearchLMDateStr = cboLMDate.text
    If Len(SearchLMDateStr) = 0 Then
        GoTo FoundNode
    End If
    If IsDate(SearchLMDateStr) Then
        SearchLMDate = DateValue(SearchLMDateStr)
        NodeLMDate = DateValue(tvAttribCol(CurrentObjKey).last_modified)
        If SearchLMDate = NodeLMDate Then
            GoTo FoundNode
        Else
            GoTo NotFound
        End If
    Else
        MsgBox "The 'Last Modified Date' search value is not a valid date." & Chr(13) & _
                "Please re-enter.", vbInformation, "Notice"
        GoTo HaltError
    End If

NotFound:
        LMDateChk = 0
        Exit Function
FoundNode:
        LMDateChk = 1
        Exit Function
HaltError:
        LMDateChk = -1

End Function

Private Function LMTimeChk(CurrentObjKey As String) As Integer

Dim SearchLMTimeStr As String
Dim SearchLMTime As Variant
Dim NodeLMTime As Variant

    SearchLMTimeStr = cboLMTime.text
    If Len(SearchLMTimeStr) = 0 Then
        GoTo FoundNode
    End If
    If IsDate(SearchLMTimeStr) Then
        SearchLMTime = TimeValue(SearchLMTimeStr)
        NodeLMTime = TimeValue(tvAttribCol(CurrentObjKey).last_modified)
        If SearchLMTime = NodeLMTime Then
            GoTo FoundNode
        Else
            GoTo NotFound
        End If
    Else
        MsgBox "The 'Last Modified Time' search value is not a valid time." & Chr(13) & _
                "Please re-enter.", vbInformation, "Notice"
        GoTo HaltError
    End If

NotFound:
        LMTimeChk = 0
        Exit Function
FoundNode:
        LMTimeChk = 1
        Exit Function
HaltError:
        LMTimeChk = -1

End Function

Private Function LastModifiedByChk(CurrentObjKey As String) As Integer

Dim LastModifiedByStr As String

    LastModifiedByStr = cboModifiedBy.text
    If Len(LastModifiedByStr) = 0 Then
        GoTo FoundNode
    End If
    If InStr(1, tvAttribCol(CurrentObjKey).modified_by, LastModifiedByStr, _
            vbTextCompare) > 0 Then
        GoTo FoundNode
    Else
        GoTo NotFound
    End If
    
NotFound:
        LastModifiedByChk = 0
        Exit Function
FoundNode:
        LastModifiedByChk = 1
        Exit Function
HaltError:
        LastModifiedByChk = -1

End Function

Private Function TableNameChk(CurrentObjKey As String) As Integer

Dim TableNameStr As String

    TableNameStr = cboTableName.text
    If Len(TableNameStr) = 0 Then
        GoTo FoundNode
    End If
    If InStr(1, tvAttribCol(CurrentObjKey).table_name, TableNameStr, _
            vbTextCompare) > 0 Then
        GoTo FoundNode
    Else
        GoTo NotFound
    End If
    
NotFound:
        TableNameChk = 0
        Exit Function
FoundNode:
        TableNameChk = 1
        Exit Function
HaltError:
        TableNameChk = -1

End Function

Private Function ReadOnlyChk(CurrentObjKey As String) As Integer

Dim ReadOnlyStr As String
Dim ReadOnly As Boolean

    ReadOnlyStr = cboReadOnly.text
    If Len(ReadOnlyStr) = 0 Then
        GoTo FoundNode
    End If
    If ReadOnlyStr = "True" Then
        ReadOnly = True
    Else
        ReadOnly = False
    End If
    If ReadOnly = tvAttribCol(CurrentObjKey).read_only Then
        GoTo FoundNode
    Else
        GoTo NotFound
    End If
    
NotFound:
        ReadOnlyChk = 0
        Exit Function
FoundNode:
        ReadOnlyChk = 1
        Exit Function
HaltError:
        ReadOnlyChk = -1

End Function

Private Function SystemChk(CurrentObjKey As String) As Integer

Dim SystemStr As String
Dim System As Boolean

    SystemStr = cboSystem.text
    If Len(SystemStr) = 0 Then
        GoTo FoundNode
    End If
    If SystemStr = "True" Then
        System = True
    Else
        System = False
    End If
    If System = tvAttribCol(CurrentObjKey).system_node Then
        GoTo FoundNode
    Else
        GoTo NotFound
    End If
    
NotFound:
        SystemChk = 0
        Exit Function
FoundNode:
        SystemChk = 1
        Exit Function
HaltError:
        SystemChk = -1

End Function

Private Function CreateNodeChk(CurrentObjKey As String) As Integer

Dim CreateNodeStr As String
Dim CreateNode As Boolean

    CreateNodeStr = cboCreateNode.text
    If Len(CreateNodeStr) = 0 Then
        GoTo FoundNode
    End If
    If CreateNodeStr = "True" Then
        CreateNode = True
    Else
        CreateNode = False
    End If
    If CreateNode = tvAttribCol(CurrentObjKey).create_node Then
        GoTo FoundNode
    Else
        GoTo NotFound
    End If
    
NotFound:
        CreateNodeChk = 0
        Exit Function
FoundNode:
        CreateNodeChk = 1
        Exit Function
HaltError:
        CreateNodeChk = -1

End Function

Private Function CreateItemChk(CurrentObjKey As String) As Integer

Dim CreateItemStr As String
Dim CreateItem As Boolean

    CreateItemStr = cboCreateItem.text
    If Len(CreateItemStr) = 0 Then
        GoTo FoundNode
    End If
    If CreateItemStr = "True" Then
        CreateItem = True
    Else
        CreateItem = False
    End If
    If CreateItem = tvAttribCol(CurrentObjKey).create_item Then
        GoTo FoundNode
    Else
        GoTo NotFound
    End If
    
NotFound:
        CreateItemChk = 0
        Exit Function
FoundNode:
        CreateItemChk = 1
        Exit Function
HaltError:
        CreateItemChk = -1

End Function

Private Function QuickTypeChk(CurrentObjKey As String) As Integer

Dim SearchQuickTypeStr As String
Dim SearchQuickType As Long
Dim NodeQuickType As Long

    SearchQuickTypeStr = cboQuickType.text
    If Len(SearchQuickTypeStr) = 0 Then
        GoTo FoundNode
    End If
    If IsNumeric(SearchQuickTypeStr) Then
        SearchQuickType = val(SearchQuickTypeStr)
        NodeQuickType = tvAttribCol(CurrentObjKey).quicktypeid
        If SearchQuickType = NodeQuickType Then
            GoTo FoundNode
        Else
            GoTo NotFound
        End If
    Else
        MsgBox "The 'Quick Type' search value is not a valid number." & Chr(13) & _
                "Please re-enter.", vbInformation, "Notice"
        GoTo HaltError
    End If

NotFound:
        QuickTypeChk = 0
        Exit Function
FoundNode:
        QuickTypeChk = 1
        Exit Function
HaltError:
        QuickTypeChk = -1

End Function

Private Function LinkedChk(CurrentObjKey As String) As Integer

Dim LinkedStr As String
Dim Linked As Boolean
Dim NodeLinked As Boolean

    LinkedStr = cboLinked.text
    If Len(LinkedStr) = 0 Then
        GoTo FoundNode
    End If
    If LinkedStr = "True" Then
        Linked = True
    Else
        Linked = False
    End If
    If tvAttribCol(CurrentObjKey).Link_NodeID <> 0 Then
        NodeLinked = True
    Else
        NodeLinked = False
    End If
    If Linked = NodeLinked Then
        If Linked = True Then
            If LinkNodeChk(CurrentObjKey) Then  'chk for match of specific link node
                GoTo FoundNode
            Else
                GoTo NotFound
            End If
        Else
            GoTo FoundNode
        End If
    Else
        GoTo NotFound
    End If
    
NotFound:
        LinkedChk = 0
        Exit Function
FoundNode:
        LinkedChk = 1
        Exit Function
HaltError:
        LinkedChk = -1

End Function

Private Function LinkNodeChk(CurrentObjKey As String) As Integer

Dim SearchLinkNodeStr As String
Dim SearchLinkNode As Long
Dim NodeLinkNode As Long

    SearchLinkNodeStr = cboLinkNode.text
    If Len(SearchLinkNodeStr) = 0 Then
        GoTo FoundNode
    End If
    If IsNumeric(SearchLinkNodeStr) Then
        SearchLinkNode = val(SearchLinkNodeStr)
        NodeLinkNode = tvAttribCol(CurrentObjKey).Link_NodeID
        If SearchLinkNode = NodeLinkNode Then
            GoTo FoundNode
        Else
            GoTo NotFound
        End If
    Else
        MsgBox "The 'Link Node' search value is not a valid number." & Chr(13) & _
                "Please re-enter.", vbInformation, "Notice"
        GoTo HaltError
    End If

NotFound:
        LinkNodeChk = 0
        Exit Function
FoundNode:
        LinkNodeChk = 1
        Exit Function
HaltError:
        LinkNodeChk = -1

End Function

