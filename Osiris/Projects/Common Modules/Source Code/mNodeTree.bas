Attribute VB_Name = "mNodeTree"
'This module requires the following components to exist in the project:
'   Microsoft Windows Common Controls (tested with versions 5.0 and 6.0)

Option Explicit

Public Function IsDescendantOf(SearchNode As Node, Ancestor As Node) _
        As Boolean
    Dim CurrentNode As Node
    
    Set CurrentNode = SearchNode
    While CurrentNode.Index <> Ancestor.Index
        Set CurrentNode = CurrentNode.Parent
        If CurrentNode Is Nothing Then
            IsDescendantOf = False
            Exit Function
        End If
    Wend
    IsDescendantOf = True
   
End Function

Public Function NextofParent(n As Node, LookinNodeIndex As Long) As Node
    
    If n.Parent.Index = LookinNodeIndex Then
        Set NextofParent = Nothing
        Exit Function
    End If
    If Not n.Parent.Next Is Nothing Then
        Set NextofParent = n.Parent.Next
    Else
        Set NextofParent = NextofParent(n.Parent, LookinNodeIndex)
    End If

End Function

Public Function GeneralNextofParent(n As Node) As Node
    ' if n has no parent (root node), return 'Nothing'.
    ' The recursion will also stop on this one if the original
    ' call to the function was with the last node in the tree,
    ' because it will recurse upward until it reaches the root.
    If n.Parent Is Nothing Then
        Set GeneralNextofParent = Nothing
        Exit Function
    End If
    ' if n has a parent, AND there exists a next sibling
    ' for that parent, then return the next one
    ' after the parent
    If Not n.Parent.Next Is Nothing Then
        Set GeneralNextofParent = n.Parent.Next
    Else
        'if n has a parent, but there is no next node after it,
        'then recurse the tree upwards until you find one,
        'or end up returning 'Nothing' due to being at the
        'end of the node tree.
        Set GeneralNextofParent = GeneralNextofParent(n.Parent)
    End If
End Function


