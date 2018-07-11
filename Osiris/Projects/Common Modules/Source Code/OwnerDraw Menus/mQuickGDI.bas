Attribute VB_Name = "mQuickGDI"
'This module requires the following module(s) to exist in the project:
'   mWin32API.bas

Option Explicit
DefLng A-Z

Const NEWTRANSPARENT = 3  '  use with SetBkMode()

Dim m_hDC As Long
Dim CurObj As Long

Public Property Get SysColor(ByVal Index As ColConst) As Long
    SysColor = GetSysColor(Index)
End Property

Public Property Let SysColor(ByVal Index As ColConst, ByVal NewCol As Long)
    Call SetSysColors(1, ByVal Index, NewCol)
End Property

Public Property Get TargethDC() As Long
    TargethDC = m_hDC
End Property

Public Property Let TargethDC(ByVal vNewValue As Long)
    'The hDC to draw to when performing operations
    'from this module's subroutines.
    m_hDC = vNewValue
End Property

Public Sub DrawRect(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
    If m_hDC = 0 Then Exit Sub
    Call Rectangle(m_hDC, X1, Y1, X2, Y2)
End Sub

Public Function GetPen(ByVal nWidth As Long, ByVal Clr As Long) As Long
    GetPen = CreatePen(0, nWidth, Clr)
End Function

Public Sub ThreedBox(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
    'Draw a raised box around the specified
    'coordinates.

    If m_hDC = 0 Then Exit Sub

    Dim CurPen As Long, OldPen As Long
    Dim dm As POINTAPI

    CurPen = GetPen(1, SysColor(COLOR_BTNHIGHLIGHT))
    OldPen = SelectObject(m_hDC, CurPen)
    'FirstLightLine
    MoveToEx m_hDC, X1, Y2, dm
    LineTo m_hDC, X1, Y1
    'SecondLightLine
    LineTo m_hDC, X2, Y1

    SelectObject m_hDC, OldPen
    DeleteObject CurPen
    CurPen = GetPen(1, SysColor(COLOR_BTNSHADOW))
    OldPen = SelectObject(m_hDC, CurPen)
    'FirstDarkLine
    MoveToEx m_hDC, X2, Y1, dm
    LineTo m_hDC, X2, Y2
    'SecondDarkLine
    LineTo m_hDC, X1, Y2

    SelectObject m_hDC, OldPen
    DeleteObject CurPen
End Sub
Public Sub ThreedBoxInvert(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
    'Draw a raised box around the specified
    'coordinates.
    If m_hDC = 0 Then Exit Sub
    
    Dim CurPen As Long, OldPen As Long
    Dim dm As POINTAPI

    CurPen = GetPen(1, SysColor(COLOR_BTNSHADOW))
    OldPen = SelectObject(m_hDC, CurPen)
    'FirstLightLine
    MoveToEx m_hDC, X1, Y1, dm
    LineTo m_hDC, X2, Y1

    SelectObject m_hDC, OldPen
    DeleteObject CurPen
    CurPen = GetPen(1, SysColor(COLOR_BTNHIGHLIGHT))
    OldPen = SelectObject(m_hDC, CurPen)
    'FirstDarkLine
    MoveToEx m_hDC, X1, Y2, dm
    LineTo m_hDC, X2, Y2

    SelectObject m_hDC, OldPen
    DeleteObject CurPen
End Sub

Public Function hPrint(ByVal X As Long, ByVal Y As Long, ByVal hStr As String, ByVal Clr As Long) As Long
    If m_hDC = 0 Then Exit Function
    
    'Equivalent to setting a form's property
    'FontTransparent = True
    SetBkMode m_hDC, NEWTRANSPARENT
        
    Dim OT As Long
    
    OT = GetTextColor(m_hDC)
    SetTextColor m_hDC, Clr
    'Print the text
    hPrint = TextOut(m_hDC, X, Y, hStr, Len(hStr))
    'Restore old text color
    SetTextColor m_hDC, OT
End Function

Public Function Stock(ByVal Obj As StockObjects) As Long
    Stock = GetStockObject(Obj)
End Function
