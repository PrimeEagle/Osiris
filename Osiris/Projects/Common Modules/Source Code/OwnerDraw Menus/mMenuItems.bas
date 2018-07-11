Attribute VB_Name = "mMenuItems"
Option Explicit
DefLng A-Z

Const NUM_MENU_ITEMS = 25
Global Caps(NUM_MENU_ITEMS) As String
Global HasIcon(NUM_MENU_ITEMS) As Integer

Dim hMenu As Long
Dim hSubMenu As Long
Dim mnuID As Long
Dim m_Form As Form

Public Property Get MenuForm() As Form
    Set MenuForm = m_Form
End Property

Public Property Let MenuForm(ByVal vNewValue As Form)
    Set m_Form = vNewValue
    hMenu = GetMenu(m_Form.hwnd)
End Property

Public Property Get subMenu() As Long
    subMenu = hSubMenu
End Property

Public Property Let subMenu(ByVal vNewValue As Long)
    hSubMenu = GetSubMenu(hMenu, vNewValue)
End Property

Public Property Get menuId() As Long
    menuId = mnuID
End Property

Public Property Let menuId(ByVal vNewValue As Long)
    mnuID = GetMenuItemID(hSubMenu, vNewValue)
End Property

Public Sub OwnerDrawMenu(ByVal ItemData As Long)
    'Change the menu's style to owner-draw. You must
    'now subclass the form that this menu is on so
    'you can respond to the WM_MEASUREITEM and WM_DRAWITEM
    'messages.
    Call ModifyMenu(hSubMenu, menuId, MF_BYCOMMAND Or MF_OWNERDRAW, menuId, ItemData)
End Sub

Public Function GetString(ByVal mID As Long) As String
    Dim tmpStr As String * 100
    
    Call GetMenuString(hSubMenu, mID, tmpStr, 100, 0&)
    GetString = VBA.Trim$(tmpStr)
End Function

Public Sub OwnToplevel(ByVal ItemData As Long)
    Call ModifyMenu(hMenu, menuId, MF_BYCOMMAND Or MF_OWNERDRAW, menuId, ItemData)
End Sub
