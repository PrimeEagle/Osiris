Attribute VB_Name = "Module1"
Public Const MF_BITMAP = &H4
                    Public Const CLR_MENUBAR = &H80000004

                    Const Number_of_Menu_Selections = 3
                    'changes depending on the number of menu items

                    Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) _
                    As Long
                    Declare Function GetSubMenu Lib "user32" (ByVal hwnd As Long, ByVal x As Long) _
                    As Long
                    Declare Function GetMenuItemID Lib "user32" (ByVal _
                    hMenu As Long, ByVal nPos As Long) As Long
                    Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu _
                    As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem _
                    As Long, ByVal lpString As String) As Long
                    Declare Function SetMenuItemBitmaps Lib "user32" _
                    (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal _
                    hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long

