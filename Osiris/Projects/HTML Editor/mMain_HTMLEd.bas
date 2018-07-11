Attribute VB_Name = "mMain_HTMLEd"
Option Explicit

Public fHTMLEd As frmHTMLEdit

Sub Main()
    Set fHTMLEd = New frmHTMLEdit
    Load fHTMLEd
End Sub

Public Sub BuildimlMenu(Optional UseProgressBar As Boolean = False)
    
    Dim LargeIconFolder As String
    Dim PBICON_BuildMenu As String
    Dim MENU_ICON_PATH As String
    
    LargeIconFolder = "c:\osiris\resources\icons\32"
    MENU_ICON_PATH = "c:\osiris\resources\icons\menu"
    
    MsgBox "TEMP:  LargeIconFolder,PBICON_BuildMenu,MENU_ICON_PATH should be stored in the registry!", vbInformation
    
    If UseProgressBar Then
        InitProgressBar "Building Image List for menus . . .", 0, 100, _
            LargeIconFolder & PBICON_BuildMenu, False
    End If
    
    On Error Resume Next    'if the pictures do not exist, then blank pictures
    imlmenu.ListImages.Add , "Copy", LoadPicture(MENU_ICON_PATH & "Edit Copy 16.bmp")
    imlmenu.ListImages.Add , "Cut", LoadPicture(MENU_ICON_PATH & "Edit Cut 16.bmp")
    imlmenu.ListImages.Add , "Paste", LoadPicture(MENU_ICON_PATH & "Edit Paste 16.bmp")
    imlmenu.ListImages.Add , "New", LoadPicture(MENU_ICON_PATH & "File New 16.bmp")
    imlmenu.ListImages.Add , "Open", LoadPicture(MENU_ICON_PATH & "File Open 16.bmp")
    imlmenu.ListImages.Add , "Print", LoadPicture(MENU_ICON_PATH & "File Print 16.bmp")
    imlmenu.ListImages.Add , "Preview", LoadPicture(MENU_ICON_PATH & "File Print Preview 16.bmp")
    imlmenu.ListImages.Add , "Save", LoadPicture(MENU_ICON_PATH & "File Save 16.bmp")
    imlmenu.ListImages.Add , "Paint", LoadPicture(MENU_ICON_PATH & "Format Color 16.bmp")
    imlmenu.ListImages.Add , "Help", LoadPicture(MENU_ICON_PATH & "Help 16.bmp")
    imlmenu.ListImages.Add , "Delete", LoadPicture(MENU_ICON_PATH & "Edit Delete 16.bmp")
    imlmenu.ListImages.Add , "Prop", LoadPicture(MENU_ICON_PATH & "Properties 16.bmp")
    imlmenu.ListImages.Add , "Security", LoadPicture(MENU_ICON_PATH & "Security 16.bmp")
    imlmenu.ListImages.Add , "AccessTable", LoadPicture(MENU_ICON_PATH & "Access Table 16.bmp")
    imlmenu.ListImages.Add , "Bold", LoadPicture(MENU_ICON_PATH & "Bold 16.bmp")
    imlmenu.ListImages.Add , "Italic", LoadPicture(MENU_ICON_PATH & "Italic 16.bmp")
    imlmenu.ListImages.Add , "Underline", LoadPicture(MENU_ICON_PATH & "Underline 16.bmp")
    imlmenu.ListImages.Add , "Right Just", LoadPicture(MENU_ICON_PATH & "Right Just 16.bmp")
    imlmenu.ListImages.Add , "Left Just", LoadPicture(MENU_ICON_PATH & "Left Just 16.bmp")
    imlmenu.ListImages.Add , "Center Just", LoadPicture(MENU_ICON_PATH & "Center Just 16.bmp")
    imlmenu.ListImages.Add , "Bullets", LoadPicture(MENU_ICON_PATH & "Bullets 16.bmp")
    imlmenu.ListImages.Add , "Font Color", LoadPicture(MENU_ICON_PATH & "Font Color 16.bmp")
    imlmenu.ListImages.Add , "Absolute Mode", LoadPicture(MENU_ICON_PATH & "Absolute Mode 16.bmp")
    imlmenu.ListImages.Add , "Absolute Pos", LoadPicture(MENU_ICON_PATH & "Absolute Pos.bmp")
    imlmenu.ListImages.Add , "BGColor", LoadPicture(MENU_ICON_PATH & "BGColor.bmp")
    imlmenu.ListImages.Add , "FGColor", LoadPicture(MENU_ICON_PATH & "FGColor.bmp")
    imlmenu.ListImages.Add , "Borders", LoadPicture(MENU_ICON_PATH & "Borders.bmp")
    imlmenu.ListImages.Add , "Delete Cell", LoadPicture(MENU_ICON_PATH & "Delete Cell.bmp")
    imlmenu.ListImages.Add , "Delete Column", LoadPicture(MENU_ICON_PATH & "Delete Column.bmp")
    imlmenu.ListImages.Add , "Delete Row", LoadPicture(MENU_ICON_PATH & "Delete Row.bmp")
    imlmenu.ListImages.Add , "Decrease Indent", LoadPicture(MENU_ICON_PATH & "Decrease Indent.bmp")
    imlmenu.ListImages.Add , "Details", LoadPicture(MENU_ICON_PATH & "Details.bmp")
    imlmenu.ListImages.Add , "Find", LoadPicture(MENU_ICON_PATH & "Find.bmp")
    imlmenu.ListImages.Add , "Image", LoadPicture(MENU_ICON_PATH & "Image.bmp")
    imlmenu.ListImages.Add , "Increase Indent", LoadPicture(MENU_ICON_PATH & "Increase Indent.bmp")
    imlmenu.ListImages.Add , "Insert Cell", LoadPicture(MENU_ICON_PATH & "Insert Cell.bmp")
    imlmenu.ListImages.Add , "Insert Column", LoadPicture(MENU_ICON_PATH & "Insert Column.bmp")
    imlmenu.ListImages.Add , "Insert Row", LoadPicture(MENU_ICON_PATH & "Insert Row.bmp")
    imlmenu.ListImages.Add , "Insert Table", LoadPicture(MENU_ICON_PATH & "Insert Table.bmp")
    imlmenu.ListImages.Add , "Link", LoadPicture(MENU_ICON_PATH & "Link.bmp")
    imlmenu.ListImages.Add , "Merge Cell", LoadPicture(MENU_ICON_PATH & "Merge Cell.bmp")
    imlmenu.ListImages.Add , "Redo", LoadPicture(MENU_ICON_PATH & "Redo.bmp")
    imlmenu.ListImages.Add , "Snap Grid 16", LoadPicture(MENU_ICON_PATH & "Snap Grid 16.bmp")
    imlmenu.ListImages.Add , "Split Cell 16", LoadPicture(MENU_ICON_PATH & "Split Cell 16.bmp")
    imlmenu.ListImages.Add , "Undo", LoadPicture(MENU_ICON_PATH & "Undo.bmp")
    imlmenu.ListImages.Add , "Numbers", LoadPicture(MENU_ICON_PATH & "Numbers.bmp")
    On Error GoTo 0
    
    If UseProgressBar Then
        fProgForm.Hide
    End If
End Sub


