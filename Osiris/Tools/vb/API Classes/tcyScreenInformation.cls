VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsScreenInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SM_CXVSCROLL = 2      'Width of arrow bitmap on vertical scroll bar
Private Const SM_CYHSCROLL = 3      'Height of arrow bitmap on horizontal scroll bar
Private Const SM_CYCAPTION = 4      'Height of caption
Private Const SM_CXBORDER = 5       ''Width of window frame that cannot be sized
Private Const SM_CYBORDER = 6       ''Height of window frame that cannot be sized
Private Const SM_CXDLGFRAME = 7     'Width of frame when window has WS_DLGFRAME style
Private Const SM_CYDLGFRAME = 8     'Height of frame when window has WS_DLGFRAME style
Private Const SM_CYVTHUMB = 9       'Height of thumb on horizontal scroll bar
Private Const SM_CXHTHUMB = 10      'Width of thumb on horizontal scroll bar
Private Const SM_CXICON = 11        ''Width of icon
Private Const SM_CYICON = 12        ''Height of icon
Private Const SM_CXCURSOR = 13      ''Width of cursor
Private Const SM_CYCURSOR = 14      ''Height of cursor
Private Const SM_CYMENU = 15        'Height of single-line menu
Private Const SM_CXFULLSCREEN = 16  ''Width of window client area for full-screen window
Private Const SM_CYFULLSCREEN = 17  ''Height of window client area for full-screen window
Private Const SM_CYKANJIWINDOW = 18 'Height of Kanji window
Private Const SM_MOUSEPRESENT = 19  'Mouse present
Private Const SM_CYVSCROLL = 20     'Height of arrow bitmap on vertical scroll bar
Private Const SM_CXHSCROLL = 21     'Width of arrow bitmap on horizontal scroll bar
Private Const SM_DEBUG = 22         'Nonzero if Windows debug version
Private Const SM_SWAPBUTTON = 23    'Non-zero if the left and right mouse buttons are swapped
Private Const SM_CXMIN = 28         ''Minimum width of window
Private Const SM_CYMIN = 29         ''Minimum height of window
Private Const SM_CXSIZE = 30        'Width of bitmaps contained in the title bar
Private Const SM_CYSIZE = 31        'Height of bitmaps contained in the title bar
Private Const SM_CXFRAME = 32       'Width of window frame
Private Const SM_CYFRAME = 33       'Height of window frame
Private Const SM_CXMINTRACK = 34    'Minimum tracking width of window
Private Const SM_CYMINTRACK = 35    'Minimum tracking height of window
Private Const SM_CXDOUBLECLK = 36       'Width of rectangle around the location of the first click. The second click must occur in the same rectangular location.
Private Const SM_CYDOUBLECLK = 37       'Height of rectangle around the location of the first click. The second click must occur in the same rectangular location.
Private Const SM_CXICONSPACING = 38     'Width of rectangles the system uses to position tiled icons
Private Const SM_CYICONSPACING = 39     'Height of rectangles the system uses to position tiled icons
Private Const SM_MENUDROPALIGNMENT = 40 'Alignment of pop-up menus. If zero, left side is aligned with corresponding left side of menu-bar item. If non-zero, left side is aligned with right side of corresponding menu bar item
Private Const SM_PENWINDOWS = 41        'Handle of Pen Windows dynamic link library if Pen Windows is installed
Private Const SM_DBCSENABLED = 42       'Returns a non-zero if the current Windows version uses double-byte characters, otherwise returns zero
Private Const SM_CMOUSEBUTTONS = 43

'Default Property Values:
Private Const m_def_Window_MaxHeight = 0
Private Const m_def_ScrollH_ArrowHeight = 0
Private Const m_def_ScrollH_ArrowWidth = 0
Private Const m_def_ScrollV_ArrowHeight = 0
Private Const m_def_ScrollV_ArrowWidth = 0
Private Const m_def_Dialog_FrameWidth = 0
Private Const m_def_Dialog_FrameHeight = 0
Private Const m_def_ScrollV_ThumbHeight = 0
Private Const m_def_ScrollH_ThumbWidth = 0
Private Const m_def_PenWindows = 0
Private Const m_def_Mouse_Present = False
Private Const m_def_IsDebugVersion = False
Private Const m_def_Mouse_ButtonsSwapped = False
Private Const m_def_Title_BMPHeight = 0
Private Const m_def_Title_BMPWidth = 0
Private Const m_def_Window_FrameHeight = 0
Private Const m_def_Window_FrameWidth = 0
Private Const m_def_Window_TrackingHeight = 0
Private Const m_def_Window_TrackingWidth = 0
Private Const m_def_Mouse_DblClkHeight = 0
Private Const m_def_Mouse_DblClkWidth = 0
Private Const m_def_Icon_SpacingX = 0
Private Const m_def_Icon_SpacingY = 0
Private Const m_def_Menu_DropAlignment = 0
Private Const m_def_UsesDoubleByte = 0
Private Const m_def_Mouse_Buttons = 0
Private Const m_def_Border_NonSizeHeight = 0
Private Const m_def_Border_NonSizeWidth = 0
Private Const m_def_Caption_Height = 0
Private Const m_def_Cursor_Width = 0
Private Const m_def_Cursor_Height = 0
Private Const m_def_Icon_Width = 0
Private Const m_def_Icon_Height = 0
Private Const m_def_MenuBar_Height = 0
Private Const m_def_Window_MaxWidth = 0
Private Const m_def_Window_MinHeight = 0
Private Const m_def_Window_MinWidth = 0
Private Const m_def_Icon_SmallHeight = 0
Private Const m_def_Icon_SmallWidth = 0

'Property Variables:
Dim m_Window_MaxHeight As Long
Dim m_ScrollH_ArrowHeight As Long
Dim m_ScrollH_ArrowWidth As Long
Dim m_ScrollV_ArrowHeight As Long
Dim m_ScrollV_ArrowWidth As Long
Dim m_Dialog_FrameWidth As Long
Dim m_Dialog_FrameHeight As Long
Dim m_ScrollV_ThumbHeight As Long
Dim m_ScrollH_ThumbWidth As Long
Dim m_Windows_KanjiHeight As Object
Dim m_PenWindows As Long
Dim m_Mouse_Present As Boolean
Dim m_IsDebugVersion As Boolean
Dim m_Mouse_ButtonsSwapped As Boolean
Dim m_Title_BMPHeight As Long
Dim m_Title_BMPWidth As Long
Dim m_Window_FrameHeight As Long
Dim m_Window_FrameWidth As Long
Dim m_Window_TrackingHeight As Long
Dim m_Window_TrackingWidth As Long
Dim m_Mouse_DblClkHeight As Long
Dim m_Mouse_DblClkWidth As Long
Dim m_Icon_SpacingX As Long
Dim m_Icon_SpacingY As Long
Dim m_Menu_DropAlignment As Long
Dim m_UsesDoubleByte As Long
Dim m_Mouse_Buttons As Integer
Dim m_Border_NonSizeHeight As Long
Dim m_Border_NonSizeWidth As Long
Dim m_Caption_Height As Long
Dim m_Cursor_Width As Long
Dim m_Cursor_Height As Long
Dim m_Icon_Width As Long
Dim m_Icon_Height As Long
Dim m_MenuBar_Height As Long
Dim m_Window_MaxWidth As Long
Dim m_Window_MinHeight As Long
Dim m_Window_MinWidth As Long
Dim m_Icon_SmallHeight As Long
Dim m_Icon_SmallWidth As Long



Public Property Get Border_NonSizeHeight() As Long
    Border_NonSizeHeight = m_Border_NonSizeHeight
End Property

Public Property Get Border_NonSizeWidth() As Long
    Border_NonSizeWidth = m_Border_NonSizeWidth
End Property

Public Property Get Caption_Height() As Long
    Caption_Height = m_Caption_Height
End Property

Public Property Get Cursor_Width() As Long
    Cursor_Width = m_Cursor_Width
End Property

Public Property Get Cursor_Height() As Long
    Cursor_Height = m_Cursor_Height
End Property

Public Property Get Icon_Width() As Long
    Icon_Width = m_Icon_Width
End Property

Public Property Get Icon_Height() As Long
    Icon_Height = m_Icon_Height
End Property

Public Property Get MenuBar_Height() As Long
    MenuBar_Height = m_MenuBar_Height
End Property

Public Property Get Window_MaxWidth() As Long
    Window_MaxWidth = m_Window_MaxWidth
End Property

Public Property Get Window_MinHeight() As Long
    Window_MinHeight = m_Window_MinHeight
End Property

Public Property Get Window_MinWidth() As Long
    Window_MinWidth = m_Window_MinWidth
End Property

Public Property Get Icon_SmallHeight() As Long
    Icon_SmallHeight = m_Icon_SmallHeight
End Property

Public Property Get Icon_SmallWidth() As Long
    Icon_SmallWidth = m_Icon_SmallWidth
End Property

'Initialize Properties for User Control
Private Sub Class_Initialize()
    m_Border_NonSizeHeight = GetSystemMetrics(SM_CYBORDER)
    m_Border_NonSizeWidth = GetSystemMetrics(SM_CXBORDER)
    m_Caption_Height = GetSystemMetrics(SM_CYCAPTION)
    m_Cursor_Width = GetSystemMetrics(SM_CXCURSOR)
    m_Cursor_Height = GetSystemMetrics(SM_CYCURSOR)
    m_Icon_Width = GetSystemMetrics(SM_CXICON)
    m_Icon_Height = GetSystemMetrics(SM_CYICON)
    m_MenuBar_Height = GetSystemMetrics(SM_CYMENU)
    m_Window_MaxHeight = GetSystemMetrics(SM_CYFULLSCREEN)
    m_Window_MaxWidth = GetSystemMetrics(SM_CXFULLSCREEN)
    m_Window_MinHeight = GetSystemMetrics(SM_CYMIN)
    m_Window_MinWidth = GetSystemMetrics(SM_CXMIN)
    
    'm_Icon_SmallHeight = GetSystemMetrics(SM_CYSMICON)
    'm_Icon_SmallWidth = GetSystemMetrics(SM_CXSMICON)
    m_PenWindows = GetSystemMetrics(SM_PENWINDOWS)
    
    m_ScrollH_ArrowHeight = GetSystemMetrics(SM_CYHSCROLL)
    m_ScrollH_ArrowWidth = GetSystemMetrics(SM_CXHSCROLL)
    m_ScrollV_ArrowHeight = GetSystemMetrics(SM_CYVSCROLL)
    m_ScrollV_ArrowWidth = GetSystemMetrics(SM_CXVSCROLL)
    m_Dialog_FrameWidth = GetSystemMetrics(SM_CXDLGFRAME)
    m_Dialog_FrameHeight = GetSystemMetrics(SM_CYDLGFRAME)
    m_ScrollV_ThumbHeight = GetSystemMetrics(SM_CYVTHUMB)
    m_ScrollH_ThumbWidth = GetSystemMetrics(SM_CXHTHUMB)
    m_Mouse_Present = (GetSystemMetrics(SM_MOUSEPRESENT) <> 0)
    m_IsDebugVersion = (GetSystemMetrics(SM_DEBUG) <> 0)
    m_Mouse_ButtonsSwapped = (GetSystemMetrics(SM_SWAPBUTTON) <> 0)
    m_Title_BMPHeight = GetSystemMetrics(SM_CYSIZE)
    m_Title_BMPWidth = GetSystemMetrics(SM_CXSIZE)
    m_Window_FrameHeight = GetSystemMetrics(SM_CYFRAME)
    m_Window_FrameWidth = GetSystemMetrics(SM_CXFRAME)
    m_Window_TrackingHeight = GetSystemMetrics(SM_CYMINTRACK)
    m_Window_TrackingWidth = GetSystemMetrics(SM_CXMINTRACK)
    m_Mouse_DblClkHeight = GetSystemMetrics(SM_CYDOUBLECLK)
    m_Mouse_DblClkWidth = GetSystemMetrics(SM_CXDOUBLECLK)
    m_Icon_SpacingX = GetSystemMetrics(SM_CXICONSPACING)
    m_Icon_SpacingY = GetSystemMetrics(SM_CYICONSPACING)
    m_Menu_DropAlignment = GetSystemMetrics(SM_MENUDROPALIGNMENT)
    m_UsesDoubleByte = (GetSystemMetrics(SM_DBCSENABLED) <> 0)
    m_Mouse_Buttons = GetSystemMetrics(SM_CMOUSEBUTTONS)
End Sub

Public Property Get Window_MaxHeight() As Long
    Window_MaxHeight = m_Window_MaxHeight
End Property


Public Property Get ScrollH_ArrowHeight() As Long
    ScrollH_ArrowHeight = m_ScrollH_ArrowHeight
End Property

Public Property Get ScrollH_ArrowWidth() As Long
    ScrollH_ArrowWidth = m_ScrollH_ArrowWidth
End Property

Public Property Get ScrollV_ArrowHeight() As Long
    ScrollV_ArrowHeight = m_ScrollV_ArrowHeight
End Property

Public Property Get ScrollV_ArrowWidth() As Long
    ScrollV_ArrowWidth = m_ScrollV_ArrowWidth
End Property

Public Property Get Dialog_FrameWidth() As Long
    Dialog_FrameWidth = m_Dialog_FrameWidth
End Property

Public Property Get Dialog_FrameHeight() As Long
    Dialog_FrameHeight = m_Dialog_FrameHeight
End Property

Public Property Get ScrollV_ThumbHeight() As Long
    ScrollV_ThumbHeight = m_ScrollV_ThumbHeight
End Property

Public Property Get ScrollH_ThumbWidth() As Long
    ScrollH_ThumbWidth = m_ScrollH_ThumbWidth
End Property

Public Property Get Windows_KanjiHeight() As Object
    Set Windows_KanjiHeight = m_Windows_KanjiHeight
End Property

Public Property Get PenWindows() As Long
    PenWindows = m_PenWindows
End Property

Public Property Get Mouse_Present() As Boolean
    Mouse_Present = m_Mouse_Present
End Property

Public Property Get IsDebugVersion() As Boolean
    IsDebugVersion = m_IsDebugVersion
End Property

Public Property Get Mouse_ButtonsSwapped() As Boolean
    Mouse_ButtonsSwapped = m_Mouse_ButtonsSwapped
End Property

Public Property Get Title_BMPHeight() As Long
    Title_BMPHeight = m_Title_BMPHeight
End Property

Public Property Get Title_BMPWidth() As Long
    Title_BMPWidth = m_Title_BMPWidth
End Property

Public Property Get Window_FrameHeight() As Long
    Window_FrameHeight = m_Window_FrameHeight
End Property

Public Property Get Window_FrameWidth() As Long
    Window_FrameWidth = m_Window_FrameWidth
End Property

Public Property Get Window_TrackingHeight() As Long
    Window_TrackingHeight = m_Window_TrackingHeight
End Property

Public Property Get Window_TrackingWidth() As Long
    Window_TrackingWidth = m_Window_TrackingWidth
End Property

Public Property Get Mouse_DblClkHeight() As Long
    Mouse_DblClkHeight = m_Mouse_DblClkHeight
End Property

Public Property Get Mouse_DblClkWidth() As Long
    Mouse_DblClkWidth = m_Mouse_DblClkWidth
End Property

Public Property Get Icon_SpacingX() As Long
    Icon_SpacingX = m_Icon_SpacingX
End Property

Public Property Get Icon_SpacingY() As Long
    Icon_SpacingY = m_Icon_SpacingY
End Property

Public Property Get Menu_DropAlignment() As Long
    Menu_DropAlignment = m_Menu_DropAlignment
End Property

Public Property Get UsesDoubleByte() As Long
    UsesDoubleByte = m_UsesDoubleByte
End Property

Public Property Get Mouse_Buttons() As Integer
    Mouse_Buttons = m_Mouse_Buttons
End Property


