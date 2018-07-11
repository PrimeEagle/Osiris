Attribute VB_Name = "mWin32API"
Option Explicit
DefInt A-Z

'---------------------------------TYPE DEFINITIONS------------------------------
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Public Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Type DRAWITEMSTRUCT
        CtlType As Long
        CtlID As Long
        itemID As Long
        itemAction As Long
        itemState As Long
        hwndItem As Long
        hdc As Long
        rcItem As RECT
        ItemData As Long
End Type

Public Type MEASUREITEMSTRUCT
        CtlType As Long
        CtlID As Long
        itemID As Long
        itemWidth As Long
        ItemHeight As Long
        ItemData As Long
End Type

Public Enum ColConst
    ActiveBorder = 10
    ActiveCaption = 2
    ADJ_MAX = 100
    ADJ_MIN = -100 'shorts
    APPWORKSPACE = 12
    Background = 1
    BTNFACE = 15
    BTNHIGHLIGHT = 20
    BTNSHADOW = 16
    BTNTEXT = 18
    CAPTIONTEXT = 9
    GRAYTEXT = 17
    HIGHLIGHT = 13
    HIGHLIGHTTEXT = 14
    INACTIVEBORDER = 11
    INACTIVECAPTION = 3
    INACTIVECAPTIONTEXT = 19
    Menu = 4
    MENUTEXT = 7
    SCROLLBAR = 0
    Window = 5
    WINDOWFRAME = 6
    WINDOWTEXT = 8
End Enum

Public Enum StockObjects
   soWHITE_BRUSH = 0
   soLTGRAY_BRUSH = 1
   soGRAY_BRUSH = 2
   soDKGRAY_BRUSH = 3
   soBLACK_BRUSH = 4
   soNULL_BRUSH = 5
   soHOLLOW_BRUSH = 5
   soWHITE_PEN = 6
   soBLACK_PEN = 7
   soNULL_PEN = 8
   soOEM_FIXED_FONT = 10
   soANSI_FIXED_FONT = 11
   soANSI_VAR_FONT = 12
   soSYSTEM_FONT = 13
   soDEVICE_DEFAULT_FONT = 14
   soDEFAULT_PALETTE = 15
   soSYSTEM_FIXED_FONT = 16
   soSTOCK_LAST = 16
End Enum

'Type OFSTRUCT '136 bytes --- Data Structure for OpenFile call
' cBytes As String * 1
' fFixedDisk As String * 1
' nErrCode As Integer
' reserved As String * 4
' szPathName As String * 128
'End Type

'--------------------------------WINDOWS MESSAGES--------------------------------------

'Global Const BM_GETCHECK = &HF0
'Global Const BM_GETSTATE = &HF2
'Global Const BM_SETCHECK = &HF1
'Global Const BM_SETSTATE = &HF3
'Global Const BM_SETSTYLE = &HF4
'Global Const BN_CLICKED = 0
'Global Const BN_DISABLE = 4
'Global Const BN_DOUBLECLICKED = 5
'Global Const BN_HILITE = 2
'Global Const BN_PAINT = 1
'Global Const BN_UNHILITE = 3
'Global Const BS_3STATE = &H5&
'Global Const BS_AUTO3STATE = &H6&
'Global Const BS_AUTOCHECKBOX = &H3&
'Global Const BS_AUTORADIOBUTTON = &H9&
'Global Const BS_CHECKBOX = &H2&
'Global Const BS_DEFPUSHBUTTON = &H1&
'Global Const BS_DIBPATTERN = 5
'Global Const BS_DIBPATTERN8X8 = 8
'Global Const BS_DIBPATTERNPT = 6
'Global Const BS_GROUPBOX = &H7&
'Global Const BS_HATCHED = 2
'Global Const BS_HOLLOW = BS_NULL
'Global Const BS_INDEXED = 4
'Global Const BS_LEFTTEXT = &H20&
'Global Const BS_NULL = 1
'Global Const BS_OWNERDRAW = &HB&
'Global Const BS_PATTERN = 3
'Global Const BS_PATTERN8X8 = 7
'Global Const BS_PUSHBUTTON = &H0&
'Global Const BS_RADIOBUTTON = &H4&
'Global Const BS_SOLID = 0
'Global Const BS_USERBUTTON = &H8&
'Global Const CB_ADDSTRING = &H143
'Global Const CB_DELETESTRING = &H144
'Global Const CB_DIR = &H145
'Global Const CB_ERR = (-1)
'Global Const CB_ERRSPACE = (-2)
'Global Const CB_FINDSTRING = &H14C
'Global Const CB_FINDSTRINGEXACT = &H158
'Global Const CB_GETCOUNT = &H146
'Global Const CB_GETCURSEL = &H147
'Global Const CB_GETDROPPEDCONTROLRECT = &H152
'Global Const CB_GETDROPPEDSTATE = &H157
'Global Const CB_GETEDITSEL = &H140
'Global Const CB_GETEXTENDEDUI = &H156
'Global Const CB_GETITEMDATA = &H150
'Global Const CB_GETITEMHEIGHT = &H154
'Global Const CB_GETLBTEXT = &H148
'Global Const CB_GETLBTEXTLEN = &H149
'Global Const CB_GETLOCALE = &H15A
'Global Const CB_INSERTSTRING = &H14A
'Global Const CB_LIMITTEXT = &H141
'Global Const CB_MSGMAX = &H15B
'Global Const CB_OKAY = 0
'Global Const CB_RESETCONTENT = &H14B
'Global Const CB_SELECTSTRING = &H14D
'Global Const CB_SETCURSEL = &H14E
'Global Const CB_SETEDITSEL = &H142
'Global Const CB_SETEXTENDEDUI = &H155
'Global Const CB_SETITEMDATA = &H151
'Global Const CB_SETITEMHEIGHT = &H153
'Global Const CB_SETLOCALE = &H159
'Global Const CB_SHOWDROPDOWN = &H14F
'Global Const CN_EVENT = &H4
'Global Const CN_RECEIVE = &H1
'Global Const CN_TRANSMIT = &H2
'Global Const EM_CANUNDO = &HC6
'Global Const EM_EMPTYUNDOBUFFER = &HCD
'Global Const EM_FMTLINES = &HC8
'Global Const EM_GETFIRSTVISIBLELINE = &HCE
'Global Const EM_GETHANDLE = &HBD
'Global Const EM_GETLINE = &HC4
'Global Const EM_GETLINECOUNT = &HBA
'Global Const EM_GETMODIFY = &HB8
'Global Const EM_GETPASSWORDCHAR = &HD2
'Global Const EM_GETRECT = &HB2
'Global Const EM_GETSEL = &HB0
'Global Const EM_GETTHUMB = &HBE
'Global Const EM_GETWORDBREAKPROC = &HD1
'Global Const EM_LIMITTEXT = &HC5
'Global Const EM_LINEFROMCHAR = &HC9
'Global Const EM_LINEINDEX = &HBB
'Global Const EM_LINELENGTH = &HC1
'Global Const EM_LINESCROLL = &HB6
'Global Const EM_REPLACESEL = &HC2
'Global Const EM_SCROLL = &HB5
'Global Const EM_SCROLLCARET = &HB7
'Global Const EM_SETHANDLE = &HBC
Global Const EM_SETLIMITTEXT = &HC5
'Global Const EM_SETMODIFY = &HB9
'Global Const EM_SETPASSWORDCHAR = &HCC
'Global Const EM_SETREADONLY = &HCF
'Global Const EM_SETRECT = &HB3
'Global Const EM_SETRECTNP = &HB4
'Global Const EM_SETSEL = &HB1
'Global Const EM_SETTABSTOPS = &HCB
'Global Const EM_SETWORDBREAKPROC = &HD0
Global Const EM_UNDO = &HC7
'Global Const EN_CHANGE = &H300
'Global Const EN_ERRSPACE = &H500
'Global Const EN_HSCROLL = &H601
'Global Const EN_KILLFOCUS = &H200
'Global Const EN_MAXTEXT = &H501
'Global Const EN_SETFOCUS = &H100
'Global Const EN_UPDATE = &H400
'Global Const EN_VSCROLL = &H602
'Global Const ES_AUTOHSCROLL = &H80&
'Global Const ES_AUTOVSCROLL = &H40&
'Global Const ES_CENTER = &H1&
'Global Const ES_LEFT = &H0&
'Global Const ES_LOWERCASE = &H10&
'Global Const ES_MULTILINE = &H4&
'Global Const ES_NOHIDESEL = &H100&
'Global Const ES_OEMCONVERT = &H400&
'Global Const ES_PASSWORD = &H20&
'Global Const ES_READONLY = &H800&
'Global Const ES_RIGHT = &H2&
'Global Const ES_UPPERCASE = &H8&
'Global Const ES_WANTRETURN = &H1000&
'Global Const GCL_CBCLSEXTRA = (-20)
'Global Const GCL_CBWNDEXTRA = (-18)
'Global Const GCL_HBRBACKGROUND = (-10)
'Global Const GCL_HCURSOR = (-12)
'Global Const GCL_HICON = (-14)
'Global Const GCL_HMODULE = (-16)
'Global Const GCL_MENUNAME = (-8)
'Global Const GCL_STYLE = (-26)
'Global Const GCL_WNDPROC = (-24)
'Global Const GCW_ATOM = (-32)
Public Const GWL_WNDPROC = (-4&)
'Global Const HTCAPTION = 2
'Global Const LB_ADDFILE = &H196
'Global Const LB_ADDSTRING = &H180
'Global Const LB_CTLCODE = 0&
'Global Const LB_DELETESTRING = &H182
'Global Const LB_DIR = &H18D
'Global Const LB_ERR = (-1)
'Global Const LB_ERRSPACE = (-2)
'Global Const LB_FINDSTRING = &H18F
'Global Const LB_FINDSTRINGEXACT = &H1A2
'Global Const LB_GETANCHORINDEX = &H19D
'Global Const LB_GETCARETINDEX = &H19F
'Global Const LB_GETCOUNT = &H18B
'Global Const LB_GETCURSEL = &H188
'Global Const LB_GETHORIZONTALEXTENT = &H193
'Global Const LB_GETITEMDATA = &H199
'Global Const LB_GETITEMHEIGHT = &H1A1
'Global Const LB_GETITEMRECT = &H198
'Global Const LB_GETLOCALE = &H1A6
'Global Const LB_GETSEL = &H187
'Global Const LB_GETSELCOUNT = &H190
'Global Const LB_GETSELITEMS = &H191
'Global Const LB_GETTEXT = &H189
'Global Const LB_GETTEXTLEN = &H18A
'Global Const LB_GETTOPINDEX = &H18E
'Global Const LB_INSERTSTRING = &H181
'Global Const LB_MSGMAX = &H1A8
'Global Const LB_OKAY = 0
'Global Const LB_RESETCONTENT = &H184
'Global Const LB_SELECTSTRING = &H18C
'Global Const LB_SELITEMRANGE = &H19B
'Global Const LB_SELITEMRANGEEX = &H183
'Global Const LB_SETANCHORINDEX = &H19C
'Global Const LB_SETCARETINDEX = &H19E
'Global Const LB_SETCOLUMNWIDTH = &H195
'Global Const LB_SETCOUNT = &H1A7
'Global Const LB_SETCURSEL = &H186
'Global Const LB_SETHORIZONTALEXTENT = &H194
'Global Const LB_SETITEMDATA = &H19A
'Global Const LB_SETITEMHEIGHT = &H1A0
'Global Const LB_SETLOCALE = &H1A5
'Global Const LB_SETSEL = &H185
'Global Const LB_SETTABSTOPS = &H192
'Global Const LB_SETTOPINDEX = &H197
'Global Const LBN_DBLCLK = 2
'Global Const LBN_ERRSPACE = (-2)
'Global Const LBN_KILLFOCUS = 5
'Global Const LBN_SELCANCEL = 3
'Global Const LBN_SELCHANGE = 1
'Global Const LBN_SETFOCUS = 4
'Global Const LBS_DISABLENOSCROLL = &H1000&
'Global Const LBS_EXTENDEDSEL = &H800&
'Global Const LBS_HASSTRINGS = &H40&
'Global Const LBS_MULTICOLUMN = &H200&
'Global Const LBS_MULTIPLESEL = &H8&
'Global Const LBS_NODATA = &H2000&
'Global Const LBS_NOINTEGRALHEIGHT = &H100&
'Global Const LBS_NOREDRAW = &H4&
'Global Const LBS_NOTIFY = &H1&
'Global Const LBS_OWNERDRAWFIXED = &H10&
'Global Const LBS_OWNERDRAWVARIABLE = &H20&
'Global Const LBS_SORT = &H2&
'Global Const LBS_USETABSTOPS = &H80&
'Global Const LBS_WANTKEYBOARDINPUT = &H400&
Global Const LVM_FIRST = &H1000
Global Const LVM_GETSELECTEDCOUNT = (LVM_FIRST + 50)
'Global Const MF_APPEND = &H100&
'Global Const MF_BITMAP = &H4&
Global Const MF_BYCOMMAND = &H0&
'Global Const MF_BYPOSITION = &H400&
'Global Const MF_CALLBACKS = &H8000000
'Global Const MF_CHANGE = &H80&
'Global Const MF_CHECKED = &H8&
'Global Const MF_CONV = &H40000000
'Global Const MF_DELETE = &H200&
'Global Const MF_DISABLED = &H2&
'Global Const MF_ENABLED = &H0&
'Global Const MF_END = &H80
'Global Const MF_ERRORS = &H10000000
'Global Const MF_GRAYED = &H1&
'Global Const MF_HELP = &H4000&
'Global Const MF_HILITE = &H80&
'Global Const MF_HSZ_INFO = &H1000000
'Global Const MF_INSERT = &H0&
'Global Const MF_LINKS = &H20000000
'Global Const MF_MASK = &HFF000000
'Global Const MF_MENUBARBREAK = &H20&
'Global Const MF_MENUBREAK = &H40&
'Global Const MF_MOUSESELECT = &H8000&
Global Const MF_OWNERDRAW = &H100&
'Global Const MF_POPUP = &H10&
'Global Const MF_POSTMSGS = &H4000000
'Global Const MF_REMOVE = &H1000&
'Global Const MF_SENDMSGS = &H2000000
'Global Const MF_SEPARATOR = &H800&
'Global Const MF_STRING = &H0&
'Global Const MF_SYSMENU = &H2000&
'Global Const MF_UNCHECKED = &H0&
'Global Const MF_UNHILITE = &H0&
'Global Const MF_USECHECKBITMAPS = &H200&
'Global Const MFT_STRING = MF_STRING
'Global Const MFT_BITMAP = MF_BITMAP
'Global Const MFT_MENUBARBREAK = MF_MENUBARBREAK
'Global Const MFT_MENUBREAK = MF_MENUBREAK
'Global Const MFT_OWNERDRAW = MF_OWNERDRAW
'Global Const MFT_RADIOCHECK = &H200
'Global Const MFT_SEPARATOR = MF_SEPARATOR
'Global Const MFT_RIGHTORDER = &H2000
'Global Const ODS_CHECKED = &H8
'Global Const ODS_DISABLED = &H4
'Global Const ODS_FOCUS = &H10
'Global Const ODS_GRAYED = &H2
Global Const ODS_SELECTED = &H1
'Global Const PWR_CRITICALRESUME = 3
'Global Const PWR_FAIL = (-1)
'Global Const PWR_OK = 1
'Global Const PWR_SUSPENDREQUEST = 1
'Global Const PWR_SUSPENDRESUME = 2
'Global Const RBHT_CAPTION = &H2
'Global Const RBHT_GRABBER = &H4
'Global Const RBHT_CLIENT = &H3
'Global Const RB_HITTEST = (WM_USER + 8)
'Global Const REBARCLASSNAME = "ReBarWindow32"
'Global Const ICC_COOL_CLASSES = &H400       ' rebar (coolbar) control
'Global Const ICC_BAR_CLASSES = &H4             ' toolbar, statusbar, trackbar, tooltips
'Global Const RBS_VARHEIGHT = &H200
'Global Const CCS_NODIVIDER = &H40
'Global Const RB_SETBARINFO = WM_USER + 4
'Global Const RBBIM_COLORS = &H2
'Global Const RBBIM_TEXT = &H4
'Global Const RBBIM_BACKGROUND = &H80
'Global Const RBBIM_STYLE = &H1
'Global Const RBBIM_CHILD = &H10
'Global Const RBBIM_CHILDSIZE = &H20
'Global Const RBBIM_SIZE = &H40
'Global Const RBBS_CHILDEDGE = &H4 ' edge around top & bottom of child window
'Global Const RBBS_FIXEDBMP = &H20  ' bitmap doesn't move during band resize
Global Const SW_NORMAL = 1
'Global Const TB_GETSTYLE = &H400 + 57
'Global Const TB_SETSTYLE = &H400 + 56
'Global Const TBSTYLE_FLAT = &H800
'Global Const WA_ACTIVE = 1
'Global Const WA_CLICKACTIVE = 2
'Global Const WA_INACTIVE = 0
'Global Const WM_ACTIVATE = &H6
'Global Const WM_ACTIVATEAPP = &H1C
'Global Const WM_ASKCBFORMATNAME = &H30C
'Global Const WM_CANCELJOURNAL = &H4B
'Global Const WM_CANCELMODE = &H1F
'Global Const WM_CAPTURECHANGED = &H215
'Global Const WM_CHANGECBCHAIN = &H30D
Global Const WM_CHAR = &H102
'Global Const WM_CHARTOITEM = &H2F
'Global Const WM_CHILDACTIVATE = &H22
'Global Const WM_CLEAR = &H303
'Global Const WM_CLOSE = &H10
Global Const WM_COMMAND = &H111
'Global Const WM_COMPACTING = &H41
'Global Const WM_COMPAREITEM = &H39
Global Const WM_COPY = &H301
'Global Const WM_COPYDATA = &H4A
'Global Const WM_CREATE = &H1
'Global Const WM_CTLCOLORBTN = &H135
'Global Const WM_CTLCOLORDLG = &H136
'Global Const WM_CTLCOLOREDIT = &H133
'Global Const WM_CTLCOLORLISTBOX = &H134
'Global Const WM_CTLCOLORMSGBOX = &H132
'Global Const WM_CTLCOLORSCROLLBAR = &H137
'Global Const WM_CTLCOLORSTATIC = &H138
Global Const WM_CUT = &H300
'Global Const WM_DEADCHAR = &H103
'Global Const WM_DELETEITEM = &H2D
'Global Const WM_DESTROY = &H2
'Global Const WM_DESTROYCLIPBOARD = &H307
'Global Const WM_DEVMODECHANGE = &H1B
'Global Const WM_DRAWCLIPBOARD = &H308
Global Const WM_DRAWITEM = &H2B
'Global Const WM_DROPFILES = &H233
'Global Const WM_ENABLE = &HA
'Global Const WM_ENDSESSION = &H16
Global Const WM_ENTERIDLE = &H121
'Global Const WM_ENTERMENULOOP = &H211
Public Const WM_ENTERSIZEMOVE = &H231
Public Const WM_EXITSIZEMOVE = &H232
Global Const WM_ERASEBKGND = &H14
'Global Const WM_EXITMENULOOP = &H212
'Global Const WM_FONTCHANGE = &H1D
'Global Const WM_GETDLGCODE = &H87
'Global Const WM_GETHOTKEY = &H33
Global Const WM_GETFONT = &H31
Global Const WM_GETMINMAXINFO = &H24
Global Const WM_GETTEXT = &HD
Global Const WM_GETTEXTLENGTH = &HE
'Global Const WM_HOTKEY = &H312
'Global Const WM_HSCROLL = &H114
'Global Const WM_HSCROLLCLIPBOARD = &H30E
'Global Const WM_ICONERASEBKGND = &H27
'Global Const WM_INITDIALOG = &H110
'Global Const WM_INITMENU = &H116
Global Const WM_INITMENUPOPUP = &H117
Global Const WM_KEYDOWN = &H100
'Global Const WM_KEYFIRST = &H100
'Global Const WM_KEYLAST = &H108
Global Const WM_KEYUP = &H101
'Global Const WM_KILLFOCUS = &H8
'Global Const WM_LBUTTONDBLCLK = &H203
Global Const WM_LBUTTONDOWN = &H201
'Global Const WM_LBUTTONUP = &H202
'Global Const WM_MBUTTONDBLCLK = &H209
'Global Const WM_MBUTTONDOWN = &H207
'Global Const WM_MBUTTONUP = &H208
'Global Const WM_MDIACTIVATE = &H222
'Global Const WM_MDICASCADE = &H227
'Global Const WM_MDICREATE = &H220
'Global Const WM_MDIDESTROY = &H221
'Global Const WM_MDIGETACTIVE = &H229
'Global Const WM_MDIICONARRANGE = &H228
'Global Const WM_MDIMAXIMIZE = &H225
'Global Const WM_MDINEXT = &H224
'Global Const WM_MDIREFRESHMENU = &H234
'Global Const WM_MDIRESTORE = &H223
'Global Const WM_MDISETMENU = &H230
'Global Const WM_MDITILE = &H226
Global Const WM_MEASUREITEM = &H2C
'Global Const WM_MENUCHAR = &H120
'Global Const WM_MENUSELECT = &H11F
'Global Const WM_MOUSEACTIVATE = &H21
'Global Const WM_MOUSEFIRST = &H200
'Global Const WM_MOUSELAST = &H209
'Global Const WM_MOUSEMOVE = &H200
'Global Const WM_MOVE = &H3
'Global Const WM_NCACTIVATE = &H86
'Global Const WM_NCCALCSIZE = &H83
'Global Const WM_NCCREATE = &H81
'Global Const WM_NCDESTROY = &H82
'Global Const WM_NCHITTEST = &H84
'Global Const WM_NCLBUTTONDBLCLK = &HA3
'Global Const WM_NCLBUTTONDOWN = &HA1
'Global Const WM_NCLBUTTONUP = &HA2
'Global Const WM_NCMBUTTONDBLCLK = &HA9
'Global Const WM_NCMBUTTONDOWN = &HA7
'Global Const WM_NCMBUTTONUP = &HA8
'Global Const WM_NCMOUSEMOVE = &HA0
'Global Const WM_NCPAINT = &H85
'Global Const WM_NCRBUTTONDBLCLK = &HA6
'Global Const WM_NCRBUTTONDOWN = &HA4
'Global Const WM_NCRBUTTONUP = &HA5
'Global Const WM_NEXTDLGCTL = &H28
'Global Const WM_NULL = &H0
Global Const WM_PAINT = &HF
'Global Const WM_PAINTCLIPBOARD = &H309
'Global Const WM_PAINTICON = &H26
'Global Const WM_PALETTECHANGED = &H311
'Global Const WM_PALETTEISCHANGING = &H310
Global Const WM_PARENTNOTIFY = &H210
Global Const WM_PASTE = &H302
'Global Const WM_PENWINFIRST = &H380
'Global Const WM_PENWINLAST = &H38F
'Global Const WM_POWER = &H48
'Global Const WM_QUERYDRAGICON = &H37
'Global Const WM_QUERYENDSESSION = &H11
'Global Const WM_QUERYNEWPALETTE = &H30F
'Global Const WM_QUERYOPEN = &H13
'Global Const WM_QUEUESYNC = &H23
'Global Const WM_QUIT = &H12
'Global Const WM_RBUTTONDBLCLK = &H206
'Global Const WM_RBUTTONDOWN = &H204
'Global Const WM_RBUTTONUP = &H205
'Global Const WM_RENDERALLFORMATS = &H306
'Global Const WM_RENDERFORMAT = &H305
Global Const WM_SETCURSOR = &H20
'Global Const WM_SETFOCUS = &H7
'Global Const WM_SETFONT = &H30
'Global Const WM_SETHOTKEY = &H32
Global Const WM_SETREDRAW = &HB
'Global Const WM_SETTEXT = &HC
Global Const WM_SHOWWINDOW = &H18
Global Const WM_SIZE = &H5
'Global Const WM_SIZECLIPBOARD = &H30B
'Global Const WM_SPOOLERSTATUS = &H2A
'Global Const WM_SYSCHAR = &H106
'Global Const WM_SYSCOLORCHANGE = &H15
'Global Const WM_SYSCOMMAND = &H112
'Global Const WM_SYSDEADCHAR = &H107
'Global Const WM_SYSKEYDOWN = &H104
'Global Const WM_SYSKEYUP = &H105
'Global Const WM_TIMECHANGE = &H1E
'Global Const WM_TIMER = &H113
'Global Const WM_UNDO = &H304
Global Const WM_USER = &H400
'Global Const WM_VKEYTOITEM = &H2E
'Global Const WM_VSCROLL = &H115
'Global Const WM_VSCROLLCLIPBOARD = &H30A
'Global Const WM_WINDOWPOSCHANGED = &H47
'Global Const WM_WINDOWPOSCHANGING = &H46
'Global Const WM_WININICHANGE = &H1A

'--------------------------------CONSTANTS--------------------------------------
Global Const COLOR_BTNHIGHLIGHT = 20
Global Const COLOR_BTNSHADOW = 16
Global Const NULL_BRUSH = 5

'--------------------------------VARIABLES--------------------------------------
'Global PrevWorkspcColor As Long

'--------------------------------DECLARATAIONS----------------------------------
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Integer
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wcmd As Integer) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Sub GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT)
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal aint As Integer) As Integer
Public Declare Sub InvalidateRect Lib "user32" (ByVal hwnd As Integer, lpRect As RECT, ByVal bErase As Integer)
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hdc As Long, ByVal nCount As Long, ByVal lpFacename As String) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Integer) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Integer
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Integer) As Long
Public Declare Sub SetSysColors Lib "user32" (ByVal nChanges As Integer, lpSysColor As Integer, lpColorValues As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetUpdateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function CallWindowProc& Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Sub ValidateRect Lib "user32" (ByVal hwnd As Integer, lpRect As RECT)
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Integer
Public Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Function SysColor2RGB(ByVal aColor As Long) As Long
    aColor = aColor And (Not &H80000000)
    SysColor2RGB = GetSysColor(aColor)
End Function

