VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSave 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   2880
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   300
   End
   Begin VB.PictureBox picNew 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   840
      Picture         =   "Form1.frx":0102
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   300
   End
   Begin VB.PictureBox picOpen 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   1800
      Picture         =   "Form1.frx":0204
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   300
   End
   Begin VB.Menu FileMenu 
      Caption         =   "File"
      Begin VB.Menu Menu 
         Caption         =   "Open"
         Index           =   1
      End
      Begin VB.Menu Menu 
         Caption         =   "Save"
         Index           =   2
      End
      Begin VB.Menu Menu 
         Caption         =   "New"
         Index           =   3
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'API's declaration
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As _
Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal _
hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As _
Long, ByVal nPos As Long) As Long

Private Sub Form_Load()
   InsertBMP2Menu Me, 0, 0, picOpen
   InsertBMP2Menu Me, 0, 1, picSave
   InsertBMP2Menu Me, 0, 2, picNew
End Sub

'                           InsertBMP2Menu
'porpose: to insert Bitmap to the menu.
'frmForm is the form that contain the menu.
'nMenuPos: the No of the menu row ( Base 0 ), for example:
'   File = 0, Edit = 1, View = 2 ( the horizontal position.
'
'nSubMenuPos: the No of the sub menu ( Base 0 ), for example:
'   New = 0, Open = 1, Save = 2, Save As = 3, they all under File !
'picBMP is the picture control that its picture is being
'insert into the menu.
'Last update: 05/03/1998
Sub InsertBMP2Menu(frmForm As Form, nMenuPos As Integer, _
   nSubMenuPos As Integer, picBMP As PictureBox)

   Dim hMenu As Long, hSubMenu As Long, menuId As Long

   'Get the form menu handle
   hMenu = GetMenu(hwnd)
   'Get the form SubMenu handle
   hSubMenu = GetSubMenu(hMenu, nMenuPos)

   'get the menu item ID
   menuId = GetMenuItemID(hSubMenu, nSubMenuPos)

   'set picture to the menu item
   SetMenuItemBitmaps hMenu, menuId, 0, CLng(picBMP.Picture), _
CLng(picBMP.Picture)
End Sub

