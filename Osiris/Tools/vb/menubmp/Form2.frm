VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Dynamic"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Static"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Height          =   300
      Index           =   2
      Left            =   1800
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   300
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Height          =   300
      Index           =   1
      Left            =   1800
      Picture         =   "Form2.frx":0102
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   300
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Height          =   300
      Index           =   0
      Left            =   1800
      Picture         =   "Form2.frx":0444
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   300
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Height          =   300
      Left            =   2400
      Picture         =   "Form2.frx":0786
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   600
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Height          =   300
      Index           =   2
      Left            =   4080
      Picture         =   "Form2.frx":0888
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Height          =   300
      Index           =   1
      Left            =   4080
      Picture         =   "Form2.frx":098A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   840
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      Height          =   300
      Index           =   0
      Left            =   4080
      Picture         =   "Form2.frx":0A8C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      Top             =   120
      Width           =   300
   End
   Begin VB.Menu TopMenu 
      Caption         =   "BitMenu"
      Begin VB.Menu SubMenu 
         Caption         =   "SubMenu0"
         Index           =   0
      End
      Begin VB.Menu SubMenu 
         Caption         =   "SubMenu1"
         Index           =   1
      End
      Begin VB.Menu SubMenu 
         Caption         =   "SubMenu2"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub SubMenu_Click(Index As Integer)

                         ' Uncheck presently checked item, check new item
                         Static LastSelection As Integer
                         SubMenu(LastSelection).Checked = False
                         SubMenu(Index).Checked = True
                         LastSelection = Index

                    End Sub

                    Sub Command1_Click()

                         Dim i As Integer
                         Dim x As Long
                         Dim hMenu As Long
                         Dim hSubMenu As Long
                         Dim menuId As Long

                         'to create a static bitmap menu
                         hMenu = GetMenu(Me.hwnd)
                         hSubMenu = GetSubMenu(hMenu, 0)
                         For i = 0 To Number_of_Menu_Selections - 1
                         menuId = GetMenuItemID(hSubMenu, i)
                         x = ModifyMenu(hMenu, menuId, MF_BITMAP, menuId, _
                         CLng(Picture1(i).Picture))
                         x = SetMenuItemBitmaps(hMenu, menuId, 0, 0, _
                         CLng(Picture2.Picture))
                         Next

                    End Sub

                    Sub Command2_Click()

                         Dim i As Integer
                         Dim x As Long
                         Dim hMenu As Long
                         Dim hSubMenu As Long
                         Dim menuId As Long

                         'to create a dynamic menu system
                         hMenu = GetMenu(Me.hwnd)
                         hSubMenu = GetSubMenu(hMenu, 0)
                         For i = 0 To Number_of_Menu_Selections - 1
                         'Place some text into the menu.
                         SubMenu(i).Caption = Picture3(i).FontName & _
                         Str$(Picture3(i).FontSize) + " Pnt"
                         '1. Must be AutoRedraw for Image().
                         '2. Set Backcolor of Picture control to that of the
                         ' current system Menu Bar color, so Dynamic bitmaps will appear
                         ' as normal menu items when menu bar color is changed via the
                         ' control panel
                         '3. See the bitmaps on screen, this could all be done at design time.
                         Picture3(i).AutoRedraw = True
                         Picture3(i).BackColor = CLR_MENUBAR
                         ' You can uncomment this
                         ' Picture3(i).Visible = False
                         ' Set the width and height of the Picture controls based on their
                         ' corresponding Menu items caption, and the Picture controls
                         ' Font and FontSize.
                         'DoEvents() is necessary to make new dimension
                         'values to take affect prior to exiting this Sub.
                         Picture3(i).Width = Picture3(i).TextWidth(SubMenu(i).Caption)
                         Picture3(i).Height = Picture3(i).TextHeight(SubMenu(i).Caption)
                         Picture3(i).Print SubMenu(i).Caption
                         'Set picture controls backgroup picture (Bitmap) to its Image.
                         Picture3(i).Picture = Picture3(i).Image
                         x = DoEvents()
                         Next
                         'Get handle to forms menu.
                         hMenu = GetMenu(Me.hwnd)
                         'Get handle to the specific menu in top level menu.
                         hSubMenu = GetSubMenu(hMenu, 0)
                         For i = 0 To Number_of_Menu_Selections - 1
                         'Get ID of sub menu
                         menuId = GetMenuItemID(hSubMenu, i)
                         'Replace menu text w/bitmap from corresponding picture control
                         x = ModifyMenu(hMenu, menuId, MF_BITMAP, menuId, _
                         CLng(Picture3(i).Picture))
                         'Replace bitmap for menu check mark with custom check bitmap
                         x = SetMenuItemBitmaps(hMenu, menuId, 0, 0, _
                         CLng(Picture2.Picture))
                         Next

                    End Sub

