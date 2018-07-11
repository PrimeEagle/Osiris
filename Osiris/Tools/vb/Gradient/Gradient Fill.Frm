VERSION 4.00
Begin VB.Form Fill 
   AutoRedraw      =   -1  'True
   Caption         =   "Gradient Backround"
   ClientHeight    =   3255
   ClientLeft      =   3225
   ClientTop       =   2160
   ClientWidth     =   4215
   Height          =   3660
   Icon            =   "Gradient Fill.frx":0000
   Left            =   3165
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleMode       =   0  'User
   ScaleWidth      =   4215
   Top             =   1815
   Width           =   4335
   Begin VB.PictureBox Mainframe 
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2955
      ScaleWidth      =   3915
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   3975
      Begin VB.PictureBox Redpicture 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   3000
         ScaleHeight     =   315
         ScaleWidth      =   675
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Width           =   735
      End
      Begin VB.PictureBox Greenpicture 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   3000
         ScaleHeight     =   315
         ScaleWidth      =   675
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   720
         Width           =   735
      End
      Begin VB.PictureBox Bluepicture 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   3000
         ScaleHeight     =   315
         ScaleWidth      =   675
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1320
         Width           =   735
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   3000
         ScaleHeight     =   315
         ScaleWidth      =   675
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton Update 
         Caption         =   "Update &Background..."
         Default         =   -1  'True
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   2520
         Width           =   2055
      End
      Begin VB.CommandButton Cancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2640
         TabIndex        =   10
         Top             =   2520
         Width           =   1095
      End
      Begin ComctlLib.Slider Redslider 
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   873
         _StockProps     =   64
         LargeChange     =   10
         Max             =   255
         TickFrequency   =   10
      End
      Begin ComctlLib.Slider Greenslider 
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   873
         _StockProps     =   64
         Max             =   255
         TickFrequency   =   10
      End
      Begin ComctlLib.Slider Blueslider 
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   873
         _StockProps     =   64
         Max             =   255
         SelStart        =   255
         TickFrequency   =   10
         Value           =   255
      End
      Begin VB.Label Actual 
         Caption         =   "&Actual Backround:"
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   2040
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Fill"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Blueslider_Click()

Gradient Bluepicture, 0, 0, Blueslider.Value
Gradient Picture4, Redslider.Value, Greenslider.Value, Blueslider.Value

End Sub

Private Sub Blueslider_Scroll()

Blueslider_Click

End Sub


Private Sub Cancel_Click()

Unload Me 'exit the program

End Sub




Private Sub Form_Resize()

'Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2 'centre the form on the screen

Blueslider_Click
Update_Click

'This project was downloaded from
'
'http://www.brianharper.demon.co.uk/
'
'Please use this project and all of its source code however you want.
'
'UNZIPPING
'To unzip the project files you will need a 32Bit unzipper program that
'can handle long file names. If you have a latest copy of Winzip installed
'on your system then you may use that. If you however dont have a copy,
'then visit my web site, go into the files section and from there you can
'click on the Winzip link to goto their site and download a copy of the
'program. By doing this you will now beable to unzip the project files
'retaining their proper long file names.
'Once upzipped, load up your copy of Visual Basic and goto
'File/Open Project. Locate the project files to where ever you unzipped
'them, then click Open. The project files will be loaded and are now ready
'for use.
'
'THE PROJECT
'If your ever creating your own Setup Utility program then this is the
'project that your going to need. With just about all Setup programs
'having them, then there is no reason why your shouldnt neither.
'
'NOTES
'I have only provided the necessary project files with the zip. This keeps
'the size of the zip files down to a minimum and enables me to upload more
'prjects files to my site.
'
'I hope you find the project usful in what ever you are programming. I
'have tried to write out a small explanation of what each line of code
'does in the project, although most of it is pretty simple to understand.
'
'If you find any bugs in the code then please dont hesitate to Email me and
'I will get back to you as soon as possible. If you however need help on a
'different matter concerning Visual Basic then please please Email me as
'I like to here from people and here what they are programming.
'
'My Email address is:
'Brian@brianharper.demon.co.uk
'
'My web site is:
'http://www.brianharper.demon.co.uk/
'
'Please visit my web site and find many other useful projects like this.
'

End Sub




Private Sub Greenslider_Click()

Gradient Greenpicture, 0, Greenslider.Value, 0
Gradient Picture4, Redslider.Value, Greenslider.Value, Blueslider.Value

End Sub

Private Sub Greenslider_Scroll()

Greenslider_Click

End Sub


Private Sub Redslider_Click()

Gradient Redpicture, Redslider.Value, 0, 0
Gradient Picture4, Redslider.Value, Greenslider.Value, Blueslider.Value

End Sub



Private Sub Redslider_Scroll()

Redslider_Click

End Sub


Private Sub Update_Click()

Gradient Me, Redslider.Value, Greenslider.Value, Blueslider.Value

End Sub


