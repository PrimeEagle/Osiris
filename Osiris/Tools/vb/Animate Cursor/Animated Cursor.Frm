VERSION 4.00
Begin VB.Form AniCursor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Animated Cursor"
   ClientHeight    =   1830
   ClientLeft      =   2955
   ClientTop       =   2475
   ClientWidth     =   3990
   Height          =   2235
   Icon            =   "Animated Cursor.frx":0000
   Left            =   2895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1830
   ScaleWidth      =   3990
   Top             =   2130
   Width           =   4110
   Begin VB.CommandButton Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   1320
      Width           =   975
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   1320
   End
   Begin VB.CommandButton Animate 
      Caption         =   "&Animate Cursor"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   1320
      Width           =   1515
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   9
      Left            =   2220
      Picture         =   "Animated Cursor.frx":000C
      Top             =   705
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   8
      Left            =   1740
      Picture         =   "Animated Cursor.frx":0316
      Top             =   705
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   7
      Left            =   1260
      Picture         =   "Animated Cursor.frx":0620
      Top             =   705
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   6
      Left            =   780
      Picture         =   "Animated Cursor.frx":092A
      Top             =   705
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   5
      Left            =   270
      Picture         =   "Animated Cursor.frx":0C34
      Top             =   705
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   4
      Left            =   2205
      Picture         =   "Animated Cursor.frx":0F3E
      Top             =   225
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   3
      Left            =   1710
      Picture         =   "Animated Cursor.frx":1248
      Top             =   225
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   2
      Left            =   1215
      Picture         =   "Animated Cursor.frx":1552
      Top             =   225
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   1
      Left            =   720
      Picture         =   "Animated Cursor.frx":185C
      Top             =   225
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgPtr 
      Height          =   480
      Index           =   0
      Left            =   270
      Picture         =   "Animated Cursor.frx":1B66
      Top             =   225
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "AniCursor"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit
'
' Index variable to track current cell in animation
'
Private CustomPtr As Integer

Private Sub Command1_Click()
   If Not Timer1.Enabled Then
      '
      ' Turn on animation if Timer isn't enabled.
      '
      Timer1.Enabled = True
      Me.MousePointer = vbCustom
      Command1.Caption = "Stop Cursor Animation"
   Else
      '
      ' Turn off animation if Timer is enabled.
      ' Reset MouseIcon and index to position 0.
      '
      Timer1.Enabled = False
      Me.MousePointer = vbDefault
      Command1.Caption = "Animate Cursor"
      CustomPtr = 0
      Me.MouseIcon = imgPtr(CustomPtr)
   End If
End Sub

Private Sub Animate_Click()

If Not Timer.Enabled Then 'turn on animation if Timer isn't enabled.
   Timer.Enabled = True
   Me.MousePointer = vbCustom
   Animate.Caption = "Stop Animation"
Else
   'turn off animation if Timer is enabled.
   'reset MouseIcon and index to position 0.
   Timer.Enabled = False
   Me.MousePointer = vbDefault
   Animate.Caption = "Animate Cursor"
   CustomPtr = 0
   Me.MouseIcon = imgPtr(CustomPtr)
End If
   
End Sub

Private Sub Timer1_Timer()
   '
   ' Set MouseIcon to current index, and increment
   ' index (looping back if neccessary).
   ' Use cursors designed for B&W if Win16, or
   ' full color cursors in Win32.
   '
   #If Win16 Then
      Me.MouseIcon = imgPtr(CustomPtr + 10)
   #ElseIf Win32 Then
      Me.MouseIcon = imgPtr(CustomPtr)
   #End If
   CustomPtr = (CustomPtr + 1) Mod 10
End Sub

Private Sub Cancel_Click()

Unload Me 'exit the program

End Sub

Private Sub Form_Load()

Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2 'centre the form on the screen

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
'This project can be very handy if your programming something that needs to
'add an animation to it. This project is very basic and uses the actual mouse
'cursor for the animation. When the mouse is moved over the form, or your
'program, then the cursor will change to an animation. This can be usfull when
'first starting your program or in a splash screen of some sort.
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

Private Sub imgPtr_Click(Index As Integer)

End Sub

Private Sub Timer_Timer()

'set MouseIcon to current index, and increment index (looping back if neccessary).
   
Me.MouseIcon = imgPtr(CustomPtr)
CustomPtr = (CustomPtr + 1) Mod 10
   
End Sub


