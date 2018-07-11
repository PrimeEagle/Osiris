VERSION 4.00
Begin VB.Form Help 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help File Executor"
   ClientHeight    =   1935
   ClientLeft      =   1905
   ClientTop       =   1950
   ClientWidth     =   4335
   ForeColor       =   &H80000008&
   Height          =   2340
   Icon            =   "Help Files.frx":0000
   Left            =   1845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4335
   Top             =   1605
   Width           =   4455
   Begin VB.CommandButton Apply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   1440
      Width           =   975
   End
   Begin VB.Frame Actions 
      Caption         =   "Actions:"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   3015
      Begin VB.OptionButton Usinghelp 
         Caption         =   "&Using Help"
         Height          =   195
         Left            =   1680
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Search 
         Caption         =   "&Search"
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Find 
         Caption         =   "&Find"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Contents 
         Caption         =   "C&ontents"
         Height          =   195
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.TextBox Filename 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.CommandButton Browse 
      Caption         =   "&Browse..."
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   3720
      Top             =   0
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      DialogTitle     =   "Select A Help File"
      Flags           =   6148
   End
   Begin VB.Label Helpfilelabel 
      Caption         =   "&Help file:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Help"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private Sub Apply_Click()

If Contents.Value = True Then
   HelpFunction Me.hwnd, HELP_INDEX, ""
ElseIf Search.Value = True Then
   HelpFunction Me.hwnd, HELP_PARTIALKEY, ""
ElseIf Usinghelp.Value = True Then
   HelpFunction Me.hwnd, HELP_HELPONHELP, ""
Else
   SearchFor.Show 1 'show the search string dialog
End If
    
End Sub


Private Sub Browse_Click()

dialog.ShowOpen
filename.Text = dialog.filename

End Sub

Private Sub Cancel_Click()

Unload Me 'exit the program
            
End Sub




Public Sub Form_Load()
   
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
'When I have been creating a project for some reason, I have always needed
'to call the Windows help, wether it be the actuall Windows help files, or
'a help file of my own. This project uses a function from the API and allows
'four options. The user can execute the help file and see its contents,
'search the help file, see the help for help files, or can manually search
'for strings in the help file.
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



