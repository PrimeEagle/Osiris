VERSION 4.00
Begin VB.Form SearchFor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help File Executor"
   ClientHeight    =   1335
   ClientLeft      =   2310
   ClientTop       =   2490
   ClientWidth     =   3375
   ClipControls    =   0   'False
   Height          =   1740
   Icon            =   "SearchFor.frx":0000
   Left            =   2250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   Top             =   2145
   Width           =   3495
   Begin VB.CommandButton Find 
      Caption         =   "&Find"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Searchstring 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Searchforlabel 
      Caption         =   "&Search for:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "SearchFor"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub Add_Click()

If Addnewlabel.Caption = "Add &new section:" Then 'if adding a new section
   WritePrivateProfileStringByKeyName Newtext.Text, "", "", Ini.Filename.Text
ElseIf Addnewlabel.Caption = "Add &new entry:" Then 'if adding a new entry
   WritePrivateProfileStringByKeyName Ini.Sectionslist.Text, Newtext.Text, Newvalue.Text, Ini.Filename.Text
ElseIf Addnewlabel.Caption = "Edit e&ntry:" Then 'if editing an entry then
   WritePrivateProfileStringByKeyName Ini.Sectionslist.Text, Newtext.Text, Newvalue.Text, Ini.Filename.Text
End If

'This code changes the filename text box so that the new edits that have
'been made can be updated in the program options.
'The "Ini.filename_change" procedure will not work here so this code is
'used instead.

A = Ini.Filename.Text
Ini.Filename.Text = ""
Ini.Filename.Text = A

Unload Me 'exit the add dialog

End Sub

Private Sub Cancel_Click()

Unload Me 'exit the add dialog

End Sub


Private Sub Find_Click()

HelpFunction Me.hwnd, HELP_PARTIALKEY, Searchstring.Text

End Sub

Private Sub Form_Load()

Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2 'centre the form on the screen

End Sub


