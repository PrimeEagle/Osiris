VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHideFromTaskManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Declare Function RegisterServiceProcess Lib "KERNEL32" (ByVal ProcessID As Long, ByVal ServiceFlags As Long) As Long
Private Declare Function GetCurrentProcessId Lib "KERNEL32" () As Long

Public Sub Hide()
    RegisterServiceProcess GetCurrentProcessId, 1
End Sub

Public Sub Show()
    RegisterServiceProcess GetCurrentProcessId, 0
End Sub

