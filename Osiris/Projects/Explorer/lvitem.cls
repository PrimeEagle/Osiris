VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "lvitem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public parent_node As Long
Public read_only As Boolean
Public data_type As String
Public icon_large As String
Public icon_small As String
Public created As Variant
Public created_by As String
Public last_modified As Variant
Public modified_by As String
Public variation As Boolean

