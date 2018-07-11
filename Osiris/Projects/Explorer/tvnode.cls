VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tvnode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public read_only As Boolean
Public table_name As String
Public quicktypeid As Long
Public icon_normal As String
Public icon_selected As String
Public create_node As Boolean
Public create_item As Boolean
Public system_node As Boolean
Public Created As Variant
Public created_by As String
Public last_modified As Variant
Public modified_by As String
Public Link_NodeID As Long
Public sublink As Boolean
