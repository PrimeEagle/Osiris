VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private Type SYSTEM_INFO
    dwOEMID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessormask As Long
    dwNumberofProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type

'Default Property Values:
Const m_def_AvailablePhysical = 0
Const m_def_AvailableVirtual = 0
Const m_def_HighAddress = ""
Const m_def_LowAddress = ""
Const m_def_PageSize = 0
Const m_def_TotalPhysical = 0
Const m_def_TotalVirtual = 0

'Property Variables:
Dim m_AvailablePhysical As Long
Dim m_AvailableVirtual As Long
Dim m_HighAddress As String
Dim m_LowAddress As String
Dim m_PageSize As Long
Dim m_TotalPhysical As Long
Dim m_TotalVirtual As Long


Public Property Get AvailablePhysical() As Long
    AvailablePhysical = m_AvailablePhysical
End Property

Public Property Get AvailableVirtual() As Long
    AvailableVirtual = m_AvailableVirtual
End Property

Public Property Get HighAddress() As String
    HighAddress = m_HighAddress
End Property

Public Property Get LowAddress() As String
    LowAddress = m_LowAddress
End Property

Public Property Get PageSize() As Long
    PageSize = m_PageSize
End Property

Public Property Get TotalPhysical() As Long
    TotalPhysical = m_TotalPhysical
End Property

Public Property Get TotalVirtual() As Long
    TotalVirtual = m_TotalVirtual
End Property

Public Function Update() As Boolean
    Dim mySys  As SYSTEM_INFO
    Dim memsts As MEMORYSTATUS
    
    GlobalMemoryStatus memsts
    
    With memsts
        m_TotalPhysical = CStr(.dwTotalPhys)
        m_AvailablePhysical = CStr(.dwAvailPhys)
        m_TotalVirtual = CStr(.dwTotalVirtual)
        m_AvailableVirtual = CStr(.dwAvailVirtual)
    End With

    GetSystemInfo mySys
        
    m_PageSize = CStr(mySys.dwPageSize)
    m_LowAddress = "&H" & CStr(Hex(mySys.lpMinimumApplicationAddress))
    m_HighAddress = "&H" & CStr(Hex(mySys.lpMaximumApplicationAddress))
End Function

'Initialize Properties for User Control
Private Sub Class_Initialize()
    m_AvailablePhysical = m_def_AvailablePhysical
    m_AvailableVirtual = m_def_AvailableVirtual
    m_HighAddress = m_def_HighAddress
    m_LowAddress = m_def_LowAddress
    m_PageSize = m_def_PageSize
    m_TotalPhysical = m_def_TotalPhysical
    m_TotalVirtual = m_def_TotalVirtual
    
    Update
End Sub

