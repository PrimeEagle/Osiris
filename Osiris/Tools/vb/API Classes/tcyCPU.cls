VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLScpu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As TSYSTEM_INFO)

Private Type TSYSTEM_INFO
    dwOEMID                     As Long
    dwPageSize                  As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessormask       As Long
    dwNumberofProcessors        As Long
    dwProcessorType             As Long
    dwAllocationGranularity     As Long
    wProcessorLevel             As Integer
    wProcessorRevision          As Integer
End Type

Enum EProcessor
    is386 = 386
    is486 = 486
    isPentium = 586
    isMIPSR4000 = 4000
    isAlpha21064 = 21064
End Enum

'local variable(s) to hold property value(s)
Private mvarNumberOfProcessors        As Byte 'local copy
Private mvarBrandOfCPU                As String 'local copy
Private mvarOEMID                     As Long 'local copy
Private mvarPageSize                  As Long 'local copy
Private mvarMinimumApplicationAddress As Long 'local copy
Private mvarMaximumApplicationAddress As Long 'local copy
Private mvarAllocationGranularity     As Long 'local copy
Private mvarProcessorLevel            As Integer 'local copy
Private mvarProcessorRevision         As Integer 'local copy

Public Property Get ProcessorRevision() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProcessorRevision
    ProcessorRevision = mvarProcessorRevision
End Property



Public Property Get ProcessorLevel() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ProcessorLevel
    ProcessorLevel = mvarProcessorLevel
End Property


Public Property Get AllocationGranularity() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AllocationGranularity
    AllocationGranularity = mvarAllocationGranularity
End Property


Public Property Get MaximumApplicationAddress() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MaximumApplicationAddress
    MaximumApplicationAddress = mvarMaximumApplicationAddress
End Property



Public Property Get MinimumApplicationAddress() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MinimumApplicationAddress
    MinimumApplicationAddress = mvarMinimumApplicationAddress
End Property


Public Property Get PageSize() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PageSize
    PageSize = mvarPageSize
End Property


Public Property Get OEMID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.OEMID
    OEMID = mvarOEMID
End Property


Public Property Get BrandOfCPU() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BrandOfCPU
    BrandOfCPU = mvarBrandOfCPU
End Property


Public Property Get NumberOfProcessors() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NumberOfProcessors
    NumberOfProcessors = mvarNumberOfProcessors
End Property



Private Sub GatherInformation()
    Dim mySys  As TSYSTEM_INFO
    GetSystemInfo mySys
        
    mvarNumberOfProcessors = CStr(mySys.dwNumberofProcessors)
    
    Select Case mySys.dwProcessorType
        Case Is = is386
            mvarBrandOfCPU = "Intel 386"
            
        Case Is = is486
            mvarBrandOfCPU = "Intel 486"
        
        Case Is = isPentium
            mvarBrandOfCPU = "Intel Pentium"
        
        Case Is = isMIPSR4000
            mvarBrandOfCPU = "MIPS R4000"
            
        Case Is = isAlpha21064
            mvarBrandOfCPU = "Alpha 21064"
    End Select

    mvarOEMID = mySys.dwOEMID
    mvarPageSize = mySys.dwPageSize
    mvarMinimumApplicationAddress = mySys.lpMinimumApplicationAddress
    mvarMaximumApplicationAddress = mySys.lpMaximumApplicationAddress
    mvarAllocationGranularity = mySys.dwAllocationGranularity
    mvarProcessorLevel = mySys.wProcessorLevel
    mvarProcessorRevision = mySys.wProcessorRevision
End Sub

Private Sub Class_Initialize()
    GatherInformation
End Sub
