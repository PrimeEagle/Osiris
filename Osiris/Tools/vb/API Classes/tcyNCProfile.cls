VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tcyNCProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum TDetectionMethods
    vbManual = 0
    vbCOM1 = 1
    vbCOM2 = 2
    vbCOM3 = 3
    vbCOM4 = 4
    vbDialupNetworking = 5
    vbClass = 6
End Enum

Public Enum TIntervals
    vbPerHour = 0
    vbPerMinute = 1
    vbPerSecond = 2
End Enum

'local variable(s) to hold property value(s)
Private mvarDescription     As String 'local copy
Private mvarDetectionMethod As TDetectionMethods 'local copy
Private mvarFreeTime        As Long 'local copy
Private mvarWAVAlarm        As String 'local copy
Private mvarRangeTracking   As Boolean 'local copy
Private mvarRangeStart      As String 'local copy
Private mvarRangeStop       As String 'local copy
Private mvarRangeInterval   As Byte 'local copy
Private mvarRangePrice      As Currency 'local copy
Private mvarNormalPrice     As Currency 'local copy
Private mvarNormalInterval  As Byte 'local copy
Private mvarFileName        As String 'local copy
Private mvarToday           As Long 'local copy
Private mvarInRange         As Long 'local copy
Private mvarMonth           As Long 'local copy
Private mvarToDate          As Long 'local copy
Private mvarDetectionClass  As String 'local copy
Private mvarCurrent         As Long 'local copy

Public Property Let Current(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Current = 5
    mvarCurrent = vData
End Property


Public Property Get Current() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Current
    Current = mvarCurrent
End Property




Public Property Let DetectionClass(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DetectionClass = 5
    mvarDetectionClass = vData
End Property


Public Property Get DetectionClass() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DetectionClass
    DetectionClass = mvarDetectionClass
End Property




Public Property Let ToDate(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ToDate = 5
    mvarToDate = vData
End Property


Public Property Get ToDate() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ToDate
    ToDate = mvarToDate
End Property



Public Property Let Month(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Month = 5
    mvarMonth = vData
End Property


Public Property Get Month() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Month
    Month = mvarMonth
End Property



Public Property Let InRange(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.InRange = 5
    mvarInRange = vData
End Property


Public Property Get InRange() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.InRange
    InRange = mvarInRange
End Property



Public Property Let Today(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Today = 5
    mvarToday = vData
End Property


Public Property Get Today() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Today
    Today = mvarToday
End Property




Public Property Let FileName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FileName = 5
    mvarFileName = vData
End Property


Public Property Get FileName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FileName
    FileName = mvarFileName
End Property




Public Property Let NormalInterval(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NormalInterval = 5
    mvarNormalInterval = vData
End Property


Public Property Get NormalInterval() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NormalInterval
    NormalInterval = mvarNormalInterval
End Property



Public Property Let NormalPrice(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NormalPrice = 5
    mvarNormalPrice = vData
End Property


Public Property Get NormalPrice() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NormalPrice
    NormalPrice = mvarNormalPrice
End Property



Public Property Let RangePrice(ByVal vData As Currency)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RangePrice = 5
    mvarRangePrice = vData
End Property


Public Property Get RangePrice() As Currency
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RangePrice
    RangePrice = mvarRangePrice
End Property



Public Property Let RangeInterval(ByVal vData As Byte)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RangeInterval = 5
    mvarRangeInterval = vData
End Property


Public Property Get RangeInterval() As Byte
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RangeInterval
    RangeInterval = mvarRangeInterval
End Property



Public Property Let RangeStop(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RangeStop = 5
    mvarRangeStop = vData
End Property


Public Property Get RangeStop() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RangeStop
    RangeStop = mvarRangeStop
End Property



Public Property Let RangeStart(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RangeStart = 5
    mvarRangeStart = vData
End Property


Public Property Get RangeStart() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RangeStart
    RangeStart = mvarRangeStart
End Property



Public Property Let RangeTracking(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RangeTracking = 5
    mvarRangeTracking = vData
End Property


Public Property Get RangeTracking() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RangeTracking
    RangeTracking = mvarRangeTracking
End Property



Public Property Let WAVAlarm(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WAVAlarm = 5
    mvarWAVAlarm = vData
End Property


Public Property Get WAVAlarm() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WAVAlarm
    WAVAlarm = mvarWAVAlarm
End Property



Public Property Let FreeTime(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FreeTime = 5
    mvarFreeTime = vData
End Property


Public Property Get FreeTime() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FreeTime
    FreeTime = mvarFreeTime
End Property



Public Property Let DetectionMethod(ByVal vData As TDetectionMethods)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DetectionMethod = 5
    mvarDetectionMethod = vData
End Property


Public Property Get DetectionMethod() As TDetectionMethods
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DetectionMethod
    DetectionMethod = mvarDetectionMethod
End Property



Public Property Let Description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
End Property


Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property



Public Function Save() As Boolean
    Dim I As New tcyIniFile
    
    If mvarFileName = "" Then
        mvarFileName = App.Path & "\Default Provider.ncp"
    End If
    
    With I
        .FileName = mvarFileName
        .Section = "ProfileInformation"
        
        If mvarDescription = "" Then
            mvarDescription = "Default Provider"
        End If
        
        .Parameter = "Description":         .Value = mvarDescription
        .Parameter = "DetectionMethod":     .Value = CStr(mvarDetectionMethod)
        .Parameter = "FreeTime":            .Value = CStr(mvarFreeTime)
        .Parameter = "WAVAlarm":            .Value = mvarWAVAlarm
        .Parameter = "RangeTracking":       .Value = CStr(mvarRangeTracking)
        .Parameter = "RangeStart":          .Value = CStr(mvarRangeStart)
        .Parameter = "RangeStop":           .Value = CStr(mvarRangeStop)
        .Parameter = "RangeInterval":       .Value = CStr(mvarRangeInterval)
        .Parameter = "RangePrice":          .Value = CStr(mvarRangePrice)
        .Parameter = "NormalInterval":      .Value = CStr(mvarNormalInterval)
        
        .Parameter = "Today":               .Value = CStr(mvarToday)
        .Parameter = "Month":               .Value = CStr(mvarMonth)
        .Parameter = "InRange":             .Value = CStr(mvarInRange)
        .Parameter = "ToDate":              .Value = CStr(mvarToDate)
    End With
    
    Set I = Nothing
End Function

Public Function Load() As Boolean
    Dim I As New tcyIniFile
    
    If mvarFileName = "" Then
        mvarFileName = App.Path & "\Default Provider.ncp"
    End If
    
    With I
        .FileName = mvarFileName
        .Section = "ProfileInformation"
        
        .Default = "Default Provider"
        .Parameter = "Description":         mvarDescription = .Value
        
        .Default = CStr(vbManual)
        .Parameter = "DetectionMethod":     mvarDetectionMethod = .Value
        
        .Default = "#32770"
        .Parameter = "DetectionClass":     mvarDetectionClass = .Value
        
        .Default = "0"
        .Parameter = "FreeTime":            mvarFreeTime = .Value
        .Parameter = "WAVAlarm":            mvarWAVAlarm = .Value
        
        .Default = CStr(False)
        .Parameter = "RangeTracking":       mvarRangeTracking = CBool(.Value)
        
        .Default = "0"
        .Parameter = "RangeStart":          mvarRangeStart = CDate(.Value)
        
        .Default = "0"
        .Parameter = "RangeStop":           mvarRangeStop = CDate(.Value)
        
        .Default = CStr(vbPerHour)
        .Parameter = "RangeInterval":       mvarRangeInterval = CByte(.Value)
        
        .Default = "$0.00"
        .Parameter = "RangePrice":          mvarRangePrice = CCur(.Value)
        
        .Default = CStr(vbPerHour)
        .Parameter = "NormalInterval":      mvarNormalInterval = CByte(.Value)
        
        .Default = "$0.00"
        .Parameter = "NormalPrice":         mvarNormalPrice = CCur(.Value)
        
        mvarCurrent = 0
        
        .Default = "0"
        .Parameter = "Today":               mvarToday = CLng(.Value)
        .Parameter = "Month":               mvarMonth = CLng(.Value)
        .Parameter = "InRange":             mvarInRange = CLng(.Value)
        .Parameter = "ToDate":              mvarToDate = CLng(.Value)
    End With
    
    Load = True
    
    Set I = Nothing
End Function

Private Sub Class_Initialize()
    mvarFileName = App.Path & "\Default Provider.ncp"
    
    mvarDescription = ""
    
    mvarDetectionMethod = vbManual
    mvarDetectionClass = "#32770"
    
    mvarFreeTime = 0
    mvarWAVAlarm = ""
    
    mvarRangeTracking = False
    mvarRangeStart = 0
    mvarRangeStop = 0
    mvarRangeInterval = vbPerHour
    mvarRangePrice = 0#
    
    mvarNormalInterval = vbPerHour
    mvarNormalPrice = 0#
    
    mvarCurrent = 0
    mvarToday = 0
    mvarMonth = 0
    mvarInRange = 0
End Sub

Private Sub Class_Terminate()
    Save
End Sub
