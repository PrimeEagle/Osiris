Attribute VB_Name = "mSubclass"
Option Explicit

Type Instances
    in_use As Boolean 'This instance is alive
    ClassAddr As Long 'Pointer to self
    hwnd As Long 'hWnd being hooked
    PrevWndProc As Long 'Stored for unhooking
End Type

Public Const MIN_INSTANCES = 1
Public Const MAX_INSTANCES = 256

Public m_MyInstance As Integer
Global Instances(MIN_INSTANCES To MAX_INSTANCES) As Instances



'Hooks a window or acts as if it does if the window is
'already hooked by a previous instance of myUC.
Public Sub Hook_Window(ByVal hwnd As Long, ByVal instance_ndx As Integer)
        
        Instances(instance_ndx).PrevWndProc = Is_Hooked(hwnd)
        If Instances(instance_ndx).PrevWndProc = 0& Then
            Instances(instance_ndx).PrevWndProc = SetWindowLong(hwnd, _
                GWL_WNDPROC, AddressOf SwitchBoard)
        End If
        Instances(instance_ndx).hwnd = hwnd
        
End Sub
' Unhooks only if no other instances need the hWnd
Public Sub UnHookWindow(ByVal instance_ndx As Integer)
        If TimesHooked(Instances(instance_ndx).hwnd) = 1 Then
        SetWindowLong Instances(instance_ndx).hwnd, GWL_WNDPROC, _
            Instances(instance_ndx).PrevWndProc
        End If
        Instances(instance_ndx).hwnd = 0&
End Sub
'Determine if we have already hooked a window,
'and returns the PrevWndProc if true, 0& if false
Public Function Is_Hooked(ByVal hwnd As Long) As Long
        
        Dim ndx As Integer
        Is_Hooked = 0&
        For ndx = MIN_INSTANCES To MAX_INSTANCES
            If Instances(ndx).hwnd = hwnd Then
                Is_Hooked = Instances(ndx).PrevWndProc
                Exit For
            End If
        Next ndx
        
End Function
'Returns a count of the number of times a given
'window has been hooked by instances of myUC.
Public Function TimesHooked(ByVal hwnd As Long) As Long
        Dim ndx As Integer
        Dim cnt As Integer
        
        For ndx = MIN_INSTANCES To MAX_INSTANCES
            If Instances(ndx).hwnd = hwnd Then
                cnt = cnt + 1
            End If
        Next ndx
        TimesHooked = cnt
End Function


