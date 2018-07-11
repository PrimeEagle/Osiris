VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChangeOwner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarMDAPath As String 'local copy
Private mvarMDWPath As String 'local copy
Private mvarMDBPath As String 'local copy
Private mvarFromUser As String 'local copy
Private mvarToUser As String 'local copy

Public Function Execute() As Boolean
    Dim mdb     As Database
    Dim U       As Container
    Dim C       As Container
    Dim D       As Document
    Dim D2      As Document
    Dim wrkJet  As Workspace

    Execute = False
    
    If (Dir$(mvarMDAPath, vbNormal) <> "") Then
        DBEngine.SystemDB = mvarMDAPath
        DBEngine.DefaultType = dbUseJet
            
        Set wrkJet = CreateWorkspace("JetWorkspace", "ew_user", "di", dbUseJet)
        Workspaces.Append wrkJet
        
        If (Dir$(mvarMDWPath, vbNormal) <> "") Then
            DBEngine.SystemDB = mvarMDWPath
            Set wrkJet = DBEngine.Workspaces(0)
            
            If (Dir$(mvarMDBPath, vbNormal) <> "") Then
                Set mdb = wrkJet.OpenDatabase(mvarMDBPath)
                        
                For Each C In mdb.Containers
                    Debug.Print C.Name & ": " & C.Owner
                    
                    If C.Owner = mvarFromUser Then
                        Set U = C
                        
                        U.UserName = mvarToUser
                        U.Owner = mvarToUser
                        U.Permissions = dbSecFullAccess Or dbSecWriteOwner Or dbSecWriteSec
                    End If
                Next
                
                Set C = mdb.Containers("Databases")
                Set D = C.Documents("MSysDb")
                D.UserName = mvarToUser
                D.Owner = mvarToUser
                D.Permissions = dbSecDBAdmin Or dbSecDBCreate Or dbSecDBOpen Or dbSecDBExclusive
                
                For Each D In mdb.Containers("Tables").Documents
                    If (D.Owner = mvarFromUser) Then
                        Set D2 = D
                    
                        On Error Resume Next
                        D2.UserName = mvarToUser
                        D2.Permissions = dbSecReadDef Or dbSecCreate Or dbSecWriteDef Or dbSecRetrieveData Or dbSecInsertData Or dbSecReplaceData Or dbSecDelete
                        D2.Owner = mvarToUser
                        Err.Clear
                    End If
                    
                    Debug.Print "Name: " & D.Name & " - " & D.Owner
                Next
                
                Execute = True
            End If
        End If
    End If
    
    Set U = Nothing
    Set mdb = Nothing
    Set wrkJet = Nothing
    Set C = Nothing
    Set D = Nothing
    Set D2 = Nothing
End Function

Public Property Let ToUser(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ToUser = 5
    mvarToUser = vData
End Property


Public Property Get ToUser() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ToUser
    ToUser = mvarToUser
End Property



Public Property Let FromUser(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FromUser = 5
    mvarFromUser = vData
End Property


Public Property Get FromUser() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FromUser
    FromUser = mvarFromUser
End Property



Public Property Let MDBPath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MDBPath = 5
    mvarMDBPath = vData
End Property


Public Property Get MDBPath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MDBPath
    MDBPath = mvarMDBPath
End Property



Public Property Let MDWPath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MDWPath = 5
    mvarMDWPath = vData
End Property


Public Property Get MDWPath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MDWPath
    MDWPath = mvarMDWPath
End Property



Public Property Let MDAPath(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MDAPath = 5
    mvarMDAPath = vData
End Property


Public Property Get MDAPath() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MDAPath
    MDAPath = mvarMDAPath
End Property



