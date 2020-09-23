Attribute VB_Name = "Module1"
Global fMainForm As frmMain
Global frmNewList As frmListDetail
Global iRcount As Integer
Global EntityDescription As String
Global option1Tag As String ' tag should equal fieldname at beginning of its selected case
Global option2Tag As String ' tag should equal fieldname at beginning of its selected case
Global entityTbl As String
Global latestentity As String
Global g_strUser As String
Global pwJubill As Boolean
Global permOk As Boolean
Global permResulted As Boolean
Global EntityMnu As String
Const PermAmount = 1



Sub Main()
    bAmDebugging = False ' change this if app hangs while debugging
    Dim fLogin As New frmLogin
    fLogin.Show vbModal
    If Not fLogin.OK Then
        'Login Failed so exit app
        End
    End If
    Unload fLogin

    frmSplash.Show
    frmSplash.Refresh
    Set fMainForm = New frmMain
    
    Load fMainForm
    Unload frmSplash

    'DEnv1.cn.Open
    With DEnv1.rsClientTable
        .CursorLocation = adUseClient
        .LockType = adLockPessimistic
    End With
    
    With DEnv1.rstblUserPermissions
        .CursorLocation = adUseClient
        .LockType = adLockPessimistic
    End With
    
    With DEnv1.rstblUsers
        .CursorLocation = adUseClient
        .LockType = adLockPessimistic
        
    End With
    
    With DEnv1.rstblAgents
        .CursorLocation = adUseClient
        .LockType = adLockPessimistic
    End With
    
    With DEnv1.rstblUserLog
        .CursorLocation = adUseClient
        .LockType = adLockPessimistic
    End With
    
    Set frmNewList = New frmListDetail
    
    fMainForm.Show
    
End Sub

Public Function PopulateFrm(Entity As String, sortBy As String, Ascending As Boolean, SearchBy As String)
    'sortBy = UCase(sortBy)

    Select Case Entity
    
    Case "Customers"
        option1Tag = "Account"
        option2Tag = "Customer"
        entityTbl = "ClientTable"
    
        latestentity = Entity

        If SearchBy <> "" Then
            On Error Resume Next
                DEnv1.rsClientTable.Close
            If sortBy = "option1" Then
                DEnv1.rsClientTable.Source = ("Select * From " & entityTbl & " Where " & option1Tag & " Like '" & SearchBy & "%'")
            Else
                DEnv1.rsClientTable.Source = ("Select * From " & entityTbl & " Where " & option2Tag & " Like '" & SearchBy & "%'")
            End If
            'DEnv1.rsClientTable.Requery
        Else
            On Error Resume Next
            DEnv1.rsClientTable.Close
            DEnv1.rsClientTable.Source = ("Select * From " & entityTbl)
            'DEnv1.rsClientTable.Requery
        End If
    
        DEnv1.rsClientTable.Open
        On Error Resume Next
        DEnv1.rsClientTable.MoveFirst
        iRcount = 0
    
        Do While Not DEnv1.rsClientTable.EOF
        iRcount = iRcount + 1
        DEnv1.rsClientTable.MoveNext
        Loop
            
            
            
    Case "Salesmen"
        option1Tag = "SalesID"
        option2Tag = "Name"
        entityTbl = "tblAgents"
        latestentity = Entity

    If SearchBy <> "" Then
        On Error Resume Next
        DEnv1.rstblAgents.Close
        If sortBy = "option1" Then
            DEnv1.rstblAgents.Source = ("Select * From " & entityTbl & " Where " & option1Tag & " Like '" & SearchBy & "%'")
        Else
            DEnv1.rstblAgents.Source = ("Select * From " & entityTbl & " Where " & option2Tag & " Like '" & SearchBy & "%'")
        'DEnv1.rstblAgents.Requery
        End If
        On Error Resume Next
        DEnv1.rstblAgents.Close
        DEnv1.rstblAgents.Source = ("Select * From " & entityTbl)
        'DEnv1.rstblAgents.Requery
    End If
    
    DEnv1.rstblAgents.Open
    On Error Resume Next
    DEnv1.rstblAgents.MoveFirst
    iRcount = 0
    
    Do While Not DEnv1.rstblAgents.EOF
        iRcount = iRcount + 1
        DEnv1.rstblAgents.MoveNext
    Loop
        
    
    Case "Users"
        option1Tag = "Username"
        option2Tag = "Fullname"
        entityTbl = "tblUsers"
        latestentity = Entity

    If SearchBy <> "" Then
        On Error Resume Next
        DEnv1.rstblUsers.Close
            If sortBy = "option1" Then
                DEnv1.rstblUsers.Source = ("Select * From " & entityTbl & " Where " & option1Tag & " Like '" & SearchBy & "%'")
            Else
                DEnv1.rstblUsers.Source = ("Select * From " & entityTbl & " Where " & option2Tag & " Like '" & SearchBy & "%'")
            End If
    Else
        On Error Resume Next
        DEnv1.rstblUsers.Close
        DEnv1.rstblUsers.Source = ("Select * From " & entityTbl)
        DEnv1.rstblUsers.Requery
    End If
    
    DEnv1.rstblUsers.Open
    On Error Resume Next
    DEnv1.rstblUsers.MoveFirst
    iRcount = 0
    
    Do While Not DEnv1.rstblUsers.EOF
        iRcount = iRcount + 1
        DEnv1.rstblUsers.MoveNext
    Loop
        
    End Select
                  
End Function

Public Function checkPermission(permId)
permOk = False
permResulted = False

    With DEnv1.rstblUserPermissions
        If .State = 0 Then
            .Open
        End If
        .MoveFirst
        
        Do Until permResulted = True
        On Error Resume Next
            If .Fields("username") <> g_strUser Then
                On Error Resume Next
                .MoveNext
                On Error Resume Next
                If .EOF Then
                    permOk = False
                    permResulted = True
                End If
            Else
                If .Fields("permissionid") = permId Then
                    If .Fields("value") = "Yes" Then
                        permOk = True
                    Else
                    permOk = False
                    End If
                    permResulted = True
                Else
                .MoveNext
                End If
            End If
        Loop
    End With
    
End Function
