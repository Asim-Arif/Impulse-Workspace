Attribute VB_Name = "UserManagement"
Dim CurrentUser As New ADODB.Recordset
Dim CurrentUserEX As New ADODB.Recordset

Dim UserID As Long

Public Sub LoadUserSettings()
    With CurrentUser
        If .State = 1 Then .Close
        .Open "Select * From Users Where UserName='" & CurrentUserName & "'", con, adOpenForwardOnly, adLockReadOnly
    End With
End Sub

Public Function UserHasAccess(AccessOption As String, Optional ArgUserID As Long) As Boolean

    With CurrentUser
       ' If No User Setings Loaded Then Load Them First
       If UserID <> ArgUserID And .State = 1 Then .Close
       If .State = 0 Then Call LoadUserSettings
       ' If No Current User Then Exit Function
       If .RecordCount = 0 Then UserHasAccess = False: Exit Function
            UserID = ArgUserID
            'UserHasAccess = IIf(IsNull(.Fields(AccessOption)), False, .Fields(AccessOption))
            UserHasAccess = True
    End With
    
End Function



Public Sub LoadUserSettingsEX()
    With CurrentUserEX
        If .State = 1 Then .Close
        .Open "SELECT * FROM VUserMenuOptions WHERE UserName='" & CurrentUserName & "'", con, adOpenForwardOnly, adLockReadOnly
    End With
End Sub


Public Function UserHasAccessEX(AccessOption As String, Optional ArgUserID As Long) As Boolean
    
    If CurrentUserName = "Administrator" Then
        UserHasAccessEX = True
        Exit Function
    End If
    With CurrentUserEX
       ' If No User Setings Loaded Then Load Them First
       If UserID <> ArgUserID And .State = 1 Then .Close
       
       If .State = 0 Then Call LoadUserSettingsEX
       
       On Error Resume Next
       .Filter = ""
       .Filter = "OptionID = '" & AccessOption & "'" ', , adSearchForward, 0
       
       ' If No Current User Then Exit Function
        If .RecordCount > 0 Then
            UserHasAccessEX = True
            Exit Function
        Else
            UserHasAccessEX = False
            Exit Function
        End If
            UserID = ArgUserID
            'UserHasAccess = IIf(IsNull(.Fields(AccessOption)), False, .Fields(AccessOption))
            UserHasAccessEX = True
    End With
    
    
End Function

Public Function CheckPasswordComplexity(p_lUserID As Long, strPassword As String, iOperation As Integer) As Boolean
    'iOperation=0 means Login...
    'iOperation=1 means Password Change...
    'iOperation=2 means New User...
    Dim bEnforceComplexity As Boolean, iMinLength As Integer, ExpiryDT As Date, iExpireDays As Integer
    Dim lCount As Long
    If iOperation = 0 Then
    
        iExpireDays = Val(getGeneralDataValue("Passwords_ExpireAfter"))
        If iExpireDays > 0 Then
            ExpiryDT = GetSingleDateValue("PasswordChangedDT", "Users", " WHERE UserID=" & p_lUserID)
            ExpiryDT = DateAdd("d", iExpireDays, ExpiryDT)
            If GetServerDate(False) > ExpiryDT Then
                MsgBox "Your Password has expired, Please change your password.", vbInformation
                Exit Function
            End If
        End If
        Dim bRestrictSystems As Boolean
        bRestrictSystems = CBool(getGeneralDataValue("Login_RestrictedSystems"))
        
        If bRestrictSystems And p_lUserID <> 1 Then
            lCount = GetSingleLongValue("COUNT(*)", "Users_AllowedComputers", " WHERE UserID=" & p_lUserID & " AND MachineName='" & strComputerName & "'")
            If lCount = 0 Then
                MsgBox "You are not allowed to log in to this Computer, Please contact Administrator.", vbInformation
                Exit Function
            End If
        End If
        
    End If
    
    iMinLength = Val(getGeneralDataValue("Passwords_MinLength"))
    If iMinLength > 0 Then
        If Len(strPassword) < iMinLength Then
            MsgBox "Your Password must have atleast " & iMinLength & ". Please Change Your Password.", vbInformation
            Exit Function
        End If
    End If
    
    'Passwords_EnforceComplexity: means It should have atleast 1 Caps & atleast 1 UnCap Letter.
    bEnforceComplexity = CBool(getGeneralDataValue("Passwords_EnforceComplexity"))
    If bEnforceComplexity Then
        Dim i As Integer, strChar As String, bFoundLower As Boolean, bFoundUpper As Boolean, iAscVal As Integer, bFoundNumeric As Boolean
        For i = 1 To Len(strPassword)
            strChar = Mid(strPassword, i, 1)
            iAscVal = Asc(strChar)
            If iAscVal >= 65 And iAscVal <= 90 Then
                bFoundLower = True
            ElseIf iAscVal >= 97 And iAscVal <= 122 Then
                bFoundUpper = True
            ElseIf iAscVal >= 48 And iAscVal <= 57 Then
                bFoundNumeric = True
            End If
            If bFoundLower And bFoundUpper And bFoundNumeric Then
                Exit For
            End If
        Next
        
        If Not (bFoundLower And bFoundUpper And bFoundNumeric) Then
            MsgBox "Your Password must be alphanumeric & should contain atleast 1 Upper & Lower Character. Please Change Your Password", vbInformation
            Exit Function
        End If
        
    End If
    
    If iOperation = 1 Then
    
        Dim bRepeatPrevious As Boolean
        bRepeatPrevious = CBool(getGeneralDataValue("Passwords_RepeatPrevious"))
        If bRepeatPrevious = False Then
            'If Changing Passwords, Check for Previous Passwords
            lCount = GetSingleLongValue("COUNT(*)", "Users_OldPasswords", " WHERE UserID=" & p_lUserID & " AND Password='" & strPassword & "'")
            If lCount > 0 Then
                MsgBox "You can not repeat your previous password, Please Change your password.", vbInformation
                Exit Function
            End If
        End If
        
    End If
    
    CheckPasswordComplexity = True
    
End Function

Public Function GetUserRightAccess(strOption As String) As Boolean

    Dim bAccess As Boolean
    bAccess = GetSingleBooleanValue(strOption, "Users", " WHERE UserName='" & CurrentUserName & "'")
    
    GetUserRightAccess = bAccess
    If bAccess = False Then
        MsgBox "You don't have access.", vbInformation
        Exit Function
    End If
    
End Function
