VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FrmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impulse Sign in"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4005
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frmlogin.frx":0000
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   4005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2550
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4965
      Width           =   1305
   End
   Begin VB.TextBox Pass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   105
      PasswordChar    =   "l"
      TabIndex        =   0
      Top             =   4425
      Width           =   3750
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sign in"
      Default         =   -1  'True
      Height          =   390
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4950
      Width           =   1305
   End
   Begin VB.Label lblChangePassword 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   105
      MouseIcon       =   "Frmlogin.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   5610
      Width           =   1725
   End
   Begin VB.Image ImgUser 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3240
      Left            =   105
      Picture         =   "Frmlogin.frx":0594
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3750
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   375
      Left            =   105
      TabIndex        =   3
      Top             =   3750
      Width           =   3750
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "6615;661"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Password :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   5
      Top             =   4200
      Width           =   900
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Username :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   3510
      Width           =   960
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LoginSucceeded As Boolean

Dim bMouseDown As Boolean
Const lCaptionHeight = 600
Dim lDownX As Long, lDownY As Long
Dim iColNo As Integer
Dim iColNoTemp As Integer
Dim bSaved As Boolean
Dim lEntryID As Long

Private Sub DrawBorder(Pic As PictureBox)

    Dim lBorderColor As Long
    lBorderColor = RGB(0, 200, 0)    'RGB(69, 140, 201) 'RGB(69, 140, 201)
    Pic.Line (0, 0)-(Pic.Width, 0), lBorderColor 'Top
    Pic.Line (0, Pic.Height - 85)-(Pic.Width, Pic.Height - 85), lBorderColor 'Bottom
    Pic.Line (0, 0)-(0, Pic.Height), lBorderColor    'Left
    Pic.Line (Pic.Width - 85, 0)-(Pic.Width - 85, Pic.Height), lBorderColor  'Right
    
End Sub



Private Sub cmdCancel_Click()

    If MsgBox("Are You Sure To Quit ?", vbYesNo + vbQuestion) = vbYes Then
        con.Close
        Set con = Nothing
        Unload FrmLogin
        Unload MainForm
        End
    End If
    
End Sub

Private Sub cmdImport_Click()
    
    Exit Sub
    Dim ConAcc As New ADODB.Connection, rsAcc As New ADODB.Recordset
    ConAcc.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Custom Softwares\SMBI\Surgical Industry\db\db_qel_dent.mdb;Persist Security Info=False"
    
    Dim rs As New ADODB.Recordset
    Dim rsPic As New ADODB.Stream
    Dim strSize As String, lSize As Long
    With rsAcc
        .Open "SELECT * FROM tbl_prods", con, adOpenDynamic, adLockOptimistic
        Do Until .EOF
            rs.Open "SELECT * FROM Items WHERE ItemID='" & !pModel & "'", con, adOpenDynamic, adLockOptimistic
            If rs.RecordCount > 0 Then GoTo MoveNEXT:
            rs.AddNew
            rs!CatID = 52
            rs!ITemID = !pModel & ""
            rs!ItemName = !pName & ""
            strSize = Trim(!PSize & "")
            lSize = Val(strSize)
            rs!ItemSize = lSize
            rs!SizeUnit = Mid(!PSize & "", Len(CStr(lSize)) + 1)
            On Error Resume Next
            rs!Attributes = !Auth & ""
            If err.Number = -2147217887 Then
                rs!Attributes = Trim(Mid(!Auth & "", 2))
                On Error GoTo 0
            End If
            On Error Resume Next
            rs!PackingInstructions = !Pdetail & ""
            If err.Number = -2147217887 Then
                rs!PackingInstructions = Trim(Mid(!Pdetail & "", 2))
                On Error GoTo 0
            End If
            If Not IsNull(!sImg) Then
                rsPic.Open
                rsPic.Type = adTypeBinary
                On Error Resume Next
                rsPic.LoadFromFile "D:\Images\pimgs_dent\" & !sImg
                If err.Number = 3002 Then GoTo skipPicS
                rs.Fields("ItemPic") = rsPic.Read
                
skipPicS:
                err.Clear
                On Error GoTo 0
                rsPic.Close

            End If
             If Not IsNull(!lImg) Then
                rsPic.Open
                rsPic.Type = adTypeBinary
                On Error Resume Next
                rsPic.LoadFromFile "D:\Images\pimgs_dent\" & !lImg
                If err.Number = 3002 Then GoTo skipPic
                rs.Fields("ItemPicOther") = rsPic.Read
skipPic:
                err.Clear
                On Error GoTo 0
                rsPic.Close
            End If
            rs!Inactive = False
            rs!FOB = 0
            rs!CreateDT = Now
            rs!GroupID = !SecID & ""
            rs.Update
            
MoveNEXT:
            rs.Close
            Debug.Print !pModel & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing

End Sub

Private Sub cmdOK_Click()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    rs.Open "Select * From DateRange", con, adOpenForwardOnly, adLockReadOnly
    If rs.EOF = False Then
        YearStartDate = rs!StartDate & ""
        YearLastDate = rs!LastDate & ""
    End If
    rs.Close
    
    Dim strLen As Long
    strLen = 255
    strComputerName = Space$(255)
    Call GetComputerName(strComputerName, strLen)
    strComputerName = Left(strComputerName, strLen)
    
    With rs
        .Open "SELECT ComputerName FROM UrduSupportedSystems WHERE ComputerName='" & strComputerName & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            bUrduSupported = False
        Else
            bUrduSupported = True
        End If
        .Close
    End With
    
'    If Command = "FromVBIDE" And Combo1.Text = "Administrator" Then
'        UserID = 1
'        CurrentUserName = "Administrator"
'        Call SaveSetting(App.Title, "UserLevelSettings", "UserName", Combo1.Text)
'        Call DoStartUpActivities(1)
'        Screen.MousePointer = vbDefault
'        HicoVisible = True
'        Unload Me
'        GoTo ComeHere
'    End If
    
    Dim FullUserName As String, strPassword As String
    'check for correct password
    With rs
        .Open "SELECT UserID,Password,FullUserName,FinancialMainLink,StockMainLink,PasswordEncrypted FROM Users WHERE UserName='" & Combo1.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            strPassword = !Password & ""
            If !PasswordEncrypted Then
                strPassword = CryptRC4(FromHexDump(strPassword), "AwmWrangler")
            Else
                strPassword = !Password & ""
            End If
        End If
        
        'If UCase(Pass & "") = UCase(![Password] & "") Then
        If Pass & "" = strPassword Then
            'Check Password Complexity...
            If CheckPasswordComplexity(Val(!UserID & ""), !Password & "", 0) = False Then
                Exit Sub
            End If
            
            Screen.MousePointer = vbHourglass
            Me.Hide
            DoEvents
            CurrentUserName = Combo1
            'UserReqAuth = Abs(![ReqAuthentication]) 'Whether Or Not User Required Authorization ON Issue and Receive.
            'MainForm.mnuSubOptions(5).Caption = "Log Off " & CurrentUserName & " ..."
            Me.Hide
            UserID = ![UserID]
            FullUserName = !FullUserName & ""
             
            Call SaveSetting(App.Title, "UserLevelSettings", "UserName", Combo1.Text)
            
            Call DoStartUpActivities(![UserID])
            Screen.MousePointer = vbDefault
            'HicoVisible = ![HicoVisible]
            .Close
        
            'Unload Me
        Else
            MsgBox "Cannot Login Due To Invalid Password, Note That Password is Case Sensitive." & vbNewLine & _
             "Make Sure That Caps Lock Is Not Accidently ON.", vbCritical, "Login"
         
            Pass.SelStart = 0
            Pass.SelLength = Len(Pass)
            Pass.SetFocus
            Exit Sub
        End If
        
    End With
    Set rs = Nothing
        
    Set MainForm.cmdMin.Container = MainForm.MainToolbar
    
    MainForm.cmdMin.Move MainForm.MainToolbar.Buttons("PlaceHolder1").Left - 60, MainForm.MainToolbar.Buttons("PlaceHolder1").Top + 300
    MainForm.cmdMin.Visible = True
    MainForm.lblUserName.Caption = UCase(CurrentUserName)
    
ComeHere:
    'MainForm.Label1.Caption = "Logged on as " & vbNewLine & CurrentUserName '& vbNewLine & "Click Here To Log Off."
    'MainForm.Label1.Caption = "Logged on : " & CurrentUserName '& vbNewLine & "Click Here To Log Off."
    
    If (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
        'Show Alerts... Userwise
        Call MainForm.ShowPurchaseCalendarAlerts
        MainForm.TmrAlerts.Enabled = True
    End If
    
    'CreateExplorerBarForExport
    
'    If App.Title = "HicoExport" Then
'        Call
'    ElseIf App.Title = "Star" Then
'        Call CreateExplorerBarForStar
'    End If

'    Dim f As New frmUserDetails
'    Load f
'    f.lblUserName = FullUserName
'    Set f.ImgUser.Picture = ImgUser.Picture
'
'    f.ShowMe
    
    Unload Me
    
    Exit Sub
    
err:
    MsgBox err.Description
End Sub


Private Sub Combo1_Click()
    
    On Error GoTo err
    
    UserName = Combo1.Text
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select UserName,UserPic From Users Where UserName='" & Combo1.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            If Not IsNull(![UserPic]) Then
                Set ImgUser.DataSource = rs
                ImgUser.DataField = "UserPic"
            Else
                ImgUser.Picture = LoadPicture(App.Path & "\login.jpg")
            End If
        End If
        .Close
    End With
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description

End Sub


Private Sub Form_Load()
    
'    Dim hRegion As Long
'    hRegion = CreateRoundRectRgn(0, 0, Me.Width / 15, Me.Height / 15, 20, 20)
'    SetWindowRgn Me.hwnd, hRegion, True
    
    Call AddToCombo(Combo1, "UserName", "Users", "WHERE InActive=0 ORDER BY UserName")
    'If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
        
    Dim strSavedUserName As String
    strSavedUserName = GetSetting(App.Title, "UserLevelSettings", "UserName", "")
    If strSavedUserName <> "" Then
        On Error Resume Next
        Combo1.Text = strSavedUserName
        If err.Number = 380 Then
            Combo1.ListIndex = -1
            On Error GoTo 0
        End If
    End If
    
    Call SetToolTip(cmdOk.hwnd, "Logs You On To The Main Database.")
    
    
    
    'Call DrawPicCaptions
    
End Sub


Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Call cmdCancel_Click
        cancel = True
    End If
End Sub

Private Sub Form_Unload(cancel As Integer)
    OurOwnForm = True
    Unload Me
End Sub

Private Sub DoStartUpActivities(UserID As Long)
    
    On Error GoTo err
    
    'If Not OutOfFinancialYear Then Call CheckForAutoVouchers
    Set rsUserSettings = New ADODB.Recordset
    With rsUserSettings
        '.Open "Select * From VExportUsers Where UserID=" & UserID, con, adOpenForwardOnly, adLockReadOnly
        .Open "SELECT * FROM Users WHERE UserID=" & UserID, con, adOpenForwardOnly, adLockReadOnly
        If strCompany = "Weldon Industries" Or strCompany = "All Seas" Then
            'Hide Following Links for Weldon Industries...
            'Show them for Paksmith & Weldon Instruments
            Dim iExportIndex As Integer, iFinancialIndex As Integer, iPayrollIndex As Integer
            
            iExportIndex = MainForm.MainToolbar.Buttons("Export").Index
            MainForm.MainToolbar.Buttons(iExportIndex).Visible = False
            MainForm.MainToolbar.Buttons(iExportIndex + 1).Visible = False
                
            iFinancialIndex = MainForm.MainToolbar.Buttons("Financial").Index
            MainForm.MainToolbar.Buttons(iFinancialIndex).Visible = False
            MainForm.MainToolbar.Buttons(iFinancialIndex + 1).Visible = False
        
            iPayrollIndex = MainForm.MainToolbar.Buttons("Payroll").Index
            MainForm.MainToolbar.Buttons(iPayrollIndex).Visible = False
            MainForm.MainToolbar.Buttons(iPayrollIndex + 1).Visible = False
            
        End If
        
        If strCompany = "All Seas" Then
            Dim iStockIndex As Integer
            iStockIndex = MainForm.MainToolbar.Buttons("Stock").Index
            MainForm.MainToolbar.Buttons(iStockIndex).Visible = False
            MainForm.MainToolbar.Buttons(iStockIndex + 1).Visible = False
        End If
        MainForm.MainToolbar.Buttons("Financial").Enabled = !FinancialMainLink
        MainForm.MainToolbar.Buttons("Export").Enabled = !ExportMainLink
        MainForm.MainToolbar.Buttons("Payroll").Enabled = !PayrollMainLink
        MainForm.MainToolbar.Buttons("Stock").Enabled = !StockMainLink
        MainForm.MainToolbar.Buttons("Production").Enabled = IIf(IsNull(!ProductionMainLink), False, !ProductionMainLink)
        MainForm.MainToolbar.Buttons("Company").Enabled = IIf(IsNull(!CompanyMainLink), False, !CompanyMainLink)
        
        MainForm.MainToolbar.Buttons("Dashboard").Enabled = IIf(IsNull(!DashBoardMainLink), False, !DashBoardMainLink)
        
        MainForm.MainToolbar.Buttons("QMS").Enabled = IIf(IsNull(!QMSMainLink), False, !QMSMainLink)
        MainForm.MainToolbar.Buttons("FixedAssets").Enabled = IIf(IsNull(!FixedAssetsMainLink), False, !FixedAssetsMainLink)
'        MainForm.MainToolbar.Buttons("Dashboard").Visible = IIf(IsNull(!DashBoardMainLink), False, !DashBoardMainLink)
'        MainForm.MainToolbar.Buttons("DashboardDash").Visible = IIf(IsNull(!DashBoardMainLink), False, !DashBoardMainLink)
        If !OpenCommandCenter Then
            Load frmDashBoardCC
            frmDashBoardCC.ShowMe
        End If
    End With
    
    Dim i As Integer
    For i = 1 To MainForm.MainToolbar.Buttons.count
        If MainForm.MainToolbar.Buttons(i).Enabled And MainForm.MainToolbar.Buttons(i).Visible Then
            Call MainForm.MainToolbar_ButtonClick(MainForm.MainToolbar.Buttons(i))
            'frmMDIBG.Move (Me.ScaleWidth - frmMDIBG.ScaleWidth) / 2, (Me.ScaleHeight - frmMDIBG.ScaleHeight) / 2
            Exit For
        End If
    Next
    
    If Not (strCompany = "Tecno" Or strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne") Then
        Dim iQAIndex As Integer
        iQAIndex = MainForm.MainToolbar.Buttons("QMS").Index
        MainForm.MainToolbar.Buttons(iQAIndex).Visible = False
        MainForm.MainToolbar.Buttons(iQAIndex + 1).Visible = False
    End If
    
    Call LoadUserSettings
    Call MainForm.ApplySecuritySettings(UserID)
     
    MainForm.MainToolbar.Width = MainForm.MainToolbar.Buttons("PlaceHolder1").Left + MainForm.MainToolbar.Buttons("PlaceHolder1").Width
    MainForm.FraWelcome.Left = MainForm.MainToolbar.Left + MainForm.MainToolbar.Width
    
    Call CreateExplorerBarForFavourites
    'Show User Favourites...
    'Call LoadUserFavourites
    'bPrintWhenSaved = Abs(CBool(getGeneralDataValue("PrintWhenSavedVouchers")))
'    If rsUserSettings!ShowTips Then
'        frmTipViewer.Show
'    End If
    Call EncryptSensitiveData
    
    Exit Sub
err:
    MsgBox err.Description
    'Resume Next
End Sub

Private Sub CheckforUtilityBills()

    With con.Execute("Select Count(*) From UBPymts Where DueDT<=#" & Format(Date, "dd-MMM-yyyy") & "# and isnull(Vchrno)")
        If Val(.Fields(0) & "") > 0 Then
            If MsgBox("There Are " & Val(.Fields(0)) & " Utility Bill(s) Payable, Do You Want To Get Details ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            frmBillList.Show
            'Call frmBillList.cmdRefresh_Click
        End If
    End With
    
End Sub


Private Sub lblChangePassword_Click()

    If Combo1.ListIndex = -1 Then
        MsgBox "Select your user name from the List and try again.", vbInformation
        Exit Sub
    End If
    
    Dim bCanChangePassword As Boolean
    bCanChangePassword = GetSingleBooleanValue("ChangePassword", "Users", " WHERE UserName='" & Combo1.Text & "'")
    
    If bCanChangePassword = False Then
        MsgBox "User Can't change password.", vbInformation
        Exit Sub
    End If
    
    Dim f As New frmUserInfo
    Load f
    f.Tag = Combo1.Text
    f.Show_Option = Opt_ChangeUserPassword
    f.Show 1
    
End Sub
