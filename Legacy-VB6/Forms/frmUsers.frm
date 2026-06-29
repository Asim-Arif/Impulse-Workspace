VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Manager"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   Icon            =   "frmUsers.frx":0000
   LinkTopic       =   "Form24"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7185
      Left            =   105
      TabIndex        =   7
      Top             =   90
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   12674
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      BackColor       =   -2147483644
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "User List"
      TabPicture(0)   =   "frmUsers.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdNew"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdDelete"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdCopy"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdMarkInactive"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdPrint"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "LV"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "List1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Opts(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Opts(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Opts(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtSearch"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   345
         Left            =   240
         TabIndex        =   18
         Tag             =   "Type Username to Search"
         Text            =   "Type Username to Search"
         Top             =   4890
         Width           =   3345
      End
      Begin VB.OptionButton Opts 
         Caption         =   "All"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   225
         TabIndex        =   16
         Top             =   5445
         Width           =   1035
      End
      Begin VB.OptionButton Opts 
         Caption         =   "Active"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   1380
         TabIndex        =   15
         Top             =   5445
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Opts 
         Caption         =   "In-Active"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   2535
         TabIndex        =   14
         Top             =   5460
         Width           =   1185
      End
      Begin VB.Frame Frame1 
         Caption         =   "Settings for default"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Left            =   210
         TabIndex        =   10
         Top             =   5745
         Width           =   4815
         Begin MSForms.CommandButton cmdSettings 
            Height          =   375
            Left            =   2370
            TabIndex        =   5
            Top             =   810
            Width           =   1515
            Caption         =   "Change Settings..."
            Size            =   "2672;661"
            MousePointer    =   99
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdPass 
            Height          =   375
            Left            =   750
            TabIndex        =   4
            Top             =   810
            Width           =   1515
            Caption         =   "Set Password..."
            Size            =   "2672;661"
            MousePointer    =   99
            Accelerator     =   83
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Label Label3 
            Caption         =   "Use these buttons to specify a password or to setup the user rights for this project."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   750
            TabIndex        =   11
            Top             =   330
            Width           =   3885
         End
         Begin VB.Image Image2 
            Height          =   675
            Left            =   150
            Picture         =   "frmUsers.frx":045E
            Top             =   300
            Width           =   510
         End
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Left            =   4680
         TabIndex        =   0
         Top             =   2700
         Visible         =   0   'False
         Width           =   3285
      End
      Begin MSComctlLib.ListView LV 
         Height          =   3345
         Left            =   240
         TabIndex        =   17
         Top             =   1560
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   5900
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "User Name"
            Object.Width           =   5424
         EndProperty
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   375
         Left            =   3630
         TabIndex        =   13
         Top             =   3525
         Width           =   1425
         Caption         =   "Print"
         Size            =   "2514;661"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdMarkInactive 
         Height          =   375
         Left            =   3630
         TabIndex        =   12
         Top             =   2565
         Width           =   1425
         Caption         =   "Inactive"
         Size            =   "2514;661"
         Accelerator     =   73
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdCopy 
         Height          =   375
         Left            =   3630
         TabIndex        =   3
         Top             =   3045
         Width           =   1425
         Caption         =   "Make a Copy"
         Size            =   "2514;661"
         Accelerator     =   77
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdDelete 
         Height          =   375
         Left            =   3630
         TabIndex        =   2
         Top             =   2070
         Width           =   1425
         Caption         =   "Delete"
         Size            =   "2514;661"
         Accelerator     =   68
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdNew 
         Height          =   375
         Left            =   3630
         TabIndex        =   1
         Top             =   1590
         Width           =   1425
         ForeColor       =   0
         Caption         =   "New User..."
         PicturePosition =   327683
         Size            =   "2514;661"
         Accelerator     =   78
         MouseIcon       =   "frmUsers.frx":08C9
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label2 
         Caption         =   "&Users"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   $"frmUsers.frx":0BE3
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   870
         TabIndex        =   8
         Top             =   510
         Width           =   4095
      End
      Begin VB.Image Image1 
         Height          =   675
         Left            =   240
         Picture         =   "frmUsers.frx":0C96
         Top             =   450
         Width           =   540
      End
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   405
      Left            =   3870
      TabIndex        =   6
      Top             =   7380
      Width           =   1470
      ForeColor       =   0
      Caption         =   "Close"
      Size            =   "2593;714"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Rec As New ADODB.Recordset

Private Sub cmdClose_Click()
 Unload Me
End Sub


Private Sub cmdCopy_Click()
    
    On Error GoTo err
    Dim NewUserName As String
    NewUserName = InputBox("Please Enter New User Name :", "Copy User", "Copy Of " & List1.Text)
    
    If NewUserName = "" Then Exit Sub
    
    Call StartTrans(con)
    Dim cmd As New ADODB.Command
    With cmd
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_CopyUser"
        .Parameters("@FromID").Value = Val(LV.SelectedItem.Tag) 'List1.ItemData(List1.ListIndex)
        .Parameters("@NewUserName").Value = NewUserName
        .Execute
    End With
    con.CommitTrans
    Call FillList
    Exit Sub
err:
    MsgBox err.Description
    
End Sub



Private Sub cmdDelete_Click()

    On Error GoTo err
    ' Is it Attempt to Delete The Administrator ? Fail IT
    
    'MsgBox "This option is temporarily DISABLED.", vbCritical
    'Exit Sub
    
    If UCase(LV.SelectedItem.Text) = UCase("Administrator") Then
        MsgBox "User Administrator Cannot Be Deleted.", vbInformation
        Exit Sub
    End If


    If MsgBox("Are You Sure To Delete " & LV.SelectedItem.Text & " ?", vbQuestion + vbYesNo, "User Settings") = vbYes Then
        con.Execute "Delete From Users Where UserName='" & LV.SelectedItem.Text & "'"
        Call FillList
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdMarkInactive_Click()

    On Error GoTo err
     
    If UCase(LV.SelectedItem.Text) = UCase("Administrator") Then
        MsgBox "User Administrator cannot be made Inactive.", vbInformation
        Exit Sub
    End If

    Dim strAction As String
    If cmdMarkInactive.Tag = "0" Then
        strAction = "Active"
    Else
        strAction = "Inactive"
    End If
    
    If MsgBox("Are You Sure To " & strAction & " " & LV.SelectedItem.Text & " ?", vbQuestion + vbYesNo, "User Settings") = vbYes Then
        con.Execute "UPDATE Users SET InActive=" & Val(cmdMarkInactive.Tag) & " WHERE UserName='" & LV.SelectedItem.Text & "'"
        Call FillList
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdNew_Click()

    Me.Hide
    With frmUserInfo
        .Show_Option = Opt_NewUser
        .Show 1
        Call FillList
    End With

    Me.Show 1

End Sub

Private Sub cmdPass_Click()

    If UCase(LV.SelectedItem.Text) = UCase("Administrator") And CurrentUserName <> "Administrator" Then
        MsgBox "Only Administrator Can Change His Password.", vbInformation
        Exit Sub
    End If

    Me.Hide
    With frmUserInfo
        .Show_Option = Opt_ResetPassword
        .Tag = LV.SelectedItem.Text
        .Show 1
    End With
    Me.Show 1
    
End Sub


Private Sub cmdPrint_Click()

    If LV.ListItems.count = 0 Then
        MsgBox "Please Select User.", vbInformation
        Exit Sub
    End If
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\UsersMenuOptions.rpt")
    Dim f As New frmPrevRpt
    Me.Hide
    f.ShowReport "{Users.UserName}='" & LV.SelectedItem.Text & "'", rpt, , , , Me
    
End Sub

Private Sub cmdSettings_Click()
    
    If LV.ListItems.count = 0 Then Exit Sub
    
    If LCase(LV.SelectedItem.Text) = "administrator" Then
        MsgBox "User Administrator Has Full Control Its Access Rights Cannot Be Changed.", vbInformation
        Exit Sub
    End If
    On Error GoTo err
    
    Me.Hide
    
    Load frmUserSettings
    
    With frmUserSettings
        Call .ShowUserSettings(Val(LV.SelectedItem.Tag))
    End With
    
    Me.Show 1
    
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub




Private Sub CommandButton1_Click()

End Sub

Private Sub Form_Load()
   Call FillList
End Sub


Private Sub Disable()
    cmdDelete.Enabled = False
    cmdCopy.Enabled = False
    cmdPass.Enabled = False
    cmdSettings.Enabled = False
End Sub

Private Sub Enable()

    cmdDelete.Enabled = True
    'cmdCopy.Enabled = True
    cmdPass.Enabled = True
    cmdSettings.Enabled = True
    
    If LV.SelectedItem.ForeColor = vbRed Then   'If In-Active
        cmdMarkInactive.Caption = "Make Active"
        cmdMarkInactive.Tag = "0"
    Else
        cmdMarkInactive.Caption = "Make Inactive"
        cmdMarkInactive.Tag = "1"
    End If
    
End Sub

Private Sub Form_Unload(cancel As Integer)
    OurOwnForm = True
End Sub

Private Sub List1_Click()
    Call Enable
End Sub


Public Sub FillList()

    On Error GoTo err
    
    Dim i As Integer, ITM As ListItem, lForeColor As Long
    If Rec.State = 1 Then Rec.Close
    Rec.CursorLocation = adUseClient
    Dim strCondition As String
    If Opts(1).Value Then
        strCondition = " WHERE InActive=0"
    ElseIf Opts(2).Value Then
        strCondition = " WHERE InActive=1"
    End If
    strCondition = strCondition & " " & strOrderBy
    Rec.Open "SELECT UserName,UserID,InActive FROM Users " & strCondition, con, adOpenForwardOnly, adLockReadOnly
    
    LV.ListItems.Clear
    For i = 0 To Rec.RecordCount - 1
        Set ITM = LV.ListItems.add(, , Rec.Fields(0) & "")
        ITM.Tag = Rec![UserID]
        If Rec!Inactive Then
            lForeColor = vbRed
        Else
            lForeColor = LV.ForeColor
        End If
        ITM.ForeColor = lForeColor
        Rec.MoveNEXT
    Next i
     
    If LV.ListItems.count > 0 Then
        LV.ListItems(1).Selected = True
    Else
        cmdSettings.Enabled = False
        cmdPass.Enabled = False
        cmdDelete.Enabled = False
    End If
    
    Rec.Close
    Set Rec = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'Call FillList(" ORDER BY UserName")
End Sub

Private Sub LV_DblClick()
    If LV.ListItems.count = 0 Then Exit Sub
    Call cmdSettings_Click
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call Enable
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If LV.ListItems.count = 0 Then Exit Sub
    If LV.SelectedItem.ForeColor = vbRed Then 'In-Active
        Dim strRet As String
        strRet = ShowPopUpMenu1("Re-Activate")
        If strRet = "Re-Activate" Then
            If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            con.Execute "UPDATE Users SET InActive=0 WHERE UserName='" & LV.SelectedItem.Text & "'"
            Call FillList
        End If
    End If
End Sub


Private Sub List1_DblClick()
    If List1.ListIndex = -1 Then Exit Sub
    Call cmdSettings_Click
End Sub


Private Sub Opts_Click(Index As Integer)
    Call FillList
End Sub

Private Sub txtSearch_GotFocus()
    If txtSearch.Text = txtSearch.Tag Then
        txtSearch.Text = ""
        txtSearch.FontItalic = False
        txtSearch.ForeColor = vbBlack
    End If
End Sub



Private Sub txtSearch_LostFocus()
    If txtSearch = "" Then
        txtSearch.Text = txtSearch.Tag
        txtSearch.FontItalic = True
        txtSearch.ForeColor = &HC0C0C0
    End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        Dim i As Integer, iStartIndex As Integer
        If Not (LV.SelectedItem Is Nothing) Then
            If LV.SelectedItem.Index = LV.ListItems.count Then
                iStartIndex = 1
            Else
                iStartIndex = LV.SelectedItem.Index + 1
            End If
        Else
            iStartIndex = 1
        End If
        
        For i = iStartIndex To LV.ListItems.count
            If UCase(LV.ListItems(i).Text) Like "*" & UCase(txtSearch) & "*" Then
                LV.ListItems(i).Selected = True
                LV.SelectedItem.EnsureVisible
                'LV.SetFocus
                Exit For
            End If
        Next
        If i = LV.ListItems.count + 1 Then
            MsgBox "No Mathes Found.", vbInformation
            Set LV.SelectedItem = Nothing
        End If
    End If
    
End Sub

