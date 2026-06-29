VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmVendNewPORM 
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11880
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   9945
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame FALL 
      Height          =   9870
      Left            =   105
      TabIndex        =   0
      Top             =   30
      Width           =   11685
      Begin ComboList.Usercontrol1 cmbItem 
         Height          =   285
         Left            =   30
         TabIndex        =   5
         Top             =   930
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   503
         BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   -1
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4635
         Left            =   30
         TabIndex        =   14
         Top             =   4380
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   8176
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Raw Material"
         TabPicture(0)   =   "frmVendNewPORM.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label3(1)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label2(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "LVRM"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txtTotalOrdQty"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cmdPost"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "txtOrdQty"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "cmbVender"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "cmbMaterial"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtRate"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txtQty"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "cmdAdd"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txtAvgQty"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "txtTotalReq"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).ControlCount=   13
         TabCaption(1)   =   "Order RM"
         TabPicture(1)   =   "frmVendNewPORM.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         Begin VB.TextBox txtTotalReq 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   555
            Width           =   825
         End
         Begin VB.TextBox txtAvgQty 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   555
            Width           =   825
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   10860
            TabIndex        =   25
            Top             =   540
            Width           =   675
         End
         Begin VB.TextBox txtQty 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   10005
            TabIndex        =   24
            Top             =   540
            Width           =   825
         End
         Begin VB.TextBox txtRate 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   9150
            TabIndex        =   23
            Top             =   540
            Width           =   825
         End
         Begin ComboList.Usercontrol1 cmbMaterial 
            Height          =   285
            Left            =   45
            TabIndex        =   21
            Top             =   555
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   503
            BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ListIndex       =   -1
         End
         Begin ComboList.Usercontrol1 cmbVender 
            Height          =   285
            Left            =   5040
            TabIndex        =   18
            Top             =   555
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   503
            BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ListIndex       =   -1
         End
         Begin VB.TextBox txtOrdQty 
            Height          =   300
            Left            =   4965
            TabIndex        =   20
            Top             =   1755
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.CommandButton cmdPost 
            Caption         =   "&Post"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   9825
            TabIndex        =   19
            Top             =   4200
            Width           =   1665
         End
         Begin VB.TextBox txtTotalOrdQty 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   8595
            TabIndex        =   16
            Top             =   4230
            Width           =   1110
         End
         Begin MSComctlLib.ListView LVRM 
            Height          =   3315
            Left            =   45
            TabIndex        =   15
            Top             =   855
            Width           =   11490
            _ExtentX        =   20267
            _ExtentY        =   5847
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   8388608
            BackColor       =   15199215
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S. No."
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Process"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Group"
               Object.Width           =   3175
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "RM Description"
               Object.Width           =   6350
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Rate"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Qty"
               Object.Width           =   1764
            EndProperty
         End
         Begin MSForms.Label Label2 
            Height          =   225
            Index           =   0
            Left            =   45
            TabIndex        =   22
            Top             =   330
            Width           =   11490
            ForeColor       =   4194304
            BackColor       =   12504280
            Caption         =   $"frmVendNewPORM.frx":0038
            Size            =   "20267;397"
            BorderColor     =   10860487
            BorderStyle     =   1
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Total Order Qty :"
            Height          =   195
            Index           =   1
            Left            =   7320
            TabIndex        =   17
            Top             =   4275
            Width           =   1245
         End
      End
      Begin VB.TextBox txtTotalPlanQty 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   7560
         TabIndex        =   12
         Top             =   3885
         Width           =   1110
      End
      Begin VB.TextBox txtPlanQty 
         Height          =   300
         Left            =   1740
         TabIndex        =   11
         Top             =   2040
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   6345
         TabIndex        =   7
         Top             =   9255
         Width           =   5280
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3480
            TabIndex        =   10
            Top             =   165
            Width           =   1665
         End
         Begin VB.CommandButton cmdSaveClose 
            Caption         =   "&Save && Print"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1770
            TabIndex        =   9
            Top             =   135
            Width           =   1665
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save && &New   "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   60
            TabIndex        =   8
            Top             =   165
            Width           =   1665
         End
      End
      Begin VB.Frame Frame1 
         Height          =   645
         Left            =   0
         TabIndex        =   1
         Top             =   75
         Width           =   11685
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Raw Material Purchase Order"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   18
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   390
            Index           =   0
            Left            =   60
            TabIndex        =   2
            Top             =   105
            Width           =   11565
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Raw Material Purchase Order"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   18
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Index           =   1
            Left            =   75
            TabIndex        =   3
            Top             =   120
            Width           =   11550
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   2580
         Left            =   30
         TabIndex        =   6
         Top             =   1230
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   4551
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   15199215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. No."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Buyer"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Order No."
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Delivery Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Ord. Qty"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Plan Qty"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Plan Qty :"
         Height          =   195
         Index           =   0
         Left            =   6405
         TabIndex        =   13
         Top             =   3930
         Width           =   1125
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   2
         Left            =   45
         TabIndex        =   4
         Top             =   720
         Width           =   10755
         ForeColor       =   4194304
         BackColor       =   12504280
         Caption         =   " Item Code"
         Size            =   "18971;397"
         BorderColor     =   10860487
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
End
Attribute VB_Name = "frmVendNewPORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrderNo As String

Private Sub cmbProcess_matched()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Max(Cast(SubString(OrderNo,8,Len(OrderNo)-1) As Int)) From VendPurchaseOrder", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            txtOrderNo = "HS-" & Left(cmbProcess.Text, 3) & "-1"
        Else
            txtOrderNo = "HS-" & Left(cmbProcess.Text, 3) & "-" & Val(.Fields(0) & "") + 1
        End If
        .Close
    End With
    Set rs = Nothing
    
    Dim cmd As New ADODB.Command
    
    Call cmbVend.ClearVals
    
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "SP_ReqAuthForRcv"
        Set .ActiveConnection = con
        .Parameters("@UserID").Value = UserID
        .Parameters("@ProcessID").Value = cmbProcess.ID
        .Execute
        UserReqAuth = Val(.Parameters("@ReqAuth").Value & "")
        
        '.CommandText = "SP_VendersForRcving"
        '.Parameters("@ProcessID").value = cmbProcess.ID
        '.Parameters("@HicoVisible").value = HicoVisible
        
        'Set rs = .Execute
    End With
    
    If RcvLot(cmbProcess.ID) Then
        RcvingLot = True
    Else
        RcvingLot = False
    End If
    
    LV.ListItems.Clear
    GotRcvID = 0
    
    
    'If HicoVisible Then
    '    Call cmbVend.AddVals(con, "VenderName", "VVendIssItems", "VendID", " Where PhaseID=" & cmbProcess.ID & " Or PhaseID=-1")
    'Else
    '    Call cmbVend.AddVals(con, "VenderName", "VVendIssItems", "VendID", " Where PhaseID=" & cmbProcess.ID)
    'End If
    
'    If HicoVisible Then
'        Call cmbVend.AddVals(con, "VenderName", "Venders", "VendID", " Where (PhaseID=" & cmbProcess.ID & " Or PhaseID=-1) And VendID In(Select VendID From VVendIssItems Where ProcessID=" & cmbProcess.ID & " Group By VendID Having Sum(RemQty)>0)")
'    Else
    Call cmbVend.AddVals(con, "VenderName", "VVenders", "VendID", " Where ProcessID=" & cmbProcess.ID & " and VendType=0")
'    End If
    
    Dim lSNo1 As Long, lSNo2 As Long
    lSNo1 = Val(con.Execute("Select SNo From VInitialLotProcess").Fields(0).Value & "")
    lSNo2 = Val(con.Execute("Select SNo From Processes Where ProcessID=" & cmbProcess.ID).Fields(0).Value & "")
    
    If lSNo2 >= lSNo1 Then
        'If LV.ColumnHeaders.Count = 5 Then
        '    LV.ColumnHeaders.Add(, , "Lot #", "2000").Tag = "2000"
        'End If
        cmbProcess.Tag = "Lot No"
    Else
        'If LV.ColumnHeaders.Count = 6 Then
        '    LV.ColumnHeaders.Remove 6
        'End If
        cmbProcess.Tag = ""
    End If
    
    
    'Set Focus To CmbVenders And Also Drop It Down!
    cmbVend.SetFocus
    SendKeys "a"
    SendKeys Chr(8)
    
    Exit Sub
    
err:
    MsgBox err.Description
End Sub

Private Sub cmbVend_matched()
    
    On Error GoTo err
    
    cmbItem.AddVals con, "ItemID", "VVendAssItems", "ItemID", " Where VendID=" & Val(cmbVend.ID) & "  Order By ItemID"
    cmbItemName.AddVals con, "ItemName", "VVendAssItems", "ItemID", " Where VendID=" & Val(cmbVend.ID) & "  Order By ItemID"
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdSClear_Click()
'
'    If Saved Then
'        FormProcessID = cmbProcess.ID
'        Unload Me
'        Call ShowRcvVForm(FormProcessID)
'    End If
'    'Call ClearFields

End Sub

Private Sub cmdSClose_Click()

    'If Not AskForAutoApply Then Exit Sub

    If Saved Then
        Unload Me
    End If

End Sub

Private Sub cmbMaterial_matched()
    cmbVender.ClearVals
    cmbVender.AddVals con, "AccTitle", "VVenders", "VendID", "WHERE VendID IN(SELECT VendID FROM VenderAssItems WHERE RMID=" & cmbMaterial.ID & ")"
        
    txtAvgQty = GetSingleLongValue("Qty", "VItemsRMComp", " WHERE ItemID='" & cmbItem.ID & "' AND RMID=" & cmbMaterial.ID)
    txtTotalReq = Val(txtAvgQty) * Val(txtTotalPlanQty)
End Sub

Private Sub cmbVender_matched()
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT Rate FROM VenderAssItems WHERE VendID=" & cmbVender.ID & " AND RMID=" & cmbMaterial.ID, con
        If .EOF = False Then
            txtRate = Val(!Rate & "")
        End If
        .Close
    End With
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdAdd_Click()

    On Error GoTo err
    If Val(txtQty) <= 0 Then
        MsgBox "Invalid Qty.", vbInformation
        txtQty.SetFocus
        Exit Sub
    ElseIf Val(txtRate) <= 0 Then
        MsgBox "Invalid Rate.", vbInformation
        txtRate.SetFocus
        Exit Sub
    End If
     
    Dim ITM As ListItem
    Dim rs As New ADODB.Recordset
    Dim strProcess As String, strRM As String, strGroup As String
    With rs
        .Open "SELECT * FROM VItemsRMComp WHERE ItemID='" & cmbItem.ID & "' AND RMID=" & cmbMaterial.ID, con, adOpenForwardOnly, adLockReadOnly
        strProcess = !ProcDesc & ""
        strRM = !Description & ""
        strGroup = !GroupDescription & ""
        .Close
    End With
    Set rs = Nothing
    Set ITM = LVRM.ListItems.add(, , LV.ListItems.count + 1)
    ITM.ListSubItems.add(, , strProcess).Tag = cmbVender.ID
    ITM.ListSubItems.add(, , strGroup).Tag = cmbMaterial.ID
    ITM.ListSubItems.add , , strRM
    ITM.ListSubItems.add , , Val(txtRate)
    ITM.ListSubItems.add , , Val(txtQty)
    
    txtRate = 0
    txtQty = 0
    cmbMaterial.SetFocus
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbItem_matched()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "SELECT OrderNo,CustCode,DeliveryDT,Qty,CompItemID FROM VrptOrders WHERE CompItemID='" & cmbItem.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, !OrderNo & "-" & !CompITemID, LV.ListItems.count + 1)
            ITM.SubItems(1) = !CustCode & ""
            ITM.SubItems(2) = !OrderNo & ""
            ITM.SubItems(3) = Format(!DeliveryDT, "dd-MMM-yyyy")
            ITM.SubItems(4) = Val(!Qty & "")
            .MoveNEXT
        Loop
        .Close
'        .Open "SELECT * FROM VItemsRMComp WHERE ItemID='" & cmbItem.ID & "'", con, adOpenForwardOnly, adLockReadOnly
'        LVRM.ListItems.Clear
'        Do Until .EOF
'            Set ITM = LVRM.ListItems.add(, !RMID & "'", LVRM.ListItems.count + 1)
'            ITM.ListSubItems.add , , !GroupDescription & ""
'            ITM.ListSubItems.add , , !Description & ""
'            ITM.ListSubItems.add(, , Val(!Qty & "") * Val(txtTotalPlanQty)).Tag = Val(!Qty & "")
'            .MoveNext
'        Loop
'        .Close
    End With
    Set rs = Nothing
    cmbMaterial.ClearVals
    cmbMaterial.AddVals con, "Description + '{' + ProcDesc + '}'", "VItemsRMComp", "RMID", " WHERE ItemID='" & cmbItem.ID & "'"
     
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    con.Execute "Delete From PrintSession"
End Sub


Private Sub cmdPost_Click()

    On Error GoTo err
    If cmbVender.MatchFound = False Then Exit Sub
    If Val(txtTotalOrdQty) <= 0 Then Exit Sub
    Dim ITM As ListItem, i As Integer
    
    For i = 1 To LVRM.ListItems.count
        With LVRM.ListItems(i)
            Set ITM = LVOrd.ListItems.add(, LV.SelectedItem.SubItems(2) & "-" & .key & "-" & cmbVender.ID, LVOrd.ListItems.count + 1)
            ITM.ListSubItems.add(, , cmbVender.Text).Tag = cmbVender.ID
            ITM.ListSubItems.add(, , .SubItems(2)).Tag = Val(.key)
            ITM.ListSubItems.add , , Val(.SubItems(4))
        End With
    Next
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdSave_Click()
'
    'If Saved Then
'        FormProcessID = cmbProcess.ID
'        Unload Me
'        Call ShowIssPOForm(FormProcessID)
'    End If
    
End Sub

Private Sub cmdSaveClose_Click()

    If Saved Then
        Unload Me
    End If
    
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Pic.Visible = False
    ElseIf KeyCode = 13 Then
        If Val(txtQty) = 0 Then
                If InitLotProcess Then
                    LV.SelectedItem.SubItems(5) = Val(txtQty)
                Else
                    LV.SelectedItem.SubItems(4) = Val(txtQty)
                End If
                LV.SelectedItem.SubItems(LV.ColumnHeaders.count - 1) = ""
                'Reset The IssID's
                Dim TempID As Long
                TempID = 0
                
                For i = LV.SelectedItem.Index + 1 To LV.ListItems.count
                    TempID = Val(Right(LV.ListItems(i).SubItems(LV.ColumnHeaders.count - 1), 3)) - 1
                    If TempID > 0 Then
                        LV.ListItems(i).SubItems(LV.ColumnHeaders.count - 1) = "RCV-" & _
                         Format(ServerDate, "ddMMyy") & TempID
                    End If
                Next i
                If GotRcvID > 0 Then GotRcvID = GotRcvID - 1
                
                
                
                Pic.Visible = False
                Exit Sub
            End If
            
            
            
            If InitLotProcess Then
                LV.SelectedItem.SubItems(4) = Val(txtQty)
                LV.SelectedItem.SubItems(5) = Format(DTPicker2.Value, "dd-MMM-yyyy")
            Else
                LV.SelectedItem.SubItems(4) = Val(txtQty)
                LV.SelectedItem.SubItems(5) = Format(DTPicker2.Value, "dd-MMM-yyyy")
            End If
            
            '-----------------------
            
            'Dim RcvdID As String
            
            If cmbProcess.Tag = "Lot No" Then
                If GotRcvID = 0 Then
                    GotRcvID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-9) As Int)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(DTPicker2, "dd MMM yy") & "' and UPPER(Left(RecieptID,2))='HS'").Fields(0).Value & "") + 1  'and UPPER(Left(RecieptID,3))<>'HSR'
                Else
                    GotRcvID = GotRcvID + 1
                End If
                LV.SelectedItem.SubItems(LV.ColumnHeaders.count - 1) = "HS-" & _
                Format(DTPicker2, "ddMMyy") & GotRcvID
            Else
                If GotRcvID = 0 Then
                    GotRcvID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-10) As Int)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(DTPicker2, "dd MMM yy") & "' and UPPER(Left(RecieptID,3))='RCV'").Fields(0).Value & "") + 1
                Else
                    GotRcvID = GotRcvID + 1
                End If
                'LV.SelectedItem.SubItems(LV.ColumnHeaders.Count - 1) = "RCV-" & _
                 Format(DTPicker1, "ddMMyy") & "-" & Format(GotRcvID, "000")
                 LV.SelectedItem.SubItems(LV.ColumnHeaders.count - 1) = "RCV-" & _
                  Format(DTPicker2, "ddMMyy") & GotRcvID
            End If
                
            '-----------------------
            
            'Call AdjustQty
            Pic.Visible = False
            
            If LV.SelectedItem.Index < LV.ListItems.count Then
                'If Val(LV.SelectedItem.SubItems(3)) = 0 Then
                 LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
                'End If
                LV.SetFocus
            End If
            
        ElseIf KeyAscii = 27 Then
            Pic.Visible = False
        End If
End Sub

Private Sub DTPicker1_Change()
    DTDelivery.Value = DateAdd("m", 1, DTPicker1)
End Sub

Private Sub DTPicker1_Click()
    Call DTPicker1_Change
End Sub

Private Sub Form_Activate()

    'Check For User Access
    If PermissionDenied Then
        Me.Hide
        MsgBox "You Donot Have Permissions To Use This Option. Please Contact Your Administrator", vbInformation
        Unload Me
        Exit Sub
    End If


End Sub

Private Sub Form_Load()
        
    On Error GoTo err
    With cmbItem
        .ListHeight = 6000
        .AddVals con, "ItemID", "Items", "ItemID", " ORDER BY ItemID"
    End With
    
    With cmbVender
        .ListHeight = 2600
        .AddVals con, "AccTitle", "VVenders", "VendID", " ORDER BY AccTitle"
    End With
    cmbMaterial.ListHeight = 2600
    SSTab1.TabVisible(1) = False
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Public Sub LockProc(ProcID As Long)
    Call Form_Resize
    cmbProcess.ID = ProcID
    cmbProcess.Enabled = False
    Me.Show
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
    
    Set Col_OrdQty = Nothing
    Set frmVendRcvItems = Nothing
    
End Sub

Private Sub Form_Resize()
    FAll.Move (Me.ScaleWidth - FAll.Width) / 2, (Me.ScaleHeight - FAll.Height) / 2
End Sub

Private Sub LV_DblClick()
    If LV.ListItems.count = 0 Then Exit Sub
    With txtPlanQty
        .Move LV.Left + LV.ColumnHeaders(6).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(6).Width
        .Text = Val(LV.SelectedItem.SubItems(5))
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
End Sub


Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        Item.SubItems(4) = Item.SubItems(3)
    Else
        Item.SubItems(4) = 0
    End If
End Sub

Private Sub LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call LV_DblClick
    End If
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
   ' Exit Sub
   ' If LV.ListItems.Count = 0 Then Exit Sub
    'If Val(LV.SelectedItem.SubItems(6)) = 0 Then Exit Sub
    'Dim iCol As Integer
    'iCol = GetColNo(LV, X)
    
    'If iCol = 8 Then
    '    Load frmVendRcvItems1
    '    LV.SelectedItem.ListSubItems(7).Text = frmVendRcvItems1.FillList(Col_OrdQty, LV.SelectedItem.SubItems(1))
    '    LV.SelectedItem.ListSubItems(7).Tag = LV.SelectedItem.ListSubItems(7).Text
    'End If
    
End Sub

Private Sub txtQtyOld_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    If KeyAscii = 13 Then
        
        If Val(txtQty) = 0 Then
            If InitLotProcess Then
                LV.SelectedItem.SubItems(5) = Val(txtQty)
            Else
                LV.SelectedItem.SubItems(4) = Val(txtQty)
            End If
            LV.SelectedItem.SubItems(LV.ColumnHeaders.count - 1) = ""
            'Reset The IssID's
            Dim TempID As Long
            TempID = 0
            
            For i = LV.SelectedItem.Index + 1 To LV.ListItems.count
                TempID = Val(Right(LV.ListItems(i).SubItems(LV.ColumnHeaders.count - 1), 3)) - 1
                If TempID > 0 Then
                    LV.ListItems(i).SubItems(LV.ColumnHeaders.count - 1) = "RCV-" & _
                     Format(ServerDate, "ddMMyy") & "-" & Format(TempID, "000")
                End If
            Next i
            If GotRcvID > 0 Then GotRcvID = GotRcvID - 1
            
            
            
            Pic.Visible = False
            Exit Sub
        End If
        
        If InitLotProcess Then
            If Val(txtQty) > Val(LV.SelectedItem.SubItems(4)) Then
                MsgBox "Invalid Quantities!", vbInformation
                Exit Sub
            End If
        Else
            If Val(txtQty) > Val(LV.SelectedItem.SubItems(3)) Then
                MsgBox "Invalid Quantities!", vbInformation
                Exit Sub
            End If
        End If
        
        If Not txtQty.Locked Then
            If cmbProcess.Tag = "Lot No" Then
                If lLotNo = 0 Then
                    lLotNo = Val(con.Execute("Select Max(Cast(LotNo As int)) From VendRcvdDetail").Fields(0).Value & "") + 1
                Else
                    lLotNo = lLotNo + 1
                End If
                
                If InitLotProcess Then
                    LV.SelectedItem.ListSubItems(4).Tag = lLotNo
                Else
                    LV.SelectedItem.ListSubItems(4).Tag = lLotNo
                End If
            End If
        End If
        If InitLotProcess Then
            LV.SelectedItem.SubItems(5) = Val(txtQty)
        Else
            LV.SelectedItem.SubItems(4) = Val(txtQty)
        End If
        
        '-----------------------
        
        'Dim RcvdID As String
        
            If cmbProcess.Tag = "Lot No" Then
                If GotRcvID = 0 Then
                    GotRcvID = Val(con.Execute("Select Max(right(RecieptID,3)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "' and UPPER(Left(RecieptID,2))='HS'").Fields(0).Value & "") + 1  'and UPPER(Left(RecieptID,3))<>'HSR'
                Else
                    GotRcvID = GotRcvID + 1
                End If
                LV.SelectedItem.SubItems(LV.ColumnHeaders.count - 1) = "HS-" & _
                 Format(DTPicker1, "ddMMyy") & "-" & Format(GotRcvID, "000")
            Else
                If GotRcvID = 0 Then
                    GotRcvID = Val(con.Execute("Select Max(right(RecieptID,3)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "' and UPPER(Left(RecieptID,3))='RCV'").Fields(0).Value & "") + 1
                Else
                    GotRcvID = GotRcvID + 1
                End If
                 LV.SelectedItem.SubItems(LV.ColumnHeaders.count - 1) = "RCV-" & _
                 Format(DTPicker1, "ddMMyy") & "-" & Format(GotRcvID, "000")
            End If
            
        '-----------------------
        
        'Call AdjustQty
        Pic.Visible = False
    ElseIf KeyAscii = 27 Then
        Pic.Visible = False
    End If
    
    Call OnlyNums(KeyAscii, False)
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub AdjustQty()
'
'    On Error GoTo err
'    Dim rs As New ADODB.Recordset
'    Dim lQty As Long, TotQty As Long, WasteQty As Long
'    Dim TempQty As Long, TempWasteQty As Long
'    Dim StrDetail As String
'
'    TempQty = 0
'    TotQty = Val(LV.SelectedItem.SubItems(6))
'    WasteQty = Val(LV.SelectedItem.SubItems(5))
'
'    With rs
'        .Open "Select OrderNo,ItemCode,TotQty From VTotIssdByOrder Where ItemCode='" & LV.SelectedItem.SubItems(1) & "'", con, adOpenForwardOnly
'
'        Do Until .EOF
'            Dim Obj As New RcvAdjOrderQty
'            Obj.OrderNo = ![OrderNo] & ""
'            Obj.ItemCode = ![ItemCode] & ""
'            Obj.lOrdQty = ![TotQty] & ""
'            Obj.lWastage = 0
'            Obj.lQty = 0
'
'            Col_OrdQty.add Obj, ![OrderNo] & "|" & ![ItemCode]
'            Set Obj = Nothing
'            .MoveNext
'        Loop
'        .MoveFirst
'
'        Do While lQty < TotQty
'
'            TempQty = TotQty - lQty
'            TempWasteQty = 0
'            If TempQty > ![TotQty] Then
'                TempQty = ![TotQty]
'            End If
'
'            If WasteQty > 0 Then
'                If TempQty > WasteQty Then
'                    TempQty = TempQty - WasteQty
'                    TempWasteQty = WasteQty
'                    WasteQty = 0
'                Else
'                    TempWasteQty = WasteQty
'                    TempQty = 0
'                    WasteQty = WasteQty - TempQty
'                End If
'            End If
'
'
'            Col_OrdQty(![OrderNo] & "|" & ![ItemCode]).lQty = TempQty
'            Col_OrdQty(![OrderNo] & "|" & ![ItemCode]).lWastage = TempWasteQty
'
'            lQty = lQty + TempQty + TempWasteQty
'
'            StrDetail = StrDetail & ![OrderNo] & "=" & TempQty + TempWasteQty & ","
'
'            .MoveNext
'        Loop
'        .Close
'    End With
'
'    StrDetail = Left(StrDetail, Len(StrDetail) - 1)
'    LV.SelectedItem.ListSubItems(7) = StrDetail
'    LV.SelectedItem.ListSubItems(7).ToolTipText = StrDetail
'    Set rs = Nothing
'    Exit Sub
'err:
'    MsgBox err.Description
'
End Sub


Private Function OLDSaved() As Boolean

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim Col As New Collection
    
    Dim cmd As New ADODB.Command
    
    Dim iRefID As Long
    
    Call StartTrans(con)
    
    con.Execute "Insert Into VendReceived(DT,RecieptNo,VendID) Values('" & DTPicker1 & "','" & txtRecNo & "'," & cmbVend.ID & ")"
    
    iRefID = Val(con.Execute("Select Max(EntryID) From VendReceived").Fields(0).Value & "")
    
    'Get The Next Process ID For Current Item...
    For i = 1 To LV.ListItems.count
        
            If Val(LV.ListItems(i).SubItems(6)) > 0 Then
                Dim iNextProcID As Integer
                'Dim cmd As New ADODB.Command
                With cmd
                    .ActiveConnection = con
                    .CommandType = adCmdStoredProc
                    .CommandText = "SP_GetNextProcID"
                    .Parameters("@ItemCode").Value = LV.ListItems(i).SubItems(1)
                    .Parameters("@ProcID").Value = cmbProcess.ID
                    .Execute
                    iNextProcID = Val(.Parameters("@NextProcID").Value & "")
                End With
                
                For j = 1 To Col_OrdQty.count
                    With Col_OrdQty(j)
                        If .ItemCode = LV.ListItems(i).SubItems(1) And .lQty + .lWastage > 0 Then
                        
                            con.Execute "Insert Into VendRcvdDetail(RefID,OrderNo,ItemCode,NextProcessID,ProcessID,RcvdQty,Wastage) Values(" & _
                            iRefID & ",'" & .OrderNo & "','" & .ItemCode & "'," & iNextProcID & "," & cmbProcess.ID & "," & .lQty & "," & .lWastage & ")"
                            
                            '2. Now If Rcvd After Issued Then
                            'Update The RcvdQty In VendIssdDetails....
                            '-----------------------------------------------------------------
                            With cmd
                                .ActiveConnection = con
                                .CommandType = adCmdStoredProc
                                .CommandText = "SP_UpdateIssued"
                                .Parameters("@OrderNo").Value = Col_OrdQty(j).OrderNo
                                .Parameters("@ItemCode").Value = Col_OrdQty(j).ItemCode
                                .Parameters("@ProcessID").Value = cmbProcess.ID
                                .Parameters("@Qty").Value = Col_OrdQty(j).lQty + Col_OrdQty(j).lWastage
                                .Execute
                            End With
                            '-----------------------------------------------------------------
                                
                                                    
                            'con.Execute "Update VendIssdDetail Set RcvdQty=RcvdQty+" & .lQty + .lWastage & " Where OrderNo='" & .OrderNo & "' and ItemCode='" & .ItemCode & "'"
                        End If
                    End With
                Next j
                
                'For j = 1 To Col.Count
                '    Col.Remove 1
                'Next
                
                
                
                
                '---------------------------------------------------------------------------------------------------------------------------------------
                'Update RcvdQty In The VendIssdDetail Table....
                
                'With rs
                '    .Open "Select EntryID,IssQty-RcvdQty From VendIssdDetail Where ItemCode='" & LV.ListItems(i).SubItems(1) & "' and IssQty-RcvdQty>0 and ProcessID=" & cmbProcess.ID & " Order By EntryID", con, adOpenForwardOnly, adLockReadOnly
                    
                '    Do Until .EOF
                '        Dim Obj As New AdjIssQty
                '        Obj.EntryID = ![EntryID]
                '        Obj.lQty = .Fields(1)
                '        Col.Add Obj
                '        Set Obj = Nothing
                '        .MoveNext
                '    Loop
                '    .Close
                'End With
                
                'Dim TempQty As Long, TotQty As Long
                'Dim TempTot As Long
                'TotQty = Val(LV.ListItems(i).SubItems(6))
                'TempQty = 0
                'TempTotQty = 0
                
                'For j = 1 To Col.Count
                
                '    If TotQty - TempTot >= Col(j).lQty Then
                '        TempQty = Col(j).lQty
                '    Else
                '        TempQty = TotQty - TempTot
                '    End If
                '    TempTot = TempTot + TempQty
                '
                '    con.Execute "Update VendIssdDetail Set RcvdQty=" & TempQty & " Where EntryID=" & Col(j).EntryID
                '
                '    If TempTot >= TotQty Then Exit For
                '
                'Next j
                '---------------------------------------------------------------------------------------------------------------------------------------
            End If
        
    Next i
    
    Set rs = Nothing
    con.CommitTrans
    OLDSaved = True
    Exit Function
err:
    MsgBox err.Description
    
End Function
    
Private Function Saved() As Boolean

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim Col As New Collection
    
    Dim cmd As New ADODB.Command
    
    Dim iRefID As Long
    
    Dim j As Integer, k As Integer
    
    Call StartTrans(con)
    
    Dim i As Integer
    Dim myPOSavedCollection As New Collection
    Dim strMyOrderNo As String, lMyRefID As Long
    
    For i = 1 To LVRM.ListItems.count
        strMyOrderNo = ""
        Dim myPOObjTemp As clsRMPOEntry
        For j = 1 To myPOSavedCollection.count
            Set myPOObjTemp = myPOSavedCollection(j)
            If myPOObjTemp.iVendID = Val(LVRM.ListItems(i).ListSubItems(1).Tag) Then
                strMyOrderNo = myPOObjTemp.strOrderNo
                Set myPOObjTemp = Nothing
                Exit For
            End If
            Set myPOObjTemp = Nothing
        Next
'        For k = 1 To myPOSavedCollection.count
'            Set myPOObjTemp = myPOSavedCollection(j)
'            If myPOObjTemp.iVendID = Val(.ListSubItems(1).Tag) And myPOObjTemp.iMaterialID = Val(.ListSubItems(2).Tag) Then
'                lMyRefID = myPOObjTemp.lRefID
'                Set myPOObjTemp = Nothing
'                Exit For
'            End If
'            Set myPOObjTemp = Nothing
'        Next
        If strMyOrderNo = "" Then
            strMyOrderNo = GetNextPO
            con.Execute "INSERT INTO VendOrders(OrderNo,DT,VendID) Values('" & strMyOrderNo & "','" & Date & "'," & LVRM.ListItems(i).ListSubItems(1).Tag & ")"
            Dim myPOObj As New clsRMPOEntry
            myPOObj.iVendID = Val(LVRM.ListItems(i).ListSubItems(1).Tag)
            myPOObj.strOrderNo = strMyOrderNo
            myPOSavedCollection.add myPOObj
            Set myPOObj = Nothing
        End If
        con.Execute "INSERT INTO VendOrderDetail(OrderNo,RMID,Rate,Qty) Values('" & strMyOrderNo & "'," & LVRM.ListItems(i).ListSubItems(2).Tag & "," & _
         Val(LVRM.ListItems(i).SubItems(4)) & "," & Val(LVRM.ListItems(i).SubItems(5)) & ")"
         
    Next
    
    
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            con.Execute "INSERT INTO VendPurchOrdDetail(OrderNo,ItemCode,Rate,Qty) Values('" & txtOrderNo & "','" & .SubItems(1) & "'," & Val(.SubItems(3)) & "," & Val(.SubItems(4)) & ")"
        End With
    Next
    
    con.CommitTrans
    
    'Print Report
    'Now Print
    
    For j = 1 To myPOSavedCollection.count
        Set myPOObjTemp = myPOSavedCollection(j)
        strMyOrderNo = myPOObjTemp.strOrderNo
        Call PrintRMPO(strMyOrderNo)
            
        Set myPOObjTemp = Nothing
    Next
    
    Saved = True
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Private Sub PrintRMPO(strOrderNo As String)
    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\VendRMPO.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{IssItemsSimple.OrderNo}='" & strOrderNo & "' ", rpt
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Function GetNextPO() As String
    Dim Prefix As String
    'ORD-DDMM-001
    Prefix = "RMPO-" & Format(Date, "ddMMyy") & "-"
    GetNextPO = Prefix & Format(GetMax("RIGHT(RTrim(OrderNo),3)", "VendOrders", "WHERE Left(OrderNo," & Len(Prefix) & ")='" & Prefix & "'") + 1, "000")
    
End Function

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
'    If KeyAscii = 13 Then
'        LV.SelectedItem.SubItems(4) = Val(txtQty)
'        'Pic.Visible = False
'    ElseIf KeyAscii = 27 Then
'        Pic.Visible = False
'    End If
    
    Call OnlyNums(KeyAscii, False)
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub LVRM_DblClick()
    If LVRM.ListItems.count = 0 Then Exit Sub
    With txtOrdQty
        .Move LVRM.Left + LVRM.ColumnHeaders(5).Left, LVRM.Top + LVRM.SelectedItem.Top, LVRM.ColumnHeaders(5).Width
        .Text = Val(LVRM.SelectedItem.SubItems(4))
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub txtOrdQty_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(txtOrdQty) > Val(LVRM.SelectedItem.SubItems(3)) Or Val(txtOrdQty) < 0 Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        End If
        LVRM.SelectedItem.SubItems(4) = Val(txtOrdQty)
        txtOrdQty.Visible = False
        Call UpdateORDTotals
    ElseIf KeyAscii = 27 Then
        txtOrdQty.Visible = False
    End If

End Sub

Private Sub txtOrdQty_LostFocus()
    txtOrdQty.Visible = False
End Sub

Private Sub txtPlanQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txtPlanQty) > Val(LV.SelectedItem.SubItems(4)) Or Val(txtPlanQty) < 0 Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        End If
        LV.SelectedItem.SubItems(5) = Val(txtPlanQty)
        txtPlanQty.Visible = False
        Call UpdateTotals
    ElseIf KeyAscii = 27 Then
        txtPlanQty.Visible = False
    End If
End Sub

Private Sub UpdateTotals()

    Dim i As Integer, lTotal As Long
    lTotal = 0
    For i = 1 To LV.ListItems.count
        lTotal = lTotal + Val(LV.ListItems(i).SubItems(5))
    Next
    txtTotalPlanQty = lTotal
    For i = 1 To LVRM.ListItems.count
        With LVRM.ListItems(i)
            .SubItems(3) = Val(.ListSubItems(3).Tag) * lTotal
        End With
    Next
    
End Sub

Private Sub txtPlanQty_LostFocus()
    txtPlanQty.Visible = False
End Sub

Private Sub UpdateORDTotals()

    Dim i As Integer, lTotal As Long
    lTotal = 0
    For i = 1 To LVRM.ListItems.count
        lTotal = lTotal + Val(LVRM.ListItems(i).SubItems(4))
    Next
    txtTotalOrdQty = lTotal
    
End Sub

