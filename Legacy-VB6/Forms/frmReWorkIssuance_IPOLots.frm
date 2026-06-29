VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmReWorkIssuance_IPOLots 
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   9390
   ScaleMode       =   0  'User
   ScaleWidth      =   10721.69
   WindowState     =   2  'Maximized
   Begin VB.Frame FAcc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   9420
      Left            =   60
      TabIndex        =   1
      Top             =   -45
      Width           =   11865
      Begin VB.CommandButton cmdPrintPending 
         Caption         =   "Pending"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7875
         TabIndex        =   26
         Top             =   8940
         Visible         =   0   'False
         Width           =   1125
      End
      Begin ComboList.Usercontrol1 cmbArticle 
         Height          =   285
         Left            =   105
         TabIndex        =   2
         Top             =   930
         Width           =   11670
         _ExtentX        =   20585
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
         Appearance      =   0
      End
      Begin VB.Frame Frame1 
         Height          =   3480
         Left            =   105
         TabIndex        =   4
         Top             =   1950
         Width           =   11700
         Begin VB.TextBox txtTotalShipped 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   7185
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   3135
            Width           =   1215
         End
         Begin MSComctlLib.ListView LV 
            Height          =   2850
            Left            =   90
            TabIndex        =   5
            Top             =   240
            Width           =   11490
            _ExtentX        =   20267
            _ExtentY        =   5027
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
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
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Sr. No."
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Lot #"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Customer Order"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Repair Type"
               Object.Width           =   6006
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Qty"
               Object.Width           =   1764
            EndProperty
         End
      End
      Begin VB.TextBox txtLotNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4605
         TabIndex        =   23
         Top             =   1485
         Width           =   825
      End
      Begin VB.Frame fTitle 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   11865
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Re-Work Issuance"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   390
            Index           =   2
            Left            =   0
            TabIndex        =   21
            Top             =   135
            Width           =   11850
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Re-Work Issuance"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Index           =   3
            Left            =   0
            TabIndex        =   22
            Top             =   150
            Width           =   11850
         End
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "Print PTC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6675
         TabIndex        =   17
         Top             =   9015
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.ComboBox cmbRange 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmReWorkIssuance_IPOLots.frx":0000
         Left            =   135
         List            =   "frmReWorkIssuance_IPOLots.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1485
         Width           =   1410
      End
      Begin VB.Frame Frame2 
         Height          =   3555
         Left            =   105
         TabIndex        =   8
         Top             =   5355
         Width           =   11700
         Begin VB.TextBox txtRate 
            Height          =   300
            Left            =   5400
            TabIndex        =   19
            Top             =   1335
            Visible         =   0   'False
            Width           =   930
         End
         Begin ComboList.Usercontrol1 cmbMaker 
            Height          =   285
            Left            =   675
            TabIndex        =   16
            Top             =   225
            Width           =   7590
            _ExtentX        =   13388
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
         Begin VB.TextBox txtTotalUnshipped 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   7425
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   3210
            Width           =   1215
         End
         Begin MSComctlLib.ListView LV1 
            Height          =   2610
            Left            =   90
            TabIndex        =   10
            Top             =   555
            Width           =   11490
            _ExtentX        =   20267
            _ExtentY        =   4604
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
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
               Text            =   "Sr. No."
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Lot #"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Customer Order"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Repair Type"
               Object.Width           =   6006
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Qty"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Rate"
               Object.Width           =   1764
            EndProperty
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Maker :"
            Height          =   195
            Left            =   90
            TabIndex        =   15
            Top             =   270
            Width           =   540
         End
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   3075
         TabIndex        =   12
         Top             =   1485
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   151781379
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   1545
         TabIndex        =   13
         Top             =   1485
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   151781379
         CurrentDate     =   37055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         Caption         =   " Range                    From                           To                             Lot #         Repair Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   1
         Left            =   105
         TabIndex        =   25
         Top             =   1230
         Width           =   10620
      End
      Begin MSForms.ComboBox cmbRepairType 
         Height          =   315
         Left            =   5445
         TabIndex        =   24
         Top             =   1485
         Width           =   5295
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "9340;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblStatus 
         Height          =   285
         Left            =   105
         TabIndex        =   18
         Top             =   9000
         Width           =   6375
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         Caption         =   " Article                                                                            "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   0
         Left            =   105
         TabIndex        =   14
         Top             =   675
         Width           =   11670
      End
      Begin MSForms.CommandButton cmdSave 
         Height          =   375
         Left            =   9030
         TabIndex        =   7
         Top             =   8940
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Save"
         PicturePosition =   327683
         Size            =   "2355;661"
         Accelerator     =   83
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   375
         Left            =   10440
         TabIndex        =   3
         Top             =   8940
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "2355;661"
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   10905
         TabIndex        =   0
         Top             =   1365
         Width           =   870
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "1535;635"
         Accelerator     =   82
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopMenu"
   End
End
Attribute VB_Name = "frmReWorkIssuance_IPOLots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iColNo As Integer

Private Sub cmbArticle_matched()

    On Error GoTo err
    
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim lTotalQty As Long
    Dim lBalanceQty As Long, lTotalUnshipped As Long
    Dim strCondition As String
    strCondition = "WHERE DT BETWEEN '" & DtFrom & "' AND '" & DtTo & "' AND CompItemCode='" & cmbArticle.ID & "'"
    
    strCondition = strCondition & " AND IPO_Repair_RefID NOT IN(SELECT IPODL_RefID FROM VendIssdDetail)"
    
    If Val(txtLotNo) > 0 Then
        strCondition = strCondition & " AND IPOLotNo_Original=" & Val(txtLotNo)
    End If
    
    If cmbRepairType.ListIndex > 0 Then
        strCondition = strCondition & " AND Repair_RefID=" & Val(cmbRepairType.List(cmbRepairType.ListIndex, 1))
    End If
    
    With rs
        .Open "SELECT * FROM VIPOReworkRecord " & strCondition & " ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        'LV1.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, !EntryID & "'", LV.ListItems.count + 1)
            ITM.ListSubItems.add , , Val(!IPO_Repair_RefID & "")
            ITM.ListSubItems.add , , !CustomerPO & ""
            ITM.ListSubItems.add(, , !RepairType).Tag = Val(!RepairType_RefID & "")
            ITM.ListSubItems.add , , Val(!ReWorkQty & "")
            
            'lTotalQty = lTotalQty + Val(!Qty & "") - Val(!IssQty & "")
            .MoveNEXT
        Loop
        .Close
        'txtTotalShipped = lTotalQty
        
    End With
    Set rs = Nothing
    
    Call GetLVTotal
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub GetLVTotal()

    Dim i As Integer, lTotal As Long
    For i = 1 To LV.ListItems.count
        lTotal = lTotal + Val(LV.ListItems(i).SubItems(4))
    Next
    txtTotalShipped = lTotal
    
End Sub

Private Sub GetLV1Total()

    Dim i As Integer, lTotal As Long
    For i = 1 To LV1.ListItems.count
        lTotal = lTotal + Val(LV1.ListItems(i).SubItems(4))
    Next
    txtTotalUnshipped = lTotal
    
End Sub


Private Sub cmbRange_Change()
    Call cmbRange_Click
End Sub

Private Sub cmbRange_Click()
    DtTo = GetServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
        Case 0
            DtFrom = GetServerDate
        Case 1
            DtFrom = DateAdd("d", -7, GetServerDate)
        Case 2
            DtFrom = DateAdd("d", -15, GetServerDate)
        Case 3
            DtFrom = DateAdd("d", -30, GetServerDate)
        Case 4
            DtFrom.Enabled = True
            DtTo.Enabled = True
            DtFrom.SetFocus
            Exit Sub
    End Select
    
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
    FAcc.Enabled = True
End Sub

Private Sub cmdPrintPending_Click()

    Dim strSelection As String
    
    strSelection = "{@Balance}>0"
    
    If cmbArticle.MatchFound Then
        strSelection = strSelection & " AND {VVendRcvdDetailReWorkDetail.ItemCode}='" & cmbArticle.ID & "'"
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\PendingReWorks.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub cmdRefresh_Click()
    Call cmbArticle_matched
End Sub

Private Sub cmdSave_Click()
    
    On Error GoTo err
    If LV1.ListItems.count = 0 Then
        MsgBox "No Items Added.", vbInformation
        Exit Sub
    ElseIf cmbMaker.MatchFound = False Then
        MsgBox "No Maker Selected.", vbInformation
        Exit Sub
    End If
    
    Dim strItemID As String, lVendID As Long, lQty As Long, dRate As Double, lProcessID As Long, strOrderNo As String, iRepairType As Integer
    strItemID = cmbArticle.ID
    lVendID = cmbMaker.ID 'Val(LV1.ListItems(1).ListSubItems(3).Tag)
    dRate = 0 'Val(LV1.ListItems(1).SubItems(5))
    strOrderNo = LV1.ListItems(1).SubItems(2)
    iRepairType = (LV1.ListItems(1).ListSubItems(3).Tag)
    
'    Dim i As Integer
'    For i = 1 To LV1.ListItems.count
'        lQty = lQty + Val(LV1.ListItems(i).SubItems(5))
'    Next
    lProcessID = GetSingleLongValue("TOP 1 ProcessID", "RepairTypeProcesses", " WHERE Repair_RefID=" & Val(LV1.ListItems(1).ListSubItems(3).Tag) & " ORDER BY SeqNo")
    
    Dim lMakerIssNo As Long, strMakerIssNo As String
    
    
    Dim strIssIDDetail As String
   
         
    Dim lLotNo As String
'    lLotNo = getNextLotNo
        
    Dim lEntryIDDetail As Long, lRefID As Long
    Dim i As Integer
    Call StartTrans(con)
    
    For i = 1 To LV1.ListItems.count
        
        lMakerIssNo = Val(con.Execute("SELECT MAX(CAST(RIGHT(RecieptID,Len(RecieptID)-12) As Int)) FROM VendIssued WHERE Convert(varchar,DT,6)='" & Format(Date, "dd MMM yy") & "'").Fields(0).Value & "")
        lMakerIssNo = lMakerIssNo + 1
        strMakerIssNo = "M-ISU-" & Format(Date, "ddMMyy") & lMakerIssNo
        
        strIssIDDetail = Val(con.Execute("SELECT MAX(CAST(RIGHT(RecieptID,Len(RecieptID)-10) As Int)) FROM VMakerIssItems WHERE Convert(varchar,DT,6)='" & Format(Date, "dd MMM yy") & "'").Fields(0).Value & "")
              
        strIssIDDetail = "ISU-" & Format(Date, "ddMMyy") & (Val(strIssIDDetail) + 1)
       
        con.Execute "INSERT INTO VendIssued(VendID,DT,RecieptID,UserID,ProcessID,ItemID," & _
         "UserName,MachineName,SpecialInstructions,ExcessQtyPercentage," & _
         "MaximumRcvingsAgainstPO) VALUES(" & lVendID & ",'" & _
         Date & "','" & strMakerIssNo & "'," & UserID & "," & _
         lProcessID & ",'" & strItemID & "','" & CurrentUserName & "','" & _
         strComputerName & "','',100,100)"
         
        lRefID = Val(con.Execute("SELECT Max(EntryID) From VendIssued").Fields(0).Value & "")
        
        con.Execute "UPDATE VendIssued SET Authorized=1 WHERE EntryID=" & lRefID
          
        con.Execute "INSERT INTO VendIssdDetail(RefID,RecieptID,ItemCode" & _
         ",Rate,IssQty,ReqAuth,OrderNo,RcvProcessID,ReturnDT,Priority,Rcvd_RefID,LotNo,ReWorkLot,Repair_RefID,IPODL_RefID) VALUES(" & lRefID & ",'" & _
         strIssIDDetail & "','" & strItemID & "'," & Val(LV1.ListItems(i).SubItems(5)) & "," & Val(LV1.ListItems(i).SubItems(4)) & ",0,'" & strOrderNo & "'," & lProcessID & ",'" & _
         Date & "',1,0," & lLotNo & ",1," & iRepairType & "," & Val(LV1.ListItems(i).SubItems(1)) & ")"
    Next
    
    
    
'    If cmbVender.ID = iFactoryMakerID And (strCompany = "Weldon Instruments" Or strCompany = "QEL") Then
'        con.Execute "UPDATE VendIssued SET Authorized=1 WHERE EntryID=" & lRefID
'    End If
    
    
        
'    lEntryIDDetail = GetSingleLongValue("MAX(EntryID)", "VendIssdDetail")
'    For i = 1 To LV1.ListItems.count
'        With LV1.ListItems(i)
'            con.Execute "UPDATE VendRcvdDetailReWorkDetail SET IssQty=IssQty+" & Val(.SubItems(5)) & " WHERE EntryID=" & Val(.key)
'            con.Execute "INSERT INTO VendReWorkIssuanceDetail(VID_RefID,VRDRWD_RefID,Qty) VALUES(" & lEntryIDDetail & _
'             "," & Val(.key) & "," & Val(.SubItems(5)) & ")"
'        End With
'    Next
    
    con.CommitTrans
    
'    If chkPrint.Value = vbChecked Then
'        Dim rpt As CRAXDDRT.Report
'        Set rpt = rptApp.OpenReport(RptPath & "\PTCQel.rpt")
'        Dim f As New frmPrevRpt
'        f.ShowReport "{VFOrderItemPTC.LotNo}='" & lLotNo & "'", rpt
'    End If
    
'    lblStatus.Caption = "Lot #:" & lLotNo & " Just Generated."
'    Call cmbArticle_matched
    LV1.ListItems.Clear
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()

    iColNo = 0
    mnuPop.Visible = False
    cmbArticle.ListHeight = 5800
        
    cmbArticle.DropDownOffLine = True
    Call cmbArticle.AddVals(con, "'{ ' + ItemID + ' } ' + ItemName ", "Items", "ItemID", " WHERE InActive=0")
    cmbRange.ListIndex = 0
    
    
    cmbRepairType.AddItem "<All>"
    cmbRepairType.List(0, 1) = "0"
    Call AddToCombo(cmbRepairType, "RepairType", "RepairTypes", "", True, "EntryID")
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    FAcc.Top = (Me.ScaleHeight - FAcc.Height) / 2
    
    
    
End Sub


Private Sub LV_DblClick()

    'Check If Added Lot's Types match with this ones...
    'same Artice,order, Maker & Repair Type
    If LV1.ListItems.count > 0 Then
        If LV.SelectedItem.SubItems(2) <> LV1.ListItems(1).SubItems(2) Then
            MsgBox "Can't Add. Order No. Does not Match.", vbInformation
            Exit Sub
        ElseIf Val(LV.SelectedItem.ListSubItems(3).Tag) <> Val(LV1.ListItems(1).ListSubItems(3).Tag) Then
            MsgBox "Can't Add. Maker Does not Match.", vbInformation
            Exit Sub
        ElseIf Val(LV.SelectedItem.ListSubItems(4).Tag) <> Val(LV1.ListItems(1).ListSubItems(4).Tag) Then
            MsgBox "Can't Add. Repair Type Does not Match.", vbInformation
            Exit Sub
        End If
    End If
    
    Dim ITM As ListItem
    Set ITM = LV1.ListItems.add(, LV.SelectedItem.key, LV1.ListItems.count + 1)
    ITM.ListSubItems.add , , LV.SelectedItem.SubItems(1)
    ITM.ListSubItems.add , , LV.SelectedItem.SubItems(2)
    ITM.ListSubItems.add(, , LV.SelectedItem.SubItems(3)).Tag = LV.SelectedItem.ListSubItems(3).Tag
    ITM.ListSubItems.add(, , LV.SelectedItem.SubItems(4)).Tag = LV.SelectedItem.ListSubItems(4).Tag
    
    
    LV.ListItems.Remove LV.SelectedItem.Index
    
    If LV1.ListItems.count = 1 Then
        Dim lProcessID As Long, strItemID As String, lRepairType As Long
        lRepairType = Val(LV1.ListItems(1).ListSubItems(3).Tag)
        strItemID = cmbArticle.ID
        lProcessID = GetSingleLongValue("TOP 1 ProcessID", "RepairTypeProcesses", " WHERE Repair_RefID=" & lRepairType & " ORDER BY SeqNo")
        cmbMaker.ClearVals
        'cmbMaker.AddVals con, "'{' + VendID1 + '} ' + VenderName", "Makers", "VendID", " WHERE VendID IN(SELECT VendID FROM VendAssItems WHERE ProcessID=" & lProcessID & " AND ItemID='" & strItemID & "')"
        cmbMaker.AddVals con, "'{' + VendID1 + '} ' + VenderName", "Makers", "VendID", " WHERE VendID IN(SELECT VendID FROM MakerProcesses WHERE ProcessID=" & lProcessID & ")"
    End If
    
    GetLVTotal
    GetLV1Total
    
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If x > LV.ColumnHeaders(1).Left And x < LV.ColumnHeaders(i).Left + LV.ColumnHeaders(i).Width Then
            iColNo = i
            Exit For
        End If
    Next
    
End Sub

Private Sub LV1_DblClick()
    If LV1.ListItems.count = 0 Then Exit Sub
    With txtRate
        .Move LV1.Left + LV1.ColumnHeaders(6).Left, LV1.Top + LV1.SelectedItem.Top, LV1.ColumnHeaders(6).Width
        .Text = Val(LV1.SelectedItem.SubItems(5))
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txtRate) < 0 Then
            MsgBox "Invalid Rate.", vbInformation
            Exit Sub
        End If
        LV1.SelectedItem.SubItems(5) = Val(txtRate)
        txtRate.Visible = False
    ElseIf KeyAscii = 27 Then
        txtRate.Visible = False
    End If
End Sub

Private Sub txtRate_LostFocus()
    txtRate.Visible = False
End Sub
