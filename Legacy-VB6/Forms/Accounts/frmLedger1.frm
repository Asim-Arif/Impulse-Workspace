VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmLedger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleMode       =   0  'User
   ScaleWidth      =   10667.95
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   975
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLedger.frx":0000
            Key             =   "PND"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLedger.frx":0458
            Key             =   "CLR"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLedger.frx":08AC
            Key             =   "DUE"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLedger.frx":0D04
            Key             =   "BNC"
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   15
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Frame FDrag 
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   60
      MousePointer    =   7  'Size N S
      TabIndex        =   11
      Top             =   4155
      Width           =   11805
   End
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   705
      Left            =   15
      TabIndex        =   9
      Top             =   30
      Width           =   11190
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   7965
         TabIndex        =   3
         Top             =   345
         Width           =   1605
         _ExtentX        =   2831
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
         Format          =   24903683
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   6210
         TabIndex        =   2
         Top             =   345
         Width           =   1755
         _ExtentX        =   3096
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
         Format          =   24903683
         CurrentDate     =   37055
      End
      Begin MSForms.CommandButton cmdRefresh 
         Default         =   -1  'True
         Height          =   360
         Left            =   9615
         TabIndex        =   4
         Top             =   300
         Width           =   1500
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "2646;635"
         Accelerator     =   82
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         Caption         =   $"frmLedger.frx":1158
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
         Height          =   210
         Index           =   0
         Left            =   15
         TabIndex        =   10
         Top             =   120
         Width           =   11130
      End
      Begin MSForms.ComboBox CmbAccTitle 
         Height          =   315
         Left            =   1905
         TabIndex        =   1
         Top             =   345
         Width           =   4305
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "7594;556"
         MatchEntry      =   1
         ListStyle       =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbAccNo 
         Height          =   315
         Left            =   30
         TabIndex        =   0
         Top             =   345
         Width           =   1890
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3334;556"
         MatchEntry      =   1
         ListStyle       =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   3375
      Left            =   45
      TabIndex        =   5
      Top             =   765
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   5953
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   15199215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Dated"
         Object.Width           =   2363
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Voucher #"
         Object.Width           =   2363
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   6695
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Debit"
         Object.Width           =   2363
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Credit"
         Object.Width           =   2363
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Balance"
         Object.Width           =   2560
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Dpst Slip #"
         Object.Width           =   1969
      EndProperty
   End
   Begin MSComctlLib.ListView LVCheque 
      Height          =   2610
      Left            =   60
      TabIndex        =   6
      Top             =   4320
      Visible         =   0   'False
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   4604
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   15330543
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Dated"
         Object.Width           =   2363
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cheque #"
         Object.Width           =   2068
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bank & Branch"
         Object.Width           =   3939
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Description"
         Object.Width           =   4529
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Amount"
         Object.Width           =   2363
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Type"
         Object.Width           =   1969
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Chq. Date"
         Object.Width           =   2068
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Status"
         Object.Width           =   1575
      EndProperty
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   360
      Left            =   10380
      TabIndex        =   8
      Top             =   7365
      Width           =   1425
      ForeColor       =   64
      Caption         =   " Close      "
      PicturePosition =   327683
      Size            =   "2514;635"
      Picture         =   "frmLedger.frx":11F8
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   360
      Left            =   8895
      TabIndex        =   7
      Top             =   7365
      Width           =   1425
      ForeColor       =   64
      Caption         =   "Print         "
      PicturePosition =   327683
      Size            =   "2514;635"
      Picture         =   "frmLedger.frx":130A
      Accelerator     =   80
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Menu mnuPop 
      Caption         =   "Pop Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuToCash 
         Caption         =   "Transfer to &Cash In Hand"
      End
      Begin VB.Menu mnuToBank 
         Caption         =   "Transfer to Bank &Account"
         Begin VB.Menu mnuBanks 
            Caption         =   "< No Bank Account Defined >"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChqBounced 
         Caption         =   "Cheque &Bounced"
      End
   End
   Begin VB.Menu mnuclear 
      Caption         =   "mnuClear"
      Visible         =   0   'False
      Begin VB.Menu mnuChqCashed 
         Caption         =   "Make This Cheque &Cleared"
      End
      Begin VB.Menu mnuBounced 
         Caption         =   "Chque &Bounced"
      End
   End
End
Attribute VB_Name = "frmLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReadyToDrag As Boolean
Dim BankAccNo As String





Private Sub CmdClose_Click()
Unload Me
End Sub








Private Sub cmdPrint_Click()










Screen.MousePointer = vbHourglass

With cr1
    .ReportFileName = RptPath & "\ledger.rpt"
    .SubreportToChange = ""
    .DataFiles(0) = DBName
    .SelectionFormula = "{VLedger.AccNo}='" & cmbAccNo & "' AND ({VLedger.VDate}=date(" & Format(DtFrom.Tag, "yyyy,MM,dd") & ") TO date(" & Format(DtTo.Tag, "yyyy,MM,dd") & "))"
    .Formulas(0) = "Company='" & company & "'"
    .Formulas(1) = "FromTo='From " & Format(DtFrom.Tag, "dd-MMM-yyyy") & " To " & Format(DtTo.Tag, "dd-MMM-yyyy") & "'"
    .Formulas(2) = "SubReport=" & LVCheque.Visible
    'if Cheques Are Present then Show Cheque Detail Also
    If LVCheque.Visible Then
        .SubreportToChange = "ChqLedger"
        .DataFiles(0) = DBName
        .SelectionFormula = "{VChqLedger.AccNo}='" & cmbAccNo & "' AND ({VChqLedger.CDate}=date(" & Format(DtFrom.Tag, "yyyy,MM,dd") & ") TO date(" & Format(DtTo.Tag, "yyyy,MM,dd") & "))"
    End If
    .Action = 1
    .PageZoomNext
End With
Screen.MousePointer = vbDefault
End Sub



Public Sub cmdRefresh_Click()
    Call RefreshLedger
    cmbAccNo.Tag = cmbAccNo
    CmbAccTitle.Tag = CmbAccTitle
   Call Form_Resize
End Sub



Private Sub FDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReadyToDrag = True
    FDrag.BackColor = vbBlack
    FDrag.ZOrder
End Sub

Private Sub FDrag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim TheTop As Long
TheTop = FDrag.Top + Y

    If ReadyToDrag Then
        If (TheTop > (FAcc.Top + FAcc.Height + 500)) And (TheTop < cmdClose.Top - 500) Then
            FDrag.Top = TheTop
        End If
    End If

    
End Sub

Private Sub FDrag_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If ReadyToDrag Then
    
        FDrag.BackColor = Me.BackColor
        LV.Height = FDrag.Top - LV.Top
        LVCheque.Move 50, FDrag.Top + FDrag.Height, LVCheque.Width, Me.ScaleHeight - (FAcc.Height + cmdClose.Height + LV.Height + FDrag.Height + 100)
        If LVCheque.Top > cmdClose.Top - 500 Then
            LVCheque.Top = cmdClose.Top - 500
        End If
        
        ReadyToDrag = False
        
        FDrag.ZOrder 1
    End If
    
End Sub

Private Sub Form_Load()

DtTo = Date
DtFrom = DateAdd("M", -1, Date)
    Call AddToCombo(cmbAccNo, "AccNo", "Accounts", " Where Not Parent Order By AccNo")
    Call AddToCombo(CmbAccTitle, "AccTitle", "Accounts", " Where Not Parent Order By AccNo")
   
   Call GetBankAccountsInMenu
    
    
End Sub

Private Sub Form_Resize()

On Error Resume Next
Dim TheHeight As Long


 FAcc.Move (Me.ScaleWidth - FAcc.Width) / 2
cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 100
cmdPrint.Move cmdClose.Left - cmdPrint.Width - 50, cmdClose.Top

FDrag.Move 50, Me.ScaleHeight * 3 / 5, Me.ScaleWidth - 100
 
 If LVCheque.Visible Then
   TheHeight = FDrag.Top
   FDrag.Visible = True
 Else
   TheHeight = Me.ScaleHeight - cmdClose.Height - 200
   FDrag.Visible = False
 End If
 
 LV.Move 50, FAcc.Top + FAcc.Height, Me.ScaleWidth - 100, TheHeight - (FAcc.Top + FAcc.Height)
 LVCheque.Move 50, FDrag.Top + FDrag.Height, Me.ScaleWidth - 100, Me.ScaleHeight - (FAcc.Height + cmdClose.Height + LV.Height + 250)
 FDrag.ZOrder 1
 
End Sub




Private Sub cmbAccNo_Change()
    If cmbAccNo.MatchFound Then Call cmbAccNo_Click
End Sub

Private Sub cmbAccNo_Click()

    If CmbAccTitle.ListIndex <> cmbAccNo.ListIndex Then CmbAccTitle.ListIndex = cmbAccNo.ListIndex
    
End Sub

Private Sub CmbAccTitle_Change()
    If CmbAccTitle.MatchFound Then Call cmbAccTitle_Click
End Sub

Private Sub cmbAccTitle_Click()
    If cmbAccNo.ListIndex <> CmbAccTitle.ListIndex Then cmbAccNo.ListIndex = CmbAccTitle.ListIndex
End Sub


Private Sub RefreshLedger()
If Not cmbAccNo.MatchFound Then Exit Sub
Dim rsLedger As New ADODB.Recordset
SQL = "select * from VLedger where AccNo='" & cmbAccNo & "' and (VDate between #" & DtFrom & "# and #" & DtTo & "#) Order By VDate,SNo"

With rsLedger
    .CursorLocation = adUseClient
    .Open SQL, Con, adOpenForwardOnly, adLockReadOnly
    Dim itm As ListItem, Balance As String
 
    LV.ListItems.Clear
 
    If .RecordCount > 0 Then
        Balance = Val(![Balance] & "") + Val(![Credit] & "") - Val(![Debit] & "")
    Else
        Balance = Val(Con.Execute("Select Balance From Accounts Where AccNo='" & cmbAccNo & "'")![Balance] & "")
    End If
    
    Set itm = LV.ListItems.Add(, , "")
    itm.SubItems(2) = "          <<< Balance Forward >>>"
    If Val(Balance) > 0 Then
        Balance = Format(Abs(Val(Balance)), "###,##0.00") & "  DR"
    ElseIf Val(Balance) < 0 Then
        Balance = Format(Abs(Val(Balance)), "###,##0.00") & "  CR"
    Else
        Balance = Format(Val(Balance), "###,##0.00") & "       "
    End If
 
 
    itm.SubItems(5) = Balance
 
    For i = 0 To .RecordCount - 1
        Set itm = LV.ListItems.Add(, ![Sno] & "'", Format(![VDate], "dd-MMM-yyyy"))
        itm.SubItems(1) = ![VchrNo] & ""
        itm.SubItems(2) = ![Desc] & ""
        itm.SubItems(3) = Format(Val(![Debit] & ""), "###,##0.00")
        itm.SubItems(4) = Format(Val(![Credit] & ""), "###,##0.00")
        itm.Tag = Val(![CSNo] & "")
        
        'Specify Balance as DBT OR CDT
        If Val(![Balance] & "") > 0 Then
            Balance = Format(Abs(Val(![Balance] & "")), "###,##0.00") & "  DR"
        ElseIf Val(![Balance] & "") < 0 Then
            Balance = Format(Abs(Val(![Balance] & "")), "###,##0.00") & "  CR"
        Else
            Balance = Format(Val(![Balance] & ""), "###,##0.00") & "       "
        End If
        
        itm.SubItems(5) = Balance
        itm.SubItems(6) = ![DpstSlip]
        .MoveNext
        
    Next
    
End With


If LV.ListItems.Count > 1 Then
    Set itm = LV.ListItems.Add(, , "")
    itm.SubItems(2) = "             <<< Last Balance >>>"
    itm.SubItems(5) = LV.ListItems(LV.ListItems.Count - 1).SubItems(5)
End If

     DtTo.Tag = DtTo
     DtFrom.Tag = DtFrom
     
     
     '"(VDate" & Format(DtFrom.value, "(yyyy,MM,dd)") & " to Date" & Format(DtTo.value, "(yyyy,MM,dd)") & ") and ({ledgerrpt.acno}='" & Trim(Text16.Text) & "')"
'Now Fill The Other List With Cheques
Call RefreshChequeLedger

End Sub

Private Sub RefreshChequeLedger()
    Set LVCheque.SmallIcons = ImageList1
    Dim ListColor As Long
    Dim CrsLedger As New ADODB.Recordset
    
    With CrsLedger
        .CursorLocation = adUseClient
        .Open "Select * from VChqLedger Where (AccNo='" & cmbAccNo & "') " & _
            "and (CDate between #" & DtFrom & "# and #" & DtTo & "#) " & _
            "Order by CDate,Sno", Con, 0, 1
            
        LVCheque.ListItems.Clear
        If .RecordCount > 0 Then
            Dim itm As ListItem
            Do Until .EOF
                If ![Bounced] Then
                  ListColor = vbRed
                ElseIf (![ChequeDate] <= Date And Not ![Posted]) Then
                  ListColor = vbBlue
                ElseIf ![Posted] Then
                  ListColor = RGB(0, 125, 0)
                Else
                  ListColor = vbBlack
                End If
                
                Set itm = LVCheque.ListItems.Add(, ![Sno] & "'", Format(![CDate], "dd-MMM-yyyy"))
                itm.ForeColor = ListColor
                
                
                
                For i = 1 To 7
                    With itm.ListSubItems.Add(, , "")
                     .ForeColor = ListColor
                    End With
                    
                Next i
                
                itm.ListSubItems(1) = ![ChequeNo] & ""
                itm.ListSubItems(2) = ![Bank] & Space$(3) & ![Branch]
                itm.ListSubItems(3) = ![Desc]
                itm.ListSubItems(4) = Format(Val(![Amount] & ""), "###,##0.00")
                itm.ListSubItems(5) = ![ChequeType] & ""
                itm.ListSubItems(6) = Format(![ChequeDate], "dd-MMM-yyyy")
                
                'Check For The Chq Status
                Dim Status As String, img As String
                If ![Posted] Then
                  Status = "Cleared"
                  img = "CLR"
                ElseIf ![Bounced] Then
                  Status = "Bounced"
                  img = "BNC"
                ElseIf itm.ForeColor = vbBlue Then
                  Status = "Due"
                  img = "DUE"
                Else
                  Status = "Pending"
                  img = "PND"
                End If
                 
                 itm.ListSubItems(7) = Status
                 itm.SmallIcon = img
                 
                'Set Flag That it Was Payment Or Recept
                If ![Payment] Then itm.Tag = "P" Else itm.Tag = "R"
                
                
                
                .MoveNext
                LVCheque.Visible = True
                FDrag.Visible = True
            Loop
            
        Else
            LVCheque.Visible = False
            FDrag.Visible = False
        End If
        .Close
    End With
    
    Set CrsLedger = Nothing
End Sub





Private Sub LV_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
If Not LVCheque.Visible Then Exit Sub
 If Val(ITEM.Tag) = 0 Then
   LVCheque.HideSelection = True
 Else
    LVCheque.HideSelection = False
    LVCheque.ListItems(Val(ITEM.Tag) & "'").Selected = True
 End If
 
 
 
 
End Sub

Private Sub LVCheque_ItemCheck(ByVal ITEM As MSComctlLib.ListItem)
 ITEM.Checked = Not ITEM.Checked
End Sub

Private Sub LVCheque_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With LVCheque
    
        If .ListItems.Count = 0 Or Button <> 2 Then Exit Sub
        
        If .SelectedItem.ForeColor <> vbBlue Then Exit Sub
        
        If .SelectedItem.ListSubItems(5).Text = "Open" Then
            mnuToCash.Visible = True
        Else
            mnuToCash.Visible = False
        End If
        If .SelectedItem.Tag = "R" Then         ' Receipt Chq
         Me.PopupMenu mnuPop
        ElseIf .SelectedItem.Tag = "P" Then         ' Payment Chq
          Me.PopupMenu mnuclear
        End If
    End With
    
End Sub

Private Function SaveReceiptTrans(ReceiveInAccNo As String, Optional SlipNo As String) As Boolean

    Dim i As Integer
    Dim NextVchrNo As String
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim dAmount As Double
    Dim sScNo As String
    Dim sCIH As String
    Dim Bal As Double
    Dim Desc As String
    'Get Serial # and Balances
    NextVchrNo = getNextVchrNo(Date, "RV")
    INextSno = getNextSno(Date)
    SnoToUse = INextSno
    Bal = getBalance(cmbAccNo.Tag, Date)
    dAmount = Format$(LVCheque.SelectedItem.ListSubItems(4).Text, "###")
    sScNo = Left$(LVCheque.SelectedItem.key, Len(LVCheque.SelectedItem.key) - 1)
    sCIH = ReceiveInAccNo
    Bal = Bal - dAmount
    'Start New Transaction
    Con.BeginTrans
    'Credit Entry
    Con.Execute "insert into Vouchers Values(" & INextSno & ",#" & _
        Date & "#,'" & NextVchrNo & "','" & cmbAccNo.Tag & _
        "','" & LVCheque.SelectedItem.ListSubItems(3).Text & _
        "',0," & dAmount & "," & Bal & ",'" & SlipNo & "'," & sScNo & ")"
    
    
    
    Call adjustLedger(cmbAccNo.Tag, Date, INextSno, -Val(dAmount))
    
    Dim ChqNo As String
    INextSno = INextSno + 1
    Bal = getBalance(sCIH, Date)
    Bal = Bal + dAmount
    'Description
    Dim ChqRec As New ADODB.Recordset
    With ChqRec
        .CursorLocation = adUseClient
        .Open "Select ChequeNo From Cheque Where (Sno=" & sScNo & ")", Con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            Desc = "Cheque Recieved from " & CmbAccTitle.Tag & " Chq# : " & _
                .Fields(0) & ""
             ChqNo = .Fields(0) & ""
        Else
            Desc = "Cheque Recieved from " & CmbAccTitle.Tag & " Chq# : "
        End If
        
        .Close
    End With
    Set ChqRec = Nothing
    'Debit Entry
    
    Con.Execute "insert into Vouchers Values(" & INextSno & ",#" & _
        Date & "#,'" & NextVchrNo & "','" & sCIH & _
        "','" & Desc & _
        "'," & dAmount & ",0," & Bal & ",'" & SlipNo & "'," & sScNo & ")"
        
    
    
    Call adjustLedger(sCIH, Date, INextSno, Val(dAmount))
    
    Dim rsCopy As New ADODB.Recordset
    With rsCopy
        .CursorLocation = adUseClient
        .Open "Select * from ChequeDetail Where (Sno=" & sScNo & ")", Con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            Con.Execute "Insert into VoucherDetail Values('" & SnoToUse & "','" & ![BalancedSno] & "')"
            .MoveNext
        Loop
        .Close
    End With
    Set rsCopy = Nothing
    Con.Execute "Update Cheque Set Posted=True Where Sno=" & sScNo
    'Commit Transaction
    Con.CommitTrans
    
    
   
    
    SaveReceiptTrans = True
    Exit Function

err:
    MsgBox "Transaction Terminated Due To Following Error;" & vbNewLine & err.Description
    Con.RollbackTrans
        SaveReceiptTrans = False
End Function

Private Sub mnuBanks_Click(Index As Integer)

If MsgBox("Are You Sure To Transfer This Cheque To " & mnuBanks(Index).Caption, vbInformation + vbYesNo) = vbNo Then Exit Sub
   Dim SlipNo As String
   SlipNo = InputBox("Enter Deposit Slip No.")
   
   If Trim(SlipNo) = "" Then
      MsgBox "Transaction Terminated Due To Invlalid Bank Deposit SlipNo", vbInformation
      Exit Sub
   End If
   
   
   Screen.MousePointer = vbHourglass
   If SaveReceiptTrans(mnuBanks(Index).Tag, SlipNo) Then
      Sleep (1000)
      Call RefreshLedger
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub mnuBounced_Click()
If MsgBox("Are You Sure To Make This Cheque Bounced.", vbInformation + vbYesNo) = vbNo Then Exit Sub
Screen.MousePointer = vbHourglass
   Con.Execute "Update Cheque Set Bounced=True Where Sno=" & Val(LVCheque.SelectedItem.key)
   Sleep (1000)
   Call RefreshChequeLedger
   Screen.MousePointer = vbDefault
End Sub

Private Sub mnuChqBounced_Click()
   If MsgBox("Are You Sure To Make This Cheque Bounced.", vbInformation + vbYesNo) = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
      Con.Execute "Update Cheque Set Bounced=True Where Sno=" & Val(LVCheque.SelectedItem.key)
      Sleep (1000)
      Call RefreshChequeLedger
      Screen.MousePointer = vbDefault
End Sub

Private Sub mnuChqCashed_Click()
If MsgBox("Are You Sure To Make This Cheque Cleared.", vbInformation + vbYesNo) = vbNo Then Exit Sub

   Screen.MousePointer = vbHourglass
   If SavePaymentTrans Then
      Sleep (1000)
      Call RefreshLedger
   End If
       
   Screen.MousePointer = vbDefault
   
   
End Sub

Private Sub mnuToCash_Click()
If MsgBox("Are You Sure To Transfer This Cheque To Cash In Hand.", vbInformation + vbYesNo) = vbNo Then Exit Sub
   Screen.MousePointer = vbHourglass
   If SaveReceiptTrans(CashInHand) Then
      Sleep (1000)
      Call RefreshLedger
   End If
   Screen.MousePointer = vbDefault
End Sub



Private Sub GetBankAccountsInMenu()
  Dim rs As New ADODB.Recordset
 With rs
  .Open "Select * from VBankAccounts Order By AccNo", Con, adOpenForwardOnly, adLockReadOnly
  
  For i = 0 To .RecordCount - 1
   Load mnuBanks(i + 1)
   mnuBanks(i + 1).Caption = "[" & ![AccNo] & "] " & ![AccTitle]
   mnuBanks(i + 1).Enabled = True
   mnuBanks(i + 1).Tag = ![AccNo] & ""
  .MoveNext
  Next
  mnuBanks(0).Visible = .RecordCount = 0
  .Close
 End With
 Set rs = Nothing
  
End Sub






Private Function SavePaymentTrans() As Boolean

    Dim i As Integer
    Dim NextVchrNo As String
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim dAmount As Double
    Dim sScNo As String
    Dim sCIH As String
    Dim Bal As Double
    Dim Desc As String
    'Get Serial # and Balances
    NextVchrNo = getNextVchrNo(Date, "RV")
    INextSno = getNextSno(Date)
    SnoToUse = INextSno
    Bal = getBalance(cmbAccNo.Tag, Date)
    dAmount = Format$(LVCheque.SelectedItem.ListSubItems(4).Text, "###")
    sScNo = Left$(LVCheque.SelectedItem.key, Len(LVCheque.SelectedItem.key) - 1)
    
    
    
    sCIH = Con.Execute("Select AccNo From VPayChqs Where Sno=" & Val(LVCheque.SelectedItem.key)).Fields(0)
    
       
    
    Bal = Bal + dAmount
    'Start New Transaction
    Con.BeginTrans
    'Debit Entry
    Con.Execute "insert into Vouchers Values(" & INextSno & ",#" & _
        Date & "#,'" & NextVchrNo & "','" & cmbAccNo.Tag & _
        "','" & LVCheque.SelectedItem.ListSubItems(3).Text & _
        "'," & dAmount & ",0," & Bal & ",''," & sScNo & ")"
    
    
    
    Call adjustLedger(cmbAccNo.Tag, Date, INextSno, Val(dAmount))
    
    Dim ChqNo As String
    INextSno = INextSno + 1
    Bal = getBalance(sCIH, Date)
    Bal = Bal - dAmount
    'Description
    Dim ChqRec As New ADODB.Recordset
    With ChqRec
        .CursorLocation = adUseClient
        .Open "Select ChequeNo From Cheque Where (Sno=" & sScNo & ")", Con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            Desc = "Cheque Paid To " & CmbAccTitle.Tag & " Chq# : " & _
                .Fields(0) & ""
             ChqNo = .Fields(0) & ""
        Else
            Desc = "Cheque Paid To " & CmbAccTitle.Tag & " Chq# : "
        End If
        
        .Close
    End With
    Set ChqRec = Nothing
    'Credit Entry
    
    Con.Execute "insert into Vouchers Values(" & INextSno & ",#" & _
        Date & "#,'" & NextVchrNo & "','" & sCIH & _
        "','" & Desc & _
        "',0," & dAmount & "," & Bal & ",''," & sScNo & ")"
        
    
    
    Call adjustLedger(sCIH, Date, INextSno, -Val(dAmount))
    
    Dim rsCopy As New ADODB.Recordset
    With rsCopy
        .CursorLocation = adUseClient
        .Open "Select * from ChequeDetail Where (Sno=" & sScNo & ")", Con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            Con.Execute "Insert into VoucherDetail Values('" & SnoToUse & "','" & ![BalancedSno] & "')"
            .MoveNext
        Loop
        .Close
    End With
    Set rsCopy = Nothing
    Con.Execute "Update Cheque Set Posted=True Where Sno=" & sScNo
    'Commit Transaction
    Con.CommitTrans
   
    
    
    SavePaymentTrans = True
    Exit Function
err:
    MsgBox "Transaction Terminated Due To Following Error;" & vbNewLine & err.Description
    Con.RollbackTrans
        SavePaymentTrans = False
End Function



Private Sub PrintLedger()
With cr1
        

End With
End Sub
