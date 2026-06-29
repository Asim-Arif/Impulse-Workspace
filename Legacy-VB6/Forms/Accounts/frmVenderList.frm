VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmVenderList 
   ClientHeight    =   6795
   ClientLeft      =   1110
   ClientTop       =   240
   ClientWidth     =   10365
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   10365
   WindowState     =   2  'Maximized
   Begin ComboList.Usercontrol1 cmbVendName 
      Height          =   285
      Left            =   15
      TabIndex        =   6
      Top             =   735
      Width           =   7320
      _ExtentX        =   12912
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
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   570
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   10
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVenderList.frx":0000
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVenderList.frx":01BC
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1110
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVenderList.frx":0378
            Key             =   "Active"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVenderList.frx":07CC
            Key             =   "Inactive"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   15
      TabIndex        =   0
      Top             =   -90
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vender List"
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
         Index           =   0
         Left            =   4260
         TabIndex        =   1
         Top             =   120
         Width           =   1830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vender List"
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
         Index           =   1
         Left            =   4275
         TabIndex        =   2
         Top             =   135
         Width           =   1830
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   5010
      Left            =   15
      TabIndex        =   3
      Top             =   1050
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   8837
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   8388608
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
         Text            =   "Vender ID"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Vender Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Contact Person"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Phone 1"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Phone 2"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Balance"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Description"
         Object.Width           =   3528
      EndProperty
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   0
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
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   1
      Left            =   15
      TabIndex        =   7
      Top             =   510
      Width           =   7320
      BackColor       =   11517387
      Caption         =   " Vender"
      Size            =   "12912;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   375
      Left            =   8685
      TabIndex        =   5
      Top             =   6135
      Width           =   1665
      ForeColor       =   64
      Caption         =   "Close          "
      PicturePosition =   327683
      Size            =   "2937;661"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdOpt 
      Height          =   375
      Left            =   7005
      TabIndex        =   4
      Top             =   6135
      Width           =   1680
      Caption         =   "Options         "
      PicturePosition =   327683
      Size            =   "2963;661"
      Accelerator     =   79
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Menu mnupop 
      Caption         =   "mnupop"
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuShowAll 
         Caption         =   "Show &Inactive Venders"
      End
      Begin VB.Menu mnuLedger 
         Caption         =   "View &Ledger For Last 30 Days"
      End
      Begin VB.Menu mnuduePymts 
         Caption         =   "View Due Payments"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "Add &New Vender"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit Vender Information"
      End
      Begin VB.Menu mnuInActive 
         Caption         =   "Make This Vender Ina&ctive"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrchs 
         Caption         =   "Enter Purc&hase From This Vender"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPymt 
         Caption         =   "Enter Pa&yment To This Vender"
         Visible         =   0   'False
      End
      Begin VB.Menu dash3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintDetail 
         Caption         =   "Print Selected Vender's Detail"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintList 
         Caption         =   "&Print Vender List"
      End
      Begin VB.Menu mnuPrintVendListGroup 
         Caption         =   "Print Vender List ( Groupwise )"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintVendLabel 
         Caption         =   "Print This Vender Label"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintLabels 
         Caption         =   "Print Mailing Labels"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintDetailList 
         Caption         =   "Print Vender &Detailed List"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintVenderFileLable 
         Caption         =   "Print Vender File Lable"
      End
      Begin VB.Menu mnuPrintVenderProfile 
         Caption         =   "Print Vender Profile"
      End
      Begin VB.Menu mnuPrintVenderDetails 
         Caption         =   "Print Vender Details"
      End
      Begin VB.Menu mnuPrintVenderLable 
         Caption         =   "Print Vender Lable"
      End
      Begin VB.Menu mnuPrintVenderShippingLable 
         Caption         =   "Print Vender Shipping Lable"
      End
   End
End
Attribute VB_Name = "frmVenderList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public add As Boolean, edit As Boolean
Dim PermissionDenied As Boolean

Private Sub cmbVendName_matched()
    Call GoToSelected
End Sub

Private Sub GoToSelected()

    On Error Resume Next
    LV.ListItems(cmbVendName.ID & "'").EnsureVisible
    LV.ListItems(cmbVendName.ID & "'").Selected = True
    LV.SetFocus
    
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOpt_Click()
    With cmdOpt
        Call CheckMenu
        Me.PopUpMenu mnuPop, 255, .Left, .Top + .Height
    End With
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

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    LV.Sorted = True
    If LV.SortKey = ColumnHeader.Index - 1 Then
        Set LV.ColumnHeaderIcons = ImageList2
        If LV.SortOrder = lvwAscending Then
            LV.SortOrder = lvwDescending
            ColumnHeader.Icon = "Up"
        Else
            LV.SortOrder = lvwAscending
            ColumnHeader.Icon = "Down"
        End If
    Else
        LV.ColumnHeaderIcons = Nothing
        Set LV.ColumnHeaderIcons = ImageList2
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        ColumnHeader.Icon = "Down"
    End If

    On Error Resume Next
    LV.SelectedItem.EnsureVisible
    
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    Call mnuedit_Click
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()

    mnuPop.Visible = False
    'PermissionDenied = Not UserHasAccess("VenderList")
    Call RefreshList
    mnuAdd.Enabled = Not OutOfFinancialYear
    mnuEdit.Enabled = Not OutOfFinancialYear
    'mnuDelete.Enabled = Not OutOfFinancialYear
    
    cmbVendName.ListHeight = 2800
    
    'cmbVendName.AddVals con, "'{' + AccNo + '} ' + AccTitle", "VVenders", "VendID", " "
    cmbVendName.AddVals con, "'{' + MKNo + '} ' + AccTitle", "VVenders", "VendID", " "
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
 
    Dim TitleBottom As Long

    TitleBottom = fTitle.Top + fTitle.Height + 600
 
    LV.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + cmdClose.Height + 100)
 
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 50
    cmdOpt.Move cmdClose.Left - cmdOpt.Width - 75, cmdClose.Top
 
 
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    Call CheckMenu
    Me.PopUpMenu mnuPop
End Sub

Public Sub mnuAdd_Click()
    Dim f As New frmVender
    With f
        .add = True
        .Show 1
    End With
    mnuRefresh_Click
End Sub


Private Sub mnuClose_Click()
Unload Me
End Sub

Private Sub mnuduePymts_Click()
With frmDueRcptsPymts
   Call .FillCmbs(True)
   .Show
   .cmbAccNo = LV.SelectedItem
   .SetFocus
End With
End Sub

Private Sub mnuInActive_Click()

    If LV.ListItems.count = 0 Then Exit Sub

    If mnuInActive.Tag = "Active" Then
        If Val(LV.SelectedItem.SubItems(6)) <> 0 Then
            MsgBox "This Vender Cannot Be Inactive Unless Balance Becomes Zero.", vbInformation
            Exit Sub
        Else
            If MsgBox("Are You Sure To Make '" & LV.SelectedItem.SubItems(1) & "' Inactive", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            con.Execute "Update Accounts Set Active=0 Where Accno='" & LV.SelectedItem & "'"
        End If
    Else
        con.Execute "Update Accounts Set Active=1 Where Accno='" & LV.SelectedItem & "'"
    End If
    Call RefreshList
    
End Sub


Private Sub mnuedit_Click()

    Dim f As New frmVender
    With f
        .edit = True
        .Tag = Val(LV.SelectedItem.key)
        .Show 1
        mnuRefresh_Click
    End With
    
End Sub



Private Sub mnuLedger_Click()
    With frmLedger
        .cmbAccNo = LV.SelectedItem
        .cmdRefresh_Click
        .Show
    End With
End Sub

Private Sub mnuPrintDetail_Click()
   Screen.MousePointer = vbHourglass
   With cr1
      .ReportFileName = RptPath & "\VendDetailList.rpt"
      .Connect = ConnectStr
      .Formulas(0) = "Company='" & strCompany & "'"
      .Formulas(1) = "Address='" & COMPANYADDRESS & "'"
      
      .SelectionFormula = IIf(mnuShowAll.Checked, "", "{VVenders.Active}=True AND {VVenders.AccNo}='" & LV.SelectedItem & "'")
      .Action = 1
      .PageZoomNext
   End With
   Screen.MousePointer = vbDefault
End Sub


Private Sub mnuPrintDetailList_Click()
Screen.MousePointer = vbHourglass
   With cr1
   .ReportFileName = RptPath & "\VendDetailList.rpt"
   .Connect = ConnectStr
   .Formulas(0) = "Company='" & strCompany & "'"
   .Formulas(1) = "Address='" & COMPANYADDRESS & "'"
   .SelectionFormula = IIf(mnuShowAll.Checked, "", "{VVenders.Active}=True")
   .Action = 1
   .PageZoomNext
End With
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuPrintLabels_Click()
Screen.MousePointer = vbHourglass
With cr1
   .ReportFileName = RptPath & "\VendAddress.rpt"
   .Connect = ConnectStr
   .Formulas(0) = ""
   .Formulas(1) = ""
   .SelectionFormula = IIf(mnuShowAll.Checked, "", "{VVenders.Active}=True")
   .Action = 1
   .PageZoomNext
End With
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuPrintList_Click()
    
        
    Dim rpt As CRAXDDRT.Report, frm As New frmPrevRpt
    Set rpt = rptApp.OpenReport(RptPath & "\VenderList.rpt")
    
    frm.ShowReport "", rpt
    
    
    
End Sub

Private Sub mnuPrintVendLabel_Click()

    On Error GoTo err
    Screen.MousePointer = vbHourglass
    With cr1
       .ReportFileName = RptPath & "\VendLabel.rpt"
       .Connect = ConnectStr
       .Formulas(0) = ""
       .Formulas(1) = ""
       .SelectionFormula = IIf(mnuShowAll.Checked, "{VCustomers.AccNo}='" & LV.SelectedItem & "'", "{Vcustomers.Active}=True AND {VCustomers.AccNo}='" & LV.SelectedItem & "'")
       .Action = 1
       .PageZoomNext
    End With
    Screen.MousePointer = vbdefualt
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintVendListGroup_Click()
        
    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report, frm As New frmPrevRpt
    Set rpt = rptApp.OpenReport(rptAccountsPath & "\VenderListWithMaterialCats.rpt")
        
    frm.ShowReport IIf(mnuShowAll.Checked, "", "{VVenders.Active}=True"), rpt
    
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Sub mnuPymt_Click()
    Dim f As New frmCashPV
    With f
        .cmbAccNo = LV.SelectedItem
        .Show
    End With
End Sub

Private Sub mnuRefresh_Click()
   Call RefreshList
End Sub

Private Sub RefreshList()
    
    Dim bShowBalance As Boolean
    bShowBalance = GetSingleBooleanValue("Show_Balance_Vendor_List", "Users", " WHERE UserName='" & CurrentUserName & "'")
    
    Screen.MousePointer = vbHourglass

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim PrevIndex As Double
    Dim TheColor As Long
    Dim ActiveCondition As String
    Dim TheTag As String

    If LV.ListItems.count > 0 Then PrevIndex = LV.SelectedItem.Index
    Set LV.SmallIcons = ImageList1
    
    Dim strBalance As String
    With rs
        'Check For Show All Option
        'ActiveCondition = IIf(mnuShowAll.Checked, "", " Where Active=1")
    
        .CursorLocation = adUseClient
        .Open "SELECT * FROM VVenders " & IIf(mnuShowAll.Checked = False, " Where Active=1", "") & " Order By AccTitle", con, adOpenForwardOnly, adLockReadOnly
    
        LV.ListItems.Clear
    
        For i = 0 To .RecordCount - 1
            'Set Color For Inactive
            If !Active Then TheTag = "Active" Else TheTag = "Inactive"
            If LV.ListItems.count Mod 2 = 1 Then
                TheColor = vbBlue
            Else
                TheColor = vbBlack
            End If
            
            Set ITM = LV.ListItems.add(, ![VendID] & "'", !MKNo & "", , TheTag)
            ITM.ForeColor = TheColor
            ITM.Tag = TheTag
            ITM.ListSubItems.add(, , ![AccTitle] & "").ToolTipText = ![AccTitle] & ""
            ITM.ListSubItems(1).ForeColor = TheColor
            ITM.ListSubItems.add(, , ![ContactPerson] & "").ToolTipText = ![ContactPerson] & ""
            ITM.ListSubItems(2).ForeColor = TheColor
            ITM.ListSubItems.add(, , ![Phone1] & " ").ForeColor = TheColor
            ITM.ListSubItems.add(, , ![Phone2] & " ").ForeColor = TheColor
                    
            If bShowBalance Then
                strBalance = Format(Val(![Balance] & ""), "0.00 DR;0.00 CR;0.00     ")
            Else
                strBalance = "-"
            End If
            
            ITM.ListSubItems.add(, , strBalance).ForeColor = TheColor
            ITM.ListSubItems.add(, , !VenderDescription & "").ForeColor = TheColor
            
            .MoveNEXT
        Next
        .Close
    End With
    Set rs = Nothing

    If PrevIndex > 0 And PrevIndex <= LV.ListItems.count Then LV.ListItems(PrevIndex).Selected = True
   
    On Error Resume Next
    LV.SelectedItem.EnsureVisible
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub mnuPrchs_Click()
'    Dim f As New frmPurchaseJV
'    With f
'        .cmbVendAccNo = LV.SelectedItem
'        .Show
'    End With
End Sub

Private Sub mnuShowAll_Click()
    mnuShowAll.Checked = Not mnuShowAll.Checked
    Call RefreshList
End Sub
 
 Private Sub CheckMenu()
    
    mnuInActive.Enabled = False
    If OutOfFinancialYear Then
        mnuEdit.Enabled = False
        mnuLedger.Enabled = False
        mnuPrintList.Enabled = True
        mnuPrintDetail.Enabled = True
        mnuPrintDetailList.Enabled = True
        mnuAdd.Enabled = False
        mnuInActive.Enabled = False
        mnuPrchs.Enabled = False
        mnuPymt.Enabled = False
        mnuduePymts.Enabled = True
    ElseIf LV.ListItems.count = 0 Then
        mnuEdit.Enabled = False
        mnuLedger.Enabled = False
        mnuPrintList.Enabled = False
        mnuPrintDetail.Enabled = False
        mnuPrintDetailList.Enabled = False
        mnuAdd.Enabled = True
        mnuInActive.Enabled = True
        mnuPrchs.Enabled = False
        mnuPymt.Enabled = False
        mnuduePymts.Enabled = False
    Else
        mnuEdit.Enabled = True
        mnuLedger.Enabled = True
        mnuPrintList.Enabled = True
        mnuPrintDetail.Enabled = True
        mnuPrintDetailList.Enabled = True
        mnuAdd.Enabled = True
        
        mnuInActive.Caption = IIf(LV.SelectedItem.Tag = "Active", "Make This Vender Ina&ctive", "Make This Vender A&ctive")
        mnuInActive.Tag = LV.SelectedItem.Tag
        mnuInActive.Enabled = True
     
        mnuPrchs.Enabled = (mnuInActive.Tag = "Active")
        mnuPymt.Enabled = (mnuInActive.Tag = "Active")
        mnuduePymts.Caption = "View Due Payments From " & LV.SelectedItem.SubItems(1)
        mnuduePymts.Enabled = True
    End If
    mnuPrintLabels.Enabled = mnuPrintList.Enabled
    mnuPrintVendLabel.Enabled = mnuPrintList.Enabled
   
    mnuLedger.Enabled = False
    mnuduePymts.Enabled = False
End Sub


Private Sub mnuPrintVenderDetails_Click()
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Print_Vender_Details.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VMakers.VendID}=" & Val(LV.SelectedItem.key), rpt
End Sub

Private Sub mnuPrintVenderFileLable_Click()
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Print_Vender_File_Label.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VMakers.VendID}=" & Val(LV.SelectedItem.key), rpt
End Sub

Private Sub mnuPrintVenderLable_Click()
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Print_Vender_Lable.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VMakers.VendID}=" & Val(LV.SelectedItem.key), rpt
End Sub

Private Sub mnuPrintVenderProfile_Click()
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Print_Vender_Profile.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VMakers.VendID}=" & Val(LV.SelectedItem.key), rpt
End Sub

Private Sub mnuPrintVenderShippingLable_Click()
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Print_Vender_Shipping_Lable.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VMakers.VendID}=" & Val(LV.SelectedItem.key), rpt
End Sub


