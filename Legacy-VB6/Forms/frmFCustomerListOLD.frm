VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmFCustomerListOLD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Foreign Customers"
   ClientHeight    =   6795
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   10365
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   10365
   WindowState     =   2  'Maximized
   Begin ComboList.Usercontrol1 cmbCust 
      Height          =   285
      Left            =   195
      TabIndex        =   6
      Top             =   750
      Width           =   2310
      _ExtentX        =   4075
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
   Begin Crystal.CrystalReport cr1 
      Left            =   6735
      Top             =   615
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
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   375
      Top             =   5835
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
            Picture         =   "frmFCustomerListOLD.frx":0000
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFCustomerListOLD.frx":01BC
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   510
      Top             =   6225
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
            Picture         =   "frmFCustomerListOLD.frx":0378
            Key             =   "Active"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFCustomerListOLD.frx":07CC
            Key             =   "Inactive"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   180
      TabIndex        =   2
      Top             =   -75
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Foreign Customers"
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
         Left            =   3705
         TabIndex        =   3
         Top             =   120
         Width           =   2940
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Foreign Customers"
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
         Left            =   3720
         TabIndex        =   4
         Top             =   135
         Width           =   2940
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4845
      Left            =   45
      TabIndex        =   0
      Top             =   1215
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   8546
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "ImageList2"
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
         Text            =   "Cust Code"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Customer Name"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Country"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Currency"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Balance"
         Object.Width           =   2293
      EndProperty
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   4440
      TabIndex        =   9
      Top             =   675
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
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Customer                                Country"
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
      Height          =   240
      Index           =   2
      Left            =   195
      TabIndex        =   8
      Top             =   525
      Width           =   4215
   End
   Begin MSForms.ComboBox cmbCountry 
      Height          =   285
      Left            =   2490
      TabIndex        =   7
      Top             =   750
      Width           =   1920
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3387;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdOpt 
      Height          =   390
      Left            =   7155
      TabIndex        =   1
      Top             =   6075
      Width           =   1560
      Caption         =   "Options         "
      PicturePosition =   327683
      Size            =   "2752;688"
      Accelerator     =   79
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   390
      Left            =   8745
      TabIndex        =   5
      Top             =   6075
      Width           =   1560
      ForeColor       =   0
      Caption         =   "Close            "
      PicturePosition =   327683
      Size            =   "2752;688"
      Picture         =   "frmFCustomerListOLD.frx":0C20
      Accelerator     =   67
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
      Visible         =   0   'False
      Begin VB.Menu mnuRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuShowAll 
         Caption         =   "Show &Inactive Customers"
      End
      Begin VB.Menu mnuLedger 
         Caption         =   "View &Ledger For Last 30 Days"
      End
      Begin VB.Menu mnuDueRcpts 
         Caption         =   "View Due Receipts"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "Add &New Customer"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit Customer Information"
      End
      Begin VB.Menu mnuInActive 
         Caption         =   "Make This Customer Ina&ctive"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSales 
         Caption         =   "Enter &Sales To This Customer"
      End
      Begin VB.Menu mnuRcpt 
         Caption         =   "&Receive Payments From This Customer"
      End
      Begin VB.Menu dash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintDetail 
         Caption         =   "Print This Customer's Detail"
      End
      Begin VB.Menu mnuPrintList 
         Caption         =   "&Print Customer List"
      End
      Begin VB.Menu mnuPrintLabels 
         Caption         =   "Print Mailing Labels"
      End
      Begin VB.Menu mnuPrintDetailList 
         Caption         =   "Print Customer &Detailed List"
      End
   End
End
Attribute VB_Name = "frmFCustomerListOLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Add As Boolean, Edit As Boolean
Dim PermissionDenied As Boolean

Private Sub cmbCust_matched()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    SQL = "Select Distinct Country From ForeignCustomers "
    If cmbCust.ID <> "0" Then
        SQL = SQL & " Where CustCode='" & cmbCust.ID & "'"
    End If
    
    With rs
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        cmbCountry.AddItem "<All Countries>"
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            'cmbCountry.List(cmbCountry.ListCount - 1, 1) = ![Curr] & ""
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOpt_Click()
   
   With cmdOpt
      Call CheckMenu
      Me.PopupMenu mnuPop, 255, .Left, .Top + .Height
   End With
   
End Sub

Private Sub cmdRefresh_Click()

    Dim Cond As String
    If cmbCust.ID <> "0" Then
        Cond = "  CustCode='" & cmbCust.ID & "' "
    Else
        Cond = ""
    End If
    
    If cmbCountry.ListIndex <> 0 Then
        If Cond = "" Then
            Cond = "  Country='" & cmbCountry & "'"
        Else
            Cond = Cond & " and Country='" & cmbCountry & "'"
        End If
    End If
        
    Call RefreshList(Cond)
    
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

End Sub

Private Sub Lv_DblClick()
If LV.ListItems.Count = 0 Then Exit Sub

Call mnuEdit_Click
    
End Sub




Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()

'PermissionDenied = Not UserHasAccess("CustomerList")
    cmbCust.ListHeight = 2500
    cmbCust.AddItem "<All Customers>", "0"
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode")
    
  
    Call RefreshList
  'mnuAdd.Enabled = Not OutOfFinancialYear And UserHasAccess("NewCustomer")
  'mnuEdit.Enabled = Not OutOfFinancialYear And UserHasAccess("EditCustomer")
  
'mnuDelete.Enabled = Not OutOfFinancialYear
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
 
    Dim TitleBottom As Long
    Dim OtherHeight As Long
    TitleBottom = fTitle.Top + fTitle.Height
    OtherHeight = cmbCust.Height + Label1(2).Height + 50
    
    LV.Move 50, TitleBottom + OtherHeight, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + cmdClose.Height + 100 + OtherHeight)
 
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 50
    cmdOpt.Move cmdClose.Left - cmdOpt.Width - 75, cmdClose.Top
 
 
End Sub






Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    Call CheckMenu
    Me.PopupMenu mnuPop
End Sub

Public Sub mnuAdd_Click()
   Load frmFCustomer
   
   frmFCustomer.Show 1
End Sub


Private Sub mnuClose_Click()
Unload Me
End Sub

Private Sub mnuDueRcpts_Click()
With frmDueRcptsPymts
   Call .FillCmbs
   .Show
   .cmbAccNo = LV.SelectedItem
   .SetFocus
End With
End Sub

Private Sub mnuInActive_Click()
If LV.ListItems.Count = 0 Then Exit Sub

If mnuInActive.Tag = "Active" Then
    'If Val(LV.SelectedItem.SubItems(6)) <> 0 Then
    '    MsgBox "This Customer Cannot Be Inactive Unless Balance Becomes Zero.", vbInformation
    '    Exit Sub
    'Else
        If MsgBox("Are You Sure To Make '" & LV.SelectedItem.SubItems(1) & "' Inactive", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        con.Execute "Update ForeignCustomers Set Active=0 Where CustCode='" & LV.SelectedItem & "'"
    'End If
Else
    con.Execute "Update FForeignCustomers Set Active=1 Where CustCode='" & LV.SelectedItem & "'"
End If
  
  Call RefreshList
End Sub


Private Sub mnuEdit_Click()
'Check For User Access

'If Not UserHasAccess("EditCustomer") Then
'   MsgBox "You Donot Have Permissions To Edit Customer.", vbInformation
'   Exit Sub
'End If

Load frmFCustomer
frmFCustomer.EditCustomer (LV.SelectedItem)
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
With Cr1
   .ReportFileName = RptPath & "\CustDetailList.rpt"
   .Connect = ConnectStr
   .Formulas(0) = "Company='" & company & "'"
   .Formulas(1) = "Address='" & COMPANYADDRESS & "'"
   .SelectionFormula = IIf(mnuShowAll.Checked, "", "{Vcustomers.Active}=True AND {VCustomers.AccNo}='" & LV.SelectedItem & "'")
   .Action = 1
   .PageZoomNext
End With
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuPrintDetailList_Click()
Screen.MousePointer = vbHourglass
   With Cr1
   .ReportFileName = RptPath & "\CustDetailList.rpt"
   .Connect = ConnectStr
   .Formulas(0) = "Company='" & company & "'"
   .Formulas(1) = "Address='" & COMPANYADDRESS & "'"
   
   .SelectionFormula = IIf(mnuShowAll.Checked, "", "{Vcustomers.Active}=True")
   .Action = 1
   .PageZoomNext
End With
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuPrintLabels_Click()
Screen.MousePointer = vbHourglass
With Cr1
   .ReportFileName = RptPath & "\CustAddress.rpt"
   .Connect = ConnectStr
   .Formulas(0) = ""
   .Formulas(1) = ""
   .SelectionFormula = IIf(mnuShowAll.Checked, "", "{Vcustomers.Active}=True")
   .Action = 1
   .PageZoomNext
End With
Screen.MousePointer = vbdefualt
End Sub

Private Sub mnuPrintList_Click()
Screen.MousePointer = vbHourglass
With Cr1
   .ReportFileName = RptPath & "\CustomerList.rpt"
   .Connect = ConnectStr
   .Formulas(0) = "Company='" & company & "'"
   .Formulas(1) = "Address='" & COMPANYADDRESS & "'"
   .SelectionFormula = IIf(mnuShowAll.Checked, "", "{Vcustomers.Active}=True")
   .Action = 1
   .PageZoomNext
End With
Screen.MousePointer = vbdefualt
End Sub

Private Sub mnuRcpt_Click()
'Dim f As New frmCashRV
With f
    .cmbAccNo = LV.SelectedItem
    .Show
End With
End Sub

Private Sub mnuRefresh_Click()
   Call RefreshList
End Sub




Private Sub RefreshList(Optional Cond As String)


Screen.MousePointer = vbHourglass

Dim rs As New ADODB.Recordset
Dim ITM As ListItem
Dim PrevIndex As Double
Dim TheColor As Long
Dim ActiveCondition As String


If LV.ListItems.Count > 0 Then PrevIndex = LV.SelectedItem.Index
Set LV.SmallIcons = ImageList1
With rs
    'Check For Show All Option
    ActiveCondition = IIf(mnuShowAll.Checked, "", " Where Active=1")
    
    If Cond <> "" Then
        If ActiveCondition = "" Then
            ActiveCondition = " Where " & Cond
        Else
            ActiveCondition = ActiveCondition & " and " & Cond
        End If
    End If
    
    .CursorLocation = adUseClient
    .Open "Select * From ForeignCustomers" & ActiveCondition, con, adOpenForwardOnly, adLockReadOnly
    
    LV.ListItems.Clear
    
    For i = 0 To .RecordCount - 1
         'Set Color For Inactive
        If Not ![Active] Then
            TheColor = &H808080
            TheTag = "Inactive"
        Else
            TheColor = &H800000
            TheTag = "Active"
        End If
        
        Set ITM = LV.ListItems.Add(, ![CustCode] & "'", ![CustCode], , TheTag)
        ITM.ForeColor = TheColor
        ITM.Tag = TheTag
        ITM.ListSubItems.Add(, , ![Name] & "").ToolTipText = ![Name] & ""
        ITM.ListSubItems(1).ForeColor = TheColor
        ITM.ListSubItems.Add(, , ![Country] & "").ToolTipText = ![Country] & ""
        ITM.ListSubItems(2).ForeColor = TheColor
        ITM.ListSubItems.Add(, , ![Curr] & " ").ForeColor = TheColor
        'itm.ListSubItems.Add(, , ![Cont1Name] & " ").ForeColor = TheColor
        'itm.ListSubItems.Add(, , ![Cont1Phone] & " ").ForeColor = TheColor
        'itm.ListSubItems.Add(, , Format(Val(![Balance] & ""), "0.00 DR;0.00 CR;0.00     ")).ForeColor = TheColor
        .MoveNext
    Next
    .Close
End With
Set rs = Nothing


If PrevIndex > 0 And PrevIndex <= LV.ListItems.Count Then LV.ListItems(PrevIndex).Selected = True
   
    On Error Resume Next
    LV.SelectedItem.EnsureVisible
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuSales_Click()
'Dim f As New frmSaleJV
With f
    .cmbCustAccNo = LV.SelectedItem
    .Show
End With
End Sub

Private Sub mnuShowAll_Click()
mnuShowAll.Checked = Not mnuShowAll.Checked
Call RefreshList
End Sub
 
 
 
 Private Sub CheckMenu()
 If OutOfFinancialYear Then
     mnuEdit.Enabled = False
     mnuLedger.Enabled = False
     mnuPrintList.Enabled = True
     mnuPrintDetail.Enabled = True
     mnuPrintDetailList.Enabled = True
     mnuAdd.Enabled = False
     mnuInActive.Enabled = False
     mnuSales.Enabled = False
     mnuRcpt.Enabled = False
     mnuDueRcpts.Enabled = True
ElseIf LV.ListItems.Count = 0 Then
     mnuEdit.Enabled = False
     mnuLedger.Enabled = False
     mnuPrintList.Enabled = False
     mnuPrintDetailList.Enabled = False
     mnuPrintDetail.Enabled = False
     
     mnuAdd.Enabled = True
     mnuInActive.Enabled = True
     mnuSales.Enabled = False
     mnuRcpt.Enabled = False
     mnuDueRcpts.Enabled = False
Else
     mnuEdit.Enabled = True
     mnuLedger.Enabled = True
     mnuPrintList.Enabled = True
     mnuPrintDetail.Enabled = True
     mnuPrintDetailList.Enabled = True
     mnuAdd.Enabled = True
     
     mnuInActive.Caption = IIf(LV.SelectedItem.Tag = "Active", "Make This Customer Ina&ctive", "Make This Customer A&ctive")
     mnuInActive.Tag = LV.SelectedItem.Tag
     mnuInActive.Enabled = True
     mnuDueRcpts.Caption = "View Due Receipts From " & LV.SelectedItem.SubItems(1)
     mnuDueRcpts.Enabled = True
     mnuSales.Enabled = (mnuInActive.Tag = "Active")
     mnuRcpt.Enabled = (mnuInActive.Tag = "Active")

End If
mnuPrintLabels.Enabled = mnuPrintList.Enabled
 End Sub
