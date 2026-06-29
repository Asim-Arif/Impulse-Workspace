VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmCustomInvoiceAuth 
   ClientHeight    =   7665
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   12750
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   7665
   ScaleMode       =   0  'User
   ScaleWidth      =   11420.35
   WindowState     =   2  'Maximized
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   7665
      Left            =   45
      TabIndex        =   1
      Top             =   -45
      Width           =   12645
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
         ItemData        =   "frmCustomInvoiceAuth.frx":0000
         Left            =   4455
         List            =   "frmCustomInvoiceAuth.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   405
         Width           =   2010
      End
      Begin MSComCtl2.DTPicker DTVchr 
         Height          =   285
         Left            =   6930
         TabIndex        =   10
         Top             =   7275
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   152240131
         CurrentDate     =   38594
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   2580
         Top             =   2070
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
               Picture         =   "frmCustomInvoiceAuth.frx":0078
               Key             =   "Down"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCustomInvoiceAuth.frx":0234
               Key             =   "Up"
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2475
         TabIndex        =   5
         Top             =   1605
         Visible         =   0   'False
         Width           =   3615
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   195
         TabIndex        =   3
         Top             =   405
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
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2010
         Top             =   2070
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin MSComctlLib.ListView LV 
         Height          =   6390
         Left            =   60
         TabIndex        =   6
         Top             =   765
         Width           =   12495
         _ExtentX        =   22040
         _ExtentY        =   11271
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Country"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Invoice #"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Invoice DT"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Invoice Amt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Exch. Rate"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Amt PK"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   8205
         TabIndex        =   14
         Top             =   405
         Width           =   1560
         _ExtentX        =   2752
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
         Format          =   152240131
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   6480
         TabIndex        =   15
         Top             =   405
         Width           =   1710
         _ExtentX        =   3016
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
         Format          =   152240131
         CurrentDate     =   37055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sales Account :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   7290
         Width           =   1275
      End
      Begin MSForms.ComboBox CmbAccTitle 
         Height          =   315
         Left            =   1710
         TabIndex        =   11
         Top             =   7260
         Width           =   3810
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "6720;556"
         ListWidth       =   8819
         ListStyle       =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label2 
         Caption         =   "Voucher Date :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5595
         TabIndex        =   9
         Top             =   7305
         Width           =   1275
      End
      Begin MSForms.CommandButton cmdUpdate 
         Height          =   360
         Left            =   8370
         TabIndex        =   8
         Top             =   7230
         Width           =   2835
         Caption         =   "Authorize Selected Invoices"
         PicturePosition =   327683
         Size            =   "5001;635"
         Accelerator     =   85
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
         Left            =   11220
         TabIndex        =   7
         Top             =   7215
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
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   2550
         TabIndex        =   4
         Top             =   405
         Width           =   1860
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3281;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   9840
         TabIndex        =   0
         Top             =   375
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
         BackStyle       =   0  'Transparent
         Caption         =   $"frmCustomInvoiceAuth.frx":03F0
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
         Left            =   165
         TabIndex        =   2
         Top             =   180
         Width           =   9570
      End
   End
End
Attribute VB_Name = "frmCustomInvoiceAuth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strExchDiffAccNo As String

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
            .MoveNEXT
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

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub



Private Sub DtFrom_Change()
    'Call setDateRange(DtFrom)
End Sub

Private Sub DtTo_Change()
    'Call setDateRange(DtTo)
End Sub


Private Sub cmdUpdate_Click()

    On Error GoTo err
    If CmbAccTitle.MatchFound = False Then
        MsgBox "Select Account No.", vbInformation
        Exit Sub
    End If
    Dim strVchrNo As String
    Call StartTrans(con)
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                strVchrNo = AddVoucher(Val(.ListSubItems(5).Tag), Val(.SubItems(6)), Val(.SubItems(7)), .ListSubItems(1).Tag, CmbAccTitle.List(CmbAccTitle.ListIndex, 1), .SubItems(3) & "," & .SubItems(5) & " @ " & .SubItems(6) & "," & .SubItems(1) & " " & .SubItems(2), .SubItems(3))
                con.Execute "INSERT INTO CustomInvoiceAuthorization(CustomInvoice,VchrNo,AmtInPakRs,ExchRate,UserName,MachineName) " & _
                "Values('" & .SubItems(3) & "','" & strVchrNo & "'," & Val(.SubItems(7)) & "," & Val(.SubItems(6)) & ",'" & CurrentUserName & _
                "','" & strComputerName & "')"
            End If
        End With
    Next
    con.CommitTrans
    Call RefreshLV
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub

Private Function AddVoucher(dAmtF As Double, dExchRate As Double, dAmt As Double, CustAccNo As String, FSAccNo As String, Desc As String, strInvNo As String) As String
    
    
    Dim i As Long, INextSno As Double, Bal As Double, Amt As Double, dCustomerAmt As Double, dGLDebitAmt As Double, dGLCreditAmt As Double
    Dim VchrNo As String, theDesc As String, AccNo As String
    
    Dim Dated As Date
    Dated = DTVchr
    'Charge All Amount to Cash
    Amt = dAmt
    
    If Amt = 0 Then Exit Function
     
    Dim rs As New ADODB.Recordset, dTotalAdvAmt As Double, dTotalAdvAmtPK As Double
    dTotalAdvAmt = 0
    With rs
        .Open "SELECT CustomInvoice,FCustAdvancePaymentsAdjusted.Amount,ExchRate FROM FCustAdvancePayments" & _
                " INNER JOIN FCustAdvancePaymentsAdjusted ON  FCustAdvancePayments.EntryID=FCustAdvancePaymentsAdjusted.RefID WHERE CustomInvoice='" & strInvNo & "'", con, adOpenForwardOnly, adLockReadOnly
                
        Do Until .EOF
            dTotalAdvAmt = dTotalAdvAmt + Val(!Amount & "")
            dTotalAdvAmtPK = dTotalAdvAmtPK + (Val(!Amount & "") * Val(!ExchRate & ""))
            dTotalAdvAmtPK = Round(dTotalAdvAmtPK)
            .MoveNEXT
        Loop
        .Close
    End With
    'ConAcc.BeginTrans
     
    'Get The Required Values
     
    INextSno = getNextSno(Dated)
    VchrNo = getNextVchrNo(Dated, "JV")
     
    'Give The Description
    theDesc = Desc '"Invoice # : " & InvNo
     
    ''''''''''''''''Charge To  Foreign Sales Account... Charge Full Amount'''''''''''''
    
    AccNo = FSAccNo
    
    Bal = getBalance(AccNo, Dated)
    
    Bal = Bal - Amt
        
    'Credit Entry
    con.Execute "INSERT INTO Vouchers VALUES(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
    
    
    Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
     
     
     INextSno = INextSno + 1
     
    '''''''''''''''''' Charge to Customer Account... Amt Depending on Advance Posting against Invoice... '''''''''''''
      
      AccNo = CustAccNo
     
      Bal = getBalance(AccNo, Dated)
         
      If dTotalAdvAmt = 0 Then
            dCustomerAmt = Amt
      Else
            dCustomerAmt = dExchRate * (dAmtF - dTotalAdvAmt)
            dCustomerAmt = dCustomerAmt + (dTotalAdvAmtPK)
            dCustomerAmt = Round(dCustomerAmt)
      End If
      
      Bal = Bal + dCustomerAmt
    
    con.Execute "INSERT INTO Vouchers VALUES(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & theDesc & "'," & dCustomerAmt & ",0," & Bal & ",'',0)"
        
     
    Call adjustLedger(AccNo, Dated, INextSno, Val(dCustomerAmt))
    
      'Now If Foreing Exch. Gain or Loss
    If dTotalAdvAmt <> 0 Then   '
        dGLCreditAmt = 0
        dGLDebitAmt = 0
        If dCustomerAmt > Amt Then
            dGLCreditAmt = dCustomerAmt - Amt
            dGLDebitAmt = 0
        ElseIf Amt > dCustomerAmt Then
            dGLCreditAmt = 0
            dGLDebitAmt = Amt - dCustomerAmt
        End If
        
        If dGLCreditAmt > 0 Or dGLDebitAmt > 0 Then
            dGLCreditAmt = Round(dGLCreditAmt)
            dGLDebitAmt = Round(dGLDebitAmt)
            Bal = getBalance(strExchDiffAccNo, Dated)
            Bal = Bal + (dGLDebitAmt - dGLCreditAmt)
            
            INextSno = INextSno + 1
            con.Execute "INSERT INTO Vouchers VALUES(" & INextSno & ",'" & _
                Dated & "','" & VchrNo & "','" & strExchDiffAccNo & _
                "','" & theDesc & "'," & dGLDebitAmt & "," & dGLCreditAmt & "," & Bal & ",'',0)"
    
            Call adjustLedger(strExchDiffAccNo, Dated, INextSno, (dGLDebitAmt - dGLCreditAmt))
            
        End If
    End If
        
      
    
    con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & VchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
    
    AddVoucher = VchrNo
    
End Function

Private Sub DTVchr_Change()
    Call setDateRange(DTVchr)
End Sub

Private Sub Form_Load()
    

    cmbCust.ListHeight = 2500
    cmbCust.AddItem "<All Customers>", "0"
    Call cmbCust.AddVals(con, "Distinct CustCode", "ForeignCustomers", "CustCode")
    'AddToCombo CmbAccTitle, "AccTitle", "Accounts", " Where AccNo='" & FSALESACC & "' OR AccNo='" & FSAMPLESALESACC & "'", , "AccNo"
    AddToCombo CmbAccTitle, "AccTitle", "Accounts", " WHERE SubAccOf='31-001' OR AccNo='33-001-12001'", , "AccNo"
    If CmbAccTitle.ListCount > 0 Then
'        If strCompany = "IAA" Or strCompany = "Kami" Then
'            CmbAccTitle.Text = "Export Sales"
'        Else
            CmbAccTitle.ListIndex = 0
'        End If
    End If
    
    
    DTVchr = GetServerDate
    
    strExchDiffAccNo = getGeneralDataValue("PRC_ExchDiff_AccNo")   '45-23004
    
    If strExchDiffAccNo = "" Then
        MsgBox "Exch. Diff. Account not found in general data.", vbCritical
        cmdUpdate.Enabled = False
    End If
    
    DtFrom = Date
    DtTo = Date
    cmbRange.ListIndex = 5
    
    
End Sub

Private Sub RefreshLV()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim Cond As String
    Dim DTRange As String
    Dim ITM As ListItem
    
    Screen.MousePointer = vbHourglass
    
    Cond = " WHERE DT BETWEEN '" & DtFrom.Value & "' AND '" & DtTo.Value & "'"
    
    If cmbCust.ID <> "0" Then
        Cond = Cond & " AND CustCode='" & cmbCust.ID & "' "
    End If
    
    If cmbCountry.ListIndex <> 0 Then
        Cond = Cond & " and Country='" & cmbCountry & "'"
    End If
    
'    If Cond = "" Then
'        Cond = " Where CustomInvoice Not In(Select CustomInvoice From CustomInvoiceAuthorization) "
'    Else
'        Cond = " AND CustomInvoice Not In(Select CustomInvoice From CustomInvoiceAuthorization) "
'    End If
    With rs
        
        .Open "Select * From VCustomInvoiceAuth " & Cond & " Order By DT", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        'LV.Sorted = False
        Do Until .EOF
            Set ITM = LV.ListItems.add(, ![CustomInvoice] & "'", .AbsolutePosition)
            ITM.Tag = ![Curr] & ""
            ITM.ListSubItems.add(, , ![CustCode] & "").Tag = ![AccNo] & ""
            'ITM.SubItems(1) = ![CustCode] & ""
            ITM.ListSubItems.add(, , ![Country] & "").Tag = ![PackingInfo] & ""
            'ITM.SubItems(2) = ![Country] & ""
            ITM.ListSubItems.add , , ![CustomInvoice] & ""
            'ITM.SubItems(3) = ![OrderNo] & ""
            ITM.ListSubItems.add , , Format(![DT] & "", "dd-MMM-yyyy")
            'ITM.SubItems(4) = Format(![Dt] & "", "dd-MMM-yyyy")
            ITM.ListSubItems.add(, , ![Curr] & " " & Val(![TotalCustomAmt] & "")).Tag = Val(![TotalCustomAmt] & "")
            'ITM.SubItems(5) = Val(![OrderAmt] & "")
            ITM.SubItems(6) = ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    FAcc.Move (Me.Width - FAcc.Width) / 2, (Me.Height - FAcc.Height) / 2
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    If LV.SortKey + 1 = ColumnHeader.Index Then
        If LV.SortOrder = lvwAscending Then
            LV.SortOrder = lvwDescending
            ColumnHeader.Icon = "Down"
        Else
            LV.SortOrder = lvwAscending
            ColumnHeader.Icon = "Up"
        End If
    Else
        LV.ColumnHeaderIcons = Nothing
        Set LV.ColumnHeaderIcons = ImageList2
        
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        LV.Sorted = True
        ColumnHeader.Icon = "Down"
    End If
    Call RefreshSnos
    'If LV.Sorted = False Then LV.Sorted = True
End Sub

Private Sub LV_DblClick()
    
    If LV.ListItems.count = 0 Then Exit Sub
    With txtDesc
        .Move LV.ColumnHeaders(7).Left + LV.Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(7).Width
        .Text = Val(LV.SelectedItem.SubItems(6))
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
    
End Sub


Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        Item.Selected = True
        Call LV_DblClick
    End If
End Sub


Private Sub txtDesc_GotFocus()
    cmdRefresh.Default = False
    txtDesc.SelStart = 0
    txtDesc.SelLength = Len(txtDesc)
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
  
    If KeyAscii = 13 Then
    
        If Val(txtDesc) = 0 Then
            MsgBox "Invalid Exchange Rate.", vbInformation
            Exit Sub
        End If
    
        'con.Execute "Update Vouchers Set [Description]='" & txtDesc & "' Where Sno=" & Val(LV.SelectedItem.key)
        LV.SelectedItem.ListSubItems(6) = Val(txtDesc)
        LV.SelectedItem.SubItems(7) = Round(Val(txtDesc) * Val(LV.SelectedItem.ListSubItems(5).Tag))
        txtDesc.Visible = False
        'cmdRefresh.Default = False
    ElseIf KeyAscii = 27 Then
        txtDesc.Visible = False
        'cmdRefresh.Default = False
    End If

End Sub

Private Sub txtDesc_LostFocus()
    txtDesc.Visible = False
End Sub
Private Sub RefreshSnos()
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Text = i
    Next i
End Sub

Private Sub cmbRange_Click()

    DtTo = GetServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
        Case 0
            DtFrom = Date
        Case 1
            DtFrom = DateAdd("d", -15, Date)
        Case 2
            DtFrom = DateAdd("d", -30, Date)
        Case 3
            DtFrom = DateAdd("d", -60, Date)
        Case 4
            DtFrom = DateAdd("d", -90, Date)
        Case 5
            DtFrom = DateAdd("d", -180, Date)
        Case 6
            DtFrom.Enabled = True
            DtTo.Enabled = True
            DtFrom.SetFocus
            Exit Sub
    End Select
    
End Sub


