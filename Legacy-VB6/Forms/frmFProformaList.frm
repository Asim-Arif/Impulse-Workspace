VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmFProformaList 
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   10530
   ScaleMode       =   0  'User
   ScaleWidth      =   10641.08
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10560
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Proforma List"
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
         Index           =   8
         Left            =   75
         TabIndex        =   9
         Top             =   165
         Width           =   10605
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Proforma List"
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
         Height          =   465
         Index           =   9
         Left            =   90
         TabIndex        =   10
         Top             =   180
         Width           =   10590
      End
   End
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   9765
      Left            =   15
      TabIndex        =   3
      Top             =   600
      Width           =   11850
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   60
         TabIndex        =   15
         Top             =   9225
         Width           =   5475
         Begin VB.ComboBox cmbfield 
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
            ItemData        =   "frmFProformaList.frx":0000
            Left            =   60
            List            =   "frmFProformaList.frx":0013
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   135
            Width           =   1875
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4395
            TabIndex        =   17
            Top             =   135
            Width           =   1020
         End
         Begin VB.TextBox txtFind 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1965
            TabIndex        =   16
            Top             =   135
            Width           =   2430
         End
      End
      Begin ComboList.Usercontrol1 cmbCompany 
         Height          =   285
         Left            =   60
         TabIndex        =   14
         Top             =   345
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
               Picture         =   "frmFProformaList.frx":0057
               Key             =   "Down"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmFProformaList.frx":0213
               Key             =   "Up"
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2475
         TabIndex        =   11
         Top             =   1605
         Visible         =   0   'False
         Width           =   3615
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
         ItemData        =   "frmFProformaList.frx":03CF
         Left            =   6060
         List            =   "frmFProformaList.frx":03E5
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   345
         Width           =   1710
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   9255
         TabIndex        =   1
         Top             =   345
         Width           =   1500
         _ExtentX        =   2646
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
         Format          =   156237827
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   7755
         TabIndex        =   0
         Top             =   345
         Width           =   1500
         _ExtentX        =   2646
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
         Format          =   156237827
         CurrentDate     =   37055
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   2370
         TabIndex        =   6
         Top             =   345
         Width           =   1965
         _ExtentX        =   3466
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
      Begin Crystal.CrystalReport Cr1 
         Left            =   270
         Top             =   2040
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
      Begin MSComctlLib.ListView LV 
         Height          =   8445
         Left            =   60
         TabIndex        =   12
         Top             =   765
         Width           =   11715
         _ExtentX        =   20664
         _ExtentY        =   14896
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Company"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Customer"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Country"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Proforma #"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Proforma DT"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Proforma Amt"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Order(s)"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Internal"
            Object.Width           =   1940
         EndProperty
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   375
         Left            =   10425
         TabIndex        =   13
         Top             =   9285
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
         Left            =   4350
         TabIndex        =   7
         Top             =   345
         Width           =   1695
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2990;503"
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
         Left            =   10755
         TabIndex        =   2
         Top             =   255
         Width           =   1020
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "1799;635"
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
         Caption         =   $"frmFProformaList.frx":043C
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
         Left            =   60
         TabIndex        =   4
         Top             =   120
         Width           =   10665
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopMenu"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add New Proforma"
      End
      Begin VB.Menu dd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit This Proforma"
      End
      Begin VB.Menu mnuDashAdvance 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeleteProforma 
         Caption         =   "Delete Proforma"
      End
      Begin VB.Menu mnuAdvancePaymentDetails 
         Caption         =   "Advance Payment Details"
      End
      Begin VB.Menu d1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGenExpDocs 
         Caption         =   "&Generate Export Documents"
      End
      Begin VB.Menu d5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print Proforma"
      End
      Begin VB.Menu mnuPrintUSAProforma 
         Caption         =   "Print USA Proforma"
      End
      Begin VB.Menu mnuPrintProformaCompanySelection 
         Caption         =   "Print Proforma (Company Selection)"
      End
      Begin VB.Menu mnuPrintProformaDiscount 
         Caption         =   "Print Proforma (Discount)"
      End
      Begin VB.Menu mnuPrintProformaWD 
         Caption         =   "Print Proforma Without Details"
      End
      Begin VB.Menu mnuDashPrintList 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintList 
         Caption         =   "Print List"
      End
   End
End
Attribute VB_Name = "frmFProformaList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
    If cmbCountry.ListCount > 0 Then cmbCountry.ListIndex = 0
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbRange_Click()

    DtTo = ServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
        
        Case 0
            DtFrom = DateAdd("d", -15, ServerDate)
        Case 1
            DtFrom = DateAdd("d", -30, ServerDate)
        Case 2
            DtFrom = DateAdd("d", -60, ServerDate)
        Case 3
            DtFrom = DateAdd("d", -90, ServerDate)
        Case 4
            DtFrom = DateAdd("d", -180, ServerDate)
        Case 5
            DtFrom.Enabled = True
            DtTo.Enabled = True
            DtFrom.SetFocus
            Exit Sub
    End Select
    
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


Private Sub Form_Load()
    
    mnuPop.Visible = False
    cmbCust.ListHeight = 2500
    cmbCust.AddItem "<All Customers>", "0"
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where custCode<>'Hico'")
    cmbCust.ID = "0"
    cmbRange.ListIndex = 4
        
    '\/\/\/\/\/\/\/\
    cmbCompany.ListHeight = 4000
    cmbCompany.AddItem "< All Companies>", "0"
    cmbCompany.AddVals con, "CompanyName", "Companies", "EntryID", " Order By CompanyName"
    cmbCompany.ID = "0"
    
    '\/\/\/\/\/\/\/\
    cmbfield.ListIndex = 3
    
    If strCompany = "IAA" Or strCompany = "Kami" Then
        mnuPrintProformaCompanySelection.Visible = True
    Else
        mnuPrintProformaCompanySelection.Visible = False
    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Dim TheHeight As Long
    Frame1.Left = (Me.ScaleWidth - Frame1.Width) / 2
    
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    
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
    
    If cmbCompany.ID <> "0" Then
        Cond = Cond & " AND CompanyRefID=" & cmbCompany.ID
    End If
    
    If cmbCountry.ListIndex <> 0 Then
        Cond = Cond & " AND Country='" & cmbCountry & "'"
    End If
    
    
    Dim iRecNo As Integer
    Dim lForeColor As Long
    With rs
        
        .Open "SELECT * FROM VFProformaList " & Cond & " Order By DT DESC,PInvoice", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        iRecNo = .RecordCount
        'LV.Sorted = False
        Do Until .EOF
            Set ITM = LV.ListItems.add(, ![PInvoice] & "'", iRecNo - .AbsolutePosition + 1)
            ITM.Tag = ![Curr] & ""
            ITM.ListSubItems.add , , ![CompanyName] & ""
            ITM.ListSubItems.add , , ![CustCode] & ""
            'ITM.SubItems(1) = ![CustCode] & ""
            ITM.ListSubItems.add , , ![Country] & ""
            'ITM.SubItems(2) = ![Country] & ""
            ITM.ListSubItems.add , , ![PInvoice] & ""
            'ITM.SubItems(3) = ![OrderNo] & ""
            ITM.ListSubItems.add , , Format(![DT] & "", "dd-MMM-yyyy")
            'ITM.SubItems(4) = Format(![Dt] & "", "dd-MMM-yyyy")
            ITM.ListSubItems.add(, , ![Curr] & " " & Val(![ProformaAmt] & "")).Tag = Val(![ProformaAmt] & "")
            'ITM.SubItems(5) = Val(![ProformaAmt] & "")
            ITM.SubItems(6) = Val(![ProformaAmt] & "") + Val(!ChargesAmt & "") - Val(!DiscountAmount & "") + Val(!Insurance_Freight_Amount & "") + Val(!Insurance_Freight_Amount_II & "")
            ITM.SubItems(7) = !Orders & ""
            ITM.SubItems(8) = !InternalOrders & ""
            
            If Val(!AdvanceAmount & "") > 0 Then
                lForeColor = vbBlue
            ElseIf Val(!ChargesAmt & "") > 0 Then
                lForeColor = vbRed
            Else
                lForeColor = LV.ForeColor
            End If
            ITM.ForeColor = lForeColor
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ForeColor = lForeColor
            Next
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

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If LV.SortKey + 1 = ColumnHeader.Index Then
        If LV.SortOrder = lvwAscending Then
            LV.SortOrder = lvwDescending
            'ColumnHeader.Icon = "Down"
        Else
            LV.SortOrder = lvwAscending
            'ColumnHeader.Icon = "Up"
        End If
    Else
        LV.ColumnHeaderIcons = Nothing
        Set LV.ColumnHeaderIcons = ImageList2
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        'ColumnHeader.Icon = "Down"
    End If
    Call RefreshSnos
    'If LV.Sorted = False Then LV.Sorted = True
End Sub

Private Sub LV_DblClick()
    'If LV.ListItems.Count = 0 Or OutOfFinancialYear Then Exit Sub
    'If InStr(LV.SelectedItem, "-") > 0 Then Call mnuEditDesc_Click
    If LV.ListItems.count = 0 Then Exit Sub
        
    If strCompany = "IAA" Or strCompany = "Kami" Then
        If getDBPassword("EditExportDocs") = False Then Exit Sub
    End If
    
    Load frmFNewProforma
    
    Call frmFNewProforma.EditProforma(Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1))
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button <> 2 Then Exit Sub
    If LV.ListItems.count = 0 Then
        Call ShowMenus(False)
        
    Else
        Call ShowMenus(True)
        
    End If
    
    'If LV.ListItems.Count = 0 Or Button <> 2 Or OutOfFinancialYear Then Exit Sub

    'if a Voucher Entry Show Delete Menu
    'If InStr(LV.SelectedItem, "-") > 0 Then Me.PopupMenu mnuPop
    
End Sub
Private Sub ShowMenus(vis As Boolean)
    d1.Visible = vis



    d5.Visible = vis
    'mnuAdd.Visible = Vis
    mnuEdit.Visible = vis




    mnuPrint.Visible = vis
    
    
    Me.PopUpMenu mnuPop
End Sub

Private Sub mnuAdd_Click()
    Load frmFNewProforma
    frmFNewProforma.Show
    frmFNewProforma.AddNewProforma
End Sub

Private Sub mnuAdvancePaymentDetails_Click()
    
    If LV.ListItems.count = 0 Then Exit Sub
    
    Dim f As New frmGetTextAndNumericValue
    Dim strValue As String, dValue As Double
    
    Load f
    If f.getTextANDNumericValue(strValue, dValue, False, "Description", "Amount", "Advance Payments") = False Then Exit Sub
         
    con.Execute "UPDATE FPInvoice SET AdvanceDescription='" & strValue & "',AdvanceAmount=" & dValue & " WHERE PInvoice='" & LV.SelectedItem.SubItems(4) & "'"
    
End Sub

Private Sub mnuDeleteProforma_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(*)", "FProformaOrders INNER JOIN CustomInvoiceItems ON FProformaOrders.EntryID=CustomInvoiceItems.RefID", " WHERE FProformaOrders.PInvoice='" & LV.SelectedItem.SubItems(4) & "'")
    If lCount > 0 Then
        MsgBox "One or more invoices are using the selected profroma, Can't Delete.", vbInformation
        Exit Sub
    End If
     
    If getDBPassword("DeleteExportEntries") = False Then Exit Sub
    
    Call StartTrans(con)
    con.Execute "DELETE FROM FPInvoice WHERE PInvoice='" & LV.SelectedItem.SubItems(4) & "'"
    con.CommitTrans
End Sub

Private Sub mnuedit_Click()
    Call LV_DblClick
End Sub

Private Sub mnuGenExpDocs_Click()
    'frmGenExpDocs.Show 1
End Sub

Private Sub mnuPrint_Click()
    Call PrintProforma
End Sub

Private Sub PrintProforma(Optional bWithoutDetails As Boolean = False, Optional dDiscount As Double = 0, Optional lCompanyRefID As Long = 0, Optional bUSAProforma As Boolean = False)
    
    On Error GoTo err
    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    Dim rpt_Temp_App As New CRAXDDRT.Application
    If bWithoutDetails Then
        Set rpt = rpt_Temp_App.OpenReport(RptPath & "\rptFProformaWD.rpt") '
    ElseIf bUSAProforma Then
        Set rpt = rpt_Temp_App.OpenReport(RptPath & "\Proforma_USA.rpt") '
    Else
        Set rpt = rpt_Temp_App.OpenReport(RptPath & "\rptFProforma.rpt") '
        'If dDiscount > 0 Then
        rpt.FormulaFields.GetItemByName("Discount").Text = dDiscount
        
        If strCompany = "IAA" Or strCompany = "Kami" Then
            If lCompanyRefID = 0 Then
                lCompanyRefID = GetSingleLongValue("CompanyRefID", "FPInvoice", " WHERE PInvoice='" & LV.SelectedItem.SubItems(4) & "'")
            End If
            Dim rptSub As CRAXDDRT.Report
        
            Set rptSub = rpt.OpenSubreport("CompanyHeader")
            rptSub.RecordSelectionFormula = "{Companies.EntryID}=" & lCompanyRefID
        End If
        
        'End If
    End If
    
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VrptProforma.PInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt
    
    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description

End Sub


Private Sub mnuPrintSummary_Click()
    On Error GoTo err
    Me.MousePointer = vbHourglass
    
    Call StartTrans(con)
    con.Execute "Delete From PrintFOrderList"
    
    For i = 1 To LV.ListItems.count
        con.Execute "Insert Into PrintFOrderList(Sno,Customer,Country," & _
         "OrderNo,OrderDT,OrderAmt,Status,Curr) Values(" & LV.ListItems(i) & _
         ",'" & LV.ListItems(i).SubItems(1) & "','" & LV.ListItems(i).SubItems(2) & _
         "','" & LV.ListItems(i).SubItems(3) & "','" & _
         LV.ListItems(i).SubItems(4) & "'," & Val(LV.ListItems(i).ListSubItems(5).Tag & "") & _
         ",'" & LV.ListItems(i).SubItems(6) & "','" & LV.ListItems(i).Tag & "')"
    Next i
    
    con.CommitTrans
    
    Load frmPrevRpt
    
    Dim strFromTo As String
    Dim iNumber As Integer
    If cmbRange.ListIndex = 0 Then
        iNumber = -15
    ElseIf cmbRange.ListIndex = 1 Then
        iNumber = -30
    ElseIf cmbRange.ListIndex = 2 Then
        iNumber = -60
    ElseIf cmbRange.ListIndex = 3 Then
        iNumber = -90
    End If
    
    If cmbRange.ListIndex = 4 Then
        strFromTo = Format(DtFrom, "dd-MMM-yyyy") & " To  " & Format(DtTo, "dd-MMM-yyyy")
    Else
        strFromTo = Format(DateAdd("d", iNumber, ServerDate), "dd-MMM-yyyy") & " To  " & Format(ServerDate, "dd-MMM-yyyy")
    End If
    
    Dim rptFProformaList As CRAXDDRT.Report
    Set rptFProformaList = rptApp.OpenReport(RptPath & "\rptFProformaList.rpt") '
    rptFProformaList.FormulaFields(3).Text = "'" & cmbCust.Text & "'"
    rptFProformaList.FormulaFields(4).Text = "'" & strFromTo & "'"
    
    frmPrevRpt.ShowReport "", rptFProformaList
    
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub mnuPrintList_Click()
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ProformaList.rpt")
    Dim strSelection As String
    strSelection = "{VFProformaList.DT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#"
    If cmbCust.ID <> "0" Then
        strSelection = strSelection & " AND {VFProformaList.CustCode}='" & cmbCust.ID & "'"
    End If
    
    If cmbCountry.ListIndex > 0 Then
        strSelection = strSelection & " AND {VFProformaList.Country}='" & cmbCountry.Text & "'"
    End If
    
    rpt.FormulaFields.GetItemByName("Customers").Text = "'" & cmbCust.Text & " " & cmbCountry.Text & "'"
    rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & Format(DtFrom, "dd-MMM-yyyy") & " to " & Format(DtTo.Value, "dd-MMM-yyyy") & "'"
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt

End Sub

Private Sub mnuPrintProformaCompanySelection_Click()

    Dim strCompanyRefID As String
    Dim f As New frmGetSingleComboValue
    Load f
    f.cmbCombo.AddVals con, "CompanyName", "Companies", "EntryID"
    If f.getComboValue(strCompanyRefID, "Company", "Company") = False Then Exit Sub
    Dim lCompanyRefID As Long
    lCompanyRefID = Val(strCompanyRefID)
    Call PrintProforma(False, 0, lCompanyRefID)
    
End Sub

Private Sub mnuPrintProformaDiscount_Click()

    Dim strDiscount As String
    Dim f As New frmGetSingleTextValue
    If f.getTextValue(strDiscount, , "Discount", "Discount") = False Then Exit Sub
         
    Dim dDiscount As Double
    dDiscount = Val(strDiscount)
    
    Call PrintProforma(False, dDiscount)
    
End Sub

Private Sub mnuPrintProformaWD_Click()
    Call PrintProforma(True)
End Sub

Private Sub mnuPrintUSAProforma_Click()
    Call PrintProforma(False, 0, 0, True)
End Sub

Private Sub txtDesc_GotFocus()
    cmdRefresh.Default = False
    txtDesc.SelStart = 0
    txtDesc.SelLength = Len(txtDesc)
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
  
If KeyAscii = 13 Then
    
    If Trim(txtDesc = "") Then
        MsgBox "Invalid Description", vbInformation
        Exit Sub
    End If
    
   con.Execute "Update Vouchers Set [Desc]='" & txtDesc & "' Where Sno=" & Val(LV.SelectedItem.key)
   LV.SelectedItem.ListSubItems(3) = txtDesc & txtDesc.Tag
   txtDesc.Visible = False
   cmdRefresh.Default = False
ElseIf KeyAscii = 27 Then
   txtDesc.Visible = False
   cmdRefresh.Default = False
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

Private Sub txtFind_Change()
    cmdFind.Enabled = Len(Trim(txtFind)) > 0
End Sub

Private Sub txtFind_GotFocus()
   txtFind.SelStart = 0
   txtFind.SelLength = Len(txtFind)
   cmdFind.Default = True
End Sub

Private Sub txtFind_LostFocus()
   cmdFind.Default = False
End Sub

Private Sub cmdFind_Click()
    
    Dim StartIdx As Integer
    
    If LV.SelectedItem Is Nothing Then
       StartIdx = 1
    Else
       If LV.SelectedItem.Index = LV.ListItems.count Then
          StartIdx = 1
       Else
          StartIdx = LV.SelectedItem.Index
       End If
    End If

    Dim fldValue As String

    For i = StartIdx To LV.ListItems.count
        
        If cmbfield.ListIndex > 2 Then
            fldValue = LV.ListItems(i).SubItems(cmbfield.ListIndex + 4)
        Else
            fldValue = LV.ListItems(i).SubItems(cmbfield.ListIndex + 2)
        End If
        If LCase(fldValue) Like "*" & LCase(Trim(txtFind)) & "*" Then
            LV.ListItems(i).Selected = True
            LV.ListItems(i).EnsureVisible
            LV.SetFocus
            If i <> StartIdx Then Exit For
        End If
       
    Next
    Dim rpt As New CRAXDDRT.Report
    
    If i = LV.ListItems.count + 1 Then
        MsgBox "Record Not Found.", vbInformation
        Set LV.SelectedItem = Nothing
    End If

End Sub

