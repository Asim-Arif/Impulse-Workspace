VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmCustomInvoiceList 
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13875
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   8925
   ScaleMode       =   0  'User
   ScaleWidth      =   12428.03
   WindowState     =   2  'Maximized
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   8970
      Left            =   60
      TabIndex        =   3
      Top             =   -45
      Width           =   13770
      Begin VB.Frame Frame9 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   13755
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Custom Invoice List"
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
            Index           =   33
            Left            =   75
            TabIndex        =   14
            Top             =   135
            Width           =   13620
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Custom Invoice List"
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
            Index           =   34
            Left            =   90
            TabIndex        =   15
            Top             =   150
            Width           =   13605
         End
      End
      Begin MSComCtl2.DTPicker DTEdit 
         Height          =   345
         Left            =   2520
         TabIndex        =   12
         Top             =   3705
         Visible         =   0   'False
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
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
         Format          =   157351939
         CurrentDate     =   40684
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   300
         Left            =   1110
         TabIndex        =   11
         Top             =   8535
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2475
         TabIndex        =   8
         Top             =   2190
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
         ItemData        =   "frmCustomInvoiceList.frx":0000
         Left            =   4440
         List            =   "frmCustomInvoiceList.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   930
         Width           =   2010
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   8190
         TabIndex        =   1
         Top             =   930
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
         Format          =   157351939
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   6465
         TabIndex        =   0
         Top             =   930
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
         Format          =   157351939
         CurrentDate     =   37055
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   60
         TabIndex        =   6
         Top             =   930
         Width           =   2460
         _ExtentX        =   4339
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
      Begin MSComctlLib.ListView LV 
         Height          =   7110
         Left            =   60
         TabIndex        =   9
         Top             =   1350
         Width           =   13605
         _ExtentX        =   23998
         _ExtentY        =   12541
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Country"
            Object.Width           =   1722
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Invoice #"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Invoice DT"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Invoice Amt"
            Object.Width           =   2565
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Ship. Mode"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "GP Date"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Maturity"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Orders"
            Object.Width           =   3889
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "D.P Nos"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   375
         Left            =   12330
         TabIndex        =   10
         Top             =   8520
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
         TabIndex        =   7
         Top             =   930
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
         Left            =   9810
         TabIndex        =   2
         Top             =   900
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
         Caption         =   $"frmCustomInvoiceList.frx":0078
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
         Left            =   45
         TabIndex        =   4
         Top             =   705
         Width           =   11400
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
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
            Picture         =   "frmCustomInvoiceList.frx":0112
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomInvoiceList.frx":02CE
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopMenu"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add New Courier Invoice"
      End
      Begin VB.Menu dd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit This Invoice"
      End
      Begin VB.Menu mnuDeleteInvoice 
         Caption         =   "Delete Invoice"
      End
      Begin VB.Menu d1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUpdateCustomPrice 
         Caption         =   "Update Custom Prices"
      End
      Begin VB.Menu mnuUpdateRebateRate 
         Caption         =   "Update Rebate Rates"
      End
      Begin VB.Menu mnuGenerateBatches 
         Caption         =   "Generate Batches"
      End
      Begin VB.Menu mnuImpMaterial 
         Caption         =   "Auto Calc &Imported Material"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPacking 
         Caption         =   "Add Packing Info."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDelPackingInfo 
         Caption         =   "Delete Packing Info."
         Visible         =   0   'False
      End
      Begin VB.Menu dashship 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAddShippInfo 
         Caption         =   "Add Shipping Bill Info."
         Visible         =   0   'False
      End
      Begin VB.Menu dashshipinfo 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuShippingInstructions 
         Caption         =   "Shipping Instructions"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEForm 
         Caption         =   "'E' Form"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGSP 
         Caption         =   "G S P"
         Visible         =   0   'False
      End
      Begin VB.Menu dashGSP 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintPList 
         Caption         =   "Print Packing List"
         Visible         =   0   'False
         Begin VB.Menu mnuCustomPList 
            Caption         =   "Custom Packing List"
         End
         Begin VB.Menu mnuCustomPListOther 
            Caption         =   "Custom Packing List (Other Format)"
         End
         Begin VB.Menu mnuCommercialPList 
            Caption         =   "Commercial Packing List"
         End
         Begin VB.Menu mnuCommercialPListOther 
            Caption         =   "Commercial Packing List (Other Format)"
         End
         Begin VB.Menu mnuDashPackingLabels 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPrintPackingLabels 
            Caption         =   "Print Packing Labels"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu d6 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintActualInvoice 
         Caption         =   "Print Actual Invoice"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print Invoice"
      End
      Begin VB.Menu mnuPrintUSAInvoice 
         Caption         =   "Print USA Invoice"
      End
      Begin VB.Menu mnuPrintInvoiceasHandTools 
         Caption         =   "Print Invoice as Hand Tools"
      End
      Begin VB.Menu mnuPrintInvoiceWOCustomerAddress 
         Caption         =   "Print Invoice (W/O Customer Address)"
      End
      Begin VB.Menu mnuPrintInvoiceDiscount 
         Caption         =   "Print Invoice (Discount)"
      End
      Begin VB.Menu mnuPrintPaymentInvoice 
         Caption         =   "Print Payment Invoice"
      End
      Begin VB.Menu mnuPrintUSAPaymentInvoice 
         Caption         =   "Print USA Payment Invoice"
      End
      Begin VB.Menu mnuPrintPaymentDocs 
         Caption         =   "Print Payment Docs"
      End
      Begin VB.Menu mnuPrintSpecialCustomInvoice 
         Caption         =   "Print Special Custom Invoice"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintWithBatchNo 
         Caption         =   "Print Invoice (With Batch No.)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintPackingList 
         Caption         =   "Print Packing List"
      End
      Begin VB.Menu mnuPrintGSPInvoice 
         Caption         =   "Print GSP Invoice"
      End
      Begin VB.Menu mnuCustomPListCombined 
         Caption         =   "Packing List Combined"
      End
      Begin VB.Menu mnuCustomPListCombinedCustom 
         Caption         =   "Packing List Combined (Custom)"
      End
      Begin VB.Menu mnuPrintBillofExchange 
         Caption         =   "Print Bill of Exchange"
      End
      Begin VB.Menu mnuPrintCoverLetter 
         Caption         =   "Print Covering Letter"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintInvWDetail 
         Caption         =   "Print Invoice With Detail of Imp. Material"
         Visible         =   0   'False
      End
      Begin VB.Menu DashComInvPrint 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuComDocs 
         Caption         =   "Commercial Documents"
         Visible         =   0   'False
         Begin VB.Menu mnuPrintComInv 
            Caption         =   "Print Commercial Invoice"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPrint5 
            Caption         =   "Print With 5 % &Compliance"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuPrint10 
            Caption         =   "Print with 10 % Premium"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuComDocsDash1 
            Caption         =   "-"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuBillOfExchange 
            Caption         =   "Print Bill of Exchange (I)"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuBillOfExchange2 
            Caption         =   "Print Bill of Exchange (II)"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuComDocsDash2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCoveringLetter 
            Caption         =   "Covering Letter"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu PrintDecathDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintDecDocs 
         Caption         =   "Print Decathlon Docs..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuInvPListComparison 
         Caption         =   "Print Invoice / Packing List Comparison"
      End
      Begin VB.Menu mnuPrintLabLetter 
         Caption         =   "Print Lab Test Letter"
      End
      Begin VB.Menu mnuPrintBankSlipLabTest 
         Caption         =   "Print Bank Slip for Lab Test"
      End
      Begin VB.Menu mnuPrintUnderTaking 
         Caption         =   "Print Undertaking"
      End
      Begin VB.Menu mnuPrintLetterToBank 
         Caption         =   "Letter to Bank"
      End
      Begin VB.Menu mnuPrintLetterToBankSimple 
         Caption         =   "Letter to Bank (Simple)"
      End
      Begin VB.Menu mnuPrintBankSalesContract 
         Caption         =   "Bank Sales Contract"
      End
      Begin VB.Menu mnuPrintReasonForDelay 
         Caption         =   "Reason For Delay"
      End
      Begin VB.Menu mnuPrintReasonForDelaySBP 
         Caption         =   "Reason For Delay (SBP)"
      End
      Begin VB.Menu mnuPrintLabReport 
         Caption         =   "Lab Report"
      End
      Begin VB.Menu mnuPrintBankUnderTaking1 
         Caption         =   "Bank Undertaking 1"
      End
      Begin VB.Menu mnuPrintBankUnderTaking2 
         Caption         =   "Bank Undertaking 2"
      End
      Begin VB.Menu mnuPrintBankUnderTaking3 
         Caption         =   "Bank Undertaking 3"
      End
      Begin VB.Menu mnuDashPrintList 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintList 
         Caption         =   "Print List"
      End
      Begin VB.Menu mnuDashCustomerFeedback 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCustomerFeedback 
         Caption         =   "Customer Feedback"
      End
      Begin VB.Menu mnuPrintInvoiceExport 
         Caption         =   "Print Invoice (Export)"
      End
      Begin VB.Menu mnuDeclarationOfConformity 
         Caption         =   "Print Declaration of Conformity"
      End
      Begin VB.Menu mnuPrintCertificateOfCompliance 
         Caption         =   "Print Certificate of Compliance"
      End
   End
End
Attribute VB_Name = "frmCustomInvoiceList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCourierInvoice As Boolean
Dim bShipping As Boolean
Dim bCustomInvoice As Boolean
Dim iColNo As Integer
Private Const iMaturityColIndex As Integer = 9
Private Const iGPColIndex As Integer = 8
Dim bSampleInvoice As Boolean
Dim iInvoiceType As Integer

Public Sub ShowMe(Optional p_bCustomInvoice As Boolean = True, Optional m_bSampleInvoice As Boolean = False, Optional p_iInvoiceType As Integer = 0)
    
    'p_iInvoiceType=0:Custom Invoice
    'p_iInvoiceType=1:Commercial Invoice
    'p_iInvoiceType=2:Bank Invoice
    
    iInvoiceType = p_iInvoiceType
    
    bCustomInvoice = p_bCustomInvoice
    bSampleInvoice = m_bSampleInvoice
    
    If bSampleInvoice Then
        Label1(33).Caption = "Sample Invoice List"
        Label1(34).Caption = "Sample Invoice List"
    Else
        If iInvoiceType = 0 Then
            Label1(33).Caption = "Custom Invoice List"
            Label1(34).Caption = "Custom Invoice List"
        ElseIf iInvoiceType = 1 Then
            Label1(33).Caption = "Commercial Invoice List"
            Label1(34).Caption = "Commercial Invoice List"
        ElseIf iInvoiceType = 2 Then
            Label1(33).Caption = "Bank Invoice List"
            Label1(34).Caption = "Bank Invoice List"
        End If
    End If
    
    
    mnuPrintActualInvoice.Visible = False
    mnuPrintLetterToBank.Visible = False
    mnuPrintLetterToBankSimple.Visible = False
    mnuPrintBankSalesContract.Visible = False
    mnuPrintReasonForDelay.Visible = False
    mnuPrintReasonForDelaySBP.Visible = False
    mnuPrintLabReport.Visible = False
    mnuPrintBankUnderTaking1.Visible = False
    mnuPrintBankUnderTaking2.Visible = False
    mnuPrintBankUnderTaking3.Visible = False
        
    mnuPrintUSAPaymentInvoice.Visible = False
    mnuPrintCertificateOfCompliance.Visible = False
    mnuPrintInvoiceasHandTools.Visible = False
    mnuPrintGSPInvoice.Visible = False
    
    If iInvoiceType = 0 Then    'Custom Invoice
    
        If strCompany = "Dr-Frgz" Then
            mnuPrintSpecialCustomInvoice.Visible = True
        End If
        mnuDeclarationOfConformity.Visible = False
        mnuPrintPaymentInvoice.Visible = False
        mnuPrintPaymentDocs.Visible = False
        mnuPrintActualInvoice.Visible = True
        
    ElseIf iInvoiceType = 1 Then    'Commercial Invoice
    
        If strCompany = "Dr-Frgz" Then
            mnuPrint.Caption = "Print Invoice (With Order)"
            mnuPrintWithBatchNo.Visible = True
        End If
        mnuDeclarationOfConformity.Visible = True
        mnuPrintPaymentInvoice.Visible = True
        mnuPrintPaymentDocs.Visible = True
        mnuPrintCertificateOfCompliance.Visible = True
        mnuPrintInvoiceasHandTools.Visible = True
        mnuPrintUSAPaymentInvoice.Visible = True
        mnuPrintGSPInvoice.Visible = True
        
    ElseIf iInvoiceType = 2 Then    'Bank Invoice
        'Hide for Bank Invoice
        mnuAdd.Visible = False
        dd.Visible = False
        mnuEdit.Visible = False
        mnuDeleteInvoice.Visible = False
        d1.Visible = False
        mnuUpdateCustomPrice.Visible = False
        mnuUpdateRebateRate.Visible = False
        mnuGenerateBatches.Visible = False
        mnuPrintUSAInvoice.Visible = False
        mnuPrintInvoiceWOCustomerAddress.Visible = False
        mnuPrintInvoiceDiscount.Visible = False
        mnuPrintPaymentInvoice.Visible = False
        mnuPrintPaymentDocs.Visible = False
        mnuCustomPListCombined.Visible = False
        mnuCustomPListCombinedCustom.Visible = False
        mnuPrintBillofExchange.Visible = False
        mnuInvPListComparison.Visible = False
        mnuPrintLabLetter.Visible = False
        mnuPrintBankSlipLabTest.Visible = False
        mnuPrintUnderTaking.Visible = False
        mnuCustomerFeedback.Visible = False
        mnuPrintInvoiceExport.Visible = False
        mnuDeclarationOfConformity.Visible = False
        'Show for Bank Invoice
        mnuPrintLetterToBank.Visible = True
        mnuPrintLetterToBankSimple.Visible = True
        mnuPrintBankSalesContract.Visible = True
        mnuPrintReasonForDelay.Visible = True
        mnuPrintReasonForDelaySBP.Visible = True
        mnuPrintLabReport.Visible = True
        mnuPrintBankUnderTaking1.Visible = True
        mnuPrintBankUnderTaking2.Visible = True
        mnuPrintBankUnderTaking3.Visible = True
    End If
    
    Me.Show
    
End Sub

Public Sub ShowForShipping()
    bShipping = True
    Me.Show
End Sub

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

Private Sub cmdClose_Click()
   Unload Me
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub


Private Sub DTEdit_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        If iColNo = iMaturityColIndex Then
            con.Execute "UPDATE CustomInvoice SET DTMaturity='" & DTEdit.Value & "' WHERE CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'"
        ElseIf iColNo = iGPColIndex Then
            con.Execute "UPDATE CustomInvoice SET GatePassDT='" & DTEdit.Value & "' WHERE CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'"
        End If
        LV.SelectedItem.SubItems(iColNo - 1) = Format(DTEdit, "dd-MMM-yyyy")
        DTEdit.Visible = False
    ElseIf KeyCode = 27 Then
        DTEdit.Visible = False
    End If
    
End Sub

Private Sub DTEdit_LostFocus()
    DTEdit.Visible = False
End Sub

Private Sub DtFrom_Change()
    'Call setDateRange(DtFrom)
End Sub

Private Sub DtTo_Change()
    'Call setDateRange(DtTo)
End Sub


Private Sub Form_Load()
    
    bShipping = False
    bCourierInvoice = False
    mnuPop.Visible = False
    cmbCust.ListHeight = 2500
    cmbCust.AddItem "<All Customers>", "0"
    Call cmbCust.AddVals(con, "DISTINCT CustCode", "ForeignCustomers", "CustCode")
    cmbCust.ID = "0"
    DtFrom = Date
    DtTo = Date
    cmbRange.ListIndex = 5
    
    If strCompany = "IAA" Or strCompany = "Kami" Then
        mnuPrintInvoiceDiscount.Visible = True
        mnuPrintInvoiceWOCustomerAddress.Visible = True
    Else
        mnuPrintInvoiceDiscount.Visible = False
        mnuPrintInvoiceWOCustomerAddress.Visible = False
    End If
    bSampleInvoice = False
    
End Sub

Private Sub RefreshLV()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim Cond As String
    Dim DTRange As String
    Dim ITM As ListItem
    
    Screen.MousePointer = vbHourglass
    
'    If bCourierInvoice Then
'        Cond = " Where Left(CustomInvoice,4)='" & strCourierInvoicePrefix & "'"
'    Else
'        Cond = " Where Left(CustomInvoice,4)='VTL '"
'    End If
    
    If cmbCust.ID <> "0" Then
        Cond = " Where CustCode='" & cmbCust.ID & "' "
    End If
    
    If cmbCountry.ListIndex <> 0 Then
        If Cond = "" Then
            Cond = " Where Country='" & cmbCountry & "'"
        Else
            Cond = Cond & " and Country='" & cmbCountry & "'"
        End If
    End If
    
    If cmbRange.ListIndex = 0 Then
        DTRange = " DT Between '" & DateAdd("d", 0, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 1 Then
        DTRange = " DT Between '" & DateAdd("d", -15, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 2 Then
        DTRange = " DT Between '" & DateAdd("d", -30, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 3 Then
        DTRange = " DT Between '" & DateAdd("d", -60, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 4 Then
        DTRange = " DT Between '" & DateAdd("d", -90, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 5 Then
        DTRange = " DT Between '" & DateAdd("d", -180, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 6 Then
        DTRange = " DT Between '" & DtFrom & "' and '" & DtTo & "'"
    End If
    
    If Cond = "" Then
        Cond = " Where "
    Else
        Cond = Cond & " and "
    End If
    Cond = Cond & DTRange
    
    Cond = Cond & " AND SampleInvoice=" & Abs(bSampleInvoice)
    'Cond = Cond & " AND Left(CustomInvoice,4)='VTL '"
    With rs
        
        .Open "Select * From VCustomInvoiceList " & Cond & " Order By DT", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        'LV.Sorted = False
        Do Until .EOF
            Set ITM = LV.ListItems.add(, ![CustomInvoice] & "'", .AbsolutePosition)
            ITM.Tag = ![Curr] & ""
            ITM.ListSubItems.add(, , ![CustCode] & "").Tag = "" '![ImpMaterial] & ""
            'ITM.SubItems(1) = ![CustCode] & ""
            ITM.ListSubItems.add(, , ![Country] & "").Tag = ![PackingInfo] & ""
            'ITM.SubItems(2) = ![Country] & ""
            ITM.ListSubItems.add , , ![CustomInvoice] & ""
            'ITM.SubItems(3) = ![OrderNo] & ""
            ITM.ListSubItems.add , , Format(![DT] & "", "dd-MMM-yyyy")
            'ITM.SubItems(4) = Format(![Dt] & "", "dd-MMM-yyyy")
'            If bShipping = False Then
'                ITM.ListSubItems.Add(, , ![Curr] & " " & Val(![TotalAmt] & "")).Tag = Val(![TotalAmt] & "")
'            Else
'                ITM.ListSubItems.Add(, , "-").Tag = "-"
'            End If
            'ITM.SubItems(5) = Val(![OrderAmt] & "")
            If bCustomInvoice Then
                ITM.ListSubItems.add(, , ![Curr] & " " & Format(Val(![TotalCustomAmt] & ""), "#.###")).Tag = Val(![TotalCustomAmt] & "")
            Else
                ITM.ListSubItems.add(, , ![Curr] & " " & Format(Val(![TotalAmt] & ""), "#.###")).Tag = Val(![TotalAmt] & "")
            End If
            ITM.SubItems(6) = !ShippingMode & ""
            If IsNull(!GatePassDT) Then
                ITM.SubItems(7) = "-"
            Else
                ITM.SubItems(7) = Format(!GatePassDT, "dd-MMM-yyyy")
            End If
            
            If IsNull(!DTMaturity) Then
                ITM.SubItems(8) = "-"
            Else
                ITM.SubItems(8) = Format(!DTMaturity, "dd-MMM-yyyy")
            End If
            If strCompany = "IAA" Or strCompany = "Kami" Then
                ITM.ListSubItems.add , , !InvoiceInternalOrders & ""
            Else
                ITM.ListSubItems.add , , !InvoiceOrders & ""
            End If
            
            ITM.ListSubItems.add , , !DispatchListNos & ""
            
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
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    FAcc.Top = (Me.ScaleHeight - FAcc.Height) / 2
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

    If LV.ListItems.count = 0 Then Exit Sub
    
    If iColNo = iMaturityColIndex Or iColNo = iGPColIndex Then
        DTEdit.Value = Date
        DTEdit.Move LV.ColumnHeaders(iColNo).Left + LV.Left, LV.Top + LV.SelectedItem.Top
        DTEdit.Visible = True
        DTEdit.SetFocus
        Exit Sub
    End If
    
    If bShipping Then Exit Sub
     
    If strCompany = "Dr-Frgz" Then
        'Stop Editing of Older Invoices which were created without DispatchList.
        
'        If CDate(LV.SelectedItem.SubItems(4)) <= #12/31/2010# Then
'            MsgBox "Invoie is Locked.", vbInformation
'            Exit Sub
'        End If
        Dim strInvoiceNo As String
        strInvoiceNo = LV.SelectedItem.SubItems(3)
        If Val(Mid(strInvoiceNo, 5, InStr(strInvoiceNo, "/") - 5)) <= 73 Then
            If MsgBox("Invoice is Locked. Do you want to Edit it anyway ?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
            MsgBox "Still Not Allowed.", vbInformation
            Exit Sub
                'If getDBPassword("EditOlderInvoices") = False Then Exit Sub
        End If
    End If

    If strCompany = "IAA" Or strCompany = "Kami" Then
        If getDBPassword("EditExportDocs") = False Then Exit Sub
    End If
    
    Dim f As New frmNewCustomInvoice_Adv 'frmNewCustomInvoice
    Load f
    Call f.EditProforma(Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1), bCustomInvoice)
    
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        With LV.ColumnHeaders(i)
            If x > Val(.Left) And x < (Val(.Left) + Val(.Width)) Then
                iColNo = i
                Exit For
            End If
        End With
    Next
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 2 Then Exit Sub
    If LV.ListItems.count = 0 Then
        Call ShowMenus(False)
    Else
        Call ShowMenus(True)
    End If
'    If bShipping Then
'        ShowMenusForShipping
'    ElseIf bCourierInvoice Then
'        ShowMenusforCourier
'    Else
'        If LV.ListItems.Count = 0 Then
'            Call ShowMenus(False)
'        Else
'            Call ShowMenus(True)
'        End If
'    End If
    'If LV.ListItems.Count = 0 Or Button <> 2 Or OutOfFinancialYear Then Exit Sub

    'if a Voucher Entry Show Delete Menu
    'If InStr(LV.SelectedItem, "-") > 0 Then Me.PopupMenu mnuPop
    
End Sub

Private Sub ShowMenusForShipping()

    On Error Resume Next
    Dim c As Control
    For Each c In Controls
        If TypeOf c Is Menu Then
            If c.Name <> "mnuPrintPList" Then
                c.Visible = False
            End If
        End If
    Next
    
    On Error GoTo 0
    
    If LV.ListItems.count > 0 Then
        mnuPrintPList.Visible = True
        mnuCustomPList.Visible = True
        mnuCustomPListOther.Visible = True
        mnuCommercialPList.Visible = True
        mnuCommercialPListOther.Visible = True
        mnuDashPackingLabels.Visible = True
        mnuPrintPackingLabels.Visible = True
    End If
'    mnuPrintDecDocs.Visible = False
    Me.PopUpMenu mnuPop
    
End Sub

Private Sub ShowMenus(vis As Boolean)
        
    mnuPrint5.Visible = False
    mnuPrint10.Visible = False
    Dim strCustomer As String
    If strCompany = "PAKSMITH GROUP" Then
        mnuUpdateCustomPrice.Visible = True
    Else
        mnuUpdateCustomPrice.Visible = False
    End If
'    If LV.ListItems.Count > 0 Then
'        strCustomer = LV.SelectedItem.SubItems(1)
'        If strCustomer = "FAIR TRADE" Or strCustomer = "EquoSolidale" Or strCustomer = "Sportsystem" Or strCustomer = "METRO" Or strCustomer = "ASA Sports" Then 'Or strCustomer = "SCANMAN OY"
'            mnuPrint5.Visible = True
'            mnuPrint10.Visible = True
'        End If
'    End If
    
    'd1.Visible = vis

    'mnuPrintCoverLetter.Visible = False
    'PrintDecathDash.Visible = False
    'mnuPrintDecDocs.Visible = False
    
    d6.Visible = vis
    
    If iInvoiceType = 2 Then
        mnuEdit.Visible = False 'vis
    Else
        mnuEdit.Visible = vis
    End If
    'mnuImpMaterial.Visible = vis
 
    'mnuPrint.Visible = vis
    'mnuPrintComInv.Visible = vis
    If bSampleInvoice Then
        mnuAdd.Caption = "&Add New Sample Invoice"
        mnuDashCustomerFeedback.Visible = True
        mnuCustomerFeedback.Visible = True
    Else
        mnuAdd.Caption = "&Add New Custom Invoice"
        mnuDashCustomerFeedback.Visible = False
        mnuCustomerFeedback.Visible = False
    End If
    If LV.ListItems.count > 0 Then
'        mnuImpMaterial.Visible = True
'        mnuImpMaterial.Caption = IIf(Val(LV.SelectedItem.ListSubItems(1).Tag) > 0, "Edit &Imported Material", "Auto Calc &Imported Material")
'        mnuPacking.Visible = True
        
'        If Val(LV.SelectedItem.ListSubItems(2).Tag) > 0 Then
'            mnuPacking.Caption = "Edit &Packing Info."
'            mnuDelPackingInfo.Visible = True
'            mnuPrintPList.Visible = True
'        Else
'            mnuPacking.Caption = "Add &Packing Info."
'            mnuDelPackingInfo.Visible = False
'            mnuPrintPList.Visible = False
'        End If
        
'        If Val(LV.SelectedItem.ListSubItems(1).Tag) > 0 Then
'            mnuPrintInvWDetail.Visible = True
'        Else
'            mnuPrintInvWDetail.Visible = False
'        End If
        
'        If LV.SelectedItem.SubItems(1) = "DECATHLON" Then
'            mnuPrintCoverLetter.Visible = True
'            PrintDecathDash.Visible = True
'            mnuPrintDecDocs.Visible = True
'        End If
        
    Else
        'mnuPrintInvWDetail.Visible = False
        'mnuPacking.Visible = False
       'mnuImpMaterial.Visible = False
        'mnuPrintPList.Visible = False
        'mnuDelPackingInfo.Visible = False
    End If
       
    'mnuAddShippInfo.Visible = IIf(Left(mnuImpMaterial.Caption, 4) = "Edit", True, False)
    'dashshipinfo.Visible = mnuAddShippInfo.Visible
    If strCompany = "Instrumed" Then
        mnuCustomPListCombined.Visible = False
        mnuCustomPListCombinedCustom.Visible = False
        mnuGenerateBatches.Visible = False
        mnuPrintSpecialCustomInvoice.Visible = False
    End If
    Me.PopUpMenu mnuPop
    
End Sub

Private Sub ShowMenusforCourier()
         
    On Error Resume Next
    Dim c As Control
    For Each c In Controls
        If TypeOf c Is Menu Then
            c.Visible = False
        End If
    Next
    
    On Error GoTo 0
    mnuAdd.Visible = True
    If LV.ListItems.count > 0 Then
        dd.Visible = True
        mnuEdit.Visible = True
        mnuComDocs.Visible = True
        mnuPrintComInv.Visible = True
    End If
    mnuPrintDecDocs.Visible = False
    Me.PopUpMenu mnuPop
    
End Sub


Private Sub mnuAdd_Click()
    'Load frmNewCustomInvoice
    'frmNewCustomInvoice.Show
    'ShowForm frmNewCustomInvoice, MainForm, False, "New Custom Invoice"
    'frmNewCustomInvoice.AddNewProforma
    Load frmNewCustomInvoice_Adv
    frmNewCustomInvoice_Adv.AddNewInvoice True, bCourierInvoice, bSampleInvoice
End Sub

Private Sub mnuAddShippInfo_Click()

    Load frmEditShippInfo
    'ShowForm f, MainForm, , "Shipp. Info"
    Call frmEditShippInfo.EditEntry(LV.SelectedItem.SubItems(3))
    
End Sub

Private Sub mnuBillOfExchange_Click()
    Call PrintBillOfExchange
End Sub

Private Sub PrintBillOfExchange(Optional Firstrpt As Boolean = True)

    On Error GoTo err
    Me.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptBillOfExchange.rpt")
    rpt.FormulaFields(3).Text = Firstrpt
    Dim frm As New frmPrevRpt
    frm.ShowReport "{CustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    Me.MousePointer = vbNormal
    
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description

End Sub

    
Private Sub mnuBillOfExchange2_Click()
    Call PrintBillOfExchange(False)
End Sub

Private Sub mnuCommercialPList_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3), False, True, LV.SelectedItem.SubItems(1) = "DECATHLON", LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2))
End Sub

Private Sub mnuCommercialPListOther_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3), True, True, LV.SelectedItem.SubItems(1) = "DECATHLON", LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2))
End Sub

Private Sub mnuCoveringLetter_Click()
    
    On Error GoTo err
    Me.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptComCoveringLetter.rpt")
    If LCase(CurrentUserName) = "mnaeem" Then
        rpt.FormulaFields(3).Text = "'Naeem Ahmed'"
    End If
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "{CustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub mnuCustomerFeedback_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    Dim f As New frmSampleInvoice_CustomerFeedback
    Load f
    
    f.AddFeedback LV.SelectedItem.SubItems(3)
    
End Sub

Private Sub mnuCustomPList_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3))
End Sub

Private Sub mnuCustomPListCombined_Click()

    On Error GoTo err

    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    If bCustomInvoice Then
        Set rpt = rptApp.OpenReport(RptPath & "\rptCustomPListDetailed.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\Commercial_Packing_List_Detailed.rpt")
    End If
    'rpt.FormulaFields(15).Text = bImpMaterialDetail
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VCustomPList.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt
    
    Screen.MousePointer = vbDefault


    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub mnuCustomPListCombinedCustom_Click()

    On Error GoTo err

    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    
    If bCustomInvoice Then
        Set rpt = rptApp.OpenReport(RptPath & "\rptCustomPListDetailedCustom.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\Commercial_Packing_List_Detailed_Custom.rpt")
    End If
    'rpt.FormulaFields(15).Text = bImpMaterialDetail
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VCustomPList.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt
    
    Screen.MousePointer = vbDefault


    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub mnuCustomPListOther_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3), True)
End Sub

Private Sub mnuDeclarationOfConformity_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Declartion_Of_Conformity.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VCustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt

End Sub

Private Sub mnuDeleteInvoice_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(*)", "CustomInvoiceAuthorization", " WHERE CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'")
    If lCount > 0 Then
        MsgBox "Invoice is Authorized, Can't Delete.", vbInformation
        Exit Sub
    End If
    
    If MsgBox("Are you sure to Delete Invoice & Packing List?", vbQuestion + vbYesNo) = vbYes Then
        If getDBPassword("DeleteInvoice") = False Then Exit Sub
        Call StartTrans(con)
        con.Execute "DELETE FROM CustomInvoice WHERE CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'"
        con.Execute "DELETE FROM CustomInvoiceItems WHERE CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'"
        con.Execute "DELETE FROM CustomPList WHERE CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'"
        con.CommitTrans
    End If

End Sub

Private Sub mnuDelPackingInfo_Click()
    
    If LV.ListItems.count = 0 Then Exit Sub
    Dim iAffect As Integer
    If MsgBox("Are you sure to Delete Packing List of Invoice # " & LV.SelectedItem.SubItems(3) & " ?", vbQuestion + vbYesNo) = vbYes Then
        Call StartTrans(con)
        con.Execute "Delete From CustomPList Where CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'", iAffect
        con.CommitTrans
        If iAffect > 0 Then
            LV.SelectedItem.ListSubItems(2).Tag = 0
        End If
    End If
    
End Sub

Private Sub mnuedit_Click()
    Call LV_DblClick
End Sub

Private Sub mnuEForm_Click()

    Load frmPrintFormE
    If LV.ListItems.count > 0 Then
        frmPrintFormE.ShowForm LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
    Else
        frmPrintFormE.Show , MainForm
    End If
    
End Sub

Private Sub mnuGenerateBatches_Click()

    Dim cmd As New ADODB.Command, iRecAffect As Integer
    With cmd
        Set .ActiveConnection = con_ServerSide
        .CommandType = adCmdStoredProc
        .CommandText = "AddBatchNosToInvoice"
        .Parameters("@CustomInvoice").Value = LV.SelectedItem.SubItems(3)
        .Execute iRecAffect
    End With
    If iRecAffect > 0 Then
        MsgBox "Generated.", vbInformation
        Exit Sub
    End If
    
End Sub

Private Sub mnuGSP_Click()

'    Load frmGSP
'    If LV.ListItems.Count > 0 Then
'        frmGSP.ShowForm LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
'    Else
'        frmGSP.Show , MainForm
'    End If
    
End Sub

Private Sub mnuInvPListComparison_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\InvoicePackingListComparison.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{CustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    
End Sub

Private Sub mnuPacking_Click()

'    If Left(mnuPacking.Caption, 3) = "Add" Then
'
'        Load frmNewPList
'        frmNewPList.ShowMe LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
'    Else
'        Load frmNewPList
'        Call frmNewPList.EditPList(LV.SelectedItem.SubItems(3))
'        ShowForm frmNewPList, MainForm, , "Packing List"
'    End If
    
End Sub

Private Sub mnuPrint_Click()
    Call PrintCustomInvoice
End Sub

Private Sub PrintCustomInvoice(Optional bImpMaterialDetail As Boolean = False, Optional bWithBatchNo As Boolean = False, Optional bSpecialCustomInvoice As Boolean = False, Optional dDiscount As Double = 0, Optional bWithoutAddress As Boolean)

    On Error GoTo err

    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    If bCustomInvoice Then
        If bSpecialCustomInvoice Then
            Set rpt = rptApp.OpenReport(RptPath & "\rptSpecialCustomInvoice.rpt")
        Else
            Set rpt = rptApp.OpenReport(RptPath & "\rptCustomInvoice.rpt")
        End If
    Else
'        If bWithBatchNo Then
'            Set rpt = rptApp.OpenReport(RptPath & "\rptCommercialInvoiceWithBatchNo.rpt")
'        Else
'            Set rpt = rptApp.OpenReport(RptPath & "\rptCommercialInvoice.rpt")
'        End If
        If iInvoiceType = 2 Then
            Set rpt = rptApp.OpenReport(RptPath & "\Bank_Invoice.rpt")
        Else
            Set rpt = rptApp.OpenReport(RptPath & "\rptCommercialInvoice.rpt")
            rpt.FormulaFields.GetItemByName("ShowBatchNo").Text = bWithBatchNo
        End If
    End If
    'rpt.FormulaFields(15).Text = bImpMaterialDetail
    
    
'    If dDiscount > 0 Then
'        rpt.FormulaFields.GetItemByName("Discount").Text = dDiscount
'    End If
    rpt.FormulaFields.GetItemByName("Discount").Text = dDiscount
    
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt
    
    Screen.MousePointer = vbDefault


    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub mnuPrint10_Click()
    Call PrintCommercialInvoice(True, False)
End Sub

Private Sub mnuPrint5_Click()
    Call PrintCommercialInvoice(True, True)
End Sub

Private Sub mnuPrintActualInvoice_Click()

    On Error GoTo err

    
    Dim rpt As CRAXDDRT.Report
    
    Set rpt = rptApp.OpenReport(RptPath & "\Print_Actual_Invoice.rpt")
    
    rpt.FormulaFields.GetItemByName("Discount").Text = dDiscount
    
    Dim f As New frmPrevRpt
    
    f.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt
    

    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintBankSalesContract_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Bank_Sales_Contract.rpt")
    
    Dim strSelection As String
    strSelection = "{VrptProformas.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'"
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub mnuPrintBankSlipLabTest_Click()

    Dim dAmt As Double, strAmt As String
    Dim frmGet As New frmGetSingleTextValue
    If strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
        If frmGet.getTextValue(strAmt, False, "Exch. Rate", "Exch. Rate for Lab Test") = False Then Exit Sub
    Else
        If frmGet.getTextValue(strAmt, False, "Amount", "Amount for Lab Test") = False Then Exit Sub
    End If
    
    dAmt = Val(strAmt)
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\NationalBankSlip.rpt")
    
    If strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
        rpt.FormulaFields.GetItemByName("ExchRate").Text = dAmt
    Else
        rpt.FormulaFields.GetItemByName("Amount").Text = dAmt
    End If
    
    Dim f As New frmPrevRpt
    f.ShowReport "{CustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    
End Sub

Private Sub mnuPrintBankUnderTaking1_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Bank_Undertaking1.rpt")
    
    Dim strSelection As String
    strSelection = "{VrptProformas.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'"
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt

End Sub

Private Sub mnuPrintBankUnderTaking2_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Bank_Undertaking2.rpt")
    
    Dim strSelection As String
    strSelection = "{VrptProformas.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'"
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
End Sub

Private Sub mnuPrintBankUnderTaking3_Click()
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Bank_Undertaking3.rpt")
    
    Dim strSelection As String
    strSelection = "{VrptProformas.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'"
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
End Sub

Private Sub mnuPrintBillofExchange_Click()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptBillOfExchange.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VRptProformas.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintCertificateOfCompliance_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Certificate_Of_Compliance.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VCustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt


End Sub

Private Sub mnuPrintComInv_Click()
    Call PrintCommercialInvoice
End Sub

Private Sub PrintCommercialInvoice(Optional bPer As Boolean = False, Optional bFivePer As Boolean = True)
    
    On Error GoTo err

    Screen.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt
    Dim rpt As CRAXDDRT.Report
    
    If UCase(LV.SelectedItem.SubItems(1)) = "MIKASA" And LV.SelectedItem.SubItems(2) = "JAPAN" Then
        Set rpt = rptApp.OpenReport(rptexportpath & "\rptCommercialInvoiceMikasaJP.rpt")
    ElseIf bPer Then
        Set rpt = rptApp.OpenReport(rptexportpath & "\rptCommercialInvoiceFairTrade.rpt")
        rpt.FormulaFields(22).Text = bFivePer
    ElseIf LV.SelectedItem.SubItems(1) = "DECATHLON" Then
        Set rpt = rptApp.OpenReport(rptexportpath & "\rptDecathlonCommercialInvoice.rpt")
    Else
        Set rpt = rptApp.OpenReport(rptexportpath & "\rptCommercialInvoice.rpt")
    End If
    f.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt
    
    If Not bPer Then
        If LV.SelectedItem.SubItems(2) = "Canada" Then
            Dim rpt1 As CRAXDDRT.Report
            Set rpt1 = rptApp.OpenReport(rptexportpath & "\rptCommercialInvoice.rpt")
            Dim f1 As New frmPrevRpt
            Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
            Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
            
            Set FormulaFields = rpt1.FormulaFields
            
            For Each FormulaField In FormulaFields
                If FormulaField.Name = "{@forHeading}" Then
                    FormulaField.Text = "'C A N A D I A N   C U S T O M S   I N V O I C E'"
                    Exit For
                End If
            Next
            f1.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt1
        End If
    End If
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
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptFProformaList.rpt")
    rpt.FormulaFields(3).Text = "'" & cmbCust.Text & "'"
    rpt.FormulaFields(4).Text = "'" & strFromTo & "'"
    
    frmPrevRpt.ShowReport "", rpt
    
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub mnuPrintCoverLetter_Click()

    On Error GoTo err

    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptDecathlonCoveringLetter.rpt")
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VCustomInvoice.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt

    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub mnuPrintDecDocs_Click()
    Load frmPrintDecathlonDocs
    frmPrintDecathlonDocs.ShowMe (LV.SelectedItem.SubItems(3))
End Sub

Private Sub mnuPrintGSPInvoice_Click()
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rpt_GSP_Invoice.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VrptProformas.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
End Sub

Private Sub mnuPrintInvoiceasHandTools_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Invoice_Handtools.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VrptProformas.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt

End Sub

Private Sub mnuPrintInvoiceDiscount_Click()

    Dim strDiscount As String
    Dim f As New frmGetSingleTextValue
    If f.getTextValue(strDiscount, , "Discount", "Discount") = False Then Exit Sub
         
    Dim dDiscount As Double
    dDiscount = Val(strDiscount)
     
    Dim frm As New frmPrevRpt
    Call PrintCustomInvoice(False, False, False, dDiscount)
    
End Sub

Private Sub mnuPrintInvoiceExport_Click()

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    
     
    Set rpt = rptApp.OpenReport(RptPath & "\Commercial_Invoice_Export.rpt")
     
    Dim f As New frmPrevRpt
    
    f.ShowReport "{VCustomInvoice.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt
  
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintInvoiceWOCustomerAddress_Click()

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    
    Dim strDiscount As String
    Dim f As New frmGetSingleTextValue
    If f.getTextValue(strDiscount, , "Discount", "Discount") = False Then Exit Sub
         
    Dim dDiscount As Double
    dDiscount = Val(strDiscount)
    If bCustomInvoice Then
        Set rpt = rptApp.OpenReport(RptPath & "\rptSampleInvoiceWOAddress.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\Commercial_Invoice_WO_Address.rpt")
    End If
    rpt.FormulaFields.GetItemByName("Discount").Text = dDiscount
    
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt
  
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintInvWDetail_Click()

    Call PrintCustomInvoice(True)
    
End Sub

Private Sub mnuPrintLabLetter_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\LaboratoryTestLetter.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{CustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "' AND {CustomInvoiceItems.ItemForLabTest}=TRUE", rpt
    
End Sub

Private Sub mnuPrintLabReport_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Bank_Lab_Report.rpt")
    
    Dim strSelection As String
    strSelection = "{VrptProformas.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'"
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub mnuPrintLetterToBank_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptLetterToBank.rpt")
    
    Dim strSelection As String
    strSelection = "{VrptProformas.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'"
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub mnuPrintLetterToBankSimple_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptLetterToBankSimple.rpt")
    
    Dim strSelection As String
    strSelection = "{VrptProformas.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'"
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub mnuPrintList_Click()
        
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\CustomInvoiceList.rpt")
    Dim strSelection As String
    strSelection = "{VCustomInvoiceList.DT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#"
    If cmbCust.ID <> "0" Then
        strSelection = strSelection & " AND {VCustomInvoiceList.CustCode}='" & cmbCust.ID & "'"
    End If
    
    If cmbCountry.ListIndex > 0 Then
        strSelection = strSelection & " AND {VCustomInvoiceList.Country}='" & cmbCountry.Text & "'"
    End If
    
    rpt.FormulaFields.GetItemByName("Customers").Text = "'" & cmbCust.Text & " " & cmbCountry.Text & "'"
    rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & Format(DtFrom, "dd-MMM-yyyy") & " to " & Format(DtTo.Value, "dd-MMM-yyyy") & "'"
    rpt.FormulaFields.GetItemByName("bCustom").Text = bCustomInvoice
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt

End Sub

Private Sub mnuPrintPackingLabels_Click()

    On Error GoTo err
    
    If LV.SelectedItem.SubItems(1) <> "DECATHLON" Then
        MsgBox "Packing Labels are available only for DECATHLON at the moment.", vbInformation
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report, f As New frmPrevRpt
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptDecathlonPackingLabels.rpt")
    f.ShowReport "{VCustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    
    Me.MousePointer = vbNormal
    
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub mnuPrintPackingList_Click()
    Call PrintPackingList(LV.SelectedItem.SubItems(3), bCustomInvoice, iInvoiceType)
End Sub

Private Sub mnuPrintPaymentDocs_Click()

    On Error GoTo err

    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    
    Set rpt = rptApp.OpenReport(RptPath & "\Print_Payment_Docs.rpt")
    
    Dim f As New frmPrevRpt
    f.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt
    
    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub mnuPrintPaymentInvoice_Click()

    On Error GoTo err

    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    
    Set rpt = rptApp.OpenReport(RptPath & "\Payment_Invoice.rpt")
    
    Dim f As New frmPrevRpt
    f.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt
    
    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub mnuPrintReasonForDelay_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Bank_Reason_For_Delay.rpt")
    
    Dim strSelection As String
    strSelection = "{VrptProformas.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'"
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub mnuPrintReasonForDelaySBP_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Bank_Reason_For_Delay_SBP.rpt")
    
    Dim strSelection As String
    strSelection = "{VrptProformas.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'"
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub mnuPrintSpecialCustomInvoice_Click()
    Call PrintCustomInvoice(, , True)
End Sub

Private Sub mnuPrintUnderTaking_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Undertaking.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VCustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt

End Sub

Private Sub mnuPrintUSAInvoice_Click()

    On Error GoTo err

    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    If bCustomInvoice Then
        Set rpt = rptApp.OpenReport(RptPath & "\Invoice_USA.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\Commercial_Invoice_USA.rpt")
    End If
    
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt
    
    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintUSAPaymentInvoice_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Commercial_USA_Payment_Invoice.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VrptProformas.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt


End Sub

Private Sub mnuPrintWithBatchNo_Click()
    Call PrintCustomInvoice(, True)
End Sub

Private Sub mnuShippingInstructions_Click()
'    Load frmShippingInstructions
'    If LV.ListItems.Count = 0 Then
'        frmShippingInstructions.Show , MainForm
'    Else
'        frmShippingInstructions.ShowForm LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
'    End If
End Sub

Private Sub mnuUpdateCustomPrice_Click()
    If LV.ListItems.count = 0 Then Exit Sub
    If MsgBox("Are you Sure, You want to Update Custom Prices?", vbInformation + vbYesNo) = vbNo Then Exit Sub
    
    Dim dCustomPrice As Double
    dCustomPrice = Val(con.Execute("Select SUM(Qty*Price)/SUM(Qty) From CustomInvoiceItems Where CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'").Fields(0).Value & "")
    con.Execute "Update CustomInvoiceItems SET CustomPrice=" & dCustomPrice & " WHERE CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'"
    
End Sub

Private Sub mnuUpdateRebateRate_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    
    Dim strInvoiceNo As String
    strInvoiceNo = LV.SelectedItem.SubItems(3)
    
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(*)", "Rebate", " WHERE CustomInvoice='" & strInvoiceNo & "'")
    
    If lCount > 0 Then
        MsgBox "Can't Update Rate, Rebate Already Generated.", vbInformation
        Exit Sub
    End If
    
    Dim iRecAffect As Integer
    If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbYes Then
        con.Execute "UPDATE CustomInvoiceItems SET RebateRate=CRebate FROM CustomInvoiceItems INNER JOIN " & _
         "FProformaOrders ON CustomInvoiceItems.RefID=FProformaOrders.EntryID INNER JOIN " & _
         "FOrderItems ON FProformaOrders.OrderEntryID=FOrderItems.ID INNER JOIN " & _
         "Items ON FOrderItems.CompItemCode=Items.ItemID INNER JOIN " & _
         "ItemCatagories ON Items.CatID=ItemCatagories.CatID INNER JOIN " & _
         "SROs ON ItemCatagories.SRONo=SROs.EntryID WHERE CustomInvoice='" & strInvoiceNo & "'", iRecAffect
         
        If iRecAffect > 0 Then
            MsgBox "Updated.", vbInformation
        End If
    End If
    
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
    
        con.Execute "Update Vouchers Set [Description]='" & txtDesc & "' Where Sno=" & Val(LV.SelectedItem.key)
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
