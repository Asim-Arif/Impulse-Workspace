VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.MDIForm MainForm 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Contractor Package"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10725
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   5325
      Top             =   2430
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4515
      Top             =   2430
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4920
      Top             =   2430
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   5730
      Top             =   2910
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5745
      Top             =   2415
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Please Select Payroll Database to Connect."
      Filter          =   """Payroll Database|Payroll.mdb"""
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5115
      Top             =   2835
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":0000
            Key             =   "Cont"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":0502
            Key             =   "PaySheet"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":05A1
            Key             =   "PayEntry"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":067F
            Key             =   "Issue"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":0754
            Key             =   "Recieve"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":07F6
            Key             =   "Items"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":08D9
            Key             =   "ContLedger"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   4515
      Top             =   2835
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   108
      ImageHeight     =   17
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":0EBB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   5115
      Top             =   3390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   150
      ImageHeight     =   17
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":1028
            Key             =   "Accounts"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":22B2
            Key             =   "test"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":2553
            Key             =   "Company"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":2FBA
            Key             =   "Glossary"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":398D
            Key             =   "Developers"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":4557
            Key             =   "Export"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":56F9
            Key             =   "Management"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":6494
            Key             =   "Orders"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":75AB
            Key             =   "Payroll"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":88A7
            Key             =   "Policies"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":99AE
            Key             =   "Production"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":AB69
            Key             =   "Purchases"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":B618
            Key             =   "Stats"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":C035
            Key             =   "Inventory"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":D15F
            Key             =   "BusPart"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4515
      Top             =   3390
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   130
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":E23B
            Key             =   "test"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":F90D
            Key             =   "Company"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":10374
            Key             =   "Glossary"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":10D47
            Key             =   "Management"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":11AE2
            Key             =   "Policies"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":12BE9
            Key             =   "Purchases"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":13698
            Key             =   "BusPart"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":14774
            Key             =   "Export"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":15916
            Key             =   "Developers"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":164E0
            Key             =   "Inventory"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":1760A
            Key             =   "Orders"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":18721
            Key             =   "Payroll"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":19A1D
            Key             =   "Production"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":1ABD8
            Key             =   "Stats"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":1B5F5
            Key             =   "Accounts"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuSubOptions 
         Caption         =   "Use Calculator"
         Index           =   1
      End
      Begin VB.Menu mnuSubOptions 
         Caption         =   "User Manager"
         Index           =   2
      End
      Begin VB.Menu mnuSubOptions 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuSubOptions 
         Caption         =   "Change Password"
         Index           =   4
      End
      Begin VB.Menu mnuSubOptions 
         Caption         =   "Log Off"
         Index           =   5
      End
      Begin VB.Menu mnuSubOptions 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuSubOptions 
         Caption         =   "E&xit"
         Index           =   7
      End
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "General Settings"
      Visible         =   0   'False
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Set Initial Lot Process"
         Index           =   1
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Percentage Ranges"
         Index           =   3
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "&Components"
         Index           =   5
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "&Manufacturing Processes"
         Index           =   7
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "&FDA Processes"
         Index           =   8
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Steel Vendors"
         Index           =   10
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Test Labs"
         Index           =   11
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Measuring &Units"
         Index           =   13
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Compositions"
         Index           =   14
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Cities && Ports"
         Index           =   16
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Bank List"
         Index           =   18
         Shortcut        =   ^B
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSubGeneral 
         Caption         =   "Payment Terms"
         Index           =   19
      End
   End
   Begin VB.Menu mnupopdept 
      Caption         =   "popMenuDept"
      Visible         =   0   'False
      Begin VB.Menu mnuDeptRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu popdeptdash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowDept 
         Caption         =   "Show Inactive Departments"
      End
      Begin VB.Menu mnuDeptdash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeptInact 
         Caption         =   "Make This Department Inactive"
      End
      Begin VB.Menu mnuDeptdash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeptnew 
         Caption         =   "Add &New Department"
      End
      Begin VB.Menu mnuDepdash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeptEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuDeptdash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDeptPrint 
         Caption         =   "&Print Department List"
      End
   End
   Begin VB.Menu mnuVenders 
      Caption         =   "Venders"
      Visible         =   0   'False
      Begin VB.Menu mnuVendList 
         Caption         =   "&Vender List"
      End
      Begin VB.Menu vendlistdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVendAssItems 
         Caption         =   "Assigned Items"
      End
      Begin VB.Menu VendAssItemsdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVendPymt 
         Caption         =   "Vender Advance Payment"
      End
      Begin VB.Menu mnuVendPaySheetdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVendPaySheet 
         Caption         =   "Vender Payment Sheet"
      End
      Begin VB.Menu mnuAdvdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVendLoans 
         Caption         =   "Loans"
         Begin VB.Menu mnuVendAdvShort 
            Caption         =   "New Short Term Advance"
         End
         Begin VB.Menu mnuVendAdvLong 
            Caption         =   "New Long Term Loan"
         End
         Begin VB.Menu mnuVendLoanLedgersdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuVendShortLoanLedger 
            Caption         =   "Short Term Loan Ledger"
         End
         Begin VB.Menu mnuVendLongLoanLedger 
            Caption         =   "Long Term Loan Ledger"
         End
      End
      Begin VB.Menu VendLedgerdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVendLedger 
         Caption         =   "Vender Ledger"
      End
   End
   Begin VB.Menu mnuRMVend 
      Caption         =   "Raw Material"
      Visible         =   0   'False
      Begin VB.Menu mnuRMGroups 
         Caption         =   "RM Groups"
      End
      Begin VB.Menu mnuRM 
         Caption         =   "Raw Material"
      End
      Begin VB.Menu RMPOdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRMPO 
         Caption         =   "New Purchase Order"
      End
      Begin VB.Menu mnuPOrderList 
         Caption         =   "Purchase Order List"
      End
      Begin VB.Menu RRMdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecRM 
         Caption         =   "Receive Raw Material"
      End
      Begin VB.Menu mnuIssRM 
         Caption         =   "Issue Raw Material"
      End
   End
   Begin VB.Menu mnuHICORes 
      Caption         =   "Hico Resources"
      Visible         =   0   'False
      Begin VB.Menu mnuHRGroups 
         Caption         =   "Hico Resources Groups"
      End
   End
   Begin VB.Menu mnuCompanya 
      Caption         =   "Company"
      Visible         =   0   'False
      Begin VB.Menu mnuItems 
         Caption         =   "Items"
         Begin VB.Menu mnuItmGrps 
            Caption         =   "Item Groups"
         End
         Begin VB.Menu mnuItemsFinQuality 
            Caption         =   "Item's Finishing Quality"
         End
         Begin VB.Menu mnuItemTypes 
            Caption         =   "Item Types"
         End
      End
   End
   Begin VB.Menu mnuCustomers 
      Caption         =   "Customers"
      Visible         =   0   'False
      Begin VB.Menu mnuNewOrd 
         Caption         =   "New Order"
      End
      Begin VB.Menu mnuNewStockOrder 
         Caption         =   "New Stock Order"
      End
      Begin VB.Menu mnuCustOrdList 
         Caption         =   "Customer Order List"
      End
      Begin VB.Menu custdash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrdStatus 
         Caption         =   "Order Status"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOrdStatus1 
         Caption         =   "Order Status(Other Format)"
      End
      Begin VB.Menu mnuOrdStatusComp 
         Caption         =   "Order Status Complete"
         Visible         =   0   'False
      End
      Begin VB.Menu custdash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewProforma 
         Caption         =   "New Proforma"
      End
      Begin VB.Menu mnuCustProformaList 
         Caption         =   "Customer Proforma List"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Visible         =   0   'False
      Begin VB.Menu mnuDailyRpts 
         Caption         =   "Daily Reports"
      End
      Begin VB.Menu mnuDailyRptsdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIssRpts 
         Caption         =   "Issue Reports"
      End
      Begin VB.Menu mnuRcvRpts 
         Caption         =   "Receive Reports"
      End
      Begin VB.Menu mnuRcvRptsdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRMListrpt 
         Caption         =   "Raw Material List"
      End
      Begin VB.Menu mnuRMVendListrpt 
         Caption         =   "Raw Material Vender List"
      End
      Begin VB.Menu mnuRMIssRpts 
         Caption         =   "Raw Material Issue Reports"
      End
      Begin VB.Menu mnuRMRcvRpts 
         Caption         =   "Raw Material ReceiveReports"
      End
      Begin VB.Menu mnuRMRcvRptdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHRrpt 
         Caption         =   "Hico Resources"
      End
      Begin VB.Menu mnuAssHRrpt 
         Caption         =   "Assigned Hico Resources"
      End
      Begin VB.Menu mnuCustListrptdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCustListrpt 
         Caption         =   "Customer List"
      End
      Begin VB.Menu mnuReqrptdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReqrpt 
         Caption         =   "Requsition Report"
      End
      Begin VB.Menu mnuPORpt 
         Caption         =   "P.O. Report"
      End
      Begin VB.Menu mnuPORptdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProdPlanrpt 
         Caption         =   "Production Planning Report"
      End
      Begin VB.Menu mnuPPlanrptdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVendListrpt 
         Caption         =   "Vender List"
      End
   End
   Begin VB.Menu mnuProdPhases 
      Caption         =   "Production Phases"
      Visible         =   0   'False
      Begin VB.Menu mnuReq 
         Caption         =   "Requsition"
         Visible         =   0   'False
      End
      Begin VB.Menu prodplandash 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuProdPlan 
         Caption         =   "Prdoduction Planning"
      End
      Begin VB.Menu TestDash 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAuth 
         Caption         =   "Authorize"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAuthDash 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "FORGING"
         Index           =   1
         Begin VB.Menu mnuIssForg 
            Caption         =   "Issue Items"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecVForg 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSForg 
            Caption         =   "Receive Items From Supervisor"
         End
         Begin VB.Menu mnudForg 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastForg 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuForgSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuForgSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "MACHINING"
         Index           =   2
         Begin VB.Menu mnuIssMach 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVMach 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSMach 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudMach 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastMach 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuMachSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMachSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "DRILLING"
         Index           =   3
         Begin VB.Menu mnuIssDril 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVDril 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSDril 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudDril 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastDril 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuDrilSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDrilSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "FILLING"
         Index           =   4
         Begin VB.Menu mnuIssFill 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVFill 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSFill 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudFill 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastFill 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuFillSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFillSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "FUNTIONAL INSPECTION UN-FINISH"
         Index           =   5
         Begin VB.Menu mnuIssFunc 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVFunc 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSFunc 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudFunc 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastFunc 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuFuncSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFuncSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "HEAT TREATMENT"
         Index           =   6
         Begin VB.Menu mnuIssHeat 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVHeat 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSHeat 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudHeat 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastHeat 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuHeatSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHeatSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "POLISHING"
         Index           =   7
         Begin VB.Menu mnuIssPoli 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVPoli 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSPoli 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecWastPoli 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnudPoli 
            Caption         =   "-"
         End
         Begin VB.Menu mnuIssRepairPoli 
            Caption         =   "Issue Repair"
         End
         Begin VB.Menu mnuRcvRepairPoli 
            Caption         =   "Receive Repair"
         End
         Begin VB.Menu mnuPoliSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPoliSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "PASSIVATION"
         Index           =   8
         Begin VB.Menu mnuIsspass 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVPass 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSPass 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudPass 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastPass 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuPassSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPassSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "FUNTIONAL INSPECTION"
         Index           =   9
         Begin VB.Menu mnuIssFunt 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVFunt 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSFunt 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudFunt 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastFunt 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuFuntdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFuntSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "POLISHING && RE-BUFFING"
         Index           =   10
         Begin VB.Menu mnuIssPoRe 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVPoRe 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSPoRe 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudPoRe 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastPoRe 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuPORESSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPORESS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "ULTRASONIC CLEAING"
         Index           =   11
         Begin VB.Menu mnuIssUltr 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVUltr 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSUltr 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudUltr 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastUltr 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuUltrSSDash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuUltrSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "PACKING && LABELING"
         Index           =   12
         Begin VB.Menu mnuIssPack 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVPack 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSPack 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudPack 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastPack 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuPackSSDash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPackSS 
            Caption         =   "Stock Status"
         End
      End
      Begin VB.Menu mnuIssRcvProcs 
         Caption         =   "HOLDING && DISTRIBUTION"
         Index           =   13
         Begin VB.Menu mnuIssHold 
            Caption         =   "Issue Items"
         End
         Begin VB.Menu mnuRecVHold 
            Caption         =   "Receive Items From Venders"
         End
         Begin VB.Menu mnuRecSHold 
            Caption         =   "Receive Items From Supervisor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnudHold 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRecWastHold 
            Caption         =   "Receive Wastage"
         End
         Begin VB.Menu mnuHoldSSdash 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHoldSS 
            Caption         =   "Stock Status"
         End
      End
   End
   Begin VB.Menu mnumainSteel 
      Caption         =   "Steel"
      Visible         =   0   'False
      Begin VB.Menu mnuSteel 
         Caption         =   "&Steel List"
      End
      Begin VB.Menu mnuGuage 
         Caption         =   "Guage List"
      End
   End
   Begin VB.Menu mnuPopCust 
      Caption         =   "Customers"
      Visible         =   0   'False
      Begin VB.Menu PopSubCust 
         Caption         =   "Add New Customer"
         Index           =   0
      End
      Begin VB.Menu PopSubCust 
         Caption         =   "Customer List"
         Index           =   1
      End
      Begin VB.Menu PopSubCust 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu PopSubCust 
         Caption         =   "Customer Assigned Items"
         Index           =   3
      End
      Begin VB.Menu PopSubCust 
         Caption         =   "Customer Ledger"
         Index           =   4
      End
   End
   Begin VB.Menu mnuPopProdVend 
      Caption         =   "Venders"
      Visible         =   0   'False
      Begin VB.Menu PopSubProdVend 
         Caption         =   "Add New Vender"
         Index           =   0
      End
      Begin VB.Menu PopSubProdVend 
         Caption         =   "Vender List"
         Index           =   1
      End
      Begin VB.Menu PopSubProdVend 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu PopSubProdVend 
         Caption         =   "Vender Assigned Items"
         Index           =   3
      End
      Begin VB.Menu PopSubProdVend 
         Caption         =   "Vender Ledger"
         Index           =   4
      End
   End
   Begin VB.Menu mnuPopRMVend 
      Caption         =   "R.M. Venders"
      Visible         =   0   'False
      Begin VB.Menu PopSubRMVend 
         Caption         =   "Add New R.M. Vender"
         Index           =   0
      End
      Begin VB.Menu PopSubRMVend 
         Caption         =   "R.M. Vender List"
         Index           =   1
      End
      Begin VB.Menu PopSubRMVend 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu PopSubRMVend 
         Caption         =   "Assigned Items"
         Index           =   3
      End
      Begin VB.Menu PopSubRMVend 
         Caption         =   "R.M. Vender Ledger"
         Index           =   4
      End
   End
   Begin VB.Menu mnuPopSteelVend 
      Caption         =   "Steel Venders"
      Visible         =   0   'False
      Begin VB.Menu PopSubSteelVend 
         Caption         =   "Add New Steel Vender"
         Index           =   0
      End
      Begin VB.Menu PopSubSteelVend 
         Caption         =   "Steel Vender List"
         Index           =   1
      End
   End
   Begin VB.Menu mnupopemp 
      Caption         =   "PopMenuEmp"
      Visible         =   0   'False
      Begin VB.Menu mnuempRefresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu mnuempdash0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowInactive 
         Caption         =   "Show Inactive Employees/Departments"
      End
      Begin VB.Menu mnuempdash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMake 
         Caption         =   "Make This Employee Inactive"
      End
      Begin VB.Menu dd1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuempNewDept 
         Caption         =   "Add New &Department"
      End
      Begin VB.Menu mnuEmpNewEmp 
         Caption         =   "Add New &Employee"
      End
      Begin VB.Menu mnuempdash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmpEdit 
         Caption         =   "&Edit"
      End
      Begin VB.Menu mnuEmpdash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAssItems 
         Caption         =   "Show &Assigned Items"
      End
      Begin VB.Menu mnuEmpPrint 
         Caption         =   "&Print Employee List (Internal Salary)"
      End
      Begin VB.Menu mnuPrintEmpExt 
         Caption         =   "Print Employee List (E&xternal Salary)"
      End
      Begin VB.Menu mnuPrintEmpTrans 
         Caption         =   "Print Employee List (Company Transport)"
      End
      Begin VB.Menu dashemps 
         Caption         =   "-"
      End
      Begin VB.Menu mnuappForm 
         Caption         =   "Print Application Form"
      End
      Begin VB.Menu mnuPrintCard 
         Caption         =   "Print Employee Card"
         Begin VB.Menu mnuPrintEmpCard 
            Caption         =   "This Employee Only"
         End
         Begin VB.Menu mnuPrintAllCard 
            Caption         =   "For All Employees"
         End
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TempForeWin As Long
Dim ForeWin As Long
Dim Ready As Boolean
Private Sub dashBankLIst_Click()
    frmBankList.Show
End Sub


Private Sub BusPartToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.key
        Case "Customers"
            Me.PopUpMenu mnuPopCust, 512, CoolBar1.Width
        Case "ProdVend"
            Me.PopUpMenu mnuPopProdVend, 512, CoolBar1.Width
        Case "RMVend"
            Me.PopUpMenu mnuPopRMVend, 512, CoolBar1.Width
    End Select
    
End Sub

Private Sub CoolBar1_Resize()
    For i = 1 To CoolBar1.Bands.Count
        If CoolBar1.Bands(i).NewRow Then CoolBar1.Bands(i).NewRow = False
    Next i
End Sub

Private Sub MDIForm_DblClick()
    
    If con.State = 1 Then con.Close
    Set con = Nothing
    End
    
    
End Sub

Private Sub MDIForm_Load()

    Dim i As Long
    
    
    On Error Resume Next
    
    'Dim ctr As Control
    'For Each ctr In Controls
    '    If TypeOf ctr Is Menu Then
    '        ctr.Visible = False
    '    End If
    'Next ctr
    
    On Error GoTo err
    
    'Hide The Title Bar  Of Main Form.
    
    frmFlash.Show
    SetWindowLong Me.hwnd, -16, WS_DLGFRAME
    
    'PicBar.Visible = False
        
    FormsMaxWidth = Me.ScaleWidth
    'frmFlash.Move 0, 0, Me.Width, Me.Height
    
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub MDIForm_QueryUnload(cancel As Integer, UnloadMode As Integer)
    Set con = Nothing
End Sub

Private Sub MDIForm_Unload(cancel As Integer)
    On Error Resume Next
   ' Con.RollbackTrans
    'If Con.State = 1 Then Con.Close
    Set con = Nothing
End Sub

Private Sub mnuAssHRrpt_Click()
    Call ShowAssHRrpt
End Sub

Private Sub mnuAssItems_Click()
    Call ContractorList.showItems
End Sub

Private Sub mnuAuth_Click()
    frmAuthorizeIssRcv.Show
End Sub

Private Sub mnuCalc_Click()

    On Error GoTo err
    Call Shell("calc.exe")
    Exit Sub
err:
    MsgBox "Cannot Run Calculater Due To Error: " & vbNewLine & err.Description, vbInformation

End Sub

Private Sub mnuComp_Click()
    frmComponents.Show 1
End Sub

Private Sub mnuCompCatal_Click()
    frmItems.Show
End Sub

Private Sub mnuCompositions_Click()
    frmCompositions.Show
End Sub

Private Sub mnuCustCatal_Click()
    frmCustCatal.Show
End Sub

Private Sub mnuCustList_Click()
    frmFCustomerList.Show
End Sub

Private Sub mnuCustListrpt_Click()
    Call ShowCustListrpt
End Sub

Private Sub mnuCustOrdList_Click()
    frmfCustOrderList.Show
End Sub

Private Sub mnuCustProformaList_Click()
    frmFProformaList.Show
End Sub

Private Sub mnuDailyRpts_Click()
    frmDailyIssRcvRpt.Show
End Sub

Private Sub mnuDeptEdit_Click()
    If mnupopdept.Tag = "EmpDept" Then
        Call DeptList.mnuEdit(mnuDeptEdit.Tag)
    Else
        Call ContDeptList.mnuEdit(mnuDeptEdit.Tag)
    End If
End Sub

Private Sub mnuDeptInact_Click()
    If mnupopdept.Tag = "EmpDept" Then
        Call DeptList.MakeInactive
    Else
        Call ContDeptList.MakeInactive
    End If
End Sub

Private Sub mnuDeptList_Click()
    DeptList.Show
End Sub

Private Sub mnuDeptnew_Click()
    If mnupopdept.Tag = "EmpDept" Then
        Call DeptList.mnuAddDept
    Else
        Call ContDeptList.mnuAddDept
    End If
End Sub

Private Sub mnuDeptPrint_Click()
    If mnupopdept.Tag = "EmpDept" Then
        Call DeptList.PrintDeptList
    Else
        'Call ContDeptList.PrintDeptList
    End If
End Sub

Private Sub mnuDeptRefresh_Click()
    If mnupopdept.Tag = "EmpDept" Then
        Call DeptList.RefreshIt
    Else
        Call ContDeptList.RefreshIt
    End If
End Sub

Private Sub mnuDrilSS_Click()
    Call ShowStockStatus(3)
End Sub

Private Sub mnuEmpEdit_Click()
    If mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.mnuEdit(mnuEmpEdit.Tag)
    Else
        Call ContractorList.mnuEdit(mnuEmpEdit.Tag)
    End If
End Sub

Private Sub mnuempNewDept_Click()
    If Me.mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.mnuAddDept
    Else
        Call ContractorList.mnuAddDept
    End If
End Sub
Private Sub mnuEmpNewEmp_Click()
    If Me.mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.mnuAdd
    Else
        Call ContractorList.mnuAdd
    End If
End Sub


Private Sub mnuEmpPrint_Click()
    If mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.PrintEmpList
    Else
          'Call ContractorList.PrintEmpList
    End If
End Sub

Private Sub mnuempRefresh_Click()
    If mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.RefreshIt
    Else
        Call ContractorList.RefreshIt
    End If
End Sub
Private Sub mnuExit_Click()
    If con.State = 1 Then con.Close
    Set con = Nothing
    End
End Sub

Private Sub mnuFDAProcess_Click()
    frmFDAProcesses.Show 1
End Sub

Private Sub mnuFillSS_Click()
    Call ShowStockStatus(4)
End Sub

Private Sub mnuForeignCatal_Click()
    frmForeignCatalog.Show 1
End Sub

Private Sub mnuForgSS_Click()
    Call ShowStockStatus(1)
End Sub

Private Sub mnuFuncSS_Click()
    Call ShowStockStatus(5)
End Sub

Private Sub mnuFuntSS_Click()
    Call ShowStockStatus(9)
End Sub

Private Sub mnuGuage_Click()
    frmGages.Show 1
End Sub

Private Sub mnuHeatSS_Click()
    Call ShowStockStatus(6)
End Sub

Private Sub mnuHICOResources_Click()
    frmHICOR.Show
End Sub

Private Sub mnuHoldSS_Click()
    Call ShowStockStatus(13)
End Sub

Private Sub mnuHolidays_Click()
    Load Holidays
    'Holidays.Move 3000, 500
    With Holidays
        .Move (Me.Width - .Width) / 2, Me.Top + 500
        .Show
    End With
End Sub

Private Sub mnuHRGroups_Click()
    Dim f As New frmItemGroups
    Load f
    Call f.ShowForm("HRGroups", "Hico Resources Groups", "Hico Resource Name")
End Sub

Private Sub mnuHRrpt_Click()
    Call ShowHRrpt
End Sub

Private Sub mnuIssDril_Click()
    Call ShowIssForm(3)
    
End Sub

Private Sub mnuIssFill_Click()
    Call ShowIssForm(4)
End Sub

Private Sub mnuIssForg_Click()
    Call ShowIssForm(1)
End Sub

Private Sub mnuIssFunc_Click()
    Call ShowIssForm(5)
End Sub

Private Sub mnuIssFunt_Click()
    Call ShowIssForm(9)
End Sub

Private Sub mnuIssHeat_Click()
    Call ShowIssForm(6)
End Sub

Private Sub mnuIssHold_Click()
    frmVendRcvItemsHnD.Show
End Sub

Private Sub mnuIssMach_Click()
    Call ShowIssForm(2)
End Sub

Private Sub mnuIssPack_Click()
    Call ShowIssForm(12)
End Sub

Private Sub mnuIsspass_Click()
    Call ShowIssForm(8)
End Sub

Private Sub mnuIssPoli_Click()
    Call ShowIssForm(7)
End Sub

Private Sub mnuIssPoRe_Click()
    Call ShowIssForm(10)
End Sub

Private Sub mnuIssRepair_Click()
    frmVendIssRepair.Show
End Sub

Private Sub mnuIssRepairPoli_Click()
    frmVendIssRepair.Show
End Sub

Private Sub mnuIssRM_Click()
    frmRMIss.Show
End Sub

Private Sub mnuIssRpts_Click()
    Dim f As New frmRptIssRcv
    Call f.ShowForIss
End Sub

Private Sub mnuIssUltr_Click()
    Call ShowIssForm(11)
End Sub

Private Sub mnuItemsFinQuality_Click()
    Dim f As New frmMiscS
    Load f
    Call f.ShowForm("ItemsFinQuality", "Items Fin. Quality", "Finishing Quality")
End Sub

Private Sub mnuItemTypes_Click()
    Dim f As New frmMiscS
    Load f
    Call f.ShowForm("ItemTypes", "Item Types", "Item Types")
End Sub

Private Sub mnuItmGrps_Click()
    Dim f As New frmItemGroups
    Load f
    Call f.ShowForm("ItemGroups", "Item Groups", "Group Name")
End Sub

Private Sub mnuLabs_Click()
    frmTestLabs.Show
End Sub

Private Sub mnuLogoff_Click()

    If MsgBox("Are You Sure To Log Off ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Unload MainForm
    LoggingOff = True
    Call Main
    
End Sub

Private Sub mnuLotProc_Click()
    frmInitialLotProcess.Show
End Sub

Private Sub mnuMachSS_Click()
    Call ShowStockStatus(2)
End Sub

Private Sub mnuMake_Click()

    If mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.MakeInactive
    Else
        Call ContractorList.MakeInactive
    End If
    
End Sub

Private Sub mnuMatList_Click()
    frmMaterial.Show
End Sub

Private Sub mnuNewOrd_Click()
    frmFNewCustOrder.Show
End Sub

Private Sub mnuNewProforma_Click()
    frmFNewProforma.Show
End Sub

Private Sub mnuNewStockOrder_Click()
    frmFNewStockOrder.Show
End Sub

Private Sub mnuOrdStatus_Click()
    frmFCustOrdStatus.Show
End Sub

Private Sub mnuOrdStatus1_Click()
    frmFCustOrdStatus1.Show
End Sub

Private Sub mnuOrdStatusComp_Click()
    frmFCustOrdStatusComp.Show
End Sub

Private Sub mnuPackSS_Click()
    Call ShowStockStatus(12)
End Sub

Private Sub mnuPassSS_Click()
    Call ShowStockStatus(8)
End Sub

Private Sub mnuPassword_Click()
    
    With frmUserInfo
        .Show_Option = Opt_ChangeUserPassword
        .Tag = CurrentUserName
        .Show 1
    End With
    
End Sub

Private Sub mnuPer_Click()
    frmPercentRanges.Show
End Sub

Private Sub mnuPoliSS_Click()
    Call ShowStockStatus(7)
End Sub

Private Sub mnuPOrderList_Click()
    frmRMPOrdList.Show
End Sub

Private Sub mnuPORESS_Click()
    Call ShowStockStatus(10)
End Sub

Private Sub mnuPORpt_Click()
    On Error GoTo err
    
    Dim f As New frmPrint
    Load f
    With f
        Set .Rpt = New rptPOrder
        .cmbCust.ClearVals
        '.cmbCust.AddVals con, "Distinct OrderNo", "VFItemsQtyToPlan", "OrderNo", " Where RemQty>0"
        .cmbCust.AddItem "<All Venders>", "0"
        .cmbCust.AddVals con, "VendID1", "Venders", "VendID"
        .SelFor = "{VPOrdersrpt.RecieptID}='"
        .JobCondField = " VendID "
        .JobField = " Distinct RecieptID "
        .JobIDField = " RecieptID "
        .JobTable = " VPOrdersrpt "
        .TheEnd = "'"
        .Show , MainForm
    End With
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPorts_Click()
    frmPorts.Show
End Sub

Private Sub mnuProcess_Click()
    frmProcesses.Show 1
End Sub


Private Sub mnuReqForm_Click()
    'frmRequisition.Show
    frmGMPs.Show
End Sub

Private Sub mnuProdPlan_Click()
    frmProdPlan.Show
End Sub

Private Sub mnuProdPlanrpt_Click()

    On Error GoTo err
    
    Dim f As New frmPrint
    Load f
    With f
        Set .Rpt = New rptProdPlan
        .cmbCust.ClearVals
        '.cmbCust.AddVals con, "Distinct OrderNo", "VFItemsQtyToPlan", "OrderNo", " Where RemQty>0"
        .cmbCust.AddItem "<All Customers>", "0"
        .cmbCust.AddVals con, "CustCode", "ForeignCustomers", "CustCode"
        .SelFor = "{VOrderPlanning.OrderNo}='"
        .JobCondField = " CustCode "
        .JobField = " Distinct OrderNo "
        .JobIDField = " OrderNo "
        .JobTable = " VFItemsQtyToPlan "
        .TheEnd = "'"
        .Show , MainForm
    End With
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub mnuPymtTerms_Click()
    Dim f As New frmMiscS
    Load f
    Call f.ShowForm("PaymentTerms", "Payment Terms", "Payment Terms")
End Sub

Private Sub mnuRcvRepair_Click()
    frmVendRcvRepair.Show
End Sub

Private Sub mnuRcvRepairPoli_Click()
    frmVendRcvRepair.Show
End Sub

Private Sub mnuRcvRepairWastage_Click()
    frmVendRcvRepairWaste.Show
End Sub

Private Sub mnuRcvRpts_Click()
    Dim f As New frmRptIssRcv
    Call f.ShowForIss(True)
End Sub

Private Sub mnuRecRM_Click()
    frmRMRcvItems.Show
End Sub

Private Sub mnuRecSDril_Click()
    Call ShowRcvSForm(3)
End Sub

Private Sub mnuRecSFill_Click()
    Call ShowRcvSForm(4)
End Sub

Private Sub mnuRecSForg_Click()
    Call ShowRcvSForm(1)
End Sub

Private Sub mnuRecSFunc_Click()
    Call ShowRcvSForm(5)
End Sub

Private Sub mnuRecSFunt_Click()
    Call ShowRcvSForm(9)
End Sub

Private Sub mnuRecSHeat_Click()
    Call ShowRcvSForm(6)
End Sub

Private Sub mnuRecSHold_Click()
    Call ShowRcvSForm(13)
End Sub

Private Sub mnuRecSMach_Click()
    Call ShowRcvSForm(2)
End Sub

Private Sub mnuRecSPack_Click()
    Call ShowRcvSForm(12)
End Sub

Private Sub mnuRecSPass_Click()
    Call ShowRcvSForm(8)
End Sub

Private Sub mnuRecSPoli_Click()
    Call ShowRcvSForm(7)
End Sub

Private Sub mnuRecSPoRe_Click()
    Call ShowRcvSForm(10)
End Sub

Private Sub mnuRecSUltr_Click()
    Call ShowRcvSForm(11)
End Sub

Private Sub mnuRecVDril_Click()
    Call ShowRcvVForm(3)
End Sub

Private Sub mnuRecVFill_Click()
    Call ShowRcvVForm(4)
End Sub

Private Sub mnuRecVForg_Click()
    Call ShowRcvVForm(1)
End Sub

Private Sub mnuRecVFunc_Click()
    Call ShowRcvVForm(5)
End Sub

Private Sub mnuRecVFunt_Click()
    Call ShowRcvVForm(9)
End Sub

Private Sub mnuRecVHeat_Click()
    Call ShowRcvVForm(6)
End Sub

Private Sub mnuRecVHold_Click()
    Call ShowRcvVForm(13)
End Sub

Private Sub mnuRecVMach_Click()
    Call ShowRcvVForm(2)
End Sub

Private Sub mnuRecVPack_Click()
    Call ShowRcvVForm(12)
End Sub

Private Sub mnuRecVPass_Click()
    Call ShowRcvVForm(8)
End Sub

Private Sub mnuRecVPoli_Click()
    Call ShowRcvVForm(7)
End Sub

Private Sub mnuRecVPoRe_Click()
    Call ShowRcvVForm(10)
End Sub

Private Sub mnuRecVUltr_Click()
    Call ShowRcvVForm(11)
End Sub

Private Sub mnuRecWastDril_Click()
    Call ShowWastRcvForm(3)
End Sub

Private Sub mnuRecWastFill_Click()
    Call ShowWastRcvForm(4)
End Sub

Private Sub mnuRecWastForg_Click()
    Call ShowWastRcvForm(1)
End Sub

Private Sub mnuRecWastFunc_Click()
    Call ShowWastRcvForm(5)
End Sub

Private Sub mnuRecWastFunt_Click()
    Call ShowWastRcvForm(9)
End Sub

Private Sub mnuRecWastHeat_Click()
    Call ShowWastRcvForm(6)
End Sub

Private Sub mnuRecWastHold_Click()
    Call ShowWastRcvForm(13)
End Sub

Private Sub mnuRecWastMach_Click()
    Call ShowWastRcvForm(2)
End Sub

Private Sub mnuRecWastPack_Click()
    Call ShowWastRcvForm(12)
End Sub

Private Sub mnuRecWastPass_Click()
    Call ShowWastRcvForm(8)
End Sub

Private Sub mnuRecWastPoli_Click()
    Call ShowWastRcvForm(7)
End Sub

Private Sub mnuRecWastPoRe_Click()
    Call ShowWastRcvForm(10)
End Sub

Private Sub mnuRecWastUltr_Click()
    Call ShowWastRcvForm(11)
End Sub

Private Sub mnuReq_Click()
    'frmRequisition.Show
    frmGMPs.Show
End Sub

Private Sub mnuReqRpt_Click()

    On Error GoTo err
    
    Dim f As New frmPrint
    Load f
    With f
        Set .Rpt = New rptRequsition
        .cmbCust.ClearVals
        '.cmbCust.AddVals con, "Distinct OrderNo", "VFItemsQtyToPlan", "OrderNo", " Where RemQty>0"
        .cmbCust.AddItem "<All Processes>", "0"
        .cmbCust.AddVals con, "Description", "Processes", "ProcessID"
        .SelFor = "{VRequsitionrpt.RecieptID}='"
        .JobCondField = " ProcessID "
        .JobField = " Distinct RecieptID "
        .JobIDField = " RecieptID "
        .JobTable = " VRequsitionrpt "
        .TheEnd = "'"
        .Show , MainForm
    End With
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub mnuRM_Click()
    frmRM.Show
End Sub

Private Sub mnuRMAssItems_Click()
    frmRMVendAssItems.Show
End Sub

Private Sub mnuRMGroups_Click()
    Dim f As New frmItemGroups
    Load f
    Call f.ShowForm("RMGroups", "RM Groups", "RM Name")
End Sub

Private Sub mnuRMListrpt_Click()
    Call ShowRMListrpt
End Sub

Private Sub mnuRMPO_Click()
    frmRMPurchOrd.Show
End Sub

Private Sub mnuRMVendList_Click()
    frmRMVenderList.Show
End Sub

Private Sub mnuRMVendListrpt_Click()
    Call ShowRMVendListrpt
End Sub

Private Sub mnuShowDept_Click()
    If mnupopdept.Tag = "EmpDept" Then
        Call DeptList.ShowInactive(mnuShowDept)
    Else
        Call ContDeptList.ShowInactive(mnuShowDept)
    End If
End Sub

Private Sub mnuShowInactive_Click()
    If mnupopemp.Tag = "From Emp" Then
        Call EmployeeList.ShowInactive(mnuShowInactive)
    Else
        Call ContractorList.ShowInactive(mnuShowInactive)
    End If
End Sub

Private Sub mnuSteel_Click()
    frmSteelList.Show
End Sub

Private Sub mnuSteelVends_Click()
    frmSteelVends.Show
End Sub

Private Sub mnuSubGeneral_Click(Index As Integer)

    Select Case Index
        Case 1
            frmInitialLotProcess.Show
        Case 3
            frmPercentRanges.Show
        Case 5
            frmComponents.Show 1
        Case 7
            frmProcesses.Show 1
        Case 8
            frmFDAProcesses.Show 1
        Case 10
            frmSteelVends.Show
        Case 11
            frmTestLabs.Show
        Case 13
            frmUnits.Show
        Case 14
            frmCompositions.Show
        Case 16
            frmPorts.Show
        Case 19
            Call mnuPymtTerms_Click
    End Select
    
End Sub

Private Sub mnuSubOptions_Click(Index As Integer)

    Select Case Index
        Case 1
            Call mnuCalc_Click
        Case 2
            frmUsers.Show 1
        Case 4
            Call mnuPassword_Click
        Case 5
            Call mnuLogoff_Click
        Case 7
            Call mnuExit_Click
    End Select
    
End Sub

Private Sub mnuTaxRanges_Click()
    Load TaxRanges
    'Holidays.Move 3000, 500
    With TaxRanges
        .Move (Me.Width - .Width) / 2, Me.Top + 500
        .Show
    End With
End Sub

Private Sub mnuUltrSS_Click()
    Call ShowStockStatus(11)
End Sub

Private Sub mnuUnits_Click()
    frmUnits.Show
End Sub

Private Sub mnuUserManager_Click()
    frmUsers.Show 1
End Sub

Public Sub ApplySecuritySettings(UserID As Long)
   
    On Error GoTo err
    
    'General Settings...
    mnuSubGeneral(1).Enabled = UserHasAccess("SetInitLotProc", UserID)
    mnuSubGeneral(3).Enabled = UserHasAccess("PercentRanges")
    mnuSubGeneral(5).Enabled = UserHasAccess("Components")
    mnuSubGeneral(7).Enabled = UserHasAccess("Processes")
    mnuSubGeneral(8).Enabled = UserHasAccess("FDAProcesses")
    mnuSubGeneral(10).Enabled = UserHasAccess("SteelVenders")
    mnuSubGeneral(11).Enabled = UserHasAccess("TestLabs")
    mnuSubGeneral(13).Enabled = UserHasAccess("MeasuringUnits")
    mnuSubGeneral(14).Enabled = UserHasAccess("Compositions")
    mnuSubGeneral(16).Enabled = UserHasAccess("CitiesNPorts")
    mnuSubGeneral(18).Enabled = UserHasAccess("PymtTerms")
    
    'Venders
    'mnuVendList.Enabled = UserHasAccess("VenderList")
    'mnuVendAssItems.Enabled = UserHasAccess("VendAssItems")
    mnuVendPymt.Enabled = UserHasAccess("VendAdvPymt")
    mnuVendPaySheet.Enabled = UserHasAccess("VendPymtSheet")
    mnuVendAdvShort.Enabled = UserHasAccess("NewShortTerm")
    mnuVendAdvLong.Enabled = UserHasAccess("NewLongTerm")
    mnuVendShortLoanLedger.Enabled = UserHasAccess("ShortTermLedger")
    mnuVendLongLoanLedger.Enabled = UserHasAccess("LongTermLedger")
    'mnuVendLedger.Enabled = UserHasAccess("VendLedger")
    
    'Raw Material
    mnuRMGroups.Enabled = UserHasAccess("RMGroups")
    mnuRM.Enabled = UserHasAccess("RawMaterial")
    'mnuRMVendList.Enabled = UserHasAccess("RMVendList")
    'mnuRMAssItems.Enabled = UserHasAccess("RMVAssItems")
    mnuRMPO.Enabled = UserHasAccess("RMVNewPO")
    mnuPOrderList.Enabled = UserHasAccess("RMVPOList")
    mnuRecRM.Enabled = UserHasAccess("RMRcvRM")
    mnuIssRM.Enabled = UserHasAccess("RMIssRM")
    
    'HICO Resources
    mnuHRGroups.Enabled = UserHasAccess("HICOResGroups")
    'mnuHICOResources.Enabled = UserHasAccess("HICOResources")
    'mnuVendAssR.Enabled = UserHasAccess("AssHICORes")
        
    'Company
    'mnuCompDeatil.Enabled = UserHasAccess("CompanyDetail")
    'mnuCompCatal.Enabled = UserHasAccess("CompanyCatalog")
    mnuItmGrps.Enabled = UserHasAccess("ItemGroups")
    mnuItemsFinQuality.Enabled = UserHasAccess("ItemFinQuality")
    mnuItemTypes.Enabled = UserHasAccess("ItemTypes")
    'mnuForeignCatal.Enabled = UserHasAccess("ForeignCatalog")
    'mnuCustList.Enabled = UserHasAccess("CustomerList")
    'mnuCustCatal.Enabled = UserHasAccess("CustomerCatalog")
    mnuNewOrd.Enabled = UserHasAccess("CustNewOrder")
    mnuNewStockOrder.Enabled = UserHasAccess("StockNewOrder")
    mnuCustOrdList.Enabled = UserHasAccess("CustOrdList")
    mnuOrdStatus1.Enabled = UserHasAccess("CustOrdStatus")
    mnuNewProforma.Enabled = UserHasAccess("CustNewProforma")
    mnuCustProformaList.Enabled = UserHasAccess("CustProList")
        
    mnuReq.Enabled = UserHasAccess("Requsition")
    mnuProdPlan.Enabled = UserHasAccess("AddProdPlan")
    mnuAuth.Enabled = UserHasAccess("Authorize")
    
    'Reports
    
    'Users
    mnuSubOptions(2).Enabled = UserHasAccess("UserManagement")
    mnuSubOptions(4).Enabled = UserHasAccess("ChangePassword")
    
    
    HicoVisible = UserHasAccess("HicoVisible")
    'If UserID = 1 Then mnuAuth.Enabled = True Else mnuAuth.Enabled = False
    
    If UserID <> 1 Then Call ApplySecurityForIssRcv(UserID)
    '
    'With Toolbar1.Buttons
    '   '.Item("TakeAtt").Enabled = mnutake.Enabled
    '   '.Item("NewShort").Enabled = mnuShort.Enabled
    '   '.Item("NewLong").Enabled = mnuLong.Enabled
    '   '.Item("EmpLedger").Enabled = mnuEmpLedger.Enabled
    '   '.Item("SalSheet").Enabled = mnuSalSheet.Enabled
    'End With
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub ApplySecurityForIssRcv(UserID As Long)
    
    For i = 1 To mnuIssRcvProcs.Count
        mnuIssRcvProcs(i).Enabled = False
    Next i
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select ProcessID From UserAssProcesses Where UserID=" & UserID, con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            mnuIssRcvProcs(![ProcessID]).Enabled = True
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub
Public Sub mnuVendAdvLong_Click()

    With frmVendNewAdv
        .LongTerm = True
        .Show 1
    End With
    
End Sub

Public Sub mnuVendAdvShort_Click()
    
    With frmVendNewAdv
        .LongTerm = False
        .Show 1
    End With
    
End Sub

Private Sub mnuVendAssItems_Click()
    frmVendAssItems.Show
End Sub

Private Sub mnuVendAssR_Click()
    frmVendAssHR1.Show
End Sub

Private Sub mnuVendIssItems_Click()
    frmVendIssItems.Show
End Sub

Private Sub mnuVendIssItemsByItems_Click()
    frmVendIssItemsByItem.Show
End Sub

Private Sub mnuVendLedger_Click()
    frmVendLedger.Show
End Sub

Private Sub mnuVendList_Click()
    frmVenderList.Show
End Sub

Private Sub mnuVendListrpt_Click()
    Call ShowVendListrpt
End Sub

Private Sub mnuVendLongLoanLedger_Click()
    frmVendLoanLedger.Show
End Sub

Private Sub mnuVendPaySheet_Click()
    frmVendSS.Show    'frmVendPaySheet.Show
End Sub

Private Sub mnuVendPymt_Click()
    frmVendAdvSalary.Show
End Sub

Private Sub mnuVendRcvItems_Click()
    frmVendRcvItems.Show
End Sub

Private Sub mnuVendRcvItemsFromSupervisor_Click()
    frmVendRcvItems1.Show
End Sub

Private Sub mnuVendRcvWst_Click()
    frmVendRcvWastage.Show
End Sub

Private Sub mnuVendShortLoanLedger_Click()
    frmVendShortAdvLedger.Show
End Sub

'-----------------------------------------------------------
'Code Of Pop Menu Ports.
Private Sub NewCity_Click()
    Call frmPorts.NewCity_Click
End Sub


Private Sub Rename_Click()
    Call frmPorts.Rename_Click
End Sub

Private Sub del_Click()
    Call frmPorts.del_Click
End Sub

Private Sub PicBusPart_Resize()

    Dim Temp As Long
    PicBusPart.Left = 0
    PicBusPart.Width = 1950
    
    'If PicBusPart.Height = 1995 Then 'Means Its Open
    Temp = Round(PicBusPart.Height / 350)
     'MsgBox "Total Buttons=" & Round(PicBusPart.Height / 350)
    'End If
    
    VSBusPart.Move PicBusPart.Width - VSBusPart.Width - 50, VSBusPart.Top, VSBusPart.Width, PicBusPart.Height - 15
    VSBusPart.Max = (BusPartToolBar.Buttons.Count - Temp) + 1
    'VSBusPart.Left = PicBusPart.Width - VSBusPart.Width - 50
    
    'PicBusPart.Height = 1665
End Sub


Private Sub PicTitle_Click()
    For i = 1 To Toolbar.Count
        PicToolBar(i).SetFocus
    Next i
End Sub


Private Sub PicToolBar_Paint(Index As Integer)
    'If Index = 1 Then Stop
    'If PicToolBar(Index).Height > 15 Then
    '    ToolBar(Index).Visible = False
    '    ToolBar(Index).Visible = True
    'End If
    
End Sub



Private Sub PopSubCust_Click(Index As Integer)
    Select Case Index
        Case 0
            frmFCustomer.Show 1
        Case 1
            frmFCustomerList.Show
        Case 3
            frmCustCatal.Show
        Case 4
    End Select
End Sub

Private Sub PopSubProdVend_Click(Index As Integer)
    
    On Error GoTo err
    
    Select Case Index
        Case 0
            Dim f As New frmVender
            With f
                .Add = True
                .Show 1
            End With
        Case 1
            frmVenderList.Show
        Case 3
            frmVendAssItems.Show
        Case 4
            frmVendLedger.Show
            'Call showform(frmVendLedger)
    End Select
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub PopSubRMVend_Click(Index As Integer)
    Select Case Index
        Case 0
            Dim f As New frmRMVender
            With f
                .Add = True
                .Show 1
            End With
        Case 1
            frmRMVenderList.Show
        Case 3
            frmRMVendAssItems.Show
        Case 4
    End Select
End Sub

Private Sub ProdToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 4 Then
        Me.PopUpMenu mnuProdPhases, 512, CoolBar1.Width
    End If
End Sub


'-----------------------------------------------------------
Private Sub Timer1_Timer()
    '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
    '*************************************************
        'This Is Special Code...
        'It Will Check If User Is Back From Some Other Form.
        'If So Then It Will Do Something To Get The Buttons Back
        'Of The Toolbar.
    '*************************************************
    Dim a As String
    Dim Temp As Long
    a = Space$(255)
    Temp = GetClassName(GetForegroundWindow, a, 255)
    a = Left(a, Temp)
    
    TempForeWin = GetForegroundWindow
    If TempForeWin <> Me.hwnd And a <> "msvb_lib_monthview" Then
        'Now Check If Screen.ActiveForm.Hwnd Not In The Loaded forms Collection...
        
        On Error Resume Next
        If Not InLoadedForms(Screen.ActiveForm.hwnd) Then
            If AwayTime = "12:00:00 AM" Then
                AwayTime = ServerTime
            End If
            ForeWin = TempForeWin
        End If
        On Error GoTo 0
    Else
        If ForeWin <> Me.hwnd Then
            
            ForeWin = Me.hwnd
            If Not OurOwnForm Then
                'FrmUnlock.Show 1 'Ask For Pass
            End If
            
            Call ResetToolbar
            OurOwnForm = False
            
            'Sleep (5000)
            'Call ABC
            Ready = True
        End If
    End If
    '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
    
End Sub
Private Function InLoadedForms(lHwnd As Long) As Boolean

    Dim f As Form
    InLoadedForms = False
    'Debug.Print lHwnd
    For Each f In Forms
        If f.hwnd = lHwnd Then
            InLoadedForms = True
            Exit Function
        End If
    Next
    
End Function
Private Sub ResetToolbar()
    
    
    
End Sub
Private Sub PicToolBar_Resize(Index As Integer)

    Dim Temp As Long
    PicToolBar(Index).Left = 0
    PicToolBar(Index).Width = 2310
    Toolbar(Index).Move 0, 0, PicToolBar(Index).Width - SBar(Index).Width - 40
    'If PicBusPart.Height = 1995 Then 'Means Its Open
    Temp = Round(PicToolBar(Index).Height / 350)
     'MsgBox "Total Buttons=" & Round(PicBusPart.Height / 350)
    'End If
    
    SBar(Index).Move PicToolBar(Index).Width - SBar(Index).Width - 50, 10, SBar(Index).Width, PicToolBar(Index).Height - 15
    SBar(Index).Max = (Toolbar(Index).Buttons.Count - Temp) + 1
    
End Sub
Private Sub SBar_Change(Index As Integer)
    Toolbar(Index).Top = -((SBar(Index).value - 1) * 350)
End Sub

Private Sub Timer2_Timer()
    AwaySeconds = AwaySeconds + 1
End Sub

Private Sub Timer3_Timer()
    Timer3.Interval = 1000
    Dim SeverTime1 As Date
    SeverTime1 = con.Execute("Select GetDate()").Fields(0)
    ServerTime = Format(SeverTime1, "hh:mm:ss AMPM")
    If Format(ServerDate, "dd-MMM-yyyy") <> Format(SeverTime1, "dd-MMM-yyyy") Then ServerDate = Format(SeverTime1, "dd-MMM-yyyy")
End Sub

Private Sub ToolBar_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    
    On Error GoTo err
    
    Dim MenuTop As Long
    'MenuTop = Button.Top + CoolBar1.Top + PicToolBar(Index).Top - ((SBar(Index).value - 1) * 350)
    MenuTop = Button.Top + CoolBar1.Top + PicToolBar(Index).Top + Toolbar(Index).Top
    Select Case Index
        Case 1
            Call company(Button, MenuTop)
        Case 2
            'Call Policies(Button)
        Case 3
            Call BusPart(Button, MenuTop)
        Case 4
            Call Orders(Button, MenuTop)
        Case 7
            Call Production(Button, MenuTop)
        Case 8
            Call Payroll(Button, MenuTop)
        Case 9
            Call Accounts(Button, MenuTop)
        Case 10
            Call Export(Button, MenuTop)
    End Select
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub company(ByVal Button As MSComctlLib.Button, MenuTop As Long)

    'Error Traping Of Toolbar_ButtonClick Will be Used.
    
    Dim strRet As String
    
    Select Case Button.key
        Case "CompDetail"
            frmCompany.Show
        Case "CompCatalog"
            'frmItems.Show
            'Exit Sub
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "&Group List", "-", "&Item List", "Add &New Item")
            Select Case strRet
                Case "&Group List"
                    Dim f As New frmItemGroups
                    Load f
                    Call f.ShowForm("ItemGroups", "Item Groups", "Group Name")
                Case "&Item List"
                    frmItems.Show
                    
                Case "Add &New Item"
                    With frmNewItem
                        Load frmNewItem
                        .Show 1
                    End With
            End Select
        Case "RawMaterial"
        
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "Raw Material &Groups", "&Raw Material", "-", "&New Purchase Order", "&Purchase Order List", "-", "&Receive Raw Material", "&Issue Raw Material")
            Select Case strRet
                Case "Raw Material &Groups"
                    
                    Load frmItemGroups
                    Call frmItemGroups.ShowForm("RMGroups", "RM Groups", "RM Name")
                Case "&Raw Material"
                    frmRM.Show
                Case "&New Purchase Order"
                    frmRMPurchOrd.Show
                Case "&Purchase Order List"
                    frmRMPOrdList.Show
                Case "&Receive Raw Material"
                    frmRMRcvItems.Show
                Case "&Issue Raw Material"
                    frmRMIss.Show
            End Select
        
        Case "SteelList"
            frmSteelList.Show
        Case "ForCatalog"
            frmForeignCatalog.Show 1
        Case "CompDept"
                
        Case "CompResources"
            frmHICOR.Show
            
        Case "Banks"
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "&Bank List", "&Add New Bank", "-", "Add &New Cheque Book")
            
        Case "Processes"
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "&Process List", "-", "&FDA Process List")
            Select Case strRet
                Case "&Process List"
                    frmProcesses.Show 1
                Case "&FDA Process List"
                    frmFDAProcesses.Show 1
            End Select
            
        Case "GeneralOptions"
            'strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "&Process List", "-", "&FDA Process List")
            Me.PopUpMenu mnuGeneral, 512, CoolBar1.Width, MenuTop
            
    End Select
    
    
    
End Sub
Private Sub BusPart(ByVal Button As MSComctlLib.Button, MenuTop As Long)

    'Error Traping Of Toolbar_ButtonClick Will be Used.
    Dim strRet As String
    Select Case Button.key
        Case "FCustomers"
        
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "&Customer List", "&Add New Customer", "-", "Customer Catalog", "-", "Customer Ledger", "-", "Customer Mailing Labels")
            Select Case strRet
                Case "&Customer List"
                    frmFCustomerList.Show
                Case "&Add New Customer"
                    frmFCustomer.Show 1
                Case "Customer Ledger"
                
                Case "Customer Catalog"
                    frmCustCatal.Show
                Case "Customer Mailing Labels"
            End Select
            
        Case "Contractors"
            Me.PopUpMenu mnuVenders, 512, CoolBar1.Width, MenuTop
            Exit Sub
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, Button.Description & " &List", "-", "&New " & Button.Description, "-", Button.Description & " &Assigned Items", "-", Button.Description & " Ledger", "-", "Vender Advance Payment", "-", "Vender Payment Sheet")
            
            If Right(strRet, 4) = "List" Then
                Call ShowBusPartList(Button.Description)
            ElseIf Left(strRet, 4) = "&New" Then
                Call AddNewBusPart(Button.Description)
            ElseIf Right(strRet, 5) = "Items" Then
                Call ShowBusPartAssItems(Button.Description)
            ElseIf Right(strRet, 6) = "Ledger" Then
                Call ShowBusPartLedger(Button.Description)
            ElseIf strRet = "Vender Advance Payment" Then
                frmVendAdvSalary.Show
            ElseIf strRet = "Vender Payment Sheet" Then
                frmVendSS.Show
            End If
            
        Case "RMVenders"
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, Button.Description & " &List", "-", "&New " & Button.Description, "-", Button.Description & " &Assigned Items", "-", Button.Description & " Ledger")
            If Right(strRet, 4) = "List" Then
                Call ShowBusPartList(Button.Description)
            ElseIf Left(strRet, 4) = "&New" Then
                Call AddNewBusPart(Button.Description)
            ElseIf Right(strRet, 5) = "Items" Then
                Call ShowBusPartAssItems(Button.Description)
            ElseIf Right(strRet, 6) = "Ledger" Then
                Call ShowBusPartLedger(Button.Description)
            End If
        Case Else
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, Button.Description & " &List", "-", "&New " & Button.Description)
            If Right(strRet, 4) = "List" Then
                Call ShowBusPartList(Button.Description)
            ElseIf Left(strRet, 4) = "&New" Then
                Call AddNewBusPart(Button.Description)
            End If
        End Select
        
    
    
End Sub
Private Sub Orders(ByVal Button As MSComctlLib.Button, MenuTop As Long)
    'Error Traping Of Toolbar_ButtonClick Will be Used.
    'Toolbar Index 4.
    
    Select Case Button.key
        Case "NewOrder"
            frmFNewCustOrder.Show
        Case "NewStockOrder"
            frmFNewStockOrder.Show
        Case "Batches"
            frmBatches.Show 1
        Case "OrderList"
            frmfCustOrderList.Show
        Case "OrderStatus"
            frmFCustOrdStatus1.Show
        Case "OrderPlanning"
            frmProdPlan.Show
    End Select
    
End Sub
Private Sub Production(ByVal Button As MSComctlLib.Button, MenuTop As Long)
    'Error Traping Of Toolbar_ButtonClick Will be Used.
    
    
    Dim strRet As String
    
    Select Case Button.key
        Case "Requsition"
            frmGMPs.Show
        Case "PPlan"
            frmProdPlan.Show
        Case "Authorise"
            frmAuthorizeIssRcv.Show
        Case "Processes"
            'mnuReq.Visible = False
            'mnuProdPlan.Visible = False
            'mnuAuth.Visible = False
            Me.PopUpMenu mnuProdPhases, 512, CoolBar1.Width
    End Select
        
    
End Sub
Private Sub Payroll(ByVal Button As MSComctlLib.Button, MenuTop As Long)
    'Error Traping Of Toolbar_ButtonClick Will be Used.
    
    
    Dim strRet As String
    
    Select Case Button.key
    
        Case "Departments"
            DeptList.Show
        Case "Employees"
            EmployeeList.Show
        Case "Attendance"
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "&Take Attendance", "Attendance &Sheet", "-", "&Over Time Record", "-", "Employee &Penalty Days")
            Select Case strRet
                Case "&Take Attendance"
                    Load Attend
                    Attend.Show
                Case "Attendance &Sheet"
                    AttSheet.Show
                Case "&Over Time Record"
                    OverTime.Show
                Case "Employee &Penalty Days"
                    EmpAbsents.Show
            End Select
        Case "AbsentSheet"
            AbsentSheet.Show
        Case "OverTime"
            OverTime.Show
        Case "ShortTerm"
            AdvanceShort.Show
        Case "LongTerm"
            LoanLedger.Show
        Case "Outstanding"
            PrintOutStanding.Show 1
        Case "AdvSalSheet"
            AdvSalary.Show
        Case "Deductions"
            
        Case "Benefits"
            
        Case "Bonuses"
            BonusSheet.Show
        Case "SalSheet"
            SalarySheet.Show
        Case "PaySlips"
            Load Payslip
            'Holidays.Move 3000, 500
            With Payslip
                .Move (Me.Width - .Width) / 2, Me.Top + 500
                .Show
            End With
        Case "EmpLedger"
            SalaryLedger.Show
        Case "SalHistory"
            SalaryHistory.Show
        Case "JoinedDis"
            PrintEmpJoinLeave.Show 1
        Case "Assets"
            
        Case "OldAge"
            OldAge.Show
        Case "SSecurity"
            SocialSecurity.Show
    End Select
        
    
End Sub
Private Sub Accounts(ByVal Button As MSComctlLib.Button, MenuTop As Long)
    'Error Traping Of Toolbar_ButtonClick Will be Used.
    
    
    Dim strRet As String
    Dim f As Form
    
    Select Case Button.key
    
        Case "CashnBank"
            frmCashStatus.Show
        Case "DueCheques"
            frmChqs.Show
        Case "DailyTrans"
        
            Set f = New frmTransaction
            With f
               .Opts(0).value = True
               .Show
            End With

        Case "CashBook"
            'For Helicon and Bhinder Show Cash in hand Ledger
            
            'If LCase(company) Like "helicon*" Or LCase(company) Like "bhinder*" Then
            'If LCase(company) Like "bhinder*" Then
            '    With frmLedger
            '        .Show
            '        .cmbAccNo = CashInHand
            '        .DTFrom = Date
            '        .DTTo = Date
            '        .cmdRefresh_Click
            '        .SetFocus
            '    End With
               
            'Else
                Set f = New frmTransaction
                With f
                    .Opts(1).value = True
                    .Show
                End With
            'End If

        Case "Banking"
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "&New Bank Account", "Bank Accounts && Cheque Book &Detail", "-", "Bank && Branch &Profile")
            
            Select Case strRet
            
                Case "&New Bank Account"
                    With frmNewAccount
                        .chkBankAcc = vbChecked
                        .Show 1
                    End With
                Case "Bank Accounts && Cheque Book &Detail"
                    frmChqBooks.Show
                Case "Bank && Branch &Profile"
                    frmBankList.Show
            End Select
            
        Case "FundsTransfer"
            frmTransferFunds.Show 1
        Case "Receipt"
            frmCashRV.Show
        Case "Payment"
            frmCashPV.Show
        Case "Journal"
            frmJV.Show
        Case "Expense"
            frmExpenses.Show
        Case "ClearAccrued"
            frmAcrExpenses.Show
        Case "Chart"
            Load frmChartOfAccounts
            frmChartOfAccounts.Show
        Case "AccRec"
            Set f = New frmAccPayRcv
            With f
               .Tag = ACCRECNO
               .Show
            End With
        Case "AccRecAging"
        
        Case "AccPay"
            Set f = New frmAccPayRcv
            With f
               .Tag = ACCPAYNO
               .Show
            End With

        Case "AccPayAging"
        
        Case "MonTrans"
            frmComparison.Show
        Case "TrailBalance"
        
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "Trial Balance ( Summary )", "-", "Trial Balance ( Detail )")
            
            Select Case strRet
            
                Case "Trial Balance ( Summary )"
                    Set f = New frmTrialBalance
                    f.Show
                Case "Trial Balance ( Detail )"
                    Set f = New frmTrialBalance
                    f.chkDetail.value = vbChecked
                    f.Show
            End Select
            
        Case "ProfitnLoss"
        
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "Profit && Loss ( Standard )", "-", "Profit && Loss ( Summary )", "-", "Profit && Loss ( Detail )")
            
            Select Case strRet
            
                Case "Profit && Loss ( Standard )"
                
                    Set f = New frmPLS
                    With f
                        .Summary = False
                        .Standered = True
                        .Show
                    End With

                Case "Profit && Loss ( Summary )"
                
                    Set f = New frmPLS
                    With f
                        .Summary = True
                        .Standered = False
                        .Show
                    End With
                    
                Case "Profit && Loss ( Detail )"
                
                    Set f = New frmPLS
                    With f
                        .Summary = False
                        .Show
                    End With
        
            End Select
            
        Case "BalSheet"
            frmBalSheet.Show
            
    End Select
    
        
    
End Sub
Private Sub Export(ByVal Button As MSComctlLib.Button, MenuTop As Long)
    'Error Traping Of Toolbar_ButtonClick Will be Used.
    
    
    Dim strRet As String
    
    Select Case Button.key
        Case "Agents"
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "&Shipping Agents", "-", "&Insurance Agents")
            Select Case strRet
                Case "&Shipping Agents"
                    frmShippingAgents.Show
                Case "&Insurance Agents"
                    frmInsuranceAgents.Show
            End Select
        Case "PInvoice"
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "&Proforma Invoice List", "-", "&New Proforma Invoice")
            Select Case strRet
                Case "&Proforma Invoice List"
                    frmFProformaList.Show
                Case "&New Proforma Invoice"
                    'frmFNewProforma.Show
                    Load frmFNewProforma
                    frmFNewProforma.Show
                    frmFNewProforma.AddNewProforma
            End Select
        Case "CustomDocs"
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "Covering &Letter", "-", "&Custom Invoice", "-", "&Packing List", "-", "&Shipping Instruction", "&E Form")
            
            Select Case strRet
                Case "Covering &Letter"
                    frmCustomCovering.Show
                Case "&Custom Invoice"
                    frmCustomInvoice.Show
                Case "&Packing List"
                    frmCustomPList.Show
                Case "&E Form"
                    frmEForm.Show
                Case "&Shipping Instruction"
                    frmShippingInstructions.Show
            End Select
            
        Case "ComDocs"
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "Covering &Letter", "-", "&Commercial Invoice", "-", "&Packing List", "-", "&Inspection Certificate", "-", "&GSP")
            
            Select Case strRet
                Case "Covering &Letter"
                    frmCommercialCovering.Show
                Case "&Commercial Invoice"
                    'frmComInvList.Show
                    frmCommercialInvoice.Show
                Case "&Packing List"
                    frmCustomPList.Show
                Case "&Inspection Certificate"
                    frmInspection.Show
                
                Case "&GSP"
                    frmGSP.Show
                
            End Select
            
        Case "DHR"
            frmDHR.Show
            
        Case "RebateDocs"
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "Rebate &Documents", "-", "&Payment && Rebate Status")
            Select Case strRet
                Case "Rebate &Documents"
                    frmRebate.Show
                Case "&Payment && Rebate Status"
                    frmRebateStatus.Show 1
            End Select
            
        Case "Statistics"
        
            strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "&Total Export", "-", "&Payment Ledger")
            Select Case strRet
                Case "&Total Export"
                    frmTotalExport.Show
                Case "&Payment Ledger"
            End Select
            
        Case "EEStatement"
            
    End Select
    
        
    
End Sub



