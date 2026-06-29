VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "CTW Accounting System"
   ClientHeight    =   6330
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9495
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3195
      Top             =   2910
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3240
      Top             =   1650
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select Database File"
      FileName        =   "Accounts.CTW"
      Filter          =   "CTW Accounts Files|*.CTW"
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   1005
      ButtonWidth     =   1270
      ButtonHeight    =   953
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Accnt"
            Key             =   "Accounts"
            Object.ToolTipText     =   "Chart Of Accounts  [F2]"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ChartOfAccounts"
                  Text            =   "&Chart Of Accounts"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AccPay"
                  Text            =   "Accounts &Payables"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AccRec"
                  Text            =   "Accounts &Receivables"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cust"
            Key             =   "Customers"
            Object.ToolTipText     =   "Customers List  [F3]"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "NewCustomer"
                  Text            =   "New Customer"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "CustomerList"
                  Text            =   "Customer List"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Vends"
            Key             =   "Venders"
            Object.ToolTipText     =   "Venders List  [F4]"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "NewVender"
                  Text            =   "New Vender"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "VenderList"
                  Text            =   "Vender &List"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Banks"
            Key             =   "Banks"
            Object.ToolTipText     =   "Bank & Bank Accounts Details  [F5]"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "NewBank"
                  Text            =   "New Bank Account"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BankAccs"
                  Text            =   "Bank Accounts Detail"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "BankList"
                  Text            =   "Banks && Branch &Profile"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rcpt"
            Key             =   "RV"
            Object.ToolTipText     =   "Receipt Transaction  [F6]"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Pymt"
            Key             =   "PV"
            Object.ToolTipText     =   "Payment Transaction  [F7]"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "JV"
            Key             =   "JV"
            Object.ToolTipText     =   "Journal Transaction  [F8]"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sales"
            Key             =   "SJV"
            Object.ToolTipText     =   "Sale Transaction  [F9]"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SJV"
                  Text            =   "Sales Voucher"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SRJV"
                  Text            =   "Sale Return Voucher"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Prchs"
            Key             =   "PJV"
            Object.ToolTipText     =   "Purchase Transaction  [F11]"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PJV"
                  Text            =   "Purchase Voucher"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PRJV"
                  Text            =   "Purchase Return Vouhcer"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Emps"
            Key             =   "EmpAdv"
            Object.ToolTipText     =   "Employees List  [Ctrl+E]"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   6
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "EmpAdv"
                  Text            =   "Short Term Advances"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "EmpLongAdv"
                  Text            =   "Long Term Advances"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "EmpSal"
                  Text            =   "Employee Salaries"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "EmpList"
                  Text            =   "Employees List"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "NewEmp"
                  Text            =   "New Employee"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Expense"
            Key             =   "Expenses"
            Object.ToolTipText     =   "Expense Transaction  [F12]"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ExpVchr"
                  Text            =   "Expense Voucher"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AcrExp"
                  Text            =   "Clear Accured Expense Voucher"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "MonthAcr"
                  Text            =   "Monthly Accured/Prepaid Expenses"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ExpAcc"
                  Text            =   "Define Expense Account"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ledger"
            Key             =   "Ledger"
            Object.ToolTipText     =   "View Detailed Account Ledger  [Ctrl+L]"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Ledger"
                  Text            =   "Detailed Ledger"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "MonLgr"
                  Text            =   "Comparison Ledger"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "TransRpt"
                  Text            =   "Transaction Report"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Window"
            Key             =   "Window"
            Description     =   "`"
            Object.ToolTipText     =   "Toggle Between Active Windows"
         EndProperty
      EndProperty
      Begin ComboList.Usercontrol1 Usercontrol11 
         Height          =   285
         Left            =   8400
         TabIndex        =   1
         Top             =   450
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   -1
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4050
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E42
            Key             =   "Banks"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":134A
            Key             =   "Venders"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1742
            Key             =   "Users"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23F6
            Key             =   "Expenses"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2842
            Key             =   "Ledger"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C96
            Key             =   "Accounts"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":319E
            Key             =   "RV"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":36A6
            Key             =   "PV"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3BAE
            Key             =   "Customers"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":40B6
            Key             =   "PJV"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4502
            Key             =   "EmpAdv"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4A0A
            Key             =   "SJV"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4FEE
            Key             =   "Items"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":54F6
            Key             =   "Window"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":594A
            Key             =   "JV"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5D9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":61F2
            Key             =   "Expense"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":66FA
            Key             =   "Venders1"
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   240
      Top             =   2265
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuPrintOnSave 
         Caption         =   "Print Vouchers When Saved"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuCalculater 
         Caption         =   "Use Calculater"
      End
      Begin VB.Menu dashcalc 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "User Manager"
      End
      Begin VB.Menu dashh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup Data"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   "Restore Data"
      End
      Begin VB.Menu mnudashdb 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangeDBPath 
         Caption         =   "Change Database Path"
      End
      Begin VB.Menu dashdata 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompanyInfo 
         Caption         =   "Company Information"
      End
      Begin VB.Menu mnuRegister 
         Caption         =   "Register CTW Accounting System"
      End
      Begin VB.Menu dash0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChangePassword 
         Caption         =   "Change Password"
      End
      Begin VB.Menu mnuLogOff 
         Caption         =   "Log Off"
      End
      Begin VB.Menu dashexit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuPress 
      Caption         =   "Press"
      Visible         =   0   'False
      Begin VB.Menu mnuRates 
         Caption         =   "Define Rates"
      End
      Begin VB.Menu mnuPapers 
         Caption         =   "Papers Types"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuJobs 
         Caption         =   "Customer Jobs"
      End
      Begin VB.Menu mnuPurchaseOrders 
         Caption         =   "Paper Purchase Orders"
      End
      Begin VB.Menu mnuCanceled 
         Caption         =   "Order Canceled Items"
      End
      Begin VB.Menu dashpress 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOrderS 
         Caption         =   "Receive Orders"
      End
   End
   Begin VB.Menu mnuAccounts 
      Caption         =   "Accounts"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mnuNewAccounts 
         Caption         =   "Define &New Accounts"
      End
      Begin VB.Menu mnuNewExpAcc 
         Caption         =   "Define New &Expense Account"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccPay 
         Caption         =   "Accounts &Payable"
      End
      Begin VB.Menu mnuAccRcv 
         Caption         =   "Accounts &Receivable"
      End
      Begin VB.Menu dashacc 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChaccounts 
         Caption         =   "&Chart Of Accounts"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuStock 
      Caption         =   "Stock"
      Begin VB.Menu mnuItems 
         Caption         =   "Item List"
      End
      Begin VB.Menu dashpch 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPchOrders 
         Caption         =   "Purchase Orders"
      End
      Begin VB.Menu mnuItemRcpt 
         Caption         =   "Receive Order Item"
      End
      Begin VB.Menu mnuPurchase 
         Caption         =   "Item Purchase"
      End
      Begin VB.Menu dashrcpt 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaleOrders 
         Caption         =   "Customer Orders"
      End
      Begin VB.Menu mnuOrderSale 
         Caption         =   "Customer Ordered Sale"
      End
      Begin VB.Menu mnuItemSales 
         Caption         =   "Item Sales"
      End
      Begin VB.Menu dashSRT 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSRT 
         Caption         =   "Sales Return From Customer"
      End
      Begin VB.Menu mnuPRT 
         Caption         =   "Purchase Return To Vender"
      End
      Begin VB.Menu dashsales 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStockAdj 
         Caption         =   "Adjust Quantity In Hand"
      End
      Begin VB.Menu mnuItemLedger 
         Caption         =   "Item Ledger"
      End
   End
   Begin VB.Menu mnuBanks 
      Caption         =   "&Banks"
      Begin VB.Menu mnuNewBankAcc 
         Caption         =   "&New Bank Account"
      End
      Begin VB.Menu mnuChqBooks 
         Caption         =   "Bank Accounts && Cheque Book &Detail"
         Shortcut        =   {F5}
      End
      Begin VB.Menu dash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBankList 
         Caption         =   "Bank and Branch &Profile"
      End
   End
   Begin VB.Menu mnuVouchers 
      Caption         =   "&Vouchers"
      Begin VB.Menu mnuRVoucher 
         Caption         =   "&Receipt Voucher"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuPVoucher 
         Caption         =   "&Payment Voucher"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuJVoucher 
         Caption         =   "&Journal Voucher"
         Shortcut        =   {F8}
      End
      Begin VB.Menu dash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaleJV 
         Caption         =   "&Sale Voucher"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuPurchaseJV 
         Caption         =   "Pur&chase Voucher"
         Shortcut        =   {F11}
      End
      Begin VB.Menu dashSale 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaleReturn 
         Caption         =   "Sale &Return"
      End
      Begin VB.Menu mnuPurchaseReturn 
         Caption         =   "Purchase Re&turn"
      End
      Begin VB.Menu dashexp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExpenses 
         Caption         =   "&Expenses"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuAcrExp 
         Caption         =   "&Clear Accured Expenses"
      End
      Begin VB.Menu dashAcr 
         Caption         =   "-"
      End
      Begin VB.Menu dash11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMonthlyAcrExps 
         Caption         =   "Monthly Accured/Prepaid Expenses"
      End
      Begin VB.Menu mnuTransfer 
         Caption         =   "&Transfer Funds"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuCustomer 
      Caption         =   "&Customers"
      Begin VB.Menu mnuNewCustomer 
         Caption         =   "&New Customer"
      End
      Begin VB.Menu mnuCustomers 
         Caption         =   "&Customers List"
         Shortcut        =   {F3}
      End
      Begin VB.Menu dashcust 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCustLedger 
         Caption         =   "View Last 30 Days &Ledger Of"
         Begin VB.Menu mnuCusts 
            Caption         =   "< No Customers Found >"
            Enabled         =   0   'False
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuVenders 
      Caption         =   "Venders"
      Begin VB.Menu mnuNewVender 
         Caption         =   "&New Vender"
      End
      Begin VB.Menu mnuVenderList 
         Caption         =   "Vender &List"
         Shortcut        =   {F4}
      End
      Begin VB.Menu dashvend 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVendLedger 
         Caption         =   "View Last 30 Days &Ledger Of"
         Begin VB.Menu mnuvends 
            Caption         =   "< No Venders Found >"
            Enabled         =   0   'False
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuDetailedLedger 
         Caption         =   "Detailed Account Ledger"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuMonthLedger 
         Caption         =   "Comparison  Ledger"
      End
      Begin VB.Menu dash123 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCStatus 
         Caption         =   "Cash && Bank Status"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuDueChqs 
         Caption         =   "Due Cheque"
      End
      Begin VB.Menu mnuDueRcpts 
         Caption         =   "Due Receipts From Customers"
      End
      Begin VB.Menu mnuDuePymts 
         Caption         =   "Due Payments To Venders"
      End
      Begin VB.Menu mnucomparison 
         Caption         =   "Monthly Transaction Analysis"
      End
      Begin VB.Menu dashtb 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrialBalance 
         Caption         =   "&Trial Balance ( Summary )"
      End
      Begin VB.Menu mnuTBalDetail 
         Caption         =   "Trial &Balance ( Detail )"
      End
      Begin VB.Menu dashbalsheet 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlsStd 
         Caption         =   "Profit && Loss ( Standered )"
      End
      Begin VB.Menu mnuPl 
         Caption         =   "&Profit && Loss ( Summary )"
      End
      Begin VB.Menu mnuPLDetail 
         Caption         =   "Profit && &Loss ( Detail )"
      End
      Begin VB.Menu dashtbal 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBSheet 
         Caption         =   "Balance Sheet"
      End
      Begin VB.Menu dashPLS 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCashBook 
         Caption         =   "Cash Book Report"
      End
      Begin VB.Menu mnuTrans 
         Caption         =   "Transaction Report"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuReception 
      Caption         =   "Reception"
      Begin VB.Menu mnuCallDetail 
         Caption         =   "Calls Detail"
      End
      Begin VB.Menu mnuNewCall 
         Caption         =   "New Calls"
      End
   End
   Begin VB.Menu mnuPayroll 
      Caption         =   "Payroll"
   End
   Begin VB.Menu mnugeneral 
      Caption         =   "Settings"
      Begin VB.Menu mnuChange 
         Caption         =   "Employee Settings"
      End
      Begin VB.Menu mnugendash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHolidays 
         Caption         =   "Define Holidays"
      End
      Begin VB.Menu mnuTaxRanges 
         Caption         =   "Tax Ranges"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuEmp 
      Caption         =   "Employees"
      Begin VB.Menu mnuDeptList 
         Caption         =   "&Department List"
      End
      Begin VB.Menu mnuNewDept 
         Caption         =   "&Add New Department"
      End
      Begin VB.Menu empdash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmpList 
         Caption         =   "Employee &List"
      End
      Begin VB.Menu mnunewEmp 
         Caption         =   "Add &New Employee"
      End
      Begin VB.Menu dashempsal 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmpLedger 
         Caption         =   "Employee Ledger"
      End
      Begin VB.Menu lnledger 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmpAssets 
         Caption         =   "Employee Assets"
      End
      Begin VB.Menu mnuSalaryLedger 
         Caption         =   "Employee &Salary History"
      End
   End
   Begin VB.Menu mnuAttSheet 
      Caption         =   "Attendance"
      Begin VB.Menu mnutake 
         Caption         =   "Take Attendance"
      End
      Begin VB.Menu mnuAttdash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSheet 
         Caption         =   "Attendance Sheet"
      End
   End
   Begin VB.Menu mnuAdvHistory 
      Caption         =   "Advances"
      Begin VB.Menu mnuShort 
         Caption         =   "New &Short Term Advance"
      End
      Begin VB.Menu lnLongloan 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFine 
         Caption         =   "Fine/Bonus To Employees"
      End
      Begin VB.Menu mnuOtherBenefits 
         Caption         =   "Employee's Other Benefits"
      End
      Begin VB.Menu advdash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShortAdv 
         Caption         =   "Short Term Advance Ledger"
      End
   End
   Begin VB.Menu mnuSalary 
      Caption         =   "Salary"
      Begin VB.Menu mnuSalSheet 
         Caption         =   "Salary Sheet"
      End
      Begin VB.Menu mnuAttdash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPaySlip 
         Caption         =   "Payslip"
      End
   End
   Begin VB.Menu mnuEmpReports 
      Caption         =   "Reports"
      Begin VB.Menu mnuApplicationForm 
         Caption         =   "Application Form"
      End
      Begin VB.Menu mnurptemplist 
         Caption         =   "Employee List "
      End
      Begin VB.Menu mnurptEmpCards 
         Caption         =   "Employee Cards"
      End
      Begin VB.Menu mnuAbsentSheet 
         Caption         =   "Absent Sheet"
      End
      Begin VB.Menu mnuPrintJoinLeave 
         Caption         =   "Employees Joined and Discontinued"
         Visible         =   0   'False
      End
      Begin VB.Menu dashRpts1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAttRegister 
         Caption         =   "Attendance Register"
      End
      Begin VB.Menu mnuExtSalsheet 
         Caption         =   "External Salary Sheet"
         Visible         =   0   'False
      End
      Begin VB.Menu dashsal 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSalaryComparison 
         Caption         =   "Monthly Salary Comparison"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOutStanding 
         Caption         =   "Out Standing Loans Report"
         Visible         =   0   'False
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

  
  
Private Sub MDIForm_Load()

mnuPrintOnSave.Checked = GetSetting(App.Title, "Settings", "PrintOnSave", True)

On Error Resume Next

   With Toolbar1
     Set .ImageList = ImageList2
     For i = 1 To .Buttons.Count
       If .Buttons(i).Style <> tbrSeparator Then .Buttons(i).Image = .Buttons(i).key
     Next
   
      If UseSundry Then

         With .Buttons("Accounts")
            .ButtonMenus("AccRec").Text = "Sundry Debtors"
            .ButtonMenus("AccPay").Text = "Sundry Creditors"
         End With
   
         mnuAccRcv.Caption = "Sundry Debtors"
         mnuAccPay.Caption = "Sundry Creditors"
   
      End If
   End With
   
   
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Call CleanUp
End Sub

Private Sub mnuAccPay_Click()
Dim f As New frmAccPayRcv
   With f
      .Tag = ACCPAYNO
      .Show
   End With
End Sub

Private Sub mnuAccRcv_Click()
Dim f As New frmAccPayRcv
With f
   .Tag = ACCRECNO
   .Show
End With
End Sub

Private Sub mnuAcrExp_Click()
   frmAcrExpenses.Show
End Sub

Private Sub mnuBackup_Click()
MsgBox "This Option is not Available.", vbInformation
End Sub

Private Sub mnuBankList_Click()
   frmBankList.Show
End Sub

Private Sub mnuBSheet_Click()
    frmBalSheet.Show
'   MsgBox "Balance Sheet is Under Developement", vbInformation
End Sub


Private Sub mnuCalculater_Click()
On Error GoTo ERR

Shell "calc.exe", vbNormalFocus

Exit Sub
ERR:
MsgBox "Windows calculater is not installed.", vbInformation

End Sub

Private Sub mnuCanceled_Click()
 frmCanceledItems.Show
End Sub

Private Sub mnuChaccounts_Click()

      Load frmChartOfAccounts
      frmChartOfAccounts.Show

End Sub

Private Sub mnuChangeDBPath_Click()

On Error GoTo ERR

If MsgBox("Current Database Path is '" & DBName & "'" & vbNewLine & _
       "Are You Sure To Change Database Path ?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
     
TheStart:
CD1.InitDir = App.Path
CD1.ShowOpen

Dim NewPath As String
NewPath = CD1.FileName

Set con = Nothing

If Not ConnectToDatabase() Then
      If MsgBox("Cannot Connect To Database Following Error Occured..." & vbNewLine & _
         ERR.Description & vbNewLine & "Do You Want To Select Database File ?", vbQuestion + vbSystemModal + vbYesNo, "DataBase Connection Error") = vbYes Then
         GoTo TheStart
              
      Else
         
         Call ConnectToDatabase
      End If
Else
   DBName = NewPath
   
   If MsgBox("Do You Want To Make This Database As Default Database ?", vbYesNo + vbQuestion) = vbYes Then
      SaveSetting App.Title, "Settings", "DatabasePath", DBName
   End If
   
   Unload frmMain
   Call getFinancialYear
   Load frmMain
   frmMain.Show
   FrmLogin.Show
End If


Exit Sub

ERR:
   If Not ERR.Number = 32755 Then MsgBox "Following Error Occured..." & vbNewLine & ERR.Description

End Sub

Private Sub mnuChangePassword_Click()
 With frmUserInfo
   .Show_Option = Opt_ChangeUserPassword
   .Tag = CurrentUserName
   .Show 1
 End With
End Sub

Private Sub mnuChqBooks_Click()
   frmChqBooks.Show
End Sub

Private Sub mnuCompanyInfo_Click()
    frmCompanyInformation.Show
End Sub

Private Sub mnuComparison_Click()
   frmComparison.Show
End Sub

Private Sub mnuCStatus_Click()
   frmCashStatus.Show
End Sub

Private Sub mnuCustomer_Click()
Call AddCustomersToMenu
End Sub

Private Sub mnuCustomers_Click()
    frmCustomerList.Show
End Sub

Private Sub mnuCusts_Click(Index As Integer)
Dim f As New frmLedger
With f
   .Show
   Call FindInCombo(.cmbAccNo, mnuCusts(Index).Tag)
   
   Call .cmdRefresh_Click
   .SetFocus
End With
End Sub

Private Sub mnuDailyIncomeSheet_Click()
    frmDailyIncome.Show
End Sub

Private Sub mnuDetailedLedger_Click()
frmLedger.Show
End Sub

Private Sub mnuDueChqs_Click()
frmChqs.Show
End Sub

Private Sub mnuDueRcpts_Click()
Dim f As New frmDueRcptsPymts
With f
   Call .FillCmbs(False)
   .Show
   .SetFocus
End With
End Sub
Private Sub mnuduePymts_Click()
Dim f As New frmDueRcptsPymts
With f
   Call .FillCmbs(True)
   .Show
   .SetFocus
End With
End Sub



Private Sub mnuEmpList_Click()
frmEmpList.Show
End Sub

Private Sub mnuEmployees_Click()
Call AddEmployeesToMenu
End Sub

Private Sub mnuEmps_Click(Index As Integer)
Dim f As New frmLedger
With f
   .Show
   Call FindInCombo(.cmbAccNo, mnuEmps(Index).Tag)
   
   Call .cmdRefresh_Click
   .SetFocus
End With
End Sub

Private Sub mnuExit_Click()
Call CleanUp
Unload Me
End Sub

Private Sub mnuExpenses_Click()
   frmExpenses.Show
End Sub

Private Sub mnuIncomePercentages_Click()
    frmPercentages.Show 1
End Sub

Private Sub mnuItemLedger_Click()
   frmItemLedger.Show
End Sub

Private Sub mnuItemRcpt_Click()
   'frmItemRcpt.Show
   frmOrderRcpt.Show
End Sub

Private Sub mnuItems_Click()
frmItems.Show
End Sub

Private Sub mnuItemSales_Click()
frmItemSale.Show
End Sub

Private Sub mnuJobs_Click()
frmJobs.Show
End Sub

Private Sub mnuJVoucher_Click()
    frmJV.Show
End Sub

Private Sub mnuLedger_Click()
    frmLedger.Show
End Sub


Private Sub mnuLogOff_Click()
If MsgBox("Are You Sure To Log Off ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
   Unload frmMain
   Load frmMain
   frmMain.Show
   FrmLogin.Show 1

End Sub

Private Sub mnuLongAdv_Click()
'Dim f As New frmEmpAdv'

'With f
'   .LongTermAdvances = True
'   .Show
'End With

End Sub

Private Sub mnuMonthLedger_Click()
   frmMonthLedger.Show
End Sub

Private Sub mnuMonthlyAcrExps_Click()
   frmMonthlyAccured.Show
End Sub

Private Sub mnuNewAccounts_Click()
'If getPassword Then
frmNewAccount.Show 1
End Sub

Private Sub mnuNewBankAcc_Click()

'If getPassword Then
   With frmNewAccount
      .chkBankAcc = vbChecked
      .Show 1
   End With
'End If
End Sub

Public Sub mnuNewCustomer_Click()
Dim f As New frmCustomer
With f
 .Add = True
 .Show 1
End With

End Sub

Public Sub mnuNewEmp_Click()
'Dim f As New frmEmp
'With f
' .Add = True
'  .Show 1
'End With
End Sub

Private Sub mnuNewExpAcc_Click()
'If getPassword Then
   Load frmNewAccount
   With frmNewAccount
      .ExpAccount = True
      .Show 1
   End With
'End If
End Sub

Public Sub mnuNewVender_Click()
Dim f As New frmVender
With f
 .Add = True
 .Show 1
End With

End Sub

Private Sub mnuOrders_Click()
frmOrders.Show
End Sub

Private Sub mnuOrderSale_Click()
frmOrderSale.Show
End Sub

Private Sub mnuPapers_Click()
frmPapers.Show
End Sub

Private Sub mnuPchOrders_Click()
frmPurchaseOrders.Show
End Sub

Private Sub mnuPL_Click()
Dim f As New frmPLS
With f
    .Summary = True
    .Standered = False
    .Show
End With
End Sub

Private Sub mnuPLDetail_Click()
Dim f As New frmPLS
With f
    .Summary = False
    .Show
End With
End Sub

Private Sub mnuPlsStd_Click()
Dim f As New frmPLS
With f
    .Summary = False
    .Standered = True
    .Show
End With
End Sub

Private Sub mnuPrintOnSave_Click()
   mnuPrintOnSave.Checked = Not mnuPrintOnSave.Checked
   Call SaveSetting(App.Title, "Settings", "PrintOnSave", mnuPrintOnSave.Checked)
End Sub

Private Sub mnuPRT_Click()
   frmPurchaseReturn.Show
End Sub

Private Sub mnuPurchase_Click()
   frmItemPurchase.Show
End Sub

Private Sub mnuPurchaseJV_Click()
If mnuStock.Visible Then
   frmItemPurchase.Show
Else
   frmPurchaseJV.Show
End If
End Sub

Private Sub mnuPurchaseOrders_Click()
   frmPaperOrder.Show
End Sub

Private Sub mnuPurchseReturn_Click()
   frmPchReturnJV.Show
End Sub

Private Sub mnuPurchaseReturn_Click()
If mnuStock.Visible Then
   frmPurchaseReturn.Show
Else
   frmPchReturnJV.Show
End If
End Sub

Private Sub mnuPVoucher_Click()
   frmCashPV.Show
End Sub

Private Sub mnuRates_Click()
frmRates.Show
End Sub

Private Sub mnuRegister_Click()
   frmRegister.Show
End Sub

Private Sub mnuRestore_Click()
   MsgBox "This Option is not Available.", vbInformation
End Sub

Private Sub mnuRVoucher_Click()
    frmCashRV.Show
End Sub

Private Sub mnuSalaries_Click()
frmSalaries.Show
End Sub

Private Sub mnuSaleJV_Click()
If mnuStock.Visible Then
   frmItemSale.Show
Else
   frmSaleJV.Show
End If
End Sub

Private Sub mnuSaleOrders_Click()
frmSaleOrders.Show
End Sub

Private Sub mnuSaleReturn_Click()
If mnuStock.Visible Then
   frmSaleReturn.Show
Else
   frmSaleReturnJV.Show
End If
End Sub

Private Sub mnuShortAdv_Click()
'Dim f As New frmEmpAdv
'With f
'   .LongTermAdvances = False
'   .Show
'End With


End Sub

Private Sub mnuSRT_Click()
   frmSaleReturn.Show
End Sub

Private Sub mnuStockAdj_Click()
If Not getPassword Then Exit Sub
frmStockAdj.Show
End Sub

Private Sub mnuTBalDetail_Click()
   Dim f As New frmTrialBalance
   f.chkDetail.value = vbChecked
   f.Show
End Sub

Private Sub mnuTestStat_Click()
    frmTestStats.Show 1
End Sub

Private Sub mnuTHTCIncome_Click()
     frmTHTC.Show 1
End Sub

Private Sub mnuTrans_Click()
Dim f As New frmTransaction
With f
   .Opts(0).value = True
   .Show
End With
End Sub

Private Sub mnuTransfer_Click()
   frmTransferFunds.Show 1
End Sub

Private Sub mnuTrialBalance_Click()
   Dim f As New frmTrialBalance
   f.Show
End Sub

Private Sub mnuUsers_Click()
   'MsgBox "User Management is Disabled At This Time.", vbInformation
frmUsers.Show 1
End Sub

Private Sub mnuVenderList_Click()
    frmVenderList.Show
End Sub

Private Sub mnuVenders_Click()
   Call AddVendersToMenu
End Sub

Private Sub AddCustomersToMenu()
Screen.MousePointer = vbHourglass
 Dim rs As New ADODB.Recordset
 With mnuCusts
.Item(0).Visible = True

 For i = 1 To .Count - 1
  Unload .Item(i)
 Next
 
 rs.CursorLocation = adUseClient
 rs.Open "Select AccNo,AccTitle From VCustomers Order By AccTitle", con, adOpenForwardOnly
 
  For i = 0 To rs.RecordCount - 1
   Load mnuCusts(i + 1)
   mnuCusts(i + 1).Caption = rs![AccTitle] & ""
   mnuCusts(i + 1).Enabled = True
   mnuCusts(i + 1).Tag = rs![AccNo] & ""
   rs.MoveNext
  Next
 End With
 
 If rs.RecordCount > 0 Then mnuCusts(0).Visible = False
 
   rs.Close
   Set rs = Nothing
  
Screen.MousePointer = vbDefault
End Sub


Private Sub AddEmployeesToMenu()
Screen.MousePointer = vbHourglass
 Dim rs As New ADODB.Recordset
 With mnuEmps
.Item(0).Visible = True
 For i = 1 To .Count - 1
  Unload .Item(i)
 Next
 
 rs.CursorLocation = adUseClient
 rs.Open "Select AccNo,AccTitle From VEmps Order By AccTitle", con, adOpenForwardOnly
 
  For i = 0 To rs.RecordCount - 1
   Load mnuEmps(i + 1)
   mnuEmps(i + 1).Caption = rs![AccTitle] & ""
   mnuEmps(i + 1).Enabled = True
   mnuEmps(i + 1).Tag = rs![AccNo] & ""
   rs.MoveNext
  Next
 End With
 
 If rs.RecordCount > 0 Then mnuEmps(0).Visible = False
 
   rs.Close
   Set rs = Nothing
  
Screen.MousePointer = vbDefault
End Sub
Private Sub AddVendersToMenu()
Screen.MousePointer = vbHourglass
 Dim rs As New ADODB.Recordset
 With mnuvends
 .Item(0).Visible = True
 For i = 1 To .Count - 1
  Unload .Item(i)
 Next
 
 rs.CursorLocation = adUseClient
 rs.Open "Select AccNo,AccTitle From VVenders Order By AccTitle", con, adOpenForwardOnly
 
  For i = 0 To rs.RecordCount - 1
   Load mnuvends(i + 1)
   mnuvends(i + 1).Caption = rs![AccTitle] & ""
   mnuvends(i + 1).Enabled = True
   mnuvends(i + 1).Tag = rs![AccNo] & ""
   rs.MoveNext
  Next
End With
 If rs.RecordCount > 0 Then mnuvends(0).Visible = False
 
   rs.Close
   Set rs = Nothing
  
Screen.MousePointer = vbDefault
End Sub

Private Sub mnuVends_Click(Index As Integer)
Dim f As New frmLedger
With f
   .Show
   Call FindInCombo(.cmbAccNo, mnuvends(Index).Tag)

   Call .cmdRefresh_Click
      .SetFocus
End With

End Sub

Private Sub mnuCashBook_Click()

'For Helicon and Bhinder Show Cash in hand Ledger

'If LCase(company) Like "helicon*" Or LCase(company) Like "bhinder*" Then

If LCase(company) Like "bhinder*" Then

   With frmLedger
      .Show
      .cmbAccNo = CashInHand
      .DTFrom = Date
      .DTTo = Date
      .cmdRefresh_Click
      .SetFocus
   End With
   
Else
   Dim f As New frmTransaction
   With f
      .Opts(1).value = True
      .Show
   End With
End If

End Sub

Private Sub Timer1_Timer()
   Me.Caption = App.Title & ".   [" & Format(Now, "dddd dd/MM/yyyy hh:mm:ss AMPM") & "]"
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
Case "Window"
   Me.PopupMenu mnuWindow, 8, Toolbar1.Left + Button.Left + Button.Width, Toolbar1.Top + Toolbar1.Height
Case "Accounts"
  Call mnuChaccounts_Click
Case "Customers"
   Call frmCustomerList.Show
Case "Venders"
   Call frmVenderList.Show
Case "Banks"
   Call mnuChqBooks_Click
Case "RV"
   Call mnuRVoucher_Click
Case "PV"
   Call mnuPVoucher_Click
Case "JV"
   Call mnuJVoucher_Click
Case "Ledger"
   frmLedger.Show
Case "Expenses"
   Call mnuExpenses_Click
Case "SJV"
    mnuSaleJV_Click
Case "PJV"
    mnuPurchaseJV_Click
Case "EmpAdv"
   mnuEmpList_Click
End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.key
Case "ChartOfAccounts"
   Call mnuChaccounts_Click
Case "AccPay"
   Call mnuAccPay_Click
Case "AccRec"
   Call mnuAccRcv_Click
Case "NewCustomer"
   Call mnuNewCustomer_Click
Case "CustomerList"
   frmCustomerList.Show
Case "NewVender"
   Call mnuNewVender_Click
Case "VenderList"
   frmVenderList.Show
Case "NewBank"
   Call mnuNewBankAcc_Click
Case "BankAccs"
   Call mnuChqBooks_Click
Case "BankList"
   Call mnuBankList_Click
Case "EmpAdv"
   mnuShortAdv_Click
Case "EmpLongAdv"
   mnuLongAdv_Click
Case "EmpSal"
   mnuSalaries_Click
Case "ExpAcc"
   mnuNewExpAcc_Click
Case "ExpVchr"
   mnuExpenses_Click
Case "AcrExp"
   mnuAcrExp_Click
Case "MonthAcr"
   mnuMonthlyAcrExps_Click
Case "Ledger"
   frmLedger.Show
Case "MonLgr"
   frmMonthLedger.Show
Case "TransRpt"
   Call mnuTrans_Click
Case "PJV"
   mnuPurchaseJV_Click
Case "PRJV"
   mnuPurchaseReturn_Click
Case "SJV"
   mnuSaleJV_Click
Case "SRJV"
   mnuSaleReturn_Click
Case "EmpList"
   mnuEmpList_Click
Case "NewEmp"
   mnuNewEmp_Click
   
End Select
End Sub








Public Sub ApplySecuritySettings()
   
   Dim Out As Boolean

   Out = OutOfFinancialYear
    
   mnuVouchers.Visible = Not Out
         
   'Accounts
   mnuChaccounts.Enabled = UserHasAccess("ChartofAccounts")
   mnuAccPay.Enabled = UserHasAccess("AccountsPayableReceivable")
   mnuAccRcv.Enabled = UserHasAccess("AccountsPayableReceivable")
   
   mnuNewBankAcc.Enabled = UserHasAccess("BanksChqBooks") And Not Out   'of out of financial year date then disable this option
   mnuChqBooks.Enabled = UserHasAccess("BanksChqBooks")
   
   
   
   mnuNewAccounts.Enabled = UserHasAccess("NewAccount") And Not Out
   mnuNewExpAcc.Enabled = UserHasAccess("NewExpenseAccount") And Not Out
   'Vouchers
   mnuJVoucher.Enabled = UserHasAccess("JournalVoucher") And Not Out
   mnuRVoucher.Enabled = UserHasAccess("ReceiptVoucher") And Not Out
   mnuPVoucher.Enabled = UserHasAccess("PaymentVoucher") And Not Out
   mnuExpenses.Enabled = UserHasAccess("ExpenseVoucher") And Not Out
   mnuTransfer.Enabled = UserHasAccess("TransferFunds") And Not Out
   mnuMonthlyAcrExps.Enabled = UserHasAccess("AccuredPrepaidVoucher") And Not Out
   'Customers
   mnuCustomers.Enabled = UserHasAccess("CustomerList")
   mnuNewCustomer.Enabled = UserHasAccess("NewCustomer") And Not Out
   mnuSaleJV.Enabled = UserHasAccess("SaleVoucher")
   mnuSaleReturn.Enabled = UserHasAccess("SaleReturnVoucher")
   
   'Venders
   mnuVenderList.Enabled = UserHasAccess("VenderList")
   mnuNewVender.Enabled = UserHasAccess("NewVender") And Not Out
   mnuPurchaseJV.Enabled = UserHasAccess("PurchaseVoucher")
   mnuPurchaseReturn.Enabled = UserHasAccess("PurchaseReturnVoucher")
   
   'Employees
      
      
      ' Employees Options Are Included in Package Or Not
   
      'mnuEmployees.Visible = EmployeesEnabled
      'mnuSalaries.Visible = EmployeesEnabled
      'mnuAdvances.Visible = EmployeesEnabled
      'dash11.Visible = EmployeesEnabled
   
   
   
   
   mnuEmpList.Enabled = UserHasAccess("EmployeeList")
   mnunewEmp.Enabled = UserHasAccess("NewEmployee") And Not Out
'   mnuShortAdv.Enabled = UserHasAccess("ShortTermAdvance")
'   mnuLongAdv.Enabled = UserHasAccess("LongTermAdvance")
'   mnuShortAdv.Enabled = UserHasAccess("ShortTermAdvance")
'   mnuSalaries.Enabled = UserHasAccess("SalariesVoucher")
   
   
   'Detailed Ledgers
   mnuDetailedLedger.Enabled = UserHasAccess("DetailedLedger")
   mnuCustLedger.Enabled = mnuDetailedLedger.Enabled
   mnuVendLedger.Enabled = mnuDetailedLedger.Enabled
   mnuEmpLedger.Enabled = mnuDetailedLedger.Enabled
   
   'Reports
   mnuMonthLedger.Enabled = UserHasAccess("MonthlyLedger")
   mnuDueChqs.Enabled = UserHasAccess("DueCheques")
   mnuCStatus.Enabled = UserHasAccess("CashBankStatus")
   mnuTrans.Enabled = UserHasAccess("TransactionReport")
   mnuTrialBalance.Enabled = UserHasAccess("TrialBalance")
   mnuTBalDetail.Enabled = UserHasAccess("TrialBalance")
   mnuCashBook.Enabled = UserHasAccess("CashBookReport")
   mnuPLDetail.Enabled = UserHasAccess("ProfitLoss")
   mnuPlsStd.Enabled = UserHasAccess("ProfitLoss")
   mnuPl.Enabled = UserHasAccess("ProfitLoss")
   mnuDueRcpts.Enabled = UserHasAccess("DueReceiptsPayments")
   mnuduePymts.Enabled = UserHasAccess("DueReceiptsPayments")
   
   mnuBSheet.Enabled = UserHasAccess("BalanceSheet")
   
   'Users
   mnuUsers.Enabled = UserHasAccess("UserManagement")
   
   'Backup,Restore & Other Option
   mnuBackup.Enabled = UserHasAccess("BackupData")
   mnuRestore.Enabled = UserHasAccess("RestoreData")
   
   mnuChangePassword.Enabled = UserHasAccess("ChangePassword")
   
    
  With Toolbar1
   
   .Buttons("Accounts").Enabled = mnuChaccounts.Enabled
      With .Buttons("Accounts").ButtonMenus
         .Item("AccPay").Enabled = mnuAccPay.Enabled
         .Item("AccRec").Enabled = mnuAccRcv.Enabled
         .Item("ChartOfAccounts").Enabled = mnuChaccounts.Enabled
      End With
   
   'Banks Button
      .Buttons("Banks").Enabled = mnuChqBooks.Enabled
      With .Buttons("Banks").ButtonMenus
         .Item("BankList").Enabled = mnuBankList.Enabled
         .Item("BankAccs").Enabled = mnuChqBooks.Enabled
         .Item("NewBank").Enabled = mnuNewBankAcc.Enabled And Not Out
      End With
      
      'Customer Button
      .Buttons("Customers").Enabled = mnuCustomers.Enabled
      With .Buttons("Customers").ButtonMenus
         .Item("CustomerList").Enabled = mnuCustomers.Enabled
         .Item("NewCustomer").Enabled = mnuNewCustomer.Enabled And Not Out
      End With
      
       'Vender Button
      .Buttons("Venders").Enabled = mnuVenderList.Enabled
      With .Buttons("Venders").ButtonMenus
         .Item("VenderList").Enabled = mnuVenderList.Enabled
         .Item("NewVender").Enabled = mnuNewVender.Enabled And Not Out
      End With
      
        'Employees Button
       
       .Buttons("EmpAdv").Visible = EmployeesEnabled
       
      .Buttons("EmpAdv").Enabled = mnuEmpList.Enabled
      With .Buttons("EmpAdv").ButtonMenus
         .Item("EmpAdv").Enabled = mnuShortAdv.Enabled And Not Out
'         .Item("EmpLongAdv").Enabled = mnuLongAdv.Enabled And Not Out
'         .Item("EmpSal").Enabled = mnuSalaries.Enabled And Not Out
         .Item("EmpList").Enabled = mnuEmpList.Enabled
         .Item("NewEmp").Enabled = mnunewEmp.Enabled And Not Out
      End With
      
      
      'Expense Voucher Button
      .Buttons("Expenses").Enabled = mnuExpenses.Enabled And Not Out
      
      With .Buttons("Expenses").ButtonMenus
         .Item("ExpVchr").Enabled = mnuExpenses.Enabled
         .Item("ExpAcc").Enabled = mnuNewExpAcc.Enabled
         .Item("AcrExp").Enabled = mnuAcrExp.Enabled
         .Item("MonthAcr").Enabled = mnuMonthlyAcrExps.Enabled
      End With
     
      .Buttons("JV").Enabled = mnuJVoucher.Enabled And Not Out
      .Buttons("SJV").Enabled = mnuSaleJV.Enabled And Not Out
      .Buttons("PJV").Enabled = mnuPurchaseJV.Enabled And Not Out
      .Buttons("RV").Enabled = mnuRVoucher.Enabled And Not Out
      .Buttons("PV").Enabled = mnuPVoucher.Enabled And Not Out
      
  
 
    
    'Ledger Button
      .Buttons("Ledger").Enabled = mnuDetailedLedger.Enabled
      
      With .Buttons("Ledger").ButtonMenus
         .Item("Ledger").Enabled = mnuDetailedLedger.Enabled
         .Item("MonLgr").Enabled = mnuMonthLedger.Enabled
      End With
      
End With
   'Show Stock And Press Menus Accourding To Package Type
   
   mnuStock.Visible = (InStr(UCase(App.Title), "STOCK") > 0)
   mnuPress.Visible = (InStr(UCase(App.Title), "PRESS") > 0)

End Sub

