VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmUserSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Access Rights"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9675
   Icon            =   "frmUserSettings.frx":0000
   LinkTopic       =   "Form24"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9675
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Vouchers :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1350
      Left            =   4815
      TabIndex        =   48
      Top             =   630
      Width           =   4770
      Begin VB.CheckBox Check1 
         Caption         =   "Edit Voucher Description"
         DataField       =   "EditVoucherDesc"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   15
         Left            =   2265
         TabIndex        =   9
         Top             =   645
         Width           =   2205
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Transfer Funds"
         DataField       =   "TransferFunds"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   13
         Left            =   2265
         TabIndex        =   7
         Top             =   300
         Width           =   2205
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Delete A Voucher"
         DataField       =   "DeleteVoucher"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   11
         Left            =   2265
         TabIndex        =   11
         Top             =   990
         Width           =   2205
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add Expe&nse Voucher"
         DataField       =   "AccuredPrepaidVoucher"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   10
         Top             =   996
         Width           =   2205
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Accured/Prepaid Voucher"
         DataField       =   "ReceiptVoucher"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   648
         Width           =   2130
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add &Journal Voucher"
         DataField       =   "JournalVoucher"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   2205
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Accounts :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1350
      Left            =   30
      TabIndex        =   47
      Top             =   630
      Width           =   4770
      Begin VB.CheckBox Check1 
         Caption         =   "&View Chart of Accounts"
         DataField       =   "ChartOfAccounts"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   0
         Top             =   300
         Width           =   2100
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Receivable / Payables"
         DataField       =   "AccountsPayableReceivable"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   2
         Top             =   645
         Width           =   2100
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Bank A/C && ChqBooks"
         DataField       =   "BanksChqBooks"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   4
         Top             =   990
         Width           =   2100
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add &New Account"
         DataField       =   "NewAccount"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   2280
         TabIndex        =   1
         Top             =   300
         Width           =   2310
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add New E&xpense Account"
         DataField       =   "NewExpenseAccount"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   2280
         TabIndex        =   3
         Top             =   645
         Width           =   2310
      End
      Begin VB.CheckBox Check1 
         Caption         =   "&Edit Account Information"
         DataField       =   "EditAccount"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   2280
         TabIndex        =   5
         Top             =   990
         Width           =   2310
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Customers :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1350
      Left            =   30
      TabIndex        =   49
      Top             =   2010
      Width           =   4770
      Begin VB.CheckBox Check1 
         Caption         =   "Add Sale Return Voucher"
         DataField       =   "SaleReturnVoucher"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   10
         Left            =   2265
         TabIndex        =   17
         Top             =   1020
         Width           =   2205
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add Sale Voucher"
         DataField       =   "SaleVoucher"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   9
         Left            =   2265
         TabIndex        =   15
         Top             =   660
         Width           =   2205
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add Re&ceipt Voucher"
         DataField       =   "ReceiptVoucher"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   2265
         TabIndex        =   13
         Top             =   315
         Width           =   2205
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Edit Customer Info."
         DataField       =   "EditCustomer"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   20
         Left            =   135
         TabIndex        =   16
         Top             =   1005
         Width           =   1980
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add New Customer"
         DataField       =   "NewCustomer"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   19
         Left            =   120
         TabIndex        =   14
         Top             =   660
         Width           =   1980
      End
      Begin VB.CheckBox Check1 
         Caption         =   "View List of Customers"
         DataField       =   "CustomerList"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   18
         Left            =   135
         TabIndex        =   12
         Top             =   300
         Width           =   1980
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Venders :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1350
      Left            =   4815
      TabIndex        =   50
      Top             =   2010
      Width           =   4770
      Begin VB.CheckBox Check1 
         Caption         =   "View List of Venders"
         DataField       =   "VenderList"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   23
         Left            =   120
         TabIndex        =   18
         Top             =   300
         Width           =   2115
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add New Vender"
         DataField       =   "NewVender"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   22
         Left            =   135
         TabIndex        =   20
         Top             =   660
         Width           =   2115
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Edit Vender Information"
         DataField       =   "EditVender"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   21
         Left            =   120
         TabIndex        =   22
         Top             =   1005
         Width           =   2115
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add Payment Voucher"
         DataField       =   "PaymentVoucher"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   17
         Left            =   2235
         TabIndex        =   19
         Top             =   300
         Width           =   2205
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add Purchase Voucher"
         DataField       =   "PurchaseVoucher"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   16
         Left            =   2235
         TabIndex        =   21
         Top             =   645
         Width           =   2205
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add Purchase Return Voucher"
         DataField       =   "PurchaseReturnVoucher"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   12
         Left            =   2235
         TabIndex        =   23
         Top             =   1005
         Width           =   2490
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Employees :"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1350
      Left            =   4815
      TabIndex        =   51
      Top             =   3360
      Width           =   4785
      Begin VB.CheckBox Check1 
         Caption         =   "View Employee List"
         DataField       =   "EmployeeList"
         DataSource      =   "Adodc1"
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
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   29
         Left            =   135
         TabIndex        =   35
         Top             =   300
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add New Employee"
         DataField       =   "NewEmployee"
         DataSource      =   "Adodc1"
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
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   28
         Left            =   135
         TabIndex        =   37
         Top             =   645
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Edit Employee Info."
         DataField       =   "EditEmployee"
         DataSource      =   "Adodc1"
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
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   27
         Left            =   135
         TabIndex        =   39
         Top             =   990
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add Salaries Voucher"
         DataField       =   "SalariesVoucher"
         DataSource      =   "Adodc1"
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
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   26
         Left            =   2220
         TabIndex        =   40
         Top             =   975
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add Short Term  Adv. Voucher"
         DataField       =   "ShortTermAdvance"
         DataSource      =   "Adodc1"
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
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   25
         Left            =   2220
         TabIndex        =   36
         Top             =   285
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add Long Term  Adv. Voucher"
         DataField       =   "LongTermAdvance"
         DataSource      =   "Adodc1"
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
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   24
         Left            =   2220
         TabIndex        =   38
         Top             =   630
         Width           =   2535
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Reports :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2400
      Left            =   30
      TabIndex        =   52
      Top             =   3360
      Width           =   4770
      Begin VB.CheckBox Check1 
         Caption         =   "Balance Sheet"
         DataField       =   "BalanceSheet"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   40
         Left            =   2220
         TabIndex        =   31
         Top             =   1350
         Width           =   2205
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Transaction Report"
         DataField       =   "TransactionReport"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   39
         Left            =   165
         TabIndex        =   34
         Top             =   2010
         Width           =   2265
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Cash Book Report"
         DataField       =   "CAshBookReport"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   36
         Left            =   165
         TabIndex        =   32
         Top             =   1665
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Trial Balance Reports"
         DataField       =   "TrialBalance"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   38
         Left            =   2220
         TabIndex        =   25
         Top             =   300
         Width           =   2205
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Profit && Loss Reports"
         DataField       =   "ProfitLoss"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   37
         Left            =   2220
         TabIndex        =   27
         Top             =   650
         Width           =   2205
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Monthly Trans. Comparison"
         DataField       =   "MonthlyTransComparison"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   35
         Left            =   2220
         TabIndex        =   29
         Top             =   1000
         Width           =   2505
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Cash && Bank Status"
         DataField       =   "CashBankStatus"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   34
         Left            =   165
         TabIndex        =   30
         Top             =   1323
         Width           =   2235
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Due Receipts && Payments"
         DataField       =   "DueReceiptsPayments"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   33
         Left            =   2220
         TabIndex        =   33
         Top             =   1665
         Width           =   2205
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Due Cheques"
         DataField       =   "DueCheques"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   32
         Left            =   165
         TabIndex        =   28
         Top             =   982
         Width           =   2235
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Monthly Ledger"
         DataField       =   "MonthlyLedger"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   31
         Left            =   165
         TabIndex        =   26
         Top             =   641
         Width           =   2235
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Detailed Ledger"
         DataField       =   "DetailedLedger"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   30
         Left            =   165
         TabIndex        =   24
         Top             =   300
         Width           =   2235
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "General :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1035
      Left            =   4815
      TabIndex        =   53
      Top             =   4725
      Width           =   4785
      Begin VB.CheckBox Check1 
         Caption         =   "Change Password"
         DataField       =   "ChangePassword"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   41
         Left            =   2205
         TabIndex        =   44
         Top             =   690
         Width           =   1650
      End
      Begin VB.CheckBox Check1 
         Caption         =   "User Management"
         DataField       =   "UserManagement"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   44
         Left            =   2205
         TabIndex        =   42
         Top             =   300
         Width           =   1650
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Backup Data"
         DataField       =   "BackupData"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   43
         Left            =   120
         TabIndex        =   41
         Top             =   330
         Width           =   1440
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Restore Data"
         DataField       =   "RestoreData"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   42
         Left            =   120
         TabIndex        =   43
         Top             =   675
         Width           =   1440
      End
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   375
      Left            =   8190
      TabIndex        =   46
      Top             =   5850
      Width           =   1425
      Caption         =   " Cancel       "
      PicturePosition =   327683
      Size            =   "2514;661"
      Picture         =   "frmUserSettings.frx":000C
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdSave 
      Height          =   375
      Left            =   6720
      TabIndex        =   45
      Top             =   5850
      Width           =   1425
      Caption         =   "Save         "
      PicturePosition =   327683
      Size            =   "2514;661"
      Picture         =   "frmUserSettings.frx":011E
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label USer 
      AutoSize        =   -1  'True
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   675
      TabIndex        =   56
      Top             =   0
      Width           =   915
   End
   Begin VB.Label lblUserName 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "UserName"
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
      Left            =   645
      TabIndex        =   55
      Top             =   225
      Width           =   3555
   End
   Begin VB.Label Label1 
      Caption         =   "This User Has Following Access Rights, Selected Option Shows That Access Is Granted While Unselected Access is Denied  . . . "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   4275
      TabIndex        =   54
      Top             =   105
      Width           =   5385
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   90
      Picture         =   "frmUserSettings.frx":0230
      Top             =   -30
      Width           =   540
   End
End
Attribute VB_Name = "frmUserSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsUser As New ADODB.Recordset
Dim Loaded As Boolean
Private Sub cmdCancel_Click()
On Error GoTo ERR
   rsUser.CancelUpdate
   Unload Me
 Exit Sub
ERR:
MsgBox ERR.Description
End Sub

Private Sub cmdSave_Click()
   rsUser.Update
   Unload Me
End Sub



Private Sub Form_Activate()
If Loaded Then Exit Sub
Me.Hide
With rsUser
   
   .Open "Select * From Users Where UserName='" & Me.Tag & "'", con, adOpenDynamic, adLockOptimistic
   'On Error GoTo err
   For i = 0 To Check1.Count - 1
       Set Check1(i).DataSource = rsUser
   Next
   lblUserName.Caption = " " & Me.Tag
   Loaded = True
End With
Me.Show 1
End Sub


Private Sub Form_Load()
   Loaded = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   If rsUser.State = 1 Then rsUser.Close
   Set rsUser = Nothing
End Sub
