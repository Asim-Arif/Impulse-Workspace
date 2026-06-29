VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmUserSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Access Rights"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "frmUserSettings.frx":0000
   LinkTopic       =   "Form24"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Departements :"
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
      Height          =   615
      Left            =   105
      TabIndex        =   38
      Top             =   765
      Width           =   4770
      Begin VB.CheckBox Check1 
         Caption         =   "Departement List"
         DataField       =   "DeptList"
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
         Left            =   120
         TabIndex        =   40
         Top             =   255
         Width           =   1950
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add New Departement"
         DataField       =   "NewDept"
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
         Index           =   29
         Left            =   2385
         TabIndex        =   39
         Top             =   255
         Width           =   2070
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Loans And Salaries :"
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
      Height          =   1605
      Left            =   120
      TabIndex        =   29
      Top             =   4035
      Width           =   4770
      Begin VB.CheckBox Check1 
         Caption         =   "Employee Salary History"
         DataField       =   "SalaryLedger"
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
         Left            =   120
         TabIndex        =   44
         Top             =   1260
         Width           =   2115
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Employee Loan History"
         DataField       =   "LoanLedger"
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
         Left            =   2340
         TabIndex        =   43
         Top             =   1260
         Width           =   2205
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Make Advance. Salary Final"
         DataField       =   "FinalAdvSal"
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
         Left            =   2340
         TabIndex        =   20
         Top             =   620
         Width           =   2340
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Issue Long Term Loan"
         DataField       =   "LongTermLoan"
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
         Left            =   2340
         TabIndex        =   18
         Top             =   300
         Width           =   2205
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Make Salary Sheet Final"
         DataField       =   "FinalSalSheet"
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
         Left            =   2340
         TabIndex        =   22
         Top             =   940
         Width           =   2205
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Calculate Salary Sheet"
         DataField       =   "CalculateSalSheet"
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
         TabIndex        =   21
         Top             =   940
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Calculate Advance Salary"
         DataField       =   "CalculateAdvSal"
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
         TabIndex        =   19
         Top             =   620
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Issue Short Term Loan"
         DataField       =   "ShortTermLoan"
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
         TabIndex        =   17
         Top             =   300
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Employees :"
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
      Height          =   1290
      Left            =   105
      TabIndex        =   28
      Top             =   1395
      Width           =   4770
      Begin VB.CheckBox Check1 
         Caption         =   "View  Employee List"
         DataField       =   "EmployeeList"
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
         Left            =   105
         TabIndex        =   42
         Top             =   255
         Width           =   2130
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add New Employee"
         DataField       =   "NewEmployee"
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
         Left            =   2370
         TabIndex        =   0
         Top             =   255
         Width           =   2100
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Change Emp. Salary"
         DataField       =   "ChangeEmpSal"
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
         Left            =   105
         TabIndex        =   2
         Top             =   1005
         Width           =   2100
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Print Employee List"
         DataField       =   "PrintEmpList"
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
         Left            =   2370
         TabIndex        =   4
         Top             =   1005
         Width           =   2100
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Edit Employee Information"
         DataField       =   "EditEmployee"
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
         Left            =   2370
         TabIndex        =   1
         Top             =   630
         Width           =   2310
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Make Emp. Active/Inactive"
         DataField       =   "EmpActive"
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
         Left            =   105
         TabIndex        =   3
         Top             =   615
         Width           =   2310
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Attendance/Over Time : "
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
      Height          =   1290
      Left            =   105
      TabIndex        =   30
      Top             =   2715
      Width           =   4770
      Begin VB.CheckBox Check1 
         Caption         =   "Edit Over Time Record"
         DataField       =   "EditOverTime"
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
         Left            =   2355
         TabIndex        =   10
         Top             =   1020
         Width           =   2205
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Attendance Shift Changable"
         DataField       =   "AttShiftChange"
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
         Left            =   2355
         TabIndex        =   8
         Top             =   660
         Width           =   2400
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Edit Employee Attendance"
         DataField       =   "EditAttendance"
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
         Left            =   2355
         TabIndex        =   6
         Top             =   315
         Width           =   2205
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Add Over Time"
         DataField       =   "AddOverTime"
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
         Left            =   135
         TabIndex        =   9
         Top             =   1005
         Width           =   1980
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Att. Date Changable"
         DataField       =   "AttDateChange"
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
         Left            =   135
         TabIndex        =   7
         Top             =   660
         Width           =   1980
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Take Attendance "
         DataField       =   "TakeAttendance"
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
         Left            =   135
         TabIndex        =   5
         Top             =   300
         Width           =   1980
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
      Height          =   3255
      Left            =   4905
      TabIndex        =   31
      Top             =   765
      Width           =   2460
      Begin VB.CheckBox Check1 
         Caption         =   "Print Employee Cards"
         DataField       =   "PrintEmpCards"
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
         Left            =   105
         TabIndex        =   41
         Top             =   947
         Width           =   2310
      End
      Begin VB.CheckBox Check1 
         Caption         =   "OldAge"
         DataField       =   "OldAge"
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
         Index           =   28
         Left            =   105
         TabIndex        =   37
         Top             =   2940
         Width           =   2280
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Social Security"
         DataField       =   "SocialSecurity"
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
         Index           =   27
         Left            =   105
         TabIndex        =   36
         Top             =   2602
         Width           =   2280
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Monthly Salary Comparison"
         DataField       =   "SalaryComparison"
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
         Left            =   105
         TabIndex        =   14
         Top             =   1609
         Width           =   2280
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Attendance Register"
         DataField       =   "AttendanceRegister"
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
         Left            =   105
         TabIndex        =   15
         Top             =   1940
         Width           =   2280
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Employee Pay Slip"
         DataField       =   "EmpPaySlip"
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
         Left            =   105
         TabIndex        =   16
         Top             =   2271
         Width           =   2280
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Employee Ledger"
         DataField       =   "EmployeeLedger"
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
         Left            =   105
         TabIndex        =   13
         Top             =   1278
         Width           =   1845
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Application Form"
         DataField       =   "ApplicationForm"
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
         Left            =   105
         TabIndex        =   12
         Top             =   616
         Width           =   1845
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Daily Absents Report"
         DataField       =   "DailyAbsents"
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
         Left            =   105
         TabIndex        =   11
         Top             =   285
         Width           =   1845
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
      Height          =   1605
      Left            =   4920
      TabIndex        =   32
      Top             =   4035
      Width           =   2460
      Begin VB.CheckBox Check1 
         Caption         =   "Show Reminders"
         DataField       =   "Reminders"
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
         Left            =   105
         TabIndex        =   46
         Top             =   1335
         Width           =   1950
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Edit Tax Ranges"
         DataField       =   "TaxRanges"
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
         Left            =   105
         TabIndex        =   45
         Top             =   1050
         Width           =   1950
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
         Index           =   26
         Left            =   105
         TabIndex        =   25
         Top             =   765
         Width           =   1650
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Departement Settings"
         DataField       =   "DeptSettings"
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
         Index           =   25
         Left            =   105
         TabIndex        =   23
         Top             =   195
         Width           =   1950
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Define Holidays"
         DataField       =   "DefineHolidays"
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
         Index           =   24
         Left            =   105
         TabIndex        =   24
         Top             =   480
         Width           =   1440
      End
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   0
      Picture         =   "frmUserSettings.frx":000C
      Top             =   120
      Width           =   480
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   375
      Left            =   5940
      TabIndex        =   27
      Top             =   5730
      Width           =   1425
      Caption         =   " Cancel       "
      PicturePosition =   327683
      Size            =   "2514;661"
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
      Left            =   4470
      TabIndex        =   26
      Top             =   5730
      Width           =   1425
      Caption         =   "Save         "
      PicturePosition =   327683
      Size            =   "2514;661"
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
      Left            =   615
      TabIndex        =   35
      Top             =   120
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
      Left            =   615
      TabIndex        =   34
      Top             =   360
      Width           =   3105
   End
   Begin VB.Label Label1 
      Caption         =   "This User Has Following Access Rights, Selected Option Shows That Access Is Granted While Unselected Access is Denied  . . . "
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
      Height          =   570
      Left            =   3765
      TabIndex        =   33
      Top             =   105
      Width           =   3585
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
   
   .Open "Select * From Users Where UserName='" & Me.Tag & "'", Con, adOpenDynamic, adLockOptimistic
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
