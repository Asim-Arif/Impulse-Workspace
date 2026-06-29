VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmEmpDashBoard 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Delete Voucher"
   ClientHeight    =   10560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmpDashBoard.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10560
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAbsents 
      BackColor       =   &H00FFC0C0&
      Caption         =   "7 Absents"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10860
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   10110
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdSalaryLedger 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Salary Ledger"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3270
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   10110
      Width           =   1545
   End
   Begin VB.CommandButton cmdPrintProdReport 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Productivity Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   10110
      Width           =   1905
   End
   Begin VB.TextBox txtTotalSTRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Left            =   6150
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   10110
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox txtTotalRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Left            =   7185
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   10110
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   405
      Left            =   8220
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   10110
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10110
      Width           =   1215
   End
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   9495
      Index           =   0
      Left            =   60
      ScaleHeight     =   9495
      ScaleWidth      =   12120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   570
      Width           =   12120
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         Height          =   9495
         Left            =   30
         TabIndex        =   7
         Top             =   -60
         Width           =   12060
         Begin ComboList.Usercontrol1 cmbEmp 
            Height          =   330
            Left            =   60
            TabIndex        =   35
            Top             =   1305
            Width           =   9120
            _ExtentX        =   16087
            _ExtentY        =   582
            BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ListIndex       =   -1
            Appearance      =   0
         End
         Begin VB.TextBox txtAddress 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   750
            Left            =   2910
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   47
            Top             =   5985
            Width           =   4725
         End
         Begin VB.TextBox txtNotices 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   915
            Left            =   2910
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   43
            Top             =   7860
            Width           =   4725
         End
         Begin VB.TextBox txtReview 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   690
            Left            =   2910
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   41
            Top             =   6945
            Width           =   4725
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Factory Details :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   930
            Left            =   660
            TabIndex        =   22
            Top             =   120
            Width           =   10980
            Begin VB.TextBox txtTotalEmps 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   795
               TabIndex        =   28
               Top             =   525
               Width           =   1560
            End
            Begin VB.TextBox txtTotalPresents 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   2385
               TabIndex        =   27
               Top             =   525
               Width           =   1560
            End
            Begin VB.TextBox txtTotalAbsents 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   3975
               TabIndex        =   26
               Top             =   525
               Width           =   1560
            End
            Begin VB.TextBox txtTotalSalary 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   5565
               TabIndex        =   25
               Top             =   525
               Width           =   1560
            End
            Begin VB.TextBox txtTotalSTLoan 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   7155
               TabIndex        =   24
               Top             =   525
               Width           =   1560
            End
            Begin VB.TextBox txtTotalLTLoan 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   330
               Left            =   8745
               TabIndex        =   23
               Top             =   525
               Width           =   1560
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Active Employees"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   795
               TabIndex        =   34
               Top             =   300
               Width           =   1440
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Present Emps."
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   1
               Left            =   2400
               TabIndex        =   33
               Top             =   300
               Width           =   1170
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Absent Emps."
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   2
               Left            =   3990
               TabIndex        =   32
               Top             =   300
               Width           =   1110
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Monthly Salary"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   3
               Left            =   5565
               TabIndex        =   31
               Top             =   300
               Width           =   1215
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Temp. Loan"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   4
               Left            =   7170
               TabIndex        =   30
               Top             =   300
               Width           =   945
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Permanent Loan"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   5
               Left            =   8760
               TabIndex        =   29
               Top             =   300
               Width           =   1365
            End
         End
         Begin VB.TextBox txtOTHrs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   6135
            TabIndex        =   20
            Top             =   9000
            Width           =   1485
         End
         Begin VB.TextBox txtLateTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   4623
            TabIndex        =   18
            Top             =   9000
            Width           =   1485
         End
         Begin VB.TextBox txtLeaves 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   3097
            TabIndex        =   16
            Top             =   9000
            Width           =   1485
         End
         Begin VB.TextBox txtAbsents 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   1571
            TabIndex        =   14
            Top             =   9000
            Width           =   1485
         End
         Begin VB.TextBox txtTotalDays 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   420
            Left            =   45
            TabIndex        =   12
            Top             =   9000
            Width           =   1485
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5355
            TabIndex        =   8
            Top             =   3405
            Visible         =   0   'False
            Width           =   1000
         End
         Begin MSComctlLib.ListView LV 
            Height          =   7755
            Left            =   7650
            TabIndex        =   9
            Top             =   1695
            Width           =   4365
            _ExtentX        =   7699
            _ExtentY        =   13679
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   16711680
            BackColor       =   16761024
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Date"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "In"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Out"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "OT Hrs"
               Object.Width           =   1587
            EndProperty
         End
         Begin MSComctlLib.ListView LVSalHistory 
            Height          =   4275
            Left            =   45
            TabIndex        =   10
            Top             =   4500
            Width           =   2850
            _ExtentX        =   5027
            _ExtentY        =   7541
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   16711680
            BackColor       =   16761024
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Date"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Salary"
               Object.Width           =   2117
            EndProperty
         End
         Begin MSComctlLib.ListView LVDetails 
            Height          =   4100
            Left            =   2910
            TabIndex        =   11
            Top             =   1695
            Width           =   4725
            _ExtentX        =   8334
            _ExtentY        =   7223
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   16711680
            BackColor       =   16761024
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Object.Width           =   5292
            EndProperty
         End
         Begin MSComCtl2.DTPicker DT 
            Height          =   330
            Left            =   9195
            TabIndex        =   36
            Top             =   1305
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   582
            _Version        =   393216
            CustomFormat    =   "MMMM, yyyy"
            Format          =   148176899
            CurrentDate     =   40597
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   2925
            TabIndex        =   48
            Top             =   5760
            Width           =   660
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Notices"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   2925
            TabIndex        =   44
            Top             =   7620
            Width           =   615
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Latest Review"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   2925
            TabIndex        =   40
            Top             =   6720
            Width           =   1155
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Employee"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   60
            TabIndex        =   39
            Top             =   1095
            Width           =   810
         End
         Begin MSForms.CommandButton cmdRefresh 
            Height          =   420
            Left            =   10890
            TabIndex        =   38
            Top             =   1260
            Width           =   1125
            ForeColor       =   65280
            BackColor       =   4210752
            Caption         =   "Refresh        "
            PicturePosition =   327683
            Size            =   "1984;741"
            Accelerator     =   82
            FontName        =   "Calibri"
            FontEffects     =   1073741825
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Month"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   9165
            TabIndex        =   37
            Top             =   1065
            Width           =   555
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OT Hrs"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   10
            Left            =   6607
            TabIndex        =   21
            Top             =   8775
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Late Time"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   9
            Left            =   4960
            TabIndex        =   19
            Top             =   8775
            Width           =   810
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Leaves"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   8
            Left            =   3562
            TabIndex        =   17
            Top             =   8775
            Width           =   555
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Absents"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   7
            Left            =   1980
            TabIndex        =   15
            Top             =   8775
            Width           =   660
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Presents"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   6
            Left            =   330
            TabIndex        =   13
            Top             =   8775
            Width           =   720
         End
         Begin VB.Image Image1 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   2790
            Left            =   45
            Stretch         =   -1  'True
            Top             =   1695
            Width           =   2850
         End
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   285
      Left            =   5460
      TabIndex        =   3
      Top             =   10170
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEE DASHBOARD"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   12120
   End
End
Attribute VB_Name = "frmEmpDashBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bMouseDown As Boolean
Const lCaptionHeight = 600
Dim lDownX As Long, lDownY As Long
Dim iColNo As Integer
Dim iColNoTemp As Integer
Dim bSaved As Boolean
Dim lEntryID As Long

Private Sub DrawBorder(Pic As PictureBox)

    Dim lBorderColor As Long
    lBorderColor = RGB(220, 219, 244) 'RGB(0, 200, 0)    'RGB(69, 140, 201) 'RGB(69, 140, 201)
    Pic.Line (0, 0)-(Pic.Width, 0), lBorderColor 'Top
    Pic.Line (0, Pic.Height - 85)-(Pic.Width, Pic.Height - 85), lBorderColor 'Bottom
    Pic.Line (0, 0)-(0, Pic.Height), lBorderColor    'Left
    Pic.Line (Pic.Width - 85, 0)-(Pic.Width - 85, Pic.Height), lBorderColor  'Right
    
End Sub

Private Sub DrawPicCaptions()

    Dim i As Integer
    Dim lBorderColor As Long, lForeColor As Long
    For i = 0 To PicWC.count - 1
        If i = 1 Then
            lBorderColor = RGB(153, 204, 255)
            lForeColor = 0
        ElseIf i = 0 Then
            lBorderColor = RGB(220, 219, 244) 'RGB(255, 204, 153)
            lForeColor = 0
        ElseIf i = 3 Then
            lBorderColor = RGB(255, 153, 204)
            lForeColor = 0
        ElseIf i = 2 Then
            lBorderColor = RGB(0, 255, 0)
            lForeColor = 0
        ElseIf i = 4 Then
            lBorderColor = RGB(200, 200, 192)
            lForeColor = 0
        End If
        Call DrawCaptions(PicWC(i), lBorderColor, lForeColor)
    Next
    
End Sub

Private Sub DrawCaptions(Pic As PictureBox, lBorderColor As Long, lForeColor As Long)

    Dim lFillColor
    'lBorderColor = 0 'RGB(226, 103, 161) 'RGB(70, 70, 181)
    lFillColor = lBorderColor 'RGB(226, 103, 161) 'RGB(70, 70, 181)
    Pic.Line (0, 20)-(Pic.Width, 20), lBorderColor
    
    If Pic.Tag <> "" Then
        Pic.Line (0, 150)-(Pic.Width, 150), lBorderColor
        
        Pic.Line (0, 50)-(Pic.Width, 300), lFillColor, BF
        Pic.ForeColor = lForeColor 'vbGreen
        Pic.CurrentX = (Pic.Width - Pic.TextWidth(Pic.Tag)) / 2
        Pic.CurrentY = 0
        Pic.Print Pic.Tag
    End If
        
    'Pic.Line (0, Pic.Height - 30)-(Pic.Width, Pic.Height - 30), lBorderColor 'Bottom
    Pic.Line (0, Pic.Height - 50)-(Pic.Width, Pic.Height), lFillColor, BF
    
    Pic.Line (20, 0)-(20, Pic.Height), lBorderColor                          'Left
    Pic.Line (Pic.Width - 20, 0)-(Pic.Width - 30, Pic.Height), lBorderColor  'Right
    
    'Pic.Line (0, Pic.Height - 500)-(50, Pic.Height), lBorderColor
'    Me.BackColor = lBorderColor
End Sub


Private Sub cmdAbsents_Click()

    Dim myDT As Date
    myDT = DateAdd("d", -7, Date)
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\AbsentEmployees.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VEmpWithLastActivityDT.LastActivityDT}<=#" & myDT & "# AND {VEmp.Active}", rpt
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrintProdReport_Click()

    Dim f As Form
    
    Dim DtFrom As Date, DtTo As Date, strCond As String, strEmpID As String, strDeptID As String
'    Set f = New frmGetEmpIDAndDateRange
'    If f.getEmpIDAndDateRange(DtFrom, DtTo, strDeptID, strEmpID) = False Then Exit Sub
    DtFrom = DateSerial(year(Date), Month(Date), 1)
    DtTo = DateSerial(year(Date), Month(Date) + 1, 0)
    strEmpID = cmbEmp.ID
    
    Dim rpt As CRAXDDRT.Report, strSelection As String
    
    Set rpt = rptApp.OpenReport(RptPath & "\EmployeeProductivityReport.rpt")
    rpt.FormulaFields.GetItemByName("SubHeading").Text = "'" & Format(DtFrom, "dd-MMM-yyyy") & " to " & Format(DtTo, "dd-MMM-yyyy") & "'"
    strSelection = "{VEmpProductivityReport.EmpID}='" & strEmpID & "' AND {VEmpProductivityReport.DTOnly}=#" & DtFrom & "# TO #" & DtTo & "#"
    
    Dim frm As New frmPrevRpt
    frm.ShowReport strSelection, rpt
    
End Sub

Private Sub cmdRefresh_Click()

    On Error GoTo err
    
    If cmbEmp.MatchFound = False Then
        MsgBox "Please Select Employee.", vbInformation
        Exit Sub
    End If
    
    Dim rs As New ADODB.Recordset, ITM As ListItem, DTLastMonth As Date
    With rs
        .Open "SELECT EmpPic,EmpID,Name,Designation,DeptName,Phone1,Phone2,StartingSalary,JoinDate,Address FROM VEmp WHERE EmpID='" & cmbEmp.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            If Not IsNull(![EmpPic]) Then
                Set Image1.DataSource = rs
                Image1.DataField = "EmpPic"
            Else
                Image1.Picture = LoadPicture("")
            End If
            LVDetails.ListItems("EmpID").SubItems(1) = !EmpId & ""
            LVDetails.ListItems("EmployeeName").SubItems(1) = !Name & ""
            LVDetails.ListItems("Designation").SubItems(1) = !Designation & ""
            LVDetails.ListItems("Department").SubItems(1) = !DeptName & ""
            LVDetails.ListItems("PhoneNo1").SubItems(1) = !Phone1 & ""
            LVDetails.ListItems("JoiningDate").SubItems(1) = Format(!JoinDate, "dd-MMM-yyyy")
            LVDetails.ListItems("CurrentSalary").SubItems(1) = Format(Val(!StartingSalary & ""), "#,##")
            txtAddress = !Address & ""
            
            LVDetails.ListItems("Salary").SubItems(1) = Format(GetSingleLongValue("Total-OAmt", "MonthlySalaries", " WHERE EntryID IN(SELECT MAX(EntryID) FROM MonthlySalaries WHERE EmpID='" & cmbEmp.ID & "')"), "#,##")
            LVDetails.ListItems("OverTime").SubItems(1) = Format(GetSingleLongValue("OAmt", "MonthlySalaries", " WHERE EntryID IN(SELECT MAX(EntryID) FROM MonthlySalaries WHERE EmpID='" & cmbEmp.ID & "')"), "#,##")
            
            
        End If
        .Close
        
        Dim strStatus As String
        .Open "SELECT Attendance FROM AttendanceSheet WHERE EmpID='" & cmbEmp.ID & "' AND DT='" & Date & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            strStatus = "Absent"
        Else
            If !Attendance = 2 Then
                strStatus = "On Leave"
            ElseIf !Attendance = 1 Then
                strStatus = "Present"
            Else
                strStatus = "Absent"
            End If
        End If
        .Close
        LVDetails.ListItems("Status").SubItems(1) = strStatus
        
        Dim lTotalLoan As Long, lTotalCleared As Long, lLoanBalance As Long, lDedAmt As Long
        .Open "SELECT SUM(Amount) AS TotalLoan FROM Advances WHERE EmpID='" & cmbEmp.ID & "' AND Type=1", con, adOpenForwardOnly, adLockReadOnly
        lTotalLoan = Val(!TotalLoan & "")
        .Close
        .Open "SELECT SUM(AmtClrd) AS TotalCleared FROM AmtCleared WHERE EmpID='" & cmbEmp.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        lTotalCleared = Val(!TotalCleared & "")
        .Close
        lLoanBalance = lTotalLoan - lTotalCleared
        LVDetails.ListItems("LoanBalance").SubItems(1) = Format(lLoanBalance, "#,##")
        .Open "SELECT DAmount FROM Advances WHERE EntryID IN(SELECT MAX(EntryID) FROM Advances WHERE EmpID='" & cmbEmp.ID & "' AND Type=1)", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            lDedAmt = 0
        Else
            lDedAmt = Val(!dAmount & "")
        End If
        .Close
        LVDetails.ListItems("LoanRe-Payment").SubItems(1) = Format(lDedAmt, "#,##")
        .Open "SELECT SUM(Amount) AS TempLoan FROM Advances WHERE EmpID='" & cmbEmp.ID & "' AND Type=0 AND MONTH(DT)=" & Month(Date) & " AND YEAR(DT)=" & year(DT), con, adOpenForwardOnly, adLockReadOnly
        LVDetails.ListItems("TempLoan").SubItems(1) = Format(Val(!TempLoan & ""), "#,##")
        .Close
        
        Dim strGName As String, strGPhone As String, strGCell As String, strKnownBy As String
        
        .Open "SELECT strFullName,strPhoneNo,strCellNo,strKnownBy FROM Employees INNER JOIN GuarantorList_Imported ON Employees.intGuarantorID=GuarantorList_Imported.intGuarantorId WHERE Employees.EmpID='" & cmbEmp.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            strGName = ""
            strGPhone = ""
            strGCell = ""
            strKnownBy = ""
        Else
            strGName = !strFullName & ""
            strGPhone = !strPhoneNo & ""
            strGCell = !strCellNo & ""
            strKnownBy = !strKnownBy & ""
        End If
        .Close
        
        LVDetails.ListItems("Guarantor").SubItems(1) = strGName
        LVDetails.ListItems("GuarantorPhone").SubItems(1) = strGPhone
        LVDetails.ListItems("GuarantorCell").SubItems(1) = strGCell
        LVDetails.ListItems("GuarantorKnownBy").SubItems(1) = strKnownBy
        
        Dim strReveiew As String, strReviewDT As String
        .Open "SELECT * FROM EmpReviews WHERE EntryID IN(SELECT MAX(EntryID) FROM EmpReviews WHERE EmpID='" & cmbEmp.ID & "')", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            strReveiew = !Review & ""
            strReviewDT = Format(!DT, "dd-MMM-yyyy")
        Else
            strReveiew = ""
            strReviewDT = ""
        End If
        .Close
        txtReview = strReviewDT & " - " & strReveiew
        
        Dim strNotice As String, strNoticeDT As String
        .Open "SELECT * FROM EmpNotices WHERE EmpID='" & cmbEmp.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        txtNotices = ""
        Do Until .EOF
            txtNotices = Format(!DT, "dd-MMM-yyyy") & " " & !Notice & ""
            .MoveNEXT
        Loop
        .Close
'        LVDetails.ListItems("LatestReview").SubItems(1) = strReveiew
'        LVDetails.ListItems("ReviewDate").SubItems(1) = strReviewDT
    
        .Open "SELECT * FROM VSalaryLedger WHERE EmpID='" & cmbEmp.ID & "' ORDER BY DT ASC", con, adOpenForwardOnly, adLockReadOnly
        LVSalHistory.ListItems.Clear
        Do Until .EOF
            Set ITM = LVSalHistory.ListItems.add(, , Format(!DT, "dd-MMM-yyyy"))
            ITM.ForeColor = vbWhite
            ITM.ListSubItems.add , , Format(Val(![Salary] & ""), "#,##")
            .MoveNEXT
        Loop
        .Close
        
        Dim myDT As Date, LastDT As Date
        myDT = DateSerial(DT.year, DT.Month, 1)
        LastDT = DateSerial(DT.year, DT.Month + 1, 0)
        LV.ListItems.Clear
        While myDT <= LastDT
            Set ITM = LV.ListItems.add(, Format(myDT, "dd-MMM-yyyy"), Format(myDT, "dd-dddd"))
            If Weekday(myDT) = 1 Then
                ITM.ForeColor = RGB(20, 150, 20)
            ElseIf GetSingleLongValue("COUNT(*)", "Holidays", " WHERE DT='" & myDT & "'") > 0 Then
                ITM.ForeColor = RGB(20, 150, 20)
            Else
                ITM.ForeColor = vbWhite
            End If
            myDT = DateAdd("d", 1, myDT)
        Wend
        
        .Open "SELECT DT,FirstInTime,SecondOutTime,OTHrs FROM VEmpTimes2 WHERE EmpID='" & cmbEmp.ID & "' AND MONTH(DT)=" & DT.Month & " AND YEAR(DT)=" & DT.year, con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            Set ITM = LV.ListItems(Format(!DT, "dd-MMM-yyyy"))
            
            ITM.ListSubItems.add , , !FirstInTime & ""
            ITM.ListSubItems.add , , !SecondOutTime & ""
            ITM.ListSubItems.add , , Round(Val(!OTHrs & ""), 1)
            .MoveNEXT
        Loop
        .Close
        
        
    End With
    Set rs = Nothing
    
    Dim lTotalPresents As Long, lTotalSundays As Long, lTotalHolidays As Long, lTotalLeaves As Long
    lTotalPresents = GetSingleLongValue("COUNT(Attendance)", "AttendanceSheet", " WHERE EmpID='" & cmbEmp.ID & "' AND MONTH(DT)=" & DT.Month & " AND YEAR(DT)=" & DT.year & " AND Attendance=1")
    lTotalLeaves = GetSingleLongValue("COUNT(EmpID)", "Leaves", " WHERE EmpID='" & cmbEmp.ID & "' AND MONTH(DT)=" & DT.Month & " AND YEAR(DT)=" & DT.year)
    lTotalHolidays = GetSingleLongValue("COUNT(DT)", "Holidays", " WHERE MONTH(DT)=" & DT.Month & " AND YEAR(DT)=" & DT.year)
    lTotalSundays = TotalSundays(DT.Month, DT.year)
    
    'lTotalPresents = lTotalPresents + lTotalHolidays + lTotalSundays
    
    txtTotalDays = lTotalPresents
    
    If Date < DateSerial(DT.year, DT.Month + 1, 0) Then
        txtAbsents = Day(Date) - (lTotalPresents + lTotalHolidays + TotalSundays(DT.Month, DT.year, Day(Date)))
    Else
        txtAbsents = LV.ListItems.count - (lTotalPresents + lTotalHolidays + lTotalSundays)
    End If
    
    txtLeaves = lTotalLeaves
    
    Dim dTotalLateHrs As Double
    dTotalLateHrs = GetSingleDoubleValue("SUM(LateHours)", "AttendanceSheet", " WHERE EmpID='" & cmbEmp.ID & "' AND MONTH(DT)=" & DT.Month & " AND YEAR(DT)=" & DT.year)
    txtLateTime = getHrsMin(dTotalLateHrs)
    'txtPayHrs = getHrsMin(Round(GetSingleDoubleValue("SUM(PayableHrs)", "EmpTimes", " WHERE ISNULL(OverTime,0)=0 AND EmpID='" & cmbEmp.ID & "' AND MONTH(DT)=" & DT.Month & " AND YEAR(DT)=" & DT.year), 4))
    
    txtOTHrs = getHrsMin(Round(GetSingleDoubleValue("SUM(PayableHrs)", "EmpTimes", " WHERE ISNULL(OverTime,0)=1 AND EmpID='" & cmbEmp.ID & "' AND MONTH(DT)=" & DT.Month & " AND YEAR(DT)=" & DT.year), 4))
        
    Call UpdateFactoryDetails
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub UpdateFactoryDetails()
    txtTotalEmps = GetSingleLongValue("COUNT(EmpID)", "Employees", " WHERE Active=1")
    txtTotalPresents = GetSingleLongValue("COUNT(Attendance)", "AttendanceSheet", "WHERE DT='" & Date & "' AND Attendance=1")
    txtTotalAbsents = Val(txtTotalEmps) - Val(txtTotalPresents)
    
    txtTotalSalary = Format(GetSingleLongValue("SUM(StartingSalary)", "Employees", " WHERE Active=1"), "#,##")
    txtTotalSTLoan = Format(GetSingleLongValue("SUM(Amount)", "Advances", " WHERE Type=0 AND MONTH(DT)=" & Month(Date) & " AND YEAR(DT)=" & year(Date)), "#,##")
    txtTotalLTLoan = Format(GetSingleLongValue("SUM(Amount)", "Advances", " WHERE Type=1") - GetSingleLongValue("SUM(AmtClrd)", "AmtCleared"), "#,##")
End Sub
Private Function GetTotal() As Double

    Dim i As Integer
    Dim dTotalAmt As Double, dTotalRate As Double, dTotalStRate As Double
    For i = 1 To LV.ListItems.count
        dTotalStRate = dTotalStRate + Val(LV.ListItems(i).SubItems(2))
        dTotalRate = dTotalRate + Val(LV.ListItems(i).SubItems(3))
        dTotalAmt = dTotalAmt + Val(LV.ListItems(i).SubItems(5))
        
    Next
    txtTotal = Format(dTotalAmt, "#,##0.00")
    txtTotalRate = dTotalRate
    txtTotalSTRate = dTotalStRate
    
End Function
    

Private Sub cmdSalaryLedger_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    
    Dim myDTFrom As Date, myDTTo As Date, Cond As String
    'myDTFrom = DateAdd("m", -6, Date)
    myDTTo = Date
     
    If Month(Date) > 6 Then
        myDTFrom = DateSerial(year(Date), 7, 1)
    Else
        myDTFrom = DateSerial(year(Date) - 1, 7, 1)
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\EmpSalaryLedger.rpt")
    Dim f As New frmPrevRpt
    'f.ShowReport "{VEmp.EmpID}='" & cmbEmp.id & "' AND {MonthlySalaries.DT}=#" & myDTFrom & "# TO #" & myDTTo & "#", rpt
    f.ShowReport "{VEmp.EmpID}='" & cmbEmp.ID & "'", rpt
    
End Sub

Private Sub Form_Load()

    Call DrawPicCaptions
    bMouseDown = False
    bSaved = False
    lEntryID = 0
    
End Sub

Private Sub PicWC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y < 600 Then
        If Button = 1 Then
            If bMouseDown = False Then
                bMouseDown = True
                lDownX = X
                lDownY = Y
            End If
        End If
    End If
End Sub

Private Sub PicWC_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bMouseDown Then
        Me.Move Me.Left + ((X - lDownX) / 15), Me.Top + ((Y - lDownY) / 15)
    End If
End Sub

Private Sub PicWC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    bMouseDown = False
End Sub

Public Function ShowMe() As Boolean
    
    'Will hold the handle to the new region.
    Dim hRegion As Long
    '
    'This API can create a Round or Oval Shaped Region.
'    hRegion = CreateRoundRectRgn(0, 0, Me.Width / 15, Me.Height / 15, 20, 20)
    '
    'Change the shape of the Form to the newly created Region.
'    SetWindowRgn Me.hwnd, hRegion, True
    '
    'Show the new form with the specified shape.
    'frmRegion.Show
    
    DT.Value = Date
    cmbEmp.ListHeight = 5800
    
    Call cmbEmp.AddVals(con, "'{ ' + EmpID + ' } ' + Name + ' {' + ISNULL(Designation,'') + '} (' + DeptName + ')' + CASE WHEN Active=0 THEN '(InActive)' ELSE '' END", "VEmp", "EmpID", " ORDER BY EmpID")
    
    LVDetails.ListItems.add(, "EmpID", "Employee Code").ForeColor = vbWhite
    LVDetails.ListItems.add(, "EmployeeName", "Employee Name").ForeColor = vbWhite
    LVDetails.ListItems.add(, "Designation", "Designation").ForeColor = vbWhite
    LVDetails.ListItems.add(, "Department", "Department").ForeColor = vbWhite
    LVDetails.ListItems.add(, "PhoneNo1", "Phone No.").ForeColor = vbWhite
    LVDetails.ListItems.add(, "JoiningDate", "Joining Date").ForeColor = vbWhite
    LVDetails.ListItems.add(, "CurrentSalary", "CurrentSalary").ForeColor = vbWhite
    LVDetails.ListItems.add(, "Salary", "Last Salary").ForeColor = vbWhite
    LVDetails.ListItems.add(, "OverTime", "Last Over Time").ForeColor = vbWhite
    
    LVDetails.ListItems.add(, "Status", "Status").ForeColor = vbWhite
    LVDetails.ListItems.add(, "LoanBalance", "Loan Balance").ForeColor = vbWhite
    LVDetails.ListItems.add(, "LoanRe-Payment", "Loan Re-Payment").ForeColor = vbWhite
    LVDetails.ListItems.add(, "TempLoan", "Temp Loan").ForeColor = vbWhite
    
    
    LVDetails.ListItems.add(, "Guarantor", "Guarantor").ForeColor = vbWhite
    LVDetails.ListItems.add(, "GuarantorPhone", "Phone").ForeColor = vbWhite
    LVDetails.ListItems.add(, "GuarantorCell", "Cell").ForeColor = vbWhite
    LVDetails.ListItems.add(, "GuarantorKnownBy", "Known By").ForeColor = vbWhite
'    LVDetails.ListItems.add , "", ""
'    LVDetails.ListItems.add , "", ""
'    LVDetails.ListItems.add , "", ""
'    LVDetails.ListItems.add , "", ""
'    LVDetails.ListItems.add(, "LatestReview", "Latest Review").ForeColor = vbWhite
'    LVDetails.ListItems.add(, "ReviewDate", "Date - Review").ForeColor = vbWhite
    
    Call UpdateFactoryDetails
    Me.Show
    ShowMe = bSaved
    
End Function

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub
