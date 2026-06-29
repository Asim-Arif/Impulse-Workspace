VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmDCEmp 
   Caption         =   "Discontinue  ..........."
   ClientHeight    =   7920
   ClientLeft      =   2625
   ClientTop       =   2190
   ClientWidth     =   10575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   10575
   Begin VB.CheckBox chkPayThroughSalarySheet 
      Caption         =   "Pay Salary Through Salary Sheet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   75
      TabIndex        =   79
      Top             =   7305
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Frame Frame1 
      Height          =   7275
      Left            =   30
      TabIndex        =   4
      Top             =   30
      Width           =   10485
      Begin VB.CheckBox chkExternal 
         Caption         =   "External Salary Sheet"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         TabIndex        =   82
         Top             =   0
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.TextBox txtTotalHrs 
         Height          =   285
         Left            =   150
         TabIndex        =   80
         Top             =   165
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.TextBox txtNetPay 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8925
         TabIndex        =   68
         Top             =   4050
         Width           =   1440
      End
      Begin VB.TextBox txtNetPayable 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   8925
         TabIndex        =   65
         Top             =   6060
         Width           =   1440
      End
      Begin ComboList.Usercontrol1 cmbDept 
         Height          =   285
         Left            =   1500
         TabIndex        =   5
         Top             =   675
         Width           =   2300
         _ExtentX        =   4048
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
      End
      Begin ComboList.Usercontrol1 cmbEmp 
         Height          =   285
         Left            =   1500
         TabIndex        =   6
         Top             =   1080
         Width           =   6180
         _ExtentX        =   10901
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
      End
      Begin VB.Frame Frame7 
         Caption         =   "Other Receivable :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2460
         Left            =   6405
         TabIndex        =   47
         Top             =   4710
         Width           =   2475
         Begin VB.TextBox txtOthers 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1260
            TabIndex        =   71
            Top             =   990
            Width           =   1110
         End
         Begin VB.TextBox txtTotalReceivable 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   480
            TabIndex        =   53
            Top             =   2040
            Width           =   1600
         End
         Begin VB.TextBox txtLBalance 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1245
            TabIndex        =   51
            Top             =   630
            Width           =   1110
         End
         Begin VB.TextBox txtNoticePay2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1245
            TabIndex        =   49
            Top             =   285
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Others :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   17
            Left            =   510
            TabIndex        =   72
            Top             =   1020
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Total Receivables"
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
            Index           =   26
            Left            =   525
            TabIndex        =   52
            Top             =   1800
            Width           =   1500
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Loan Bal :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   25
            Left            =   495
            TabIndex        =   50
            Top             =   660
            Width           =   705
         End
         Begin MSForms.CheckBox chkNoticePay2 
            Height          =   285
            Left            =   75
            TabIndex        =   48
            Top             =   270
            Width           =   1230
            BackColor       =   -2147483633
            ForeColor       =   10488870
            DisplayStyle    =   4
            Size            =   "2170;503"
            Value           =   "0"
            Caption         =   "Notice Pay :"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Other Payables :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2460
         Left            =   480
         TabIndex        =   32
         Top             =   4725
         Width           =   5880
         Begin VB.TextBox txtOtherPayable 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4125
            TabIndex        =   77
            Top             =   1560
            Width           =   1605
         End
         Begin VB.TextBox txtTotalPayable 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4125
            TabIndex        =   63
            Top             =   2085
            Width           =   1600
         End
         Begin VB.TextBox txtOvertimeDay 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1500
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   1260
            Width           =   1600
         End
         Begin VB.TextBox txtOvertimeAmt 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4125
            Locked          =   -1  'True
            TabIndex        =   54
            Top             =   1260
            Width           =   1600
         End
         Begin VB.TextBox txtNoticePay 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1500
            TabIndex        =   46
            Top             =   1605
            Width           =   1600
         End
         Begin VB.TextBox txtBLYear 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4125
            TabIndex        =   43
            Top             =   915
            Width           =   1600
         End
         Begin VB.TextBox txtBonus10c 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1500
            TabIndex        =   42
            Top             =   915
            Width           =   1600
         End
         Begin VB.TextBox txtGAmount 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4125
            TabIndex        =   39
            Top             =   570
            Width           =   1600
         End
         Begin VB.TextBox txtGPeriod 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1500
            TabIndex        =   38
            Top             =   570
            Width           =   1600
         End
         Begin VB.TextBox txtELAmt 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4125
            TabIndex        =   36
            Top             =   225
            Width           =   1600
         End
         Begin VB.TextBox txtELeaves 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1500
            MaxLength       =   2
            TabIndex        =   34
            Top             =   225
            Width           =   1600
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Others :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   30
            Left            =   3495
            TabIndex        =   78
            Top             =   1605
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "Total Payables"
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
            Index           =   24
            Left            =   4290
            TabIndex        =   64
            Top             =   1875
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Overtime Hrs :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   10
            Left            =   390
            TabIndex        =   57
            Top             =   1305
            Width           =   1050
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "OT Amount :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   11
            Left            =   3180
            TabIndex        =   56
            Top             =   1305
            Width           =   915
         End
         Begin MSForms.CheckBox chNoticePay1 
            Height          =   285
            Left            =   210
            TabIndex        =   45
            Top             =   1605
            Width           =   1230
            BackColor       =   -2147483633
            ForeColor       =   10488870
            DisplayStyle    =   4
            Size            =   "2170;503"
            Value           =   "0"
            Caption         =   "Notice Pay :"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Bonus L Y :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   23
            Left            =   3300
            TabIndex        =   44
            Top             =   960
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Bonus (10C) :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   22
            Left            =   450
            TabIndex        =   41
            Top             =   960
            Width           =   990
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "G. Amount :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   21
            Left            =   3225
            TabIndex        =   40
            Top             =   615
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Graduity Period :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   20
            Left            =   225
            TabIndex        =   37
            Top             =   615
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "E. L. Amt :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   19
            Left            =   3330
            TabIndex        =   35
            Top             =   270
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Earned Leaves :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   18
            Left            =   270
            TabIndex        =   33
            Top             =   270
            Width           =   1170
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Total Deducation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1350
         Left            =   6405
         TabIndex        =   30
         Top             =   3360
         Width           =   2490
         Begin VB.TextBox txtDeducation 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   330
            TabIndex        =   31
            Top             =   585
            Width           =   1845
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Deducation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1350
         Left            =   480
         TabIndex        =   23
         Top             =   3360
         Width           =   5880
         Begin VB.TextBox txtPF 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1500
            Locked          =   -1  'True
            TabIndex        =   74
            Top             =   945
            Width           =   1600
         End
         Begin VB.TextBox txtOtherDed 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4125
            TabIndex        =   73
            Top             =   945
            Width           =   1605
         End
         Begin VB.TextBox txtTAdvance 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4125
            Locked          =   -1  'True
            TabIndex        =   69
            Top             =   285
            Width           =   1600
         End
         Begin VB.TextBox txtRent 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4125
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   630
            Width           =   1605
         End
         Begin VB.TextBox txtEOBI 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1500
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   630
            Width           =   1600
         End
         Begin VB.TextBox txtTax 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1500
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   285
            Width           =   1600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "P.F. :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   29
            Left            =   990
            TabIndex        =   76
            Top             =   990
            Width           =   405
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Others :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   28
            Left            =   3495
            TabIndex        =   75
            Top             =   990
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "T. Adv:"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   14
            Left            =   3555
            TabIndex        =   70
            Top             =   330
            Width           =   540
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Flat Rent :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   16
            Left            =   3270
            TabIndex        =   29
            Top             =   675
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "EOBI :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   15
            Left            =   945
            TabIndex        =   28
            Top             =   675
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tax :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   13
            Left            =   1035
            TabIndex        =   25
            Top             =   330
            Width           =   375
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Total Payable"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   6420
         TabIndex        =   21
         Top             =   2610
         Width           =   2490
         Begin VB.TextBox txtTPayable 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   330
            TabIndex        =   22
            Top             =   270
            Width           =   1845
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Employee Work Detail"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   495
         TabIndex        =   18
         Top             =   2610
         Width           =   5865
         Begin VB.TextBox txtDWork 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1500
            Locked          =   -1  'True
            TabIndex        =   60
            Top             =   300
            Width           =   1600
         End
         Begin VB.TextBox txtWorkAmt 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   4125
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   300
            Width           =   1600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Days Worked :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   8
            Left            =   420
            TabIndex        =   61
            Top             =   330
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Amount :"
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   9
            Left            =   3255
            TabIndex        =   20
            Top             =   330
            Width           =   660
         End
      End
      Begin VB.TextBox txtSalary 
         Height          =   285
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   2205
         Width           =   2300
      End
      Begin MSComCtl2.DTPicker DTP 
         Height          =   285
         Left            =   5400
         TabIndex        =   13
         Top             =   1440
         Width           =   2300
         _ExtentX        =   4048
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dddd, dd-MMM-yyyy"
         Format          =   171048963
         CurrentDate     =   38894
      End
      Begin VB.TextBox txtDuration 
         Height          =   285
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1845
         Width           =   2300
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1500
         TabIndex        =   8
         Top             =   1440
         Width           =   2300
         _ExtentX        =   4048
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dddd, dd-MMM-yyyy"
         Format          =   171048963
         CurrentDate     =   38829
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   5400
         TabIndex        =   58
         Top             =   1845
         Width           =   2300
         _ExtentX        =   4048
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dddd, dd-MMM-yyyy"
         Format          =   171048963
         CurrentDate     =   38894
      End
      Begin MSForms.CommandButton cmdShow 
         Height          =   345
         Left            =   5400
         TabIndex        =   67
         Top             =   2205
         Width           =   2295
         Caption         =   "Refresh"
         PicturePosition =   327683
         Size            =   "4048;609"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "NET PAYABLE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   27
         Left            =   9000
         TabIndex        =   66
         Top             =   5805
         Width           =   1290
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "GROSS PAYABLE"
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
         Left            =   8925
         TabIndex        =   62
         Top             =   3780
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Termination Date :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   12
         Left            =   3960
         TabIndex        =   59
         Top             =   1890
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "TERMINATION AND DISCONTINUATION OF EMPLOYEE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   150
         TabIndex        =   17
         Top             =   165
         Width           =   10260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   " Salary :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   15
         Top             =   2235
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Discontinue Date :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   3975
         TabIndex        =   14
         Top             =   1485
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Department :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   12
         Top             =   675
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Emp Code :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   11
         Top             =   1110
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date of Joining :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   10
         Top             =   1515
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Duration :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   225
         TabIndex        =   9
         Top             =   1875
         Width           =   720
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1875
         Left            =   8235
         Stretch         =   -1  'True
         Top             =   675
         Width           =   1650
      End
   End
   Begin VB.CheckBox chkPaySlip 
      Caption         =   "Payment Slip"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3090
      TabIndex        =   3
      Top             =   7515
      Width           =   1380
   End
   Begin MSComctlLib.ListView LV 
      Height          =   5820
      Left            =   0
      TabIndex        =   81
      Top             =   0
      Visible         =   0   'False
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   10266
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
      NumItems        =   28
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "Emp ID"
         Text            =   "Emp ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "Emp Name"
         Text            =   "Emp Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Object.Tag             =   "B. Sal"
         Text            =   "B. Sal"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Object.Tag             =   "Rate/Hr"
         Text            =   "Rate/Day"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Object.Tag             =   "A. Hrs"
         Text            =   "A.Days"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Object.Tag             =   "A. Amt"
         Text            =   "A. Amt"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Object.Tag             =   "A.Allow"
         Text            =   "A.Allow"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Object.Tag             =   "Allow.Amt"
         Text            =   "Allow.Amt"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Object.Tag             =   "S. Hrs"
         Text            =   "S.Days"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Object.Tag             =   "S. Amt"
         Text            =   "S. Amt"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Object.Tag             =   "Leave Hrs"
         Text            =   "Leaves"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Object.Tag             =   "Lv. Amt."
         Text            =   "Lv. Amt."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Object.Tag             =   "OT Hrs"
         Text            =   "OT Hrs"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Object.Tag             =   "OT Amt"
         Text            =   "OT Amt"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Object.Tag             =   "L.Hrs"
         Text            =   "L.Hrs"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   15
         Object.Tag             =   "L.Amt"
         Text            =   "L.Amt"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   16
         Object.Tag             =   "Total"
         Text            =   "Total"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   17
         Object.Tag             =   "T. Amt"
         Text            =   "T. Amt"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   18
         Object.Tag             =   "Net Ttl"
         Text            =   "Net Ttl"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   19
         Object.Tag             =   "ST Loan"
         Text            =   "ST Loan"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   20
         Object.Tag             =   "Loan Ded."
         Text            =   "Loan Ded."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   21
         Object.Tag             =   "Adv Sal"
         Text            =   "Adv Sal"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   22
         Object.Tag             =   "U. Fund"
         Text            =   "U. Fund"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   23
         Object.Tag             =   "Fine"
         Text            =   "Fine"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   24
         Object.Tag             =   "Scholorship"
         Text            =   "Scholorship"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   25
         Object.Tag             =   "Lunch Ded."
         Text            =   "Lunch Ded."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   26
         Object.Tag             =   "EOBI"
         Text            =   "EOBI"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   27
         Object.Tag             =   "Balance"
         Text            =   "Balance"
         Object.Width           =   1764
      EndProperty
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   360
      Left            =   4515
      TabIndex        =   2
      Top             =   7425
      Width           =   1995
      Caption         =   "Print     "
      PicturePosition =   327683
      Size            =   "3519;635"
      Accelerator     =   80
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   360
      Left            =   6510
      TabIndex        =   0
      Top             =   7425
      Width           =   1995
      Caption         =   "Make Discontinue"
      PicturePosition =   327683
      Size            =   "3519;635"
      Accelerator     =   68
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   360
      Left            =   8505
      TabIndex        =   1
      Top             =   7425
      Width           =   1995
      Caption         =   "Cancel    "
      PicturePosition =   327683
      Size            =   "3528;635"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmDCEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LongTerm As Boolean
Dim rsLedger As New ADODB.Recordset
Dim iTotalHrs As Integer

Private Sub PrintReport()

    Screen.MousePointer = vbHourglass
    Dim lID As Long
    lID = Val(con.Execute("Select Max(EntryID) From Advances").Fields(0).Value & "")
    
    con.Execute "Delete From PrintTempAdv"
    
    con.Execute "Insert Into PrintTempAdv(EmpID,BSal,Rate,SDays,SAmt,Total,Paid,LongTerm,Balance,DT,LunchAmt,AmtPaid,AdvancesID) Values('" & _
     cmbEmp.ID & "'," & Val(txtSalary) & "," & Val(txtRate) & "," & Val(txtHrsWorked) & "," & Val(txtAmount) & "," & _
     Val(txtAmt) & "," & Val(txtPrevST) & "," & Val(txtLongTerm) & "," & Val(txtBalance) & ",'" & DTPicker1 & "'," & Val(txtLunchAmt) & "," & Val(txtBal) & "," & lID & ")"
    
    Dim f As New frmPrevRpt, rpt As New CRAXDDRT.Report
    
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptTempAdv.rpt")
    
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
    
    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & company & "'"
        ElseIf FormulaField.Name = "{@FromTo}" Then
            FormulaField.Text = "'" & Format(DTPicker1, "dd-MMM-yyyy") & "'"
        ElseIf FormulaField.Name = "{@MonthDays}" Then
            FormulaField.Text = Val(txtWorkingHrs) 'lWHrs * 8 'Val(txtTotalHrs)
        End If
    Next
    
    Me.Hide
    f.ShowReport " ", rpt
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub chkNoticePay2_Click()

    If chNoticePay1.Value = True Then
        chkNoticePay2.Value = False
    End If
    If chkNoticePay2.Value = True Then
        txtNoticePay2.Text = txtSalary.Text
        txtTotalReceivable.Text = Val(txtLBalance.Text) + Val(txtNoticePay2.Text)
    ElseIf chNoticePay1.Value = False Then
        txtTotalReceivable.Text = Val(txtTotalReceivable.Text) - Val(txtNoticePay2.Text)
        txtNoticePay2.Text = 0
    End If
    txtNetPayable.Text = Val(txtTotalPayable.Text) - Val(txtTotalReceivable.Text)
End Sub

Private Sub chNoticePay1_Click()
    
    If chNoticePay1.Value = True Then
        txtNoticePay.Text = txtSalary.Text
        txtTotalPayable.Text = TotalPayable
    ElseIf chNoticePay1.Value = False Then
        txtTotalPayable.Text = Val(txtTotalPayable.Text) - (txtNoticePay.Text)
        txtNoticePay.Text = 0
    
    End If
    txtNetPayable.Text = Val(txtTotalPayable.Text) - Val(txtTotalReceivable.Text)
End Sub

Private Sub cmbDept_matched()

    Dim strCondition As String
    cmbEmp.ClearVals
    
    If cmbDept.ID = "0" Then
        strCondition = " WHERE Active=0 AND DiscontinuedOn IS NULL"
    Else
        strCondition = " WHERE DeptID='" & cmbDept.ID & "' AND Active=0 AND DiscontinuedOn IS NULL"
    End If
    strCondition = strCondition & " ORDER BY EmpID"
    
    cmbEmp.AddVals con, "'{' + EmpID + '} ' + Name", "Employees", "EmpID", strCondition
    
End Sub

Private Function Saved() As Boolean

    Saved = False
    'On Error GoTo ERR
    Dim SQL As String
    
    If Not cmbEmp.MatchFound Then
        MsgBox "Please Select Valid Employee From The List.", vbInformation
        Exit Function
    ElseIf Val(txtBal) = 0 Then
        MsgBox "Invalid Amount.", vbInformation
        Exit Function
    ElseIf Val(txtBal) > Val(txtMax) - Val(txtPrevST) Then
        MsgBox "Sorry, Not Allowed Greater Than Rs. " & (Val(txtMax) - Val(txtPrevST)) & ".", vbInformation
        Exit Function
    End If
    '''''''''''''''''''''''''''''''''
    ' Remove This Check at the moment
    
    
'    If LongTerm = True Then
'        If IsNotLongApplicable Then
'            Exit Function
'        End If
'    End If
'
    
    
    '''''''''''''''''''''''''''''''''
    ' Remove This Check at the moment
    

    
    
    
    
'    Dim AmtAllowed As Double
'    AmtAllowed = GetShortAmt(cmbEmp.ID, DTPicker1.value)
'    If Val(txtAmt) > AmtAllowed Then
'      MsgBox "Cannot Pay More Than Rs. " & Format(AmtAllowed, "0.00"), vbInformation
'      Exit Function
'    End If
'
    Dim strVchrNo As String
    Call StartTrans(con)
    
    strVchrNo = AddVoucher
    
    SQL = "Insert Into Advances(EmpID,DT,Type,Description," & _
     "Amount,DAmount,AmountCleared,AccVoucherNo) Values('" & cmbEmp.ID & _
     "','" & DTPicker1 & "',0,'Short Term'," & Val(txtBal) & ",0,0,'" & _
     strVchrNo & "')"
    
    con.Execute SQL
    
    con.CommitTrans
    
    Saved = True
    If chkPrint = 1 Then Call PrintReport
    SQL = ""
    Exit Function
err:
    MsgBox err.Description


End Function

Private Function AddVoucher() As String

    Dim i As Long, INextSno As Double, Bal As Double, Amt As Double
    Dim VchrNo As String, theDesc As String, AccNo As String
    
    Dim Dated As Date
    Dated = DTPicker1
    'Charge All Amount to Cash
    Amt = Val(txtBal)
    
    If Amt = 0 Then Exit Function
    
    'ConAcc.BeginTrans
     
    'Get The Required Values
     
    INextSno = getNextSno(Dated)
    VchrNo = getNextVchrNo(Dated, "CPV")
     
    'Give The Description
    theDesc = "Temp. Advance Paid To [" & cmbEmp.ID & "] " & cmbEmp.Text
     
    ''''''''''''''''Charge To  Cash Entry'''''''''''''
    
    AccNo = CashInHand
    
    Bal = getBalance(AccNo, Dated)
    
    Bal = Bal - Amt
        
    'Credit Entry
    con.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
    
    
    Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
     
     
     INextSno = INextSno + 1
     
    '''''''''''''''''' D E B I T   E N T R I E S '''''''''''''
      
      'Determine the Account no for Staff Member
      'Long / Short Tem Advance
      
      AccNo = "15-015-20001" 'IIf(LongTerm, LongAdvAccNo, ShortAdvAccNo)
     
      Bal = getBalance(AccNo, Dated)
      
      Amt = Amt
      Bal = Bal + Amt
    
    
      'Debit To Staff Long/Short Advance Account
      
      con.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & theDesc & "'," & Amt & ",0," & Bal & ",'',0)"
        
     
    Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
    
    con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & VchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate & "')"
    
    AddVoucher = VchrNo
    
End Function

Private Sub cmbEmp_matched()

    Image1.Picture = LoadPicture("")
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select DeptID,Sex,JoinDate,EmpPic,StartingSalary FROM Employees WHERE EmpID='" & cmbEmp.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        DTPicker1 = !JoinDate
        
'        cmbMain.ID = !MainUnit
'        cmbDept.ID = !DeptID
        
        If !Sex = "Female" Then
            bFemale = True
        Else
            bFemale = False
        End If
        
        If Not IsNull(![EmpPic]) Then
            Set Image1.DataSource = rs
            Image1.DataField = "EmpPic"
        End If
        txtSalary.Text = Val(!StartingSalary & "")
        .Close
    End With
    Set rs = Nothing
    txtDuration.Text = DateDiffInText(DTPicker1.Value)
  
End Sub

Private Sub GetLunchAmt()

    On Error GoTo err
    
    Dim DT As Date
    DT = DateSerial(DTPicker1.year, DTPicker1.Month, 1)
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select LunchAmt From EmpLunchAmt Where EmpID='" & cmbEmp.ID & "' AND DT='" & DT & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            txtLunchAmt = 0
            txtLunchAmt.Tag = "Add"
        Else
            txtLunchAmt = Val(!LunchAmt & "")
            txtLunchAmt.Tag = "Edit"
        End If
        .Close
    End With
    
    Exit Sub
    
err:
    MsgBox err.Description
End Sub

Private Function FinalizedSalarySheetLoaded() As Boolean

    Dim rs As New ADODB.Recordset
    Dim Cond As String
    Dim IsExernalSalary As Byte
    Dim GrandTotal As Double
    
    IsExernalSalary = True
    
    Cond = "Where EmpID='" & cmbEmp.ID & "'"
    
    If Cond = "" Then
        Cond = "Where"
    Else
        Cond = Cond & " And "
    End If
    
    Cond = Cond & " Month(DT)=" & DTPicker1.Month & _
    " and Year(DT)=" & DTPicker1.year
    
    With rs
        .Open "Select * From VEmpSettlements " & Cond, con, adOpenForwardOnly, adLockReadOnly
        
        If .RecordCount = 0 Then
            FinalizedSalarySheetLoaded = False
            GoTo END_OF_FUNCTION
        End If
            
            LV.ListItems.Clear
            Do Until .EOF
        
                Set ITM = LV.ListItems.add(, ![EmpId], ![EmpId] & "")
                
'                If IsComputed Then
'                    ITM.Checked = IsPayed
'                    ITM.Bold = IsTransfered
'                End If
                
'                If Val(txtTotalHrs) <> Val(![TotalMonthHrs] & "") Then
'                    txtTotalHrs = ![TotalMonthHrs] & ""
'                    MsgBox "Salary Sheet Is Finalized According To " & txtTotalHrs & " Working Hours.", vbInformation
'                End If
                
            
                ITM.Tag = "" 'Val(![Emptype] & "")
                
                ITM.ListSubItems.add(, , ![Name] & "").Tag = Val(![PrevLTLoan] & "")
                ITM.ListSubItems.add , , ![BSal] & ""
                
                ITM.ListSubItems(2).Tag = ![CasualLeaves] & "," & ![SickLeaves] & "," & ![CompensatoryLeaves] & _
                "," & ![AnnualLeaves] & "," & ![WPLeaves] & "," & ![MaternityLeaves]
                
                ITM.ListSubItems.add(, , ![Rate] & "").Tag = ![HrsPerDay]
                ITM.ListSubItems.add(, , ![ADays] & "").Tag = Val(![SDays] & "")
                ITM.ListSubItems.add , , ![AAmt] & ""
                ITM.ListSubItems.add , , ![AAllow] & ""
                ITM.ListSubItems.add , , ![AAllowAmt] & ""
                ITM.ListSubItems.add , , IIf(chkExternal = vbChecked, getHrsMin(![SDays]), ![SDays] & "")
                ITM.ListSubItems.add , , ![SAmt] & ""
                ITM.ListSubItems.add , , ![Leaves] & ""
                ITM.ListSubItems.add , , ![LeaveAmt] & ""
                ITM.ListSubItems.add , , ![OHrs] & ""
                ITM.ListSubItems.add , , ![OAmt] & ""
                ITM.ListSubItems.add , , ![lHrs] & ""
                ITM.ListSubItems.add , , ![lAmt] & ""
                ITM.ListSubItems.add , , ![total] & ""
                ITM.ListSubItems.add , , ![Tax] & ""
                ITM.ListSubItems.add , , ![NetTtl] & ""
                ITM.ListSubItems.add , , ![ShortTerm] & ""
                ITM.ListSubItems.add , , ![LongTerm] & ""
                ITM.ListSubItems.add , , ![AdvSal] & ""
                ITM.ListSubItems.add , , ![UnionFund] & ""
                ITM.ListSubItems.add , , ![Fine] & ""
                ITM.ListSubItems.add , , ![Bonus] & ""
                ITM.ListSubItems.add , , ![Lunch] & ""
                ITM.ListSubItems.add , , ![EOBI] & ""
                ITM.ListSubItems.add , , ![Balance] & ""
                
                'GrandTotal = GrandTotal + Val(ITM.SubItems(27))
                
                .MoveNEXT
            Loop
        
        
    End With
    
    'lblTotal = Format(GrandTotal, "#,##0.00")
    
    FinalizedSalarySheetLoaded = True
    'cmdUpdate.Visible = False
END_OF_FUNCTION:

    If rs.State = 1 Then rs.Close
    Set rs = Nothing
    
End Function

Private Sub cmbMainDept_matched()

    
End Sub

Private Sub cmbEmpName_matched()
    
    cmbEmp.ID = cmbEmpName.ID
     txtDuration.Text = DateDiffInText(DTPicker1.Value)
End Sub

Private Sub cmbMain_matched()

    cmbEmp.ClearVals
    cmbEmpName.ClearVals
    
    With cmbDept
        .ClearVals
        .AddItem "< All Departments >", "0"
        If cmbMain.ID = "0" Then
            .AddVals con, "Name", "Departments", "DeptID", " Where SubDeptOf Is Not Null"
        Else
            .AddVals con, "Name", "Departments", "DeptID", " Where SubDeptOf='" & cmbMain.ID & "'"
        End If
        .ID = "0"
    End With
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim SQL As String
    DT = DateSerial(DTPicker1.year, DTPicker1.Month, 1)
'    con.Execute "Delete From EmpLunchAmt Where EmpID='" & cmbEmp.ID & "' AND DT='" & DT & "'"
'    con.Execute "Insert Into EmpLunchAmt(EmpID,DT,LunchAmt) Values('" & cmbEmp.ID & "','" & DT & "'," & Val(txtLunchAmt) & ")"

    
    
    Call StartTrans(con)
    
    Call SaveIt
'    Call AddSalaryVouchers
        
    con.Execute "UPDATE Employees SET Active=0,DiscontinuedOn='" & DTP & "' WHERE EmpID='" & cmbEmp.ID & "'"
    
     'SQL = "INSERT INTO Advances(EmpID,DT,Type,Description,Amount) VALUES('" & cmbEmp.ID & _
      "','" & DTP & "',0,'Short Term'," & Val(LV.ListItems(1).SubItems(27)) & ")"
      
'    If chkPayThroughSalarySheet.Value <> vbChecked Then
'
'        SQL = "INSERT INTO Advances(EmpID,DT,Type,Description,Amount) VALUES('" & cmbEmp.ID & _
'         "','" & DTP & "',0,'Short Term'," & Val(txtNetPayable) & ")"
'
'        con.Execute SQL

'    End If
    
     
    con.CommitTrans
    
    Call cmbEmp_matched
    Unload Me
    
    
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    
    Call InsertRecords
    
    Dim rpt As CRAXDDRT.Report
    Dim SelFor As String
    
    If chkPaySlip.Value = vbChecked Then
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptSettPaySlip.rpt")
        SelFor = "{VPaySlip.EmpID}='" & cmbEmp.ID & "'"
    Else
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptFinalSett.rpt")
        SelFor = "{Employees.EmpID}='" & cmbEmp.ID & "'"
    End If
    
    If chkPaySlip.Value = vbChecked Then
    
        Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
        Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
        
        Set FormulaFields = rpt.FormulaFields
        For Each FormulaField In FormulaFields
            If FormulaField.Name = "{@Company}" Then
                FormulaField.Text = "'" & company & "'"
            ElseIf FormulaField.Name = "{@FromTo}" Then
                FormulaField.Text = "' For " & Format(DTPicker1, "MMM") & "-" & DTPicker1.year & "'"
            ElseIf FormulaField.Name = "{@MonthDays}" Then
                FormulaField.Text = iTotalHrs
            End If
        Next
    End If
    
    Dim frm As New frmPrevRpt
    frm.ShowReport SelFor & " ", rpt
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdShow_Click()

    If DTPicker2.Value < DTP.Value Then
        MsgBox "Termination date must be greater the discontinued Date !", vbInformation
        Exit Sub
    End If
    If cmbEmp.Text = "" Then
        MsgBox "Please select the Employee !", vbInformation
        Exit Sub
    End If
    'Call RefreshLedger
    Call RefreshSalary
    Call UpdateFormNow
    LV.Visible = False
        
End Sub


Private Sub txtAmt_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = OnlyNumber(KeyAscii.Value, True)
End Sub

Private Sub txtDeduct_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = OnlyNumber(KeyAscii.Value, True)
End Sub

'To Check The Long Term Settings That Are Made For Departments
Private Function IsNotLongApplicable() As Boolean

Dim Rec As New ADODB.Recordset
'dim Item as New

With Rec
    .CursorLocation = adUseClient
    .Open "SELECT JoinDate,LongTermTime,LongTermLimit FROM VOnlyEmps WHERE EmpID='" & cmbEmp.ID & "'", con, adOpenKeyset, adLockReadOnly
    
    If DateDiff("m", !JoinDate, DTPicker1) < !longtermtime Then
        MsgBox "Could Not Give Long Term Loan To Employee." & vbNewLine & "It Can Be Given After " & !longtermtime & " Months", vbInformation
        IsNotLongApplicable = True
    ElseIf Val(txtAmt) > !LongTermLimit Then
        MsgBox "Could Not Give Long Term Loan To Employee. Because Amount " & vbNewLine & " Is Greater Than Long Term Limit.", vbInformation
        IsNotLongApplicable = True
    End If
    
End With


End Function

Private Sub RefreshLedger_Old()
    
    Dim iMonth As Integer, iYear As Integer
    iMonth = DTP.Month
    iYear = DTP.year
    
    Dim rs As New ADODB.Recordset
    Dim rsInternal As New ADODB.Recordset
    Dim iTotalMonthDays As Integer
    
    With rs
        rsInternal.Open "SELECT EmpID,LongTerm,PrevLTLoan FROM MonthlySalaries WHERE Month(DT)=" & iMonth & " AND Year(DT)=" & iYear & " AND IsForSA=0", con, adOpenStatic, adLockReadOnly
        
    End With
    
    Dim TotalSalDays As Double
    Dim TDate As Date
    Dim SDate As Date 'For Saving
    Dim Hdays As Long
    
    SDate = DateSerial(iYear, iMonth, 1)
    TDate = DateSerial(iYear, iMonth, Day(GetServerDate))
    
    TotalSalDays = DateAdd("m", 1, SDate) - SDate

    MonthDays = TotalSalDays - TotalSundays(iMonth, iYear, TotalSalDays)
    
    If chkExternal = vbUnchecked Then
        Hdays = Val(con.Execute("Select count(DT) from Holidays Where month(DT)=" & _
        iMonth & " and Year(DT)=" & Val(iYear)).Fields(0))
        MonthDays = MonthDays '- Hdays
    End If
   
    iTotalMonthDays = TotalMonthDays
    
    Dim UnionFund As Integer
    Dim GrandTotal As Double
    Dim EOBIAmt As Byte
    
    Dim SunDates() As Date
        
    Dim EmpId As String
   
    Dim TempTotalSaldays As Double
    Dim LastDay As Integer
    
    Dim ITM As ListItem
    Dim Cond As String
    Dim SQL As String
    Dim cmd As New ADODB.Command
    
    
    Dim SalHrs As Double
    Dim HrsMultiple As Double
    
    
    Dim Leaves As Double
    Dim ApprovedLeaves As Double
    Dim UnApprovedLeaves As Double
    
    EmpId = cmbEmp.ID
    
    Dim rsCheck As New ADODB.Recordset
        
    Dim LunchDays As Long
    
    
    Dim OTByQty As Boolean
    
    
    'TDate = DateSerial(iyear, imonth, Day(GetServerDate))
        
   
    ''''''''''''''''''''' This Will Refresh The LateHours '''''''
    '            con.Execute "EXEC SP_CalcOverTime '" & SDate & "'", "'" & cmbEmp.ID & "'"
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    LastDay = TotalSalDays
    
    Dim JDate As Date
    
    If TDate = GetServerDate Then
        TotalSalDays = Day(GetServerDate)
    ElseIf GetServerDate < TDate Then
        MsgBox "Can't Show Entries Of Future Months.", vbInformation
        Exit Sub
    End If
    
    TempTotalSaldays = TotalSalDays     'Store For Later Use
    
    JDate = DateSerial(iYear, iMonth, TotalSalDays)
        
    'Sql = "Select * from VAdvanceSalary " & Cond & " JoinDate<='" & Format(JDate, "dd-MMM-yyyy") & "' Order By DeptID,EmpID"
    
    cmd.CommandType = adCmdStoredProc
    
    cmd.CommandText = "VAdvanceSalary"
    
    Set cmd.ActiveConnection = con
    cmd.Parameters.Refresh
    
    cmd.Parameters(1).Type = adDBTimeStamp
    cmd.Parameters(1).Value = JDate

    cmd.Parameters(2).Type = adVarChar
    cmd.Parameters(2).Value = EmpId
    
    
    If rsLedger.State = 1 Then rsLedger.Close
    rsLedger.CursorType = adOpenForwardOnly
    rsLedger.LockType = adLockReadOnly
    
    Set rsLedger = cmd.Execute
    
    With rsLedger
        
        '.Open Sql, Con, adOpenForwardOnly, adLockReadOnly
        Dim TotalSuns As Integer
        Dim SalDays As Double
        Dim Rate As Double
        Dim AbsAllowed As Integer
        Dim AbsentDays As Double
        
        Dim total As Double
        Dim OverTimeHours As Double
        Dim OverTimeHoursSA As Double
        Dim LateHours As Double
        Dim TaxPer As Double
        Dim DedAmt As Double
        Dim Payed As Double
        Dim LongLoan As Double
        Dim TaxSal As Double
        Dim IsComputed As Boolean
        
        Dim NTime As Integer
        Dim LTime As Double
        Dim DT As String
        Dim IsPayed As Boolean
        Dim IsTransfered As Boolean
        Dim Holidays As Integer
        
        Dim EmpSalary As Double
        Dim EmpExtSal As Double
                
        LV.ListItems.Clear
        LV.Visible = False
        
   For i = 0 To .RecordCount - 1
                    
            'Reset Total Salary days For Each Employee
            
            TotalSalDays = DateAdd("m", 1, SDate) - SDate
                
            'If TDate = GetServerDate Then
             TotalSalDays = Day(DTP)
            'End If
            
            'Check that Employee Discontinued In This Month
                        
            If Not IsNull(![DiscontinuedOn]) Then
                If Month(![DiscontinuedOn]) = Month(JDate) And ![DiscontinuedOn] <= JDate Then
                  TotalSalDays = Day(![DiscontinuedOn])
                End If
            Else
               'TotalSalDays = TempTotalSaldays
            End If
            
            Leaves = Val(!Leaves & "")
                                   

           'if Joined This Month Calculate From Join Date
           
           If Format(![JoinDate], "MMM-yyyy") = Format(JDate, "MMM-yyyy") Then
               
               rs.Open "Select count(DT) from Holidays Where DT Between '" & ![JoinDate] & "' and '" & JDate & "'", con, adOpenForwardOnly, adLockReadOnly
               
                  If rs.RecordCount > 0 Then
                     Holidays = rs.Fields(0) & ""
                  End If
               rs.Close
               Set rs = Nothing
               
               SundayAbsents = getSundayAbsents(![EmpId] & "", Val(Month(DTP)), year(DTP), Val(![WeeklyAbs] & ""), Int(TotalSalDays))
               TotalSuns = TotalSundays(DTP.Month, year(DTPicker1), TotalSalDays, Day(![JoinDate]))
               Holidays = TotalSuns + Holidays - SundayAbsents
               
            Else
            
                 rs.Open "Select count(DT) from Holidays Where convert(char(3),DT,7)='" & cmbMonth & "' and Year(DT)=" & year(DTPicker1) & " and day(dt)<" & TotalSalDays, con, adOpenForwardOnly, adLockReadOnly
                  
                  If rs.RecordCount > 0 Then
                     Holidays = rs.Fields(0) & ""
                  End If
                  
                 rs.Close
               Set rs = Nothing

               'SundayAbsents = getSundayAbsents(![EmpId], dtp.month, dtp.year, Int(TotalSalDays), ![WeeklyAbs])
               
               TotalSuns = TotalSundays(Month(DTP), year(DTP), TotalSalDays)
                                              
               Holidays = TotalSuns + Holidays - SundayAbsents
               

            End If
                    
            If chkExternal = vbChecked Then
                'Exclude The Holidays From Calculation
                TotalSalDays = TotalSalDays - Holidays
            End If
            
            
            'Get Sundays Not To Be Paid

            'Call Computed(![EmpId], IsComputed, TaxPer, AbsAllowed, OverTimeHours, NTime, LTime, LongLoan, DT, IsPayed, IsTransfered, JDate, EmpSalary, EmpExtSal, ![Active], HrsMultiple)
            
            If Not IsComputed Then
                'Tax %
                'If IsNull(![PerDeduct]) Then TaxPer = 0 Else TaxPer = Val(![PerDeduct]) & ""
                
                'Absent Days....
                AbsAllowed = Val(![AbsAllowed] & "")
    
                ' Reset Orver Time Hours
                OverTimeHours = 0
                OverTimeHoursSA = 0
                
                'Late Hours
                NTime = Val(![NTimes] & "")
                LTime = Val(![TimeAllowed] & "")
                
                'Long Term Loan Deduction
                Dim lDedAmt As Long, lClsAmt As Long, lTotAmt As Long
                If IsNull(![DedAmt]) Then lDedAmt = 0 Else lDedAmt = ![DedAmt]
                If IsNull(![totamt]) Then lTotAmt = 0 Else lTotAmt = ![totamt]
                If IsNull(![ClsAmt]) Then lClsAmt = 0 Else lClsAmt = ![ClsAmt]
                
                If lDedAmt > (lTotAmt - lClsAmt) Then
                    LongLoan = (lTotAmt - lClsAmt)
                Else
                    LongLoan = lDedAmt
                End If
                                      
                'Get The Salary Depending On Salary Sheet Type
                'EmpSalary = IIf(chkExternal = vbUnchecked, Val(![StartingSalary] & ""), Val(![Salary] & ""))
                
                 EmpSalary = Val(![PresentSalary] & "")
                 'EmpExtSal = Val(![Salary] & "")
                 EmpExtSal = EmpSalary
            End If
            
            
            
            'Check If Employee is paid Over time per piece
'            If Not IsNull(![OTByQty]) Then If ![OTByQty] Then OTByQty = True
                            
'                If Not IsNull(![Contractor]) Then
'                    If ![Contractor] Then               ' Calculate Over Time On the Basis of Qty Worked If Applicable
'                        OTByQty = True
'                        If Val(![QtyPerHour] & "") > 0 Then
'
'                            OverTimeHours = Val(![OTQtyWorked] & "")
'
'                        Else
'                            OverTimeHours = 0
'                        End If
'                    Else
'                        If IsNull(![OverTimeHours]) Then
'                            OverTimeHours = 0
'                            OverTimeHoursSA = 0
'                        Else
'                            OverTimeHours = ![OverTimeHours]
'                            OverTimeHoursSA = ![OverTimeHours]
'                        End If
'                    End If
'
'                Else
'
'                    OTByQty = False
'
'
'                    If IsNull(![OverTimeHours]) Then
'                        OverTimeHours = 0
'                        OverTimeHoursSA = 0
'                    Else
'                        OverTimeHours = ![OverTimeHours]
'                        OverTimeHoursSA = ![OverTimeHours]
'                    End If
'
'                End If
            '*********************
'            OverTimeHours = 0
'            OverTimeHoursSA = 0
            '*********************
            OTByQty = False
'
'
            If IsNull(![OverTimeHours]) Then
                OverTimeHours = 0
                OverTimeHoursSA = 0
            Else
                OverTimeHours = ![OverTimeHours]
                OverTimeHoursSA = ![OverTimeHours]
            End If
            Dim theKey As String
            
            If IsComputed Then
                theKey = ![EmpId] & Format(SDate, "dd-MMM-yyyy")
            Else
                theKey = ![EmpId] & "'"
            End If
            
            Set ITM = LV.ListItems.add(, theKey, ![EmpId] & "")
            
            If IsComputed Then
                ITM.Checked = IsPayed
                ITM.Bold = IsTransfered
            End If
            
            
'            ITM.Tag = Val(![Emptype] & "")
            
'            If IsInternalFinal And chkExternal.value = vbChecked Then
'                rsInternal.MoveFirst
'                rsInternal.Find "EmpID='" & ![EmpId] & "'", , adSearchForward, 0
'                If rsInternal.EOF Then
'                    ITM.ListSubItems.add(, , ![Name] & "").Tag = 0
'                Else
'                    ITM.ListSubItems.add(, , ![Name] & "").Tag = Val(rsInternal![PrevLTLoan] & "")
'                End If
'            Else
'                ITM.ListSubItems.add(, , ![Name] & "").Tag = Val(![TotAmt] & "") - Val(![clsamt] & "")
'            End If
            
            ITM.ListSubItems.add(, , ![Name] & "").Tag = Val(![totamt] & "") - Val(![ClsAmt] & "")
            
            ITM.ListSubItems.add , , IIf(chkExternal = vbChecked, EmpExtSal, EmpSalary & "")
            
            
            ITM.ListSubItems(2).Tag = ![CasualLeaves] & "," & ![SickLeaves] & "," & ![AnnualLeaves] & _
            "," & ![SpecialLeaves] & "," & ![CompensatoryLeaves] & "," & ![WPLeaves] & "," & ![MaternityLeaves]
            
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '   Rate = Format(Val(EmpSalary) / LastDay, "0.00")
            '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '   Calculate the Rate According To Total No ........
            '   Of Working Days in This Month for External Salary
            '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            
            If chkExternal = vbChecked Then
                
                
                'External Salary Rate Calculated By Working Days
                'Rate = Format(Val(EmpSalary) / (LastDay - Holidays), "0.00")
                Rate = Format(Val(EmpSalary) / (LastDay - TotalSuns), "0.00")
                
                'Multiple Factor for Rate Per hour
                'HrsMultiple = 8
                
                'If Not IsComputed Then HrsMultiple = Val(![WorkingHrs] & "")
                
                
                ' If Working hours are not defined then consider it 8 Hours Per day.
                If HrsMultiple = 0 Then HrsMultiple = 8
            Else
                
                'Rate = Format(Val(EmpSalary) / TempTotalSaldays, "0.00")
                
                
                Rate = Format(Val(EmpSalary) / LastDay, "0.00")
                
                'Multiple Factor for Rate Per Day
                HrsMultiple = 1
                
            End If
            
            
'            Rate = Rate / HrsMultiple
            
            
            
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '   Rate Formula For October Salary Sheet ...
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            If chkExternal Then Rate = EmpSalary / Val(txtTotalHrs)
                
            
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            
            
            'Store The Total Working hours for This Employee
       
            ITM.ListSubItems.add(, , Format(Rate, "0.00")).Tag = HrsMultiple
            
            
            If IsNull(![presentDays]) And !DeptID <> "VDRS" Then
                SalDays = 0
                'IF No Presents Found Exclude This Employee
                LV.ListItems.Remove ITM.key
                GoTo NEXT_EMPLOYEE
            Else
                
                
                SalDays = ((Val(![presentDays] & "")) + Holidays) + AbsAllowed
                
                'Store for Later Calculation
                
                'LunchDays = Val(![presentDays])
                
                
                
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'Exclude the Holidays now (for Vision External Salary Sheet Only)
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                
                
                
                If chkExternal = vbChecked Then SalDays = ((Val(![presentDays] & ""))) + AbsAllowed
                
                
                'If Salary Days Exceed Total Salary Days Set Them To Salary Days
                If SalDays > TotalSalDays Then SalDays = TotalSalDays
                    
            End If
            
            Dim DaysToExclude As Integer
            
            If Format(![JoinDate], "MMM-yyyy") = Format(JDate, "MMM-yyyy") Then
               DaysToExclude = Day(![JoinDate]) - 1
            Else
               DaysToExclude = 0
            End If
            
            
            'Deduct Penalty Days
            
            'If chkExternal = vbUnchecked Then
            '    SalDays = SalDays - Val(![AbsentDays] & "")
            'End If
            
            If SalDays < 0 Then SalDays = 0
            
            
            
            If chkExternal Then
                Leaves = Val(!Leaves & "")
                UnApprovedLeaves = 0
            Else
                Leaves = Val(!ApprovedLeaves & "")
                UnApprovedLeaves = Val(!Leaves & "") - Val(!ApprovedLeaves & "")
            End If
            

            
            '''''''''''''''''''''''''''''''
            ' dont do anything with sunday/holiday related absents at the moment
            ''''''''''''''''''''''''''''''''
            ' SalDays = SalDays - CalcSpecialAbs(!EmpId, JDate, !BGHoliday, !AGHoliday, !BPHoliday, !APHoliday)
            '''''''''''''''''''''''''''''''
            
            SalDays = SalDays - UnApprovedLeaves
            
            If Abs(![DeductAbsents]) = 0 Then SalDays = TotalSalDays - DaysToExclude
            
            AbsentDays = TotalSalDays - SalDays - DaysToExclude
            
            
            ' Make Absents 0 if in Negative
            If AbsentDays < 0 Then AbsentDays = 0
            
            
            'If IsNull(![AbsentDays]) Then AbsentDays = 0 Else AbsentDays = Val(![AbsentDays]) / 2
            'AbsentDays = IIf(AbsentDays > 0, AbsentDays - AbsAllowed, AbsentDays)
            

            ' Convert To Hours if External Salary
       
                ITM.ListSubItems.add , , AbsentDays * HrsMultiple
       
            '''''''''''''''''' Save the Salary Hours in it tag Later '''''''''''''''''''''''
            
            
            ITM.SubItems(5) = Round(AbsentDays * HrsMultiple * Rate)
            
            'If Val(!Leaves & "") > 0 Then
            '    If AbsentDays - AbsAllowed >= 0 Then
                
            '    itm.SubItems(11) = IIf((Exempt), "0", (Leaves * Rate))
                
            '    End If
            'Else
            '    itm.SubItems(11) = "0"
            'End If
            
            ITM.SubItems(11) = "0"
            
            ' Calculat Net Absents Allowed
            ITM.SubItems(6) = AbsAllowed
            Dim NetAbsAllowed As Integer
            
            ' Calculate The Allowed Amount That is Subtracted From The Absent Amount
            
            If AbsentDays = 0 Then
                NetAbsAllowed = 0
            Else
                NetAbsAllowed = IIf(AbsAllowed - AbsentDays <= 0, AbsAllowed, AbsAllowed - AbsentDays)
            End If
            
            'Absent Allowed Amount
            ITM.SubItems(7) = IIf((Exempt), "0", Round(Rate * NetAbsAllowed * HrsMultiple, 2))
            
            
            SalHrs = SalDays * HrsMultiple
            
            
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '                      Fix the Salary hours of Directors
            '               and Give Relaxation of 3 hours for october salary
            '                       Add the Leave hours as defined
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            
            
            If chkExternal Then
                If ![DeptID] = "VDRS" Then
                    SalHrs = Val(txtTotalHrs)
                Else
             
                    SalHrs = Val(![HrsWorked] & "") + ((Holidays - TotalSuns) * 8)
                    
                    ' Add All the Leaves to Worked hours
                    SalHrs = SalHrs + (Val(![Leaves] & "") * 8)
                    
                    'Remove Full Hours of Without Pay leaves
                    SalHrs = SalHrs - (Val(![WPLeaves] & "") * 8)
                    
                    ' Remove half Hours of Sick Leaves
                    SalHrs = SalHrs - (Val(![SickLeaves] & "") * 4)
                    
                    ' Check if Salary Hours are greater than Total Hours Due to relaxation
                    ' Fix them to maximum hours
                    
                    If SalHrs > Val(txtTotalHrs) Then SalHrs = Val(txtTotalHrs)
                End If
            Else
                If ![DeptID] = "VDRS" Then
                    SalHrs = Val(txtTotal)
                Else
                    
                    'Remove Full Hours of Without Pay leaves
                    SalHrs = SalHrs - Val(![WPLeaves] & "")
                    
                    'Remove half Hours of Sick Leaves
                    SalHrs = SalHrs - Val(![SickLeaves] & "") / 2
                    
                    ' Check if Salary Hours are greater than Total Hours Due to relaxation
                    ' Fix them to maximum hours
                    
                    'If SalHrs > Val(txtTotal) Then SalHrs = Val(txtTotal)
                    
                End If
                
            End If
            
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '       Show in Time Format if external salary sheet
            '       and Show days in internal salarysheet instead
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            ITM.SubItems(8) = IIf(chkExternal, getHrsMin(SalHrs), SalHrs)
            txtDWork.Text = SalHrs
            ITM.SubItems(9) = Round((SalHrs * Rate), 0)
            txtWorkAmt.Text = Round((SalHrs * Rate), 0)
            txtTPayable.Text = Val(txtWorkAmt)
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '       Save the Salary Hours in Listsubitmes(4).tag To Save
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ITM.ListSubItems(4).Tag = SalHrs
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            ITM.SubItems(10) = Leaves * HrsMultiple
            
            If chkExternal = vbUnchecked Then
                LeaveDedAmt = Val(![SickLeaves] & "") * HrsMultiple * Rate / 2
                LeaveDedAmt = LeaveDedAmt + Val(![WPLeaves] & "") * HrsMultiple * Rate
                ITM.SubItems(11) = Round(LeaveDedAmt, 2)
            End If
            ITM.SubItems(11) = 0
            
            'If SalDays = 0 Then SalDays = 0 Else SalDays = (SalDays - AbsentDays)
            
            If OTByQty Then
                ITM.SubItems(12) = OverTimeHours
                'Over time By Quantity : Per Day Rate / Sheets Per Day * 1.4
                ITM.SubItems(13) = Round(OverTimeHours * (Rate * HrsMultiple / 56 * 1.4))
            Else
                
                ITM.SubItems(12) = IIf(chkExternal, MinsToTime(OverTimeHoursSA * 60), MinsToTime(OverTimeHours * 60))
                
                If chkExternal Then
                    ITM.SubItems(13) = Round(OverTimeHoursSA * (Rate * 2))
                Else
                    ITM.SubItems(13) = Round(OverTimeHours * (Rate / 8 * 1.3333))
                End If
                
            End If
            
            
            'Itm.SubItems(13) = Round(OverTimeHours * (Rate * 8 / Val(![OverTime] & "")))
            'Itm.SubItems(12) = Val(!Leaves & "") * 8
            'Itm.SubItems(13) = Val(!CurrentLong & "")
            'LateHours = GetLateHours(![EmpId], NTime, LTime, JDate)
            
            
'            If ![DeptID] = "VFNC" Or ![DeptID] = "VSLS" Or ![EmpId] = "VGLV-0001" Then
'                LateHours = Val(![Hours] & "") - LTAllowed
'            Else
'                LateHours = 0
'            End If
            
            LateHours = 0 'Val(![Hours] & "") - LTAllowed
            
            If LateHours < 0 Then LateHours = 0
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''
            '           Exempt For October 2004
            ''''''''''''''''''''''''''''''''''''''''''''''''''
            'Well It was just written as LateHours=0 And That was why it was'nt deducting LT.
            'So I Just Commented it Out.
            'LateHours = 0
            ''''''''''''''''''''''''''''''''''''''''''''''''''
                        
            ITM.SubItems(14) = MinsToTime(LateHours * 60)
            
           ' itm.SubItems(15) = IIf((Exempt), "0", Format((LateHours * (Rate / Val(!WorkingHrs & ""))), "0.00"))
            
            If chkExternal Then
                ITM.SubItems(15) = IIf((Exempt), "0", Format((LateHours * Rate), "0.00"))
            Else
                ITM.SubItems(15) = IIf((Exempt), "0", Format((LateHours * Rate / 8), "0.00"))
            End If
             
             
            'Total = Format(Val((ITM.SubItems(9)) + Val(ITM.ListSubItems(7))) - (Val(ITM.SubItems(15)) + Val(ITM.SubItems(11))), "0.00")
            total = Format(Val((ITM.SubItems(9)) + Val(ITM.ListSubItems(7))) - (Val(ITM.SubItems(11))), "0.00")
            ITM.SubItems(16) = total
            
            'TaxPer = Val(![PerDeduct])
            ITM.SubItems(17) = Val(!TaxAmt & "") 'Round(getTaxAmt((EmpSalary * 12) / 1.5) / 12, 0)
            txtTax.Text = Round(Val(!TaxAmt & "") / 30 * SalDays, 0)
            'DedAmt = Round((Val(![Salary]) / TotalSalDays) * (TaxPer / 100))
            'TaxSal = Round((Val(EmpSalary) / LastDay) * TotalSalDays)
            DedAmt = 0 'Round(TaxSal * (TaxPer / 100))
            
            'Tax Percentag That Is Zero Length
            'itm.SubItems(18) = "0"
            
            'itm.SubItems(14) = DedAmt
            ITM.SubItems(18) = total 'Total - Val(ITM.SubItems(17))
            If IsNull(![Payed]) Then Payed = 0 Else Payed = ![Payed]
            ITM.SubItems(19) = Round(Payed / 10) * 10
            txtTAdvance.Text = Round(Payed / 10) * 10
            'If chkExternal.value = vbChecked And IsInternalFinal Then
             '   If rsInternal.EOF Then
              '      ITM.SubItems(20) = 0
               ' Else
                    'ITM.SubItems(20) = Val(rsInternal![LongTerm] & "")
                'End If
            'Else
                ITM.SubItems(20) = LongLoan
            'End If
            
            ITM.SubItems(21) = Val(![AdvSalpaid] & "")
            
            
            'Dim SocialAmt As Double
            'SocialAmt = IIf((!Social = True), !SocialAmt, 0)
            
            
            ''''''''''''''''''''''''''''''''''''''''''
            '       Union Fund deduction
            ''''''''''''''''''''''''''''''''''''''''''
            
'                If Not IsNull(![UnionFund]) Then
'                    If ![UnionFund] Then UnionFund = 20 Else UnionFund = 0
'                Else
'                    UnionFund = 0
'                End If
                UnionFund = 0
            ''''''''''''''''''''''''''''''''''''''''''
                Dim LunchDed As Double
                
'                If Not IsNull(![Lunch]) Then
'                    LunchDed = Val(![LunchAmt] & "")
'                Else
'                    LunchDed = 0
'                End If
                LunchDed = Val(!RentDeduction & "")
            ''''''''''''''''''''''''''''''''''''''''''
            
            ITM.SubItems(22) = UnionFund
            ITM.SubItems(23) = Val(!TotalFine & "")
            
            ' Use This Column For Scholorships Instead Of Bonus ...
            'ITM.SubItems(24) = Val(!TotalBonus & "")
            
            ITM.SubItems(24) = Val(![ScholorshipAmt] & "")
            ITM.SubItems(25) = LunchDed
            
            'If Not IsNull(![OldAge]) Then If ![OldAge] Then EOBIAmt = 20
            EmpAge = DateDiff("M", !DOB, JDate) / 12
            
            If EmpAge <= 60 And !EOBI Then
                EOBIAmt = 60
            Else
                EOBIAmt = 0
            End If
            '/\/\//\/\ IF U CHANGE EOBI FORMULA HERE, REMEMBER TO CHANGE IN rptSalarySheet and rptPaySlip TOO./\/\/\/\/\
            
            ITM.SubItems(26) = EOBIAmt
            txtEOBI.Text = EOBIAmt
            txtDeducation.Text = Val(txtTax) + Val(txtEOBI) + Val(txtRent) + (txtTAdvance)
            txtNetPay.Text = (txtTPayable.Text) - (txtDeducation.Text)
            'Removed OverTime From Balance
            ' ITM.SubItems(27) = Round((Val(ITM.SubItems(18)) + Val(!ScholorshipAmt & "") _
            - (Payed + Val(!TotalFine & "") + UnionFund + LongLoan + LunchDed + Val(!AdvSalpaid & "") + Val(ITM.SubItems(26))) _
            + Val(ITM.SubItems(13))))
            
            ITM.SubItems(27) = Round((Val(ITM.SubItems(18)) + Val(!ScholorshipAmt & "") _
            - (Payed + Val(!TotalFine & "") + UnionFund + LongLoan + LunchDed + Val(!AdvSalpaid & "") + Val(ITM.SubItems(17)) + Val(ITM.SubItems(26)))), 0)
            
            
            GrandTotal = GrandTotal + Val(ITM.SubItems(27))
            
            Dim LastDate As Date
            LastDate = DateSerial(year(DTPicker1), Month(DTPicker1), LastDay)
            
            Dim GPeriod, gYears, gMonths As Double
            GPeriod = DateDiffInMonths(DTPicker1.Value)
            gYears = DateDiffInYears(DTPicker1.Value)
            gMonths = GPeriod Mod 12
            txtGPeriod.Text = gYears & " Years " & gMonths & " Months "
            txtGAmount.Text = Round(Val(gYears) * txtSalary.Text, 0)
            If gMonths >= 6 Then
                txtGAmount = Round(txtGAmount + (Val(txtSalary) / 12 * gMonths), 0)
            End If
            If rs.State Then rs.Close
            rs.Open "select prevLTLoan from vpayslip WHERE empid='" & Left(theKey, Len(theKey) - 1) & "'", con, adOpenForwardOnly, adLockReadOnly
            If rs.RecordCount > 0 Then
                txtLBalance.Text = rs.Fields(0)
                LongLoan = rs.Fields(0)
            Else
                txtLBalance.Text = 0
                LongLoan = 0
            End If
            If DateDiffInMonths(DTPicker1) > 3 Then
            
                txtBonus10c.Text = txtSalary.Text
            Else
                txtBonus10c = 0
            End If
            
            txtTotalPayable.Text = TotalPayable
            txtTotalReceivable.Text = Val(txtNoticePay2.Text) + LongLoan
            txtNetPayable.Text = TotalPayable - Val(txtTotalReceivable)
            If (IsComputed) Then
                'cmdUpdate.Visible = UserHasAccess("FinalSalSheet")          'And (chkExternal = vbUnchecked)
            ElseIf Not (IsComputed) And (GetServerDate > LastDate) Then
                'cmdUpdate.Visible = UserHasAccess("FinalSalSheet")          'And (chkExternal = vbUnchecked)
            Else
               ' cmdUpdate.Visible = False
            End If
            cmdPrint.Enabled = True
            
NEXT_EMPLOYEE:
            .MoveNEXT
        Next
        
'        LV.Visible = True
        
'        lblTotal = Format(GrandTotal, "#,##0.00")
        
    End With
  
    SQL = ""
    Cond = ""
    ICO = ""
    LV.Visible = True
    
End Sub
Private Function TotalPayable() As Double

    TotalPayable = Val(txtELAmt.Text) + Val(txtGAmount.Text) + Val(txtBonus10c.Text) + Val(txtNoticePay.Text) + Val(txtBLYear.Text) + Val(txtOvertimeAmt.Text)

End Function

Private Sub SaveIt()
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM EmpSettlements", con, adOpenStatic, adLockOptimistic
        .AddNew
        ![EmpId] = cmbEmp.ID
        DT = DateSerial(DTP.year, DTP.Month, 1)
        DT = DateSerial(DTP.year, DTP.Month + 1, 0)
        ![DT] = DT & ""
        
        For i = 1 To LV.ListItems.count
            !BSal = Val(LV.ListItems(i).SubItems(2) & "")
            '!HRent = Val(LV.ListItems(i).SubItems(3) & "")
            For j = 3 To 27
                If j = 8 And chkExternal Then
                    .Fields(j + 5) = LV.ListItems(i).ListSubItems(4).Tag
                Else
                    If .Fields(j + 5).Type = 5 Or .Fields(j + 5).Type = 4 Then
                        .Fields(j + 5) = Val(LV.ListItems(i).SubItems(j) & "")
                    Else
                        .Fields(j + 5) = LV.ListItems(i).SubItems(j) & ""
                    End If
                End If
            Next j
        
            ![PrevLTLoan] = Val(LV.ListItems(i).ListSubItems(1).Tag)
                
            Dim ArrLeaves() As String
        
            ArrLeaves = Split(LV.ListItems(i).ListSubItems(2).Tag, ",")
        
            ![CasualLeaves] = Val(ArrLeaves(0))
            ![SickLeaves] = Val(ArrLeaves(1))
            ![AnnualLeaves] = Val(ArrLeaves(2))
            ![CompensatoryLeaves] = Val(ArrLeaves(3))
            ![WPLeaves] = Val(ArrLeaves(4))
            ![MaternityLeaves] = Val(ArrLeaves(5))
            ![HrsPerDay] = Val(LV.ListItems(i).ListSubItems(3).Tag)
            ![IsForSA] = chkExternal
            ![TotalMonthHrs] = Val(txtTotalHrs.Tag)
            ![DTFinal] = Date
            
            '![FakeWorkingHrs] = "" 'GetFakeWorkingHrs(LV.ListItems(i))
'            ![FakeRate] = Val(LV.ListItems(i).ListSubItems(6).Tag)
'            ![FakeSalary] = Val(LV.ListItems(i).ListSubItems(5).Tag)
            
            
            !SundayOTHrs = Val(LV.ListItems(i).ListSubItems(12).Tag)
            !SundayOTRate = Val(LV.ListItems(i).ListSubItems(13).Tag)
            !FixAllowance = Val(LV.ListItems(i).ListSubItems(14).Tag)
            
            !LateComingHrs = Val(LV.ListItems(i).ListSubItems(15).Tag)
            !ShortHrs = Val(LV.ListItems(i).ListSubItems(16).Tag)
            
            !LeaveDays = Val(LV.ListItems(i).ListSubItems(25).Tag)
            !presentDays = Val(LV.ListItems(i).ListSubItems(26).Tag)
            
            !HoldSalaryAmt = Val(LV.ListItems(i).ListSubItems(27).Tag)
            
            !GraduityYears = Val(txtGPeriod)
            !GraduityAmount = Val(txtGAmount)
            !EarnedLeaves = Val(txtELeaves)
            !EarnedLeavesAmt = Val(txtELAmt)
            !NoticePayableSalary = Abs(chNoticePay1)
            !NoticeReceivableSalary = Abs(chkNoticePay2)
            !OtherDeductions = Val(txtOtherDed)
            !OtherReceivables = Val(txtOthers)
            !NetPayable = Val(txtNetPayable)
            .Update
        
            'Update Long Term ...
        
            Dim LongAmt As Double
            LongAmt = Val(LV.ListItems(i).SubItems(20))
        
            con.Execute "DELETE AmtCleared WHERE empID='" & rs![EmpId] & "' and DT='" & Format(DT, "dd-MMM-yyyy") & "'"
            con.Execute "INSERT INTO AmtCleared (EmpID,DT,AmtClrd) Values('" & rs![EmpId] & "','" & DT & "'," & Val(txtLBalance) & ")"
    Next i
        If i = 1 Then .Update
        .Close
    
    End With

    Set rs = Nothing
    
End Sub

Private Function TotalMonthDays() As Integer

    Dim TDate As Date
    Dim Tot As Integer
    Dim year As Integer
    
    year = DTPicker1.year
    
'    If Month(GetServerDate) = DTPicker1.Month Then
'        Tot = Day(GetServerDate)
'    Else
'        TDate = DateSerial(year, DTPicker1.Month, 1)
'        Tot = DateAdd("m", 1, TDate) - TDate
'    End If
    TDate = DateSerial(year, DTPicker1.Month, 1)
    Tot = DateAdd("m", 1, TDate) - TDate
    
    TotalMonthDays = Tot
    
End Function


Private Sub InsertRecords()

    Dim DT As Date
    DT = DateSerial(DTPicker1.year, DTPicker1.Month, 1)

    Dim i As Integer
    Dim j As Integer

    con.Execute "Delete from PrintSalary"

    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select * from PrintSalary", con, adOpenStatic, adLockOptimistic
        For i = 1 To LV.ListItems.count
            .AddNew
            ![EmpId] = LV.ListItems(i).Text & ""
        
            For j = 1 To 25
                If j = 25 Then
                    .Fields(j) = LV.ListItems(i).SubItems(j + 2) & ""
                ElseIf j = 8 And chkExternal Then                           '''' For Saving in Decimal Format instead of Time
                    .Fields(j) = LV.ListItems(i).ListSubItems(4).Tag
                Else
                    .Fields(j) = LV.ListItems(i).SubItems(j) & ""
                End If
            Next j
        
            ![DT] = DT & ""
            ![PrevLTLoan] = Val(LV.ListItems(i).ListSubItems(1).Tag)
                
            Dim ArrLeaves() As String
        
            ArrLeaves = Split(LV.ListItems(i).ListSubItems(2).Tag, ",")
        
            ![CasualLeaves] = Val(ArrLeaves(0))
            ![SickLeaves] = Val(ArrLeaves(1))
            ![AnnualLeaves] = Val(ArrLeaves(2))
            ![CompensatoryLeaves] = Val(ArrLeaves(3))
            ![WPLeaves] = Val(ArrLeaves(4))
            ![MaternityLeaves] = Val(ArrLeaves(5))
            ![HrsPerDay] = Val(LV.ListItems(i).ListSubItems(3).Tag)
        
            .Update
        Next i
    
        .Close
    
    End With

    Set rs = Nothing

    SQL = "Insert into EmpLeaveBalance SELECT  EmpID,Count(CL) as CasualLeaves,Count(SL) as SickLeaves," & _
     "count(AL) as AnnualLeaves" & _
     ",count(ML) as MaternityLeaves From VEmpLeaves WHERE DT <'" & DateSerial(DTPicker1.year, DTPicker1.Month, 1) & "' Group By empID"
    
    con.Execute "Delete From EmpLeaveBalance"
    con.Execute SQL, a

End Sub

Private Function AddSalaryVouchers()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim i As Long, INextSno As Double, Bal As Double, Amt As Double
    Dim VchrNo As String, theDesc As String, AccNo As String
    Dim Dated As Date
    Dim dDRTaxDiff As Double
    
    With rs
        .Open "Select Departments.Name,Departments.AccruedAccNo,Departments.SalaryAccNo,Departments.AccruedFoodAccNo,Departments.AccruedUnionAccNo,Departments.SchAccNo,Departments.AccruedEOBIAccNo,Departments.TaxAccNo, " & _
                "T1.DeptID,T1.TotalAmt,T1.SchAmt,T1.LunchAmt,T1.UnionAmt,T1.BalanceAmt,T1.ShortAmt,T1.LongAmt,T1.EOBIAmt,T1.TaxAmt " & _
                "From (Select EmpID,Left(EmpID,4) As DeptID,SAmt As TotalAmt,Bonus As SchAmt,Lunch As LunchAmt, UnionFund As UnionAmt, Balance As BalanceAmt, ShortTerm As ShortAmt, LongTerm As LongAmt,EOBI As EOBIAmt,Tax As TaxAmt From EmpSettlements " & _
                "Where Month(DT)=" & DTPicker1.Month & " and year(DT)=" & DTPicker1.year & ") " & _
                "T1 Inner JOIN Departments ON T1.DeptID=Departments.DeptID " & _
                "Where IsNull(Departments.AccruedAccNo,'')<>'' AND IsNull(Departments.SalaryAccNo,'')<>'' AND IsNull(Departments.AccruedFoodAccNo,'')<>'' AND IsNull(Departments.AccruedUnionAccNo,'')<>''  AND IsNull(Departments.SchAccNo,'')<>'' AND IsNull(Departments.AccruedEOBIAccNo,'')<>'' AND IsNull(Departments.TaxAccNo,'')<>'' AND T1.EmpID='" & cmbEmp.ID & "'", con, adOpenForwardOnly, adLockReadOnly
                
        Do Until .EOF
        
            Dated = DTPicker1.Value 'DateSerial(DTPicker1.year, DTPicker1.year, 0) 'Date
            
            If Format(Dated, "ddd") = "Sun" Then
                Dated = DateAdd("d", -1, Dated)
            End If
            
            INextSno = getNextSno(Dated)
            VchrNo = getNextVchrNo(Dated, "JV")
            
            '''''''''''''''''' T A X   E N T R I E S '''''''''''''
            If Val(!TaxAmt & "") > 0 Then
                theDesc = "Tax Of [" & cmbEmp.ID & "] " & cmbEmp.Text
                AccNo = !TaxAccNo & ""
                Bal = getBalance(AccNo, Dated)
                Amt = Val(!TaxAmt & "")
                If Amt <> 0 Then
                    Bal = Bal - Amt
                    
                    con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
                     Dated & "','" & VchrNo & "','" & AccNo & "','" & theDesc & _
                     "',0," & Amt & "," & Bal & ",'',0)"
                     
                    Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
                    INextSno = INextSno + 1
                End If
            End If
            
            'Charge All Amount to Cash
            Amt = Val(!TotalAmt & "")
    
            If Amt = 0 Then Exit Function
            'Give The Description
            theDesc = "Salary Of  [" & cmbEmp.ID & "] " & cmbEmp.Text
            
            '''''''''''''''''' D E B I T   E N T R I E S '''''''''''''
            '''''''''''''''''' Accrued Salary Account '''''''''''''
            AccNo = !SalaryAccNo
            Amt = Val(!TotalAmt & "")
            Bal = getBalance(AccNo, Dated)
      
            Amt = Amt
            Bal = Bal + Amt
            
            con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
                Dated & "','" & VchrNo & "','" & AccNo & _
                "','" & theDesc & "'," & Amt & ",0," & Bal & ",'',0)"
        
            Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
            INextSno = INextSno + 1
            '''''''''''''''''' Scholorship Account '''''''''''''
            
            AccNo = !SchAccNo
            Amt = Val(!SchAmt & "")
            If Amt <> 0 Then
              Bal = getBalance(AccNo, Dated)
        
              Amt = Amt
              Bal = Bal + Amt
              
              con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
                  Dated & "','" & VchrNo & "','" & AccNo & _
                  "','" & theDesc & "'," & Amt & ",0," & Bal & ",'',0)"
          
              Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
              INextSno = INextSno + 1
            End If
            '''''''''''''''''' C R E D I T   E N T R I E S '''''''''''''
            ''''''''''''''''Temp. Advance'''''''''''''
            
            AccNo = "15-015-20001"
            Amt = Val(!Shortamt & "")
            If Amt <> 0 Then
                Bal = getBalance(AccNo, Dated)
                
                Bal = Bal - Amt
            
                con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
                    Dated & "','" & VchrNo & "','" & AccNo & _
                    "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
        
                Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
         
                INextSno = INextSno + 1
            End If
            ''''''''''''''''Long Term Deduction'''''''''''''
            
            AccNo = "13-003-16001"
            Amt = Val(!LongAmt & "")
            If Amt <> 0 Then
                Bal = getBalance(AccNo, Dated)
    
                Bal = Bal - Amt
    
                con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
                    Dated & "','" & VchrNo & "','" & AccNo & _
                    "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
    
                Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
    
                INextSno = INextSno + 1
            End If
            ''''''''''''''''Food Amount Deduction'''''''''''''
            
            AccNo = !AccruedFoodAccNo & ""
            Amt = Val(!LunchAmt & "")
            If Amt <> 0 Then
                Bal = getBalance(AccNo, Dated)
                
                Bal = Bal - Amt
            
                con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
                    Dated & "','" & VchrNo & "','" & AccNo & _
                    "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
        
                Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
         
                INextSno = INextSno + 1
            End If
            ''''''''''''''''Food Amount Deduction'''''''''''''
            
            AccNo = !AccruedUnionAccNo & ""
            Amt = Val(!UnionAmt & "")
            If Amt <> 0 Then
                Bal = getBalance(AccNo, Dated)
                
                Bal = Bal - Amt
            
                con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
                    Dated & "','" & VchrNo & "','" & AccNo & _
                    "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
        
                Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
         
                INextSno = INextSno + 1
            End If
            ''''''''''''''''Balance Amount Deduction'''''''''''''
            'If There's Ded. Of Tax. Then Adjust The Difference Of Rounding Of Tax Amount.
            
            AccNo = !AccruedAccNo & ""
            Amt = Val(!BalanceAmt & "")
            

            dDRTaxDiff = Val(!TotalAmt & "") + Val(!SchAmt & "") - Amt - Val(!LunchAmt & "") - Val(!UnionAmt & "") - Val(!Shortamt & "") - Val(!LongAmt & "") - Val(!EOBIAmt & "") - Val(!TaxAmt & "")
            Amt = Amt + dDRTaxDiff

            
'            If !DeptID = "VDRS" Then
'                dDRTaxDiff = Val(!TotalAmt & "") - Amt - Val(!TaxAmt & "")
'                Amt = Amt + dDRTaxDiff
'            Else
'                dDRTaxDiff = Val(!TotalAmt & "") + Val(!SchAmt & "") - Amt - Val(!LunchAmt & "") - Val(!UnionAmt & "") - Val(!ShortAmt & "") - Val(!LongAmt & "") - Val(!EOBIAmt & "")
'                Amt = Amt + dDRTaxDiff
'            End If
            
            Bal = getBalance(AccNo, Dated)
            
            Bal = Bal - Amt
        
            con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
                Dated & "','" & VchrNo & "','" & AccNo & _
                "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
    
            Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
     
            INextSno = INextSno + 1
            
            ''''''''''''''''EOBI Amount Deduction'''''''''''''
            
            AccNo = !AccruedEOBIAccNo & ""
            Amt = Val(!EOBIAmt & "")
            If Amt <> 0 Then
                Bal = getBalance(AccNo, Dated)
                
                Bal = Bal - Amt
            
                con.Execute "Insert Into Vouchers Values(" & INextSno & ",'" & _
                    Dated & "','" & VchrNo & "','" & AccNo & _
                    "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
        
                Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
         
                INextSno = INextSno + 1
            End If
            
            con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & VchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate & "')"
                
            .MoveNEXT
        Loop
        .Close
        
    End With
    
    Set rs = Nothing
    AddSalaryVouchers = True
    
    Exit Function
err:
    MsgBox err.Description
    
End Function


Private Sub DTP_Change()

   ' txtDuration.Text = DateDiffInText(DTPicker1)

'
'    Dim iYears As Integer, iMonths As Integer, iDays As Integer
'    iYears = DateDiff("m", DTPicker1.value, DTP.value)
'
'    If Day(myDT) > Day(Date) Then
'        iYears = iYears - 1
'    End If
'
'    myDT = DateAdd("M", iYears, myDT)
'    iMonths = iYears Mod 12
'
'    iDays = DateDiff("d", myDT, DTP)
'    If iDays < 0 Then
'        myDT = DateAdd("M", -1, myDT)
'        iDays = DateDiff("d", myDT, Date)
'
'        iMonths = iMonths - 1
'    End If
'
'    iYears = Int(iYears / 12)
'    txtDuration.Text = "Years " & iYears & ", Months " & iMonth & " ,Days " & iDays
End Sub

Private Sub Form_Load()
    
    Me.Left = 1500
    Me.Top = 200
    cmbEmp.ListHeight = 2200

    'DTLeave.value = Date
    With cmbDept
        .ListHeight = 2200
        .ClearVals
        .AddItem "< All Departments >", "0"
        .AddVals con, "Name", "Departments", "DeptID"
        .ID = "0"
    End With
    
    DTP = Date
    Me.Width = 10695
    Me.Height = 8430
    DTPicker2.Value = GetServerDate
        
    txtTotalHrs = 240
    txtTotalHrs.Tag = 240
End Sub

Private Sub txtBLYear_Change()
    txtTotalPayable.Text = TotalPayable
End Sub

Private Sub txtELeaves_Change()
    
    If Val(txtELeaves.Text) > 14 Then
        MsgBox "You can't enter more than 14 days !", vbInformation
        txtELeaves.SetFocus
        Exit Sub
    End If
    txtELAmt.Text = Round((Val(txtSalary.Text) / 30) * Val(txtELeaves.Text), 2)
    UpdateFields
End Sub

Private Sub txtEOBI_Change()
    UpdateFields
End Sub

Private Sub txtGAmount_Change()
    UpdateFields
End Sub

Private Sub txtGPeriod_Change()
    txtGAmount = Val(txtGPeriod) * Val(txtSalary)
End Sub

Private Sub txtNoticePay_Change()
    txtTotalPayable.Text = TotalPayable
    txtNetPayable.Text = Val(txtTotalPayable.Text) - Val(txtTotalReceivable.Text)
End Sub

Private Sub txtNoticePay2_Change()

    txtTotalReceivable.Text = Val(txtNoticePay2.Text) + Val(txtLBalance)
    txtNetPayable.Text = Val(txtTotalPayable.Text) - Val(txtTotalReceivable.Text)
    
End Sub

Private Sub txtOtherDed_Change()
    'txtDeducation = Val(txtTax) + Val(txtEOBI) + Val(txtRent) + Val(txtPF) + Val(txtOtherDed)
    UpdateFields
End Sub

Private Sub txtOtherPayable_Change()
    UpdateFields
End Sub

Private Sub txtothers_Change()
        
'    txtTotalReceivable.Text = Val(txtNoticePay2.Text) + Val(txtLBalance) + Val(txtOthers.Text)
'    txtNetPayable.Text = Val(txtTotalPayable.Text) - Val(txtTotalReceivable.Text)
    UpdateFields
        
End Sub


Private Sub RefreshSalary()
    
    Dim rs As New ADODB.Recordset
    Dim rsInternal As New ADODB.Recordset
    With rs
        .Open "Select Count(EmpID) From MonthlySalaries Where Month(DT)=" & DTP.Month & " AND Year(DT)=" & DTP.year & " AND IsForSA=0", con, adOpenForwardOnly, adLockReadOnly
        If Val(.Fields(0) & "") = 0 Then
            IsInternalFinal = False
        Else
            IsInternalFinal = True
            
            'To Save Long term Loan Information....
            rsInternal.Open "Select EmpID,LongTerm,PrevLTLoan From MonthlySalaries Where Month(DT)=" & DTP.Month & " AND Year(DT)=" & DTP.year & " AND IsForSA=0", con, adOpenStatic, adLockReadOnly
            
        End If
        .Close
    End With
    
    Dim TotalSalDays As Double
    Dim TDate As Date
    Dim SDate As Date 'For Saving
    Dim Hdays As Long
    
    SDate = DateSerial(DTP.year, DTP.Month, 1)
    TDate = DateSerial(DTP.year, DTP.Month, DTP.Day)
    
    TotalSalDays = DateAdd("m", 1, SDate) - SDate

    MonthDays = TotalSalDays - TotalSundays(DTP.Month, DTP.year, TotalSalDays)
    
    If chkExternal = vbUnchecked Then
        Hdays = Val(con.Execute("Select count(DT) from Holidays Where month(DT)=" & _
        DTP.Month & " and Year(DT)=" & Val(DTP.year)).Fields(0))
        MonthDays = MonthDays - Hdays
    End If
   
    If FinalizedSalarySheetLoaded Then Exit Sub
    
    'txtTotalHrs.Tag = txtTotalHrs

    Dim UnionFund As Integer
    Dim GrandTotal As Double
    Dim EOBIAmt As Byte
    If cmbDept.MatchFound = False Or cmbEmp.MatchFound = False Then Exit Sub
    'If DTP.year.MatchFound = False Or dtp.month.MatchFound = False Then Exit Sub
    
    Dim SunDates() As Date
        
    Dim EmpId As String
    Dim EmpAge As Double
    
    Dim TempTotalSaldays As Double
    Dim LastDay As Integer
    
    Dim ITM As ListItem
    Dim Cond As String
    Dim SQL As String
    Dim cmd As New ADODB.Command
    
    
    Dim SalHrs As Double
    Dim HrsMultiple As Double
    
    
    Dim Leaves As Double
    Dim ApprovedLeaves As Double
    Dim UnApprovedLeaves As Double
    
    If cmbEmp.ID = "0" Then
       EmpId = IIf(cmbDept.ID = "0", "", cmbDept.ID)
    Else
       EmpId = IIf(cmbEmp.ID = "0", "", cmbEmp.ID)
    End If
    
    Dim rsCheck As New ADODB.Recordset
        
    Dim LunchDays As Long
    
    
    Dim OTByQty As Boolean
    
    
    'TDate = DateSerial(DTP.year, dtp.month, Day(GetServerDate))
        
   
    ''''''''''''''''''''' This Will Refresh The LateHours '''''''
    '            con.Execute "EXEC SP_CalcOverTime '" & SDate & "'", "'" & cmbEmp.ID & "'"
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        
        
    LastDay = TotalSalDays
    
    Dim JDate As Date
    
    If TDate = GetServerDate Then
        TotalSalDays = Day(GetServerDate)
    ElseIf GetServerDate < TDate Then
        MsgBox "Can't Show Entries Of Future Months.", vbInformation
        Exit Sub
    End If
    
    TempTotalSaldays = TotalSalDays     'Store For Later Use
    
    JDate = DateSerial(DTP.year, DTP.Month, DTP.Day)
        
    'Sql = "Select * from VAdvanceSalary " & Cond & " JoinDate<='" & Format(JDate, "dd-MMM-yyyy") & "' Order By DeptID,EmpID"
    
    cmd.CommandType = adCmdStoredProc
    
    cmd.CommandText = "VAdvanceSalary"
    
    Set cmd.ActiveConnection = con
    cmd.Parameters("@ParDate").Value = JDate
    'cmd.Parameters(1).value = JDate
    If cmbEmp.ID <> "0" Then
        cmd.Parameters("@EmpID").Value = cmbEmp.ID
    End If
    cmd.Parameters("@ForDC").Value = True
'    cmd.Parameters(2).value = EmpId
        
    If rsLedger.State = 1 Then rsLedger.Close
    rsLedger.CursorType = adOpenForwardOnly
    rsLedger.LockType = adLockReadOnly
    
    Set rsLedger = cmd.Execute
    
    With rsLedger
        
        '.Open Sql, con, adOpenForwardOnly, adLockReadOnly
        Dim TotalSuns As Integer
        Dim SalDays As Double
        Dim Rate As Double, FakeRate As Double
        Dim AbsAllowed As Integer
        Dim AbsentDays As Double
        
        Dim total As Double
        Dim OverTimeHours As Double
        Dim OverTimeHoursSA As Double
        Dim LateHours As Double
        Dim TaxPer As Double
        Dim DedAmt As Double
        Dim Payed As Double
        Dim LongLoan As Double
        Dim TaxSal As Double
        Dim IsComputed As Boolean
        
        Dim NTime As Integer
        Dim LTime As Double
        Dim DT As String
        Dim IsPayed As Boolean
        Dim IsTransfered As Boolean
        Dim Holidays As Integer
        
        Dim EmpSalary As Double
        Dim EmpExtSal As Double

        LV.ListItems.Clear
        LV.Visible = False
        
        Dim dLateComingHrs As Double, dOTHrs As Double, dShortHrs As Double
        
        For i = 0 To .RecordCount - 1
                    
            'Reset Total Salary days For Each Employee
            
            TotalSalDays = DateAdd("m", 1, SDate) - SDate
                
            If TDate = GetServerDate Then
                TotalSalDays = Day(GetServerDate)
            End If
            
            'Check that Employee Discontinued In This Month
                        
            If Not IsNull(![DiscontinuedOn]) Then
                If Month(![DiscontinuedOn]) = Month(JDate) And ![DiscontinuedOn] <= JDate Then
                  TotalSalDays = Day(![DiscontinuedOn])
                End If
            Else
               TotalSalDays = TempTotalSaldays
            End If
            
            Leaves = Val(!Leaves & "")
                                   
           'if Joined This Month Calculate From Join Date
           If Format(![JoinDate], "MMM-yyyy") = Format(JDate, "MMM-yyyy") Then
               
               rs.Open "SELECT count(DT) FROM Holidays WHERE DT Between '" & ![JoinDate] & "' AND '" & JDate & "'", con, adOpenForwardOnly, adLockReadOnly
               
                  If rs.RecordCount > 0 Then
                     Holidays = rs.Fields(0) & ""
                  End If
               rs.Close
               Set rs = Nothing
               
               SundayAbsents = getSundayAbsents(![EmpId] & "", Val(DTP.Month), DTP.year, Val(![WeeklyAbs] & ""), Int(TotalSalDays))
               TotalSuns = TotalSundays(DTP.Month, DTP.year, DTP.Day, Day(![JoinDate]))
               Holidays = TotalSuns + Holidays - SundayAbsents
               
            Else
            
                 rs.Open "SELECT count(DT) FROM Holidays WHERE MONTH(DT)=" & DTP.Month & " AND Year(DT)=" & DTP.year & " AND day(dt)<=" & TotalSalDays, con, adOpenForwardOnly, adLockReadOnly
                  
                  If rs.RecordCount > 0 Then
                     Holidays = rs.Fields(0) & ""
                  End If
                  
                 rs.Close
               Set rs = Nothing

               'SundayAbsents = getSundayAbsents(![EmpId], dtp.month, dtp.year, Int(TotalSalDays), ![WeeklyAbs])
               
               TotalSuns = TotalSundays(DTP.Month, DTP.year, TotalSalDays)
                                              
               Holidays = TotalSuns + Holidays - SundayAbsents
               

            End If
                    
            If chkExternal = vbChecked Then
                'Exclude The Holidays From Calculation
                TotalSalDays = TotalSalDays - Holidays
            End If
            
            'Get Sundays Not To Be Paid

            'Call Computed(![EmpId], IsComputed, TaxPer, AbsAllowed, OverTimeHours, NTime, LTime, LongLoan, DT, IsPayed, IsTransfered, JDate, EmpSalary, EmpExtSal, ![Active], HrsMultiple)
            
            If Not IsComputed Then
                'Tax %
                'If IsNull(![PerDeduct]) Then TaxPer = 0 Else TaxPer = Val(![PerDeduct]) & ""
                
                'Absent Days....
                AbsAllowed = Val(![AbsAllowed] & "")
    
                ' Reset Orver Time Hours
                OverTimeHours = 0
                OverTimeHoursSA = 0
                
                'Late Hours
                NTime = Val(![NTimes] & "")
                LTime = Val(![TimeAllowed] & "")
                
                'Long Term Loan Deduction
                Dim lDedAmt As Long, lClsAmt As Long, lTotAmt As Long
                If IsNull(![DedAmt]) Then lDedAmt = 0 Else lDedAmt = ![DedAmt]
                If IsNull(![totamt]) Then lTotAmt = 0 Else lTotAmt = ![totamt]
                If IsNull(![ClsAmt]) Then lClsAmt = 0 Else lClsAmt = ![ClsAmt]
                
                
                If IsNull(!AmtToDeduct) Then
                    If lDedAmt > (lTotAmt - lClsAmt) Then
                        LongLoan = (lTotAmt - lClsAmt)
                    Else
                        LongLoan = lDedAmt
                    End If
                Else
                    LongLoan = Val(!AmtToDeduct & "")
                End If
                                      
                'Get The Salary Depending On Salary Sheet Type
                'EmpSalary = IIf(chkExternal = vbUnchecked, Val(![StartingSalary] & ""), Val(![Salary] & ""))
                
                 EmpSalary = IIf(m_bShowFake, Val(![EmpMinSalary] & ""), Val(![StartingSalary] & ""))
                 'EmpExtSal = Val(![Salary] & "")
                 EmpExtSal = EmpSalary
            End If
                                    
            'Check If Employee is paid Over time per piece
'            If Not IsNull(![OTByQty]) Then If ![OTByQty] Then OTByQty = True
                            
                If Not IsNull(![OTByQty]) Then
                    If ![OTByQty] Then               ' Calculate Over Time On the Basis of Qty Worked If Applicable
                        OTByQty = True
                        If Val(![QtyPerHour] & "") > 0 Then
                            'OTQty = Val(![TotalQtyWorked] & "") - (Val(![MinQty] & "") * SalDays)
                            OverTimeHours = Val(![OTQtyWorked] & "")
                            'OverTimeHours = OTQty / Val(![QtyPerHour] & "")
                        Else
                            OverTimeHours = 0
                        End If
                    Else
                        If IsNull(![OverTimeHours]) Then
                            OverTimeHours = 0
                            OverTimeHoursSA = 0
                        Else
                            OverTimeHours = ![OverTimeHours]
                            OverTimeHoursSA = ![OverTimeHours]
                        End If
                    End If
                    OTByQty = ![OTByQty]
                Else
                    
                    OTByQty = False
                    'OverTime Hours
                    
                    If IsNull(![OverTimeHours]) Then
                        OverTimeHours = 0
                        OverTimeHoursSA = 0
                    Else
                        OverTimeHours = ![OverTimeHours]
                        OverTimeHoursSA = ![OverTimeHours]
                    End If
                    
                End If
            '*********************
'            OverTimeHours = 0
'            OverTimeHoursSA = 0
            '*********************
            Dim theKey As String
            
            If IsComputed Then
                theKey = ![EmpId] & Format(SDate, "dd-MMM-yyyy")
            Else
                theKey = ![EmpId] & ""
            End If
            
            Set ITM = LV.ListItems.add(, theKey, ![EmpId] & "")
            
            If IsComputed Then
                ITM.Checked = IsPayed
                ITM.Bold = IsTransfered
            End If
            
            
            ITM.Tag = Val(![Emptype] & "")
            
'            If IsInternalFinal And chkExternal.value = vbChecked Then
'                rsInternal.MoveFirst
'                rsInternal.Find "EmpID='" & ![EmpId] & "'", , adSearchForward, 0
'                If rsInternal.EOF Then
'                    ITM.ListSubItems.add(, , ![Name] & "").Tag = 0
'                Else
'                    ITM.ListSubItems.add(, , ![Name] & "").Tag = Val(rsInternal![PrevLTLoan] & "")
'                End If
'            Else
'                ITM.ListSubItems.add(, , ![Name] & "").Tag = Val(![TotAmt] & "") - Val(![clsamt] & "")
'            End If
            ITM.ListSubItems.add(, , ![Name] & "").Tag = Val(![totamt] & "") - Val(![ClsAmt] & "")
            
            ITM.ListSubItems.add , , IIf(chkExternal = vbChecked, EmpExtSal, EmpSalary & "")
            
            ITM.ListSubItems(2).Tag = ![CasualLeaves] & "," & ![SickLeaves] & "," & ![AnnualLeaves] & _
            "," & ![CompensatoryLeaves] & "," & ![WPLeaves] & "," & ![MaternityLeaves]
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '   Rate = Format(Val(EmpSalary) / LastDay, "0.00")
            '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '   Calculate the Rate According To Total No ........
            '   Of Working Days in This Month for External Salary
            '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            If chkExternal = vbChecked Then
                                
                'Rate = Format(Val(EmpSalary) / (LastDay - TotalSuns), "0.00")
                
                Rate = Format(Val(EmpSalary) / (LastDay), "0.00")
                If HrsMultiple = 0 Then HrsMultiple = 8
            Else
                
                Rate = Format(Val(EmpSalary) / LastDay, "0.00")
                
                HrsMultiple = 1
                
            End If

            Rate = Rate / HrsMultiple
                       
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '   Rate Formula For October Salary Sheet ...
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            If chkExternal Then
                If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                    Rate = EmpSalary / 240
                Else
                    Rate = EmpSalary / Val(txtTotalHrs)
                End If
                FakeRate = Val(!EmpMinSalary & "") / Val(txtTotalHrs)
            End If
                       
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
            'Store The Total Working hours for This Employee
            ITM.ListSubItems.add(, , Format(Rate, "0.00")).Tag = HrsMultiple
            
            If IsNull(![presentDays]) And !DeptID <> "VDRS" Then
                SalDays = 0
                'IF No Presents Found Exclude This Employee
                LV.ListItems.Remove ITM.key
                GoTo NEXT_EMPLOYEE
            Else
                
                
                SalDays = ((Val(![presentDays] & "")) + Holidays) + AbsAllowed
                
                'Store for Later Calculation
                
                'LunchDays = Val(![presentDays])
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'Exclude the Holidays now (for Vision External Salary Sheet Only)
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                If chkExternal = vbChecked Then SalDays = ((Val(![presentDays] & ""))) + AbsAllowed
                
                'If Salary Days Exceed Total Salary Days Set Them To Salary Days
                If SalDays > TotalSalDays Then SalDays = TotalSalDays
                    
            End If
            
            Dim DaysToExclude As Integer
            
            If Format(![JoinDate], "MMM-yyyy") = Format(JDate, "MMM-yyyy") Then
               If Day(!JoinDate) = 5 Then
                    DaysToExclude = Day(![JoinDate]) - 1
                Else
                    DaysToExclude = Day(![JoinDate]) - 1
                End If
            Else
               DaysToExclude = 0
            End If
            
            
            'Deduct Penalty Days
            
            'If chkExternal = vbUnchecked Then
            '    SalDays = SalDays - Val(![AbsentDays] & "")
            'End If
            
            If SalDays < 0 Then SalDays = 0
                        
            If chkExternal Then
                Leaves = Val(!Leaves & "")
                UnApprovedLeaves = 0
            Else
                Leaves = Val(!ApprovedLeaves & "")
                UnApprovedLeaves = Val(!Leaves & "") - Val(!ApprovedLeaves & "")
            End If
            

            
            '''''''''''''''''''''''''''''''
            ' dont do anything with sunday/holiday related absents at the moment
            ''''''''''''''''''''''''''''''''
            ' SalDays = SalDays - CalcSpecialAbs(!EmpId, JDate, !BGHoliday, !AGHoliday, !BPHoliday, !APHoliday)
            '''''''''''''''''''''''''''''''
            
            SalDays = SalDays - UnApprovedLeaves
            
            If Abs(![DeductAbsents]) = 0 Then SalDays = TotalSalDays - DaysToExclude
            
            AbsentDays = TotalSalDays - SalDays - DaysToExclude
            
            
            ' Make Absents 0 if in Negative
            If AbsentDays < 0 Then AbsentDays = 0
            
            
            'If IsNull(![AbsentDays]) Then AbsentDays = 0 Else AbsentDays = Val(![AbsentDays]) / 2
            'AbsentDays = IIf(AbsentDays > 0, AbsentDays - AbsAllowed, AbsentDays)
            

            ' Convert To Hours if External Salary
            
            ITM.ListSubItems.add , , AbsentDays * HrsMultiple
            
            '''''''''''''''''' Save the Salary Hours in it tag Later '''''''''''''''''''''''
            
            
            ITM.SubItems(5) = Round(AbsentDays * HrsMultiple * Rate)
            
            'If Val(!Leaves & "") > 0 Then
            '    If AbsentDays - AbsAllowed >= 0 Then
                
            '    itm.SubItems(11) = IIf((Exempt), "0", (Leaves * Rate))
                
            '    End If
            'Else
            '    itm.SubItems(11) = "0"
            'End If
            
            ITM.SubItems(11) = "0"
            
            ' Calculat Net Absents Allowed
            ITM.SubItems(6) = AbsAllowed
            Dim NetAbsAllowed As Integer
            
            ' Calculate The Allowed Amount That is Subtracted From The Absent Amount
            
            If AbsentDays = 0 Then
                NetAbsAllowed = 0
            Else
                NetAbsAllowed = IIf(AbsAllowed - AbsentDays <= 0, AbsAllowed, AbsAllowed - AbsentDays)
            End If
            
            'Absent Allowed Amount
            ITM.SubItems(7) = IIf((Exempt), "0", Round(Rate * NetAbsAllowed * HrsMultiple, 2))
            
            
            SalHrs = SalDays * HrsMultiple
                        
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '                      Fix the Salary hours of Directors
            '               and Give Relaxation of 3 hours for october salary
            '                       Add the Leave hours as defined
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
            
            If chkExternal Then
'                If ![DeptID] = "VDRS" Then
                If Val(!EmpAttCat_RefID & "") = 6 Or Val(!EmpAttCat_RefID & "") = 5 Then
                    SalHrs = Val(txtTotalHrs)
                Else
             
                    'SalHrs = Round(Val(![HrsWorked] & ""), 1) + ((Holidays - TotalSuns) * 8)
                    SalHrs = Round(Val(![HrsWorked] & ""), 2) + ((Holidays) * 8)
                    ' Add All the Leaves to Worked hours
                    SalHrs = SalHrs + (Val(![Leaves] & "") * 8)
                    
                    'Remove Full Hours of Without Pay leaves
                    SalHrs = SalHrs - (Val(![WPLeaves] & "") * 8)
                    
                    ' Remove half Hours of Sick Leaves
                    'SalHrs = SalHrs - (Val(![SickLeaves] & "") * 4)
                    
                    ' Check if Salary Hours are greater than Total Hours Due to relaxation
                    ' Fix them to maximum hours
                    If strCompany <> "Dr-Frgz" Then
                        If SalHrs > Val(txtTotalHrs) Then SalHrs = Val(txtTotalHrs)
                    End If
                End If
            Else
                If ![DeptID] = "VDRS" Then
                    SalHrs = Val(txtTotal)
                Else
                    'Remove Full Hours of Without Pay leaves
                    SalHrs = SalHrs - Val(![WPLeaves] & "")
                    
                    'Remove half Hours of Sick Leaves
                    SalHrs = SalHrs - Val(![SickLeaves] & "") / 2
                    
                    ' Check if Salary Hours are greater than Total Hours Due to relaxation
                    ' Fix them to maximum hours
                    
                        If SalHrs > Val(txtTotal) Then SalHrs = Val(txtTotal)
                    
                    
                End If
                
            End If
            
            If !DeductionSpecialAbsents Then
                SalHrs = SalHrs - (Val(!SPAbsentDays & "") * 8)
            End If
            If SalHrs < 0 Then SalHrs = 0
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '       Show in Time Format if external salary sheet
            '       and Show days in internal salarysheet instead
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            ITM.SubItems(8) = IIf(chkExternal, getHrsMin(SalHrs), SalHrs)
            
            dLateComingHrs = Val(!LateComingHrs & "")
            If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                'If !EmpId = "DFG-000014" Then Stop
                Dim dAbsentHrs As Double, dDailyRate As Double, dAbsentDays As Double, dPresentDays As Double, dLateHrs As Double, dLateAmt As Double
                dAbsentHrs = Val(txtTotalHrs.Tag) - SalHrs
                dDailyRate = EmpSalary / 30
                dPresentDays = Val(!presentDays & "") + Holidays + Leaves - Val(![WPLeaves] & "")
                Dim iMonthDays As Integer
                iMonthDays = Day(DateSerial(DTP.year, DTP.Month + 1, 0))
                dAbsentDays = iMonthDays - dPresentDays
                
                Dim dAbsDaysForSalCalc As Double
                dAbsDaysForSalCalc = dAbsentDays
                If !DeductionSpecialAbsents Then
                    dPresentDays = dPresentDays - Val(!SPAbsentDays & "")
                    dAbsDaysForSalCalc = dAbsDaysForSalCalc + Val(!SPAbsentDays & "")
                    'SalHrs = SalHrs - (Val(!SPAbsentDays & "") * 8)
                End If
                
                If dPresentDays > dAbsDaysForSalCalc Then
                    ITM.SubItems(9) = Round((EmpSalary - (dDailyRate * dAbsDaysForSalCalc)) / 10, 0) * 10 'Round((SalHrs * Rate) / 10, 0) * 10
                    dPresentDays = 30 - dAbsDaysForSalCalc
                Else
                    ITM.SubItems(9) = Round((dPresentDays * dDailyRate) / 10, 0) * 10
                End If
                'ITM.SubItems(9) = Round((dPresentDays * dDailyRate) / 10, 0) * 10
                dLateHrs = (Val(iMonthDays * 8)) - SalHrs - (Val(![WPLeaves] & "") * 8)
                dLateHrs = dLateHrs - (dAbsentDays * 8)
                If !DeductionSpecialAbsents Then
                    dLateHrs = dLateHrs - (Val(!SPAbsentDays & "") * 8)
                End If
                dShortHrs = (dLateHrs - dLateComingHrs)
                If OverTimeHoursSA > 0 Then
                    If OverTimeHoursSA > dLateComingHrs Then
                        OverTimeHoursSA = OverTimeHoursSA - dLateComingHrs
                        dLateComingHrs = 0
                    Else
                        dLateComingHrs = dLateComingHrs - OverTimeHoursSA
                        OverTimeHoursSA = 0
                    End If
                End If
                
                dLateAmt = (dShortHrs + dLateComingHrs) * (Rate * (Val(!DeductionPercentage & "") / 100))
                dLateAmt = Round(dLateAmt / 10, 0) * 10
                
            Else
                ITM.SubItems(9) = Round((SalHrs * Rate) / 10, 0) * 10
            End If
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '       Save the Salary Hours in Listsubitmes(4).tag To Save
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ITM.ListSubItems(4).Tag = SalHrs
            ITM.ListSubItems(5).Tag = Val(![EmpMinSalary] & "")
            ITM.ListSubItems(6).Tag = FakeRate
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            ITM.SubItems(10) = Leaves * HrsMultiple
            
            If cmbMonth = "Sep" And DTP.year = "2004" Then
                ' Salary Sheet For September 2004 Must Be shown with the leave deduction
                ' BCoz Its on SA-8000 Record Now ...
                ITM.SubItems(11) = Round(Val(!Leaves & "") * HrsMultiple * Rate, 2)
            Else        'This Bug is Corrected Afterwards as follows
                ' Do it only For Internal Salary Sheet
                If chkExternal = vbUnchecked Then
                    LeaveDedAmt = Val(![SickLeaves] & "") * HrsMultiple * Rate / 2
                    LeaveDedAmt = LeaveDedAmt + Val(![WPLeaves] & "") * HrsMultiple * Rate
                    ITM.SubItems(11) = Round(LeaveDedAmt, 2)
                End If
                ITM.SubItems(11) = 0
            End If
            
            'If SalDays = 0 Then SalDays = 0 Else SalDays = (SalDays - AbsentDays)
            
            If OTByQty Then
                ITM.SubItems(12) = OverTimeHours
                'Over time By Quantity : Per Day Rate / Sheets Per Day * 1.4
                ITM.SubItems(13) = Round(OverTimeHours * (Rate * HrsMultiple / 56 * 1.4))
            Else
                '
                ITM.SubItems(12) = IIf(chkExternal, MinsToTime(OverTimeHoursSA * 60), MinsToTime(OverTimeHours * 60))
                ITM.ListSubItems(12).Tag = Val(!SundayOTHrs & "")
                If chkExternal Then
                    If strCompany = "QEL" Then
                        'ITM.SubItems(13) = Round(OverTimeHoursSA * (Rate + (Rate * 62.5 / 100)))
                        ITM.SubItems(13) = Round(OverTimeHoursSA * (Rate * 1.333333))
                    ElseIf strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                        ITM.SubItems(13) = Round(OverTimeHoursSA * (Rate * Val(!OTRate & "")))
                    Else
                        ITM.SubItems(13) = Round(OverTimeHoursSA * (Rate * 1.333333))
                    End If
                Else
                    ITM.SubItems(13) = Round(OverTimeHours * (Rate / 8 * 1.333333))
                End If
                If strCompany = "PAKSMITH GROUP" Then
                    ITM.ListSubItems(13).Tag = Val(!SundayOTHrs & "") * (Rate * 1.333333)
                ElseIf strCompany = "QEL" Then
                    '2.5 Hours=4;5=8; I think it's 62.5 % Increase over Standard Rate
                    'ITM.ListSubItems(13).Tag = Val(!SundayOTHrs & "") * (Rate + (Rate * 62.5 / 100))
                    ITM.ListSubItems(13).Tag = Val(!SundayOTHrs & "") * (Rate * 1.333333)
                Else
                    ITM.ListSubItems(13).Tag = Val(!SundayOTHrs & "") * (Rate * 2)
                End If
            End If
            
            LateHours = Val(![HOURS] & "") - LTAllowed
            
            If LateHours < 0 Then LateHours = 0
            
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''
            '           Exempt For October 2004
            ''''''''''''''''''''''''''''''''''''''''''''''''''
            'Well It was just written as LateHours=0 And That was why it was'nt deducting LT.
            'So I Just Commented it Out.
            'LateHours = 0
            ''''''''''''''''''''''''''''''''''''''''''''''''''
                        
            If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                ITM.SubItems(14) = MinsToTime(dLateHrs * 60)
            Else
                ITM.SubItems(14) = MinsToTime(LateHours * 60)
            End If
            
            ITM.ListSubItems(14).Tag = Val(!FixAllowance & "")
           ' itm.SubItems(15) = IIf((Exempt), "0", Format((LateHours * (Rate / Val(!WorkingHrs & ""))), "0.00"))
            
            If chkExternal Then
                If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                    ITM.SubItems(15) = dLateAmt
                Else
                    ITM.SubItems(15) = IIf((Exempt), "0", Format((LateHours * Rate), "0.00"))
                End If
            Else
                ITM.SubItems(15) = IIf((Exempt), "0", Format((LateHours * Rate / 8), "0.00"))
            End If
            
            'Total = Format(Val((ITM.SubItems(9)) + Val(ITM.ListSubItems(7))) - (Val(ITM.SubItems(15)) + Val(ITM.SubItems(11))), "0.00")
            If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                total = Format(Val((ITM.SubItems(9)) + Val(ITM.ListSubItems(7)) + Val(ITM.ListSubItems(13)) + Val(ITM.ListSubItems(13).Tag) + Val(ITM.ListSubItems(14).Tag)) - Val(ITM.SubItems(15)) - (Val(ITM.SubItems(11))), "0.00")
            Else
                total = Format(Val((ITM.SubItems(9)) + Val(ITM.ListSubItems(7)) + Val(ITM.ListSubItems(13)) + Val(ITM.ListSubItems(13).Tag) + Val(ITM.ListSubItems(14).Tag)) - (Val(ITM.SubItems(11))), "0.00")
            End If
            
            ITM.SubItems(16) = total
            
            'TaxPer = Val(![PerDeduct])
            ITM.SubItems(17) = Val(!TaxAmt & "") 'Round(getTaxAmt((EmpSalary * 12) / 1.5) / 12, 0)
            'DedAmt = Round((Val(![Salary]) / TotalSalDays) * (TaxPer / 100))
            'TaxSal = Round((Val(EmpSalary) / LastDay) * TotalSalDays)
            DedAmt = 0 'Round(TaxSal * (TaxPer / 100))
            
            'Tax Percentage That Is Zero Length
            'itm.SubItems(18) = "0"
            
            'itm.SubItems(14) = DedAmt
            ITM.SubItems(18) = total 'Total - Val(ITM.SubItems(17))
            If IsNull(![Payed]) Then Payed = 0 Else Payed = ![Payed]
            ITM.SubItems(19) = Round(Payed / 10) * 10
            
            
            ITM.SubItems(20) = LongLoan
            ITM.ListSubItems(20).Tag = (lTotAmt - lClsAmt)
            
            
            ITM.SubItems(21) = Val(![AdvSalpaid] & "")
            
            
            'Dim SocialAmt As Double
            'SocialAmt = IIf((!Social = True), !SocialAmt, 0)
            
            
            ''''''''''''''''''''''''''''''''''''''''''
            '       Union Fund deduction
            ''''''''''''''''''''''''''''''''''''''''''
            
                If Not IsNull(![UnionFund]) Then
                    If ![UnionFund] Then UnionFund = 20 Else UnionFund = 0
                Else
                    UnionFund = 0
                End If
            
            ''''''''''''''''''''''''''''''''''''''''''
                Dim LunchDed As Double
                
                If Not IsNull(![Lunch]) Then
                    LunchDed = Val(![LunchAmt] & "")
                Else
                    LunchDed = 0
                End If
            ''''''''''''''''''''''''''''''''''''''''''
            
            ITM.SubItems(22) = UnionFund
            ITM.SubItems(23) = Val(!TotalFine & "")
            
            ' Use This Column For Scholorships Instead Of Bonus ...
            'ITM.SubItems(24) = Val(!TotalBonus & "")
            
            ITM.SubItems(24) = Val(![ScholorshipAmt] & "")
            ITM.SubItems(25) = LunchDed
            
            If Not IsNull(![OldAge]) Then If ![OldAge] Then EOBIAmt = 40 Else EOBIAmt = 0
'            EmpAge = DateDiff("M", !DOB, JDate) / 12
'
'            If ![DeptID] = "VDRS" Or EmpAge > 60 Or !ArmyRetired = True Then
'                EOBIAmt = 0
'            Else
'                EOBIAmt = 40
'            End If
            '/\/\//\/\ IF U CHANGE EOBI FORMULA HERE, REMEMBER TO CHANGE IN rptSalarySheet and rptPaySlip TOO./\/\/\/\/\
            
            ITM.SubItems(26) = EOBIAmt
            
            'Removed OverTime From Balance
            ' ITM.SubItems(27) = Round((Val(ITM.SubItems(18)) + Val(!ScholorshipAmt & "") _
            - (Payed + Val(!TotalFine & "") + UnionFund + LongLoan + LunchDed + Val(!AdvSalpaid & "") + Val(ITM.SubItems(26))) _
            + Val(ITM.SubItems(13))))
            
            ITM.SubItems(27) = Round((Val(ITM.SubItems(18)) + Val(!ScholorshipAmt & "") _
            - (Payed + Val(!TotalFine & "") + UnionFund + LongLoan + LunchDed + Val(!AdvSalpaid & "") + Val(ITM.SubItems(17)) + Val(ITM.SubItems(26)))) / 10, 0) * 10
            
            ITM.ListSubItems(15).Tag = dLateComingHrs
            ITM.ListSubItems(16).Tag = dShortHrs
            
            ITM.ListSubItems(25).Tag = Val(!Leaves & "") - Val(![WPLeaves] & "")
            ITM.ListSubItems(26).Tag = dPresentDays - Val(!Leaves & "") 'Val(!PresentDays & "") + Holidays
            
            If Val(ITM.SubItems(27)) < 0 Then
                LColor = vbRed
                bFoundNegatives = True
            Else
                LColor = LV.ForeColor
            End If
            
            GrandTotal = GrandTotal + Val(ITM.SubItems(27))
            
            Dim LastDate As Date
            LastDate = DateSerial(DTP.year, DTP.Month, LastDay)

            
            
            Dim iCounter As Integer
            
            For iCounter = 1 To ITM.ListSubItems.count
                ITM.ForeColor = LColor
                ITM.ListSubItems(iCounter).ForeColor = LColor
            Next
            ITM.ListSubItems(27).Tag = Val(!HoldAmt & "")
            
            txtDWork = dPresentDays
            txtWorkAmt = Val(ITM.SubItems(9))
            txtTax = Val(!TaxAmt & "")
            txtTAdvance = Val(!Payed & "")
            txtEOBI = EOBIAmt
            txtOvertimeDay = Val(ITM.SubItems(12))
            txtOvertimeDay = Val(txtOvertimeAmt)
            txtLBalance = lTotAmt - lClsAmt
                        
            Dim iEmpDays As Integer, iEmpMonths As Integer, iGraduityYear As Integer
            Call DateDiffInMonthsNDays(DTPicker1, DTP, iEmpMonths, iEmpDays)
            If iEmpMonths < 12 Then
                iGraduityYear = 0
            ElseIf iEmpMonths >= 12 And iEmpMonths <= 18 Then
                iGraduityYear = 1
            Else
                iGraduityYear = Round(iEmpMonths / 12)
            End If
            txtGPeriod = iGraduityYear
NEXT_EMPLOYEE:
            .MoveNEXT
        Next
        
        LV.Visible = True
        
        lblTotal = Format(GrandTotal, "#,##0.00")
        
    End With
   
    SQL = ""
    Cond = ""
    ICO = ""
    
End Sub

Private Sub UpdateFormNow()

'    Dim ITM As ListItem
'    If LV.ListItems.count = 0 Then Exit Sub
'    Set ITM = LV.ListItems(1)
'    txtDWork = Val(ITM.ListSubItems(10).Tag)
'    txtWorkAmt = Val(ITM.SubItems(12))
'    txtTax = Val(ITM.SubItems(20))
'    txtEOBI = Val(ITM.SubItems(28))
'    txtRent = Val(ITM.SubItems(26))
'    txtPF = Val(ITM.ListSubItems(13).Tag)
    Call UpdateFields
    
End Sub

Private Sub UpdateFields()

    txtDeducation = Val(txtTax) + Val(txtEOBI) + Val(txtRent) + Val(txtPF) + Val(txtOtherDed) + Val(txtTAdvance)
    txtNetPay = Val(txtWorkAmt) - Val(txtDeducation)
    txtTotalPayable = Val(txtOtherPayable) + Val(txtGAmount) + Val(txtELAmt)
    txtTotalReceivable = Val(txtOthers) + Val(txtLBalance)
    txtNetPayable = Val(txtNetPay) + Val(txtTotalPayable) - Val(txtTotalReceivable)
    
End Sub

Private Sub txtTAdvance_Change()
    UpdateFields
End Sub

Private Sub txtTax_Change()
    UpdateFields
End Sub

Private Sub RefreshLedger()
        
    Dim bFoundNegatives As Boolean
    bFoundNegatives = False
    Dim LColor As Long
    
    Dim rs As New ADODB.Recordset
    Dim rsInternal As New ADODB.Recordset
    
    With rs
        .Open "SELECT COUNT(EmpID) FROM MonthlySalaries WHERE Month(DT)=" & DTP.Month & " AND Year(DT)=" & DTP.year & " AND IsForSA=0", con, adOpenForwardOnly, adLockReadOnly
        If Val(.Fields(0) & "") = 0 Then
            IsInternalFinal = False
        Else
            IsInternalFinal = True
            'To Save Long term Loan Information....
            rsInternal.Open "SELECT EmpID,LongTerm,PrevLTLoan FROM MonthlySalaries WHERE Month(DT)=" & DTP.Month & " AND Year(DT)=" & DTP.year & " AND IsForSA=0", con, adOpenStatic, adLockReadOnly
        End If
        .Close
    End With
    
    Dim TotalSalDays As Double
    Dim TDate As Date
    Dim SDate As Date 'For Saving
    Dim Hdays As Long
    
    SDate = DateSerial(DTP.year, DTP.Month, 1)
    TDate = DateSerial(DTP.year, DTP.Month, Day(GetServerDate))
    
    TotalSalDays = DateAdd("m", 1, SDate) - SDate

    MonthDays = TotalSalDays - TotalSundays(DTP.Month, DTP.year, TotalSalDays)
    
    If chkExternal = vbUnchecked Then
        Hdays = Val(con.Execute("Select count(DT) from Holidays Where month(DT)=" & _
        DTP.Month & " and Year(DT)=" & Val(DTP.year)).Fields(0))
        MonthDays = MonthDays - Hdays
    End If
   
    If FinalizedSalarySheetLoaded Then
        Exit Sub
    Else
        If m_bShowFake Then
            MsgBox "Please finalize Original Salary Sheet.", vbInformation
            Exit Sub
        End If
    End If
    
    txtTotalHrs.Tag = txtTotalHrs

    Dim UnionFund As Integer
    Dim GrandTotal As Double
    Dim EOBIAmt As Byte
    If cmbDept.MatchFound = False Or cmbEmp.MatchFound = False Then Exit Sub
    If DTP.year.MatchFound = False Or cmbMonth.MatchFound = False Then Exit Sub
    
    Dim SunDates() As Date
        
    Dim EmpId As String, DeptID As String
    Dim EmpAge As Double
    
    Dim TempTotalSaldays As Double
    Dim LastDay As Integer
    
    Dim ITM As ListItem
    Dim Cond As String
    
    Dim cmd As New ADODB.Command
    
    Dim SalHrs As Double
    Dim HrsMultiple As Double
    
    Dim Leaves As Double
    Dim ApprovedLeaves As Double
    Dim UnApprovedLeaves As Double
    
    
'    If cmbEmp.ID = "0" Then
'       EmpId = IIf(cmbDept.ID = "0", "", cmbDept.ID)
'    Else
'       EmpId = IIf(cmbEmp.ID = "0", "", cmbEmp.ID)
'    End If
    
    EmpId = cmbEmp.ID
    DeptID = cmbDept.ID
    
    Dim rsCheck As New ADODB.Recordset
    
    Dim LunchDays As Long
    
    Dim OTByQty As Boolean
    
    'TDate = DateSerial(dtp.year, dtp.month, Day(GetServerDate))

    ''''''''''''''''''''' This Will Refresh The LateHours '''''''
    '            con.Execute "EXEC SP_CalcOverTime '" & SDate & "'", "'" & cmbEmp.ID & "'"
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    LastDay = TotalSalDays
    
    Dim JDate As Date
    
    If TDate = GetServerDate Then
        TotalSalDays = Day(GetServerDate)
    ElseIf GetServerDate < TDate Then
        MsgBox "Can't Show Entries Of Future Months.", vbInformation
        Exit Sub
    End If
    
    TempTotalSaldays = TotalSalDays     'Store For Later Use
    
    JDate = DateSerial(DTP.year, DTP.Month, TotalSalDays)
        
    'Sql = "Select * from VAdvanceSalary " & Cond & " JoinDate<='" & Format(JDate, "dd-MMM-yyyy") & "' Order By DeptID,EmpID"
    
    cmd.CommandType = adCmdStoredProc
    
    If DeptID <> "0" And EmpId = "0" Then
        cmd.CommandText = "VAdvanceSalaryDeptwise"
    Else
        cmd.CommandText = "VAdvanceSalary"
    End If
    
    Set cmd.ActiveConnection = con
    cmd.Parameters.Refresh
    
    
    cmd.Parameters("@DT").Value = JDate
    
    If DeptID <> "0" And EmpId = "0" Then
        cmd.Parameters("@EmpID").Value = DeptID
    Else
        cmd.Parameters("@EmpID").Value = IIf(EmpId = "0", "%", EmpId)
    End If
    cmd.Parameters("@ForDC").Value = True
    
    If rsLedger.State = 1 Then rsLedger.Close
    rsLedger.CursorType = adOpenForwardOnly
    rsLedger.LockType = adLockReadOnly
    
    Set rsLedger = cmd.Execute
    
    With rsLedger
        
        '.Open Sql, con, adOpenForwardOnly, adLockReadOnly
        Dim TotalSuns As Integer
        Dim SalDays As Double
        Dim Rate As Double, FakeRate As Double
        Dim AbsAllowed As Integer
        Dim AbsentDays As Double
        
        Dim total As Double
        Dim OverTimeHours As Double
        Dim OverTimeHoursSA As Double
        Dim LateHours As Double
        Dim TaxPer As Double
        Dim DedAmt As Double
        Dim Payed As Double
        Dim LongLoan As Double
        Dim TaxSal As Double
        Dim IsComputed As Boolean
        
        Dim NTime As Integer
        Dim LTime As Double
        Dim DT As String
        Dim IsPayed As Boolean
        Dim IsTransfered As Boolean
        Dim Holidays As Integer
        
        Dim EmpSalary As Double
        Dim EmpExtSal As Double

        LV.ListItems.Clear
        LV.Visible = False
        
        Dim dLateComingHrs As Double, dOTHrs As Double, dShortHrs As Double
        
        For i = 0 To .RecordCount - 1
                    
            'Reset Total Salary days For Each Employee
            
            TotalSalDays = DateAdd("m", 1, SDate) - SDate
                
            If TDate = GetServerDate Then
                TotalSalDays = Day(GetServerDate)
            End If
            
            'Check that Employee Discontinued In This Month
                        
            If Not IsNull(![DiscontinuedOn]) Then
                If Month(![DiscontinuedOn]) = Month(JDate) And ![DiscontinuedOn] <= JDate Then
                  TotalSalDays = Day(![DiscontinuedOn])
                End If
            Else
               TotalSalDays = TempTotalSaldays
            End If
            
            Leaves = Val(!Leaves & "")
                                   
           'if Joined This Month Calculate From Join Date
           If Format(![JoinDate], "MMM-yyyy") = Format(JDate, "MMM-yyyy") Then
               
               rs.Open "SELECT count(DT) FROM Holidays WHERE DT Between '" & ![JoinDate] & "' AND '" & JDate & "'", con, adOpenForwardOnly, adLockReadOnly
               
                  If rs.RecordCount > 0 Then
                     Holidays = rs.Fields(0) & ""
                  End If
               rs.Close
               Set rs = Nothing
               
               SundayAbsents = getSundayAbsents(![EmpId] & "", Val(DTP.Month), DTP.year, Val(![WeeklyAbs] & ""), Int(TotalSalDays))
               TotalSuns = TotalSundays(DTP.Month, DTP.year, TotalSalDays, Day(![JoinDate]))
               Holidays = TotalSuns + Holidays - SundayAbsents
               
            Else
            
                 rs.Open "SELECT count(DT) FROM Holidays WHERE CONVERT(CHAR(3),DT,7)='" & cmbMonth & "' AND Year(DT)=" & DTP.year & " AND day(dt)<=" & TotalSalDays, con, adOpenForwardOnly, adLockReadOnly
                  
                  If rs.RecordCount > 0 Then
                     Holidays = rs.Fields(0) & ""
                  End If
                  
                 rs.Close
               Set rs = Nothing

               'SundayAbsents = getSundayAbsents(![EmpId], dtp.month, dtp.year, Int(TotalSalDays), ![WeeklyAbs])
               
               TotalSuns = TotalSundays(DTP.Month, DTP.year, TotalSalDays)
                                              
               Holidays = TotalSuns + Holidays - SundayAbsents
               

            End If
                    
            If chkExternal = vbChecked Then
                'Exclude The Holidays From Calculation
                TotalSalDays = TotalSalDays - Holidays
            End If
            
            'Get Sundays Not To Be Paid

            'Call Computed(![EmpId], IsComputed, TaxPer, AbsAllowed, OverTimeHours, NTime, LTime, LongLoan, DT, IsPayed, IsTransfered, JDate, EmpSalary, EmpExtSal, ![Active], HrsMultiple)
            
            If Not IsComputed Then
                'Tax %
                'If IsNull(![PerDeduct]) Then TaxPer = 0 Else TaxPer = Val(![PerDeduct]) & ""
                
                'Absent Days....
                AbsAllowed = Val(![AbsAllowed] & "")
    
                ' Reset Orver Time Hours
                OverTimeHours = 0
                OverTimeHoursSA = 0
                
                'Late Hours
                NTime = Val(![NTimes] & "")
                LTime = Val(![TimeAllowed] & "")
                
                'Long Term Loan Deduction
                Dim lDedAmt As Long, lClsAmt As Long, lTotAmt As Long
                If IsNull(![DedAmt]) Then lDedAmt = 0 Else lDedAmt = ![DedAmt]
                If IsNull(![totamt]) Then lTotAmt = 0 Else lTotAmt = ![totamt]
                If IsNull(![ClsAmt]) Then lClsAmt = 0 Else lClsAmt = ![ClsAmt]
                
                
                If IsNull(!AmtToDeduct) Then
                    If lDedAmt > (lTotAmt - lClsAmt) Then
                        LongLoan = (lTotAmt - lClsAmt)
                    Else
                        LongLoan = lDedAmt
                    End If
                Else
                    LongLoan = Val(!AmtToDeduct & "")
                End If
                                      
                'Get The Salary Depending On Salary Sheet Type
                'EmpSalary = IIf(chkExternal = vbUnchecked, Val(![StartingSalary] & ""), Val(![Salary] & ""))
                
                 EmpSalary = IIf(m_bShowFake, Val(![EmpMinSalary] & ""), Val(![StartingSalary] & ""))
                 'EmpExtSal = Val(![Salary] & "")
                 EmpExtSal = EmpSalary
            End If
                                    
            'Check If Employee is paid Over time per piece
'            If Not IsNull(![OTByQty]) Then If ![OTByQty] Then OTByQty = True
                            
                If Not IsNull(![OTByQty]) Then
                    If ![OTByQty] Then               ' Calculate Over Time On the Basis of Qty Worked If Applicable
                        OTByQty = True
                        If Val(![QtyPerHour] & "") > 0 Then
                            'OTQty = Val(![TotalQtyWorked] & "") - (Val(![MinQty] & "") * SalDays)
                            OverTimeHours = Val(![OTQtyWorked] & "")
                            'OverTimeHours = OTQty / Val(![QtyPerHour] & "")
                        Else
                            OverTimeHours = 0
                        End If
                    Else
                        If IsNull(![OverTimeHours]) Then
                            OverTimeHours = 0
                            OverTimeHoursSA = 0
                        Else
                            OverTimeHours = ![OverTimeHours]
                            OverTimeHoursSA = ![OverTimeHours]
                        End If
                    End If
                    OTByQty = ![OTByQty]
                Else
                    
                    OTByQty = False
                    'OverTime Hours
                    
                    If IsNull(![OverTimeHours]) Then
                        OverTimeHours = 0
                        OverTimeHoursSA = 0
                    Else
                        OverTimeHours = ![OverTimeHours]
                        OverTimeHoursSA = ![OverTimeHours]
                    End If
                    
                End If
            '*********************
'            OverTimeHours = 0
'            OverTimeHoursSA = 0
            '*********************
            Dim theKey As String
            
            If IsComputed Then
                theKey = ![EmpId] & Format(SDate, "dd-MMM-yyyy")
            Else
                theKey = ![EmpId] & ""
            End If
            
            Set ITM = LV.ListItems.add(, theKey, ![EmpId] & "")
            
            If IsComputed Then
                ITM.Checked = IsPayed
                ITM.Bold = IsTransfered
            End If
            
            
            ITM.Tag = Val(![Emptype] & "")
            
'            If IsInternalFinal And chkExternal.value = vbChecked Then
'                rsInternal.MoveFirst
'                rsInternal.Find "EmpID='" & ![EmpId] & "'", , adSearchForward, 0
'                If rsInternal.EOF Then
'                    ITM.ListSubItems.add(, , ![Name] & "").Tag = 0
'                Else
'                    ITM.ListSubItems.add(, , ![Name] & "").Tag = Val(rsInternal![PrevLTLoan] & "")
'                End If
'            Else
'                ITM.ListSubItems.add(, , ![Name] & "").Tag = Val(![TotAmt] & "") - Val(![clsamt] & "")
'            End If
            ITM.ListSubItems.add(, , ![Name] & "").Tag = Val(![totamt] & "") - Val(![ClsAmt] & "")
            
            ITM.ListSubItems.add , , IIf(chkExternal = vbChecked, EmpExtSal, EmpSalary & "")
            
            ITM.ListSubItems(2).Tag = ![CasualLeaves] & "," & ![SickLeaves] & "," & ![AnnualLeaves] & _
            "," & ![CompensatoryLeaves] & "," & ![WPLeaves] & "," & ![MaternityLeaves]
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '   Rate = Format(Val(EmpSalary) / LastDay, "0.00")
            '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '   Calculate the Rate According To Total No ........
            '   Of Working Days in This Month for External Salary
            '
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            If chkExternal = vbChecked Then
                                
                'Rate = Format(Val(EmpSalary) / (LastDay - TotalSuns), "0.00")
                
                Rate = Format(Val(EmpSalary) / (LastDay), "0.00")
                If HrsMultiple = 0 Then HrsMultiple = 8
            Else
                
                Rate = Format(Val(EmpSalary) / LastDay, "0.00")
                
                HrsMultiple = 1
                
            End If

            Rate = Rate / HrsMultiple
                       
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '   Rate Formula For October Salary Sheet ...
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            If chkExternal Then
                If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                    Rate = EmpSalary / 240
                Else
                    Rate = EmpSalary / Val(txtTotalHrs)
                End If
                FakeRate = Val(!EmpMinSalary & "") / Val(txtTotalHrs)
            End If
                       
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
            'Store The Total Working hours for This Employee
            ITM.ListSubItems.add(, , Format(Rate, "0.00")).Tag = HrsMultiple
            
            If IsNull(![presentDays]) And !DeptID <> "VDRS" Then
                SalDays = 0
                'IF No Presents Found Exclude This Employee
                LV.ListItems.Remove ITM.key
                GoTo NEXT_EMPLOYEE
            Else
                
                
                SalDays = ((Val(![presentDays] & "")) + Holidays) + AbsAllowed
                
                'Store for Later Calculation
                
                'LunchDays = Val(![presentDays])
                
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'Exclude the Holidays now (for Vision External Salary Sheet Only)
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
                If chkExternal = vbChecked Then SalDays = ((Val(![presentDays] & ""))) + AbsAllowed
                
                'If Salary Days Exceed Total Salary Days Set Them To Salary Days
                If SalDays > TotalSalDays Then SalDays = TotalSalDays
                    
            End If
            
            Dim DaysToExclude As Integer
            
            If Format(![JoinDate], "MMM-yyyy") = Format(JDate, "MMM-yyyy") Then
               If Day(!JoinDate) = 5 Then
                    DaysToExclude = Day(![JoinDate]) - 1
                Else
                    DaysToExclude = Day(![JoinDate]) - 1
                End If
            Else
               DaysToExclude = 0
            End If
            
            
            'Deduct Penalty Days
            
            'If chkExternal = vbUnchecked Then
            '    SalDays = SalDays - Val(![AbsentDays] & "")
            'End If
            
            If SalDays < 0 Then SalDays = 0
                        
            If chkExternal Then
                Leaves = Val(!Leaves & "")
                UnApprovedLeaves = 0
            Else
                Leaves = Val(!ApprovedLeaves & "")
                UnApprovedLeaves = Val(!Leaves & "") - Val(!ApprovedLeaves & "")
            End If
            

            
            '''''''''''''''''''''''''''''''
            ' dont do anything with sunday/holiday related absents at the moment
            ''''''''''''''''''''''''''''''''
            ' SalDays = SalDays - CalcSpecialAbs(!EmpId, JDate, !BGHoliday, !AGHoliday, !BPHoliday, !APHoliday)
            '''''''''''''''''''''''''''''''
            
            SalDays = SalDays - UnApprovedLeaves
            
            If Abs(![DeductAbsents]) = 0 Then SalDays = TotalSalDays - DaysToExclude
            
            AbsentDays = TotalSalDays - SalDays - DaysToExclude
            
            
            ' Make Absents 0 if in Negative
            If AbsentDays < 0 Then AbsentDays = 0
            
            
            'If IsNull(![AbsentDays]) Then AbsentDays = 0 Else AbsentDays = Val(![AbsentDays]) / 2
            'AbsentDays = IIf(AbsentDays > 0, AbsentDays - AbsAllowed, AbsentDays)
            

            ' Convert To Hours if External Salary
            
            ITM.ListSubItems.add , , AbsentDays * HrsMultiple
            
            '''''''''''''''''' Save the Salary Hours in it tag Later '''''''''''''''''''''''
            
            
            ITM.SubItems(5) = Round(AbsentDays * HrsMultiple * Rate)
            
            'If Val(!Leaves & "") > 0 Then
            '    If AbsentDays - AbsAllowed >= 0 Then
                
            '    itm.SubItems(11) = IIf((Exempt), "0", (Leaves * Rate))
                
            '    End If
            'Else
            '    itm.SubItems(11) = "0"
            'End If
            
            ITM.SubItems(11) = "0"
            
            ' Calculat Net Absents Allowed
            ITM.SubItems(6) = AbsAllowed
            Dim NetAbsAllowed As Integer
            
            ' Calculate The Allowed Amount That is Subtracted From The Absent Amount
            
            If AbsentDays = 0 Then
                NetAbsAllowed = 0
            Else
                NetAbsAllowed = IIf(AbsAllowed - AbsentDays <= 0, AbsAllowed, AbsAllowed - AbsentDays)
            End If
            
            'Absent Allowed Amount
            ITM.SubItems(7) = IIf((Exempt), "0", Round(Rate * NetAbsAllowed * HrsMultiple, 2))
            
            
            SalHrs = SalDays * HrsMultiple
                        
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '                      Fix the Salary hours of Directors
            '               and Give Relaxation of 3 hours for october salary
            '                       Add the Leave hours as defined
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
            
            If chkExternal Then
'                If ![DeptID] = "VDRS" Then
                If Val(!EmpAttCat_RefID & "") = 6 Or Val(!EmpAttCat_RefID & "") = 5 Then
                    SalHrs = Val(txtTotalHrs)
                Else
             
                    'SalHrs = Round(Val(![HrsWorked] & ""), 1) + ((Holidays - TotalSuns) * 8)
                    SalHrs = Round(Val(![HrsWorked] & ""), 2) + ((Holidays) * 8)
                    ' Add All the Leaves to Worked hours
                    SalHrs = SalHrs + (Val(![Leaves] & "") * 8)
                    
                    'Remove Full Hours of Without Pay leaves
                    SalHrs = SalHrs - (Val(![WPLeaves] & "") * 8)
                    
                    ' Remove half Hours of Sick Leaves
                    'SalHrs = SalHrs - (Val(![SickLeaves] & "") * 4)
                    
                    ' Check if Salary Hours are greater than Total Hours Due to relaxation
                    ' Fix them to maximum hours
                    If strCompany <> "Dr-Frgz" Then
                        If SalHrs > Val(txtTotalHrs) Then SalHrs = Val(txtTotalHrs)
                    End If
                End If
            Else
                If ![DeptID] = "VDRS" Then
                    SalHrs = Val(txtTotal)
                Else
                    'Remove Full Hours of Without Pay leaves
                    SalHrs = SalHrs - Val(![WPLeaves] & "")
                    
                    'Remove half Hours of Sick Leaves
                    SalHrs = SalHrs - Val(![SickLeaves] & "") / 2
                    
                    ' Check if Salary Hours are greater than Total Hours Due to relaxation
                    ' Fix them to maximum hours
                    
                        If SalHrs > Val(txtTotal) Then SalHrs = Val(txtTotal)
                    
                    
                End If
                
            End If
            
            If !DeductionSpecialAbsents Then
                SalHrs = SalHrs - (Val(!SPAbsentDays & "") * 8)
            End If
            If SalHrs < 0 Then SalHrs = 0
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '       Show in Time Format if external salary sheet
            '       and Show days in internal salarysheet instead
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            ITM.SubItems(8) = IIf(chkExternal, getHrsMin(SalHrs), SalHrs)
            
            dLateComingHrs = Val(!LateComingHrs & "")
            If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                'If !EmpId = "DFG-000014" Then Stop
                Dim dAbsentHrs As Double, dDailyRate As Double, dAbsentDays As Double, dPresentDays As Double, dLateHrs As Double, dLateAmt As Double
                dAbsentHrs = Val(txtTotalHrs.Tag) - SalHrs
                dDailyRate = EmpSalary / 30
                dPresentDays = Val(!presentDays & "") + Holidays + Leaves - Val(![WPLeaves] & "")
                Dim iMonthDays As Integer
                iMonthDays = Day(DateSerial(DTP.year, cmbMonth.ListIndex + 2, 0))
                dAbsentDays = iMonthDays - dPresentDays
                
                Dim dAbsDaysForSalCalc As Double
                dAbsDaysForSalCalc = dAbsentDays
                If !DeductionSpecialAbsents Then
                    dPresentDays = dPresentDays - Val(!SPAbsentDays & "")
                    dAbsDaysForSalCalc = dAbsDaysForSalCalc + Val(!SPAbsentDays & "")
                    'SalHrs = SalHrs - (Val(!SPAbsentDays & "") * 8)
                End If
                
                If dPresentDays > dAbsDaysForSalCalc Then
                    ITM.SubItems(9) = Round((EmpSalary - (dDailyRate * dAbsDaysForSalCalc)) / 10, 0) * 10 'Round((SalHrs * Rate) / 10, 0) * 10
                    dPresentDays = 30 - dAbsDaysForSalCalc
                Else
                    ITM.SubItems(9) = Round((dPresentDays * dDailyRate) / 10, 0) * 10
                End If
                'ITM.SubItems(9) = Round((dPresentDays * dDailyRate) / 10, 0) * 10
                dLateHrs = (Val(iMonthDays * 8)) - SalHrs - (Val(![WPLeaves] & "") * 8)
                dLateHrs = dLateHrs - (dAbsentDays * 8)
                If !DeductionSpecialAbsents Then
                    dLateHrs = dLateHrs - (Val(!SPAbsentDays & "") * 8)
                End If
                dShortHrs = (dLateHrs - dLateComingHrs)
                If OverTimeHoursSA > 0 Then
                    If OverTimeHoursSA > dLateComingHrs Then
                        OverTimeHoursSA = OverTimeHoursSA - dLateComingHrs
                        dLateComingHrs = 0
                    Else
                        dLateComingHrs = dLateComingHrs - OverTimeHoursSA
                        OverTimeHoursSA = 0
                    End If
                End If
                
                dLateAmt = (dShortHrs + dLateComingHrs) * (Rate * (Val(!DeductionPercentage & "") / 100))
                dLateAmt = Round(dLateAmt / 10, 0) * 10
                
            Else
                ITM.SubItems(9) = Round((SalHrs * Rate) / 10, 0) * 10
            End If
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '       Save the Salary Hours in Listsubitmes(4).tag To Save
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ITM.ListSubItems(4).Tag = SalHrs
            ITM.ListSubItems(5).Tag = Val(![EmpMinSalary] & "")
            ITM.ListSubItems(6).Tag = FakeRate
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            ITM.SubItems(10) = Leaves * HrsMultiple
            
            If cmbMonth = "Sep" And DTP.year = "2004" Then
                ' Salary Sheet For September 2004 Must Be shown with the leave deduction
                ' BCoz Its on SA-8000 Record Now ...
                ITM.SubItems(11) = Round(Val(!Leaves & "") * HrsMultiple * Rate, 2)
            Else        'This Bug is Corrected Afterwards as follows
                ' Do it only For Internal Salary Sheet
                If chkExternal = vbUnchecked Then
                    LeaveDedAmt = Val(![SickLeaves] & "") * HrsMultiple * Rate / 2
                    LeaveDedAmt = LeaveDedAmt + Val(![WPLeaves] & "") * HrsMultiple * Rate
                    ITM.SubItems(11) = Round(LeaveDedAmt, 2)
                End If
                ITM.SubItems(11) = 0
            End If
            
            'If SalDays = 0 Then SalDays = 0 Else SalDays = (SalDays - AbsentDays)
            
            If OTByQty Then
                ITM.SubItems(12) = OverTimeHours
                'Over time By Quantity : Per Day Rate / Sheets Per Day * 1.4
                ITM.SubItems(13) = Round(OverTimeHours * (Rate * HrsMultiple / 56 * 1.4))
            Else
                '
                ITM.SubItems(12) = IIf(chkExternal, MinsToTime(OverTimeHoursSA * 60), MinsToTime(OverTimeHours * 60))
                ITM.ListSubItems(12).Tag = Val(!SundayOTHrs & "")
                If chkExternal Then
                    If strCompany = "QEL" Then
                        'ITM.SubItems(13) = Round(OverTimeHoursSA * (Rate + (Rate * 62.5 / 100)))
                        ITM.SubItems(13) = Round(OverTimeHoursSA * (Rate * 1.333333))
                    ElseIf strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                        ITM.SubItems(13) = Round(OverTimeHoursSA * (Rate * Val(!OTRate & "")))
                    Else
                        ITM.SubItems(13) = Round(OverTimeHoursSA * (Rate * 1.333333))
                    End If
                Else
                    ITM.SubItems(13) = Round(OverTimeHours * (Rate / 8 * 1.333333))
                End If
                If strCompany = "PAKSMITH GROUP" Then
                    ITM.ListSubItems(13).Tag = Val(!SundayOTHrs & "") * (Rate * 1.333333)
                ElseIf strCompany = "QEL" Then
                    '2.5 Hours=4;5=8; I think it's 62.5 % Increase over Standard Rate
                    'ITM.ListSubItems(13).Tag = Val(!SundayOTHrs & "") * (Rate + (Rate * 62.5 / 100))
                    ITM.ListSubItems(13).Tag = Val(!SundayOTHrs & "") * (Rate * 1.333333)
                Else
                    ITM.ListSubItems(13).Tag = Val(!SundayOTHrs & "") * (Rate * 2)
                End If
            End If
            
            LateHours = Val(![HOURS] & "") - LTAllowed
            
            If LateHours < 0 Then LateHours = 0
            
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''
            '           Exempt For October 2004
            ''''''''''''''''''''''''''''''''''''''''''''''''''
            'Well It was just written as LateHours=0 And That was why it was'nt deducting LT.
            'So I Just Commented it Out.
            'LateHours = 0
            ''''''''''''''''''''''''''''''''''''''''''''''''''
                        
            If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                ITM.SubItems(14) = MinsToTime(dLateHrs * 60)
            Else
                ITM.SubItems(14) = MinsToTime(LateHours * 60)
            End If
            
            ITM.ListSubItems(14).Tag = Val(!FixAllowance & "")
           ' itm.SubItems(15) = IIf((Exempt), "0", Format((LateHours * (Rate / Val(!WorkingHrs & ""))), "0.00"))
            
            If chkExternal Then
                If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                    ITM.SubItems(15) = dLateAmt
                Else
                    ITM.SubItems(15) = IIf((Exempt), "0", Format((LateHours * Rate), "0.00"))
                End If
            Else
                ITM.SubItems(15) = IIf((Exempt), "0", Format((LateHours * Rate / 8), "0.00"))
            End If
            
            'Total = Format(Val((ITM.SubItems(9)) + Val(ITM.ListSubItems(7))) - (Val(ITM.SubItems(15)) + Val(ITM.SubItems(11))), "0.00")
            If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                total = Format(Val((ITM.SubItems(9)) + Val(ITM.ListSubItems(7)) + Val(ITM.ListSubItems(13)) + Val(ITM.ListSubItems(13).Tag) + Val(ITM.ListSubItems(14).Tag)) - Val(ITM.SubItems(15)) - (Val(ITM.SubItems(11))), "0.00")
            Else
                total = Format(Val((ITM.SubItems(9)) + Val(ITM.ListSubItems(7)) + Val(ITM.ListSubItems(13)) + Val(ITM.ListSubItems(13).Tag) + Val(ITM.ListSubItems(14).Tag)) - (Val(ITM.SubItems(11))), "0.00")
            End If
            
            ITM.SubItems(16) = total
            
            'TaxPer = Val(![PerDeduct])
            ITM.SubItems(17) = Val(!TaxAmt & "") 'Round(getTaxAmt((EmpSalary * 12) / 1.5) / 12, 0)
            'DedAmt = Round((Val(![Salary]) / TotalSalDays) * (TaxPer / 100))
            'TaxSal = Round((Val(EmpSalary) / LastDay) * TotalSalDays)
            DedAmt = 0 'Round(TaxSal * (TaxPer / 100))
            
            'Tax Percentage That Is Zero Length
            'itm.SubItems(18) = "0"
            
            'itm.SubItems(14) = DedAmt
            ITM.SubItems(18) = total 'Total - Val(ITM.SubItems(17))
            If IsNull(![Payed]) Then Payed = 0 Else Payed = ![Payed]
            ITM.SubItems(19) = Round(Payed / 10) * 10
            
            If chkExternal.Value = vbChecked And IsInternalFinal Then
                If rsInternal.EOF Then
                    ITM.SubItems(20) = 0
                Else
                    ITM.SubItems(20) = Val(rsInternal![LongTerm] & "")
                End If
            Else
                ITM.SubItems(20) = LongLoan
                ITM.ListSubItems(20).Tag = (lTotAmt - lClsAmt)
            End If
            
            ITM.SubItems(21) = Val(![AdvSalpaid] & "")
            
            
            'Dim SocialAmt As Double
            'SocialAmt = IIf((!Social = True), !SocialAmt, 0)
            
            
            ''''''''''''''''''''''''''''''''''''''''''
            '       Union Fund deduction
            ''''''''''''''''''''''''''''''''''''''''''
            
                If Not IsNull(![UnionFund]) Then
                    If ![UnionFund] Then UnionFund = 20 Else UnionFund = 0
                Else
                    UnionFund = 0
                End If
            
            ''''''''''''''''''''''''''''''''''''''''''
                Dim LunchDed As Double
                
                If Not IsNull(![Lunch]) Then
                    LunchDed = Val(![LunchAmt] & "")
                Else
                    LunchDed = 0
                End If
            ''''''''''''''''''''''''''''''''''''''''''
            
            ITM.SubItems(22) = UnionFund
            ITM.SubItems(23) = Val(!TotalFine & "")
            
            ' Use This Column For Scholorships Instead Of Bonus ...
            'ITM.SubItems(24) = Val(!TotalBonus & "")
            
            ITM.SubItems(24) = Val(![ScholorshipAmt] & "")
            ITM.SubItems(25) = LunchDed
            
            If Not IsNull(![OldAge]) Then If ![OldAge] Then EOBIAmt = 40 Else EOBIAmt = 0
'            EmpAge = DateDiff("M", !DOB, JDate) / 12
'
'            If ![DeptID] = "VDRS" Or EmpAge > 60 Or !ArmyRetired = True Then
'                EOBIAmt = 0
'            Else
'                EOBIAmt = 40
'            End If
            '/\/\//\/\ IF U CHANGE EOBI FORMULA HERE, REMEMBER TO CHANGE IN rptSalarySheet and rptPaySlip TOO./\/\/\/\/\
            
            ITM.SubItems(26) = EOBIAmt
            
            'Removed OverTime From Balance
            ' ITM.SubItems(27) = Round((Val(ITM.SubItems(18)) + Val(!ScholorshipAmt & "") _
            - (Payed + Val(!TotalFine & "") + UnionFund + LongLoan + LunchDed + Val(!AdvSalpaid & "") + Val(ITM.SubItems(26))) _
            + Val(ITM.SubItems(13))))
            
            ITM.SubItems(27) = Round((Val(ITM.SubItems(18)) + Val(!ScholorshipAmt & "") _
            - (Payed + Val(!TotalFine & "") + UnionFund + LongLoan + LunchDed + Val(!AdvSalpaid & "") + Val(ITM.SubItems(17)) + Val(ITM.SubItems(26)))) / 10, 0) * 10
            
            ITM.ListSubItems(15).Tag = dLateComingHrs
            ITM.ListSubItems(16).Tag = dShortHrs
            
            ITM.ListSubItems(25).Tag = Val(!Leaves & "") - Val(![WPLeaves] & "")
            ITM.ListSubItems(26).Tag = dPresentDays - Val(!Leaves & "") 'Val(!PresentDays & "") + Holidays
            
            If Val(ITM.SubItems(27)) < 0 Then
                LColor = vbRed
                bFoundNegatives = True
            Else
                LColor = LV.ForeColor
            End If
            
            GrandTotal = GrandTotal + Val(ITM.SubItems(27))
            
            Dim LastDate As Date
            LastDate = DateSerial(DTP.year, DTP.Month, LastDay)

            If (IsComputed) Then
                cmdUpdate.Visible = UserHasAccess("FinalSalSheet")          'And (chkExternal = vbUnchecked)
            ElseIf Not (IsComputed) And (GetServerDate > LastDate) Then
                cmdUpdate.Visible = UserHasAccess("FinalSalSheet")          'And (chkExternal = vbUnchecked)
            Else
                cmdUpdate.Visible = False
            End If
            
            Dim iCounter As Integer
            
            For iCounter = 1 To ITM.ListSubItems.count
                ITM.ForeColor = LColor
                ITM.ListSubItems(iCounter).ForeColor = LColor
            Next
            ITM.ListSubItems(27).Tag = Val(!HoldAmt & "")
NEXT_EMPLOYEE:
            .MoveNEXT
        Next
        
        LV.Visible = True
        
        lblTotal = Format(GrandTotal, "#,##0.00")
        
    End With
    cmbMonth.Tag = cmbMonth.Text
    DTP.year.Tag = DTP.year
    
    SQL = ""
    Cond = ""
    ICO = ""
    
    If cmbDept.ID = "0" And cmbEmp.ID = "0" Then
        cmdUpdate.Caption = "Make This Sheet Final"
        cmdUpdate.Visible = UserHasAccess("FinalSalSheet")
    Else
        cmdUpdate.Visible = False
    End If
    If bFoundNegatives Then
        lblTransferNegative.Visible = True
    Else
        lblTransferNegative.Visible = False
    End If
    
End Sub


