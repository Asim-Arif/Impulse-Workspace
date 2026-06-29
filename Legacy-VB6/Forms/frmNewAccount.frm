VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmNewAccount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Account"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   7230
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkBankAcc 
      Caption         =   "Bank Account"
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
      Height          =   330
      Left            =   5520
      TabIndex        =   29
      Top             =   135
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.Frame Frame1 
      Height          =   2385
      Left            =   75
      TabIndex        =   31
      Top             =   480
      Width           =   5325
      Begin VB.CheckBox chkParent 
         Caption         =   "&Master Account"
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
         Height          =   330
         Left            =   3330
         TabIndex        =   10
         Top             =   1980
         Width           =   1635
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Sub Account Of :"
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
         Height          =   330
         Left            =   105
         TabIndex        =   2
         Top             =   660
         Width           =   1530
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   1695
         TabIndex        =   9
         Top             =   1920
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   54460419
         CurrentDate     =   37246
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Debit"
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
         Height          =   225
         Index           =   0
         Left            =   3315
         TabIndex        =   6
         Top             =   1635
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Credit"
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
         Height          =   225
         Index           =   1
         Left            =   4215
         TabIndex        =   7
         Top             =   1650
         Width           =   735
      End
      Begin VB.CheckBox chkAcrExp 
         Caption         =   "Add Related  Account"
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
         Height          =   240
         Left            =   3330
         TabIndex        =   8
         Top             =   1650
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.Label LblAccNo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1695
         TabIndex        =   4
         Top             =   1080
         Width           =   3180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Account # :"
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
         Left            =   750
         TabIndex        =   35
         Top             =   1125
         Width           =   855
      End
      Begin MSForms.ComboBox cmbSubOf 
         Height          =   330
         Left            =   1695
         TabIndex        =   3
         Top             =   660
         Width           =   3180
         VariousPropertyBits=   679495705
         DisplayStyle    =   3
         Size            =   "5609;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Opening Date :"
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
         Left            =   510
         TabIndex        =   34
         Top             =   1965
         Width           =   1095
      End
      Begin MSForms.TextBox TBox 
         Height          =   330
         Index           =   1
         Left            =   1695
         TabIndex        =   5
         Top             =   1530
         Width           =   1575
         VariousPropertyBits=   679495707
         Size            =   "2778;582"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Opening Balance :"
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
         Left            =   300
         TabIndex        =   33
         Top             =   1575
         Width           =   1305
      End
      Begin MSForms.TextBox TBox 
         Height          =   330
         Index           =   0
         Left            =   1695
         TabIndex        =   1
         Top             =   240
         Width           =   3180
         VariousPropertyBits=   679495707
         MaxLength       =   255
         Size            =   "5609;582"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Account Title :"
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
         Left            =   570
         TabIndex        =   32
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.Frame fBankInfo 
      Caption         =   "Bank Account Info ."
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
      Height          =   1815
      Left            =   75
      TabIndex        =   37
      Top             =   2895
      Visible         =   0   'False
      Width           =   5310
      Begin VB.CheckBox chkOD 
         Caption         =   "Allow Over Draft"
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
         Left            =   3330
         TabIndex        =   58
         ToolTipText     =   "Make This Account Default For Bank Transaction"
         Top             =   1110
         Width           =   1845
      End
      Begin VB.CheckBox chkDefault 
         Caption         =   "Default Bank Account"
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
         Left            =   3315
         TabIndex        =   25
         ToolTipText     =   "Make This Account Default For Bank Transaction"
         Top             =   1485
         Width           =   1845
      End
      Begin VB.TextBox txtBankAccNo 
         Height          =   315
         Left            =   1665
         MaxLength       =   25
         TabIndex        =   23
         Top             =   1005
         Width           =   1620
      End
      Begin VB.ComboBox cmbAccType 
         Height          =   315
         ItemData        =   "frmNewAccount.frx":0000
         Left            =   1665
         List            =   "frmNewAccount.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1380
         Width           =   1620
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Acc Type :"
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
         Left            =   795
         TabIndex        =   41
         Top             =   1425
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bank Acc# :"
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
         Left            =   690
         TabIndex        =   40
         Top             =   1080
         Width           =   870
      End
      Begin MSForms.ComboBox cmbBank 
         Height          =   300
         Left            =   1665
         TabIndex        =   21
         Top             =   270
         Width           =   3255
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "5741;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbBankBranch 
         Height          =   300
         Left            =   1665
         TabIndex        =   22
         Top             =   645
         Width           =   3255
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "5741;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bank Name :"
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
         Left            =   660
         TabIndex        =   39
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bank Branch :"
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
         Left            =   570
         TabIndex        =   38
         Top             =   705
         Width           =   990
      End
   End
   Begin VB.Frame fAsset 
      Caption         =   "Depriciation Account Information :"
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
      Height          =   1845
      Left            =   75
      TabIndex        =   42
      Top             =   2895
      Visible         =   0   'False
      Width           =   5310
      Begin VB.TextBox txtAccDep 
         Height          =   315
         Left            =   135
         MaxLength       =   255
         TabIndex        =   11
         Top             =   480
         Width           =   3570
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4185
         MaxLength       =   5
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   1455
         Width           =   750
      End
      Begin VB.TextBox txtdepBal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3750
         MaxLength       =   25
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   1065
         Width           =   1185
      End
      Begin VB.TextBox txtAccDepBal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3750
         MaxLength       =   25
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   480
         Width           =   1185
      End
      Begin VB.TextBox txtDep 
         Height          =   315
         Left            =   135
         MaxLength       =   255
         TabIndex        =   13
         Top             =   1065
         Width           =   3570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "%"
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
         Index           =   17
         Left            =   4980
         TabIndex        =   50
         Top             =   1515
         Width           =   195
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Yearly Depriciation Rate :"
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
         Left            =   2325
         TabIndex        =   49
         Top             =   1500
         Width           =   1830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CR"
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
         Index           =   15
         Left            =   4995
         TabIndex        =   48
         Top             =   525
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DR"
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
         Index           =   14
         Left            =   4995
         TabIndex        =   47
         Top             =   1125
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Opening Bal."
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
         Left            =   3765
         TabIndex        =   46
         Top             =   855
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Opening Bal."
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
         Left            =   3765
         TabIndex        =   45
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Depriciation Account"
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
         Left            =   120
         TabIndex        =   44
         Top             =   825
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Acc. Dep. Account Title"
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
         Left            =   150
         TabIndex        =   43
         Top             =   255
         Width           =   1680
      End
   End
   Begin VB.Frame fAcrAcc 
      Caption         =   "Related Expense Account Info :"
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
      Height          =   1830
      Left            =   75
      TabIndex        =   51
      Top             =   2880
      Visible         =   0   'False
      Width           =   5310
      Begin VB.Frame fExp 
         Height          =   510
         Left            =   120
         TabIndex        =   55
         Top             =   210
         Width           =   5055
         Begin VB.OptionButton OptExp 
            Caption         =   "Accured Expense Account"
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
            Height          =   225
            Index           =   0
            Left            =   180
            TabIndex        =   16
            Top             =   195
            Value           =   -1  'True
            Width           =   2280
         End
         Begin VB.OptionButton OptExp 
            Caption         =   "Prepaid  Expense Account"
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
            Height          =   225
            Index           =   1
            Left            =   2580
            TabIndex        =   17
            Top             =   195
            Width           =   2280
         End
      End
      Begin VB.TextBox txtExpBal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   19
         Text            =   "0.00"
         Top             =   1365
         Width           =   1095
      End
      Begin VB.TextBox txtAmt 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4020
         MaxLength       =   10
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   1365
         Width           =   1125
      End
      Begin VB.TextBox txtAcrAccTitle 
         Height          =   315
         Left            =   1200
         MaxLength       =   255
         TabIndex        =   18
         Top             =   915
         Width           =   3960
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " DR"
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
         Height          =   315
         Index           =   20
         Left            =   2280
         TabIndex        =   56
         Top             =   1365
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Opening Bal. :"
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
         Left            =   150
         TabIndex        =   54
         Top             =   1425
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Amount / Month :"
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
         Left            =   2745
         TabIndex        =   53
         Top             =   1425
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Account Title :"
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
         Left            =   135
         TabIndex        =   52
         Top             =   975
         Width           =   1035
      End
   End
   Begin MSForms.CommandButton cmdOK 
      Height          =   375
      Left            =   5505
      TabIndex        =   57
      Top             =   1027
      Width           =   1665
      Caption         =   "Save & Close  "
      PicturePosition =   327683
      Size            =   "2937;661"
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdNext 
      Height          =   375
      Left            =   5505
      TabIndex        =   26
      Top             =   570
      Width           =   1665
      Caption         =   "   Save & New"
      PicturePosition =   327683
      Size            =   "2937;661"
      Picture         =   "frmNewAccount.frx":0051
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   375
      Left            =   5505
      TabIndex        =   27
      Top             =   1485
      Width           =   1665
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2937;661"
      Picture         =   "frmNewAccount.frx":0163
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Type # :"
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
      Height          =   210
      Index           =   5
      Left            =   3690
      TabIndex        =   36
      Top             =   165
      Width           =   690
   End
   Begin VB.Label LblAccType 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4470
      TabIndex        =   28
      Top             =   105
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Acc Type :"
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
      Height          =   210
      Index           =   0
      Left            =   60
      TabIndex        =   30
      Top             =   165
      Width           =   855
   End
   Begin MSForms.ComboBox CmbType 
      Height          =   330
      Left            =   945
      TabIndex        =   0
      Top             =   105
      Width           =   2685
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "4736;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   1
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmNewAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim AccType As String
Dim AccountNo As String
Dim SubAccNo As String
Public ExpAccount As Boolean
Dim PermissionDenied As Boolean
Private Sub Check1_Click()
    If Check1.value = vbChecked Then
        cmbSubOf.Enabled = True
        Call CmbType_Click
    Else
        cmbSubOf.Enabled = False
        Call CmbType_Click
    End If
End Sub

Private Sub chkAcrExp_Click()
  If chkAcrExp.value = vbChecked Then
      Me.Height = 5190
      fAcrAcc.Visible = True
      Call OptExp_Click(0)
   Else
      Me.Height = 3345
      fAcrAcc.Visible = False
   End If
   
   Call Tbox_Change(0)
   
End Sub

Private Sub chkBankAcc_Click()
   If chkBankAcc.value = vbChecked Then
      Me.Height = 5190
      fBankInfo.Visible = True
      Call AddToCombo(cmbBank, "Distinct Bank", "BankList", "Order By Bank")
      Call CmbType_Click
      Check1 = vbChecked
      Call FindInCombo(cmbType, "CURRENT ASSETS")
   Else
      Me.Height = 3345
      fBankInfo.Visible = False
   End If
   
   Check1.Enabled = chkBankAcc <> vbChecked
   chkParent.Enabled = chkBankAcc <> vbChecked
   cmbType.Enabled = chkBankAcc <> vbChecked
   
End Sub

Private Sub chkParent_Click()
If chkParent = vbChecked Then
   TBox(1) = "0.00"
   TBox(1).Enabled = False
Else
   TBox(1).Enabled = True
End If
   Call LblAccNo_Change
End Sub

Private Sub cmbBank_Click()
Call AddToCombo(cmbBankBranch, "Branch", "BankList", "Where Bank='" & cmbBank & "' Order by Branch")
End Sub

Private Sub cmbBankBranch_Click()
    Dim BankIDRec As New ADODB.Recordset
    With BankIDRec
        .CursorLocation = adUseClient
        .Open "Select BankID from BankList Where (Bank='" & cmbBank & "') and (Branch='" & cmbBankBranch & "')", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            cmbBankBranch.Tag = .Fields(0) & ""
         Else
            cmbBankBranch.Tag = ""
        End If
        .Close
    End With
    Set BankIDRec = Nothing
End Sub

Private Sub cmbSubOf_Click()
 If cmbSubOf = "" Then Exit Sub
    cmbSubOf.Tag = getAccountNo(cmbSubOf, " trim(Type) ='" & LblAccType & "'")
    
 Call Tbox_Change(0)
    'Call GetAccountNo
End Sub

Private Sub cmbType_Change()
   If cmbType.MatchFound Then Call CmbType_Click
End Sub

Private Sub CmbType_Click()
If cmbType = "" Then Exit Sub
    Dim rec As New ADODB.Recordset
    
    AccType = con.Execute("Select Code from Heads Where Head='" & cmbType.Text & "'").Fields(0) & ""
    
    
    LblAccType.Caption = AccType
    cmbSubOf.Tag = AccType
        With rec
            .CursorLocation = adUseClient
            
            If AccType = 41 And ExpAccount Then
            ''''''''''''' Direct Expense'''''''''''''
               .Open "Select AccTitle From Accounts Where (trim(AccNo) like '41-003%') and Parent", con, adOpenForwardOnly, adLockReadOnly
            ElseIf chkBankAcc = vbChecked Then
               ''''''''''''' Bank Accounts '''''''''''''
               .Open "Select AccTitle From Accounts Where (trim(AccNo) like '15-002%') and Parent", con, adOpenForwardOnly, adLockReadOnly
            Else
               .Open "Select AccTitle From Accounts Where (Type='" & AccType & "') and Parent", con, adOpenForwardOnly, adLockReadOnly
            End If
            
            cmbSubOf.Clear
            Do Until .EOF
                cmbSubOf.AddItem .Fields(0) & ""
                .MoveNext
            Loop
            .Close
            
        End With
        Set rec = Nothing
        
        If TBox(0).Text <> "" Then LblAccNo = NextAccNo
        
        
         
        
End Sub

Private Sub FillCmbType()
    Dim rec As New ADODB.Recordset
        With rec
            .CursorLocation = adUseClient
            .Open "Select Head From Heads", con, adOpenForwardOnly, adLockReadOnly
            Do Until .EOF
                cmbType.AddItem .Fields(0) & ""
                .MoveNext
            Loop
            .Close
        End With
        Set rec = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
If IsDetailAccount Then Exit Sub
    If SaveAccount Then
        TBox(0).Text = ""
        TBox(1).Text = ""
        DTPicker1.value = ServerDate
        TBox(0).SetFocus
    End If
End Sub

Private Sub cmdOK_Click()
    If IsDetailAccount Then Exit Sub
    If SaveAccount Then Unload Me
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub cmdSClose_Click()

End Sub

Private Sub DtPicker1_Change()
    Call setDateRange(DTPicker1)
End Sub

Private Sub Form_Activate()

'Check For User Access
If PermissionDenied Then
   Me.Hide
   MsgBox "You Donot Have Permissions To Use This Option. Please Contact Your Administrator", vbInformation
   Unload Me
   Exit Sub
End If

If ExpAccount Then
   Call AddToCombo(cmbType, "Head", "Heads", "Where Trim(Code) in ('41','42','43','44')")
   
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
'PermissionDenied = Not UserHasAccess("NewAccount")
   
   ExpAccount = False
   Me.KeyPreview = True
    DTPicker1.value = ServerDate
    Call FillCmbType
End Sub

Private Function SaveAccount() As Boolean

On Error GoTo err

Dim OpBalance As Double
Dim rec As New ADODB.Recordset


If Trim(LblAccNo.Caption) = "" Then
    MsgBox "No Account #. Selected", vbInformation
    SaveAccount = False
    Exit Function
ElseIf Trim(LblAccType.Caption) = "" Then
    MsgBox "No Account Title Selected", vbInformation
    SaveAccount = False
    Exit Function
ElseIf Trim(LblAccNo) Like ACCPCHNO & "*" And (chkParent <> vbChecked) Then
    If MsgBox("This is a Purchase Account. Purchase Return Account Will Also Be Added.", vbInformation + vbOKCancel) = vbCancel Then
      SaveAccount = False
      Exit Function
    End If
ElseIf Trim(LblAccNo) Like "31*" And (Not (Trim(LblAccNo) Like SRTACC & "*")) And (chkParent <> vbChecked) Then
    If MsgBox("This is a Sale Account. Sale Return Account Will Also Be Added.", vbInformation + vbOKCancel) = vbCancel Then
      SaveAccount = False
      Exit Function
    End If

End If



If Check1.value = vbChecked Then
    If Trim(cmbSubOf.Text) = "" Then
        MsgBox "Select Parent Account First", vbInformation, "Accounts"
        SaveAccount = False
        Exit Function
    Else
        'With Rec
        '    .CursorLocation = adUseClient
        '    .Open "Select AccNo from Accounts Where (AccTitle='" & cmbSubOf.Text & "')", con, adOpenForwardOnly, adLockReadOnly
        '    SubAccNo = .Fields(0)
        '    .Close
        'End With
        SubAccNo = cmbSubOf.Tag
        If NameExists(Trim(TBox(0)), True) Then
            MsgBox "Account Title: " & TBox(0) & " already exists @ this Level..."
            Exit Function
        End If
        
    End If
Else
    SubAccNo = ""
    If NameExists(Trim(TBox(0)), False) Then
        MsgBox "Account Title: " & TBox(0) & " already exists @ this Level..."
        Exit Function
    End If
End If



'Checks For Bank Account Only

If chkBankAcc.value = vbChecked Then
   
   If Not cmbBank.MatchFound Then
      MsgBox "Invalid Bank Name", vbInformation
      Exit Function
   ElseIf Not cmbBankBranch.MatchFound Then
      MsgBox "Invalid Bank Branch Name", vbInformation
      cmbBankBranch.SetFocus
      Exit Function
   ElseIf cmbAccType.ListIndex = -1 Then
      MsgBox "Select Account Type.", vbInformation
      cmbAccType.SetFocus
      Exit Function
   ElseIf chkAcrExp = vbChecked And chkAcrExp.Visible And txtactacctitle = "" Then
      MsgBox "Invalid Accured Account Name.", vbInformation
      txtAcrAccTitle.SetFocus
      Exit Function
   End If

End If

If fAsset.Visible Then
   
   If txtDep = "" Then
      MsgBox "Invalid Account Title For Depriciation Account", vbInformation
      Exit Function
   ElseIf txtAccDep = "" Then
      MsgBox "Invalid Account Title For Accumulated Depriciation Account", vbInformation
      Exit Function
   ElseIf Val(txtRate) = 0 Then
      If MsgBox("Do You Want To Save The Account With 0.00 % Depriciation ?" _
      , vbQuestion + vbYesNo) = vbNo Then Exit Function
   ElseIf Val(txtRate) >= 100 Then
      MsgBox "Invalid Depriciation Rate.", vbInformation
      Exit Function
   End If

End If

If fAcrAcc.Visible Then
   
   If txtAcrAccTitle = "" Then
      MsgBox "Invalid Account Title For Related Expense Account", vbInformation
      Exit Function
   ElseIf Val(txtAmt) = 0 Then
      MsgBox "Monthly Amount Cannot Be Zero.", vbInformation
      Exit Function
   End If

End If

con.BeginTrans

Set rec = Nothing

If Option1(0) Then
    OpBalance = Val(TBox(1))
Else
    OpBalance = -Val(TBox(1))
End If
Dim ParentAcc As Boolean

 ParentAcc = chkParent = vbChecked
 
con.Execute "Insert into Accounts Values('" & LblAccNo & "','" & TBox(0).Text & "','" & IIf(Len(SubAccNo) = 2, "", SubAccNo) & "','" & _
 AccType & "'," & OpBalance & ",#" & Format(DTPicker1.value, "dd-MMM-yyyy") & "#," & OpBalance & "," & ParentAcc & ",True)"
 


'''''''''''''''Purchase/Sale Return Account Defining'''''''''''''''''''

Dim RetAccTitle As String

If Trim(LblAccNo) Like ACCPCHNO & "*" And (chkParent <> vbChecked) Then
   RetAccTitle = InputBox("Enter Purchase Return Account Title", "Purchase Return Account", "Purchase Return of " & TBox(0))
   Call AddPchReturnAcc(RetAccTitle)
ElseIf Trim(LblAccNo) Like "31*" And (Not (Trim(LblAccNo) Like SRTACC & "*")) And (chkParent <> vbChecked) Then
   RetAccTitle = InputBox("Enter Sale Return Account Title", "Sale Return Account", "Sale Return of " & TBox(0))
   Call AddSaleReturnAcc(RetAccTitle)
End If
 
''''''''''''''' Accrued Expense Account Defination '''''''''''''''''''
If chkAcrExp = vbChecked And chkAcrExp.Visible Then Call AddAcrExpAcc(txtAcrAccTitle)


 If chkBankAcc = vbChecked Then
   'Sleep (1000)
   Call AddBankAcc
 End If
 
 If fAsset.Visible Then Call AddDepAccs
 
 con.CommitTrans
 
 SaveAccount = True
Exit Function
SaveAccount = False
con.RollbackTrans
err:
    MsgBox err.Description
    
End Function




Private Sub Form_Unload(Cancel As Integer)
    OurOwnForm = True
End Sub

Private Sub LblAccNo_Change()
If ((Trim(LblAccNo) Like "42*") Or (Trim(LblAccNo) Like "43*") Or (Trim(LblAccNo) Like "44*") Or (Trim(LblAccNo) Like "41-003*")) And (chkParent.value <> vbChecked) Then
   Option1(0).Visible = False
   Option1(0).value = True
   Option1(1).Visible = False
   chkAcrExp.Visible = True
Else
   Option1(0).Visible = True
   Option1(1).Visible = True
   chkAcrExp.Visible = False
   chkAcrExp.value = vbUnchecked
   
   If AccType = "11" And chkParent.value <> vbChecked Then
      Me.Height = 5190
      fAsset.Visible = True
      fBankInfo.Visible = False
      Option1(1).Enabled = False
      Option1(0) = True
   ElseIf Not chkBankAcc = vbChecked Then
      Me.Height = 3345
      fAsset.Visible = False
      fBankInfo.Visible = False
      Option1(1).Enabled = True
   End If
      
      
   
End If



End Sub



Private Sub OptExp_Click(Index As Integer)
Call Tbox_Change(0)
Label1(20) = IIf(OptExp(0), " CR", " DR")
End Sub

Private Sub Tbox_Change(Index As Integer)
    If Index = 0 Then
        If Trim(TBox(0)) = "" Then
         LblAccNo = ""
        Else
           If cmbType.Text <> "" Then LblAccNo = NextAccNo
          
        End If
        
        If fAsset.Visible Then
            txtDep = "Depriciation of " & TBox(0)
            txtAccDep = "Acc. Depriciation of " & TBox(0)
         ElseIf fAcrAcc.Visible Then
            If OptExp(0) Then txtAcrAccTitle = "Accrued " & TBox(0) Else txtAcrAccTitle = "Prepaid " & TBox(0)
        End If
    End If
End Sub



Private Function NextAccNo() As String

Dim NumChars As Integer
NumChars = 3
'If Len(cmbSubOf.Tag) < 10 Then
    If Len(cmbSubOf.Tag) = 2 Then
        NextAccNo = cmbSubOf.Tag & "-" & Format(Val(GetMax("Right(Accno,3)", "Accounts", "Where Type='" & cmbSubOf.Tag & "' and (isnull(subaccof) Or len(subAccOf)=0)")) + 1, "000")
     Else
        NextAccNo = cmbSubOf.Tag & "-" & Format(Val(GetMax("Right(Accno,3)", "Accounts", "Where Subaccof='" & cmbSubOf.Tag & "'")) + 1, "000")
     End If
'Else
'    NextAccNo = cmbSubOf.Tag & "-" & Format(Val(GetMax("Right(Accno,2)", "Accounts", "Where Subaccof='" & cmbSubOf.Tag & "'")) + 1, "00")
'End If

End Function

Private Function NameExists(SName As String, SubAcc As Boolean) As Boolean

    Dim NRec As New ADODB.Recordset
    With NRec
        .CursorLocation = adUseClient
        If SubAcc Then
            .Open "Select ucase(AccTitle) from Accounts Where SubAccOf='" & cmbSubOf.Tag & "'", con, adOpenForwardOnly, adLockReadOnly
        Else
            .Open "Select ucase(AccTitle) from Accounts Where Subaccof=''  or subaccof is null", con, adOpenForwardOnly, adLockReadOnly
        End If
        Do Until .EOF
            If UCase(SName) = .Fields(0) & "" Then
                NameExists = True
                Exit Function
            End If
            .MoveNext
        Loop
        .Close
    End With
    Set NRec = Nothing
End Function


Private Sub CheckMainAcc()
   
   'Ask For the Correct Main Account

'If Trim(UCase(cmbType)) <> "CURRENT ASSETS" And Trim(cmbType) <> "" Then
'   msg = MsgBox("The Bank Account Should Be Defined As Sub Accounts Of " & AccBanksTITLE & _
'                " Under ""Current Assets.""" & vbNewLine & _
'                "Do you Want To Define The Bank Account ?", vbYesNo)
'
'  If msg = vbNo Then
'      chkBankAcc.value = vbUnchecked
'      Exit Sub
'   End If
   
'End If
   
   'Select the Banks Account
   
Call FindInCombo(cmbType, "Current Assets")

Check1.value = vbChecked
Call FindInCombo(cmbSubOf, ACCBANKSTITLE)
cmbAccType.ListIndex = 0


End Sub


Private Sub AddBankAcc()
If chkDefault.value = vbChecked Then
 con.Execute "Update BankAccounts Set [Default]=False"
 con.Execute "insert into BankAccounts(BankID,AccNo,Type,BankAccno,[Default],AllowOD) Values(" & Val(cmbBankBranch.Tag) & ",'" & LblAccNo & "','" & cmbAccType & "','" & txtBankAccNo & "',True," & CBool(chkOD = vbChecked) & ")"
Else
 con.Execute "insert into BankAccounts(BankID,AccNo,Type,BankAccno,AllowOD) Values(" & Val(cmbBankBranch.Tag) & ",'" & LblAccNo & "','" & cmbAccType & "','" & txtBankAccNo & "'," & CBool(chkOD = vbChecked) & ")"
End If
End Sub

Private Sub Tbox_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   If Index = 1 Then KeyAscii = OnlyNumber(Val(KeyAscii), True)
End Sub



Private Function IsDetailAccount() As Boolean
IsDetailAccount = True
'Allow To Define a Parent Account
If chkParent = vbChecked Then
   IsDetailAccount = False
   Exit Function
End If

Dim f As Form

   If LblAccNo Like ACCPAYNO & "*" Then
      If MsgBox("Vender Account Should Be Defined By Vender Profile." & vbNewLine & " Do You Want To Enter A New Vender ?", vbInformation + vbYesNo) = vbYes Then
         
         Set f = New frmVender
         GoTo TheEnd
         
      End If
   ElseIf LblAccNo Like ACCRECNO & "*" Then
      If MsgBox("Customer Account Should Be Defined By Customer Profile" & vbNewLine & " Do You Want To Enter A New Customer ?", vbInformation + vbYesNo) = vbYes Then
         
         'Set f = New frmCustomer
         GoTo TheEnd
         
      End If
   'ElseIf lblAccNo Like AccEmpNo & "*" Or lblAccNo Like AccDirNo & "*" Then
   '   If MsgBox("Employee/Director Account Should Be Defined By Employee Profile" & vbNewLine & " Do You Want To Enter A New Employee ?", vbInformation + vbYesNo) = vbYes Then
         
   '      Set f = New frmEmp
   '      Dim l As New frmEmp
   '      With l
   '         .Add = True
   '         .txtfields(0) = Tbox(0)
   '         .txtOBals(0) = Tbox(1)
   '         .cmbDrCr(0).ListIndex = IIf(Option1(0) = True, 1, 0)
   '         .DtOpenDate = DtPicker1
   '         Unload Me
   '         .Show 1
   '      End With
         
         
   '   End If
   ElseIf LblAccNo Like "15-002*" And chkBankAcc = vbUnchecked Then
      MsgBox "This Is A Bank Account. You Must Enter Bank Info Also", vbInformation
      chkBankAcc = vbChecked
   Else
      IsDetailAccount = False
   End If

Exit Function

TheEnd:

   With f
      .Add = True
      .txtfields(1) = TBox(0)
      .txtfields(11) = TBox(1)
      .cmbDrCr.ListIndex = IIf(Option1(0) = True, 0, 1)
      .DtOpenDate = DTPicker1
       Unload Me
      .Show 1
   End With
End Function







Private Sub AddPchReturnAcc(AccTitle As String)
   Dim NextAccNo As String
   ' "41-002" is Purchase Return Main Account
   NextAccNo = PRTACC & "-" & Format(Val(GetMax("Right(Accno,3)", "Accounts", "Where Subaccof='41-002'")) + 1, "000")
   con.Execute "Insert into Accounts Values('" & NextAccNo & "','" & AccTitle & "','" & PRTACC & "','41',0,#" & Format(DTPicker1.value, "dd-MMM-yyyy") & "#,0,False,True)"
  
End Sub

Private Sub AddSaleReturnAcc(AccTitle As String)
   Dim NextAccNo As String
   If AccTitle = "" Then AccTitle = "Sale Return Of " & TBox(0).Text
   ' "31-003" is Sale Return Main Account
   NextAccNo = SRTACC & "-" & Format(Val(GetMax("Right(Accno,3)", "Accounts", "Where Subaccof='" & SRTACC & "'")) + 1, "000")
   con.Execute "Insert into Accounts Values('" & NextAccNo & "','" & AccTitle & "','" & SRTACC & "','31',0,#" & Format(DTPicker1.value, "dd-MMM-yyyy") & "#,0,False,True)"
End Sub



Private Sub AddAcrExpAcc(AccTitle As String)
   Dim NextAccNo As String
   If AccTitle = "" Then
      If OptExp(0) Then    'Accured Expense
         AccTitle = "Accured " & TBox(0).Text
      ElseIf OptExp(1) Then
         AccTitle = "Prepaid " & TBox(0).Text
      End If
   End If
   Dim MainAcc As String
   
   ' "24-005" is Accrued Expense Main Account and  "14-001" is Prepaid Main Account
   
   If OptExp(0) Then MainAcc = "24-005" Else MainAcc = "14-001"
   'Opening Balance of Accured
   Bal = IIf(OptExp(0), -Val(txtExpBal), Val(txtExpBal))
   
   NextAccNo = MainAcc & "-" & Format(Val(GetMax("Right(Accno,3)", "Accounts", "Where Subaccof='" & MainAcc & "'")) + 1, "000")
   
   con.Execute "Insert into Accounts Values('" & NextAccNo & "','" & AccTitle & "','" & MainAcc & "','" & Left(MainAcc, 2) & "'," & Bal & ",#" & Format(DTPicker1.value, "dd-MMM-yyyy") & "#," & Bal & " ,False,True)"
   
   con.Execute "Insert into ExpenseAccounts values ('" & LblAccNo & "','" & NextAccNo & "'," & Val(txtAmt) & "," & OptExp(1) & ")"
End Sub



Private Sub txtAccDepBal_KeyPress(KeyAscii As Integer)
   KeyAscii = OnlyNumber(KeyAscii, True)
End Sub



Private Sub txtAmt_KeyPress(KeyAscii As Integer)
   KeyAscii = OnlyNumber(KeyAscii, True)
End Sub

Private Sub txtdepBal_KeyPress(KeyAscii As Integer)
   KeyAscii = OnlyNumber(KeyAscii, True)
End Sub

Private Sub AddDepAccs()
   'Accumulated Depriciation Account
   Dim SubAccOf As String
   SubAccOf = Left(LblAccNo, Len(LblAccNo) - 4)
   If Len(SubAccOf) = 2 Then SubAccOf = ""
   AccTitle = txtAccDep
   NextAcc = Left(LblAccNo, Len(LblAccNo) - 4) & "-" & Format(Val(Right(LblAccNo, 3)) + 1, "000")
   con.Execute "Insert into Accounts Values('" & NextAcc & "','" & AccTitle & "','" & SubAccOf & "','11'," & -Val(txtAccDepBal) & ",#" & Format(DTPicker1.value, "dd-MMM-yyyy") & "#," & -Val(txtAccDepBal) & ",False,True)"
   
   'Depriciation Account
   AccTitle = txtDep
   NExtDepAccno = "44-001-" & Format(Val(GetMax("Right(Accno,3)", "Accounts", "Where Trim(Subaccof)='44-001'")) + 1, "000")
   
   con.Execute "Insert into Accounts Values('" & NExtDepAccno & "','" & AccTitle & "','44-001','44'," & Val(txtdepBal) & ",#" & Format(DTPicker1.value, "dd-MMM-yyyy") & "#," & Val(txtdepBal) & ",False,True)"
   con.Execute "Insert into Assets values ('" & LblAccNo & "','" & NextAccNo & "','" & NExtDepAccno & "'," & Val(txtRate) & ")"
   
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
   KeyAscii = OnlyNumber(KeyAscii, True)
End Sub
