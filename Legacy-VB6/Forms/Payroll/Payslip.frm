VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form Payslip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pay Slip..."
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleMode       =   0  'User
   ScaleWidth      =   7349.331
   Begin VB.CheckBox chkExternal 
      Caption         =   "External Pay Slip"
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
      Left            =   6405
      TabIndex        =   49
      Top             =   630
      Visible         =   0   'False
      Width           =   2145
   End
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   2130
      TabIndex        =   3
      Top             =   945
      Width           =   1965
      _ExtentX        =   3466
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
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   195
      TabIndex        =   2
      Top             =   945
      Width           =   1950
      _ExtentX        =   3440
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
   Begin VB.Frame Frame1 
      Height          =   2700
      Left            =   45
      TabIndex        =   9
      Top             =   1290
      Width           =   8115
      Begin VB.TextBox txtAdvSal 
         Alignment       =   1  'Right Justify
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
         Left            =   3420
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   2190
         Width           =   1065
      End
      Begin VB.TextBox txtLong 
         Alignment       =   1  'Right Justify
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
         Left            =   5745
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   2190
         Width           =   1065
      End
      Begin VB.TextBox txtBal 
         Alignment       =   1  'Right Justify
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
         Left            =   6915
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   2190
         Width           =   1065
      End
      Begin VB.TextBox txtShort 
         Alignment       =   1  'Right Justify
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
         Left            =   4575
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   2190
         Width           =   1065
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
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
         Left            =   6900
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   255
         Width           =   945
      End
      Begin VB.TextBox txtNet 
         Alignment       =   1  'Right Justify
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
         Left            =   6900
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   1470
         Width           =   945
      End
      Begin VB.TextBox txtTaxAmt 
         Alignment       =   1  'Right Justify
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
         Left            =   6900
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1065
         Width           =   945
      End
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
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
         Left            =   6900
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   660
         Width           =   945
      End
      Begin VB.TextBox txtLAmt 
         Alignment       =   1  'Right Justify
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
         Left            =   5070
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   1470
         Width           =   945
      End
      Begin VB.TextBox txtLHrs 
         Alignment       =   1  'Right Justify
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
         Left            =   5070
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1065
         Width           =   945
      End
      Begin VB.TextBox txtOAmt 
         Alignment       =   1  'Right Justify
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
         Left            =   5070
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   660
         Width           =   945
      End
      Begin VB.TextBox txtOHrs 
         Alignment       =   1  'Right Justify
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
         Left            =   5070
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   255
         Width           =   945
      End
      Begin VB.TextBox txtADays 
         Alignment       =   1  'Right Justify
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
         Left            =   3210
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   255
         Width           =   945
      End
      Begin VB.TextBox txtPAmt 
         Alignment       =   1  'Right Justify
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
         Left            =   3210
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1470
         Width           =   945
      End
      Begin VB.TextBox txtPDays 
         Alignment       =   1  'Right Justify
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
         Left            =   3210
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1065
         Width           =   945
      End
      Begin VB.TextBox txtAAmt 
         Alignment       =   1  'Right Justify
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
         Left            =   3210
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   660
         Width           =   945
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
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
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1470
         Width           =   945
      End
      Begin VB.TextBox txtTotSalary 
         Alignment       =   1  'Right Justify
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
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1065
         Width           =   945
      End
      Begin VB.TextBox txtPBonus 
         Alignment       =   1  'Right Justify
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
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   660
         Width           =   945
      End
      Begin VB.TextBox txtSal 
         Alignment       =   1  'Right Justify
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
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   255
         Width           =   945
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   2040
         Picture         =   "Payslip.frx":0000
         Top             =   2160
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Adv. Salary"
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
         Left            =   3465
         TabIndex        =   51
         Top             =   1950
         Width           =   840
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   360
         Left            =   120
         TabIndex        =   48
         Top             =   2190
         Width           =   1425
         ForeColor       =   0
         Caption         =   "Print         "
         PicturePosition =   327683
         Size            =   "2514;635"
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Long (Total)"
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
         Left            =   5760
         TabIndex        =   47
         Top             =   1950
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Balance"
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
         Left            =   6930
         TabIndex        =   44
         Top             =   1950
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Short (Total)"
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
         Left            =   4590
         TabIndex        =   42
         Top             =   1950
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total :"
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
         Left            =   6405
         TabIndex        =   40
         Top             =   300
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Net Total :"
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
         Left            =   6105
         TabIndex        =   38
         Top             =   1500
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tax Amt :"
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
         Left            =   6165
         TabIndex        =   36
         Top             =   1110
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tax % :"
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
         Left            =   6285
         TabIndex        =   34
         Top             =   690
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Late Amt :"
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
         Left            =   4260
         TabIndex        =   32
         Top             =   1500
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Late Hrs :"
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
         Left            =   4305
         TabIndex        =   30
         Top             =   1110
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "O. Amt :"
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
         Left            =   4395
         TabIndex        =   28
         Top             =   690
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "O. Hrs :"
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
         Left            =   4440
         TabIndex        =   26
         Top             =   300
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Abs. Days :"
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
         Left            =   2310
         TabIndex        =   24
         Top             =   300
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Pres. Amt :"
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
         Left            =   2340
         TabIndex        =   22
         Top             =   1500
         Width           =   810
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Pres. Days :"
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
         Left            =   2265
         TabIndex        =   20
         Top             =   1110
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Abs. Amt :"
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
         Left            =   2385
         TabIndex        =   18
         Top             =   690
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Rate :"
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
         Left            =   675
         TabIndex        =   16
         Top             =   1500
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tot Salary :"
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
         Left            =   285
         TabIndex        =   14
         Top             =   1110
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "P. Bonus :"
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
         Left            =   390
         TabIndex        =   12
         Top             =   690
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Salary :"
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
         Left            =   570
         TabIndex        =   10
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   8070
      Begin Crystal.CrystalReport Cr1 
         Left            =   420
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Slip"
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
         Index           =   1
         Left            =   60
         TabIndex        =   7
         Top             =   120
         Width           =   7875
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pay Slip"
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
         Height          =   390
         Index           =   2
         Left            =   75
         TabIndex        =   8
         Top             =   135
         Width           =   7875
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Payslip.frx":0616
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Payslip.frx":0A68
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin MSForms.ComboBox cmbMonth 
      Height          =   300
      Left            =   5160
      TabIndex        =   5
      Top             =   960
      Width           =   1215
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2143;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbYear 
      Height          =   300
      Left            =   4110
      TabIndex        =   4
      Top             =   960
      Width           =   1020
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "1799;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Index           =   0
      Left            =   195
      TabIndex        =   1
      Top             =   690
      Width           =   6165
      BackColor       =   11517387
      Caption         =   "  Department                    Employee                     Year                  Month"
      Size            =   "10874;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   6420
      TabIndex        =   0
      Top             =   915
      Width           =   1695
      Caption         =   "Refresh        "
      PicturePosition =   327683
      Size            =   "2990;635"
      Accelerator     =   82
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "Payslip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    cmdRefresh.Visible = True
    cmdClose.Visible = True
    cmdPrint.Visible = True
    cmdSave.Visible = False
    cmdCancel.Visible = False
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub chkExternal_Click()
Call cmdRefresh_Click
End Sub

Private Sub cmdPrint_Click()

On Error GoTo ERR

Call ShowWait
Dim DT As Date
DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)

Dim i As Integer
Dim j As Integer
con.Execute "Delete from PrintSalary"
Dim rs As New ADODB.Recordset
With rs
    .Open "select * from PrintSalary", con, adOpenStatic, adLockOptimistic
    
        .AddNew
        ![EmpId] = cmbEmp.ID
        ![EmpName] = cmbEmp.Text
        ![BSal] = Format(txtSal, "0")
        ![Rate] = txtRate
        ![ADays] = Format(txtADays, "0")
        ![AAmt] = Format(txtAAmt, "0")
        ![SDays] = txtPDays
        ![SAmt] = Format(txtPAmt, "0")
        ![OHrs] = txtOHrs
        ![OAmt] = Format(txtOAmt, "0")
        ![LHrs] = txtLHrs
        ![LAmt] = Format(txtLAmt, "0")
        ![Total] = Format(txtTotal, "0")
        ![TAX] = txtTax
        ![TAmt] = Format(txtTaxAmt, "0")
        ![NetTtl] = Format(txtNet, "0")
        ![Paid] = Format(txtShort, "0")
        ![Balance] = Format(txtBal, "0")
        ![DT] = DT & ""
        ![LongTerm] = Format(txtLong, "0")
        ![AdvSal] = Format(Val(txtAdvSal), "0")
        .Update
    
End With

Sleep (1000)
'Dim TDate As Date
'Dim MonthDays As Integer
'TDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
'MonthDays = DateAdd("m", 1, TDate) - TDate

With cr1
    .ReportFileName = RptPRPath & "\Payslip.rpt"
    '.DataFiles(0) = DatabasePath
    .Connect = ConnectStr
    .Formulas(0) = "ForMonth='" & cmbMonth.Tag & "-" & cmbYear.Tag & "'"
    '.Formulas(1) = "ForMonthDays=" & MonthDays
    .Action = 1
    .PageZoomNext
End With

Call HideWait
Exit Sub
ERR:
    MsgBox ERR.Description
    Call HideWait
End Sub
Public Sub cmdRefresh_Click()
    Call ShowWait
    Call RefreshLedger
    Call HideWait
    
    'Call Form_Resize
End Sub



Private Sub cmdSave_Click()
    'Call SaveIt(rsLedger)
    cmdRefresh.Visible = True
    cmdClose.Visible = True
    cmdPrint.Visible = True
    cmdSave.Visible = False
    cmdCancel.Visible = False
End Sub

Private Sub cmdUpdate_Click()
    Dim setVal As Boolean
    Dim AmtClrd As Double
    For i = 1 To LV.ListItems.Count
        setVal = LV.ListItems(i).Checked
        con.Execute "Update MonthlySalaries Set Payed=" & setVal & ",LongLoan=" & LV.ListItems(i).SubItems(17) & " Where EmpID+Format(DT,'dd-MMM-yyyy')='" & LV.ListItems(i).key & "'"
        AmtClrd = Val(LV.ListItems(i).SubItems(17))
        con.Execute "Update AmtCleared Set AmtClrd=" & AmtClrd & " Where EmpID+Format(DT,'dd-MMM-yyyy')='" & LV.ListItems(i).key & "'"
    Next i
End Sub

Private Sub Form_Load()

    DtTo = ServerDate
    DtFrom = DateAdd("M", -1, ServerDate)
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    
    
    cmbMonth.AddItem "Jan"
    cmbMonth.List(0, 1) = 1
    cmbMonth.AddItem "Feb"
    cmbMonth.List(1, 1) = 2
    cmbMonth.AddItem "Mar"
    cmbMonth.List(2, 1) = 3
    cmbMonth.AddItem "Apr"
    cmbMonth.List(3, 1) = 4
    cmbMonth.AddItem "May"
    cmbMonth.List(4, 1) = 5
    cmbMonth.AddItem "Jun"
    cmbMonth.List(5, 1) = 6
    cmbMonth.AddItem "Jul"
    cmbMonth.List(6, 1) = 7
    cmbMonth.AddItem "Aug"
    cmbMonth.List(7, 1) = 8
    cmbMonth.AddItem "Sep"
    cmbMonth.List(8, 1) = 9
    cmbMonth.AddItem "Oct"
    cmbMonth.List(9, 1) = "10"
    cmbMonth.AddItem "Nov"
    cmbMonth.List(10, 1) = 11
    cmbMonth.AddItem "Dec"
    cmbMonth.List(11, 1) = 12
    
    Dim CYear As Long
    CYear = Format(ServerDate, "yyyy")
    For i = 2002 To CYear
        cmbYear.AddItem i
    Next i
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID"
    cmbDept.Text = "<All Departments>"
    
    'cmbEmp.Text = "<All Employees>"
    'Set LV.SmallIcons = ImageList1
    cmbMonth.ListIndex = Val(Format$(ServerDate, "MM")) - 1
    cmbYear = Format$(ServerDate, "yyyy")
    
    
    'Call cmdRefresh_Click
    
End Sub
Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    'cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals con, "Name", "Employees", "EmpID"
    Else
        cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "'"
    End If
    
End Sub

Private Sub Form_Resize()

On Error Resume Next
fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
'FLV.Move (Me.ScaleWidth - FLV.Width) / 2
'cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 100
'cmdPrint.Move cmdClose.Left - cmdPrint.Width - 50, cmdClose.Top

'FDrag.Move 50, Me.ScaleHeight * 3 / 5, Me.ScaleWidth - 100
 
 'If LVItems.Visible Then
 '  TheHeight = FDrag.Top
 '  FDrag.Visible = True
 'Else
 '  TheHeight = Me.ScaleHeight - cmdClose.Height - 200
 '  FDrag.Visible = False
 'End If
 
 'LV.Move 50, FAcc.Top + FAcc.Height, Me.ScaleWidth - 100, TheHeight - (FAcc.Top + FAcc.Height)
 'LVItems.Move 50, FDrag.Top + FDrag.Height, Me.ScaleWidth - 100, Me.ScaleHeight - (FAcc.Height + cmdClose.Height + LV.Height + 250)
 'FDrag.ZOrder 1
 
End Sub



Private Sub RefreshLedger()

    If cmbDept.MatchFound = False Or cmbEmp.MatchFound = False Then Exit Sub
    If cmbYear.MatchFound = False Or cmbMonth.MatchFound = False Then Exit Sub
    
    
    Dim TotalSalDays As Double
    Dim TempTotalSaldays As Double
    Dim LastDay As Integer
    
    Dim ITM As ListItem
    Dim Cond As String
    Dim Sql As String
    Dim cmd As New ADODB.Command
    Dim rsLedger As New ADODB.Recordset
    Dim SundayAbsents As Double
     
    Cond = " Where EmpID='" & cmbEmp.ID & "' and"
    
    
    Dim rsCheck As New ADODB.Recordset
    Dim TDate As Date
    Dim SDate As Date 'For Saving
    'TDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), Day(ServerDate))
    SDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    
    
    TotalSalDays = DateAdd("m", 1, SDate) - SDate
    
    LastDay = TotalSalDays
    
    If TDate = ServerDate Then
        TotalSalDays = Day(ServerDate)
    ElseIf ServerDate < TDate Then
        MsgBox "Can't Show Entries Of Future Months.", vbInformation
        Exit Sub
    End If
    
    TDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), TotalSalDays)
    TempTotalSaldays = TotalSalDays
    
    
    cmd.CommandType = adCmdStoredProc
    
    cmd.CommandText = "VAdvanceSalary"
    
    Set cmd.ActiveConnection = con
    
    cmd.Parameters.Refresh
    
    cmd.Parameters(1).Type = adDBTimeStamp
    cmd.Parameters(1).value = TDate
        
    cmd.Parameters(2).Type = adVarChar
    cmd.Parameters(2).value = cmbEmp.ID
    
    rsLedger.CursorType = adOpenForwardOnly
    rsLedger.LockType = adLockReadOnly
    
    Set rsLedger = cmd.Execute
    With rsLedger
        If .EOF Then
            MsgBox "This Employee Joined After The Date Selected", vbInformation
         Exit Sub
        End If
        '.Open Sql, Con, adOpenForwardOnly, adLockReadOnly
        Dim TotalSuns As Integer
        Dim SalDays As Double
        Dim Rate As Double
        Dim AbsAllowed As Integer
        Dim AbsentDays As Double
        
        
        Dim Total As Double
        Dim OverTimeHours As Double
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
        Dim Holidays As Integer
        Dim rs As New ADODB.Recordset
        Dim EmpSalary As Double
        
        
        
        
        'Check that Employee Discontinued In This Month
                    
               If Not IsNull(![DiscontinuedOn]) Then
                   If Month(![DiscontinuedOn]) = Month(JDate) And ![DiscontinuedOn] <= JDate Then
                     TotalSalDays = Day(![DiscontinuedOn])
                   End If
               Else
                  TotalSalDays = TempTotalSaldays
               End If
                    
                    
           
           'if Joined This Month Calculate From Join Date
           
           If Format(![JoinDate], "MMM-yyyy") = Format(JDate, "MMM-yyyy") Then
                rs.Open "Select count(DT) from Holidays Where DT Between '" & ![JoinDate] & "' and '" & JDate & "'", con, adOpenForwardOnly, adLockReadOnly
                  If rs.RecordCount > 0 Then
                     Holidays = rs.Fields(0) & ""
                  End If
                rs.Close
               Set rs = Nothing
               
               SundayAbsents = getSundayAbsents(![EmpId] & "", Val(cmbMonth.List(cmbMonth.ListIndex, 1)), cmbYear, ![WeeklyAbs], Int(TotalSalDays))
               
               TotalSuns = TotalSundays(cmbMonth.List(cmbMonth.ListIndex, 1), cmbYear, TotalSalDays, Day(![JoinDate]))
               
               Holidays = TotalSuns + Holidays - SundayAbsents
            Else
                 rs.Open "Select count(DT) from Holidays Where convert(char(3),DT,7)='" & cmbMonth & "' and Year(DT)=" & cmbYear & " and day(dt)<" & TotalSalDays, con, adOpenForwardOnly, adLockReadOnly
                  If rs.RecordCount > 0 Then
                     Holidays = rs.Fields(0) & ""
                  End If
                 rs.Close
               Set rs = Nothing
               
               
               SundayAbsents = getSundayAbsents(![EmpId], cmbMonth.List(cmbMonth.ListIndex, 1), cmbYear, Int(TotalSalDays), ![WeeklyAbs])
               
               TotalSuns = TotalSundays(cmbMonth.List(cmbMonth.ListIndex, 1), cmbYear, TotalSalDays)
                                              
               Holidays = TotalSuns + Holidays - SundayAbsents
               
               
               
            End If
                
        'LV.ListItems.Clear
        'LV.Visible = False
        'For i = 0 To .RecordCount - 1
        
        
        
           
                Call Computed(![EmpId], IsComputed, TaxPer, AbsAllowed, OverTimeHours, NTime, LTime, LongLoan, DT, IsPayed, TDate, EmpSalary)
            
                If Not IsComputed Then
                'Tax %
                If IsNull(![PerDeduct]) Then TaxPer = 0 Else TaxPer = Val(![PerDeduct]) & ""
                'Absent Days....
                AbsAllowed = Val(![AbsAllowed] & "")
                'OverTime Hours
                If IsNull(![OverTimeHours]) Then OverTimeHours = 0 Else OverTimeHours = ![OverTimeHours]
                'Late Hours
                NTime = Val(![NTimes] & "")
                LTime = Val(![TimeAllowed] & "")
                'Long Term Loan Deduction
                'Long Term Loan Deduction
                Dim lDedAmt As Long, lClsAmt As Long, lTotAmt As Long
                If IsNull(![DedAmt]) Then lDedAmt = 0 Else lDedAmt = ![DedAmt]
                If IsNull(![totamt]) Then lTotAmt = 0 Else lTotAmt = ![totamt]
                If IsNull(![clsamt]) Then lClsAmt = 0 Else lClsAmt = ![clsamt]
                
                If lDedAmt > (lTotAmt - lClsAmt) Then
                    LongLoan = (lTotAmt - lClsAmt)
                Else
                    LongLoan = lDedAmt
                End If
                
                'Get The Salary Depending On Salary Sheet Type
             'EmpSalary = IIf(chkExternal = vbUnchecked, Val(![StartingSalary] & ""), Val(![Salary] & ""))
             EmpSalary = Val(![StartingSalary] & "")
            End If
            
            
            
            Dim theKey As String
            If IsComputed Then
                theKey = ![EmpId] & Format(SDate, "dd-MMM-yyyy")
            Else
                theKey = ![EmpId] & ""
            End If
            'Set itm = LV.ListItems.Add(, theKey, ![EmpID] & "")
            
            'If IsComputed Then
            '    itm.Checked = IsPayed
            'End If
            'itm.SubItems(1) = ![Name] & ""
            'itm.SubItems(2) = empsalary & ""
         
          
         
            txtSal = Format(Val(EmpSalary & ""), "0.00")
            txtPBonus = Format(Val(![PProfit] & ""), "0.00")
            txtTotSalary = Format(Val(EmpSalary & ""), "0.00")
            Rate = Format(Val(EmpSalary) / LastDay, "0.00")
            txtRate = Rate
            
            If IsNull(![presentDays]) Then
                SalDays = 0
            Else
                SalDays = ((Val(![presentDays] & "") / 2) + Holidays) + AbsAllowed
                If SalDays > TotalSalDays Then SalDays = TotalSalDays
            End If
            
            AbsentDays = TotalSalDays - SalDays
            If absenstdays < 0 Then AbsentDays = 0
            'If IsNull(![AbsentDays]) Then AbsentDays = 0 Else AbsentDays = Val(![AbsentDays]) / 2
            'AbsentDays = IIf(AbsentDays > 0, AbsentDays - AbsAllowed, AbsentDays)
            txtADays = AbsentDays
            txtAAmt = Format(AbsentDays * Rate, "0.00")
            
            
            
            'If SalDays = 0 Then SalDays = 0 Else SalDays = (SalDays - AbsentDays)
            txtPDays = SalDays
            txtPAmt = Format(SalDays * Rate, "0.00")
            
            txtOHrs = OverTimeHours
            txtOAmt = Format(OverTimeHours * (Rate / Val(![OverTime] & "")), "0.00")
            
            LateHours = GetLateHours(![EmpId], NTime, LTime, TDate)
            
            txtLHrs = LateHours
            txtLAmt = Format((LateHours * (Rate / 8)), "0.00")
            
            Total = Format(Val(txtPAmt) + Val(txtOAmt) - Val(txtLAmt), "0.00")
            txtTotal = Total
            'TaxPer = Val(![PerDeduct])
            txtTax = TaxPer
            'DedAmt = format((Val(empsalary) / TotalSalDays) * (TaxPer / 100),,"0.00")
            TaxSal = Format((Val(EmpSalary) / LastDay) * TotalSalDays, , "0.00")
            DedAmt = Format(TaxSal * (TaxPer / 100), "0.00")
            txtTaxAmt = DedAmt
            txtNet = Total - DedAmt
            If IsNull(![Payed]) Then Payed = 0 Else Payed = ![Payed]
            txtShort = Payed
            
            txtLong = LongLoan
            
            txtAdvSal = Val(![AdvSalpaid] & "")
            
            txtBal = Format(Val(txtNet) - Payed - LongLoan - Val(![AdvSalpaid] & ""), "0")
            
            
        .Close
        'LV.Visible = True
    End With
    
    cmbMonth.Tag = cmbMonth.Text
    cmbYear.Tag = cmbYear.Text
    Set rsLedger = Nothing
    Sql = ""
    Cond = ""
    ICO = ""
End Sub
Private Function GetLateHours(EmpId As String, NTime As Integer, LTime As Double, ToDate As Date) As Double

    Dim LHours As Double
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    With cmd
        .CommandText = "VLHours"
        .CommandType = adCmdStoredProc
         Set .ActiveConnection = con
        .Parameters.Refresh
        
        .Parameters(1).Type = adNumeric
        .Parameters(1).value = LTime
        
        .Parameters(2).Type = adDBTimeStamp
        .Parameters(2).value = ToDate
                
        .Parameters(3).Type = adNumeric
        .Parameters(3).value = NTime
        
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        Set rs = .Execute
        
    End With
    
    If rs.RecordCount > 0 Then
        If IsNull(rs.Fields(1)) Then LHours = 0 Else LHours = rs.Fields(1)
    Else
        LHours = 0
    End If
    GetLateHours = LHours
    
End Function
Private Sub Computed(Emp As String, ByRef IsComputed As Boolean, ByRef TaxPer As Double, ByRef AbsentDays As Integer, ByRef OverTimeHours As Double, ByRef NTime As Integer, ByRef LTime As Double, ByRef LongLoan As Double, ByRef DT As String, ByRef Payed As Boolean, ToDate As Date, ByRef EmpSal As Double)

    Dim rsCheck As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    With cmd
      .CommandType = adCmdStoredProc
        .CommandText = "VSalarySettings"
        Set .ActiveConnection = con
        
        .Parameters.Refresh
        
        .Parameters(1).Type = adDBTimeStamp
        .Parameters(1).value = ToDate
                       
        .Parameters(2).Type = adVarChar
        .Parameters(2).value = Emp
        
        .Parameters(3).Type = adBoolean
        .Parameters(3).value = 0
        
        rsCheck.CursorType = adOpenForwardOnly
        rsCheck.LockType = adLockReadOnly
        Set rsCheck = .Execute()
    End With
    
    With rsCheck
        
        If .RecordCount > 0 Then
            IsComputed = True
            TaxPer = ![PerDeduct]
            AbsentDays = ![absentallowed]
            OverTimeHours = ![OverTime]
            NTime = ![NTimes]
            LTime = ![TimeAllowed]
            DT = Format(![DT], "dd-MMM-yyyy") & ""
            
            Payed = ![Payed] & ""
            If IsNull(![LongLoan]) Then LongLoan = 0 Else LongLoan = ![LongLoan]
            
            'Get The Salary Depending On Salary Sheet Type
             EmpSal = IIf(chkExternal = vbUnchecked, Val(![Salary] & ""), Val(![ExtSalary] & ""))
            
        Else
            IsComputed = False
            'LV.Checkboxes = False
            'cmdUpdate.Visible = False
        End If
    End With
    
End Sub

Private Sub SaveIt(rs As ADODB.Recordset, PerDeduct As Double)
    Dim DT As Date
    Dim LongLoan As Double
    If IsNull(rs![LongLoan]) Then LongLoan = 0 Else LongLoan = rs![LongLoan]
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    con.Execute "Insert into MonthlySalaries (EmpID,DT,Salary" & _
     ",PerDeduct,AbsentAllowed,TimeAllowed,NTimes,OverTime,Payed,ShortTermPer,LongLoan)" & _
     " Values('" & rs![EmpId] & "','" & _
     DT & _
     "'," & rs![Salary] & "," & PerDeduct & "," & rs![AbsAllowed] & _
     "," & rs![TimeAllowed] & "," & rs![NTimes] & "," & rs![OverTime] & ",0," & rs![ShortTermPer] & "," & LongLoan & ")"
    'Update Long Term ...
    Dim LongAmt As Double
    LongAmt = Val(LV.ListItems(LV.ListItems.Count).SubItems(17))
    
    con.Execute "Insert Into AmtCleared (EmpID,DT,AmtClrd) Values('" & rs![EmpId] & "','" & DT & "'," & LongAmt & ")"
    
    
    If LV.Checkboxes = False Then
        LV.Checkboxes = True
        cmdUpdate.Visible = True
    End If
    
End Sub
Private Function TotalMonthDays() As Integer
    Dim TDate As Date
    Dim Tot As Integer
    Dim year As Integer
    If cmbYear = "" Then
        year = 2002
    Else
        year = cmbYear
    End If
    
    If Month(ServerDate) = cmbMonth.List(cmbMonth.ListIndex, 1) Then
        Tot = Day(ServerDate)
    Else
        TDate = DateSerial(year, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
        Tot = DateAdd("m", 1, TDate) - TDate
    End If
    TotalMonthDays = Tot
End Function

Private Sub LV_DblClick()
    If LV.ListItems.Count = 0 Then Exit Sub
    If Val(LV.SelectedItem.SubItems(17)) = 0 Then Exit Sub
    Dim Amt As String
    Amt = InputBox("Please Enter Amount To Deduct For Long Term." & vbNewLine & "The Current Deduction Amount Is " & LV.SelectedItem.SubItems(17) & ".") & ""
    
    If Amt = "" Then
        
    Else
        LV.SelectedItem.SubItems(17) = Amt
    End If
    'With LV.SelectedItem
    '    txtEdit.Move LV.Left + LV.ColumnHeaders(18).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(18).Width
    '    txtEdit = .SubItems(17)
    '
    'End With
   '
   ' With txtEdit
   '     .Visible = True
   '     .SelStart = 0
   '     .SelLength = Len(.Text)
   '     .SetFocus
   ' End With
   ' 'MsgBox LV.FindItem (

End Sub
Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LV.SelectedItem.SubItems(17) = txtEdit.Text
        txtEdit.Text = ""
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Text = ""
        txtEdit.Visible = False
    Else
        KeyAscii = OnlyNumber(KeyAscii)
    End If
End Sub

