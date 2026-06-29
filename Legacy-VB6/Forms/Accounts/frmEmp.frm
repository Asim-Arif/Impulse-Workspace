VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEmp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Employee Info."
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   0
      TabIndex        =   40
      Top             =   -60
      Width           =   8700
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Info."
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
         Index           =   0
         Left            =   3390
         TabIndex        =   41
         Top             =   120
         Width           =   2310
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Info."
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
         Index           =   1
         Left            =   3405
         TabIndex        =   42
         Top             =   135
         Width           =   2310
      End
   End
   Begin VB.Frame FAll 
      Height          =   6345
      Left            =   60
      TabIndex        =   27
      Top             =   600
      Width           =   8625
      Begin VB.ComboBox cmbOverTime 
         Height          =   315
         ItemData        =   "frmEmp.frx":0000
         Left            =   7335
         List            =   "frmEmp.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Over Time"
         Top             =   990
         Width           =   1170
      End
      Begin VB.ComboBox cmbEmpType 
         Height          =   315
         ItemData        =   "frmEmp.frx":001E
         Left            =   6885
         List            =   "frmEmp.frx":002B
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   435
         Width           =   1620
      End
      Begin VB.TextBox txtfields 
         Alignment       =   1  'Right Justify
         DataField       =   "Salary"
         Height          =   315
         Index           =   7
         Left            =   6150
         MaxLength       =   20
         TabIndex        =   4
         Top             =   990
         Width           =   1155
      End
      Begin VB.Frame Frame2 
         Caption         =   "Reffered By :"
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
         Height          =   1440
         Left            =   5040
         TabIndex        =   34
         Top             =   1380
         Width           =   3480
         Begin VB.TextBox txtfields 
            DataField       =   "6"
            Height          =   300
            Index           =   6
            Left            =   1755
            MaxLength       =   20
            TabIndex        =   11
            Top             =   1035
            Width           =   1620
         End
         Begin VB.TextBox txtfields 
            DataField       =   "5"
            Height          =   300
            Index           =   5
            Left            =   75
            MaxLength       =   20
            TabIndex        =   10
            Top             =   1035
            Width           =   1620
         End
         Begin VB.TextBox txtfields 
            DataField       =   "4"
            Height          =   285
            Index           =   4
            Left            =   75
            MaxLength       =   100
            TabIndex        =   9
            Top             =   465
            Width           =   3300
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone2 / Mobile"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   13
            Left            =   1755
            TabIndex        =   37
            Top             =   825
            Width           =   1140
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   10
            Left            =   75
            TabIndex        =   36
            Top             =   825
            Width           =   555
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   21
            Left            =   90
            TabIndex        =   35
            Top             =   225
            Width           =   405
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Contact Information :"
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
         Height          =   1440
         Left            =   90
         TabIndex        =   30
         Top             =   1380
         Width           =   4920
         Begin VB.TextBox txtfields 
            DataField       =   "1"
            DataSource      =   "datPrimaryRS"
            Height          =   855
            Index           =   1
            Left            =   75
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Tag             =   "Address"
            Top             =   480
            Width           =   3225
         End
         Begin VB.TextBox txtfields 
            DataField       =   "3"
            DataSource      =   "datPrimaryRS"
            Height          =   300
            Index           =   3
            Left            =   3375
            MaxLength       =   20
            TabIndex        =   8
            Top             =   1035
            Width           =   1455
         End
         Begin VB.TextBox txtfields 
            DataField       =   "2"
            DataSource      =   "datPrimaryRS"
            Height          =   315
            Index           =   2
            Left            =   3375
            MaxLength       =   20
            TabIndex        =   7
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   33
            Top             =   255
            Width           =   585
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone2 / Mobile"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   4
            Left            =   3375
            TabIndex        =   32
            Top             =   810
            Width           =   1140
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   3
            Left            =   3375
            TabIndex        =   31
            Top             =   255
            Width           =   540
         End
      End
      Begin VB.TextBox txtfields 
         DataField       =   "EmpName"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Index           =   0
         Left            =   105
         MaxLength       =   255
         TabIndex        =   0
         Top             =   435
         Width           =   6735
      End
      Begin VB.TextBox txtfields 
         DataField       =   "EmpID"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   195
         TabIndex        =   29
         Top             =   435
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Frame FAcc 
         Caption         =   "Account Information :"
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
         Height          =   2745
         Left            =   90
         TabIndex        =   28
         Top             =   2835
         Width           =   8430
         Begin VB.TextBox txtOBals 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            DataSource      =   "datPrimaryRS"
            Height          =   315
            Index           =   5
            Left            =   4350
            MaxLength       =   10
            TabIndex        =   22
            Text            =   "0.00"
            Top             =   2235
            Width           =   1335
         End
         Begin VB.TextBox txtaccs 
            BackColor       =   &H00E0E0E0&
            DataSource      =   "datPrimaryRS"
            Enabled         =   0   'False
            Height          =   315
            Index           =   5
            Left            =   2175
            TabIndex        =   66
            Top             =   2235
            Width           =   2145
         End
         Begin VB.TextBox txtaccs 
            BackColor       =   &H00E0E0E0&
            DataSource      =   "datPrimaryRS"
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            Left            =   2175
            TabIndex        =   62
            Top             =   1890
            Width           =   2145
         End
         Begin VB.TextBox txtOBals 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            DataSource      =   "datPrimaryRS"
            Height          =   315
            Index           =   4
            Left            =   4350
            MaxLength       =   10
            TabIndex        =   20
            Text            =   "0.00"
            Top             =   1890
            Width           =   1335
         End
         Begin VB.TextBox txtOBals 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            DataSource      =   "datPrimaryRS"
            Height          =   315
            Index           =   3
            Left            =   4350
            MaxLength       =   10
            TabIndex        =   18
            Text            =   "0.00"
            Top             =   1545
            Width           =   1335
         End
         Begin VB.ComboBox cmbDrCr 
            Height          =   315
            Index           =   3
            ItemData        =   "frmEmp.frx":0059
            Left            =   5655
            List            =   "frmEmp.frx":0063
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1545
            Width           =   945
         End
         Begin VB.TextBox txtOBals 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            DataSource      =   "datPrimaryRS"
            Height          =   315
            Index           =   2
            Left            =   4350
            MaxLength       =   10
            TabIndex        =   16
            Text            =   "0.00"
            Top             =   1200
            Width           =   1335
         End
         Begin VB.ComboBox cmbDrCr 
            Height          =   315
            Index           =   2
            ItemData        =   "frmEmp.frx":0076
            Left            =   5655
            List            =   "frmEmp.frx":0080
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   1200
            Width           =   945
         End
         Begin VB.TextBox txtOBals 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            DataSource      =   "datPrimaryRS"
            Height          =   315
            Index           =   1
            Left            =   4350
            MaxLength       =   10
            TabIndex        =   14
            Text            =   "0.00"
            Top             =   855
            Width           =   1335
         End
         Begin VB.ComboBox cmbDrCr 
            Height          =   315
            Index           =   1
            ItemData        =   "frmEmp.frx":0093
            Left            =   5655
            List            =   "frmEmp.frx":009D
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   855
            Width           =   945
         End
         Begin VB.TextBox txtOBals 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            DataSource      =   "datPrimaryRS"
            Height          =   315
            Index           =   0
            Left            =   4350
            MaxLength       =   10
            TabIndex        =   12
            Text            =   "0.00"
            Top             =   510
            Width           =   1335
         End
         Begin VB.ComboBox cmbDrCr 
            Height          =   315
            Index           =   0
            ItemData        =   "frmEmp.frx":00B0
            Left            =   5655
            List            =   "frmEmp.frx":00BA
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   510
            Width           =   945
         End
         Begin VB.TextBox txtaccs 
            BackColor       =   &H00E0E0E0&
            DataField       =   "AcrAccNo"
            DataSource      =   "datPrimaryRS"
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            Left            =   2175
            TabIndex        =   56
            Top             =   1545
            Width           =   2145
         End
         Begin VB.TextBox txtaccs 
            BackColor       =   &H00E0E0E0&
            DataField       =   "LAdvAccno"
            DataSource      =   "datPrimaryRS"
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   2175
            TabIndex        =   53
            Top             =   1200
            Width           =   2145
         End
         Begin VB.TextBox txtaccs 
            BackColor       =   &H00E0E0E0&
            DataField       =   "SAdvAccno"
            DataSource      =   "datPrimaryRS"
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   2175
            TabIndex        =   50
            Top             =   855
            Width           =   2145
         End
         Begin VB.TextBox txtaccs 
            BackColor       =   &H00E0E0E0&
            DataField       =   "AccNo"
            DataSource      =   "datPrimaryRS"
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   2175
            TabIndex        =   47
            Top             =   510
            Width           =   2145
         End
         Begin VB.ComboBox cmbDrCr 
            Height          =   315
            Index           =   4
            ItemData        =   "frmEmp.frx":00CD
            Left            =   5655
            List            =   "frmEmp.frx":00D7
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1890
            Width           =   945
         End
         Begin VB.ComboBox cmbDrCr 
            Height          =   315
            Index           =   5
            ItemData        =   "frmEmp.frx":00EA
            Left            =   5655
            List            =   "frmEmp.frx":00F4
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   2235
            Width           =   945
         End
         Begin VB.Label lblAccs 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Accured Over Time"
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
            Index           =   5
            Left            =   90
            TabIndex        =   68
            Top             =   2235
            Width           =   2040
         End
         Begin VB.Label lblBals 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00 "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   315
            Index           =   5
            Left            =   6630
            TabIndex        =   67
            Top             =   2235
            Width           =   1680
         End
         Begin VB.Label lblBals 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00 "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   315
            Index           =   4
            Left            =   6630
            TabIndex        =   64
            Top             =   1890
            Width           =   1680
         End
         Begin VB.Label lblAccs 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Over Time"
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
            Index           =   4
            Left            =   90
            TabIndex        =   63
            Top             =   1890
            Width           =   2040
         End
         Begin VB.Label lblAccs 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Accrued Salary"
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
            Index           =   3
            Left            =   90
            TabIndex        =   59
            Top             =   1545
            Width           =   2040
         End
         Begin VB.Label lblBals 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00 "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   315
            Index           =   3
            Left            =   6630
            TabIndex        =   57
            Top             =   1545
            Width           =   1680
         End
         Begin VB.Label lblAccs 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Long Term Advances"
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
            Index           =   2
            Left            =   90
            TabIndex        =   55
            Top             =   1200
            Width           =   2040
         End
         Begin VB.Label lblBals 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00 "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   315
            Index           =   2
            Left            =   6630
            TabIndex        =   54
            Top             =   1200
            Width           =   1680
         End
         Begin VB.Label lblAccs 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Short Term Advances"
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
            Index           =   1
            Left            =   90
            TabIndex        =   52
            Top             =   855
            Width           =   2040
         End
         Begin VB.Label lblBals 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00 "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   315
            Index           =   1
            Left            =   6630
            TabIndex        =   51
            Top             =   855
            Width           =   1680
         End
         Begin VB.Label lblAccs 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Salary"
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
            Index           =   0
            Left            =   90
            TabIndex        =   49
            Top             =   510
            Width           =   2040
         End
         Begin VB.Label lblBals 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00 "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   315
            Index           =   0
            Left            =   6630
            TabIndex        =   48
            Top             =   510
            Width           =   1680
         End
         Begin VB.Label lblLabels 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Account Type                    Acc #                                      Opening Balance                   Current Balance"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   19
            Left            =   90
            TabIndex        =   58
            Top             =   270
            Width           =   8235
         End
      End
      Begin VB.CheckBox chkInactive 
         Caption         =   "Employee is inactive"
         DataField       =   "Inactive"
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
         Height          =   270
         Left            =   135
         TabIndex        =   24
         Tag             =   "False"
         Top             =   5955
         Width           =   2040
      End
      Begin VB.Frame FSave 
         Height          =   675
         Left            =   3540
         TabIndex        =   43
         Top             =   5565
         Width           =   4980
         Begin MSForms.CommandButton cmdCancel 
            Height          =   360
            Left            =   3315
            TabIndex        =   69
            Top             =   225
            Width           =   1545
            Caption         =   "Cancel            "
            PicturePosition =   327683
            Size            =   "2725;635"
            Picture         =   "frmEmp.frx":0107
            Accelerator     =   67
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdSaveClose 
            Height          =   360
            Left            =   1695
            TabIndex        =   26
            Top             =   225
            Width           =   1545
            Caption         =   "Save & Close"
            PicturePosition =   327683
            Size            =   "2725;635"
            Picture         =   "frmEmp.frx":0219
            Accelerator     =   83
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdSave 
            Height          =   360
            Left            =   105
            TabIndex        =   25
            Top             =   225
            Width           =   1545
            Caption         =   "Save & New  "
            PicturePosition =   327683
            Size            =   "2725;635"
            Picture         =   "frmEmp.frx":0377
            Accelerator     =   78
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin MSComCtl2.DTPicker DtOpenDate 
         DataField       =   "JDate"
         Height          =   315
         Left            =   3240
         TabIndex        =   3
         Top             =   990
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         CustomFormat    =   "dddd  dd-MMM-yyyy"
         Format          =   22740995
         CurrentDate     =   37055
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Over Time"
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
         Height          =   195
         Index           =   6
         Left            =   7320
         TabIndex        =   65
         Top             =   780
         Width           =   870
      End
      Begin MSForms.ComboBox cmbDept 
         Height          =   315
         Left            =   105
         TabIndex        =   2
         Top             =   990
         Width           =   3090
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "5450;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Departement "
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
         Height          =   195
         Index           =   15
         Left            =   120
         TabIndex        =   61
         Top             =   780
         Width           =   1185
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Height          =   195
         Index           =   2
         Left            =   6870
         TabIndex        =   60
         Top             =   195
         Width           =   420
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Joining Date :"
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
         Height          =   195
         Index           =   11
         Left            =   3255
         TabIndex        =   46
         Top             =   765
         Width           =   1140
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary "
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
         Height          =   195
         Index           =   5
         Left            =   6180
         TabIndex        =   45
         Top             =   780
         Width           =   585
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Salary Acc #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   7
         Left            =   585
         TabIndex        =   44
         Top             =   3840
         Width           =   1980
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   39
         Top             =   225
         Width           =   480
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vender ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   38
         Top             =   480
         Visible         =   0   'False
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Dim rsEmp As New ADODB.Recordset
   Public Add As Boolean, Edit As Boolean, View
   Dim AcNo As String

Private Sub chkInactive_Click()
'If chkInactive = vbChecked And Val(lblBalance) <> 0 Then
'   MsgBox "This Employee Cannot Be Made Inactive," & vbNewLine & _
'          "To Make An Employee Inactive Balance Must Be Zero", vbInformation
'   chkInactive = vbUnchecked
'End If
End Sub

Private Sub cmbDept_Change()
    If cmbDept.MatchFound Then Call cmbDept_Click
End Sub

Private Sub cmbDept_Click()
    If Not Add Then Exit Sub
    cmbDept.Tag = getAccountNo(cmbDept, " SubAccof Like '" & cmbEmpType.Tag & "%' AND Parent=1")
    Call GetNewAccNos
End Sub

Private Sub cmbDrCr_Click(Index As Integer)
    If cmbDrCr(i).ListIndex = 1 Then
        txtOBals(i).Tag = -Val(txtOBals(i))
    Else
        txtOBals(i).Tag = Val(txtOBals(i))
    End If

End Sub


Private Sub cmbEmpType_Click()
txtOBals(4).Enabled = False
cmbDrCr(4).Enabled = False
txtaccs(4) = "Not Applicable"

txtOBals(5).Enabled = False
cmbDrCr(5).Enabled = False
txtaccs(5) = "Not Applicable"

cmbOverTime.Enabled = False

If cmbEmpType.ListIndex = 0 Then
   cmbEmpType.Tag = "42-002"
ElseIf cmbEmpType.ListIndex = 1 Then
   cmbEmpType.Tag = "42-001"
ElseIf cmbEmpType.ListIndex = 2 Then
   cmbEmpType.Tag = "41-003-23001"
   txtOBals(4).Enabled = True
   cmbDrCr(4).Enabled = True
   txtOBals(5).Enabled = True
   cmbDrCr(5).Enabled = True
   cmbOverTime.Enabled = True
End If

Call AddToCombo(cmbDept, "AccTitle", "Accounts", "Where SubAccof Like '" & cmbEmpType.Tag & "%' AND Parent=1")

If Not Add Then Exit Sub
    cmbDept.Enabled = cmbDept.ListCount > 0
    cmbDept.Tag = cmbEmpType.Tag
If cmbDept.Enabled = False Then Call GetNewAccNos
End Sub



Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo err

If Trim(txtfields(0)) = "" Then
    MsgBox "Invalid Employee Name", vbInformation
    Exit Sub
ElseIf Val(txtfields(7)) = 0 Then
    MsgBox "Salary Cannot Be Zero.", vbInformation
    Exit Sub
End If




If Add Then
   If Not AddAccountInfo Then Exit Sub
ElseIf Edit Then
   Call UpdateAccinfo
End If
        
    rsEmp![EmpType] = cmbEmpType
    rsEmp![OverTime] = Val(cmbOverTime)
    ''''    if Over Time is Applied Then Save The
    ''''    Over Time Account No's Of Employee
    
    If cmbEmpType.ListIndex = 2 Then
        rsEmp![OTAcc] = txtaccs(4)
        rsEmp![AcrOTAcc] = txtaccs(5)
    End If
    
    rsEmp.Update
    
Me.Add = True

    Call Form_QueryUnload(0, 0)
    Call ClearFields
    Call Form_Activate

Exit Sub

err:
 MsgBox "ERROR : # " & err.Number & vbNewLine & err.Description, vbInformation
 
End Sub



Private Sub cmdSaveClose_Click()
On Error GoTo err

If Trim(txtfields(0)) = "" Then
    MsgBox "Invalid Employee Name", vbInformation
    Exit Sub
ElseIf Val(txtfields(7)) = 0 Then
    MsgBox "Salary Cannot Be Zero.", vbInformation
    Exit Sub
End If




If Add Then
   Call AddAccountInfo
ElseIf Edit Then
   Call UpdateAccinfo
End If

    rsEmp![EmpType] = cmbEmpType
    rsEmp![OverTime] = Val(cmbOverTime)
    
    ''''    if Over Time is Applied Then Save The
    ''''    Over Time Account No's Of Employee
    
    If cmbEmpType.ListIndex = 2 Then
        rsEmp![OTAcc] = txtaccs(4)
        rsEmp![AcrOTAcc] = txtaccs(5)
    End If
    rsEmp.Update
    
   Unload Me

Exit Sub

err:
 MsgBox "ERROR : # " & err.Number & vbNewLine & err.Description, vbInformation
 
End Sub


Private Sub DtOpenDate_Change()
    Call setDateRange(DtOpenDate)
End Sub

Private Sub Form_Activate()

    If AcNo = "" Then Unload Me
    

On Error GoTo err
With rsEmp

If .State <> 0 Then Exit Sub


 .Open "Select * From Employees Where EmpID=" & Val(Me.Tag), con, adOpenDynamic, adLockOptimistic


    For i = 0 To txtfields.Count - 1
      Set txtfields(i).DataSource = rsEmp
    Next
     
     
     For i = 0 To txtaccs.Count - 1
        Set txtaccs(i).DataSource = rsEmp
     Next


 If Add Then
   .AddNew
   'Generate New Account No's For This Employee
   cmbEmpType.Enabled = True
   cmbEmpType.ListIndex = 0
 ElseIf View Then
  'Call Lok
 ElseIf Edit Then
 
   Call GetAccInfo
   cmbEmpType.Enabled = False
   'cmbEmpType.ListIndex = IIf(![Director], 1, 0)
   
   Call FindInCombo(cmbEmpType, ![EmpType] & "")
   Call FindInCombo(cmbOverTime, ![OverTime])
   txtaccs(4) = IIf(![OTAcc] & "" = "", "Not Applicable", ![OTAcc] & "")
   txtaccs(5) = IIf(![AcrOTAcc] & "" = "", "Not Applicable", ![AcrOTAcc] & "")
   cmbDept.Enabled = False
   
   Call FindInCombo(cmbDept, con.Execute("Select AccTitle From Accounts Where AccNo='" & Left(![AccNo], Len(![AccNo]) - 4) & "'").Fields(0))
   
 End If
 
End With

Exit Sub

err:
 If err.Number = 545 Then Resume Next Else MsgBox "Error :" & err.Number & vbNewLine & err.Description, vbCritical

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()



Me.KeyPreview = True
 If Not AccDefined Then
  MsgBox "Account ""Staff Salaries"" is Not Defined, Please Define The Account First", vbInformation
 End If

txtaccs(0).Tag = "Salary  Of "
txtaccs(1).Tag = "Monthly Advances Of "
txtaccs(2).Tag = "Long Term Advances Of "
txtaccs(3).Tag = "Accrued Salary Of "
txtaccs(4).Tag = "Over Time of "
txtaccs(5).Tag = "Accrued Over Time Of "

DtOpenDate = Date
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

 Set rsEmp = Nothing
End Sub


Private Sub showdata()

End Sub




Private Sub GetAccInfo()
   


Dim rsAcc As New ADODB.Recordset

   For i = 0 To txtaccs.Count - 1
      With rsAcc
         .Open "Select OpenBal,Balance,Active from Accounts Where AccNo = '" & txtaccs(i) & "'", con, adOpenForwardOnly, adLockReadOnly
         
           If .RecordCount > 0 Then
            
            If i = 0 Then chkInactive = IIf(![Active], vbUnchecked, vbChecked)
            
            txtOBals(i) = Format(Abs(Val(![OpenBal] & "")), "0.00")
             txtOBals(i).Tag = Val(![OpenBal] & "")
              cmbDrCr(i).ListIndex = IIf(Val(txtOBals(i).Tag) > 0, 0, IIf(Val(txtOBals(i).Tag) < 0, 1, -1))
              ' Remember Previous Opening Balance
             cmbDrCr(i).Tag = txtOBals(i).Tag
            lblBals(i) = Format(Val(![Balance]), "0.00 DR;0.00 CR;0.00")
           End If
         .Close
      End With
   Next

Set rsAcc = Nothing

End Sub


Private Function AddAccountInfo() As Boolean
 On Error GoTo err
 Dim i As Integer
 Dim Active As Boolean
 Active = chkInactive <> vbChecked
 con.BeginTrans
 For i = 0 To txtaccs.Count - 1
   Call cmbDrCr_Click(i)
   If InStr(txtaccs(i), "-") > 0 Then
        con.Execute "Insert into Accounts(AccNo,AccTitle,SubAccOf,Type,OpenDate,OpenBal,Balance,Active) Values ('" & _
               txtaccs(i) & "','" & txtaccs(i).Tag & txtfields(0) & "','" & Left(txtaccs(i), Len(txtaccs(i)) - 4) & "','" & _
               Left(txtaccs(i), 2) & "','" & DtOpenDate & "'," & Val(txtOBals(i).Tag) & "," & Val(txtOBals(i).Tag) & "," & Active & ")"
   End If
 Next
 con.CommitTrans
 AddAccountInfo = True
 Exit Function
 
err:
   MsgBox "Cannot Add Employee Account." & vbNewLine & err.Number & vbNewLine & err.Description, vbCritical
   AddAccountInfo = False
   con.RollbackTrans
 
End Function
Private Function AccDefined() As Boolean
'Get Accounts Receiveable AccountNo
AcNo = ACCEMPNO
If AcNo = "" Then
   AccDefined = False
 Else
  AccDefined = True
 End If
End Function

Private Function UpdateAccinfo() As Boolean
 Dim i As Integer
 Dim Diff As Double
 Dim Active As Boolean
con.BeginTrans
For i = 0 To txtaccs.Count - 1
   Call cmbDrCr_Click(i)
   Diff = Val(txtOBals(i)) - Val(cmbDrCr(i).Tag)
   Active = chkInactive <> vbChecked
   con.Execute "Update Accounts Set accTitle='" & txtaccs(i).Tag & txtfields(0) & "',OpenBal=" & Val(txtOBals(i)) & ",Balance=balance + " & Diff & ",Active=" & Active & " Where Accno='" & txtaccs(i) & "'"
   Call adjustLedger(txtaccs(i), DtOpenDate, 0, Diff, True)
Next
con.CommitTrans
   UpdateAccinfo = True
   Exit Function
err:
   MsgBox "Acccount Info Cannot Be Updated Due To Error:" & err.Number & vbNewLine & err.Description
UpdateAccinfo = False

End Function



Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 9 Or Index = 8 Then KeyAscii = OnlyNumber(KeyAscii, True)
End Sub


Private Sub ClearFields()
   For i = 0 To txtfields.Count - 1
      txtfields(i) = ""
   Next
End Sub


Private Sub GetNewAccNos()
Dim AccEmp As String, AcNo As String
'Generate Employee Account #
 Dim NewVal As String
 
 'AccEmp = AccEmpNo
 
 AccEmp = cmbDept.Tag
 
 NewVal = GetMax("right(AccNo,3)", "Accounts", "Where SubAccOf='" & AccEmp & "'")
 txtaccs(0) = AccEmp & "-" & Format(Val(NewVal) + 1, "000")
 
 Dim SAdvAcc As String, LAdvAcc As String, AcrAcc As String
 
 'Get Account No's According To Directors Or Staff Salaries
 If cmbEmpType.ListIndex = 1 Then
   'Director Accounts
   SAdvAcc = "15-010-002"
   LAdvAcc = "13-002"
   AcrAcc = "24-005-001"
 Else
   'Production Staff/Staff Member Accounts
   SAdvAcc = "15-010-003"
   LAdvAcc = "13-003"
   AcrAcc = "24-005-002"
   
 End If
 
 'Short Term Advance Account
 NewVal = GetMax("right(AccNo,3)", "Accounts", "Where SubAccOf='" & SAdvAcc & "'")
txtaccs(1) = SAdvAcc & "-" & Format(Val(NewVal) + 1, "000")
 'Long Term Advance Account
 NewVal = GetMax("right(AccNo,3)", "Accounts", "Where SubAccOf='" & LAdvAcc & "'")
txtaccs(2) = LAdvAcc & "-" & Format(Val(NewVal) + 1, "000")
 'Accrued salaries
 NewVal = GetMax("right(AccNo,3)", "Accounts", "Where SubAccOf='" & AcrAcc & "'")
 txtaccs(3) = AcrAcc & "-" & Format(Val(NewVal) + 1, "000")
 
 If cmbEmpType.ListIndex = 2 Then
    
    Dim OverTimeAcc As String, AcrOverTimeAcc As String
     OverTimeAcc = "41-003-002"
     'Over Time Account
     NewVal = GetMax("right(AccNo,3)", "Accounts", "Where SubAccOf='" & OverTimeAcc & "'")
     txtaccs(4) = OverTimeAcc & "-" & Format(Val(NewVal) + 1, "000")
     
    'Accured Over Time Account
    AcrOverTimeAcc = "24-005-003"
    
    NewVal = GetMax("right(AccNo,3)", "Accounts", "Where SubAccOf='" & AcrOverTimeAcc & "'")
    txtaccs(5) = AcrOverTimeAcc & "-" & Format(Val(NewVal) + 1, "000")
 End If
 
End Sub

Private Sub txtOBals_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = OnlyNumber(KeyAscii, True)
End Sub
