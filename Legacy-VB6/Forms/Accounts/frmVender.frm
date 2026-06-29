VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmVender 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vendor Info."
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Please Select Employee's Picture"
      Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
   End
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   15
      TabIndex        =   40
      Top             =   0
      Width           =   8700
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Info."
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
         Width           =   1950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Info."
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
         Width           =   1950
      End
   End
   Begin VB.Frame FAll 
      Height          =   6435
      Left            =   60
      TabIndex        =   19
      Top             =   600
      Width           =   8625
      Begin VB.TextBox txtfields 
         DataField       =   "VenderDescription"
         DataSource      =   "datPrimaryRS"
         Height          =   465
         Index           =   16
         Left            =   5055
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   60
         Tag             =   "Address"
         Top             =   5265
         Width           =   3480
      End
      Begin VB.TextBox txtUrduFields 
         Alignment       =   1  'Right Justify
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   0
         Left            =   90
         TabIndex        =   56
         Top             =   5265
         Width           =   4920
      End
      Begin VB.CheckBox chkInactive 
         Caption         =   "Vender is inactive"
         DataField       =   "Inactive"
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
         Height          =   270
         Left            =   3330
         TabIndex        =   12
         Tag             =   "False"
         Top             =   3000
         Width           =   1560
      End
      Begin VB.TextBox txtfields 
         DataField       =   "MakerNo"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   14
         Left            =   5040
         MaxLength       =   255
         TabIndex        =   50
         Top             =   420
         Width           =   1380
      End
      Begin VB.CheckBox chkImport 
         Caption         =   "Import Vender"
         DataField       =   "ImportVender"
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
         Height          =   270
         Left            =   3345
         TabIndex        =   49
         Tag             =   "False"
         Top             =   3270
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Contact Person :"
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
         Height          =   1905
         Left            =   5055
         TabIndex        =   33
         Top             =   780
         Width           =   3480
         Begin VB.TextBox txtfields 
            DataField       =   "Mobile"
            Height          =   300
            Index           =   13
            Left            =   1755
            MaxLength       =   20
            TabIndex        =   11
            Top             =   1515
            Width           =   1620
         End
         Begin VB.TextBox txtfields 
            DataField       =   "CPhone"
            Height          =   300
            Index           =   12
            Left            =   75
            MaxLength       =   20
            TabIndex        =   10
            Top             =   1500
            Width           =   1620
         End
         Begin VB.TextBox txtfields 
            DataField       =   "CEmail"
            Height          =   285
            Index           =   9
            Left            =   75
            MaxLength       =   100
            TabIndex        =   9
            Top             =   975
            Width           =   3300
         End
         Begin VB.TextBox txtfields 
            DataField       =   "ContactPerson"
            Height          =   285
            Index           =   5
            Left            =   75
            MaxLength       =   100
            TabIndex        =   8
            Top             =   465
            Width           =   3300
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile"
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
            Top             =   1290
            Width           =   450
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
            Top             =   1290
            Width           =   555
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail"
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
            Index           =   5
            Left            =   90
            TabIndex        =   35
            Top             =   750
            Width           =   420
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
            Left            =   105
            TabIndex        =   34
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
         Height          =   2790
         Left            =   90
         TabIndex        =   26
         Top             =   780
         Width           =   4920
         Begin VB.TextBox txtfields 
            DataField       =   "BankAccNo"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   15
            Left            =   90
            MaxLength       =   100
            TabIndex        =   52
            Tag             =   "Contact Person Name"
            Top             =   2415
            Width           =   3060
         End
         Begin VB.TextBox txtfields 
            DataField       =   "Fax1"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   2
            Left            =   90
            MaxLength       =   100
            TabIndex        =   5
            Tag             =   "Contact Person Name"
            Top             =   990
            Width           =   2370
         End
         Begin VB.TextBox txtfields 
            DataField       =   "Address"
            DataSource      =   "datPrimaryRS"
            Height          =   660
            Index           =   8
            Left            =   90
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Tag             =   "Address"
            Top             =   1470
            Width           =   4740
         End
         Begin VB.TextBox txtfields 
            DataField       =   "Fax2"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   7
            Left            =   2505
            MaxLength       =   100
            TabIndex        =   6
            Top             =   990
            Width           =   2325
         End
         Begin VB.TextBox txtfields 
            DataField       =   "Phone3"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   6
            Left            =   3270
            MaxLength       =   20
            TabIndex        =   4
            Top             =   450
            Width           =   1560
         End
         Begin VB.TextBox txtfields 
            DataField       =   "Phone2"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   4
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   3
            Top             =   450
            Width           =   1560
         End
         Begin VB.TextBox txtfields 
            DataField       =   "Phone1"
            DataSource      =   "datPrimaryRS"
            Height          =   285
            Index           =   3
            Left            =   90
            MaxLength       =   20
            TabIndex        =   2
            Top             =   450
            Width           =   1560
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Account No. :"
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
            Index           =   17
            Left            =   120
            TabIndex        =   53
            Top             =   2190
            Width           =   1380
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
            TabIndex        =   32
            Top             =   1275
            Width           =   585
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NIC No. :"
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
            Index           =   8
            Left            =   2505
            TabIndex        =   31
            Top             =   765
            Width           =   675
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax :"
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
            Index           =   6
            Left            =   3285
            TabIndex        =   30
            Top             =   255
            Width           =   375
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone2"
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
            Left            =   1710
            TabIndex        =   29
            Top             =   240
            Width           =   540
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
            Left            =   120
            TabIndex        =   28
            Top             =   225
            Width           =   540
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NTN No. :"
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
            Index           =   2
            Left            =   120
            TabIndex        =   27
            Top             =   765
            Width           =   705
         End
      End
      Begin VB.TextBox txtfields 
         DataField       =   "Name"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   1
         Left            =   165
         MaxLength       =   255
         TabIndex        =   0
         Top             =   420
         Width           =   4830
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
         Height          =   1440
         Left            =   90
         TabIndex        =   20
         Top             =   3555
         Width           =   4920
         Begin VB.ComboBox cmbDrCr 
            Height          =   315
            ItemData        =   "frmVender.frx":0000
            Left            =   1545
            List            =   "frmVender.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1020
            Width           =   1110
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AccNo"
            DataSource      =   "datPrimaryRS"
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   135
            TabIndex        =   21
            Top             =   450
            Width           =   1560
         End
         Begin VB.TextBox txtfields 
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
            Index           =   11
            Left            =   165
            MaxLength       =   10
            TabIndex        =   14
            Top             =   1020
            Width           =   1350
         End
         Begin MSComCtl2.DTPicker DtOpenDate 
            Height          =   315
            Left            =   1770
            TabIndex        =   13
            Top             =   450
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "MMMM dd/MM/yyyy"
            Format          =   301072387
            CurrentDate     =   37055
         End
         Begin VB.Label lblBalance 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00 "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   315
            Left            =   2700
            TabIndex        =   44
            Top             =   1020
            Width           =   1515
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Balance"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   14
            Left            =   2685
            TabIndex        =   43
            Top             =   795
            Width           =   1380
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Acc # :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   7
            Left            =   135
            TabIndex        =   24
            Top             =   225
            Width           =   645
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Opening Bal."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   12
            Left            =   150
            TabIndex        =   23
            Top             =   795
            Width           =   1110
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Opening Date :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   11
            Left            =   1755
            TabIndex        =   22
            Top             =   225
            Width           =   1305
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         Height          =   630
         Left            =   3255
         TabIndex        =   46
         Top             =   5700
         Width           =   5280
         Begin MSForms.CommandButton cmdSaveClose 
            Height          =   360
            Left            =   1800
            TabIndex        =   17
            Top             =   180
            Width           =   1665
            Caption         =   "Save & Close  "
            PicturePosition =   327683
            Size            =   "2937;635"
            Accelerator     =   83
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
            Left            =   3510
            TabIndex        =   18
            Top             =   195
            Width           =   1665
            Caption         =   "Cancel            "
            PicturePosition =   327683
            Size            =   "2937;635"
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
            Height          =   360
            Left            =   75
            TabIndex        =   16
            Top             =   195
            Width           =   1665
            Caption         =   "Save & New   "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmVender.frx":001D
            Accelerator     =   110
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Vender Of :"
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
         Height          =   1785
         Left            =   1440
         TabIndex        =   47
         Top             =   3210
         Visible         =   0   'False
         Width           =   3480
         Begin MSComctlLib.ListView LV 
            Height          =   1470
            Left            =   60
            TabIndex        =   48
            Top             =   255
            Width           =   3330
            _ExtentX        =   5874
            _ExtentY        =   2593
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            HideColumnHeaders=   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
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
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Catagory"
               Object.Width           =   5292
            EndProperty
         End
      End
      Begin VB.TextBox txtfields 
         DataField       =   "VendID"
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   195
         TabIndex        =   25
         Top             =   435
         Visible         =   0   'False
         Width           =   1410
      End
      Begin MSForms.ComboBox cmbMasterCat 
         Height          =   300
         Left            =   90
         TabIndex        =   65
         Top             =   6030
         Visible         =   0   'False
         Width           =   3120
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "5503;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vendor Catagory"
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
         Left            =   135
         TabIndex        =   64
         Top             =   5835
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label lblPicSig 
         Alignment       =   2  'Center
         Caption         =   "Double Click To Add Picture."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   7080
         TabIndex        =   63
         Top             =   4035
         Width           =   1095
      End
      Begin VB.Label lblPic 
         Alignment       =   2  'Center
         Caption         =   "Double Click To Add Picture."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   5430
         TabIndex        =   62
         Top             =   4035
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vender Description :"
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
         Index           =   20
         Left            =   5070
         TabIndex        =   61
         Top             =   5055
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Signature"
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
         Left            =   6840
         TabIndex        =   59
         Top             =   3195
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Picture"
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
         Left            =   5130
         TabIndex        =   58
         Top             =   3195
         Width           =   495
      End
      Begin VB.Image ImgSig 
         BorderStyle     =   1  'Fixed Single
         Height          =   1590
         Left            =   6825
         Stretch         =   -1  'True
         Top             =   3405
         Width           =   1680
      End
      Begin VB.Image ImgPic 
         BorderStyle     =   1  'Fixed Single
         Height          =   1590
         Left            =   5115
         Stretch         =   -1  'True
         Top             =   3405
         Width           =   1680
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Urdu Name :"
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
         Index           =   19
         Left            =   90
         TabIndex        =   57
         Top             =   5040
         Width           =   900
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Process"
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
         Index           =   18
         Left            =   5130
         TabIndex        =   55
         Top             =   2670
         Visible         =   0   'False
         Width           =   555
      End
      Begin MSForms.ComboBox cmbProcess 
         DataField       =   "ProcessID"
         Height          =   315
         Left            =   5115
         TabIndex        =   54
         Top             =   2880
         Visible         =   0   'False
         Width           =   3405
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "6006;556"
         BoundColumn     =   0
         ColumnCount     =   2
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. "
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
         Index           =   16
         Left            =   5055
         TabIndex        =   51
         Top             =   180
         Width           =   300
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vendor Type"
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
         Left            =   6465
         TabIndex        =   45
         Top             =   180
         Width           =   1065
      End
      Begin MSForms.ComboBox cmbType 
         Height          =   300
         Left            =   6450
         TabIndex        =   1
         Top             =   420
         Width           =   2055
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "3625;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company Name"
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
         Left            =   180
         TabIndex        =   39
         Top             =   180
         Width           =   1320
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vender ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
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
Attribute VB_Name = "frmVender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsVend As New ADODB.Recordset
Public add As Boolean, edit As Boolean, View
Dim Acno As String
Dim PicFileName As String, ThumbFileName As String

Private Sub chkInactive_Click()
    If chkInactive = vbChecked And Val(lblBalance) <> 0 Then
       MsgBox "This Vender Cannot Be Made Inactive," & vbNewLine & _
                "To Make A Vender Inactive Balance Must Be Zero", vbInformation
       chkInactive = vbUnchecked
    End If
End Sub

Private Sub cmbDrCr_Click()
    If cmbDrCr.ListIndex = 1 Then
        txtfields(11).Tag = -Val(txtfields(11))
    Else
        txtfields(11).Tag = Val(txtfields(11))
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err

    If Trim(txtfields(1)) = "" Then
        MsgBox "Invalid Vender Name", vbInformation
        Exit Sub
    ElseIf cmbType.Enabled And Not cmbType.MatchFound Then
        MsgBox "Please Select Vender Type", vbInformation
        cmbType.SetFocus
        Exit Sub
    End If

    If add Then
        AccNo = AddAccountInfo
        rsVend![AccNo] = AccNo
    ElseIf edit Then
        Call UpdateAccinfo
    End If
    
    
    If cmbProcess.MatchFound And cmbProcessID <> "" Then
        rsVend!ProcessID = cmbProcess.List(cmmbProcess.ListIndex, 1)
    Else
        rsVend!ProcessID = "0"
    End If

    rsVend.Update
    
    Dim i As Integer
    Dim iVendID As Integer
    iVendID = rsVend!VendID
    If bUrduSupported Then
        con.Execute "UPDATE Venders SET VenderNameUrdu=N'" & txtUrduFields(0).Text & "' WHERE VendID=" & rsVend!VendID
    End If
    
    Call SavePics(CDbl(iVendID))
    
    Me.add = True
    Call Form_QueryUnload(0, 0)
    Call ClearFields
    Call Form_Activate
    txtfields(1).SetFocus
    If cmbType.Enabled Then cmbType = ""
    Exit Sub

err:
    MsgBox "ERROR : # " & err.Number & vbNewLine & err.Description, vbInformation
 
End Sub

Private Sub SavePics(dVendID As Double)

    If (ImgPic.Picture = 0 Or Trim(PicFileName) = "") And (ImgSig.Picture = 0 Or Trim(ThumbFileName) = "") Then Exit Sub

    Dim Rec As New ADODB.Recordset
    Dim rsPic As New ADODB.Stream
    Dim rsThumb As New ADODB.Stream
    
    Rec.Open "SELECT VendID,VenderPic,VenderSig FROM Venders WHERE VendID=" & rsVend!VendID, con, adOpenStatic, adLockOptimistic
    If ImgPic.Picture <> 0 And Trim(PicFileName) <> "" Then
      'Save Picture
        rsPic.Open
        rsPic.Type = adTypeBinary
        rsPic.LoadFromFile PicFileName
        Rec.Fields(1) = rsPic.Read
        rsPic.Close
    End If
   
    'Save Thumb Expression
    If ImgSig.Picture <> 0 And Trim(ThumbFileName) <> "" Then
        rsThumb.Open
        rsThumb.Type = adTypeBinary
        rsThumb.LoadFromFile ThumbFileName
        Rec.Fields("VenderSig") = rsThumb.Read
        rsThumb.Close
    End If
    
    Rec.Update
    Rec.Close
    Set Rec = Nothing
    Set rsPic = Nothing
    Set rsThumb = Nothing
End Sub

Private Sub cmdSaveClose_Click()

    On Error GoTo err

    If Trim(txtfields(1)) = "" Then
        MsgBox "Invalid Vender Name", vbInformation
        Exit Sub
    ElseIf cmbType.MatchFound = False And add Then
        MsgBox "Select Vender Type From List.", vbInformation
        Exit Sub
    End If

    If add Then
        AccNo = AddAccountInfo
        rsVend![AccNo] = AccNo
    ElseIf edit Then
        Call UpdateAccinfo
    End If
    
    
    If cmbProcess.MatchFound And cmbProcess <> "" Then
        rsVend!ProcessID = cmbProcess.List(cmbProcess.ListIndex, 1)
    Else
        rsVend!ProcessID = "0"
    End If
    
    
    rsVend.Update
    
    Dim i As Integer
    Dim iVendID As Integer
    iVendID = rsVend!VendID
    If bUrduSupported Then
        con.Execute "UPDATE Venders SET VenderNameUrdu=N'" & txtUrduFields(0).Text & "' WHERE VendID=" & iVendID
    End If
    
    Call SavePics(CDbl(iVendID))
    
'    con.Execute "Delete From VenderMaterialCats Where VendID=" & iVendID
'    For i = 1 To LV.ListItems.Count
'        With LV.ListItems(i)
'            If .Checked Then
'                con.Execute "Insert Into VenderMaterialCats(VendID,MaterialCat) Values(" & _
'                 iVendID & ",'" & .Tag & "')"
'            End If
'        End With
'    Next
    Unload Me

    Exit Sub

err:
    MsgBox "ERROR : # " & err.Number & vbNewLine & err.Description, vbInformation
 
End Sub

Private Sub cmbType_Change()
    If cmbType.MatchFound Then Call cmbType_Click
End Sub

Private Sub cmbType_Click()

    Dim AccType As String, NewVal As String
    If cmbType = "" Or Not cmbType.Enabled Then Exit Sub
    AccType = getAccountNo(cmbType)
    If txtfields(1) <> "" Then
        txtfields(0) = GetNextAccno(txtfields(1), Left(Acno, 2), AccType)
    End If
    
    cmbProcess.Visible = CBool(cmbType = "Makers")
    lblLabels(18).Visible = cmbProcess.Visible
    'txtfields(0) = AccType & "-" & Format(Val(NewVal) + 1, "000")
End Sub

Private Sub DtOpenDate_Change()
    'Call setDateRange(DtOpenDate)
End Sub

Private Sub Form_Activate()

    If Acno = "" Then
        Unload Me
    End If

    On Error GoTo err
    
    With rsVend
        If .State <> 0 Then Exit Sub
        .Open "Select * From Venders Where VendID=" & Val(Me.Tag), con, adOpenDynamic, adLockOptimistic

        For i = 0 To txtfields.count - 1
            If i <> 1 Then Set txtfields(i).DataSource = rsVend
        Next
        'Set chkInactive.DataSource = rsVend
        Set chkImport.DataSource = rsVend
        'Set cmbProcess.DataSource = rsVend
        If add = False Then txtUrduFields(0) = !VenderNameUrdu & ""
        If add Then
            .AddNew
            
            If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                Dim strVendID As String
                strVendID = (GetSingleLongValue("COUNT(VendID)", "Venders") + 1)
                strVendID = Format(strVendID, "0000")
                strVendID = "SUP-" & strVendID
                txtfields(14) = strVendID
            End If
            
            If Not cmbType.Enabled Then
                'Get Next AccountNo Available For Customer
                'NewVal = GetMax("right(AccNo,3)", "Accounts", "Where SubAccOf='" & AcNo & "'")
                'txtfields(0) = AcNo & "-" & Format(Val(NewVal) + 1, "000")
            Else
                Call FindInCombo(cmbType, "Regular")
            End If
            txtUrduFields(0) = ""
        ElseIf View Then
            'Call Lok
        ElseIf edit Then
            cmbType.Enabled = False
            Call GetAccInfo(txtfields(0))
            'GetVenderTypes (Val(Me.Tag))
        End If
        
        cmbProcess.Visible = CBool(cmbType = "Makers")
        lblLabels(18).Visible = cmbProcess.Visible
        
        If Not IsNull(![VenderPic]) Then
            Set ImgPic.DataSource = rsVend
            ImgPic.DataField = "VenderPic"
            lblPic.Visible = False
        Else
            lblPic.Visible = True
        End If
        
        If Not IsNull(![VenderSig]) Then
            Set ImgSig.DataSource = rsVend
            ImgSig.DataField = "VenderSig"
            lblPicSig.Visible = False
        Else
            lblPicSig.Visible = True
        End If
        
    End With

Exit Sub
err:
    If err.Number = 545 Then Resume Next

End Sub

Private Sub GetVenderTypes(p_iVendID As Integer)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select MaterialCat From VenderMaterialCats Where VendID=" & p_iVendID, con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            LV.ListItems(!MaterialCat & "'").Checked = True
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()


    Me.KeyPreview = True
    If Not AccDefined Then
        MsgBox "Account "" " & ACCRECTITLE & " "" is Not Defined, Please Define The Account First", vbInformation
    End If
    
    Call AddToCombo(cmbType, "AccTitle", "Accounts", "Where Parent=1")
    'Call AddToCombo(cmbType, "Catagory", "MaterialCatagories", " Where CatID<>'02'")
    cmbType.Enabled = cmbType.ListCount > 0
    
    DtOpenDate = Date
        
    Call AddToCombo(cmbProcess, "Description,ProcessID", "Processes", "Order by SNo")
    Call AddToCombo(cmbMasterCat, "AccTitle", "Accounts", " WHERE Parent=1", , "AccNo")
    'Call AddVenderTypes
    
End Sub

Private Sub AddVenderTypes()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select CatID,Catagory From MaterialCatagories where subcatof='' Order By Catagory", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            LV.ListItems.add(, !CatID & "'", !Catagory & "").Tag = !CatID & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)

 Set rsVend = Nothing
End Sub


Private Sub ShowData()
    
End Sub

Private Sub GetAccInfo(AccNo As String)

    Dim rsAcc As New ADODB.Recordset

    With rsAcc
        .Open "Select AccNo,AccTitle,SubaccOf,OpenDate,OpenBal,Balance,Active from Accounts Where AccNo='" & AccNo & "'", con, adOpenForwardOnly
        AccNo = ![AccNo]
        DtOpenDate = ![OpenDate]
        txtfields(11) = Abs(![OpenBal])
        txtfields(1) = ![AccTitle]
        If Not ![Active] Then chkInactive = vbChecked Else chkInactive = vbUnchecked
   
        If ![SubAccOf] <> Acno Then cmbType = con.Execute("Select AccTitle From Accounts Where AccNo='" & ![SubAccOf] & "'").Fields(0)
   
        If Val(![OpenBal]) >= 0 Then
            cmbDrCr.ListIndex = 0
        Else
            cmbDrCr.ListIndex = 1
        End If

        'Remember Previous Opening Balance
        cmbDrCr.Tag = Val(![OpenBal] & "")
        
        Dim Balance As Double
        Balance = Val(![Balance] & "")
        lblBalance = IIf(Balance > 0, Abs(Balance) & " DR", IIf(Balance < 0, Abs(Balance) & " CR", sBalance))
        .Close
    End With
   
End Sub

Private Function AddAccountInfo() As String

    Call cmbDrCr_Click
    Dim Active As Boolean
    Active = Active = chkInactive <> vbChecked
 
    Dim ParentAcc As String
 
    ParentAcc = Left(txtfields(0), InStrRev(txtfields(0), "-") - 1)
 
    AddAccountInfo = NewAccPay(txtfields(1), DtOpenDate, Val(txtfields(11).Tag), Active, ParentAcc)
    
End Function
Private Function AccDefined() As Boolean

    'Get Accounts Receiveable AccountNo
    Acno = ACCPAYNO
    If Acno = "" Then
        AccDefined = False
    Else
        AccDefined = True
    End If
    
End Function

Private Sub UpdateAccinfo()
 
    Dim Diff As Double
    Dim Active As Boolean
    Call cmbDrCr_Click
  
    Diff = Val(txtfields(11).Tag) - Val(cmbDrCr.Tag)
    Active = chkInactive <> vbChecked
         
    con.Execute "Update Accounts Set accTitle='" & txtfields(1) & "',OpenBal=" & Val(txtfields(11).Tag) & ",Balance=balance + " & Diff & ",Active=" & Abs(Active) & " Where Accno='" & txtfields(0) & "'"
      
    Call adjustLedger(txtfields(0), DtOpenDate, 0, Diff, True)

    If Diff <> 0 Then Call UpdateAdvEntry(txtfields(0), Val(txtfields(11).Tag))
 
End Sub

Private Sub Form_Unload(cancel As Integer)
    Set frmVender = Nothing
End Sub

Private Sub ImgPic_DblClick()
    Call AddVendPic
End Sub

Private Sub ImgSig_DblClick()
    Call AddVendPic(True)
End Sub

Private Sub AddVendPic(Optional bSig As Boolean)
    On Error Resume Next
    Dim Picfile As String
    CD1.ShowSave
    If Not (err.Number = 32755) Then
        Picfile = CD1.FileName
        If bSig Then
            ImgSig.Picture = LoadPicture(Picfile)
            ThumbFileName = Picfile
            lblPicSig.Visible = False
        Else
            ImgPic.Picture = LoadPicture(Picfile)
            PicFileName = Picfile
            lblPic.Visible = False
        End If
        
    End If
End Sub


Private Sub lblPic_DblClick()
    Call AddVendPic
End Sub

Private Sub lblPicSig_DblClick()
    Call AddVendPic(True)
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)

    Select Case Index
        Case 11
            KeyAscii = OnlyNumber(KeyAscii, True)
        Case 2, 3, 4, 6, 7, 12, 13
         If Not Chr(KeyAscii) = "-" Then KeyAscii = OnlyNumber(KeyAscii)
    Case 1
'        KeyAscii = Asc(UCase(Chr(KeyAscii)))
   End Select
   
End Sub


Private Sub ClearFields()
   For i = 0 To txtfields.count - 1
      If i <> 10 Then txtfields(i) = ""
   Next
   For i = 0 To txtUrduFields.count - 1
      txtUrduFields(i).Text = ""
   Next
End Sub

Private Sub txtfields_LostFocus(Index As Integer)
    If Index = 1 And add And txtfields(1) <> "" Then
        'txtfields(0) = GetNextAccno(txtfields(1), Left(AcNo, 2), AcNo)
        Call cmbType_Click
    End If
End Sub
