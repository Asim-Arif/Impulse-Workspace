VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCustomer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer Info."
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fTitle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   8700
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Info."
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
         Left            =   3217
         TabIndex        =   40
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Info."
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
         Left            =   3232
         TabIndex        =   41
         Top             =   135
         Width           =   2295
      End
   End
   Begin VB.Frame FAll 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   60
      TabIndex        =   18
      Top             =   615
      Width           =   8625
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
         Left            =   5040
         TabIndex        =   32
         Top             =   990
         Width           =   3480
         Begin VB.TextBox txtfields 
            DataField       =   "Mobile"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   13
            Left            =   1755
            MaxLength       =   20
            TabIndex        =   11
            Top             =   1500
            Width           =   1620
         End
         Begin VB.TextBox txtfields 
            DataField       =   "CPhone"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   13
            Left            =   1755
            TabIndex        =   36
            Top             =   1290
            Width           =   450
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   10
            Left            =   75
            TabIndex        =   35
            Top             =   1290
            Width           =   555
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   5
            Left            =   60
            TabIndex        =   34
            Top             =   750
            Width           =   420
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   21
            Left            =   90
            TabIndex        =   33
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
         Height          =   2220
         Left            =   90
         TabIndex        =   25
         Top             =   990
         Width           =   4920
         Begin VB.TextBox txtfields 
            DataField       =   "Fax1"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   105
            MaxLength       =   100
            TabIndex        =   5
            Tag             =   "Contact Person Name"
            Top             =   1005
            Width           =   2370
         End
         Begin VB.TextBox txtfields 
            DataField       =   "Address"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   2535
            MaxLength       =   100
            TabIndex        =   6
            Top             =   1005
            Width           =   2295
         End
         Begin VB.TextBox txtfields 
            DataField       =   "Phone3"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   3270
            MaxLength       =   20
            TabIndex        =   4
            Top             =   465
            Width           =   1560
         End
         Begin VB.TextBox txtfields 
            DataField       =   "Phone2"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
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
            Caption         =   "Address"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   9
            Left            =   90
            TabIndex        =   31
            Top             =   1275
            Width           =   585
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax 2 :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   8
            Left            =   2565
            TabIndex        =   30
            Top             =   765
            Width           =   510
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone3"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   6
            Left            =   3285
            TabIndex        =   29
            Top             =   255
            Width           =   540
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone2"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   4
            Left            =   1710
            TabIndex        =   28
            Top             =   240
            Width           =   540
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone1"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   3
            Left            =   90
            TabIndex        =   27
            Top             =   225
            Width           =   540
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax 1 :"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   2
            Left            =   105
            TabIndex        =   26
            Top             =   765
            Width           =   510
         End
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
         Height          =   870
         Left            =   90
         TabIndex        =   19
         Top             =   3270
         Width           =   8445
         Begin VB.ComboBox cmbDrCr 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmCustomer.frx":0000
            Left            =   5670
            List            =   "frmCustomer.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   450
            Width           =   1110
         End
         Begin VB.TextBox txtfields 
            DataField       =   "AccNo"
            DataSource      =   "datPrimaryRS"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   135
            TabIndex        =   20
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
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   11
            Left            =   4275
            MaxLength       =   10
            TabIndex        =   14
            Top             =   450
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
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "MMMM dd/MM/yyyy"
            Format          =   144244739
            CurrentDate     =   37055
         End
         Begin VB.Label lblBalance 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00 "
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
            Height          =   315
            Left            =   6825
            TabIndex        =   43
            Top             =   450
            Width           =   1515
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Balance"
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
            Index           =   14
            Left            =   6795
            TabIndex        =   42
            Top             =   225
            Width           =   1380
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Acc # :"
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
            Index           =   7
            Left            =   135
            TabIndex        =   23
            Top             =   225
            Width           =   570
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Opening Bal."
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
            Index           =   12
            Left            =   4275
            TabIndex        =   22
            Top             =   225
            Width           =   1035
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Opening Date :"
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
            Left            =   1755
            TabIndex        =   21
            Top             =   225
            Width           =   1230
         End
      End
      Begin VB.CheckBox chkInactive 
         Caption         =   "Customer is inactive"
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
         Left            =   5055
         TabIndex        =   12
         Tag             =   "False"
         Top             =   2925
         Width           =   2205
      End
      Begin VB.TextBox txtfields 
         DataField       =   "Name"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   180
         MaxLength       =   255
         TabIndex        =   0
         Top             =   435
         Width           =   6225
      End
      Begin VB.TextBox txtfields 
         DataSource      =   "datPrimaryRS"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   180
         TabIndex        =   24
         Top             =   435
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         Height          =   630
         Left            =   3255
         TabIndex        =   45
         Top             =   4110
         Width           =   5280
         Begin MSForms.CommandButton cmdSave 
            Height          =   360
            Left            =   75
            TabIndex        =   16
            Top             =   195
            Width           =   1665
            Caption         =   "Save & New   "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmCustomer.frx":001D
            Accelerator     =   110
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
            TabIndex        =   46
            Top             =   195
            Width           =   1665
            Caption         =   "Cancel            "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmCustomer.frx":012F
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
            Left            =   1800
            TabIndex        =   17
            Top             =   195
            Width           =   1665
            Caption         =   "Save & Close  "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmCustomer.frx":0241
            Accelerator     =   83
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Type"
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
         Left            =   6480
         TabIndex        =   44
         Top             =   210
         Width           =   1290
      End
      Begin MSForms.ComboBox cmbType 
         Height          =   300
         Left            =   6465
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
         Caption         =   "Company Name "
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
         Left            =   210
         TabIndex        =   38
         Top             =   225
         Width           =   1365
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer ID"
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
         TabIndex        =   37
         Top             =   510
         Visible         =   0   'False
         Width           =   1050
      End
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCust As New ADODB.Recordset
Public Add As Boolean, Edit As Boolean, View
Dim AcNo As String


Private Sub chkInactive_Click()
   If chkInactive = vbChecked And Val(lblBalance) <> 0 Then
      MsgBox "This Customer Cannot Be Made Inactive," & vbNewLine & _
             "To Make A Customer Inactive Balance Must Be Zero", vbInformation
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

Private Sub cmbType_Change()
If cmbType.MatchFound Then Call CmbType_Click
End Sub

Private Sub CmbType_Click()
Dim AccType As String, NewVal As String
If cmbType = "" Or Not cmbType.Enabled Then Exit Sub
   AccType = getAccountNo(cmbType)
   NewVal = GetMax("right(AccNo,3)", "Accounts", "Where SubAccOf='" & AccType & "'")
   txtfields(0) = AccType & "-" & Format(Val(NewVal) + 1, "000")
End Sub

Private Sub cmdCancel_Click()
 Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo ERR
If Trim(txtfields(1)) = "" Then
    MsgBox "Invalid Customer Name", vbInformation
    Exit Sub
End If


If Add Then
   AccNo = AddAccountInfo
   rsCust![AccNo] = AccNo
ElseIf Edit Then
   Call UpdateAccinfo
End If
    


    rsCust.Update
    
Me.Add = True
Call Form_QueryUnload(0, 0)
Call ClearFields
Call Form_Activate
If cmbType.Enabled Then cmbType = ""
txtfields(1).SetFocus
Exit Sub

ERR:
 MsgBox "ERROR : # " & ERR.Number & vbNewLine & ERR.Description, vbInformation
 
End Sub



Private Sub cmdSaveClose_Click()
On Error GoTo ERR
If Trim(txtfields(1)) = "" Then
    MsgBox "Invalid Customer Name", vbInformation
    Exit Sub
End If

If Add Then
   AccNo = AddAccountInfo
   rsCust![AccNo] = AccNo
ElseIf Edit Then
   Call UpdateAccinfo
End If
    
    
    
    
    rsCust.Update
    
Unload Me

Exit Sub

ERR:
 MsgBox "ERROR : # " & ERR.Number & vbNewLine & ERR.Description, vbInformation
 
End Sub

Private Sub cmdSClose_Click()

End Sub

Private Sub DtOpenDate_Change()
    Call setDateRange(DtOpenDate)
End Sub

Private Sub Form_Activate()



If AcNo = "" Then
 Unload Me
End If





On Error GoTo ERR
With rsCust

If .State <> 0 Then Exit Sub


 .Open "Select * From Customers Where CustID=" & Val(Me.Tag), con, adOpenDynamic, adLockOptimistic

    For i = 0 To txtfields.Count - 1
     If i <> 1 Then Set txtfields(i).DataSource = rsCust
    Next
     'Set chkInactive.DataSource = rsCust





 If Add Then
 .AddNew
  If Not cmbType.Enabled Then
   'Get Next AccountNo Available For Customer
   NewVal = GetMax("right(AccNo,3)", "Accounts", "Where SubAccOf='" & AcNo & "'")
   txtfields(0) = AcNo & "-" & Format(Val(NewVal) + 1, "000")
  Else
      Call FindInCombo(cmbType, "Foreign")
  End If
 ElseIf View Then
  'Call Lok
 ElseIf Edit Then
  cmbType.Enabled = False
  Call GetAccInfo(txtfields(0))
 End If
End With

Exit Sub
ERR:
 If ERR.Number = 545 Then Resume Next Else MsgBox "Error :" & ERR.Number & vbNewLine & ERR.Description

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()

cmdSave.Enabled = UserHasAccess("NewCustomer")

Me.KeyPreview = True
 If Not AccDefined Then
  MsgBox "Account "" " & ACCRECTITLE & " "" is Not Defined, Please Define The Account First", vbInformation
 End If
Call AddToCombo(cmbType, "AccTitle", "Accounts", "Where SubAccof Like '" & ACCRECNO & "%' And Parent=1")
cmbType.Enabled = cmbType.ListCount > 0
DtOpenDate = Date
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 Set rsCust = Nothing
End Sub


Private Sub showdata()

End Sub



Private Sub GetAccInfo(AccNo As String)
Dim rsAcc As New ADODB.Recordset

With rsAcc
   .Open "Select AccNo,AccTitle,OpenDate,OpenBal,Balance,Active,subAccOf from Accounts Where AccNo='" & AccNo & "'", con, adOpenForwardOnly
   
   AccNo = ![AccNo]
   DtOpenDate = ![OpenDate]
   txtfields(11) = Abs(![OpenBal])
   txtfields(1) = ![AccTitle]
   
   If ![SubAccOf] <> AcNo Then cmbType = con.Execute("Select AccTitle From Accounts Where AccNo='" & ![SubAccOf] & "'").Fields(0)
   
   
   If Not ![Active] Then chkInactive = vbChecked Else chkInactive = vbUnchecked
      
   
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
 AddAccountInfo = NewAccRcv(txtfields(1), DtOpenDate, Val(txtfields(11).Tag), Active, Left(txtfields(0), Len(txtfields(0)) - 4))
End Function
Private Function AccDefined() As Boolean
'Get Accounts Receiveable AccountNo
AcNo = ACCRECNO
If AcNo = "" Then
   AccDefined = False
 Else
  AccDefined = True
 End If
End Function

Private Sub UpdateAccinfo()
    Call cmbDrCr_Click
 Dim Diff As Double
 Diff = Val(txtfields(11).Tag) - Val(cmbDrCr.Tag)
Dim Active As Boolean
 Active = chkInactive <> vbChecked

      con.Execute "Update Accounts Set accTitle='" & txtfields(1) & "',OpenBal=" & Val(txtfields(11).Tag) & ",Balance=balance + " & Diff & ",Active=" & Abs(Active) & " Where Accno='" & txtfields(0) & "'"
      Call adjustLedger(txtfields(0), DtOpenDate, 0, Diff, True)

If Diff <> 0 Then Call UpdateAdvEntry(txtfields(0), Val(txtfields(11).Tag))
 
End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
      Case 11
         KeyAscii = OnlyNumber(KeyAscii, True)
      Case 2, 3, 4, 6, 7, 12, 13
         If Not Chr(KeyAscii) = "-" Then KeyAscii = OnlyNumber(KeyAscii)
   End Select
End Sub

Private Sub ClearFields()
   For i = 0 To txtfields.Count - 1
      txtfields(i) = ""
   Next
   txtfields(1).SetFocus
End Sub
