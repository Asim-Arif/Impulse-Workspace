VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmSampleInvoice_CustomerFeedback 
   AutoRedraw      =   -1  'True
   ClientHeight    =   10560
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12390
   ControlBox      =   0   'False
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
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10560
   ScaleWidth      =   12390
   WindowState     =   2  'Maximized
   Begin VB.Frame FAll 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10620
      Left            =   90
      TabIndex        =   0
      Top             =   -75
      Width           =   12270
      Begin VB.TextBox txtInvoiceNo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1545
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   915
         Width           =   3285
      End
      Begin VB.CommandButton cmdClose 
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
         Height          =   525
         Left            =   10650
         TabIndex        =   30
         Top             =   9975
         Width           =   1485
      End
      Begin VB.Frame FFB 
         Caption         =   "Feedback :"
         Height          =   4320
         Left            =   210
         TabIndex        =   29
         Top             =   2775
         Visible         =   0   'False
         Width           =   11850
         Begin VB.TextBox txtRemarks 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2025
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   42
            Top             =   1410
            Width           =   11160
         End
         Begin VB.OptionButton Opts 
            Caption         =   "Accepted"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   1
            Left            =   2205
            TabIndex        =   41
            Top             =   990
            Width           =   1365
         End
         Begin VB.OptionButton Opts 
            Caption         =   "Rejected"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   0
            Left            =   255
            TabIndex        =   40
            Top             =   990
            Width           =   1365
         End
         Begin VB.CommandButton cmdBrowseBarcode 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   8130
            TabIndex        =   37
            Top             =   3705
            Width           =   405
         End
         Begin VB.TextBox txtFilePath 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   255
            TabIndex        =   38
            Top             =   3720
            Width           =   7845
         End
         Begin VB.TextBox txtItemCode 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   255
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   585
            Width           =   11160
         End
         Begin VB.CommandButton cmdUpdateFB 
            Caption         =   "&Update"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   8730
            TabIndex        =   34
            Top             =   3630
            Width           =   1305
         End
         Begin VB.CommandButton cmdCancelFB 
            Caption         =   "Cance&l"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   10095
            TabIndex        =   33
            Top             =   3630
            Width           =   1305
         End
         Begin MSComDlg.CommonDialog CD1 
            Left            =   10875
            Top             =   3240
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            Filter          =   "PDF Files Only|*.pdf"
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "PDF Path to Attach"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   270
            TabIndex        =   39
            Top             =   3495
            Width           =   2040
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Item"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   270
            TabIndex        =   35
            Top             =   300
            Width           =   465
         End
      End
      Begin VB.Frame Frame9 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   12255
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Sample Invoice Customer Feedback"
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
            Index           =   33
            Left            =   75
            TabIndex        =   17
            Top             =   135
            Width           =   12075
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Sample Invoice Customer Feedback"
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
            Height          =   465
            Index           =   34
            Left            =   90
            TabIndex        =   18
            Top             =   150
            Width           =   12060
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   8430
         Left            =   210
         TabIndex        =   28
         Top             =   1410
         Width           =   11910
         _ExtentX        =   21008
         _ExtentY        =   14870
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   15199215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S No."
            Object.Width           =   1014
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cust. Ref. No."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Order #"
            Object.Width           =   1720
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Size"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Color"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Article #"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Article Name"
            Object.Width           =   3941
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Item No."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Price"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "Qty"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "Amount"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   11
            Text            =   "DTRE #"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   12
            Text            =   "Batch #"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Delivery"
            Object.Width           =   2364
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Invoice No. : "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   210
         TabIndex        =   31
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Hidden"
      Height          =   9045
      Left            =   3600
      TabIndex        =   1
      Top             =   705
      Visible         =   0   'False
      Width           =   6285
      Begin VB.Label lblGlovesComDecs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Declarations For Gloves"
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
         Height          =   255
         Left            =   315
         MouseIcon       =   "frmSampleInvoice_CustomerFeedback.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   5670
         Width           =   5730
      End
      Begin VB.Label lblGarmentsComDecs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Declarations For Garments"
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
         Height          =   255
         Left            =   315
         MouseIcon       =   "frmSampleInvoice_CustomerFeedback.frx":0152
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   4860
         Width           =   5730
      End
      Begin VB.Label lblBallComDecs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Declarations For Footballs"
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
         Height          =   255
         Left            =   315
         MouseIcon       =   "frmSampleInvoice_CustomerFeedback.frx":02A4
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   4050
         Width           =   5730
      End
      Begin MSForms.TextBox txtComDeclarationGarments 
         Height          =   585
         Left            =   315
         TabIndex        =   12
         Tag             =   "DeclarationsGarments"
         Top             =   5100
         Width           =   5730
         VariousPropertyBits=   -1400879077
         BorderStyle     =   1
         Size            =   "10107;1032"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtComDeclarationBalls 
         Height          =   585
         Left            =   315
         TabIndex        =   11
         Top             =   4290
         Width           =   5730
         VariousPropertyBits=   -1400879077
         BorderStyle     =   1
         Size            =   "10107;1032"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtComDeclarationGloves 
         Height          =   585
         Left            =   315
         TabIndex        =   10
         Top             =   5910
         Width           =   5730
         VariousPropertyBits=   -1400879077
         BorderStyle     =   1
         Size            =   "10107;1032"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblGlovesDecs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Declarations For Gloves"
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
         Height          =   255
         Left            =   45
         MouseIcon       =   "frmSampleInvoice_CustomerFeedback.frx":03F6
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   1815
         Width           =   5805
      End
      Begin VB.Label lblGarmentsDecs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Declarations For Garments"
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
         Height          =   255
         Left            =   45
         MouseIcon       =   "frmSampleInvoice_CustomerFeedback.frx":0548
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1050
         Width           =   5805
      End
      Begin VB.Label lblBallDecs 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Declarations For Footballs"
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
         Height          =   255
         Left            =   45
         MouseIcon       =   "frmSampleInvoice_CustomerFeedback.frx":069A
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   285
         Width           =   5805
      End
      Begin MSForms.TextBox txtDeclarationGarments 
         Height          =   540
         Left            =   45
         TabIndex        =   6
         Tag             =   "DeclarationsGarments"
         Top             =   1290
         Width           =   5805
         VariousPropertyBits=   -1400879077
         BorderStyle     =   1
         Size            =   "10239;952"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtDeclarationBalls 
         Height          =   540
         Left            =   45
         TabIndex        =   5
         Top             =   525
         Width           =   5805
         VariousPropertyBits=   -1400879077
         BorderStyle     =   1
         Size            =   "10239;952"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtDeclarationGloves 
         Height          =   540
         Left            =   45
         TabIndex        =   4
         Top             =   2055
         Width           =   5805
         VariousPropertyBits=   -1400879077
         BorderStyle     =   1
         Size            =   "10239;952"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Manual Declaration :"
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
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmSampleInvoice_CustomerFeedback.frx":07EC
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   2595
         Width           =   1875
      End
      Begin MSForms.TextBox txtManDeclaration 
         Height          =   915
         Left            =   1980
         TabIndex        =   2
         Tag             =   "ManDeclaration"
         Top             =   2595
         Width           =   3945
         VariousPropertyBits=   -1400879077
         BorderStyle     =   1
         Size            =   "6959;1614"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame FraDispatch 
      BackColor       =   &H00E7EBEF&
      Caption         =   "Dispatch List"
      Height          =   4290
      Left            =   2025
      TabIndex        =   19
      Top             =   4230
      Visible         =   0   'False
      Width           =   8820
      Begin ComboList.Usercontrol1 cmbDispatchNo 
         Height          =   285
         Left            =   60
         TabIndex        =   20
         Top             =   480
         Width           =   2520
         _ExtentX        =   4445
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
         TextForeColor   =   16711680
         ListIndex       =   -1
         Appearance      =   0
      End
      Begin VB.TextBox txtDispQtyTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3660
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   3900
         Width           =   1230
      End
      Begin VB.CommandButton cmdOKDispatch 
         Caption         =   "OK"
         Height          =   360
         Left            =   7680
         TabIndex        =   26
         Top             =   3855
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancelDispatch 
         Caption         =   "Cancel"
         Height          =   360
         Left            =   6525
         TabIndex        =   25
         Top             =   3855
         Width           =   1050
      End
      Begin MSComctlLib.ListView LVDL 
         Height          =   3105
         Left            =   60
         TabIndex        =   24
         Top             =   750
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   5477
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15199215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S #"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Inner #"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Order #"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Qty"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.TextBox txtItemCartonQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   4005
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   480
         Width           =   1230
      End
      Begin MSForms.ComboBox cmbCartonNo 
         Height          =   285
         Left            =   2580
         TabIndex        =   22
         Top             =   480
         Width           =   1440
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2540;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Dispatch List No.                        Carton No.                    Qty"
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
         Height          =   255
         Index           =   36
         Left            =   60
         TabIndex        =   21
         Top             =   240
         Width           =   8670
      End
   End
End
Attribute VB_Name = "frmSampleInvoice_CustomerFeedback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Orders As New Collection
Dim strProforma As String
Dim bCustomInvoice As Boolean
Dim strDelList As String
Dim bCourierInvoice As Boolean
Dim iColNo As Integer, iColNoLV2 As Integer
Dim bAcceptsExtraQty As Boolean
Dim strDelListPL, strEditListPL As String
Dim bItemSelectionFromDispatch As Boolean
Dim strFileName  As String
Public Sub ShowInvoiceItems(strInvoiceNo As String)
         
    On Error GoTo err
    
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim dPrice As Double
    With rs
        strOrderBy = "ORDER BY EntryID"
        
        .Open "SELECT * FROM VCustomInvoiceItems WHERE CustomInvoice='" & strInvoiceNo & "' " & strOrderBy, con_ServerSide, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, ![RefID] & "'", LV.ListItems.count + 1)
            ITM.Tag = Val(!EntryID & "")
            
            ITM.ListSubItems.add(, , ![ITemID]).Tag = Val(![EntryID] & "") '![CatID] & ""
            ITM.ListSubItems.add , , ![OrderNo] '![CatID] & ""
            ITM.ListSubItems.add , , ![ItemSize] & "" '![ItemID] & ""
            ITM.ListSubItems.add , , "" '![Color] & ""
            ITM.ListSubItems.add(, , "").Tag = "" '![ArticleNo] & ""
            
            ITM.ListSubItems.add(, ![ItemName] & "'", ![ItemName] & "").Tag = ![ItemName] & ""
            dPrice = Val(!Price & "")
            
            ITM.ListSubItems.add , , !CompItemCode & "" '![Unit] & ""
            ITM.ListSubItems.add , , Format(dPrice, "#0.0000")
            ITM.ListSubItems.add , , Val(![Qty] & "")
            
            ITM.ListSubItems(8).Tag = Val(![Qty] & "") + Val(![InvQty] & "")
            
            ITM.SubItems(10) = Val(![Qty] & "") * dPrice
            
            ITM.SubItems(11) = !DTRENo & ""
            ITM.ListSubItems(11).Tag = !ItemForLabTest
            
            ITM.SubItems(12) = !BatchNo & ""
            
            .MoveNEXT
        Loop
        .Close
         
    End With
   
    Set rs = Nothing

    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub AWBDT_Change()
    If cmbCust.ID = "DECATHLON" Then
        DTETA = DateAdd("d", 20, AWBDT.Value)
    End If
End Sub

Private Sub chkCust_Click()
    
    Exit Sub
    'If cmbCatID.MatchFound Then
    '    Call FillItems
    '    For i = 1 To LV.ListItems.Count
    '        If chkCust Then
    '            LV.ListItems(i).SubItems(2) = LV.ListItems(i).Key
    '            LV.ListItems(i).ListSubItems(3) = LV.ListItems(i).ListSubItems(3).Key
    '        Else
    '            LV.ListItems(i).SubItems(2) = LV.ListItems(i).Tag
    '            LV.ListItems(i).ListSubItems(3) = LV.ListItems(i).ListSubItems(3).Tag
    '        End If
    '    Next i
    'End If
    
End Sub

Private Sub chkShowAllItems_Click()
    If chkShowAllItems.Value = vbChecked Then
        cmbItems.AddVals con_ServerSide, "'{'+ItemID+'} '+ItemName", "Items", "ItemID"
        chkShowAllItems.Visible = False
    End If
End Sub


Private Sub cmbCartonNo_Change()
    Call cmbCartonNo_Click
End Sub

Private Sub cmbCartonNo_Click()
    
    If cmbCartonNo.MatchFound = False Then Exit Sub
    txtItemCartonQty = GetSingleLongValue("SUM(InnerQty)", "VDispatchListDetail_Ex", " WHERE CartonNo=" & cmbCartonNo.Text & " AND DispatchListEntryID=" & cmbDispatchNo.ID & " AND ItemCode='" & cmbItems.ID & "'")
    
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT * FROM VDispatchListDetail_Ex WHERE DispatchListEntryID=" & cmbDispatchNo.ID & " AND CartonNo=" & cmbCartonNo.Text & " AND ItemCode='" & cmbItems.ID & "' ORDER BY InnerEntryID", con, adOpenForwardOnly, adLockReadOnly
        LVDL.ListItems.Clear
        Do Until .EOF
            Set ITM = LVDL.ListItems.add(, , LVDL.ListItems.count + 1)
            ITM.Checked = True
            ITM.Tag = !InnerEntryID
            ITM.ListSubItems.add , , !InnerNo & ""
            ITM.ListSubItems.add , , !Orders & ""
            ITM.ListSubItems.add , , Val(!InnerQty & "")
            
            
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub cmbComConsignee_Change()
    Call cmbComConsignee_Click
End Sub

Private Sub cmbComConsignee_Click()
    If cmbComConsignee.MatchFound = False Then Exit Sub
    txtComConsignee = cmbComConsignee.List(cmbComConsignee.ListIndex, 1)
End Sub


Private Sub cmbConsignee_Change()
    Call cmbConsignee_Click
End Sub

Private Sub cmbConsignee_Click()

'    If cmbConsignee = "Customer Address" Then
'        txtBankAddress = txtBankAddress.Tag
'        cmbThrough.Text = ""
'        cmbThrough.Enabled = False
'    ElseIf cmbConsignee = "Bank Address" Then
'        cmbThrough.Enabled = True
'        If cmbThrough.ListCount > 0 Then cmbThrough.ListIndex = 0
'    End If
    
End Sub

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    
    On Error GoTo err
    If cmbCust.MatchFound = False Then Exit Sub
    If cmbCountry.MatchFound = False Then Exit Sub
    
    txtCurrency = cmbCountry.List(cmbCountry.ListIndex, 1)
'    Call GetOrders
    Call GetBanks
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select FakeAddress From ForeignCustomers Where CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry & "'", con, adOpenForwardOnly, adLockReadOnly
        txtBankAddress.Tag = ![FakeAddress] & ""
        .Close
        .Open "SELECT AcceptsExtraQty,DefaultShipMethod,Address,PaymentTerms FROM ForeignCustomers WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        bAcceptsExtraQty = !AcceptsExtraQty
        cmbSMethod = !DefaultShipMethod & ""
        txtBankAddress = !Address & ""
        txtComNotify = !Address & ""
        cmbPayment = !PaymentTerms & ""
        .Close
    End With
    cmbDisCountry = cmbCountry
    Set rs = Nothing
   
    cmbItems.ClearVals
    cmbItems.AddVals con_ServerSide, "'{'+CompItemID+'} {'+ItemID+'} '+CompDesc", "VFCustomerCatalog", "CompItemID", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'"
    
    'Add Advance Payments....
    Call FillAdvancePayments
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub GetBanks()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select BankID,BankName,Address2 From FCustBanks Where CustCode+Country='" & cmbCust.Text & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbThrough.Clear
        Do Until .EOF
            cmbThrough.AddItem ![BankName] & ""
            cmbThrough.List(cmbThrough.ListCount - 1, 1) = ![BankID]
            cmbThrough.List(cmbThrough.ListCount - 1, 2) = ![Address2] & ""
            .MoveNEXT
        Loop
        
        .Close
    End With
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbCust_matched()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Country,Curr,Address From ForeignCustomers Where CustCode='" & cmbCust.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        cmbComConsignee.Clear
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            cmbComConsignee.AddItem ![Country] & ""
            cmbComConsignee.List(.AbsolutePosition - 1, 1) = ![Address] & ""
            cmbCountry.List(cmbCountry.ListCount - 1, 1) = ![Curr] & ""
            'cmbDisCountry.AddItem ![Country] & ""
            .MoveNEXT
        Loop
        .Close
    End With
   
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbCust_UnMatched()
    'cmbCountry.ClearVals
End Sub


Private Sub cmbDispatchNo_matched()

    cmbCartonNo.Clear
    Call AddToCombo(cmbCartonNo, "DISTINCT CartonNo", "VDispatchListDetail_Ex", " WHERE DispatchListEntryID=" & cmbDispatchNo.ID & " AND ItemCode='" & cmbItems.ID & "' ORDER BY CartonNo")
    
End Sub

Private Sub cmbitems_matched()

    On Error GoTo err
        
    If bItemSelectionFromDispatch = False Then
        Call cmdBreakLink_Click
    End If
    
    cmbProformas.ClearVals
    cmbProformas.AddVals con_ServerSide, "PInvoice+' {'+OrderNo+'}-{'+CAST(InternalRefNo AS VARCHAR(255))+'} {'+CAST(InvQty AS VARCHAR(50))+'}'", "VFProformas", "EntryID", " WHERE ItemID='" & cmbItems.ID & "' AND CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "' AND InvQty>0"
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbItemSearch_matched()
    Dim strOrderNo As String, strItemCode As String
'    strOrderNo = cmbItemSearch.List(cmbItemSearch.ListIndex, 2)
'    strItemCode = cmbItemSearch.List(cmbItemSearch.ListIndex, 1)
    Dim i As Integer, strSearch As String
    For i = 1 To LV2.ListItems.count
        If strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
            strSearch = LV2.ListItems(i).SubItems(1) & "|" & LV2.ListItems(i).SubItems(2) & "|" & LV2.ListItems(i).SubItems(7) & "|" & LV2.ListItems(i).ListSubItems(11).Tag
        Else
            strSearch = LV2.ListItems(i).SubItems(1) & "|" & LV2.ListItems(i).SubItems(2) & "|" & LV2.ListItems(i).SubItems(7)
        End If
        If strSearch = cmbItemSearch.ID Then
            LV2.ListItems(i).Selected = True
            LV2.ListItems(i).EnsureVisible
            LV2.SetFocus
            Call LV2_DblClick
            Exit For
        End If
    Next
End Sub

Private Sub cmbItemSearchInvoice_GotFocus()
    cmbItemSearchInvoice.ID = ""
End Sub

Private Sub cmbItemSearchInvoice_matched()

    Dim strOrderNo As String, strItemCode As String
    Dim i As Integer, strSearch As String
    For i = 1 To LV.ListItems.count
        If strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
            strSearch = LV.ListItems(i).SubItems(1) & "|" & LV.ListItems(i).SubItems(2) & "|" & LV.ListItems(i).ListSubItems(13).Tag
        Else
            strSearch = LV.ListItems(i).SubItems(1) & "|" & LV.ListItems(i).SubItems(2)
        End If
        If strSearch = cmbItemSearchInvoice.ID Then
            LV.ListItems(i).Selected = True
            LV.ListItems(i).EnsureVisible
            LV.SetFocus
            Exit For
        End If
    Next
    
End Sub

Private Sub cmbPayment_Change()
    Call cmbPayment_Click
End Sub

Private Sub cmbPayment_Click()

    If LCase(cmbPayment.Text) = "letter of credit" Then
        txtLC.Locked = False
    Else
        txtLC.Locked = True
        txtLC.Text = ""
    End If
    
End Sub

Private Sub cmbProformas_matched()

    Dim lQtyUsed As Long
    Dim i As Integer
    For i = 1 To LV2.ListItems.count
        If Val(LV2.ListItems(i).ListSubItems(1).Tag) = Val(cmbProformas.ID) Then
            lQtyUsed = lQtyUsed + Val(LV2.ListItems(i).SubItems(4))
        End If
    Next
       
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT Qty,InvQty FROM VFProformas WHERE EntryID=" & cmbProformas.ID, con, adOpenForwardOnly, adLockReadOnly
        txtOrderQty = Val(!Qty & "")
        txtBalance = Val(!InvQty & "")
        txtUsed = lQtyUsed
        txtAvailable = Val(txtBalance) - Val(txtUsed)
        .Close
    End With
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbShippingAddresses_Change()
    Call cmbShippingAddresses_Click
End Sub

Private Sub cmbShippingAddresses_Click()

    If cmbShippingAddresses.MatchFound = False Then
        txtShippingAddress = ""
    Else
        txtShippingAddress = GetSingleStringValue("ShippingAddress", "ForeignCustomersShippingAddresses", " WHERE EntryID=" & cmbShippingAddresses.List(cmbShippingAddresses.ListIndex, 1))
    End If
    
End Sub

Private Sub cmbSMethod_Change()
    Call cmbSMethod_Click
End Sub

Private Sub cmbSMethod_Click()
    If cmbSMethod.ListIndex = 0 Then
        'Label1(12).Caption = "AWB No."
    Else
        'Label1(12).Caption = "BL No."
    End If
End Sub

Private Sub cmbTradeTerms_Change()
    Call cmbTradeTerms_Click
End Sub

Private Sub cmbTradeTerms_Click()
    
    
    If cmbTradeTerms.ListIndex = 1 Or cmbTradeTerms.ListIndex = 2 Then
        cmbInsurAgent.Enabled = True
    Else
        cmbInsurAgent.Enabled = False
    End If
    
    If cmbTradeTerms.Text Like "*Air*" Then
        cmbSMethod.ListIndex = 0
    ElseIf cmbTradeTerms.Text Like "*Sea*" Then
        cmbSMethod.ListIndex = 1
    End If
    
    Exit Sub
    
    If cmbTradeTerms.MatchFound Then
    
        cmbCatID.Locked = False
        
        cmbDesc.Locked = False
        
         
        
        
    Else
        cmbCatID.Locked = True
        
        cmbDesc.Locked = True
    End If
    
End Sub


Private Sub cmdAddAll_Click()

    If LV2.ListItems.count = 0 Then Exit Sub
    Dim i As Integer
    For i = 1 To LV2.ListItems.count
        LV2.ListItems(i).Selected = True
        txtQty1 = Val(LV2.ListItems(i).SubItems(9))
        Call AddToInvoice
    Next
    
End Sub

Private Sub AddToInvoice()

    Call MoveToFirstLV
    
    LV2.SelectedItem.SubItems(9) = LV2.SelectedItem.SubItems(9) - Val(txtQty1)
    LV2.SelectedItem.SubItems(10) = Val(txtQty1) * LV2.SelectedItem.SubItems(8)
    
    txtQty1.Visible = False
    
    Call CalculateTotals
    LV2.SetFocus
    
End Sub


Private Sub cmdAddOtherCharges_Click()
    
    On Error GoTo err
    
    If Val(txtOtherChargesSNo) > LVOtherCharges.ListItems.count + 1 Then
        MsgBox "Invalid S/No.", vbCritical
        txtOtherChargesSNo.Text = LVOtherCharges.ListItems.count + 1
        txtOtherChargesSNo.SetFocus
        Exit Sub
    End If
    
    If txtOtherCharges.Text = "" Then
        MsgBox "Please Enter Other Charges Description.", vbCritical
        txtOtherCharges.SetFocus
    End If
    
    If IsNumeric(txtOtherChargeAmt.Text) = False Then
        MsgBox "Invalid Amount.", vbCritical
        txtOtherChargeAmt.SelStart = 0
        txtOtherChargeAmt.SelLength = Len(txtOtherChargeAmt.Text)
        txtOtherChargeAmt.SetFocus
    End If
    
    Dim ITM As ListItem
    Set ITM = LVOtherCharges.ListItems.add(Val(txtOtherChargesSNo), , Val(txtOtherChargesSNo))
    ITM.SubItems(1) = txtOtherCharges.Text & ""
    ITM.SubItems(2) = Val(txtOtherChargeAmt)
    
    Call PrepareForNext
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub PrepareForNext()

    Dim i As Integer
    For i = 1 To LVOtherCharges.ListItems.count
        LVOtherCharges.ListItems(i).Text = i
    Next
    txtOtherChargesSNo.Text = LVOtherCharges.ListItems.count + 1
    txtOtherCharges.Text = ""
    txtOtherChargeAmt.Text = ""
    txtOtherCharges.SetFocus
    
End Sub

Private Sub cmdBreakLink_Click()

    txtDispatchListNo.Text = "Not Linked with any Dispatch List"
    txtDispatchListNo.Tag = "0"
    txtDispatchListCartonNo = ""
    txtDispatchListQty = ""
    txtDispatchListQtyUsed = ""
    txtDispatchListQtyBalance = ""
    
End Sub

Private Sub cmdCallDispatch_Click()
'    If cmbItems.MatchFound = False Then
'        MsgBox "Please Select Item.", vbInformation
'        Exit Sub
'    Else
'        If cmbDispatchNo.MatchFound = False Then
'            cmbDispatchNo.ClearVals
'            cmbDispatchNo.AddVals con, "DISTINCT DispatchListNo", "VDispatchListDetail_Ex", "DispatchListEntryID", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "' AND ItemCode='" & cmbItems.ID & "'"
'        End If
'        FAll.Enabled = False
'        FraDispatch.Visible = True
'    End If
    Dim f As New frmGetSingleComboValue, strDispatchListNo As String
    Load f
    Call f.cmbCombo.AddVals(con, "DispatchListNo+'{'+CAST(EntryID AS VARCHAR(10))+'}'", "DispatchList", "EntryID", " WHERE CustCode='" & cmbCust.ID & "'")
    If f.getComboValue(strDispatchListNo, "Dispatch List", "Dispatch List") = False Then Exit Sub
    
    Dim lEntryID As Long, ITM As ListItem
    lEntryID = Val(strDispatchListNo)
     
    Dim rs As New ADODB.Recordset, rsCartons As New ADODB.Recordset, rsProformas As New ADODB.Recordset, lAvailableQty As Long, lQtyToUse As Long
    With rs
        .Open "SELECT * FROM VDispatchListToExport WHERE DispatchListNo=" & lEntryID & " AND DispatchQty<=InvQty", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            With rsCartons
                .Open "SELECT CartonNo,OrderNo,ItemCode,SUM(Qty) AS CartonQty FROM VDispatchListDetail_Ex_OrderWise " & _
                    "WHERE DispatchListEntryID=" & lEntryID & " AND OrderNo='" & rs!OrderNo & "' AND ItemCode='" & rs!ItemCode & "' GROUP BY CartonNo,OrderNo,ItemCode ORDER BY CartonNo", con, adOpenForwardOnly, adLockReadOnly
                
                Do Until .EOF
                    With rsProformas
                        .Open "SELECT EntryID,OrderNo,PInvoice,ItemID,ItemName,Qty,InvQty,Price,CustomPrice " & _
                         "FROM VFProformas WHERE OrderNo='" & rsCartons!OrderNo & "' AND ItemID='" & rsCartons!ItemCode & "' AND InvQty>0", con, adOpenForwardOnly, adLockReadOnly
                        lAvailableQty = Val(rsCartons!CartonQty & "")
                        
                        Do Until (.EOF Or (lAvailableQty = 0))
                                                       
                            Set ITM = LV2.ListItems.add(, , Val(rsCartons!CartonNo & ""))
                            
                            ITM.ListSubItems.add(, , !PInvoice & "").Tag = Val(!EntryID & "")
                            ITM.ListSubItems.add(, , !OrderNo & "").Tag = Val(!Price & "")
                            ITM.ListSubItems.add(, , !ITemID & " - " & !ItemName).Tag = Val(!CustomPrice & "")
                            
                            If Val(!InvQty & "") > lAvailableQty Then
                                lQtyToUse = lAvailableQty
                            Else
                                lQtyToUse = Val(!InvQty & "")
                            End If
                            lAvailableQty = lAvailableQty - lQtyToUse
                            
                            ITM.ListSubItems.add , , lQtyToUse
                            ITM.ListSubItems.add , , ""
                            
                            'Following 2 Lines are for Dispatch...
                            ITM.ListSubItems.add(, , Val(rsCartons!CartonNo & "")).Tag = lEntryID
                            ITM.ListSubItems.add(, , !ITemID & "").Tag = ""
                            .MoveNEXT
                        Loop
                        .Close
                    End With
                    .MoveNEXT
                Loop
                .Close
            End With
            .MoveNEXT
        Loop
        .Close
        
        .Open "SELECT * FROM VDispatchListToExport WHERE DispatchListNo=" & lEntryID & " AND DispatchQty>InvQty", con, adOpenForwardOnly, adLockReadOnly
        LVNotMatched.ListItems.Clear
        Do Until .EOF
            With rsCartons
                .Open "SELECT CartonNo,OrderNo,ItemCode,SUM(Qty) AS CartonQty FROM VDispatchListDetail_Ex_OrderWise " & _
                    "WHERE DispatchListEntryID=" & lEntryID & " AND OrderNo='" & rs!OrderNo & "' AND ItemCode='" & rs!ItemCode & "' GROUP BY CartonNo,OrderNo,ItemCode ORDER BY CartonNo", con, adOpenForwardOnly, adLockReadOnly
                    
                
                Do Until .EOF
                             
                    Set ITM = LVNotMatched.ListItems.add(, , Val(rsCartons!CartonNo & ""))
                            
                    ITM.ListSubItems.add(, , !OrderNo & "").Tag = ""
                    ITM.ListSubItems.add(, , !ItemCode & "").Tag = ""
                    ITM.ListSubItems.add , , Val(!CartonQty & "")
                    ITM.ListSubItems.add(, , Val(!CartonQty & "")).Tag = lEntryID
                    
                    
                    .MoveNEXT
                Loop
                    .Close
            End With
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Set rsCartons = Nothing
    Set rsProformas = Nothing
    
End Sub

Private Sub cmdBrowseBarcode_Click()

    On Error Resume Next
    Dim Picfile As String
    CD1.ShowSave
    
    If Not (err.Number = 32755) Then
        Picfile = CD1.FileName
        txtFilePath = Picfile
        txtFilePath.Tag = ""
        strFileName = CD1.FileTitle
    End If

End Sub

Private Sub cmdCancelDispatch_Click()

    FraDispatch.Visible = False
    FAll.Enabled = True
    
End Sub

Private Sub cmdCancelFB_Click()
    FFB.Visible = False
    LV.Enabled = True
    LV.SetFocus
End Sub


Private Sub cmdRemoveAll_Click()

    If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Dim i As Integer
    For i = LV2.ListItems.count To 1 Step -1
        With LV2.ListItems(i)
            If Val(.Tag) > 0 Then
                strDelListPL = strDelListPL & Val(.Tag) & ","
            End If
            LV2.ListItems.Remove .Index
        End With
    Next
    
End Sub

Private Sub cmdSwap_Click()
    Dim Temp As String
    Temp = txtComConsignee
    txtComConsignee = txtComNotify
    txtComNotify = Temp
End Sub



Private Sub cmdAddItem_Click()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset

    Dim strMsg As String
    strMsg = ""
    If Not (cmbCust.MatchFound And cmbCountry.MatchFound And cmbProformas.MatchFound) Then Exit Sub
    If cmbCust.Text = "" Or cmbCountry.Text = "" Or cmbProformas.MatchFound = False Then Exit Sub
    
    Dim cmd As New ADODB.Command
    'Now Check If Same Item With The Same Size Already Exists Or Not.
    
    For i = 1 To LV2.ListItems.count
        If LV2.ListItems(i).SubItems(1) = cmbProformas.Text Then
            MsgBox "Proforma Already Added.", vbInformation
            Exit Sub
        End If
    Next i
    
    
    strMsg = ""
    
    With rs
        .Open "SELECT * FROM VFProformas WHERE PInvoice='" & cmbProformas.ID & "' And InvQty>0 ORDER BY CompItemCode", con, adOpenForwardOnly, adLockReadOnly
        
        If LV2.ListItems.count > 0 Then
            
            If cmbTrade <> ![TradeTerms] & "" Then strMsg = ",Trade Terms"
            If cmbPartial <> ![PartialShipment] & "" Then strMsg = strMsg & ",Partial Shipment"
            If cmbTrans <> ![TransShipment] & "" Then strMsg = strMsg & ",Trans Shipment"
            If cmbPayment <> ![PaymentTerms] & "" Then strMsg = strMsg & ",Payment Terms"
            If txtPackaging <> ![Packaging] & "" Then strMsg = strMsg & ",Packaging"
            If DeliveryDT <> ![DeliveryDT] & "" Then strMsg = strMsg & ",Delivery Date"
            
        End If
        
        If strMsg <> "" Then
            strMsg = "Following Fields Are Different From The Previous Order." & vbNewLine & Right$(strMsg, Len(strMsg) - 1) & "." & vbNewLine & "Do You Want Add This Order?"
            If MsgBox(strMsg, vbQuestion + vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
        
        cmbTradeTerms = ![TradeTerms] & ""
        cmbPartial = ![PartialShipment] & ""
        cmbTrans = ![TransShipment] & ""
        cmbPayment = ![PaymentTerms] & ""
        txtPackaging = ![Packaging] & ""
        'cmbShipAgent.ID = !ShippingAgent
        cmbSMethod = !SMethod & ""
        txtDeliveryDT = Format(![DeliveryDT], "dd-MM-yyyy")
        txtInvoiceof = !InvoiceOf & ""
        cmbShipAgent.ID = Val(!ShippingAgent & "")
'        For i = 0 To cmbBank.ListCount - 1
'            If Val(cmbBank.List(i, 1)) = Val(!OurBankID & "") Then
'                cmbBank.ListIndex = i
'                Exit For
'            End If
'        Next
        cmbBank.Text = !Bank & ""
        If OrderDT.Tag = "" Then
            'OrderDT = ![dt] & ""
            OrderDT.Tag = "One Added"
        Else
            If DateDiff("d", OrderDT, ![DT]) > 0 Then
                'OrderDT = ![dt] & ""
            End If
        End If
        
        Dim ITM As ListItem
        Dim iSNo As Integer
        
        LV2.ListItems.Clear
        
        Do Until .EOF
            
            iSNo = LV2.ListItems.count + 1
            
        
            'Set ITM = lv2.ListItems.Add(, ![OrderNo] & ![ItemCode] & "'", iSno)
            Set ITM = LV2.ListItems.add(, ![EntryID] & "'", iSNo)
            ITM.ListSubItems.add , , !ITemID & "" '![PInvoice] & ""
            ITM.ListSubItems.add(, , ![OrderNo] & "").Tag = 0
            ITM.ListSubItems.add , , ![ItemSize] & "" '![ItemID] & ""
            ITM.ListSubItems.add , , ""
            ITM.ListSubItems.add , , ""
            ITM.ListSubItems.add , , ![ItemName] & ""
            ITM.ListSubItems.add , , !CompItemCode & "" '![Unit] & ""
            ITM.ListSubItems.add , , Val(![Price] & "")
            ITM.ListSubItems(7).Tag = Val(![Price] & "") 'Val(![CustomPrice] & "")
            ITM.ListSubItems.add , , Val(![InvQty] & "")
            ITM.ListSubItems(8).Tag = Val(![InvQty] & "")
            ITM.ListSubItems.add , , Val(![Unit] & "")
            ITM.SubItems(10) = Val(![Price] & "") * Val(![InvQty] & "")
            ITM.ListSubItems(9).Tag = Val(![RebateRate] & "")
            ITM.ListSubItems(5).Tag = 0 'Val(![PackNo] & "")
            
            If strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
                ITM.ListSubItems.add(, , Format(!DeliveryDTItem, "dd-MMM-yyyy")).Tag = Format(!DeliveryDTItem, "dd/MM/yy")
            End If
            'Set itm = LV2.ListItems.Add
            'itm.SubItems(3) = "Batch # " & txtBatchNo & "   Lot # " & txtBatchNo & "/" & ![ItemCode] & ""
            
'            cmbItemSearch.AddItem !ITemID & " (" & !OrderNo & ")"
'            cmbItemSearch.List(iSNo - 1, 1) = !ITemID & ""
'            cmbItemSearch.List(iSNo - 1, 2) = !OrderNo & ""
            
            .MoveNEXT
            
        Loop
        .Close
        
        cmbItemSearch.ClearVals
        
        If strCompany = "Dr-Frgz" Or strCompany = "Towne" Then
            cmbItemSearch.AddVals con, "ItemID + ' {' + OrderNo + '}' + ' (' + CompItemCode + ') {' + CONVERT(varchar(50),DeliveryDTItem,3) + '}'", "VFProformas", "ItemID+'|'+OrderNo+'|'+CompItemCode+'|'+CONVERT(varchar(50),DeliveryDTItem,3)", "WHERE PInvoice='" & cmbProformas.ID & "' And InvQty>0"
        Else
            cmbItemSearch.AddVals con, "ItemID + ' {' + OrderNo + '}' + ' (' + CompItemCode + ')'", "VFProformas", "ItemID+'|'+OrderNo+'|'+CompItemCode", "WHERE PInvoice='" & cmbProformas.ID & "' And InvQty>0"
        End If
        
        .Open "SELECT * FROM FPInvoiceOtherCharges WHERE PInvoice='" & cmbProformas.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            If MsgBox("This Proforma has Other Charges, Do you want to Add ?", vbQuestion + vbYesNo) = vbYes Then
                Do Until .EOF
                    iSNo = LVOtherCharges.ListItems.count + 1
                    Set ITM = LVOtherCharges.ListItems.add(, , iSNo)
                    ITM.SubItems(1) = !ChargesDetail & ""
                    ITM.SubItems(2) = Val(!ChargesAmt & "")
                    Call PrepareForNext
                    .MoveNEXT
                Loop
            End If
        End If
        .Close
    End With
    
    'Dim iUp As Integer
    'iUp = UBound(Orders)
    'Orders(iUp) = cmbOrders
    'ReDim Preserve Orders(iUp + 1)
    
    Orders.add cmbProformas.ID, cmbProformas.ID
     
     
    Set rs = Nothing
    
    Call CalculateTotals
    cmbItemSearch.SetFocus
    
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    
    Exit Sub
    'LV.ListItems.Remove LV.SelectedItem.Index
    'Dim strOrder As String
    'strOrder = LV.SelectedItem.SubItems(1)
    If LV.ListItems.count = 0 Then Exit Sub
    If Val(LV.SelectedItem.ListSubItems(1).Tag) > 0 Then
        strDelList = strDelList & Val(LV.SelectedItem.ListSubItems(1).Tag) & ","
    End If
    
    
    
    LV.ListItems.Remove LV.SelectedItem.Index
    'For i = 1 To LV.ListItems.Count
    '    If i > LV.ListItems.Count Then Exit For
    '    If LV.ListItems(i).SubItems(1) = strOrder Then
    '        LV.ListItems.Remove i
    '    End If
    'Next i
    
    'Orders.Remove strOrder
    cmdDelete.Visible = False
    
    Call CalculateTotals
    
End Sub

Private Sub cmdSClear_Click()
    If Saved Then
        Unload Me
        Load frmNewCustomInvoice
        frmNewCustomInvoice.AddNewInvoice True, bCourierInvoice
    End If
End Sub

Private Sub cmdSClose_Click()
    If Saved Then
        Unload Me
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdUpdateFB_Click()
    
    If Opts(0).Value = False And Opts(1).Value = False Then
        MsgBox "Please Select Accepted or Rejected.", vbInformation
        Exit Sub
    End If
    Dim lRecAffect As Long
    
    con.Execute "UPDATE CustomInvoiceItems_CustomerFeedback SET Accepted=" & Val(Opts(0).Tag) & ",Remarks='" & txtRemarks & "' WHERE CII_RefID=" & Val(txtRemarks.Tag), lRecAffect
    
    If lRecAffect = 0 Then
        con.Execute "INSERT INTO CustomInvoiceItems_CustomerFeedback(CII_RefID,Accepted,Remarks,UserName,MachineName) VALUES(" & Val(txtRemarks.Tag) & "," & Val(Opts(0).Tag) & ",'" & txtRemarks & "','" & _
         CurrentUserName & "','" & strComputerName & "')"
         
   
    End If
           
    Dim lmyEntryID As Long
    lmyEntryID = GetSingleLongValue("EntryID", "CustomInvoiceItems_CustomerFeedback", " WHERE CII_RefID=" & Val(txtRemarks.Tag))
    Call AttachPDF(lmyEntryID)
    
    FFB.Visible = False
    LV.Enabled = True
    LV.SetFocus
    
End Sub

Private Sub AttachPDF(p_lEntryID As Long)
    
    On Error Resume Next
    Dim strPDFPath As String
    strPDFPath = txtFilePath
    If strPDFPath <> "" And txtFilePath.Tag = "" Then
        Dim Rec As New ADODB.Recordset
        Dim rsPic As New ADODB.Stream
        Rec.Open "SELECT EntryID,PDF,PDFFileName FROM CustomInvoiceItems_CustomerFeedback WHERE EntryID=" & p_lEntryID, con, adOpenStatic, adLockOptimistic
        If Rec.RecordCount > 0 Then
            rsPic.Open
            rsPic.Type = adTypeBinary
            rsPic.LoadFromFile strPDFPath
            Rec.Fields(1) = rsPic.Read
            Rec.Fields(2) = strFileName
            rsPic.Close
            Rec.Update
        End If
        
        Rec.Close
        
        Set Rec = Nothing
        Set rsPic = Nothing
        
    End If

End Sub



Private Sub Form_Resize()

    On Error Resume Next
    
    FAll.Left = (Me.ScaleWidth - FAll.Width) / 2
    FAll.Top = (Me.ScaleHeight - FAll.Height) / 2
    
    FraDispatch.Left = (Me.ScaleWidth - FraDispatch.Width) / 2
    FraDispatch.Top = (Me.ScaleHeight - FraDispatch.Height) / 2
    
End Sub

Private Sub lblBallComDecs_Click()
    If txtComDeclarationBalls = "" Then
        txtComDeclarationBalls = con.Execute("Select Declaration From ItemCatagories Where CatID='01'").Fields(0).Value & ""
    End If
End Sub

Private Sub lblBallDecs_Click()
    If txtDeclarationBalls = "" Then
        txtDeclarationBalls = con.Execute("Select Declaration From ItemCatagories Where CatID='01'").Fields(0).Value & ""
    End If
End Sub

Private Sub lblGarmentsComDecs_Click()
    If txtComDeclarationGarments = "" Then
       txtComDeclarationGarments = con.Execute("Select Declaration From ItemCatagories Where CatID='02'").Fields(0).Value & ""
    End If
End Sub

Private Sub lblGarmentsDecs_Click()
    If txtDeclarationGarments = "" Then
       txtDeclarationGarments = con.Execute("Select Declaration From ItemCatagories Where CatID='02'").Fields(0).Value & ""
    End If
End Sub

Private Sub lblGlovesComDecs_Click()
    If txtComDeclarationGloves = "" Then
       txtComDeclarationGloves = con.Execute("Select Declaration From ItemCatagories Where CatID='03'").Fields(0).Value & ""
    End If
End Sub

Private Sub lblGlovesDecs_Click()
    If txtDeclarationGloves = "" Then
       txtDeclarationGloves = con.Execute("Select Declaration From ItemCatagories Where CatID='03'").Fields(0).Value & ""
    End If
End Sub

Private Sub lblSwapAddresses_Click()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    
    With rs
        If lblSwapAddresses.Tag = "" Then
            .Open "Select Address,FakeAddress From ForeignCustomers Where CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
            lblSwapAddresses.Tag = "0"
        Else
            .Open "Select OtherAddress,OtherFakeAddress From ForeignCustomers Where CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
            lblSwapAddresses.Tag = ""
        End If
        txtComConsignee.Text = .Fields(0) & ""
        txtComNotify.Text = .Fields(1) & ""
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    LV.Enabled = False

    txtItemCode = LV.SelectedItem.SubItems(1)
    Call ShowFeedback(Val(LV.SelectedItem.Tag))
     
    FFB.Visible = True
    
End Sub

Private Sub ShowFeedback(lCII_RefID As Long)

    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM CustomInvoiceItems_CustomerFeedback WHERE CII_RefID=" & lCII_RefID, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
        
            Opts(Val(!Accepted)).Value = True
            txtRemarks = !Remarks & ""
            
            
            txtFilePath = !PDFFileName & ""
            txtFilePath.Tag = !PDFFileName & ""
        Else
        
            txtFilePath = ""
            txtFilePath.Tag = ""
            Opts(0).Value = False
            Opts(1).Value = False
            txtRemarks = ""
            
        End If
        .Close
    End With
    Set rs = Nothing
    
    txtRemarks.Tag = lCII_RefID 'Val(!EntryID & "")
    
End Sub

Private Sub CalculateTotals()


    Dim dTotal As Double
    Dim DPcsTotal As Long
    With LV
        For i = 1 To .ListItems.count
            'dQty = dQty + Val(.TextMatrix(i, 4))
            dTotal = dTotal + Val(.ListItems(i).SubItems(10) & "")
            DPcsTotal = DPcsTotal + Val(.ListItems(i).SubItems(9) & "")
        Next i
    End With
    'txttQty = dQty
    lblTotal = dTotal
    lblTotWords = ConvertInWords(CCur(dTotal), "")
    lblTotPcs = DPcsTotal
    
End Sub

Private Function Saved() As Boolean
    
    On Error GoTo err
    
    Dim iLoadID As Integer
    Dim iDisID As Integer
    Dim iBankID1 As Integer
    Saved = False
    If txtPInvoice = "" Then
        MsgBox "Please Enter Proforma Invoice No.", vbInformation
        txtPInvoice.SetFocus
        Exit Function
    End If
    
    If LV2.ListItems.count = 0 Then
        MsgBox "No Items Added!!!", vbInformation
        Exit Function
    End If
    
    If LCase(cmbPayment) = "letter of credit" Then
        If txtLC = "" Then
            MsgBox "You Must Enter LC # If Payment Terms Is 'Letter Of Credit.", vbInformation
            Exit Function
        End If
    End If
    
    If cmbBankAccNo.ListIndex = -1 Then
        'iBankID = 4
        iBankID1 = -1
    Else
        iBankID1 = cmbBankAccNo.List(cmbBankAccNo.ListIndex, 1)
    End If
    
    If cmbLoadPort.ListIndex = -1 Then
        'iLoadID = 4
        iLoadID = -1
    Else
        iLoadID = cmbLoadPort.List(cmbLoadPort.ListIndex, 1)
    End If
    
    If cmbDisPort.ListIndex = -1 Then
        'idisid=4
        iDisID = -1
    Else
        iDisID = cmbDisPort.List(cmbDisPort.ListIndex, 1)
    End If
    
    Dim strSRONo  As String
    If strProforma = "" Then
        'Get Latest SRONo...
        strSRONo = con.Execute("SELECT SRONo FROM SROs WHERE EntryID=(SELECT MAX(EntryID) FROM SROs)").Fields(0).Value & ""
    End If
    
    Call StartTrans(con)
    '----------------------------------------------------------------
'    If strProforma <> "" Then
'
'        con.Execute "UPDATE FProformaOrders SET InvQty=InvQty+IsNull((Select  " & _
'         "Qty From CustomInvoiceItems  Where CustomInvoice='" & strProforma & _
'         "' And RefID=FProformaOrders.EntryID),0) WHERE EntryID In(Select RefID" & _
'         " From CustomInvoiceItems Where CustomInvoice='" & strProforma & "')"
'
'    End If
        
    If strProforma = "" Then
        Dim ComDecBalls As String, ComDecGarments As String, ComDecGloves As String
        Dim ComNotify As String, ComConsignee As String
        Dim ComPayTerms As String
        
        Dim rs As New ADODB.Recordset
        With rs
            .Open "Select Address1 From FCustBanks Where CustCode='" & cmbCust.ID & "' And Country='" & cmbCountry.Text & "' AND DefaultBank=1", con, adOpenForwardOnly, adLockReadOnly
            If .RecordCount > 0 Then
                ComNotify = .Fields(0) & ""
            End If
            .Close
            .Open "Select Address From ForeignCustomers Where CustCode='" & cmbCust.ID & "' And Country='" & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
            If .RecordCount > 0 Then
                ComConsignee = .Fields(0) & ""
            End If
        End With
        
        ComDecBalls = txtDeclarationBalls
        ComDecGarments = txtDeclarationGarments
        ComDecGloves = txtDeclarationGloves
        ComPayTerms = cmbPayment
    Else
        ComNotify = txtComNotify
        ComConsignee = txtComConsignee
        
        ComDecBalls = txtComDeclarationBalls
        ComDecGarments = txtComDeclarationGarments
        ComDecGloves = txtComDeclarationGloves
        ComPayTerms = txtComPayTerms
        
    End If
    
    Dim strDTGatePass As String
    If IsNull(DTGatePass) Then
        strDTGatePass = "NULL"
    Else
        strDTGatePass = "'" & DTGatePass & "'"
    End If
    If strProforma = "" Then

        con.Execute "INSERT INTO CustomInvoice (CustomInvoice,DT,CustCode,Country,ShippingMode," & _
         "LoadPortID,DisPortID,TradeTerms,PaymentTerms,LCNo,PartialShipment," & _
         "TransShipment,ShippingAgentID,InsuranceAgentID,OurBankID," & _
         "AWBNo,AWBNoDT,FormE,FormEDT,Vessel,GrossWeight,DeclarationsBalls,DeclarationsGarments," & _
         "DeclarationsGloves,Consignee,ComDeclarationsBalls,ComDeclarationsGarments," & _
         "ComDeclarationsGloves,ComNotify,ComConsignee,ComPayTerms,ComAgainst,ComDrawnUnder," & _
         "ComSpecial,ManDischarge,ManDeclaration,Warehouse,Gate,ContainerNo,DTETA,SRONo,CustCurrency,CompanyRefID,SealNo,InvoiceOf,LCDate,GatePassNo,GatePassDT,DTREDescription,NoSRO,CIFTradeTerms,SampleInvoice,FreeSamples,InvoiceDescription) VALUES('" & txtPInvoice & _
         "','" & OrderDT.Value & "','" & cmbCust.Text & "','" & cmbCountry.Text & "','" & _
         cmbSMethod.Text & "'," & iLoadID & "," & iDisID & ",'" & cmbTradeTerms.Text & _
         "','" & cmbPayment & "','" & txtLC & "','" & cmbPartial & "','" & _
         cmbTrans & "'," & Val(cmbShipAgent.ID) & "," & Val(cmbInsurAgent.ID) & "," & _
         iBankID1 & ",'" & txtAWBNo & "','" & AWBDT & "','" & txtFormE & "','" & FormEDT & _
         "','" & txtVessel & "'," & Val(txtGWeight) & ",'" & _
         txtDeclarationBalls & "','" & txtDeclarationGarments & "','" & txtDeclarationGloves & _
         "','" & txtBankAddress & "','" & ComDecBalls & "','" & _
         ComDecGarments & "','" & ComDecGloves & "','" & ComNotify & _
         "','" & ComConsignee & "','" & ComPayTerms & "','" & txtShippingAddress & "','" & _
         txtComDrawn & "','" & txtComSpecial & "','" & txtManDischarge & "','" & txtManDeclaration & "','" & _
         txtWarehouse & "','" & txtGate & "','" & txtContainerNo & "','" & DTETA & "','" & strSRONo & "','" & _
         txtCurrency & "'," & cmbCompany.ID & ",'" & txtSealNo & "','" & txtInvoiceof & "','" & DTLC & "','" & _
         txtGatePassNo & "'," & strDTGatePass & ",'" & txtDTREDesc & "'," & chkNoSRO.Value & "," & chkCIFTerms.Value & "," & chkSampleInvoice.Value & "," & chkFreeSamples.Value & ",'" & txtInvoiceDescription & "')"
    Else
        con.Execute "UPDATE CustomInvoice SET DT='" & OrderDT & "',ShippingMode='" & cmbSMethod & _
         "',LoadPortID=" & iLoadID & ",DisPortID=" & iDisID & ",TradeTerms='" & cmbTradeTerms.Text & _
         "',PaymentTerms='" & cmbPayment & "',LCNo='" & txtLC & "',PartialShipment='" & _
         cmbPartial & "',TransShipment='" & cmbTrans & "',ShippingAgentID=" & Val(cmbShipAgent.ID) & _
         ",InsuranceAgentID=" & Val(cmbInsurAgent.ID) & ",OurBankID=" & iBankID1 & ",AWBNo='" & _
         txtAWBNo & "',AWBNoDT='" & AWBDT & "',FormE='" & txtFormE & "',FormEDT='" & FormEDT & _
         "',Vessel='" & txtVessel & "',GrossWeight=" & Val(txtGWeight) & ",DeclarationsBalls='" & _
         txtDeclarationBalls & "',DeclarationsGarments='" & txtDeclarationGarments & _
         "',DeclarationsGloves='" & txtDeclarationGloves & "',Consignee='" & txtBankAddress & _
         "',ComDeclarationsBalls='" & ComDecBalls & "',ComDeclarationsGarments='" & _
         ComDecGarments & "',ComDeclarationsGloves='" & ComDecGloves & "',ComNotify='" & _
         ComNotify & "',ComConsignee='" & ComConsignee & "',ComPayTerms='" & ComPayTerms & _
         "',ComAgainst='" & txtShippingAddress & "',ComDrawnUnder='" & txtComDrawn & "',ComSpecial='" & _
         txtComSpecial & "',ManDischarge='" & txtManDischarge & "',ManDeclaration='" & _
         txtManDeclaration & "',Warehouse='" & txtWarehouse & "',Gate='" & txtGate & "',ContainerNo='" & _
         txtContainerNo & "',DTETA='" & DTETA & "',SealNo='" & txtSealNo & "',Invoiceof='" & txtInvoiceof & _
         "',LCDate='" & DTLC & "',GatePassNo='" & txtGatePassNo & "',GatePassDT=" & strDTGatePass & ",DTREDescription='" & txtDTREDesc & "',NoSRO=" & chkNoSRO.Value & ",CIFTradeTerms=" & chkCIFTerms.Value & ",SampleInvoice=" & chkSampleInvoice.Value & ",FreeSamples=" & chkFreeSamples.Value & ",InvoiceDescription='" & txtInvoiceDescription & "' WHERE CustomInvoice='" & strProforma & "'"
         
'        If strDelList <> "" Then
'            strDelList = Left(strDelList, Len(strDelList) - 1)
'             con.Execute "DELETE FROM CustomInvoiceItems WHERE EntryID IN(" & strDelList & ")"
'        End If
        
    End If
   
    Dim lEntryIDs() As Double
    Dim i As Integer, iCount As Integer, j As Integer, bFound As Boolean
    
    ReDim lEntryIDs(LV2.ListItems.count, 4)
    iCount = -1
    bFound = False
    Dim iIndex As Integer, lQty As Long, dPrice As Double, dCustomPrice As Double
    
    For i = 1 To LV2.ListItems.count
        With LV2.ListItems(i)
            If Val(.Tag) = 0 Then
                bFound = False
                For j = 0 To iCount
                    If lEntryIDs(j, 0) = Val(.ListSubItems(1).Tag) Then
                        bFound = True
                        Exit For
                    End If
                Next
                If bFound Then
                    iIndex = j
                    lQty = lEntryIDs(ilindex, 1) + Val(.SubItems(4))
                    dPrice = Val(.ListSubItems(2).Tag)
                    dCustomPrice = Val(.ListSubItems(3).Tag)
                Else
                    iIndex = iCount + 1
                    lQty = Val(.SubItems(4))
                    dPrice = Val(.ListSubItems(2).Tag)
                    dCustomPrice = Val(.ListSubItems(3).Tag)
                    iCount = iCount + 1
                End If
                If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                    dCustomPrice = dPrice
                End If
                lEntryIDs(iIndex, 0) = Val(.ListSubItems(1).Tag)
                lEntryIDs(iIndex, 1) = lQty
                lEntryIDs(iIndex, 2) = dPrice
                lEntryIDs(iIndex, 3) = dCustomPrice
            End If
        End With
    Next
    
     
    If strProforma <> "" Then
        Dim lQtyDiff As Long
        For i = 1 To LV2.ListItems.count
            With LV2.ListItems(i)
                If Val(.Tag) > 0 Then
                    If Val(.SubItems(4)) <> Val(.ListSubItems(4).Tag) Then
                        lQtyDiff = Val(.SubItems(4)) - Val(.ListSubItems(4).Tag)
                        con.Execute "UPDATE CustomInvoiceItems SET CustomInvoiceItems.Qty=CustomInvoiceItems.Qty+" & _
                         lQtyDiff & " FROM CustomInvoiceitems INNER JOIN CustomPList ON CustomInvoiceItems.EntryID" & _
                         "=CustomPlist.OrderItemID WHERE CustomPList.ID=" & Val(.Tag)
                         
                        con.Execute "UPDATE CustomPList SET Qty=" & Val(.SubItems(4)) & " WHERE ID=" & Val(.Tag)
                    End If
                End If
            End With
        Next
    End If
    
    If strDelListPL <> "" Then
    
        strDelListPL = Left(strDelListPL, Len(strDelListPL) - 1)
        
        con.Execute "UPDATE CustomInvoiceItems SET CustomInvoiceItems.Qty=CustomInvoiceItems.Qty-CustomPList.Qty FROM " & _
         "CustomInvoiceItems INNER JOIN CustomPList ON CustomInvoiceItems.EntryID=CustomPList.OrderItemID WHERE " & _
         "CustomPList.ID IN(" & strDelListPL & ")"
         
        con.Execute "DELETE FROM CustomPList WHERE ID IN(" & strDelListPL & ")"
        
        con.Execute "DELETE FROM CustomInvoiceItems WHERE Qty=0"
    End If
     
    
    
    Dim lEntryID As Long, iRowAffect As Integer
    Dim lCII_EntryID As Long, lCustomPList_RefID As Long
    
    For i = 0 To iCount
    
        lEntryID = lEntryIDs(i, 0)
        dQty = lEntryIDs(i, 1)
        dPrice = lEntryIDs(i, 2)
        dCustomPrice = lEntryIDs(i, 3)
        
        con.Execute "UPDATE CustomInvoiceItems SET  Qty=" & dQty & ",Price=" & dPrice & ",CustomPrice=" & _
         dCustomPrice & " WHERE RefID=" & lEntryID & " AND CustomInvoice='" & txtPInvoice.Text & "'", iRowAffect
        
        If iRowAffect = 0 Then
            con.Execute "INSERT INTO CustomInvoiceItems(CustomInvoice,RefID,Qty,Price,CustomPrice) VALUES('" & _
             txtPInvoice.Text & "'," & lEntryID & "," & dQty & "," & dPrice & "," & dCustomPrice & ")"
        End If
               
        
        lCII_EntryID = GetSingleLongValue("EntryID", "CustomInvoiceItems", " WHERE CustomInvoice='" & txtPInvoice & "' AND RefID=" & lEntryID)
             
        For j = 1 To LV2.ListItems.count
            With LV2.ListItems(j)
                If lEntryID = Val(.ListSubItems(1).Tag) And Val(.Tag) = 0 Then
                    con.Execute "INSERT INTO CustomPlist(CustomInvoice,OrderItemID,CartonFrom,CartonTo,Qty,LotNo) VALUES('" & _
                     txtPInvoice.Text & "'," & lCII_EntryID & "," & Val(.Text) & "," & Val(.Text) & "," & _
                     Val(.SubItems(4)) & ",'" & .SubItems(5) & "')"
                
                End If
                If Val(.ListSubItems(6).Tag) > 0 And Val(.Tag) = 0 And lEntryID = Val(.ListSubItems(1).Tag) Then
                    CustomPList_DispatchListDetail = GetSingleLongValue("MAX(ID)", "CustomPList")
                    con.Execute "INSERT INTO CustomPList_DispatchListDetail(CustomPlist_RefID,ItemCode,CartonNo,DP_RefID,FromOrderNo) VALUES(" & CustomPList_DispatchListDetail & _
                     ",'" & .SubItems(7) & "'," & Val(.SubItems(6)) & "," & Val(.ListSubItems(6).Tag) & ",'" & .SubItems(8) & "')"
                End If
            End With
        Next
        
    Next
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            'If CBool(.ListSubItems(11).Tag) Then
            dPrice = Val(.SubItems(8))
            dCustomPrice = Val(.ListSubItems(7).Tag)
            dCustomPrice = dPrice

            con.Execute "UPDATE CustomInvoiceItems SET ItemForLabTest=" & Abs(CBool(.ListSubItems(11).Tag)) & ",Price=" & dPrice & ",CustomPrice=" & dCustomPrice & " WHERE EntryID=" & LV.ListItems(i).ListSubItems(1).Tag
            'End If
        End With
    Next
    
    For i = 1 To LV2.ListItems.count
        With LV2.ListItems(i)
            If Val(.Tag) > 0 Then
                con.Execute "UPDATE CustomPList SET CartonFrom=" & Val(.Text) & ",CartonTo=" & Val(.Text) & ",LotNo='" & .SubItems(5) & "' WHERE ID=" & Val(.Tag)
            End If
        End With
    Next
'    For i = 1 To LV.ListItems.count
'
'        With LV.ListItems(i)
'            dPrice = Val(.SubItems(8))
'            dCustomPrice = Val(.ListSubItems(7).Tag)
'            If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" OR strCompany = "Kami" Then
'                dCustomPrice = dPrice
'            End If
'            If Val(.ListSubItems(1).Tag) = 0 Then
'                con.Execute "INSERT INTO CustomInvoiceItems(CustomInvoice,RefID,Qty,Price," & _
'                "CustomPrice,RebateRate,DTRENo,ItemForLabTest,BatchNo) VALUES('" & txtPInvoice & "'," & Val(.key) & "," & _
'                 Val(.SubItems(9)) & "," & dPrice & "," & _
'                 dCustomPrice & "," & Val(.ListSubItems(9).Tag) & ",'" & .SubItems(11) & "'," & Abs(CBool(.ListSubItems(11).Tag)) & ",'" & .SubItems(12) & "')"
'            Else
'                If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" OR strCompany = "Kami" Then
'                    con.Execute "UPDATE CustomInvoiceItems SET Qty=" & Val(.SubItems(9)) & _
'                         ",CustomPrice=" & dPrice & ",Price=" & dPrice & ",DTRENo='" & .SubItems(11) & "',ItemForLabTest=" & Abs(CBool(.ListSubItems(11).Tag)) & ",BatchNo='" & .SubItems(12) & "' WHERE EntryID=" & Val(.ListSubItems(1).Tag)
'                Else
'                    If bCustomInvoice Then
'                        con.Execute "UPDATE CustomInvoiceItems SET Qty=" & Val(.SubItems(9)) & _
'                         ",CustomPrice=" & dPrice & ",DTRENo='" & .SubItems(11) & "',ItemForLabTest=" & Abs(CBool(.ListSubItems(11).Tag)) & ",BatchNo='" & .SubItems(12) & "' WHERE EntryID=" & Val(.ListSubItems(1).Tag)
'                    Else
'                        con.Execute "UPDATE CustomInvoiceItems SET Qty=" & Val(.SubItems(9)) & _
'                         ",Price=" & dPrice & ",DTRENo='" & .SubItems(11) & "',ItemForLabTest=" & Abs(CBool(.ListSubItems(11).Tag)) & ",BatchNo='" & .SubItems(12) & "' WHERE EntryID=" & Val(.ListSubItems(1).Tag)
'                    End If
'                End If
'            End If
'
'            con.Execute "UPDATE FProformaOrders SET InvQty=InvQty-" & _
'             Val(.SubItems(9)) & " WHERE EntryID=" & Val(.key)
'
'         End With
'
'    Next i
        
    'Packing List Info to be saved here.....
'    For i = 1 To LV2.ListItems.count
'        With LV2.ListItems(i)
'            If Val(.Tag) = 0 Then
'                con.Execute "INSERT INTO CustomPList(CustomInvoice,OrderItemID,CartonFrom,CartonTo,Qty,LotNo) " & _
'                 "VALUES('" & txtPInvoice & "'," & Val(.ListSubItems(1).Tag) & "," & Val(.Text) & "," & Val(.Text) & _
'                 "," & Val(.SubItems(4)) & ",'" & .SubItems(5) & "')"
'            End If
'        End With
'    Next
    
    Call SaveOtherCharges
    
    'Use the Following Queries Until the exact Problem is found & Resolved...
    con.Execute "UPDATE CustomInvoiceItems SET CustomInvoiceItems.Qty=T1.Qty FROM " & _
     "CustomInvoiceItems INNER JOIN (SELECT OrderItemID,SUM(((CartonTo-CartonFrom)+1)*Qty) AS Qty FROM CustomPList GROUP BY OrderItemID) T1" & _
     " ON CustomInvoiceItems.EntryID=T1.OrderItemID WHERE CustomInvoiceItems.CustomInvoice='" & txtPInvoice & "'"
     
    con.Execute "DELETE FROM CustomInvoiceItems WHERE Qty<0 AND CustomInvoice='" & txtPInvoice & "'"
    
    
    'Advance Payments
    If strProforma <> "" Then
        con.Execute "Delete From FCustAdvancePaymentsAdjusted Where CustomInvoice='" & strProforma & "'"
    End If
    
    For i = 1 To LVAdvance.ListItems.count
        With LVAdvance.ListItems(i)
            If Val(.SubItems(4)) > 0 Then
                con.Execute "Insert Into FCustAdvancePaymentsAdjusted(CustomInvoice,RefID,Amount)" & _
                 " Values('" & IIf(strProforma = "", txtPInvoice.Text, strProforma) & "'," & _
                 Val(.key) & "," & Val(.SubItems(4)) & ")"
            End If
        End With
    Next
    con.CommitTrans
    
    Saved = True
    
    Exit Function

err:
    MsgBox err.Description
    
End Function

Private Sub PrintCustomInvoice(p_strCustomInvoice As String, Optional bCustomInvoice As Boolean = True)

    On Error GoTo err

    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    If bCustomInvoice Then
        Set rpt = rptApp.OpenReport(RptPath & "\rptCustomInvoice.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\rptCommercialInvoice.rpt")
    End If
    'rpt.FormulaFields(15).Text = bImpMaterialDetail
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & p_strCustomInvoice & "' ", rpt
    
    Screen.MousePointer = vbDefault


    Exit Sub
err:
    MsgBox err.Description

End Sub

 
Private Sub GetOrders()

    On Error GoTo err
    
    Dim strTTerms As String
    
    With cmbCountry
        If Not .MatchFound Then Exit Sub
    End With
    
    
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub



Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If x > LV.ColumnHeaders(i).Left And x < LV.ColumnHeaders(i).Left + LV.ColumnHeaders(i).Width Then
            iColNo = i
            Exit For
        End If
    Next
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Exit Sub
    If Button <> 2 Then Exit Sub
    If LV.ListItems.count = 0 Then Exit Sub
    Dim strRet As String
    strRet = ShowPopUpMenu1("Mark As Lab Test Item")
    'Remove Tag from All Others...
    
    If strRet = "Mark As Lab Test Item" Then
        Dim i As Integer
        For i = 1 To LV.ListItems.count
            LV.ListItems(i).ListSubItems(11).Tag = "FALSE"
        Next
        LV.SelectedItem.ListSubItems(11).Tag = "TRUE"
    End If
    
End Sub

Private Sub LV2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    Dim myLV As ListView
    Set myLV = LV2
    
    If myLV.SortKey + 1 = ColumnHeader.Index Then
        If myLV.SortOrder = lvwAscending Then
            myLV.SortOrder = lvwDescending

        Else
            myLV.SortOrder = lvwAscending

        End If
    Else
        myLV.SortKey = ColumnHeader.Index - 1
        myLV.SortOrder = lvwAscending

    End If
    If myLV.Sorted = False Then myLV.Sorted = True

End Sub

Private Sub LV2_KeyDown(KeyCode As Integer, Shift As Integer)

    If LV2.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbYes Then
            If Val(LV2.SelectedItem.Tag) > 0 Then
                strDelListPL = strDelListPL & Val(LV2.SelectedItem.Tag) & ","
            End If
            LV2.ListItems.Remove LV2.SelectedItem.Index
        End If
    End If
    
End Sub

Private Sub LV2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i As Integer
    For i = 1 To LV2.ColumnHeaders.count
        With LV2.ColumnHeaders(i)
            If x > Val(.Left) And x < (Val(.Left) + Val(.Width)) Then
                iColNoLV2 = i
                Exit For
            End If
        End With
    Next
    
End Sub

Private Sub LVAdvance_DblClick()

    If LVAdvance.ListItems.count = 0 Then Exit Sub
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(*)", "CustomInvoiceAuthorization", " WHERE CustomInvoice='" & txtpinovice & "'")
    If lCount > 0 Then
        MsgBox "Can't Change Authorization has been done.", vbInformation
        Exit Sub
    End If
    
    With txtAdvAmt
        .Move LVAdvance.Left + LVAdvance.ColumnHeaders(5).Left, LVAdvance.Top + LVAdvance.SelectedItem.Top, LVAdvance.ColumnHeaders(5).Width
        .Text = Val(LVAdvance.SelectedItem.SubItems(4))
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub LVDL_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer, dTotal As Double
    For i = 1 To LVDL.ListItems.count
        If LVDL.ListItems(i).Checked Then
            dTotal = dTotal + Val(LVDL.ListItems(i).SubItems(3))
        End If
    Next
    txtDispQtyTotal = dTotal
End Sub


Private Sub LVOtherCharges_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then
        If MsgBox("Do you want to Delete it ?", vbQuestion + vbYesNo) = vbYes Then
            LVOtherCharges.ListItems.Remove LVOtherCharges.SelectedItem.Index
            Call PrepareForNext
        End If
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

    If SSTab1.Tab = 2 Then
        cmbItemSearch.Visible = True
    Else
        cmbItemSearch.Visible = False
    End If
    
    If SSTab1.Tab = 0 Then
        cmbItemSearchInvoice.Visible = True
    Else
        cmbItemSearchInvoice.Visible = False
    End If
    
End Sub

Private Sub txtAdvAmt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Round(Val(txtAdvAmt), 2) > Round(Val(LVAdvance.SelectedItem.SubItems(2)) - (Val(LVAdvance.SelectedItem.SubItems(3)) - Val(LVAdvance.SelectedItem.SubItems(4))), 2) Then
            MsgBox "Invalid Amount.", vbInformation
            Exit Sub
        ElseIf Val(txtAdvAmt) <= 0 Then
            MsgBox "Invalid Amount.", vbInformation
            Exit Sub
        End If
        LVAdvance.SelectedItem.SubItems(4) = Val(txtAdvAmt)
        LVAdvance.SelectedItem.SubItems(5) = Round(Val(LVAdvance.SelectedItem.SubItems(2)) - Val(LVAdvance.SelectedItem.SubItems(3)) - Val(LVAdvance.SelectedItem.SubItems(4)), 2)
        txtAdvAmt.Visible = False
    ElseIf KeyAscii = 27 Then
        txtAdvAmt.Visible = False
    End If
End Sub

Private Sub txtAdvAmt_LostFocus()
    txtAdvAmt.Visible = False
End Sub

Private Sub txtDTRENo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LV.SelectedItem.SubItems(iColNo - 1) = txtDTRENo.Text
        txtDTRENo.Visible = False
    ElseIf KeyAscii = 27 Then
        txtDTRENo.Visible = False
    End If
End Sub

Private Sub txtDTRENo_LostFocus()
    txtDTRENo.Visible = False
End Sub

Private Sub txtEditLV2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If iColNoLV2 = 1 Then
            LV2.SelectedItem.Text = Val(txtEditLV2.Text)
        Else
            LV2.SelectedItem.SubItems(iColNoLV2 - 1) = txtEditLV2.Text
        End If
        txtEditLV2.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEditLV2.Visible = False
    End If
End Sub

Private Sub txtEditLV2_LostFocus()
    txtEditLV2.Visible = False
End Sub

Private Sub txtEditPLQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txtEditPLQty) <= 0 Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        End If
        LV2.SelectedItem.SubItems(4) = Val(txtEditPLQty)
    ElseIf KeyAscii = 27 Then
        txtEditPLQty.Visible = False
    End If
End Sub

Private Sub txtEditPLQty_LostFocus()
    txtEditPLQty.Visible = False
End Sub


Private Sub txtPrice_KeyPress(KeyAscii As Integer)
        
    On Error GoTo err
    If KeyAscii = 13 Then
'        If Val(txtPrice) <= 0 Then
'            MsgBox "Invalid Price.", vbInformation
'            Exit Sub
'        End If
        LV.SelectedItem.SubItems(8) = Val(txtPrice)
        LV.SelectedItem.SubItems(10) = Val(txtPrice) * LV.SelectedItem.SubItems(9)
        
'        On Error Resume Next
'        LV2.ListItems(LV.SelectedItem.Key).SubItems(8) = Val(LV.SelectedItem.ListSubItems(8).Tag) - Val(txtQty)
'        If err.Number = 35601 Then
'            Resume Next
'        End If
        
        Dim dPrice As Double, dCustomPrice As Double
        
        dPrice = Val(LV.SelectedItem.SubItems(8))
        dCustomPrice = Val(LV.SelectedItem.ListSubItems(7).Tag)
                
        If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Tecno" Or strCompany = "Towne" Then
            dCustomPrice = dPrice
        End If

        con.Execute "UPDATE CustomInvoiceItems SET Price=" & dPrice & ",CustomPrice=" & dCustomPrice & " WHERE EntryID=" & Val(LV.SelectedItem.ListSubItems(1).Tag)
        
        txtPrice.Visible = False
        
        Call CalculateTotals
    ElseIf KeyAscii = 27 Then
        txtPrice.Visible = False
    End If
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    If KeyAscii = 13 Then
    
        If Val(txtQty) > Val(LV.SelectedItem.ListSubItems(8).Tag) And bAcceptsExtraQty = False Then
            MsgBox "Invalid Qty." & vbNewLine & "You can only add " & Val(LV.SelectedItem.ListSubItems(8).Tag) & " Qty.", vbInformation
            Exit Sub
        End If
        
        LV.SelectedItem.SubItems(9) = Val(txtQty)
        LV.SelectedItem.SubItems(10) = Val(txtQty) * LV.SelectedItem.SubItems(8)
        
        On Error Resume Next
        LV2.ListItems(LV.SelectedItem.key).SubItems(8) = Val(LV.SelectedItem.ListSubItems(8).Tag) - Val(txtQty)
        If err.Number = 35601 Then
            Resume Next
        End If
        
        txtQty.Visible = False
        
        Call CalculateTotals
    ElseIf KeyAscii = 27 Then
        txtQty.Visible = False
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub LV2_DblClick()
    
    If LV2.ListItems.count = 0 Then Exit Sub
    If iColNoLV2 <> 6 Then iColNoLV2 = 1
    With txtEditLV2
        .Move LV2.Left + LV2.ColumnHeaders(iColNoLV2).Left, LV2.Top + LV2.SelectedItem.Top, LV2.ColumnHeaders(iColNoLV2).Width
        If iColNoLV2 = 1 Then
            .Text = Val(LV2.SelectedItem.Text)
        Else
            .Text = LV2.SelectedItem.SubItems(iColNoLV2 - 1)
        End If
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
    Exit Sub
    With txtEditPLQty
        .Move LV2.ColumnHeaders(5).Left + LV2.Left, LV2.Top + LV2.SelectedItem.Top
        .Text = Val(LV2.SelectedItem.SubItems(4))
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub

Private Sub txtqty1_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    Dim ITM As ListItem
    
    If KeyAscii = 13 Then
        If Val(txtQty1) > Val(LV2.SelectedItem.ListSubItems(9)) And bAcceptsExtraQty = False Then
            MsgBox "Invalid Qty." & vbNewLine & "You can only add " & Val(LV2.SelectedItem.ListSubItems(9)) & " Qty.", vbInformation
            Exit Sub
        End If
        If Val(txtQty1) <= 0 Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        End If
        
        Call MoveToFirstLV
        
        LV2.SelectedItem.SubItems(9) = LV2.SelectedItem.SubItems(9) - Val(txtQty1)
        LV2.SelectedItem.SubItems(10) = Val(txtQty1) * Val(LV2.SelectedItem.SubItems(8) & "")
        
        txtQty1.Visible = False
        If Val(LV2.SelectedItem.ListSubItems(2).Tag) > 1 Then
            Call AddPacksOtherArticles(LV2.SelectedItem.ListSubItems(5).Tag, LV2.SelectedItem.Index, LV2.SelectedItem.SubItems(2))
        End If
'        Call CalculateTotals
        
        'LV2.SetFocus
        cmbItemSearch.SetFocus
        cmbItemSearch.ID = ""
    ElseIf KeyAscii = 27 Then
        txtQty1.Visible = False
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub AddPacksOtherArticles(iPackNo As Integer, iIndex As Integer, strOrderNo As String)
    
    For i = 1 To LV2.ListItems.count
        
        If LV2.ListItems(i).ListSubItems(5).Tag = iPackNo And strOrderNo = LV2.ListItems(i).SubItems(2) And i <> iIndex Then
            LV2.ListItems(i).Selected = True
            Call MoveToFirstLV
                
            LV2.SelectedItem.SubItems(9) = LV2.SelectedItem.SubItems(9) - Val(txtQty1)
            LV2.SelectedItem.SubItems(10) = Val(txtQty1) * LV2.SelectedItem.SubItems(8)
            txtQty1.Visible = False
        End If
    Next
    Call CalculateTotals
    
End Sub

Private Sub MoveToFirstLV()
    
    On Error GoTo err
    Dim iSNo As Integer
    Dim ITM As ListItem
     
    iSNo = LV.ListItems.count + 1
    
    Set ITM = LV.ListItems.add(, LV2.SelectedItem.key, iSNo)
    'Set ITM = LV2.SelectedItem
ComeHereForEdit:
    For i = 1 To LV2.SelectedItem.ListSubItems.count
        ITM.ListSubItems.add , , LV2.SelectedItem.ListSubItems(i).Text
        ITM.ListSubItems(i).Tag = LV2.SelectedItem.ListSubItems(i).Tag
    Next i
    ITM.ListSubItems(9).Text = txtQty1
    'ITM.ListSubItems(6).Tag = txtQty1
    ITM.SubItems(10) = Val(LV2.SelectedItem.SubItems(8) & "") * Val(txtQty1)
    ITM.SubItems(11) = ""
    ITM.ListSubItems(11).Tag = "False"
    
'    ITM.ListSubItems.Add , , ![OrderNo] & ""
'    ITM.ListSubItems.Add , , ![ItemCode] & ""
'    ITM.ListSubItems.Add , , ![Description] & ""
'    ITM.ListSubItems.Add , , ![Unit] & ""
'    ITM.ListSubItems.Add , , Val(![Price] & "")
'    ITM.ListSubItems.Add , , Val(![InvQty] & "")
'    ITM.ListSubItems(6).Tag = Val(![InvQty] & "")
'    ITM.SubItems(7) = Val(![Price] & "") * Val(![InvQty] & "")

    'Dim iIndex As Integer
    'iIndex = LV2.SelectedItem.Index
    'Set itm = LV.ListItems.Add
    'itm.SubItems(3) = LV2.ListItems(iIndex).SubItems(3)
    
    Exit Sub
err:
    If err.Description = "Key is not unique in collection" And err.Number = 35602 Then
        Set ITM = LV.ListItems(LV2.SelectedItem.key)
        ITM.SubItems(9) = Val(ITM.SubItems(9)) + Val(txtQty1)
        ITM.SubItems(10) = Val(ITM.SubItems(9)) * Val(ITM.SubItems(8))
        
    Else
        MsgBox err.Description
    End If
End Sub

Public Function FormENoExists(p_strInvNo As String, p_strFormENo As String) As Boolean
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim bReturn As Boolean
    If p_strInvNo = "" Then
        strSQL = "Select Count(FormE) From CustomInvoice Where FormE='" & p_strFormENo & "'"
    Else
        strSQL = "Select Count(FormE) From CustomInvoice Where FormE='" & p_strFormENo & "' AND CustomInvoice<>'" & p_strInvNo & "'"
    End If
    
    With rs
        .Open strSQL, con, adOpenForwardOnly, adLockReadOnly
        If .Fields(0) = 0 Then
            bReturn = False
        Else
            bReturn = True
        End If
        .Close
    End With
    Set rs = Nothing
    FormENoExists = bReturn
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Private Sub SaveOtherCharges()

    If strProforma <> "" Then
        'con.Execute "Delete From CustomInvoiceOtherCharges Where CustomInvoice='" & strProforma & "' AND Custom=" & Abs(bCustomInvoice)
        con.Execute "Delete From CustomInvoiceOtherCharges Where CustomInvoice='" & strProforma & "'"
    End If
    
    For i = 1 To LVOtherCharges.ListItems.count
        With LVOtherCharges.ListItems(i)
            con.Execute "INSERT INTO CustomInvoiceOtherCharges(CustomInvoice,ChargesDetail,ChargesAmt,SNo,Custom)" & _
             " VALUES('" & IIf(strProforma = "", txtPInvoice.Text, strProforma) & "','" & _
             .SubItems(1) & "'," & Val(.SubItems(2)) & "," & Val(.Index) & "," & Abs(bCustomInvoice) & ")"
             
             If strProforma = "" Then   'iF nEW Invoice THen save them as commercial....
                con.Execute "INSERT INTO CustomInvoiceOtherCharges(CustomInvoice,ChargesDetail,ChargesAmt,SNo,Custom)" & _
                 " VALUES('" & IIf(strProforma = "", txtPInvoice.Text, strProforma) & "','" & _
                 .SubItems(1) & "'," & Val(.SubItems(2)) & "," & Val(.Index) & ",0)"
             
             End If
        End With
    Next
    
End Sub

Private Sub FillAdvancePayments()
    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "SELECT EntryID,AdviceNo,Amount,AdjustedAmount FROM VFCustAdvancePayments Where CustCode='" & cmbCust.ID & "' AND Amount>AdjustedAmount ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LVAdvance.ListItems.Clear
        Do Until .EOF
            Set ITM = LVAdvance.ListItems.add(, !EntryID & "'", LVAdvance.ListItems.count + 1)
            ITM.ListSubItems.add , , !AdviceNo & ""
            ITM.ListSubItems.add(, , Val(!Amount & "")).Tag = Val(!Amount & "")
            ITM.ListSubItems.add(, , Val(!AdjustedAmount & "")).Tag = Val(!AdjustedAmount & "")
            ITM.ListSubItems.add(, , 0).Tag = 0
            ITM.ListSubItems.add , , Round(Val(!Amount & "") - Val(!AdjustedAmount & ""), 2)
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Opts_Click(Index As Integer)
    Opts(0).Tag = Index
End Sub

Public Sub AddFeedback(p_strCustomInvoice As String)
    txtInvoiceNo = p_strCustomInvoice
    Call ShowInvoiceItems(p_strCustomInvoice)
    Me.Show
End Sub
