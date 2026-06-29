VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmMarketingCustomer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Marketing Customer Info."
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8940
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fTitle 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   45
      TabIndex        =   29
      Top             =   0
      Width           =   8850
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marketing Customer Info."
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
         Left            =   2430
         TabIndex        =   30
         Top             =   120
         Width           =   4035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marketing Customer Info."
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
         Left            =   2445
         TabIndex        =   31
         Top             =   135
         Width           =   4035
      End
   End
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
      Height          =   7275
      Left            =   45
      TabIndex        =   26
      Top             =   570
      Width           =   8850
      Begin ComboList.Usercontrol1 cmbCountry 
         Height          =   285
         Left            =   6150
         TabIndex        =   3
         Tag             =   "Country"
         Top             =   435
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   556
         BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   -1
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         DataField       =   "Phone3"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   42
         Left            =   4620
         MaxLength       =   20
         TabIndex        =   2
         Tag             =   "City"
         Top             =   435
         Width           =   1515
      End
      Begin VB.TextBox txtfields 
         DataField       =   "Phone3"
         DataSource      =   "datPrimaryRS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   165
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   0
         Tag             =   "CustCode"
         Top             =   435
         Width           =   1200
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         Height          =   630
         Left            =   3465
         TabIndex        =   32
         Top             =   6555
         Width           =   5280
         Begin MSForms.CommandButton cmdSave 
            Height          =   360
            Left            =   75
            TabIndex        =   24
            Top             =   195
            Width           =   1665
            Caption         =   "Save & New   "
            PicturePosition =   327683
            Size            =   "2937;635"
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
            TabIndex        =   33
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
         Begin MSForms.CommandButton cmdSaveClose 
            Height          =   360
            Left            =   1815
            TabIndex        =   25
            Top             =   195
            Width           =   1665
            Caption         =   "Save & Close  "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmMarketingCustomer.frx":0000
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
      Begin VB.TextBox txtfields 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1380
         MaxLength       =   255
         TabIndex        =   1
         Tag             =   "Name"
         Top             =   435
         Width           =   3225
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   615
         Top             =   6660
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMarketingCustomer.frx":015E
               Key             =   "5'"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMarketingCustomer.frx":05B0
               Key             =   "1'"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMarketingCustomer.frx":070A
               Key             =   "3'"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMarketingCustomer.frx":0B5C
               Key             =   "2'"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMarketingCustomer.frx":1836
               Key             =   "4'"
            EndProperty
         EndProperty
      End
      Begin ComboList.Usercontrol1 cmbCurrency 
         Height          =   285
         Left            =   7785
         TabIndex        =   4
         Tag             =   "Curr"
         Top             =   435
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   556
         BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   -1
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   5730
         Left            =   150
         TabIndex        =   36
         Top             =   810
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   10107
         _Version        =   393216
         Style           =   1
         Tabs            =   6
         TabsPerRow      =   6
         TabHeight       =   617
         TabCaption(0)   =   " Company Info."
         TabPicture(0)   =   "frmMarketingCustomer.frx":1C88
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblLabels(58)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "chkExtra"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lblLabels(57)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lblLabels(56)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cmbPaymentTerms"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmbTradeTerms"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "cmbDTFormat"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "lblLabels(45)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "lblLabels(30)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "lblLabels(18)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "lblLabels(17)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "lblLabels(16)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "lblLabels(15)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "lblLabels(2)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "lblLabels(3)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "lblLabels(4)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "lblLabels(6)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "lblLabels(8)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "lblLabels(9)"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "lblLabels(60)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "lblLabels(61)"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "chkLateOrderAlert"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "txtfields(41)"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "txtfields(2)"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "txtfields(7)"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "txtfields(6)"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "txtfields(4)"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "txtfields(3)"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "txtfields(29)"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "txtfields(17)"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "txtfields(16)"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "txtfields(15)"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "txtfields(14)"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "txtfields(8)"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "txtfields(43)"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "txtfields(45)"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).ControlCount=   36
         TabCaption(1)   =   " Bank Info."
         TabPicture(1)   =   "frmMarketingCustomer.frx":1CA4
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmbBankCountry"
         Tab(1).Control(1)=   "txtbank(1)"
         Tab(1).Control(2)=   "txtbank(4)"
         Tab(1).Control(3)=   "txtbank(3)"
         Tab(1).Control(4)=   "txtbank(2)"
         Tab(1).Control(5)=   "txtbank(7)"
         Tab(1).Control(6)=   "txtbank(9)"
         Tab(1).Control(7)=   "txtbank(5)"
         Tab(1).Control(8)=   "txtbank(6)"
         Tab(1).Control(9)=   "txtbank(8)"
         Tab(1).Control(10)=   "txtbank(0)"
         Tab(1).Control(11)=   "chkDefaultBank"
         Tab(1).Control(12)=   "cmdCancelBank"
         Tab(1).Control(13)=   "cmdSaveBank"
         Tab(1).Control(14)=   "cmdRemoveBank"
         Tab(1).Control(15)=   "lblLabels(37)"
         Tab(1).Control(16)=   "lblLabels(38)"
         Tab(1).Control(17)=   "lblLabels(39)"
         Tab(1).Control(18)=   "lblLabels(40)"
         Tab(1).Control(19)=   "lblLabels(41)"
         Tab(1).Control(20)=   "lblLabels(42)"
         Tab(1).Control(21)=   "lblLabels(43)"
         Tab(1).Control(22)=   "lblLabels(44)"
         Tab(1).Control(23)=   "lblLabels(47)"
         Tab(1).Control(24)=   "cmdCustBanks"
         Tab(1).Control(25)=   "lblLabels(48)"
         Tab(1).Control(26)=   "Image3"
         Tab(1).ControlCount=   27
         TabCaption(2)   =   " Contact Info."
         TabPicture(2)   =   "frmMarketingCustomer.frx":1CC0
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame2"
         Tab(2).Control(1)=   "Frame1"
         Tab(2).Control(2)=   "Frame4"
         Tab(2).ControlCount=   3
         TabCaption(3)   =   " Agent Info."
         TabPicture(3)   =   "frmMarketingCustomer.frx":1CDC
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtfields(22)"
         Tab(3).Control(1)=   "txtfields(24)"
         Tab(3).Control(2)=   "txtfields(26)"
         Tab(3).Control(3)=   "txtfields(27)"
         Tab(3).Control(4)=   "txtfields(28)"
         Tab(3).Control(5)=   "Frame3"
         Tab(3).Control(6)=   "txtfields(30)"
         Tab(3).Control(7)=   "txtfields(23)"
         Tab(3).Control(8)=   "txtfields(25)"
         Tab(3).Control(9)=   "lblLabels(25)"
         Tab(3).Control(10)=   "lblLabels(27)"
         Tab(3).Control(11)=   "lblLabels(28)"
         Tab(3).Control(12)=   "lblLabels(29)"
         Tab(3).Control(13)=   "lblLabels(32)"
         Tab(3).Control(14)=   "lblLabels(31)"
         Tab(3).Control(15)=   "lblLabels(24)"
         Tab(3).Control(16)=   "lblLabels(49)"
         Tab(3).Control(17)=   "Image1"
         Tab(3).ControlCount=   18
         TabCaption(4)   =   " Accounts Info."
         TabPicture(4)   =   "frmMarketingCustomer.frx":1CF8
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Image5"
         Tab(4).Control(1)=   "cmbDrCr"
         Tab(4).Control(2)=   "lblBalance"
         Tab(4).Control(3)=   "lblLabels(14)"
         Tab(4).Control(4)=   "lblLabels(7)"
         Tab(4).Control(5)=   "lblLabels(12)"
         Tab(4).Control(6)=   "lblLabels(11)"
         Tab(4).Control(7)=   "DtOpenDate"
         Tab(4).Control(8)=   "txtfields(0)"
         Tab(4).Control(9)=   "txtfields(11)"
         Tab(4).ControlCount=   10
         TabCaption(5)   =   "Shipping Addresses"
         TabPicture(5)   =   "frmMarketingCustomer.frx":1D14
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "lblLabels(62)"
         Tab(5).Control(1)=   "LV"
         Tab(5).Control(2)=   "txtShippingAddress"
         Tab(5).Control(3)=   "cmdAddShippingAddress"
         Tab(5).Control(4)=   "FraShippingAddress"
         Tab(5).ControlCount=   5
         Begin ComboList.Usercontrol1 cmbBankCountry 
            Height          =   285
            Left            =   -69075
            TabIndex        =   46
            Top             =   1635
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   556
            BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ListIndex       =   -1
            Appearance      =   0
         End
         Begin VB.Frame FraShippingAddress 
            Height          =   1530
            Left            =   -74685
            TabIndex        =   151
            Top             =   2640
            Visible         =   0   'False
            Width           =   7215
            Begin VB.CommandButton cmdCancelSA 
               Caption         =   "Cancel"
               Height          =   330
               Left            =   6195
               TabIndex        =   154
               Top             =   1110
               Width           =   780
            End
            Begin VB.CommandButton cmdOK 
               Caption         =   "OK"
               Height          =   330
               Left            =   5385
               TabIndex        =   153
               Top             =   1110
               Width           =   780
            End
            Begin VB.TextBox txtEdit 
               Appearance      =   0  'Flat
               DataSource      =   "datPrimaryRS"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   930
               Left            =   45
               MaxLength       =   255
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   152
               Top             =   150
               Width           =   7080
            End
         End
         Begin VB.CommandButton cmdAddShippingAddress 
            Caption         =   "Add"
            Height          =   945
            Left            =   -67485
            TabIndex        =   150
            Top             =   675
            Width           =   780
         End
         Begin VB.TextBox txtShippingAddress 
            Appearance      =   0  'Flat
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   945
            Left            =   -74745
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   148
            Top             =   675
            Width           =   7215
         End
         Begin MSComctlLib.ListView LV 
            Height          =   3525
            Left            =   -74745
            TabIndex        =   147
            Top             =   1635
            Width           =   8040
            _ExtentX        =   14182
            _ExtentY        =   6218
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Shipping Address"
               Object.Width           =   12726
            EndProperty
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Email2"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   45
            Left            =   6615
            MaxLength       =   100
            TabIndex        =   15
            Tag             =   "PaymentDays"
            Top             =   1890
            Width           =   1785
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Fax2"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   43
            Left            =   7305
            MaxLength       =   100
            TabIndex        =   18
            Tag             =   "MaxDiscount"
            Top             =   2460
            Width           =   1110
         End
         Begin VB.TextBox txtbank 
            Appearance      =   0  'Flat
            DataField       =   "Fax1"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   -69075
            MaxLength       =   100
            TabIndex        =   102
            Top             =   1635
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Address1"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   8
            Left            =   180
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Tag             =   "Address"
            Top             =   3000
            Width           =   4155
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Email1"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   14
            Left            =   6435
            MaxLength       =   100
            TabIndex        =   8
            Tag             =   "EMail1"
            Top             =   810
            Width           =   1980
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Email2"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   15
            Left            =   6435
            MaxLength       =   100
            TabIndex        =   12
            Tag             =   "EMail2"
            Top             =   1350
            Width           =   1980
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "URL"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   16
            Left            =   180
            MaxLength       =   100
            TabIndex        =   16
            Tag             =   "URL"
            Top             =   2460
            Width           =   4155
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "FakeAddress"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   17
            Left            =   4350
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   20
            Tag             =   "FakeAddress"
            Top             =   3000
            Width           =   4065
         End
         Begin VB.Frame Frame2 
            Caption         =   "Contact Person 1 :"
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
            Height          =   2265
            Left            =   -74130
            TabIndex        =   90
            Top             =   840
            Width           =   3330
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "ContactPerson"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   5
               Left            =   75
               MaxLength       =   100
               TabIndex        =   95
               Tag             =   "Cont1Name"
               Top             =   465
               Width           =   3135
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "CEmail"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   9
               Left            =   75
               MaxLength       =   100
               TabIndex        =   94
               Tag             =   "Cont1Email"
               Top             =   975
               Width           =   3135
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "CPhone"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   12
               Left            =   75
               MaxLength       =   20
               TabIndex        =   93
               Tag             =   "cont1Phone"
               Top             =   1500
               Width           =   1545
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "Mobile"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   13
               Left            =   1665
               MaxLength       =   20
               TabIndex        =   92
               Tag             =   "Cont1Mobile"
               Top             =   1500
               Width           =   1545
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "CPhone"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   39
               Left            =   900
               MaxLength       =   20
               TabIndex        =   91
               Tag             =   "Cont1Skype"
               Top             =   1875
               Width           =   2310
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   21
               Left            =   120
               TabIndex        =   100
               Top             =   225
               Width           =   405
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "E-Mail"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   5
               Left            =   120
               TabIndex        =   99
               Top             =   750
               Width           =   420
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Phone"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   10
               Left            =   120
               TabIndex        =   98
               Top             =   1290
               Width           =   450
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mobile"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   13
               Left            =   1755
               TabIndex        =   97
               Top             =   1290
               Width           =   450
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Skype ID :"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   54
               Left            =   120
               TabIndex        =   96
               Top             =   1920
               Width           =   750
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Contact Person 2 :"
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
            Height          =   2265
            Left            =   -70755
            TabIndex        =   79
            Top             =   840
            Width           =   3330
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "Mobile"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   18
               Left            =   1665
               MaxLength       =   20
               TabIndex        =   84
               Tag             =   "Cont2Mobile"
               Top             =   1500
               Width           =   1545
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "CPhone"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   19
               Left            =   75
               MaxLength       =   20
               TabIndex        =   83
               Tag             =   "Cont2Phone"
               Top             =   1500
               Width           =   1545
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "CEmail"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   20
               Left            =   75
               MaxLength       =   100
               TabIndex        =   82
               Tag             =   "Cont2Email"
               Top             =   975
               Width           =   3135
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "ContactPerson"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   21
               Left            =   75
               MaxLength       =   100
               TabIndex        =   81
               Tag             =   "Cont2Name"
               Top             =   465
               Width           =   3135
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "CPhone"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   40
               Left            =   900
               MaxLength       =   20
               TabIndex        =   80
               Tag             =   "Cont2Skype"
               Top             =   1875
               Width           =   2310
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mobile"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   19
               Left            =   1755
               TabIndex        =   89
               Top             =   1290
               Width           =   450
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Phone"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   20
               Left            =   120
               TabIndex        =   88
               Top             =   1290
               Width           =   450
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "E-Mail"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   22
               Left            =   120
               TabIndex        =   87
               Top             =   750
               Width           =   420
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   23
               Left            =   120
               TabIndex        =   86
               Top             =   225
               Width           =   405
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Skype ID :"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   55
               Left            =   120
               TabIndex        =   85
               Top             =   1920
               Width           =   750
            End
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Fax1"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   22
            Left            =   -69000
            MaxLength       =   100
            TabIndex        =   78
            Tag             =   "AgentFax2"
            Top             =   2715
            Width           =   1560
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Fax2"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   24
            Left            =   -70590
            MaxLength       =   100
            TabIndex        =   77
            Tag             =   "AgentFax1"
            Top             =   2715
            Width           =   1560
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Phone2"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   26
            Left            =   -69000
            MaxLength       =   20
            TabIndex        =   76
            Tag             =   "AgentPhone2"
            Top             =   2160
            Width           =   1560
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Phone1"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   27
            Left            =   -70590
            MaxLength       =   20
            TabIndex        =   75
            Tag             =   "AgentPhone1"
            Top             =   2160
            Width           =   1560
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Fax1"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   28
            Left            =   -73995
            MaxLength       =   100
            TabIndex        =   74
            Tag             =   "AgentEmail"
            Top             =   1620
            Width           =   3150
         End
         Begin VB.Frame Frame3 
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
            Height          =   1365
            Left            =   -72675
            TabIndex        =   65
            Top             =   3015
            Width           =   5250
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "Mobile"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   31
               Left            =   3525
               MaxLength       =   20
               TabIndex        =   69
               Tag             =   "ACMobile"
               Top             =   375
               Width           =   1620
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "CPhone"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   32
               Left            =   3525
               MaxLength       =   20
               TabIndex        =   68
               Tag             =   "ACPhone"
               Top             =   900
               Width           =   1620
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "CEmail"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   33
               Left            =   150
               MaxLength       =   100
               TabIndex        =   67
               Tag             =   "ACEmail"
               Top             =   930
               Width           =   3300
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "ContactPerson"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   34
               Left            =   150
               MaxLength       =   100
               TabIndex        =   66
               Tag             =   "ACName"
               Top             =   405
               Width           =   3300
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mobile"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   33
               Left            =   3525
               TabIndex        =   73
               Top             =   150
               Width           =   450
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Phone"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   34
               Left            =   3555
               TabIndex        =   72
               Top             =   690
               Width           =   450
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "E-Mail"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   35
               Left            =   180
               TabIndex        =   71
               Top             =   720
               Width           =   420
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   36
               Left            =   180
               TabIndex        =   70
               Top             =   195
               Width           =   405
            End
         End
         Begin VB.TextBox txtbank 
            Appearance      =   0  'Flat
            DataField       =   "Fax1"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   -70785
            MaxLength       =   100
            TabIndex        =   64
            Top             =   2160
            Width           =   1650
         End
         Begin VB.TextBox txtbank 
            Appearance      =   0  'Flat
            DataField       =   "Phone2"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   -72465
            MaxLength       =   20
            TabIndex        =   63
            Top             =   2160
            Width           =   1635
         End
         Begin VB.TextBox txtbank 
            Appearance      =   0  'Flat
            DataField       =   "Phone1"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   -74160
            MaxLength       =   20
            TabIndex        =   62
            Top             =   2160
            Width           =   1635
         End
         Begin VB.TextBox txtbank 
            Appearance      =   0  'Flat
            DataField       =   "Fax2"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   -70785
            MaxLength       =   100
            TabIndex        =   61
            Top             =   2670
            Width           =   3315
         End
         Begin VB.TextBox txtbank 
            Appearance      =   0  'Flat
            DataField       =   "Address"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   9
            Left            =   -70785
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   60
            Top             =   3180
            Width           =   3315
         End
         Begin VB.TextBox txtbank 
            Appearance      =   0  'Flat
            DataField       =   "Fax1"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   -69075
            MaxLength       =   100
            TabIndex        =   59
            Top             =   2160
            Width           =   1620
         End
         Begin VB.TextBox txtfields 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   11
            Left            =   -73335
            MaxLength       =   10
            TabIndex        =   58
            Top             =   2445
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "AccNo"
            DataSource      =   "datPrimaryRS"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   -73335
            TabIndex        =   57
            Tag             =   "AccNo"
            Top             =   1215
            Width           =   3750
         End
         Begin VB.TextBox txtbank 
            Appearance      =   0  'Flat
            DataField       =   "Fax1"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   -74160
            MaxLength       =   100
            TabIndex        =   56
            Top             =   2670
            Width           =   3330
         End
         Begin VB.TextBox txtbank 
            Appearance      =   0  'Flat
            DataField       =   "Address"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   8
            Left            =   -74175
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   55
            Top             =   3180
            Width           =   3330
         End
         Begin VB.TextBox txtbank 
            Appearance      =   0  'Flat
            DataField       =   "Fax1"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   -74160
            MaxLength       =   100
            TabIndex        =   54
            Top             =   1635
            Width           =   5010
         End
         Begin VB.CheckBox chkDefaultBank 
            Caption         =   "Current Cutomer Bank"
            Height          =   210
            Left            =   -74175
            TabIndex        =   53
            Top             =   4185
            Width           =   1980
         End
         Begin VB.CommandButton cmdCancelBank 
            Caption         =   "Cancel"
            Height          =   300
            Left            =   -68685
            TabIndex        =   52
            Top             =   4110
            Width           =   1230
         End
         Begin VB.CommandButton cmdSaveBank 
            Caption         =   "Save"
            Height          =   300
            Left            =   -69960
            TabIndex        =   51
            Top             =   4110
            Width           =   1230
         End
         Begin VB.CommandButton cmdRemoveBank 
            Caption         =   "Remove"
            Height          =   300
            Left            =   -71235
            TabIndex        =   50
            Top             =   4110
            Width           =   1230
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Fax1"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   30
            Left            =   -70590
            MaxLength       =   100
            TabIndex        =   49
            Tag             =   "AgentURL"
            Top             =   1620
            Width           =   3150
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Address"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Index           =   23
            Left            =   -73995
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   48
            Tag             =   "AgentAddress"
            Top             =   2160
            Width           =   3150
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Fax1"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   25
            Left            =   -73995
            MaxLength       =   100
            TabIndex        =   47
            Tag             =   "AgentName"
            Top             =   1095
            Width           =   6555
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Fax1"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   29
            Left            =   2250
            MaxLength       =   100
            TabIndex        =   10
            Tag             =   "Fax2"
            Top             =   1350
            Width           =   2040
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Phone1"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   180
            MaxLength       =   20
            TabIndex        =   5
            Tag             =   "Phone1"
            Top             =   810
            Width           =   2040
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Phone2"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   2250
            MaxLength       =   20
            TabIndex        =   6
            Tag             =   "Phone2"
            Top             =   810
            Width           =   2040
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Phone3"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   4320
            MaxLength       =   20
            TabIndex        =   7
            Tag             =   "Phone3"
            Top             =   810
            Width           =   2040
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Fax2"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   7
            Left            =   4320
            MaxLength       =   100
            TabIndex        =   11
            Tag             =   "Fax3"
            Top             =   1350
            Width           =   2040
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Fax1"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   180
            MaxLength       =   100
            TabIndex        =   9
            Tag             =   "Fax1"
            Top             =   1350
            Width           =   2040
         End
         Begin VB.Frame Frame4 
            Caption         =   "Misc. Information :"
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
            Height          =   1305
            Left            =   -74130
            TabIndex        =   37
            Top             =   3165
            Width           =   6720
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "Mobile"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   35
               Left            =   4710
               MaxLength       =   20
               TabIndex        =   41
               Tag             =   "DHLNo"
               Top             =   915
               Width           =   1935
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "CPhone"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   36
               Left            =   4710
               MaxLength       =   20
               TabIndex        =   40
               Tag             =   "FedexNo"
               Top             =   420
               Width           =   1935
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "CEmail"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   37
               Left            =   75
               MaxLength       =   100
               TabIndex        =   39
               Tag             =   "NearestRailwaystation"
               Top             =   915
               Width           =   4605
            End
            Begin VB.TextBox txtfields 
               Appearance      =   0  'Flat
               DataField       =   "ContactPerson"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   38
               Left            =   75
               MaxLength       =   100
               TabIndex        =   38
               Tag             =   "NearestAirport"
               Top             =   420
               Width           =   4605
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "DHL #"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   50
               Left            =   4725
               TabIndex        =   45
               Top             =   720
               Width           =   450
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fedex #"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   51
               Left            =   4725
               TabIndex        =   44
               Top             =   225
               Width           =   615
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nearest Railwaystation"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   52
               Left            =   120
               TabIndex        =   43
               Top             =   720
               Width           =   1665
            End
            Begin VB.Label lblLabels 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nearest Airport"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   53
               Left            =   120
               TabIndex        =   42
               Top             =   225
               Width           =   1110
            End
         End
         Begin VB.TextBox txtfields 
            Appearance      =   0  'Flat
            DataField       =   "Address1"
            DataSource      =   "datPrimaryRS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1260
            Index           =   41
            Left            =   180
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Tag             =   "SpecialInstructions"
            Top             =   4050
            Width           =   8235
         End
         Begin MSComCtl2.DTPicker DtOpenDate 
            Height          =   285
            Left            =   -71745
            TabIndex        =   101
            Top             =   1845
            Visible         =   0   'False
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "MMMM dd/MM/yyyy"
            Format          =   299106307
            CurrentDate     =   37055
         End
         Begin MSForms.CheckBox chkLateOrderAlert 
            Height          =   315
            Left            =   2055
            TabIndex        =   23
            Tag             =   "LateOrderAlerts"
            Top             =   5370
            Width           =   1710
            BackColor       =   -2147483633
            ForeColor       =   4194304
            DisplayStyle    =   4
            Size            =   "3016;556"
            Value           =   "0"
            Caption         =   "Late Order Alert"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Shipping Address"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   62
            Left            =   -74700
            TabIndex        =   149
            Top             =   465
            Width           =   1230
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Days After B/L"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   61
            Left            =   6645
            TabIndex        =   146
            Top             =   1650
            Width           =   1725
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max. Discount"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   60
            Left            =   7305
            TabIndex        =   145
            Top             =   2250
            Width           =   1020
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Original Address"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   9
            Left            =   210
            TabIndex        =   142
            Top             =   2790
            Width           =   1170
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax 3"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   8
            Left            =   4320
            TabIndex        =   141
            Top             =   1140
            Width           =   405
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone3"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   6
            Left            =   4320
            TabIndex        =   140
            Top             =   600
            Width           =   540
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone2"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   4
            Left            =   2250
            TabIndex        =   139
            Top             =   600
            Width           =   540
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone1"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   3
            Left            =   210
            TabIndex        =   138
            Top             =   585
            Width           =   540
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax 1"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   2
            Left            =   210
            TabIndex        =   137
            Top             =   1125
            Width           =   405
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email 2"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   15
            Left            =   6435
            TabIndex        =   136
            Top             =   1140
            Width           =   495
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email 1"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   16
            Left            =   6435
            TabIndex        =   135
            Top             =   585
            Width           =   495
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Company URL "
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   17
            Left            =   210
            TabIndex        =   134
            Top             =   2220
            Width           =   1050
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fake Address"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   18
            Left            =   4365
            TabIndex        =   133
            Top             =   2790
            Width           =   975
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax 2"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   25
            Left            =   -68985
            TabIndex        =   132
            Top             =   2535
            Width           =   405
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone2"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   27
            Left            =   -68955
            TabIndex        =   131
            Top             =   1980
            Width           =   540
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone1"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   28
            Left            =   -70560
            TabIndex        =   130
            Top             =   1965
            Width           =   540
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax 1"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   29
            Left            =   -70560
            TabIndex        =   129
            Top             =   2505
            Width           =   405
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "URL"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   32
            Left            =   -70560
            TabIndex        =   128
            Top             =   1410
            Width           =   285
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax 2"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   30
            Left            =   2250
            TabIndex        =   127
            Top             =   1140
            Width           =   405
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Original Address"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   37
            Left            =   -74160
            TabIndex        =   126
            Top             =   2970
            Width           =   1170
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone2"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   38
            Left            =   -72450
            TabIndex        =   125
            Top             =   1935
            Width           =   540
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone1"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   39
            Left            =   -74130
            TabIndex        =   124
            Top             =   1935
            Width           =   540
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax 1"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   40
            Left            =   -70770
            TabIndex        =   123
            Top             =   1935
            Width           =   405
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email 2"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   41
            Left            =   -70770
            TabIndex        =   122
            Top             =   2460
            Width           =   495
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email 1"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   42
            Left            =   -74145
            TabIndex        =   121
            Top             =   2475
            Width           =   495
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fake Address"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   43
            Left            =   -70755
            TabIndex        =   120
            Top             =   2970
            Width           =   975
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fax 2"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   44
            Left            =   -69060
            TabIndex        =   119
            Top             =   1935
            Width           =   405
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date Format"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   45
            Left            =   4365
            TabIndex        =   118
            Top             =   2250
            Width           =   900
         End
         Begin MSForms.ComboBox cmbDTFormat 
            DataField       =   "DateFormat"
            Height          =   285
            Left            =   4350
            TabIndex        =   17
            Tag             =   "DTFormat"
            Top             =   2460
            Width           =   2925
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5159;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Opening Date"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   11
            Left            =   -71715
            TabIndex        =   117
            Top             =   1605
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Opening Bal."
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   12
            Left            =   -73335
            TabIndex        =   116
            Top             =   2220
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Acc #"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   7
            Left            =   -73305
            TabIndex        =   115
            Top             =   1005
            Width           =   420
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Balance"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   14
            Left            =   -70830
            TabIndex        =   114
            Top             =   2235
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label lblBalance 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   -71175
            TabIndex        =   113
            Top             =   2460
            Visible         =   0   'False
            Width           =   1530
         End
         Begin MSForms.ComboBox cmbDrCr 
            Height          =   285
            Left            =   -71970
            TabIndex        =   112
            Top             =   2445
            Visible         =   0   'False
            Width           =   720
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "1270;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   31
            Left            =   -73965
            TabIndex        =   111
            Top             =   1410
            Width           =   420
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Name"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   47
            Left            =   -74100
            TabIndex        =   110
            Top             =   1440
            Width           =   795
         End
         Begin MSForms.CommandButton cmdCustBanks 
            Height          =   345
            Left            =   -73530
            TabIndex        =   109
            Top             =   1020
            Width           =   2070
            Caption         =   "Customer Banks        "
            PicturePosition =   196613
            Size            =   "3651;609"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Country"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   48
            Left            =   -69060
            TabIndex        =   108
            Top             =   1395
            Width           =   585
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   24
            Left            =   -73995
            TabIndex        =   107
            Top             =   1935
            Width           =   585
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Agent Name"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   49
            Left            =   -73935
            TabIndex        =   106
            Top             =   870
            Width           =   885
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   -73665
            Picture         =   "frmMarketingCustomer.frx":1D30
            Top             =   3405
            Width           =   480
         End
         Begin VB.Image Image5 
            Height          =   615
            Left            =   -74940
            Picture         =   "frmMarketingCustomer.frx":2172
            Top             =   420
            Width           =   690
         End
         Begin VB.Image Image3 
            Height          =   480
            Left            =   -74115
            Picture         =   "frmMarketingCustomer.frx":25E8
            Top             =   900
            Width           =   480
         End
         Begin MSForms.ComboBox cmbTradeTerms 
            DataField       =   "DateFormat"
            Height          =   315
            Left            =   180
            TabIndex        =   13
            Tag             =   "TradeTerms"
            Top             =   1890
            Width           =   2565
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4524;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbPaymentTerms 
            DataField       =   "DateFormat"
            Height          =   315
            Left            =   2775
            TabIndex        =   14
            Tag             =   "PaymentTerms"
            Top             =   1890
            Width           =   3810
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "6720;556"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Trade Terms"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   56
            Left            =   210
            TabIndex        =   105
            Top             =   1680
            Width           =   900
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Terms"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   57
            Left            =   2805
            TabIndex        =   104
            Top             =   1680
            Width           =   1110
         End
         Begin MSForms.CheckBox chkExtra 
            Height          =   315
            Left            =   165
            TabIndex        =   22
            Tag             =   "AcceptsExtraQty"
            Top             =   5355
            Width           =   1710
            BackColor       =   -2147483633
            ForeColor       =   4194304
            DisplayStyle    =   4
            Size            =   "3016;556"
            Value           =   "0"
            Caption         =   "Accepts Extra Qty"
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Special Instructions"
            ForeColor       =   &H00400000&
            Height          =   195
            Index           =   58
            Left            =   195
            TabIndex        =   103
            Top             =   3825
            Width           =   1395
         End
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "City"
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
         Index           =   59
         Left            =   4605
         TabIndex        =   144
         Top             =   240
         Width           =   330
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Currency"
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
         Index           =   46
         Left            =   7785
         TabIndex        =   143
         Top             =   225
         Width           =   810
      End
      Begin MSForms.CheckBox chkinactive 
         Height          =   315
         Left            =   165
         TabIndex        =   35
         Tag             =   "Active"
         Top             =   6825
         Width           =   1785
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "3149;556"
         Value           =   "0"
         Caption         =   "Customer is inactive"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
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
         Index           =   26
         Left            =   6150
         TabIndex        =   34
         Top             =   225
         Width           =   720
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
         Left            =   1395
         TabIndex        =   28
         Top             =   225
         Width           =   1320
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer ID"
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
         Left            =   165
         TabIndex        =   27
         Top             =   225
         Width           =   1050
      End
   End
   Begin VB.Menu mnupop 
      Caption         =   "mnupop"
      Visible         =   0   'False
      Begin VB.Menu mnunewbank 
         Caption         =   "< Add New >"
      End
      Begin VB.Menu dash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBanks 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMarketingCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BnkIndex As Integer

Dim CustBanks As New Collection
Dim CustBank As New BankInfo
Dim BanksRemoved As String

Dim CustID As String
Dim Acno As String

Dim strDel As String

Private Sub chkInactive_Click()
   If chkInactive = vbChecked And Val(lblBalance) <> 0 Then
      MsgBox "This Customer Cannot Be Made Inactive," & vbNewLine & _
             "To Make A Customer Inactive Balance Must Be Zero", vbInformation
      chkInactive = vbUnchecked
   End If
End Sub

Private Sub cmbBankCountry_matched()
    txtbank(1) = cmbBankCountry.Text
End Sub

Private Sub cmbCountry_matched()
    
    If CustID = "" Then
        If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
            Dim strCountry As String, iNumber As Integer
            strCountry = cmbCountry.Text
            Dim rs As New ADODB.Recordset
            With rs
                .Open "SELECT MAX(CAST(RIGHT(CustCode,3) AS INT)) FROM ForeignCustomers_Marketing WHERE Country='" & strCountry & "' AND LEFT(CustCode,3)='" & Left(strCountry, 3) & "'", con, adOpenForwardOnly, adLockReadOnly
                iNumber = Val(.Fields(0).Value & "")
                .Close
            End With
            iNumber = iNumber + 1
            strCustCode = UCase(Left(strCountry, 3)) & "-" & Format(iNumber, "000")
            txtfields(10) = strCustCode
            txtfields(10).Locked = True
        End If
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
   'If cmbType.MatchFound Then Call CmbType_Click
End Sub

Private Sub cmdAddShippingAddress_Click()
    LV.ListItems.add(, , txtShippingAddress).Tag = "0"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelBank_Click()
    Call showBankInfo(BnkIndex)
End Sub

Private Sub cmdCancelSA_Click()
    FraShippingAddress.Visible = False
End Sub

Private Sub cmdOK_Click()
    LV.SelectedItem.Text = txtEdit.Text
    FraShippingAddress.Visible = False
End Sub

Private Sub cmdRemoveBank_Click()
If MsgBox("Remove This Bank ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
   BanksRemoved = BanksRemoved & CustBanks(BnkIndex).BankID & ","
   Call CustBanks.Remove(BnkIndex)
   If CustBanks.count = 0 Then
      showBankInfo
   Else
      showBankInfo (1)
   End If
   
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
    If Trim(txtfields(1)) = "" Then
        MsgBox "Invalid Customer Name", vbInformation
        Exit Sub
    End If

    If Saved Then
        Unload Me
        Dim f As New frmMarketingCustomer
        Load f
        f.AddNew
    End If
    
    Exit Sub

err:
 MsgBox "ERROR : # " & err.Number & vbNewLine & err.Description, vbInformation
 
End Sub



Private Sub cmdSaveBank_Click()

   If BnkIndex = 0 Then
      Call UpdateBankInfo
   Else
      Call UpdateBankInfo(CustBanks(BnkIndex))
   End If
   
 
   
End Sub

Private Sub cmdSaveClose_Click()

    On Error GoTo err
    
    If Trim(txtfields(1)) = "" Then
        MsgBox "Invalid Customer Name", vbInformation
        Exit Sub
    End If

    If Saved Then Unload Me

    Exit Sub
err:
    MsgBox "ERROR : # " & err.Number & vbNewLine & err.Description, vbInformation
 
End Sub

Private Sub cmdSClose_Click()

End Sub

Private Sub CmdCustBanks_Click()
   
   
    With mnuBanks

       For i = 1 To .count - 1
           Unload .Item(i)
       Next
    
       Dim defaultmenu As Menu
       Set defaultmenu = Nothing
    
       For i = 1 To CustBanks.count
           Load mnuBanks(i)
           mnuBanks(i).Caption = CustBanks(i).BankName & " (" & CustBanks(i).Country & ")"
           mnuBanks(i).Visible = True
           mnuBanks(i).Tag = CustBanks(i).BankID
           mnuBanks(i).Checked = (i = BnkIndex)
           If CustBanks(i).DefaultBank Then
               Set defaultmenu = mnuBanks(i)
           End If
         
        Next
   
    End With
        
    If defaultmenu Is Nothing Then
        PopUpMenu mnuPop, 2, FAll.Left + SSTab1.Left + cmdCustBanks.Left _
         , FAll.Top + SSTab1.Top + cmdCustBanks.Top + cmdCustBanks.Height
    Else
        PopUpMenu mnuPop, 2, FAll.Left + SSTab1.Left + cmdCustBanks.Left _
         , FAll.Top + SSTab1.Top + cmdCustBanks.Top + cmdCustBanks.Height, defaultmenu
    End If

End Sub

Private Sub DtOpenDate_Change()
    Call setDateRange(DtOpenDate)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    
    strDel = ""
    Me.KeyPreview = True
    CustID = ""
    cmbBankCountry.ListHeight = 1500
    DtOpenDate = ServerDate
    BnkIndex = 0
   
    cmbCurrency.ListHeight = 2000
    cmbCountry.ListHeight = 2000
    cmbCountry.AddVals con, "Distinct CountryName", "Countries", "CountryName", "Order By CountryName"
    cmbCurrency.AddVals con, "Distinct Curr", "ForeignCustomers_Marketing", "Curr"
    cmbCurrency.Text = "EUR"
    cmbBankCountry.AddVals con, "CountryName", "Countries", "CountryName", "Where CountryName<>'Pak' Order By CountryName"
   
    For i = 0 To 4
        SSTab1.Tab = i
        SSTab1.Picture = ImageList1.ListImages(i + 1 & "'").ExtractIcon
    Next
   
    SSTab1.Tab = 0
   
    With cmbDTFormat
        .AddItem "dd-MM-yyyy"
        .AddItem "dd-MMM-yyyy"
        .AddItem "MM-dd-yy"
        .AddItem "MMM-dd-yyyy"
        .AddItem "MMMM-dd-yyyy"
        .AddItem "yy-MM-dd"
        .AddItem "yyyy-MM-dd"
        .AddItem "yyyy-MMM-dd"
        .ListIndex = 1
    End With
    Call AddToCombo(cmbPaymentTerms, "Description", "PaymentTerms")
    Call AddToCombo(cmbTradeTerms, "Description", "TradeTerms")
    
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
    Set rsCust = Nothing
    For i = 1 To CustBanks.count
        CustBanks.Remove 1
    Next i
    Set CustBanks = Nothing
    Set CustBank = Nothing
    Set frmFCustomer = Nothing
    
End Sub


Private Sub ShowData()

End Sub



Private Sub GetAccInfo(AccNo As String)

    Dim rsAcc As New ADODB.Recordset

    With rsAcc
        .Open "Select AccNo,AccTitle,OpenDate,OpenBal,Balance,Active,subAccOf from Accounts Where AccNo='" & AccNo & "'", con, adOpenForwardOnly
   
        AccNo = ![AccNo]
        DtOpenDate = ![OpenDate]
        txtfields(11) = Abs(![OpenBal])
        txtfields(1) = ![AccTitle]
   
        If ![SubAccOf] <> Acno Then cmbType = con.Execute("Select AccTitle From Accounts Where AccNo='" & ![SubAccOf] & "'").Fields(0)
   
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


Private Sub AddAccountInfo()
 Call cmbDrCr_Click
 Dim Active As Boolean
 Active = Active = chkInactive <> vbChecked
 Call NewAccRcv(txtfields(1), DtOpenDate, Val(txtfields(11).Tag), Active, Left(txtfields(0), Len(txtfields(0)) - 4))
End Sub
Private Function AccDefined() As Boolean
'Get Accounts Receiveable AccountNo
Acno = ACCRECNO
If Acno = "" Then
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

      con.Execute "Update Accounts Set accTitle='" & txtfields(1) & "',OpenBal=" & Val(txtfields(11).Tag) & ",Balance=balance + " & Diff & ",Active=" & Active & " Where Accno='" & txtfields(0) & "'"
      Call adjustLedger(txtfields(0), DtOpenDate, 0, Diff, True)

If Diff <> 0 Then Call UpdateAdvEntry(txtfields(0), Val(txtfields(11).Tag))
 
End Sub

Private Sub Form_Unload(cancel As Integer)
    OurOwnForm = True
    Set CustBanks = Nothing
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    
    txtEdit.Text = LV.SelectedItem.Text
    FraShippingAddress.Visible = True
    txtEdit.SetFocus
    
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    FraShippingAddress.Visible = False
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)

    If LV.ListItems.count = 0 Then Exit Sub
    
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo) = vbYes Then
            If Val(LV.SelectedItem.Tag) <> 0 Then
                strDel = strDel & Val(LV.SelectedItem.Tag) & ","
            End If
            LV.ListItems.Remove LV.SelectedItem.Index
        End If
    End If
    
End Sub

Private Sub mnuBanks_Click(Index As Integer)
Call showBankInfo(Index)
End Sub

Private Sub mnunewbank_Click()
  BnkIndex = 0
  For i = 0 To txtbank.count - 1
   txtbank(i) = ""
   txtbank(1) = txtfields(25)
  Next
  
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

    If CustID <> "" Then Exit Sub
   
    If SSTab1.Tab = 4 And Trim(txtfields(1)) <> "" Then
        txtfields(0) = NewAccRcv(txtfields(1) & " (" & cmbCountry.Text & ")", DtOpenDate, Val(txtfields(11).Tag), (chkInactive <> vbChecked), FCUSTACCNO, True)
    End If
    
End Sub

Private Sub txtbank_Change(Index As Integer)

If Index = 1 Then
   If cmbBankCountry.Text <> txtbank(1).Text Then cmbBankCountry.Text = txtbank(1)
End If

End Sub

Private Sub txtfields_KeyPress(Index As Integer, KeyAscii As Integer)
'  Select Case Index
'     Case 11
'        KeyAscii = OnlyNumber(KeyAscii, True)
'     Case 2, 3, 4, 6, 7, 12, 13
'        If Not Chr(KeyAscii) = "-" Then KeyAscii = OnlyNumber(KeyAscii)
'  End Select
End Sub

Private Sub ClearFields()
   For i = 0 To txtfields.count - 1
      txtfields(i) = ""
   Next
   txtfields(1).SetFocus
End Sub



Private Function CreateInsertSQL(strTable As String) As String
    
    Dim c As Control
    Dim RetString As String
    
    
    'Create The Field Name Portion of Qry...
    RetString = "Insert Into " & strTable & "("
    '---------------------------------------------------------------------------
    For Each c In Controls
        
        If (TypeOf c Is VB.TextBox) Or (TypeOf c Is VB.ComboBox) Or (TypeOf c Is MSForms.ComboBox) Or (TypeOf c Is MSForms.CheckBox) Or (TypeOf c Is ComboList.Usercontrol1) Then
            If c.Tag <> "" Then
                RetString = RetString & c.Tag & ","
            End If
        End If
      
    Next
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    RetString = RetString & ") Values("
    
    
    ' Create The Values Portion of Qry....
    '---------------------------------------------------------------------------
    For Each c In Controls
        
        
        If c.Tag <> "" Then
            If (TypeOf c Is VB.TextBox) Or (TypeOf c Is VB.ComboBox) Or (TypeOf c Is MSForms.ComboBox) Then
            
                If c.Name = "NBox" Then
                    RetString = RetString & Val(c.Text & "") & ","
                Else
                    RetString = RetString & "'" & c.Text & "',"
                End If
            ElseIf (TypeOf c Is MSForms.CheckBox) Then
                If c.Name = "chkinactive" Then
                    RetString = RetString & Abs(Not c.Value) & ","
                Else
                    RetString = RetString & Abs(c.Value) & ","
                End If
            ElseIf (TypeOf c Is ComboList.Usercontrol1) Then
                RetString = RetString & "'" & c.Text & "',"
            End If
        End If
       
    Next
    
      
    'Finalaize the Qry
    '---------------------------------------------------------------------------
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    RetString = RetString & ")"
    CreateInsertSQL = RetString
    
End Function
Private Function CreateUpdateSQL(strTable As String) As String
    
    Dim c As Control
    Dim RetString As String
    
    'Create The First Part Of Insert Query...
    RetString = "Update " & strTable & " Set "
    '---------------------------------------------------------------------------
    For Each c In Controls
    
        If c.Tag <> "" Then
            If (TypeOf c Is VB.TextBox) Or (TypeOf c Is VB.ComboBox) Or (TypeOf c Is MSForms.ComboBox) Then '
                If c.Name = "NBox" Then
                    RetString = RetString & c.Tag & "=" & c.Text & ","
                Else
                    RetString = RetString & c.Tag & "='" & c.Text & "',"
                End If
            ElseIf (TypeOf c Is MSForms.CheckBox) Then
                If c.Name = "chkinactive" Then
                    RetString = RetString & c.Tag & "=" & Abs(Not c.Value) & ","
                Else
                    RetString = RetString & c.Tag & "=" & Abs(c.Value) & ","
                End If
            ElseIf (TypeOf c Is ComboList.Usercontrol1) Then
                RetString = RetString & c.Tag & "='" & c.Text & "',"
            End If
        End If
    Next
    
  
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    
    CreateUpdateSQL = RetString
    
End Function


Public Sub EditCustomer(strCustID As String)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    CustID = strCustID
    
    With rs
        .Open "SELECT * FROM ForeignCustomers_Marketing WHERE CustCode+Country='" & CustID & "'", con, adOpenForwardOnly, adLockReadOnly
        
        For Each c In Controls
            If c.Tag <> "" Then
                If TypeOf c Is MSForms.CheckBox Then
                    If c.Name = "chkinactive" Then
                        c.Value = Not .Fields(c.Tag)
                    Else
                        c.Value = .Fields(c.Tag)
                    End If
                ElseIf TypeOf c Is ComboList.Usercontrol1 Then
                    c.Text = .Fields(c.Tag) & ""
                Else
                    c.Text = .Fields(c.Tag) & ""
                End If
            End If
        Next
        
   .Close
       
       
    
End With
    Dim Bnk As BankInfo
   
   With rs
      .Open "Select * From FCustBanks Where CustCode+Country='" & CustID & "'", con, adOpenForwardOnly
      Do Until .EOF
         Set Bnk = New BankInfo
         Bnk.BankID = ![BankID]
         Bnk.BankName = ![BankName] & ""
         Bnk.Country = ![BankCountry] & ""
         Bnk.Phone1 = ![Phone1] & ""
         Bnk.Phone2 = ![Phone2] & ""
         Bnk.Fax1 = ![Fax1] & ""
         Bnk.Fax2 = ![Fax2] & ""
         Bnk.Email1 = ![Email1] & ""
         Bnk.Email2 = ![Email2] & ""
         Bnk.Address1 = ![Address1] & ""
         Bnk.Address2 = ![Address2] & ""
         Bnk.DefaultBank = ![DefaultBank]
         
         CustBanks.add Bnk
         
         .MoveNEXT
      Loop
      .Close
      .Open "SELECT * FROM ForeignCustomers_MarketingShippingAddresses WHERE CustCode+Country='" & CustID & "'", con, adOpenForwardOnly, adLockReadOnly
      LV.ListItems.Clear
      Do Until .EOF
        LV.ListItems.add(, , !ShippingAddress & "").Tag = Val(!EntryID & "")
        .MoveNEXT
      Loop
      .Close
   End With

    
   Set rs = Nothing
   Set Bnk = Nothing
   Me.Show 1
   txtfields(10).Enabled = False
   Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub showBankInfo(Optional Index As Integer = 0)
   If Index = 0 Then
      For i = 0 To txtbank.count - 1
         txtbank(i) = ""
      Next
      BnkIndex = 0
      Exit Sub
   End If

   With CustBanks(Index)
      txtbank(0) = .BankName
      txtbank(1) = .Country
      txtbank(2) = .Phone1
      txtbank(3) = .Phone2
      txtbank(4) = .Fax1
      txtbank(5) = .Fax2
      txtbank(6) = .Email1
      txtbank(7) = .Email2
      txtbank(8) = .Address1
      txtbank(9) = .Address2
      chkDefaultBank = Abs(.DefaultBank)
   End With
   BnkIndex = Index
End Sub


Private Function UpdateBankInfo(Optional Bnk As BankInfo) As Boolean
On Error GoTo err
If Bnk Is Nothing Then
   Set Bnk = New BankInfo
   CustBanks.add Bnk, UCase(txtbank(0)) & "(" & UCase(txtbank(0)) & ")"
   BnkIndex = CustBanks.count
End If

With Bnk
      .BankName = txtbank(0)
      .Country = txtbank(1)
      .Phone1 = txtbank(2)
      .Phone2 = txtbank(3)
      .Fax1 = txtbank(4)
      .Fax2 = txtbank(5)
      .Email1 = txtbank(6)
      .Email2 = txtbank(7)
      .Address1 = txtbank(8)
      .Address2 = txtbank(9)
      
      If chkDefaultBank = vbChecked Then
         For i = 1 To CustBanks.count
            CustBanks(i).DefaultBank = False
         Next
      End If
      
      .DefaultBank = chkDefaultBank = vbChecked
      
End With

Exit Function
err:
'if err.Number =
MsgBox err.Number
Stop
End Function


Private Function SaveBanks(CustomerID As String) As Boolean

   If Len(BanksRemoved) > 0 Then
      con.Execute "Delete From FCustBanks Where Bankid in (" & Left(BanksRemoved, Len(BanksRemoved) - 1) & ")"
   End If


   For i = 1 To CustBanks.count
      With CustBanks.Item(i)
         If .BankID = 0 Then
            con.Execute "INSERT INTO FCustBanks" & _
            "(CustCode,Country,BankName,BankCountry,Phone1,Phone2,Fax1,Fax2,Email1,Email2,Address1,Address2,DefaultBank)" & _
            " Values('" & CustomerID & "','" & cmbCountry.Text & "','" & .BankName & "','" & .Country & "','" & .Phone1 & "','" & .Phone2 & "','" & _
            .Fax1 & "','" & .Fax2 & "','" & .Email1 & "','" & .Email2 & "','" & .Address1 & "','" & _
            .Address2 & "'," & Abs(.DefaultBank) & ")"
         Else
            con.Execute "Update  FCustBanks set bankname='" & .BankName & "',BankCountry='" & .Country & "',Phone1='" & _
            .Phone1 & "',Phone2='" & .Phone2 & "',Fax1='" & .Fax1 & "',Fax2='" & .Fax2 & "',Email1='" & .Email1 & "',Email2='" & _
            .Email2 & "',Address1='" & .Address1 & "',Address2='" & .Address2 & "',Defaultbank=" & Abs(.DefaultBank) & " Where BankID=" & .BankID
         End If
         
      End With
   Next
SaveBanks = True
End Function



Private Function Saved() As Boolean

   Dim SQL As String
    
    If cmbDTFormat.MatchFound = False Then
        MsgBox "What About Selecting A Valid Date Fromat For This Customer ?", vbInformation, "Date Format"
        SSTab1.Tab = 0
        
        cmbDTFormat.DropDown
        cmbDTFormat.SetFocus
        Exit Function
    End If
    
    If CustID = "" Then
    
'        Dim AccNo As String, ParentAccNo As String
'        With con.Execute("Select Accno From Accounts Where Upper(AccTitle)='" & UCase(txtfields(10)) & "'")
'            If .RecordCount > 0 Then
'                ParentAccNo = ![AccNo]
'            Else
'                'ParentAccNo = NewAccRcv(txtfields(10), DtOpenDate, 0, True, FCUSTACCNO, False, True)
'                ParentAccNo = NewAccRcv(txtfields(10), DtOpenDate, 0, True, FCUSTACCNO, False)
'            End If
'            .Close
'        End With
'
'        AccNo = NewAccRcv(txtfields(1) & " (" & cmbCountry.Text & ")", DtOpenDate, Val(txtfields(11).Tag), (chkInactive <> vbChecked), ParentAccNo)
'        If txtfields(0) = "" Then
'            txtfields(0) = AccNo
'        ElseIf txtfields(0) <> AccNo Then
'            MsgBox "Account # """ & txtfields(0) & """ Is Not Available, Account # Will Be Changed To """ & AccNo & """"
'            txtfields(0) = AccNo
'        End If
'        Dim strAccNo As String
'        strAccNo = GetNextAccno(txtfields(10), "15", "15-003")
'        txtfields(0) = strAccNo
        SQL = CreateInsertSQL("ForeignCustomers_Marketing")
    Else
       SQL = CreateUpdateSQL("ForeignCustomers_Marketing")
       SQL = SQL & " Where CustCode+Country='" & CustID & "'"
    End If

    With con
        Call StartTrans(con)
   
        .Execute SQL, a
        Saved = SaveBanks(txtfields(10))
        If CustID = "" Then    'if New Customer Then Added
            Dim strAccTitle As String
            If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
                strAccTitle = txtfields(1) & "(" & txtfields(10) & ")"
            Else
                strAccTitle = txtfields(10)
            End If
            'AccNo = NewAccRcv(txtfields(1), DtOpenDate, Val(txtfields(11).Tag), (chkinactive <> vbChecked), FCUSTACCNO)
            'If txtfields(0) <> AccNo Then MsgBox "Account # """ & txtfields(0) & """ Is Not Available, Account # Will Be Changed To """ & AccNo & """"
'            con.Execute "INSERT INTO Accounts(AccNo,AccTitle,SubAccOf,Type,OpenBal,OpenDate," & _
             "Balance,UserName,ComputerName) Values('" & strAccNo & "','" & strAccTitle & _
             "','15-003','15',0,'" & Date & "',0,'" & CurrentUserName & "','" & strComputerName & "')"
        Else
            'Update AccTitle
            'con.Execute "Update Accounts SET AccTitle='" & txtfields(1) & "' Where AccNo='" & txtfields(0) & "'"
        End If
        
        'Now Save Shipping Addresses.
        If strDel <> "" Then
            strDel = Left(strDel, Len(strDel) - 1)
            con.Execute "DELETE FROM ForeignCustomers_MarketingShippingAddresses WHERE EntryID IN(" & strDel & ")"
        End If
        
        For i = 1 To LV.ListItems.count
            With LV.ListItems(i)
                If Val(.Tag) = 0 Then
                    con.Execute "INSERT INTO ForeignCustomers_MarketingShippingAddresses(CustCode,Country,ShippingAddress) VALUES('" & _
                     txtfields(10).Text & "','" & cmbCountry.Text & "','" & .Text & "')"
                Else
                    con.Execute "UPDATE ForeignCustomers_MarketingShippingAddresses SET ShippingAddress='" & .Text & _
                     "' WHERE EntryID=" & Val(.Tag)
                End If
            End With
        Next
        
        .CommitTrans
   
    End With

End Function


Public Sub AddNew()
    
    If strCompany <> "Dr-Frgz" Then
        Dim strCustCode As String, iCustCode As Integer
        Dim rs As New ADODB.Recordset
        With rs
            .Open "SELECT MAX(CAST(RIGHT(CustCode,4) AS INT)) FROM ForeignCustomers_Marketing WHERE LEN(CustCode)=8 AND LEFT(CustCode,4)='MAK-'", con, adOpenForwardOnly, adLockReadOnly
            iCustCode = Val(.Fields(0).Value & "")
            .Close
        End With
        iCustCode = iCustCode + 1
        strCustCode = "MAK-" & Format(iCustCode, "0000")
        txtfields(10) = strCustCode
        txtfields(10).Locked = False
    End If
    
    Me.Show 1
    
End Sub
