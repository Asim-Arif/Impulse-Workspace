VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmNewComInv 
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12465
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
   Icon            =   "frmNewComInv.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   12465
   WindowState     =   2  'Maximized
   Begin VB.Frame FAll 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9615
      Left            =   45
      TabIndex        =   5
      Top             =   30
      Width           =   12375
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   8745
         TabIndex        =   40
         Top             =   8865
         Width           =   3555
         Begin MSForms.CommandButton cmdCancel 
            Height          =   360
            Left            =   1800
            TabIndex        =   42
            Top             =   195
            Width           =   1665
            Caption         =   "Close            "
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
         Begin MSForms.CommandButton cmdPrint 
            Height          =   360
            Left            =   90
            TabIndex        =   41
            Top             =   195
            Width           =   1665
            Caption         =   "Print Invoice"
            PicturePosition =   327683
            Size            =   "2937;635"
            Accelerator     =   80
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   285
         TabIndex        =   0
         Top             =   1035
         Width           =   1740
         _ExtentX        =   3069
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
         Appearance      =   0
      End
      Begin VB.TextBox txtCurrency 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3810
         MaxLength       =   150
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   1035
         Width           =   1005
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   12375
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Commercial Invoice"
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
            Index           =   8
            Left            =   75
            TabIndex        =   8
            Top             =   165
            Width           =   12195
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Commercial Invoice"
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
            Index           =   9
            Left            =   90
            TabIndex        =   9
            Top             =   180
            Width           =   12180
         End
      End
      Begin MSComCtl2.DTPicker OrderDT 
         Height          =   300
         Left            =   6570
         TabIndex        =   3
         Top             =   1035
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   529
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
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   55508995
         CurrentDate     =   37250
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   7080
         Left            =   135
         TabIndex        =   14
         Top             =   1365
         Width           =   12105
         _ExtentX        =   21352
         _ExtentY        =   12488
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         ShowFocusRect   =   0   'False
         TabCaption(0)   =   "Order Detail"
         TabPicture(0)   =   "frmNewComInv.frx":030A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LV"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Other Info."
         TabPicture(1)   =   "frmNewComInv.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtVessel"
         Tab(1).Control(1)=   "txtAdd2"
         Tab(1).Control(2)=   "txtAdd1"
         Tab(1).Control(3)=   "txtOrders"
         Tab(1).Control(4)=   "txtDis"
         Tab(1).Control(5)=   "txtLoad"
         Tab(1).Control(6)=   "txtFormE"
         Tab(1).Control(7)=   "txtAWBNo"
         Tab(1).Control(8)=   "AWBDT"
         Tab(1).Control(9)=   "Label1(14)"
         Tab(1).Control(10)=   "cmdSwap"
         Tab(1).Control(11)=   "Label1(13)"
         Tab(1).Control(12)=   "Label1(12)"
         Tab(1).Control(13)=   "Label1(6)"
         Tab(1).Control(14)=   "Label1(5)"
         Tab(1).Control(15)=   "Label1(0)"
         Tab(1).Control(16)=   "cmbPayment"
         Tab(1).Control(17)=   "Label1(3)"
         Tab(1).Control(18)=   "cmbSMethod"
         Tab(1).Control(19)=   "Label1(4)"
         Tab(1).Control(20)=   "Label1(10)"
         Tab(1).Control(21)=   "Label1(11)"
         Tab(1).ControlCount=   22
         Begin VB.TextBox txtVessel 
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   -74835
            MaxLength       =   150
            TabIndex        =   38
            Top             =   2895
            Width           =   11715
         End
         Begin VB.TextBox txtAdd2 
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   1725
            Left            =   -68505
            MaxLength       =   150
            MultiLine       =   -1  'True
            TabIndex        =   36
            Top             =   3450
            Width           =   5385
         End
         Begin VB.TextBox txtAdd1 
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   1725
            Left            =   -74835
            MaxLength       =   150
            MultiLine       =   -1  'True
            TabIndex        =   35
            Top             =   3450
            Width           =   5385
         End
         Begin VB.TextBox txtOrders 
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   -74835
            MaxLength       =   150
            TabIndex        =   31
            Top             =   2340
            Width           =   11715
         End
         Begin VB.TextBox txtDis 
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   -74835
            MaxLength       =   150
            TabIndex        =   29
            Top             =   1785
            Width           =   11715
         End
         Begin VB.TextBox txtLoad 
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   -74835
            MaxLength       =   150
            TabIndex        =   27
            Top             =   1230
            Width           =   11715
         End
         Begin VB.TextBox txtFormE 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   -66135
            TabIndex        =   24
            Top             =   675
            Width           =   3000
         End
         Begin VB.TextBox txtAWBNo 
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   -72870
            MaxLength       =   150
            TabIndex        =   18
            Top             =   675
            Width           =   1890
         End
         Begin MSComctlLib.ListView LV 
            Height          =   6600
            Left            =   75
            TabIndex        =   15
            Top             =   390
            Width           =   11940
            _ExtentX        =   21061
            _ExtentY        =   11642
            View            =   3
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            TextBackground  =   -1  'True
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
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S No."
               Object.Width           =   1014
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "PInvoice"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Order #"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Item ID"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Description"
               Object.Width           =   4233
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Unit"
               Object.Width           =   1147
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Price"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "Qty"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Text            =   "Amount"
               Object.Width           =   1852
            EndProperty
         End
         Begin MSComCtl2.DTPicker AWBDT 
            Height          =   300
            Left            =   -70965
            TabIndex        =   19
            Top             =   675
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   529
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
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   55508995
            CurrentDate     =   37250
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Vessel"
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
            Index           =   14
            Left            =   -74835
            TabIndex        =   39
            Top             =   2655
            Width           =   11715
         End
         Begin MSForms.CommandButton cmdSwap 
            Height          =   975
            Left            =   -69405
            TabIndex        =   37
            Top             =   3765
            Width           =   825
            Caption         =   "SWAP"
            Size            =   "1455;1720"
            Picture         =   "frmNewComInv.frx":0342
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"frmNewComInv.frx":065C
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
            Index           =   13
            Left            =   -74835
            TabIndex        =   34
            Top             =   3210
            Width           =   11715
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Form 'E'"
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
            Index           =   12
            Left            =   -66135
            TabIndex        =   33
            Top             =   435
            Width           =   3000
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Orders"
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
            Index           =   6
            Left            =   -74835
            TabIndex        =   32
            Top             =   2100
            Width           =   11715
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Port Of Discharge"
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
            Index           =   5
            Left            =   -74835
            TabIndex        =   30
            Top             =   1545
            Width           =   11715
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Port Of Loading"
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
            Index           =   0
            Left            =   -74835
            TabIndex        =   28
            Top             =   990
            Width           =   11715
         End
         Begin MSForms.ComboBox cmbPayment 
            Height          =   285
            Left            =   -69165
            TabIndex        =   26
            Top             =   675
            Width           =   3000
            VariousPropertyBits=   746604571
            BackColor       =   16777215
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5292;503"
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
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Payment Terms"
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
            Index           =   3
            Left            =   -69165
            TabIndex        =   25
            Top             =   435
            Width           =   3000
         End
         Begin MSForms.ComboBox cmbSMethod 
            Height          =   285
            Left            =   -74835
            TabIndex        =   23
            Top             =   675
            Width           =   1935
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "3413;503"
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
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Shipping Method"
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
            Index           =   4
            Left            =   -74835
            TabIndex        =   22
            Top             =   435
            Width           =   1935
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " AWB No."
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
            Index           =   10
            Left            =   -72870
            TabIndex        =   21
            Top             =   435
            Width           =   1890
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " AWB Date"
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
            Index           =   11
            Left            =   -70965
            TabIndex        =   20
            Top             =   435
            Width           =   1755
         End
      End
      Begin MSForms.ComboBox cmbCustomInvoice 
         Height          =   285
         Left            =   4830
         TabIndex        =   17
         Top             =   1035
         Width           =   1725
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3043;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox lblTotWords 
         Height          =   270
         Left            =   2310
         TabIndex        =   16
         Top             =   8535
         Width           =   5385
         VariousPropertyBits=   746604575
         BackColor       =   15199215
         BorderStyle     =   1
         Size            =   "9499;476"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Amount :"
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
         Height          =   270
         Index           =   1
         Left            =   7680
         TabIndex        =   13
         Top             =   8535
         Width           =   885
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Amount (In Words) :"
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
         Height          =   270
         Index           =   2
         Left            =   495
         TabIndex        =   12
         Top             =   8535
         Width           =   1830
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   8550
         TabIndex        =   11
         Top             =   8535
         Width           =   1500
      End
      Begin MSForms.CheckBox chkPrint 
         Height          =   345
         Left            =   270
         TabIndex        =   10
         Top             =   8880
         Width           =   1545
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   4
         Size            =   "2725;609"
         Value           =   "0"
         Caption         =   "Print On Save"
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbTradeTerms 
         Height          =   285
         Left            =   8130
         TabIndex        =   4
         Top             =   1035
         Width           =   1920
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3387;503"
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
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmNewComInv.frx":06EA
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
         Index           =   7
         Left            =   285
         TabIndex        =   6
         Top             =   795
         Width           =   9765
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   2055
         TabIndex        =   1
         Top             =   1035
         Width           =   1725
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3043;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmNewComInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strProforma As String

Public Sub EditProforma(TempProforma As String)
    
    On Error GoTo err
    strProforma = TempProforma
    Dim rs As New ADODB.Recordset
    Dim Itm As ListItem
    
    With cmbCustomInvoice
        .AddItem strProforma
        .ListIndex = 0
        .Locked = True
    End With
    
    With rs
        .Open "Select * From VComInvoice Where ComInvoice='" & strProforma & "'", con, adOpenForwardOnly, adLockReadOnly
        
        cmbCust.Text = ![CustCode] & ""
        cmbCountry = ![Country] & ""
        cmbCust.Enabled = False
        cmbCountry.Enabled = False
        
        OrderDT = ![DT] & ""
        cmbTradeTerms = ![TradeTerms] & ""
        cmbSMethod = ![ShippingMode] & ""
        txtAWBNo = ![AWBNo] & ""
        AWBDT = ![AWBNoDT] & ""
        cmbPayment = ![PayTerms] & ""
        txtFormE = ![FormE] & ""
        
        txtLoad = ![LoadPort] & ""
        txtDis = ![DisPort] & ""
        
        txtAdd1 = ![Consignee] & ""
        
        txtAdd2 = ![Notify] & ""
        
        txtVessel = ![Vessel] & ""
        
        Dim Str As String
        
        
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set Itm = LV.ListItems.Add(, , LV.ListItems.Count + 1)
            Itm.SubItems(1) = ![PInvoice] & ""
            Itm.SubItems(2) = ![OrderNo] & ""
            Itm.SubItems(3) = ![ITemID] & ""
            Itm.SubItems(4) = ![Description] & ""
            Itm.SubItems(5) = ![Unit] & ""
            Itm.SubItems(6) = ![Price] & ""
            Itm.SubItems(7) = ![Qty] & ""
            Itm.SubItems(8) = Val(![Price] & "") * Val(![Qty] & "")
            
            If Not InStr(Str, ![OrderNo] & "") Then
                Str = Str & ![OrderNo] & ","
            End If
            .MoveNext
            
        Next i
        Str = Left(Str, Len(Str) - 1)
        
        txtOrders = Str
        
        .Close
    End With
    Call CalculateTotals

    'txtQty = 0
    'cmbItems.SetFocus
    
    Set rs = Nothing
    Me.Show
    Exit Sub
err:
    MsgBox err.Description
    
End Sub



Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    
    On Error GoTo err
    If cmbCountry.MatchFound = False Then Exit Sub
    
    cmbCust.Tag = cmbCust.Text & cmbCountry
    txtCurrency = cmbCountry.List(cmbCountry.ListIndex, 1)
    
    If strProforma <> "" Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_FPInvoices"
        
        .Parameters("@Customer").value = cmbCust.Tag
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        Set rs = .Execute
    End With
    
    cmbCustomInvoice.Clear
    With rs
        Do Until .EOF
            cmbCustomInvoice.AddItem .Fields(0) & ""
            
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    
    txtAdd1 = con.Execute("Select Address From ForeignCustomers Where CustCode+Country='" & cmbCust.Tag & "'").Fields(0).value & ""
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub cmbCust_matched()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Country,Curr From ForeignCustomers Where CustCode='" & cmbCust.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            cmbCountry.List(cmbCountry.ListCount - 1, 1) = ![Curr] & ""
            
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbCust_UnMatched()
    'cmbCountry.ClearVals
End Sub

Private Sub cmbCustomInvoice_Change()
    Call cmbCustomInvoice_Click
End Sub

Private Sub cmbCustomInvoice_Click()

    On Error GoTo err
    
    If strProforma <> "" Then Exit Sub
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select * From VrptProformas Where PInvoice='" & cmbCustomInvoice & "'", con, adOpenForwardOnly, adLockReadOnly
        
        OrderDT = ![DT] & ""
        cmbTradeTerms = ![TradeTerms] & ""
        cmbSMethod = ![SMethod] & ""
        txtAWBNo = ![BL] & ""
        AWBDT = ![BLDT] & ""
        cmbPayment = ![PaymentTerms] & ""
        txtFormE = ![FormENo] & ""
        
        txtLoad = ![LoadPort] & " " & ![LoadCity] & " " & ![LoadCountry]
        txtDis = ![DisPort] & " " & ![DisCity] & " " & ![DisCountry]
        
        txtAdd2 = ![Address1] & ""
        Dim Str As String
        Dim Itm As ListItem
        
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set Itm = LV.ListItems.Add(, , LV.ListItems.Count + 1)
            Itm.SubItems(1) = ![PInvoice] & ""
            Itm.SubItems(2) = ![OrderNo] & ""
            Itm.SubItems(3) = ![ITemID] & ""
            Itm.SubItems(4) = ![Description] & ""
            Itm.SubItems(5) = ![Unit] & ""
            Itm.SubItems(6) = ![Price] & ""
            Itm.SubItems(7) = ![Qty] & ""
            Itm.SubItems(8) = Val(![Price] & "") * Val(![Qty] & "")
            
            If Not InStr(Str, ![OrderNo] & "") Then
                Str = Str & ![OrderNo] & ","
            End If
            .MoveNext
            
        Next i
        Str = Left(Str, Len(Str) - 1)
        
        txtOrders = Str
        
        .Close
    End With
    
    Set rs = Nothing
    
    Call CalculateTotals
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbTradeTerms_Change()
    Call cmbTradeTerms_Click
End Sub

Private Sub cmbTradeTerms_Click()
    Exit Sub
    If cmbTradeTerms.MatchFound Then
        cmbCatID.Locked = False
        cmbItems.Locked = False
        cmbDesc.Locked = False
        If cmbItems.MatchFound Then
            txtPrice = cmbItems.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 3)
        End If
    Else
        cmbCatID.Locked = True
        cmbItems.Locked = True
        cmbDesc.Locked = True
    End If
End Sub


Private Sub cmdSwap_Click()

    'Dim Str As String
    
    'Str = txtAdd1
    'txtAdd1 = txtAdd2
    'txtAdd2 = Str
    
    txtAdd1 = txtAdd1 & txtAdd2
    txtAdd2 = Replace(txtAdd1, txtAdd2, "")
    txtAdd1 = Replace(txtAdd1, txtAdd2, "")
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSClear_Click()
    If Saved Then
        Unload Me
        Me.Show
    End If
End Sub

Private Sub cmdSClose_Click()
    If Saved Then
        Unload Me
    End If
End Sub


Private Sub Form_Load()
   
    On Error GoTo err
    
    SSTab1.Tab = 0
    
    Call SaveLV(LV)
    
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where custCode<>'Hico'")
    OrderDT = Date
    Me.KeyPreview = True
    cmbCust.ListHeight = 2500
    Call FillCmbs
    Exit Sub
err:
    MsgBox err.Description
   
End Sub

Private Sub Form_Resize()
On Error Resume Next
   FAll.Move ((Me.ScaleWidth - FAll.Width) / 2), ((Me.ScaleHeight - FAll.Height) / 2)
End Sub

Private Sub CalculateTotals()

    With LV
        For i = 1 To .ListItems.Count
            'dQty = dQty + Val(.TextMatrix(i, 4))
            DTotal = DTotal + Val(.ListItems(i).SubItems(8) & "")
        Next i
    End With
    'txttQty = dQty
    lblTotal = DTotal
    lblTotWords = ConvertInWords(CCur(DTotal), txtCurrency.Text)
    
End Sub
Private Function Saved() As Boolean
    
    On Error GoTo err
    
    
    Saved = False
    
    Call StartTrans(con)
    '----------------------------------------------------------------
    If strProforma <> "" Then
        con.Execute "Delete from CommercialInvoice Where ComInvoice='" & _
         strProforma & "'"
    End If
    
    
    con.Execute "Insert Into CommercialInvoice (ComInvoice,DT," & _
     "ShippingMode,AWBNo,AWBNoDT,FormE,TradeTerms,Consignee," & _
     "Notify,LoadPort,DisPort,Vessel,PayTerms) Values('" & _
     cmbCustomInvoice & "','" & OrderDT & "','" & cmbSMethod & _
     "','" & txtAWBNo & "','" & AWBDT & "','" & txtFormE & _
     "','" & cmbTradeTerms & "','" & txtAdd1 & "','" & txtAdd2 & _
     "','" & txtLoad & "','" & txtDis & "','" & txtVessel & _
     "','" & cmbPayment & "')"
    
    con.CommitTrans
    
    Saved = True
    Exit Function

err:
    MsgBox err.Description
    
End Function
Private Sub FillCmbs()

    With cmbTradeTerms
        .Clear
        .AddItem "FOB"
        .AddItem "CIF (Air)"
        .AddItem "CIF (Sea)"
        .AddItem "C&F (Air)"
        .AddItem "C&F (Sea)"
        .ListIndex = 0
    End With
    
    
    With cmbPayment
        .AddItem "CAD"
        .AddItem "30 DAYS AT SIGHT"
        .AddItem "90 DAYS AT SIGHT"
        .AddItem "60 DAYS AT SIGHT"
        .AddItem "Letter Of Credit"
        .AddItem "Advance Cheque"
        .AddItem "By Cheque"
    End With
    
    With cmbSMethod
        .AddItem "By Air"
        .AddItem "By Sea"
    End With
    
End Sub

 
