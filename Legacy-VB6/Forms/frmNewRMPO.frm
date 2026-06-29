VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmNewRMPO 
   ClientHeight    =   9645
   ClientLeft      =   885
   ClientTop       =   -705
   ClientWidth     =   12045
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   12045
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkShowPicture 
      Caption         =   "Show Picture"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   86
      Top             =   8970
      Width           =   1560
   End
   Begin ComboList.Usercontrol1 cmbCustomerPONo 
      Height          =   285
      Left            =   7440
      TabIndex        =   79
      Top             =   1110
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextBackColor   =   12640511
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin VB.ComboBox cmbProcessOrder 
      Height          =   315
      ItemData        =   "frmNewRMPO.frx":0000
      Left            =   10110
      List            =   "frmNewRMPO.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   59
      Top             =   1110
      Width           =   1920
   End
   Begin VB.CheckBox chkSteelProvided 
      Caption         =   "Steel Provided"
      Height          =   300
      Left            =   120
      TabIndex        =   57
      Top             =   7500
      Width           =   1410
   End
   Begin VB.TextBox txtSteelProvided 
      Appearance      =   0  'Flat
      DataField       =   "RptTime"
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   1575
      Locked          =   -1  'True
      MaxLength       =   150
      TabIndex        =   56
      Top             =   7515
      Width           =   10425
   End
   Begin VB.TextBox txtDrawingProvided 
      Appearance      =   0  'Flat
      DataField       =   "RptTime"
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      MaxLength       =   150
      TabIndex        =   54
      Top             =   8700
      Width           =   2160
   End
   Begin VB.CheckBox chkDrawing 
      Caption         =   "Drawing Provided"
      Height          =   300
      Left            =   6255
      TabIndex        =   53
      Top             =   8445
      Width           =   1950
   End
   Begin VB.TextBox txtSampleProvided 
      Appearance      =   0  'Flat
      DataField       =   "RptTime"
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   6240
      Locked          =   -1  'True
      MaxLength       =   150
      TabIndex        =   52
      Top             =   8085
      Width           =   2160
   End
   Begin VB.CheckBox chkSample 
      Caption         =   "Sample Provided"
      Height          =   300
      Left            =   6255
      TabIndex        =   51
      Top             =   7830
      Width           =   1950
   End
   Begin ComboList.Usercontrol1 cmbVend 
      Height          =   285
      Left            =   3195
      TabIndex        =   2
      Top             =   1980
      Width           =   5145
      _ExtentX        =   9075
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   5070
      Left            =   120
      TabIndex        =   25
      Top             =   2370
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   8943
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Order Materials"
      TabPicture(0)   =   "frmNewRMPO.frx":002D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblUnit"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblAmt"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtLastRate"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdAddItem"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtPrice"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtQty"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtEditQty"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "LV"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmbForMaterial"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmbItem"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtEdit"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "chkShowCCItems"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "CD1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "FraRemarks"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "PicProduct"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtStock"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Other Details"
      TabPicture(1)   =   "frmNewRMPO.frx":0049
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Sample Pics"
      TabPicture(2)   =   "frmNewRMPO.frx":0065
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "PicContainer"
      Tab(2).Control(1)=   "cmdAddPic"
      Tab(2).Control(2)=   "CDPics"
      Tab(2).ControlCount=   3
      Begin VB.TextBox txtStock 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   5250
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   88
         Top             =   690
         Width           =   855
      End
      Begin VB.PictureBox PicProduct 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2865
         Left            =   9030
         ScaleHeight     =   2835
         ScaleWidth      =   2760
         TabIndex        =   87
         Top             =   1290
         Visible         =   0   'False
         Width           =   2790
         Begin VB.Image ImgProduct 
            Height          =   2835
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2760
         End
      End
      Begin VB.Frame FraRemarks 
         BackColor       =   &H00E7EBEF&
         Caption         =   "Remarks :"
         Height          =   2025
         Left            =   6495
         TabIndex        =   82
         Top             =   1860
         Visible         =   0   'False
         Width           =   4410
         Begin VB.CommandButton cmdCancelPackDetails 
            Caption         =   "Cancel"
            Height          =   300
            Left            =   3540
            TabIndex        =   85
            Top             =   1650
            Width           =   795
         End
         Begin VB.CommandButton cmdOKPackingDetails 
            Caption         =   "OK"
            Height          =   300
            Left            =   2700
            TabIndex        =   84
            Top             =   1650
            Width           =   795
         End
         Begin VB.TextBox txtItemRemarks 
            Appearance      =   0  'Flat
            Height          =   1395
            Left            =   30
            MultiLine       =   -1  'True
            TabIndex        =   83
            Top             =   225
            Width           =   4305
         End
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   1770
         Top             =   3000
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "PDF Files Only|*.pdf"
      End
      Begin MSComDlg.CommonDialog CDPics 
         Left            =   -65730
         Top             =   2265
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DialogTitle     =   "Please Select Picture"
         Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
      End
      Begin VB.CommandButton cmdAddPic 
         Caption         =   "Add Picture"
         Height          =   375
         Left            =   -66060
         TabIndex        =   76
         Top             =   870
         Width           =   2070
      End
      Begin VB.PictureBox PicContainer 
         Height          =   3270
         Left            =   -74190
         ScaleHeight     =   3210
         ScaleWidth      =   10155
         TabIndex        =   74
         Top             =   1290
         Width           =   10215
         Begin VB.PictureBox PicInner 
            Height          =   2925
            Left            =   0
            ScaleHeight     =   2865
            ScaleWidth      =   10095
            TabIndex        =   77
            Top             =   0
            Width           =   10155
            Begin VB.Image ImgPics 
               Height          =   2355
               Index           =   0
               Left            =   75
               Stretch         =   -1  'True
               Top             =   465
               Width           =   2445
            End
            Begin VB.Label lblImg 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Pic Description"
               Height          =   195
               Index           =   0
               Left            =   75
               TabIndex        =   78
               Top             =   120
               Width           =   2445
            End
         End
         Begin VB.HScrollBar HScroll1 
            Height          =   270
            Left            =   0
            TabIndex        =   75
            Top             =   2940
            Width           =   10155
         End
      End
      Begin VB.CheckBox chkShowCCItems 
         BackColor       =   &H00C5D1DC&
         Caption         =   "Show Company Catalog Items"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   2190
         TabIndex        =   73
         Top             =   435
         Width           =   2955
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2280
         TabIndex        =   55
         Top             =   1950
         Visible         =   0   'False
         Width           =   1455
      End
      Begin ComboList.Usercontrol1 cmbItem 
         Height          =   285
         Left            =   60
         TabIndex        =   26
         Top             =   690
         Width           =   5175
         _ExtentX        =   9128
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
      Begin ComboList.Usercontrol1 cmbForMaterial 
         Height          =   285
         Left            =   1260
         TabIndex        =   27
         Top             =   990
         Width           =   10560
         _ExtentX        =   18627
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
      Begin MSComctlLib.ListView LV 
         Height          =   3645
         Left            =   60
         TabIndex        =   33
         Top             =   1290
         Width           =   11760
         _ExtentX        =   20743
         _ExtentY        =   6429
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Material ID"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Material Description"
            Object.Width           =   6703
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Material Type"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Price/Unit"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Order Qty"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Amount"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "PDF"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Remarks"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame1 
         Height          =   4215
         Left            =   -74250
         TabIndex        =   38
         Top             =   405
         Width           =   10470
         Begin VB.TextBox txtRef 
            Height          =   330
            Left            =   180
            TabIndex        =   43
            Top             =   495
            Width           =   2190
         End
         Begin VB.TextBox txtPayments 
            Height          =   330
            Left            =   2400
            TabIndex        =   42
            Top             =   495
            Width           =   4350
         End
         Begin VB.TextBox txtShipment 
            Height          =   330
            Left            =   6780
            TabIndex        =   41
            Top             =   495
            Width           =   2430
         End
         Begin VB.TextBox txtInstructions 
            Height          =   1680
            Left            =   180
            MultiLine       =   -1  'True
            TabIndex        =   40
            Top             =   1185
            Width           =   10125
         End
         Begin VB.TextBox txtNotes 
            Height          =   915
            Left            =   180
            MultiLine       =   -1  'True
            TabIndex        =   39
            Top             =   3180
            Width           =   10125
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ref. :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   7
            Left            =   195
            TabIndex        =   50
            Top             =   285
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Payments :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   2385
            TabIndex        =   49
            Top             =   270
            Width           =   945
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Shipment :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   6795
            TabIndex        =   48
            Top             =   285
            Width           =   900
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Instructions :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   195
            TabIndex        =   47
            Top             =   975
            Width           =   1125
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Notes :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   195
            TabIndex        =   46
            Top             =   2955
            Width           =   570
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Currency :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   9225
            TabIndex        =   45
            Top             =   270
            Width           =   855
         End
         Begin MSForms.ComboBox cmbCurrency 
            Height          =   330
            Left            =   9240
            TabIndex        =   44
            Top             =   495
            Width           =   1080
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "1905;582"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.TextBox txtEditQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8745
         TabIndex        =   32
         Top             =   1815
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   8610
         MaxLength       =   150
         TabIndex        =   31
         Top             =   690
         Width           =   915
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   7005
         MaxLength       =   150
         TabIndex        =   30
         Top             =   690
         Width           =   855
      End
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "&Add To List"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   10680
         Picture         =   "frmNewRMPO.frx":0081
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Add This Entry To Voucher"
         Top             =   390
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.TextBox txtLastRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   28
         Top             =   690
         Width           =   870
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmNewRMPO.frx":03C3
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
         Index           =   4
         Left            =   60
         TabIndex        =   37
         Top             =   405
         Width           =   10590
      End
      Begin VB.Label lblAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9555
         TabIndex        =   36
         Top             =   690
         Width           =   1095
      End
      Begin VB.Label lblUnit 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " / Unit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   7875
         TabIndex        =   35
         Top             =   690
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "For Material :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   34
         Top             =   1035
         Width           =   1110
      End
   End
   Begin VB.ComboBox cmbPOType 
      Height          =   315
      ItemData        =   "frmNewRMPO.frx":0480
      Left            =   135
      List            =   "frmNewRMPO.frx":0490
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   1110
      Width           =   2235
   End
   Begin VB.TextBox txtRunningPONo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataField       =   "RptTime"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   11250
      Locked          =   -1  'True
      MaxLength       =   150
      TabIndex        =   22
      Top             =   1980
      Width           =   765
   End
   Begin VB.CheckBox chkUrgent 
      Caption         =   "Very Urgent"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1905
      TabIndex        =   11
      Top             =   8445
      Width           =   1515
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      DataField       =   "RptTime"
      ForeColor       =   &H00400000&
      Height          =   930
      Left            =   3495
      MaxLength       =   1000
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   8070
      Width           =   2715
   End
   Begin VB.TextBox txtDept 
      Appearance      =   0  'Flat
      DataField       =   "RptTime"
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   120
      MaxLength       =   150
      TabIndex        =   12
      Top             =   8100
      Width           =   1815
   End
   Begin VB.TextBox txtPayTerms 
      Appearance      =   0  'Flat
      DataField       =   "RptTime"
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   9660
      MaxLength       =   150
      TabIndex        =   4
      Top             =   1980
      Width           =   1560
   End
   Begin VB.TextBox txtAttn 
      Appearance      =   0  'Flat
      DataField       =   "RptTime"
      ForeColor       =   &H00400000&
      Height          =   285
      Left            =   8355
      MaxLength       =   150
      TabIndex        =   3
      Top             =   1980
      Width           =   1275
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "Print When Saved"
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
      Height          =   240
      Left            =   120
      TabIndex        =   13
      Top             =   8685
      Width           =   2625
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1080
      Left            =   8475
      TabIndex        =   17
      Top             =   7830
      Width           =   3570
      Begin MSForms.CommandButton cmdNewMaterial 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   1815
         TabIndex        =   7
         Top             =   225
         Width           =   1665
         Caption         =   "New Material"
         PicturePosition =   327683
         Size            =   "2937;661"
         Accelerator     =   77
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
         Left            =   135
         TabIndex        =   5
         Top             =   645
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
         CausesValidation=   0   'False
         Height          =   360
         Left            =   1830
         TabIndex        =   8
         Top             =   645
         Width           =   1665
         Caption         =   "Close"
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
         Left            =   135
         TabIndex        =   6
         Top             =   240
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
      Begin VB.Label lbltotalBalls 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0 "
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
         Height          =   240
         Left            =   1665
         TabIndex        =   19
         Top             =   255
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Balls Issued :"
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
         Left            =   60
         TabIndex        =   18
         Top             =   285
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin MSComCtl2.DTPicker dtOrdDate 
      Height          =   285
      Left            =   1620
      TabIndex        =   1
      Top             =   1980
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   503
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
      Format          =   159580163
      CurrentDate     =   37924
   End
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
      Height          =   615
      Left            =   45
      TabIndex        =   14
      Top             =   105
      Width           =   12000
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Order"
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
         Left            =   4545
         TabIndex        =   15
         Top             =   120
         Width           =   2490
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Purchase Order"
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
         Left            =   4560
         TabIndex        =   16
         Top             =   135
         Width           =   2490
      End
   End
   Begin MSComCtl2.DTPicker DTDelivery 
      Height          =   285
      Left            =   1950
      TabIndex        =   10
      Top             =   8100
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   503
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
      Format          =   159580163
      CurrentDate     =   37924
   End
   Begin VB.Frame FraLocations 
      Height          =   3495
      Left            =   1590
      TabIndex        =   63
      Top             =   3495
      Visible         =   0   'False
      Width           =   9390
      Begin VB.TextBox txtInHand 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7515
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   510
         Width           =   960
      End
      Begin VB.TextBox txtRMQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8460
         TabIndex        =   68
         Top             =   510
         Width           =   825
      End
      Begin VB.TextBox txtLocsTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8460
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   3030
         Width           =   825
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   330
         Left            =   105
         TabIndex        =   66
         Top             =   3075
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelLocs 
         Caption         =   "Cancel"
         Height          =   330
         Left            =   1380
         TabIndex        =   65
         Top             =   3075
         Width           =   1215
      End
      Begin MSComctlLib.ListView LVLocsAll 
         Height          =   1800
         Left            =   1215
         TabIndex        =   64
         Top             =   960
         Visible         =   0   'False
         Width           =   6990
         _ExtentX        =   12330
         _ExtentY        =   3175
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Location"
            Object.Width           =   13123
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Qty"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComctlLib.ListView LVLocs 
         Height          =   2250
         Left            =   90
         TabIndex        =   70
         Top             =   795
         Width           =   9195
         _ExtentX        =   16219
         _ExtentY        =   3969
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Location"
            Object.Width           =   13123
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Qty"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmNewRMPO.frx":04C2
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
         Left            =   90
         TabIndex        =   72
         Top             =   240
         Width           =   9195
      End
      Begin MSForms.ComboBox cmbLocation 
         Height          =   285
         Left            =   90
         TabIndex        =   71
         Top             =   510
         Width           =   7440
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "13123;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "This Order Collection By :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   11
      Left            =   5595
      TabIndex        =   81
      Top             =   9210
      Width           =   2085
   End
   Begin MSForms.ComboBox cmbCollectionBy 
      Height          =   390
      Left            =   7770
      TabIndex        =   80
      Top             =   9120
      Width           =   4275
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "7541;688"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Customer Order"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   10
      Left            =   7425
      TabIndex        =   62
      Top             =   885
      Width           =   1350
   End
   Begin MSForms.ComboBox cmbSteel 
      Height          =   315
      Left            =   2385
      TabIndex        =   61
      Tag             =   "SteelUsed"
      Top             =   1110
      Width           =   5040
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "8890;556"
      ListWidth       =   8819
      MatchEntry      =   1
      ListStyle       =   1
      ShowDropButtonWhen=   1
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Steel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   9
      Left            =   2430
      TabIndex        =   60
      Top             =   870
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   8
      Left            =   10065
      TabIndex        =   58
      Top             =   870
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "PO Type"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   23
      Top             =   870
      Width           =   690
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H00C5D1DC&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Department                  Delivery Date           Ramarks"
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
      Left            =   120
      TabIndex        =   21
      Top             =   7815
      Width           =   6090
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H00C5D1DC&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmNewRMPO.frx":0579
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
      Left            =   120
      TabIndex        =   20
      Top             =   1710
      Width           =   11895
   End
   Begin MSForms.TextBox txtPOrdNo 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   1980
      Width           =   1485
      VariousPropertyBits=   679495711
      Size            =   "2619;503"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmNewRMPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VendCapacity As Long
Dim IsItemsFirst As Boolean
Dim OrderNo As String
Dim DelEntryList As String
Dim iColNo As Integer
Dim iImgWidht As Integer
Dim bSample As Boolean

Private Sub chkOrderNo_Click()
'   Call cmbEntryID_click
   
      For i = 1 To LV.ListItems.count
         With LV.ListItems(i)
            .SubItems(1) = IIf(chkOrderNo = vbChecked, Replace(.ListSubItems(1).key, "'", ""), .ListSubItems(1).Tag)
         End With
      Next
   
End Sub

Private Sub cmbEntryID_Change()
'    Call cmbEntryID_click
End Sub

Private Sub cmbCatagory_matched()

    cmbItem.ListHeight = 2600
    cmbItem.ClearVals
    Dim Cond As String
    
    If cmbCatagory.ID = "0" Then
        Cond = " WHERE CATID LIKE '" & cmbMain.List(cmbMain.ListIndex, 1) & "%' AND InActive=0"
    Else
        Cond = " WHERE CATID='" & cmbCatagory.ID & "' AND InActive=0"
    End If
    
    cmbItem.AddVals con, "MaterialName + '     {' + RMID + '}'", "Material", "RMID", Cond
    
    cmbForMaterial.ClearVals
    cmbForMaterial.AddVals con, "MaterialName + ' { ' + RMID + '}'", "Material", "RMID", Cond
    
End Sub

Private Sub cmbGatePass_matched()
    Call SetValues
End Sub

Private Sub chkDrawing_Click()
    If chkDrawing.Value = vbChecked Then
        txtDrawingProvided = ""
        txtDrawingProvided.Locked = False
    Else
        txtDrawingProvided = ""
        txtDrawingProvided.Locked = True
    End If
End Sub

Private Sub chkSample_Click()
    If chkSample.Value = vbChecked Then
        txtSampleProvided = ""
        txtSampleProvided.Locked = False
    Else
        txtSampleProvided = ""
        txtSampleProvided.Locked = True
    End If
End Sub

Private Sub chkShowCCItems_Click()
    If chkShowCCItems.Value = vbChecked Then
        cmbItem.ClearVals
        cmbItem.AddVals con, "ItemName + '     {' + ItemID + '}'", "Items", "ItemID", " WHERE InActive=0 AND Sample=1 ORDER BY ItemName"
    Else
        Call cmbPOType_Click
    End If
End Sub

Private Sub chkSteelProvided_Click()
    If chkSteelProvided.Value = vbChecked Then
        txtSteelProvided = ""
        txtSteelProvided.Locked = False
    Else
        txtSteelProvided = ""
        txtSteelProvided.Locked = True
    End If
End Sub

Private Sub cmbCustomerPONo_matched()
    cmbItem.ClearVals
    cmbItem.AddVals con, "ItemName + '     {' + ItemID + '}'", "Items", "ItemID", " WHERE ItemID IN(SELECT CompItemCode FROM FOrderItems WHERE OrderNo='" & cmbCustomerPONo.ID & "') ORDER BY ItemName"
End Sub

Private Sub cmbItem_matched()
    If Not cmbItem.MatchFound Then Exit Sub
    If cmbReturned = "Yes" Then
        cmbGatePass.ClearVals
        Call cmbGatePass.AddVals(con, "DISTINCT GPN_ID", "VVendGatePassDetail", "GPN_ID", " WHERE VendID ='" & cmbVend.Text & "' And mName ='" & cmbItem.ID & "' AND RcvdQty > 0 And OrderNo is Null Order By GPN_ID")
    End If
    Call SetValues
End Sub

Private Sub SetValues()

    Dim Rec As New ADODB.Recordset
    With Rec
        .Open "SELECT Unit,AnnealingStock,MachiningStock,QtyInStock,Rate,MaleGrindingStock FROM RM WHERE RMID1='" & cmbItem.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            lblUnit = " / " & !Unit & ""
            lblUnit.Tag = !Unit & ""
            txtPrice = Val(!Rate & "")
            Dim lStock As Long
            'lStock = GetSingleLongValue("SUM(NetQty)", "VRMStatusStorewise", " WHERE RMID1='" & cmbItem.ID & "'")
            lStock = GetSingleLongValue("QtyInStock", "VRM", " WHERE RMID1='" & cmbItem.ID & "'")
            If cmbPOType.ListIndex = 1 Then 'Annealing... but Issuance from Stock
                cmbItem.Tag = lStock - Val(!AnnealingStock & "") - Val(!MachiningStock & "") - Val(!MaleGrindingStock & "")
            ElseIf cmbPOType.ListIndex = 2 Then
                cmbItem.Tag = Val(!AnnealingStock & "")
            ElseIf cmbPOType.ListIndex = 3 Then         'Male Grinding
                cmbItem.Tag = Val(!MachiningStock & "")
            End If
            txtStock = lStock
        End If
        .Close
        
        '.Open "Select Top 1 Rate From VVendOrders Where AccTitle ='" & cmbVend.Text & "' AND RMID='" & cmbReturned.Tag & "' Order By EntryID Desc", con, adOpenForwardOnly, adLockReadOnly
        '.Open "SELECT Top 1 Rate FROM VVendOrders WHERE VendId ='" & cmbVend.id & "' AND RMID1='" & cmbItem.id & "' Order By EntryID Desc", con, adOpenForwardOnly, adLockReadOnly
        .Open "SELECT TOP 1 Rate FROM VVendOrders WHERE RMID1='" & cmbItem.ID & "' Order By EntryID Desc", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            txtLastRate = "N.A"
        Else
            txtLastRate = Val(!Rate & "")
        End If
        .Close
        Dim lRMID As Long
        lRMID = GetSingleLongValue("RMID", "RM", " WHERE RMID1='" & cmbItem.ID & "'")
        .Open "SELECT Rate FROM VenderAssItems WHERE VendID='" & cmbVend.ID & "' AND RMID=" & lRMID
        If .EOF = False Then txtPrice = Val(!Rate & "")
        .Close
        
    End With
    
End Sub

Private Function MatType(PrevMatType As Integer) As String

    Select Case PrevMatType
        Case 0
            MatType = "Raw Material"
        Case 1
            MatType = "Consumable Material"
        Case 2
            MatType = "Accessories"
    End Select

End Function

Private Sub cmbReturned_Click()
    If cmbReturned.Text = "No" Then
        cmbGatePass.ClearVals
        cmbGatePass.Enabled = False
    Else
        cmbGatePass.Enabled = True
    End If
    Call cmbVend_matched
End Sub

Private Sub cmbLocation_Change()
    Call cmbLocation_Click
End Sub

Private Sub cmbLocation_Click()
    If cmbLocation.MatchFound = False Then
        txtInHand = ""
    Else
        txtInHand.Text = Val(cmbLocation.List(cmbLocation.ListIndex, 2))
    End If
End Sub

Private Sub cmbPOType_Change()
    Call cmbPOType_Click
End Sub

Private Sub cmbPOType_Click()

    On Error GoTo err
    
    cmbItem.ClearVals
    Dim strUserCondition As String
    If CurrentUserName = "Administrator" Then
        strUserCondition = ""
    Else
        strUserCondition = " AND GroupID IN(SELECT RMGroupID FROM Users_RMGroups WHERE UserID=" & UserID & ")"
    End If
    
    If cmbPOType.ListIndex > 0 Then
        If cmbPOType.ListIndex = 3 Then         'Male Grinding
            cmbItem.AddVals con, "'{'+RMID1+'} '+RMName + ' {' + ISNULL(CAST(RMSize AS VARCHAR(50)),'') + '}'", "RM", "RMID1", " WHERE GroupID IN(SELECT Group_ID FROM RMGroupIDsForForging) AND InActive=0 AND RequiresMaleGrinding=1 " & strUserCondition & " ORDER BY RMName"
        Else
            cmbItem.AddVals con, "'{'+RMID1+'} '+RMName + ' {' + ISNULL(CAST(RMSize AS VARCHAR(50)),'') + '}'", "RM", "RMID1", " WHERE GroupID IN(SELECT Group_ID FROM RMGroupIDsForForging) AND InActive=0 " & strUserCondition & " ORDER BY RMName"
        End If
        LV.Top = 1290
        LV.Height = 3345
        Label2(1).Visible = True
        cmbForMaterial.Visible = True
    Else
        cmbItem.AddVals con, "'{'+RMID1+'} '+RMName + ' {' + ISNULL(CAST(RMSize AS VARCHAR(50)),'') + '}'", "RM", "RMID1", " WHERE InActive=0 " & strUserCondition & " ORDER BY RMName"
        LV.Top = 1005
        LV.Height = 3630
        Label2(1).Visible = False
        cmbForMaterial.Visible = False
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbVend_matched()

    If Not cmbVend.MatchFound Then Exit Sub
    
'    Dim rs As New ADODB.Recordset
'    With rs
'        .Open "Select ImportVender From VVenders Where VendId  =" & Val(cmbVend.ID), con, adOpenForwardOnly, adLockReadOnly
'        If !ImportVender And .RecordCount > 0 Then
'            cmbItem.AddVals con, "RMName + '     {' + RMID1 + '}'", "RM", "RMID1", " WHERE Imported = 1 AND InActive=0 Order By MaterialName"
'        Else
'            cmbItem.AddVals con, "RMName + '     {' + RMID1 + '}'", "RM", "RMID1", " WHERE Imported = 1 AND InActive=0 Order By RMName"
'        End If
'    End With
'    If cmbVend.ID = 281 Then
'        cmbItem.AddVals con, "RMName + '     {' + RMID1 + '}'", "RM", "RMID1", " WHERE InActive=0 Order By RMName"
'    Else
'        cmbItem.AddVals con, "RMName + '     {' + RMID1 + '}'", "RM", "RMID1", " WHERE InActive=0 AND RMID IN(SELECT RMID FROM VenderAssItems WHERE VendID=" & cmbVend.ID & ") Order By RMName"
'    End If
    
    If Left(cmbVend.ID, 10) = "24-004-001" Then
        SSTab1.TabVisible(1) = True
    Else
        SSTab1.TabVisible(1) = False
    End If
    
End Sub



Private Sub cmbVend_UnMatched()
    '    cmbEntryID.ClearVals
    LV.ListItems.Clear
End Sub

Private Sub cmbEntryID_matched()

    On Error GoTo err
    Dim LastGroup As String
    Dim rs As New ADODB.Recordset
    With rs
        
        .Open "Select * From VQuotsToOrder Where RefID=" & cmbEntryID.ID & " and VendID=" & cmbVend.ID & " and QtyToOrder > 0", con, adOpenForwardOnly, adLockReadOnly
        
        Dim ITM As ListItem
        LV.ListItems.Clear
        
        Do Until .EOF
            
            Set ITM = LV.ListItems.add(, , ![RMID] & "")
            ITM.Tag = ![RMID]
            
            'Display the Customer Order No if chkOrderNo is checked else show the Company Order No
            
            ITM.ListSubItems.add , , ![MaterialName]
            
            ITM.ListSubItems.add , , IIf(Val(![MaterialType]) = 0, "Raw Material", IIf(Val(![MaterialType]) = 1, "Consumable Materials", IIf(Val(![MaterialType]) = 2, "Accessories", "General Items")))
            ITM.ListSubItems.add , , Val(![QtyToOrder]) & " " & ![Unit]
            ITM.ListSubItems.add(, , ![Rate] & " /" & ![Unit]).Tag = ![Unit] & ""
            
            'Set Tool Tips
            
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ToolTipText = ITM.ListSubItems(i).Text
            Next
            
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
    
err:
   MsgBox err.Description
   
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo err
    
    
    If Val(txtIssQty) > Val(txtQty) Then
        MsgBox "invalid quantity", vbInformation
        Exit Sub
    ElseIf Not cmbItems.MatchFound Then
        MsgBox "Select an Item first", vbInformation
        Exit Sub
    End If
    
    Dim ITM As ListItem
    
    Set ITM = LV.ListItems.add(, cmbVend.ID & ";" & cmbOrders.Text & ";" & cmbItems.Text, cmbVend.Text & " (" & cmbEntryID.Text & ")")
    ITM.ListSubItems.add , , cmbOrders.Text
    ITM.ListSubItems.add , , cmbItems 'Format(dtOrdDate.value, "dd-MMM-yyyy")
    'ITM.ListSubItems.Add(, , cmbProcess.Text).Tag = cmbProcess.List(cmbProcess.ListIndex, 1)
    ITM.ListSubItems.add(, , cmbVend.Text).Tag = cmbVend.ID
    ITM.ListSubItems.add , , txtQty1
    ITM.ListSubItems.add , , txtIssQty
    'itm.ListSubItems.Add , , txtPOrdNo
           
    cmbVend.Enabled = LV.ListItems.count = 0
    txtQty1 = ""
    txtQty = ""
    txtIssQty = ""
    Exit Sub
err:
    If err.Number = 35602 Then
        MsgBox "This Item is Already Issued", vbInformation
    Else
        MsgBox err.Description
    End If

End Sub

Private Sub cmdAddItem_Click()
    
    If Not cmbItem.MatchFound Then
        MsgBox "Please Select The Material Which You Want To Order !", vbInformation
        cmbItem.SetFocus
        Exit Sub
    ElseIf Val(txtQty) = 0 Then
        MsgBox "Quantity Could Not Be Zero !", vbInformation
        txtQty.SetFocus
        Exit Sub
    ElseIf Val(txtPrice) = 0 Then
        If MsgBox("Do You Want To Add Zero Rate?", vbQuestion + vbYesNo) = vbNo Then
            'MsgBox "Purchase Rate Could No Be Zero !", vbInformation
            txtPrice.SetFocus
            Exit Sub
        End If
    End If
     
    For i = 1 To LV.ListItems.count
        If cmbItem.ID = LV.ListItems(i) Then
            MsgBox "This Material Already Added to List !", vbInformation
            cmbItem.SetFocus
            Exit Sub
        End If
    Next
    
    If cmbPOType.ListIndex > 0 Then
        If Val(txtQty) > Val(cmbItem.Tag) Then
            MsgBox "Not Enough Stock for this Process.", vbInformation
            Exit Sub
        End If
        If cmbForMaterial.MatchFound = False Then
            MsgBox "Please Select For Material.", vbInformation
            Exit Sub
        End If
         
        'ShowLocations
        Call ShowLocations
        Exit Sub
        
    End If
    Call AddMaterial
End Sub

Private Sub AddMaterial()

    ''''''''''''''' If Its OK, Then Add The Material '''''''''''''''''''
    Dim ITM As ListItem
    
    Set ITM = LV.ListItems.add(, , cmbItem.ID)
    ITM.Tag = 0
    ITM.ListSubItems.add , , cmbItem.Text
    ITM.ListSubItems.add , , ""
    ITM.ListSubItems.add , , txtPrice & lblUnit
    ITM.ListSubItems.add(, , txtQty).Tag = chkShowCCItems.Value
    
    If cmbPOType.ListIndex > 0 Then
        ITM.ListSubItems.add(, , lblAmt).Tag = cmbForMaterial.ID
    Else
        ITM.ListSubItems.add , , lblAmt
    End If
    ITM.ListSubItems.add , , ""
    
    cmbVend.Enabled = False
    txtQty = ""
    cmbItem.ID = ""
    cmbItem.SetFocus
    
End Sub

Private Sub ShowLocations()
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM VMaterialLocationWiseBalances WHERE MaterialID='" & cmbItem.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbLocation.Clear
        Do Until .EOF
            cmbLocation.AddItem !StoreName & " [" & !RackNo & "] [" & !ShelfNo & "] (" & Val(!ShelfQty & "") & ")"
            cmbLocation.List(.AbsolutePosition - 1, 1) = !Shelf_RefID
            cmbLocation.List(.AbsolutePosition - 1, 2) = Val(!ShelfQty & "")
            cmbLocation.List(.AbsolutePosition - 1, 3) = Val(!SheetsQty & "")
            cmbLocation.List(.AbsolutePosition - 1, 4) = ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    SSTab1.Enabled = False
    FraLocations.Visible = True
    FraLocations.ZOrder 0
    
End Sub

Private Sub cmdAddPic_Click()

    On Error Resume Next
        
    Dim f As New frmGetSingleTextValue
    Dim strDescription As String
    Load f
    If f.getTextValue(strDescription) = False Then Exit Sub
    
    PicFileName = ""
    CD1.Filter = "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
    CD1.ShowSave
     
    If Not (err.Number = 32755) Then
    
        Dim iPicIndex As Integer, iCount As Integer, iIndex As Integer
        iCount = ImgPics.count
        If iCount = 1 Then
            If lblImg(0).Tag = "" Then
                iIndex = 0
            Else
                iIndex = ImgPics.count
            End If
        Else
            iIndex = ImgPics.count
        End If
        
        If iIndex > 0 Then
            Load lblImg(iIndex)
            Load ImgPics(iIndex)
            lblImg(iIndex).Visible = True
            ImgPics(iIndex).Visible = True
                
'            lblImg(iIndex).Move lblImg(iIndex - 1).Left + lblImg(iIndex - 1).Width + 50, lblImg(iIndex - 1).Top
'            ImgPics(iIndex).Move ImgPics(iIndex - 1).Left + ImgPics(iIndex - 1).Width + 50, ImgPics(iIndex - 1).Top
        End If
        
        lblImg(iIndex).Tag = CD1.FileName
        lblImg(iIndex).Caption = strDescription
        
        ImgPics(iIndex).Tag = "0"   'Zero Means Newly Added....
        ImgPics(iIndex).Picture = LoadPicture(CD1.FileName)
         
        Call ResetPicturesPositions
        Call SetScroll
    End If
    
End Sub

Private Sub ResetPicturesPositions()
    
    Dim i As Integer, iIndex As Integer
    
    For i = 0 To ImgPics.count - 1
        If ImgPics(i).Picture <> 0 Then
        End If
    Next
    
    If iIndex > 0 Then
        Load lblImg(iIndex)
        Load ImgPics(iIndex)
        lblImg(iIndex).Visible = True
        ImgPics(iIndex).Visible = True
        
        lblImg(iIndex).Move lblImg(iIndex - 1).Left + lblImg(iIndex - 1).Width + 50, lblImg(iIndex - 1).Top
        ImgPics(iIndex).Move ImgPics(iIndex - 1).Left + ImgPics(iIndex - 1).Width + 50, ImgPics(iIndex - 1).Top
    End If
        
    PicInner.Width = ImgPics(iIndex).Left + ImgPics(iIndex).Width + 50
    
End Sub

Private Sub SetScroll()

    Dim lCount As Long
    lCount = ImgPics.count
    
    If lCount > 4 Then
        HScroll1.Min = 1
        HScroll1.Max = (lCount - 4)
        HScroll1.SmallChange = 1
        HScroll1.LargeChange = 1
        HScroll1.Value = 1
         
        HScroll1.Visible = True
    Else
        HScroll1.Visible = False
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelLocs_Click()
    SSTab1.Enabled = True
    FraLocations.Visible = False
End Sub

Private Sub cmdOK_Click()

    If Val(txtLocsTotal) <> Val(txtQty) Then
        MsgBox "Invalid Qty.", vbInformation
        Exit Sub
    End If
        
    Dim ITM As ListItem, ITMFrom As ListItem
     
    For i = LVLocsAll.ListItems.count To 1 Step -1
        If LVLocsAll.ListItems(i).Tag = cmbItem.ID Then
            LVLocsAll.ListItems.Remove i
        End If
    Next
    
    For i = 1 To LVLocs.ListItems.count
        Set ITMFrom = LVLocs.ListItems(i)
        Set ITM = LVLocsAll.ListItems.add(, , ITMFrom.Text)
        ITM.Tag = ITMFrom.Tag
        ITM.ListSubItems.add(, , ITMFrom.SubItems(1)).Tag = ITMFrom.ListSubItems(1).Tag
    Next
    
    FraLocations.Visible = False
    SSTab1.Enabled = True
    Call AddMaterial
End Sub

Private Sub HScroll1_Change()
    Call HScroll1_Scroll
End Sub

Private Sub HScroll1_Scroll()
    PicInner.Left = -((HScroll1.Value - 1) * iImgWidht)
End Sub


Private Sub ImgPics_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If ImgPics(Index).Picture <> 0 Then      'Newly Added.
        Dim strRet As String
        strRet = ShowPopUpMenu1("Remove Picture")
        If strRet = "Remove Picture" Then
            lblImg(Index).Caption = "Deleted"
            lblImg(Index).Tag = ""
            ImgPics(Index).Picture = LoadPicture("")
            ImgPics(Index).Tag = ""
        End If
    End If
    
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)

    If LV.ListItems.count = 0 Then Exit Sub
    
    If KeyCode = 46 Then
        Dim isRcvd As Boolean
        isRcvd = False
        If OrderNo <> "" And LV.SelectedItem.Tag <> "0" Then
            Dim rs As New ADODB.Recordset
            With rs
                .Open "Select QtyRcvd From VVendOrdersToRcv Where VendOrderDetailEntryID=" & LV.SelectedItem.Tag, con, adOpenForwardOnly, adLockReadOnly
                If .RecordCount = 0 Then 'Fully Received.
                    isRcvd = True
                Else
                    If Val(!QtyRcvd & "") > 0 Then
                        isRcvd = True
                    End If
                End If
                .Close
                
                If isRcvd Then
                    MsgBox "Can't Remove. Material Has Been Rcvd Against This PO.", vbInformation
                    Exit Sub
                Else
                    DelEntryList = DelEntryList & Val(LV.SelectedItem.Tag) & ","
                    LV.ListItems.Remove LV.SelectedItem.Index
                    
                End If
            End With
        Else
            LV.ListItems.Remove LV.SelectedItem.Index
        End If
    End If
    
End Sub

Private Sub LVLocs_KeyDown(KeyCode As Integer, Shift As Integer)
    If LVLocs.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        LVLocs.ListItems.Remove LVLocs.SelectedItem.Index
    End If
End Sub

Private Sub mnuDelRow_Click()
   
    
    
End Sub

Private Sub cmdSave_Click()
    
    If Saved Then
        Unload Me
        Me.Show
    End If
    
End Sub

Private Sub cmdSaveClose_Click()
    
    If Saved Then
        Unload Me
    End If
    
End Sub

Private Sub cmdUnAssign_Click()

    On Error GoTo err
    If LV.ListItems.count = 0 Then Exit Sub
    Dim iAffect As Integer
    
    'con.Execute "Delete From VendAssItems Where EntryID=" & Val(LV.SelectedItem.Key), iAffect
    'If iAffect > 0 Then
    '    'Call cmbNick_Click
    '    cmdUnAssign.Visible = False
    'End If
    
    LV.ListItems.Remove LV.SelectedItem.Index
    cmbVend.Enabled = LV.ListItems.count = 0
    Call GetTotal
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub dtOrdDate_Change()
    txtPOrdNo = GetNextPOrdNo
    DTDelivery = DateAdd("d", 7, dtOrdDate)
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Exit Sub 'For Now
    LV.Sorted = True
    If LV.SortKey = ColumnHeader.Index - 1 Then
        If LV.SortOrder = lvwAscending Then
            LV.SortOrder = lvwDescending
            ColumnHeader.Icon = "Up"
        Else
            LV.SortOrder = lvwAscending
            ColumnHeader.Icon = "Down"
        End If
    Else
        LV.ColumnHeaderIcons = Nothing
        Set LV.ColumnHeaderIcons = ImageList2
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        ColumnHeader.Icon = "Down"
    End If

    On Error Resume Next
    LV.SelectedItem.EnsureVisible

End Sub

Private Sub LV_DblClick()
    
    If LV.ListItems.count = 0 Then Exit Sub
    
    If cmbVend.MatchFound = False Then Exit Sub
    
    For i = 1 To LV.ColumnHeaders.count
        LV.ColumnHeaders(i).Width = Val(LV.ColumnHeaders(i).Tag)
    Next
    If strCompany = "IAA" Or strCompany = "Kami" Then
        If iColNo = 4 Then
            iColNo = 5
        End If
    End If
    
    If iColNo <> 4 Then
        iColNo = 5
    End If
    
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(iColNo).Width
        .Text = Val(LV.SelectedItem.SubItems(iColNo - 1))
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
 
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    
'    If strCompany = "IAA" Or strCompany = "Kami" Then
'        txtPrice.Locked = True
'    End If
    
    Call AddToCombo(cmbCollectionBy, "DISTINCT Order_Collection_By", "VendOrders", " WHERE ISNULL(Order_Collection_By,'')<>''")
    cmbVend.ListHeight = 3000
    iImgWidht = 2445
    dtOrdDate = GetServerDate
    Call dtOrdDate_Change
    For i = 1 To LV.ColumnHeaders.count
        LV.ColumnHeaders(i).Tag = LV.ColumnHeaders(i).Width
    Next
    
    
    SSTab1.TabVisible(2) = False
    
    txtPOrdNo = GetNextPOrdNo

    cmbVend.ListHeight = 2600
    cmbVend.ClearVals
'    If LCase(CurrentUserName) = "mohsin" Or LCase(CurrentUserName) = "administrator" Then
'        cmbVend.AddVals con, "AccTitle", "VVenders", "VendID", " WHERE Active = 1 AND RestrictPO=0 ORDER BY AccTitle"
'    Else
'        cmbVend.AddVals con, "AccTitle", "VVenders", "VendID", " WHERE Active = 1  AND ImportVender = 0 AND RestrictPO=0 Order By AccTitle"
'    End If
'    If strCompany = "Kami" Then
        cmbVend.AddVals con, "AccTitle", "VVenders", "AccNo", " WHERE Active = 1  ORDER BY AccTitle"
'    Else
'        cmbVend.AddVals con, "AccTitle", "VVendersANDMakers", "AccNo", " WHERE Active = 1  ORDER BY AccTitle"
'    End If
'    cmbVend.AddVals con, "AccTitle", "VVenders", "VendID"
    cmbItem.ListHeight = 2400
'    cmbItem.ClearVals
'    cmbItem.AddVals con, "MaterialName", "Material", "RMID", ""
    'txtRemarks = "1. Bill & Issue slip should state the number of PO"
    OrderNo = ""
    DelEntryList = ""
    
    txtRunningPONo = GetSingleLongValue("MAX(RunningPONo)", "VendOrders") + 1
        
    cmbPOType.ListIndex = 0
    cmbProcessOrder.ListIndex = 0
    
    cmbForMaterial.ListHeight = 4000
    cmbForMaterial.ClearVals
    cmbForMaterial.AddVals con, "RMName + '     {' + RMID1 + '}'", "RM", "RMID", " WHERE InActive=0 ORDER BY RMName"
        
    Call AddToCombo(cmbCurrency, "Distinct Curr", "ForeignCustomers")
    SSTab1.TabVisible(1) = False
    
    Call AddToCombo(cmbSteel, "SteelType", "SteelTypes", , , "SteelID")
    
    cmbCustomerPONo.ListHeight = 3000
    cmbCustomerPONo.AddVals con, "OrderNo + ' {' + InternalRefNo + '}'", "FCustomerOrders", "OrderNo"
'    If strCompany = "Dr-Frgz" Then
'        chkShowCCItems.Visible = True
'    Else
'        chkShowCCItems.Visible = False
'    End If
    chkShowCCItems.Visible = False
    bSample = False
End Sub

Private Sub Form_Resize()

    Exit Sub
    On Error Resume Next
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
    
    Dim TitleBottom As Long
    
    TitleBottom = lblAmt.Top + lblAmt.Height 'fTitle.Top + fTitle.Height
    
    LV.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + Frame6.Height + 100)
    
    Frame6.Move Me.ScaleWidth - Frame6.Width - 50, Me.ScaleHeight - Frame6.Height - 50
    
    chkPrint.Move LV.Left, Frame6.Top + 50
    chkOrderNo.Move LV.Left, Frame6.Top + 350
 
End Sub


'Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)

'cmdRemove.Move LV.Left + LV.ColumnHeaders(6).Left + LV.ColumnHeaders(6).Width, LV.Top + LV.SelectedItem.Top
'cmdRemove.Visible = True

'End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If x > Val(LV.ColumnHeaders(i).Left) And x < (Val(LV.ColumnHeaders(i).Left) + Val(LV.ColumnHeaders(i).Width)) Then
            iColNo = i
            Exit For
        End If
    Next
    
End Sub


Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            
        LV.SelectedItem.SubItems(iColNo - 1) = Val(txtEdit)
        txtEdit.Visible = False
        
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
    
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub

Private Sub txtPrice_Change()
    lblAmt = "Rs. " & Format(Val(txtQty) * Val(txtPrice), "0.00")
End Sub

Private Sub txtQty_Change()

    lblAmt = "Rs. " & Format(Val(txtQty) * Val(txtPrice), "0.00")

End Sub

Private Sub txtQty_GotFocus()

    txtQty = Val(txtQty)

    txtQty.SelStart = 0
    txtQty.SelLength = Len(txtQty.Text)
End Sub

Private Sub txtEditQty_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    
    If KeyAscii = 13 Then
        
        If iColNo = 5 Then
            LV.SelectedItem.SubItems(iColNo - 1) = Val(txtEditQty) & " " & LV.SelectedItem.ListSubItems(3).Tag
        ElseIf iColNo = 4 Then
            LV.SelectedItem.SubItems(iColNo - 1) = Val(txtEditQty) '& " " & LV.SelectedItem.ListSubItems(3).Tag
        End If
        LV.SelectedItem.SubItems(5) = Format(Val(LV.SelectedItem.SubItems(4)) * Val(LV.SelectedItem.SubItems(3)), "0.00")
        
        txtEditQty.Visible = False
        Call GetTotal
    ElseIf KeyAscii = 27 Then
        txtEditQty.Visible = False
    End If
    Call OnlyNums(KeyAscii)
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Function Saved() As Boolean

    Saved = False
    On Error GoTo err
    
    If LV.ListItems.count = 0 And OrderNo = "" Then
        MsgBox "Order Must Contain At Least One Material !", vbInformation
        cmbItem.SetFocus
        Exit Function
    ElseIf LV.ListItems.count = 0 Then
        If MsgBox("There is no item in this order.Do You Want to delete it?", vbQuestion + vbYesNo) = vbYes Then
            con.Execute "Delete From VendOrders Where OrderNo='" & OrderNo & "'"
            con.Execute "Delete From VendOrderDetail Where OrderNo='" & OrderNo & "'"
            Saved = True
            Exit Function
        Else
            MsgBox "Order Must Contain At Least One Material !", vbInformation
            cmbItem.SetFocus
            Exit Function
        End If
    End If
    If DTDelivery.Value < dtOrdDate.Value Then
        MsgBox "Delivery Date Can't be Less than Order Date.", vbInformation
        Exit Function
    End If
    
    If cmbPOType.ListIndex = -1 Then
        MsgBox "Please Select Po Type.", vbInformation
        Exit Function
    End If
    Dim TempIndex As Integer
    TempIndex = 1
    
    Dim iSteelID As Integer
    If cmbSteel.MatchFound = False Then
        iSteelID = 0
    Else
        iSteelID = cmbSteel.List(cmbSteel.ListIndex, 1)
    End If
    
    Dim lRefID As Long
    Call StartTrans(con)
    
    Dim lRunningPONo As Long
    lRunningPONo = GetSingleLongValue("MAX(RunningPONo)", "VendOrders") + 1
    If OrderNo <> "" Then
        'con.Execute "DELETE FROM VendOrderDetail WHERE OrderNo='" & OrderNo & "'"
        con.Execute "UPDATE VendOrders Set DT='" & dtOrdDate & "',AttnPerson='" & txtAttn & "',PaymentTerms='" & _
         txtPayTerms & "',Department='" & txtDept & "',DeliveryDT='" & DTDelivery & "',Remarks='" & _
         txtRemarks & "',ReqNo='" & txtReqNo & "',VeryUrgent=" & Abs(chkUrgent.Value) & ",Ref='" & txtRef & _
         "',Payments='" & txtPayments & "',Shipment='" & txtShipment & "',Instructions='" & _
         txtInstructions & "',Notes='" & txtNotes & "',Currency='" & cmbCurrency.Text & _
         "',SampleProvided=" & chkSample.Value & ",SampleProvidedDescription='" & txtSampleProvided & _
         "',DrawingProvided=" & chkDrawing.Value & ",DrawingProvidedDescription='" & txtDrawingProvided & _
         "',SteelProvided=" & chkSteelProvided.Value & ",SteelProvidedDescription='" & txtSteelProvided & "',POType=" & cmbProcessOrder.ListIndex & ",CustomerPONo='" & cmbCustomerPONo.ID & "',SteelType_RefID=" & iSteelID & ",Order_Collection_By='" & cmbCollectionBy.Text & "' WHERE OrderNo='" & OrderNo & "'"
    ElseIf Val(cmbVend.ID) = 0 Then
        con.Execute "INSERT INTO VendOrders(OrderNo,VendID,DT,AttnPerson,PaymentTerms,Department,DeliveryDT,Remarks,ReqNo,VeryUrgent,VendName,RunningPONo,IssuanceType) Values ('" & _
        txtPOrdNo & "','" & cmbVend.ID & "','" & Format(dtOrdDate, "dd-MMM-yyyy") & "','" & _
        txtAttn & "','" & txtPayTerms & "','" & txtDept & "','" & DTDelivery & "','" & txtRemarks & "','" & txtReqNo & "'," & chkUrgent.Value & ",'" & cmbVend.Text & "'," & lRunningPONo & "," & cmbPOType.ListIndex & ")"
    Else
        con.Execute "INSERT INTO VendOrders(OrderNo,VendID,DT,AttnPerson,PaymentTerms,Department,DeliveryDT,Remarks,ReqNo," & _
         "VeryUrgent,RunningPONo,IssuanceType,Ref,Payments,Shipment,Instructions,Notes,Currency,SampleProvided," & _
         "SampleProvidedDescription,DrawingProvided,DrawingProvidedDescription,SteelProvided,SteelProvidedDescription,POType,CustomerPONo,SteelType_RefID,SampleOrder,Order_Collection_By) VALUES ('" & txtPOrdNo & _
         "','" & cmbVend.ID & "','" & Format(dtOrdDate, "dd-MMM-yyyy") & "','" & txtAttn & "','" & txtPayTerms & _
         "','" & txtDept & "','" & DTDelivery & "','" & txtRemarks & "','" & txtReqNo & "'," & chkUrgent.Value & _
         "," & lRunningPONo & "," & cmbPOType.ListIndex & ",'" & txtRef & "','" & txtPayments & "','" & txtShipment & _
         "','" & txtInstructions & "','" & txtNotes & "','" & cmbCurrency.Text & "'," & chkSample.Value & ",'" & _
         txtSampleProvided & "'," & chkDrawing.Value & ",'" & txtDrawingProvided & "'," & chkSteelProvided.Value & _
         ",'" & txtSteelProvided.Text & "'," & cmbProcessOrder.ListIndex & ",'" & cmbCustomerPONo.ID & "'," & iSteelID & "," & Abs(bSample) & ",'" & cmbCollectionBy.Text & "')"
    End If
        
    If OrderNo <> "" Then
        If DelEntryList <> "" Then
            DelEntryList = Left(DelEntryList, Len(DelEntryList) - 1)
            con.Execute "DELETE FROM VendOrderDetail WHERE EntryID In(" & DelEntryList & ")"
        End If
    End If
    
    Dim lVOD_EntryID As Long
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If Val(.SubItems(4)) > 0 Then
                If OrderNo <> "" And .Tag <> "0" Then
                    con.Execute "UPDATE VendOrderDetail SET Rate=" & Val(.SubItems(3)) & ",Qty=" & Val(.SubItems(4)) & ",ItemRemarks='" & .SubItems(7) & "' WHERE EntryID=" & Val(.Tag)
                    lVOD_EntryID = Val(.Tag)
                    If .ListSubItems(6).Tag <> "" Then
                        If .ListSubItems(6).Tag = "Remove" Then
                            con.Execute "DELETE FROM VendOrderDetail_PDF WHERE VOD_RefID=" & lVOD_EntryID
                        ElseIf .ListSubItems(6).Tag = "Added" Then
                            'Do Nothing....
                        Else
                            Call AttachPDF(lVOD_EntryID, .ListSubItems(6).Tag)
                        End If
                    End If
                    
                Else
                    con.Execute "INSERT INTO VendOrderDetail(OrderNo,RMID,Rate,Qty,ForRMID,CCItem,ItemRemarks) " & _
                     " VALUES('" & txtPOrdNo & "','" & .Text & "'," & _
                     Val(.SubItems(3)) & "," & Val(.SubItems(4)) & "," & Val(.ListSubItems(5).Tag) & "," & Val(.ListSubItems(4).Tag) & ",'" & .SubItems(7) & "')"
                    
                    lVOD_EntryID = GetSingleLongValue("MAX(EntryID)", "VendOrderDetail", " WHERE RMID='" & .Text & "'")
                    If .ListSubItems(6).Tag <> "" Then
                        If .ListSubItems(6).Tag = "Remove" Then
                            con.Execute "DELETE FROM VendOrderDetail_PDF WHERE VOD_RefID=" & lVOD_EntryID
                        ElseIf .ListSubItems(6).Tag = "Added" Then
                            'Do Nothing....
                        Else
                            Call AttachPDF(lVOD_EntryID, .ListSubItems(6).Tag)
                        End If
                    End If
                    
                     'Update Rate For Stock Value
                    If Val(.ListSubItems(4).Tag) = 0 Then
                        con.Execute "UPDATE RM SET Rate=" & Val(.SubItems(3)) & " Where RMID1='" & .Text & "'"
                        If cmbPOType.ListIndex > 0 Then
                            For j = 1 To LVLocsAll.ListItems.count
                                With LVLocsAll.ListItems(j)
                                    If .Tag = LV.ListItems(i).Text Then
                                        con.Execute "INSERT INTO VendOrderDetail_More(VOD_RefID,Shelf_RefID,IssdQty) VALUES(" & lVOD_EntryID & _
                                         "," & Val(.ListSubItems(1).Tag) & "," & Val(.SubItems(1)) & ")"
                                         
                                        Call UpdateMaterialLocation(.Tag, Val(.ListSubItems(1).Tag), Val(.SubItems(1)), 0)
                                    End If
                                End With
                            Next
                        End If
                        If cmbPOType.ListIndex = 1 Then
                            'Remove From QtyInStock
                            con.Execute "UPDATE RM SET QtyInStock=QtyInStock-" & Val(.SubItems(4)) & " WHERE RMID1='" & .Text & "'"
                        ElseIf cmbPOType.ListIndex = 2 Then
                            'Remove From Annealing
                            con.Execute "UPDATE RM SET QtyInStock=QtyInStock-" & Val(.SubItems(4)) & ",AnnealingStock=AnnealingStock-" & Val(.SubItems(4)) & " WHERE RMID1='" & .Text & "'"
                        ElseIf cmbPOType.ListIndex = 3 Then
                            'Remove From Male Grinding
                            con.Execute "UPDATE RM SET QtyInStock=QtyInStock-" & Val(.SubItems(4)) & ",MachiningStock=MachiningStock-" & Val(.SubItems(4)) & " WHERE RMID1='" & .Text & "'"
                        End If
                    End If
                End If
            End If
            
            
        End With
        
    Next
    
    con.CommitTrans
    Saved = True
    
   If chkPrint = vbChecked Then Call PrintReport
    
    Screen.MousePointer = vbDefault
    Exit Function
err:
    MsgBox err.Description
    con.RollbackTrans
    Screen.MousePointer = vbDefault
    
End Function

Private Sub txtQty_LostFocus()

    txtQty = Format(Val(txtQty), "0.00  ") & lblUnit.Tag

End Sub

Private Function GetNextEntryID() As String
    GetNextEntryID = GetMax("EntryID", "VendQuots") + 1
End Function

Private Sub PrintReport()
   
    Screen.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\RMPO.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VendOrders.OrderNo}='" & txtPOrdNo & "' ", rpt
    
    Screen.MousePointer = vbDefault
   
End Sub


Private Sub GetTotal()

   Exit Sub

   Dim total As Double
   For i = 1 To LV.ListItems.count
      total = total + Val(LV.ListItems(i).SubItems(7))
   Next
   lbltotalBalls = total
End Sub

Public Sub EditOrder(strOrderNo As String)
    
    OrderNo = strOrderNo
    Dim Rec As New ADODB.Recordset, ITM As ListItem
    With Rec
        .Open "SELECT * FROM VVendOrders WHERE OrderNo='" & OrderNo & "'", con, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtPOrdNo = OrderNo
            cmbVend.ID = !VendID & ""
            dtOrdDate = !DT & ""
            txtReqNo = !ReqNo & ""
            txtAttn = !AttnPerson & ""
            txtPayTerms = !PaymentTerms
            txtRemarks = !Remarks & ""
            txtDept = !Department & ""
            DTDelivery = !DeliveryDT & ""
            chkUrgent.Value = Abs(!VeryUrgent)
            
            chkSample.Value = Abs(!SampleProvided)
            txtSampleProvided = !SampleProvidedDescription & ""
            chkDrawing.Value = Abs(!DrawingProvided)
            txtDrawingProvided = !DrawingProvidedDescription & ""
            chkSteelProvided.Value = Abs(!SteelProvided)
            txtSteelProvided = !SteelProvidedDescription & ""
            
            txtRef = !Ref & ""
            txtPayments = !Payments & ""
            txtShipment = !Shipment & ""
            txtInstructions = !Instructions & ""
            txtNotes = !Notes & ""
            cmbCurrency = !Currency & ""
            cmbProcessOrder.ListIndex = Val(!POType & "")
            
            cmbCustomerPONo.ID = !CustomerPONo & ""
            bSample = !SampleOrder
            cmbCollectionBy.Text = !Order_Collection_By & ""
            
            For i = 0 To cmbSteel.ListCount - 1
                If Val(cmbSteel.List(i, 1)) = Val(!SteelType_RefID & "") Then
                    cmbSteel.ListIndex = i
                    Exit For
                End If
            Next
            LV.ListItems.Clear
            For i = 1 To .RecordCount
                Set ITM = LV.ListItems.add(, , !RMID1 & "")
                ITM.Tag = !VendOrderDetailEntryID
                ITM.ListSubItems.add , , !RMName & " {" & !RMSize & "} {" & !RMID1 & "}" ' & ""
                ITM.ListSubItems.add , , "" 'MatType(Val(!MaterialType & ""))
                ITM.ListSubItems.add , , Format(!Rate & "  / " & !Unit, "0.00")
                ITM.ListSubItems.add , , Val(!Qty & "") & "  " & !Unit & ""
                ITM.ListSubItems.add , , Format(Val(!Rate & "") * Val(!Qty & ""), "0.00")
                
                If Val(!VOD_RefID & "") <> 0 Then
                    ITM.ListSubItems.add().Tag = "Added"
                Else
                    ITM.ListSubItems.add().Tag = ""
                End If
                ITM.ListSubItems.add , , !ItemRemarks & ""
                .MoveNEXT
            Next
            cmbVend.Enabled = False
            Editing = True
        End If
    End With
    Me.Show
    
End Sub


Private Function GetNextPOrdNo() As String

    Dim Prefix As String
    'ORD-DDMM-001
    If strCompany = "Dr-Frgz" Or strCompany = "Banzai" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne" Then
        Dim lOrderNo As Long
        lOrderNo = GetSingleLongValue("MAX(CAST(LEFT(OrderNo,CHARINDEX('/',OrderNo)-1) AS INT))", "VendOrders", " WHERE RIGHT(OrderNo,4)='" & Format(dtOrdDate, "MMyy") & "'")
        lOrderNo = lOrderNo + 1
        GetNextPOrdNo = lOrderNo & Format(dtOrdDate, "/MMyy")
    Else
        Prefix = "ORD-" & Format(dtOrdDate, "ddMMyy") & "-"
        GetNextPOrdNo = Prefix & Format(GetMax("RIGHT(RTRIM(OrderNo),3)", "VendOrders", "WHERE LEFT(OrderNo," & Len(Prefix) & ")='" & Prefix & "'") + 1, "000")
    End If
    
End Function

Private Sub cmdNewMaterial_Click()

    Dim f As New frmNewRM
    Load f
    f.AddNew

End Sub
Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If LV.ListItems.count = 0 Then Exit Sub
    If Button <> 2 Then Exit Sub
    Dim strRet As String
     
    Dim strValue As String
    strValue = LV.SelectedItem.ListSubItems(6).Tag
     
    If bSample Then
        strRet = ShowPopUpMenu1("View PDF", "Remarks")
    ElseIf strValue <> "" And strValue <> "Remove" Then
        strRet = ShowPopUpMenu1("Remove PDF", "Remarks")
    Else
        strRet = ShowPopUpMenu1("Attach PDF", "Remarks")
    End If
    
    If strRet = "Attach PDF" Then
        Call BrowsePDF
    ElseIf strRet = "Remove PDF" Then
        LV.SelectedItem.ListSubItems(6).Tag = "Remove"
    ElseIf strRet = "View PDF" Then
        Call ShowTechDrawing("FOrderItems_PDF INNER JOIN FOrderItems ON FOrderItems.ID=FOrderItems_PDF.FOI_RefID", "PDF", " WHERE OrderNo='" & cmbCustomerPONo.ID & "' AND CompItemCode='" & LV.SelectedItem.Text & "'", "Sample Attached", Me.CD1)
    ElseIf strRet = "Remarks" Then
    
        With txtItemRemarks
            .Text = LV.SelectedItem.SubItems(7)
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        
        With FraRemarks
            .Caption = "Remarks of " & LV.SelectedItem.SubItems(1) & ":"
            If LV.SelectedItem.Top + LV.Top > 3045 Then
                .Move FraRemarks.Left, 3045
            Else
                .Move FraRemarks.Left, LV.SelectedItem.Top + LV.Top
            End If
            .Visible = True
        End With
        txtItemRemarks.SetFocus
    End If
    
End Sub

Private Sub txtRMQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmbLocation.MatchFound = False Then
            MsgBox "Invalid Location.", vbInformation
            Exit Sub
        End If
        If Val(txtRMQty) > Val(txtInHand) Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        End If
        Dim i As Integer
        For i = 1 To LVLocs.ListItems.count
            If Val(cmbLocation.List(cmbLocation.ListIndex, 1)) = Val(LVLocs.ListItems(i).ListSubItems(1).Tag) Then
                MsgBox "Location already Added.", vbInformation
                Exit Sub
            End If
        Next
        
        Dim ITM As ListItem
        Set ITM = LVLocs.ListItems.add(, , cmbLocation.Text)
        ITM.Tag = cmbItem.ID
        ITM.ListSubItems.add(, , Val(txtRMQty)).Tag = cmbLocation.List(cmbLocation.ListIndex, 1)
        
        Call getTotalLocs
    End If
    
End Sub

Private Sub getTotalLocs()
    Dim i As Integer, dTotal As Double
    For i = 1 To LVLocs.ListItems.count
        dTotal = dTotal + Val(LVLocs.ListItems(i).SubItems(1))
    Next
    txtLocsTotal = dTotal
End Sub
   
    
Private Sub BrowsePDF()

    On Error Resume Next
    Dim Picfile As String
    CD1.ShowSave
    
    If Not (err.Number = 32755) Then
        Picfile = CD1.FileName
        LV.SelectedItem.ListSubItems(6).Tag = Picfile
    End If
       
End Sub

Private Sub AttachPDF(p_lEntryID As Long, strPDFPath As String)
    
    On Error Resume Next
    con.Execute "DELETE FROM VendOrderDetail_PDF WHERE VOD_RefID=" & p_lEntryID
    con.Execute "INSERT INTO VendOrderDetail_PDF(VOD_RefID) VALUES(" & p_lEntryID & ")"
    
    If strPDFPath <> "" Then
        Dim Rec As New ADODB.Recordset
        Dim rsPic As New ADODB.Stream
        Rec.Open "SELECT EntryID,PDF FROM VendOrderDetail_PDF WHERE VOD_RefID=" & p_lEntryID, con, adOpenStatic, adLockOptimistic
        If Rec.RecordCount > 0 Then
            rsPic.Open
            rsPic.Type = adTypeBinary
            rsPic.LoadFromFile strPDFPath
            Rec.Fields(1) = rsPic.Read
            rsPic.Close
            Rec.Update
        End If
        
        Rec.Close
        
        Set Rec = Nothing
        Set rsPic = Nothing
    End If

End Sub

Public Sub ShowMe(Optional p_bSample As Boolean = False)

    bSample = p_bSample
    
    cmbCustomerPONo.ClearVals
    cmbCustomerPONo.AddVals con, "OrderNo + ' {' + InternalRefNo + '}'", "FCustomerOrders", "OrderNo", "WHERE OrderType='Sample Order'"
     
    chkShowCCItems.Value = vbChecked
    chkShowCCItems.Enabled = False
    
    If p_bSample Then
        Label1(0).Caption = "Samples PO"
        Label1(1).Caption = "Samples PO"
    End If
    
    Me.Show
    
End Sub
    
Private Sub cmdOKPackingDetails_Click()
    
    LV.SelectedItem.SubItems(7) = txtItemRemarks.Text
    
    FraRemarks.Visible = False
    
End Sub

Private Sub cmdCancelPackDetails_Click()
    FraRemarks.Visible = False
End Sub


Private Sub chkShowPicture_Click()

    If chkShowPicture.Value = vbChecked Then
        Call LV_ItemClick(LV.SelectedItem)
    Else
        PicProduct.Visible = False
    End If
    
End Sub


Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)

    If chkShowPicture.Value = vbChecked Then
        Dim rs As New ADODB.Recordset
        With rs
            .Open "SELECT Pic FROM RM WHERE RMID1='" & (LV.SelectedItem.Text) & "'", con, adOpenForwardOnly, adLockReadOnly
            If Not IsNull(![Pic]) Then
                Set ImgProduct.DataSource = rs
                ImgProduct.DataField = "Pic"
            Else
                ImgProduct.Picture = LoadPicture("")
            End If
            .Close
        End With
        Set rs = Nothing
        Dim lTop As Long
        lTop = LV.Top + LV.SelectedItem.Top
        If (lTop + PicProduct.Height) > LV.Top + LV.Height Then
            lTop = LV.Top + LV.Height - PicProduct.Height
        End If
        PicProduct.Move LV.Left + LV.Width - PicProduct.Width, lTop
        PicProduct.Visible = True
    Else
'        If PicProduct.Visible = True Then PicProduct.Visible = False
    End If
End Sub

