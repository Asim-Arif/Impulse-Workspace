VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmTransferToReadyFinishStock 
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   9030
   ScaleMode       =   0  'User
   ScaleWidth      =   10721.69
   WindowState     =   2  'Maximized
   Begin VB.Frame FraLocation 
      BackColor       =   &H00E7EBEF&
      Height          =   2790
      Left            =   420
      TabIndex        =   18
      Top             =   3600
      Visible         =   0   'False
      Width           =   10980
      Begin ComboList.Usercontrol1 cmbStore 
         Height          =   285
         Left            =   0
         TabIndex        =   23
         Top             =   315
         Width           =   2175
         _ExtentX        =   3836
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
      Begin ComboList.Usercontrol1 cmbLocation 
         Height          =   285
         Left            =   2175
         TabIndex        =   22
         Top             =   315
         Width           =   7665
         _ExtentX        =   13520
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
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   0
         TabIndex        =   25
         Top             =   795
         Width           =   10965
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E3FBFB&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9825
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   315
         Width           =   1140
      End
      Begin VB.CommandButton cmdPost 
         Caption         =   "&Post"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8790
         TabIndex        =   20
         Top             =   2355
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cance&l"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9885
         TabIndex        =   19
         Top             =   2355
         Width           =   1050
      End
      Begin MSForms.Label Label5 
         Height          =   240
         Left            =   0
         TabIndex        =   26
         Top             =   570
         Width           =   10965
         BackColor       =   11517387
         Caption         =   " Remarks"
         Size            =   "19341;423"
         BorderColor     =   16512
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label4 
         Height          =   240
         Left            =   0
         TabIndex        =   24
         Top             =   90
         Width           =   10965
         BackColor       =   11517387
         Caption         =   $"frmTransferToReadyFinishStock.frx":0000
         Size            =   "19341;423"
         BorderColor     =   16512
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.Frame FAcc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   9075
      Left            =   60
      TabIndex        =   0
      Top             =   -60
      Width           =   11865
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   75
         TabIndex        =   9
         Top             =   1110
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
      Begin ComboList.Usercontrol1 cmbCustCode 
         Height          =   285
         Left            =   60
         TabIndex        =   37
         Top             =   1620
         Width           =   1230
         _ExtentX        =   2170
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
      Begin VB.Frame FTitle 
         Height          =   645
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   11865
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Transfer to Ready Finish"
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
            TabIndex        =   35
            Top             =   135
            Width           =   11745
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Transfer to Ready Finish"
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
            TabIndex        =   36
            Top             =   150
            Width           =   11730
         End
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   3135
         TabIndex        =   17
         Top             =   2850
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox txtLotNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10065
         TabIndex        =   5
         Top             =   1620
         Width           =   825
      End
      Begin VB.Frame Frame1 
         Height          =   480
         Left            =   2640
         TabIndex        =   12
         Top             =   8565
         Width           =   5475
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   1965
            TabIndex        =   15
            Top             =   135
            Width           =   2430
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4395
            TabIndex        =   14
            Top             =   135
            Width           =   1020
         End
         Begin VB.ComboBox cmbfield 
            Height          =   315
            ItemData        =   "frmTransferToReadyFinishStock.frx":00DA
            Left            =   60
            List            =   "frmTransferToReadyFinishStock.frx":00F3
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   135
            Width           =   1875
         End
      End
      Begin ComboList.Usercontrol1 cmbItems 
         Height          =   285
         Left            =   1305
         TabIndex        =   4
         Top             =   1620
         Width           =   8745
         _ExtentX        =   15425
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
      Begin VB.ComboBox cmbRange 
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
         ItemData        =   "frmTransferToReadyFinishStock.frx":013A
         Left            =   5265
         List            =   "frmTransferToReadyFinishStock.frx":0150
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1110
         Width           =   1995
      End
      Begin VB.CommandButton cmdTransfer 
         Caption         =   "Transfer to Ready Finish"
         Height          =   345
         Left            =   60
         TabIndex        =   3
         Top             =   8640
         Width           =   2205
      End
      Begin MSComctlLib.ListView LV 
         Height          =   6600
         Left            =   60
         TabIndex        =   1
         Top             =   1950
         Width           =   11715
         _ExtentX        =   20664
         _ExtentY        =   11642
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Maker"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "PO #"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Process"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Item ID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Pass Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Transfer"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Lot #"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   9075
         TabIndex        =   7
         Top             =   1110
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
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
         Format          =   154599427
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   7260
         TabIndex        =   8
         Top             =   1110
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
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
         Format          =   154599427
         CurrentDate     =   37055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTransferToReadyFinishStock.frx":01A0
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
         Index           =   1
         Left            =   45
         TabIndex        =   16
         Top             =   1410
         Width           =   10785
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTransferToReadyFinishStock.frx":0298
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
         Left            =   45
         TabIndex        =   11
         Top             =   885
         Width           =   11130
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   810
         Left            =   10905
         TabIndex        =   10
         Top             =   1110
         Width           =   840
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "1482;1429"
         Accelerator     =   82
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   375
         Left            =   10455
         TabIndex        =   2
         Top             =   8610
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "2355;661"
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
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   10
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferToReadyFinishStock.frx":0353
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferToReadyFinishStock.frx":050F
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraLocation_QEL 
      BackColor       =   &H00E7EBEF&
      Height          =   2790
      Left            =   420
      TabIndex        =   27
      Top             =   3600
      Visible         =   0   'False
      Width           =   10980
      Begin VB.TextBox txtRemarks_QEL 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4065
         TabIndex        =   31
         Top             =   315
         Width           =   6900
      End
      Begin VB.TextBox txtQty_QEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E3FBFB&
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2940
         TabIndex        =   30
         Top             =   315
         Width           =   1140
      End
      Begin VB.CommandButton cmdPost_QEL 
         Caption         =   "&Post"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8790
         TabIndex        =   29
         Top             =   2355
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancel_QEL 
         Caption         =   "Cance&l"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9885
         TabIndex        =   28
         Top             =   2355
         Width           =   1050
      End
      Begin MSForms.Label Label3 
         Height          =   240
         Left            =   0
         TabIndex        =   33
         Top             =   90
         Width           =   10965
         BackColor       =   11517387
         Caption         =   " Location                                                   Qty                  Remarks"
         Size            =   "19341;423"
         BorderColor     =   16512
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbLocation_QEL 
         Height          =   285
         Left            =   0
         TabIndex        =   32
         Tag             =   "PaymentTerms"
         Top             =   315
         Width           =   2940
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "5186;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopMenu"
      Begin VB.Menu mnuPrintPurchaseOrder 
         Caption         =   "Print Slip"
      End
      Begin VB.Menu mnuPrintPTC 
         Caption         =   "Print PTC"
      End
   End
End
Attribute VB_Name = "frmTransferToReadyFinishStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCourierInvoice As Boolean
Dim bShipping As Boolean
Dim bCustomInvoice As Boolean

Public Sub ShowMe(Optional p_bCustomInvoice As Boolean = True)
    bCustomInvoice = p_bCustomInvoice
    Me.Show
End Sub

Public Sub ShowForShipping()
    bShipping = True
    Me.Show
End Sub

Private Sub cmbLocation_matched()
    txtRemarks = GetSingleStringValue("Remarks", "RcvItemsSimpleDetail_Placement", " WHERE EntryID IN(SELECT MAX(EntryID) FROM VRcvItemsSimpleDetail_Placement WHERE ItemID='" & cmbItems.ID & "' AND Shelf_RefID=" & cmbLocation.ID & ")")
End Sub

Private Sub cmbRange_Click()

    DtTo = GetServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
        Case 0
            DtFrom = Date
        Case 1
            DtFrom = DateAdd("d", -15, Date)
        Case 2
            DtFrom = DateAdd("d", -30, Date)
        Case 3
            DtFrom = DateAdd("d", -60, Date)
        Case 4
            DtFrom = DateAdd("d", -90, Date)
        Case 5
            DtFrom.Enabled = True
            DtTo.Enabled = True
            DtFrom.SetFocus
            Exit Sub
    End Select
    
End Sub

Private Sub cmbStore_matched()
    cmbLocation.ClearVals
    cmbLocation.AddVals con, "RackNo + ' {' + ShelfNo + '}'", "VStoreShelfs", "EntryID", " WHERE Store_RefID=" & cmbStore.ID & " ORDER BY StoreName,RackNo,ShelfNo"
End Sub

Private Sub cmdCancel_Click()
    FraLocation.Visible = False
    FAcc.Enabled = True
End Sub

Private Sub cmdCancel_QEL_Click()
    FraLocation_QEL.Visible = False
    FAcc.Enabled = True
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdFind_Click()

    Dim StartIdx As Integer

    If LV.SelectedItem Is Nothing Then
       StartIdx = 1
    Else
       If LV.SelectedItem.Index = LV.ListItems.count Then
          StartIdx = 1
       Else
          StartIdx = LV.SelectedItem.Index
       End If
    End If

    Dim fldValue As String




    For i = StartIdx To LV.ListItems.count
       
       
       If cmbfield.ListIndex = 0 Then
          fldValue = Trim(Replace(LV.ListItems(i), "»", ""))
       Else
          fldValue = Trim(Replace(LV.ListItems(i).SubItems(cmbfield.ListIndex), "»", ""))
       End If
    
       If LCase(fldValue) Like "*" & LCase(Trim(txtFind)) & "*" Then
          LV.ListItems(i).Selected = True
          LV.ListItems(i).EnsureVisible
          If i <> StartIdx Then Exit For
       End If
    Next
    
    If i = LV.ListItems.count + 1 Then
        MsgBox "No Mathes Found.", vbInformation
        Set LV.SelectedItem = Nothing
    End If
End Sub


Private Sub cmdPost_Click()
    
    If Val(txtQty) = 0 Then
        MsgBox "Invalid Qty.", vbInformation
        Exit Sub
    ElseIf cmbLocation.MatchFound = False Then
        MsgBox "Invalid Location.", vbInformation
        Exit Sub
    End If
    
    Dim i As Integer, lEntryID As Long, dQty As Double, strPOs As String, strLots As String, lRISDP_EntryID As Long
    Call StartTrans(con)
    con.Execute "INSERT INTO TransferredToReadyFinishLots(UserName,MachineName,DT,ItemID,Remarks) VALUES('" & CurrentUserName & _
     "','" & strComputerName & "','" & GetServerDate & "','" & cmbItems.ID & "','" & txtRemarks & "')"
     
    lEntryID = GetSingleLongValue("MAX(EntryID)", "TransferredToReadyFinishLots", "WHERE MachineName='" & strComputerName & "'")
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                dQty = Val(.SubItems(7))
                con.Execute "INSERT INTO TransferredToReadyFinishLotsDetail(RefID,VRD_RefID,Qty) VALUES(" & _
                 lEntryID & "," & Val(.Tag) & "," & dQty & ")"
                
'                con.Execute "UPDATE Items SET InHand=ISNULL(InHand,0)+" & dQty & " WHERE ItemID='" & .SubItems(5) & "'"
                con.Execute "UPDATE VendRcvdDetail SET IssQty=IssQty+" & dQty & " WHERE EntryID=" & Val(.Tag)
                
                strPOs = strPOs & .SubItems(2) & ","
                strLots = strLots & .SubItems(8) & ","
                
            End If
        End With
    Next
    
    If strPOs <> "" Then
        strPOs = Left(strPOs, Len(strPOs) - 1)
    End If
    If strLots <> "" Then
        strLots = Left(strLots, Len(strLots) - 1)
    End If
    
    '''''''''''''''''''''''
    Dim strRcvNo As String
    strRcvNo = getRcvNo
    
    con.Execute "INSERT INTO RcvItemsSimple(RcvNo,DT,OrderNo,RcvdBy,RcvdFrom,UserName,ComputerName,CustCode) VALUES('" & _
     strRcvNo & "','" & Date & " " & ServerTime & "','" & strPOs & "','" & CurrentUserName & "','" & CurrentUserName & "','" & CurrentUserName & "','" & _
     strComputerName & "','" & cmbCustCode.ID & "')"
    
    lRefID = con.Execute("SELECT MAX(EntryID) FROM RcvItemsSimple").Fields(0).Value & ""
        
    Dim strMaterialID As String
    Dim lRISD_EntryID As Long
    
    strMaterialID = cmbItems.ID
    dQty = Val(txtQty)
    dBal = getItemBalance(strMaterialID, Date)
    dBal = dBal + dQty
    
    theDesc = "Transferred to Ready Finish"
    
    con.Execute "INSERT INTO ItemLedgerSimple(ItemID,DT,Description,Qty,Balance,OrderNo,Location,Remarks) VALUES('" & _
     strMaterialID & "','" & Date & "','" & strLots & "'," & dQty & "," & dBal & ",'" & strPOs & "','" & cmbLocation.Text & "','" & strLots & "')"
             
    lRefID = Val(con.Execute("Select Max(EntryID) From ItemLedgerSimple").Fields(0).Value & "")
            
    Call adjustItemLedger(strMaterialID, Date, dQty)

    con.Execute "INSERT INTO RcvItemsSimpleDetail(RcvNo,ItemID,Qty,InHand,LedgerRefID,OrderNo,Location,Remarks) Values('" & _
     strRcvNo & "','" & strMaterialID & "'," & dQty & "," & dBal + dQty & "," & lRefID & _
     ",'" & strPOs & "','" & cmbLocation.Text & "','" & strLots & "')"
            
    lRISD_EntryID = GetSingleLongValue("MAX(EntryID)", "RcvItemsSimpleDetail")
            
    con.Execute "INSERT INTO RcvItemsSimpleDetail_Placement(RISD_RefID,Shelf_RefID,RcvdQty,Remarks) VALUES(" & _
     lRISD_EntryID & "," & cmbLocation.ID & "," & dQty & ",'" & txtRemarks & "')"
    
    lRISDP_EntryID = GetSingleLongValue("MAX(EntryID)", "VRcvItemsSimpleDetail_Placement", " WHERE ItemID='" & strMaterialID & "' AND Shelf_RefID=" & cmbLocation.ID)
    con.Execute "UPDATE RcvItemsSimpleDetail_Placement SET Remarks='" & txtRemarks & "' WHERE EntryID=" & lRISDP_EntryID
            
    con.CommitTrans
    Call RefreshLV
    FAcc.Enabled = True
    FraLocation.Visible = False
    
End Sub

Private Function getRcvNo() As String
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select MAX(CAST(RIGHT(RcvNo,3) AS INT)) FROM RcvItemsSimple Where Convert(Char(8),DT,1)='" & Format(Date, "MM/dd/yy") & "'", con, adOpenForwardOnly, adLockReadOnly
        
        getRcvNo = "RCV-" & Format(Date, "yyMMdd-") & Format(Val(.Fields(0) & "") + 1, "000")
        
        .Close
    End With
    Set rs = Nothing
    
End Function

Private Sub cmdPost_QEL_Click()
    
    If Val(txtQty_QEL) = 0 Then
        MsgBox "Invalid Qty.", vbInformation
        Exit Sub
'    ElseIf cmbLocation_QEL.MatchFound = False Then
'        MsgBox "Invalid Location.", vbInformation
'        Exit Sub
    End If
    
    Dim i As Integer, lEntryID As Long, dQty As Double, strPOs As String, strLots As String, lRISDP_EntryID As Long
    Call StartTrans(con)
    con.Execute "INSERT INTO TransferredToReadyFinishLots(UserName,MachineName,DT,ItemID,Remarks) VALUES('" & CurrentUserName & _
     "','" & strComputerName & "','" & GetServerDate & "','" & cmbItems.ID & "','" & txtRemarks_QEL & "')"
     
    lEntryID = GetSingleLongValue("MAX(EntryID)", "TransferredToReadyFinishLots", "WHERE MachineName='" & strComputerName & "'")
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                dQty = Val(.SubItems(7))
                con.Execute "INSERT INTO TransferredToReadyFinishLotsDetail(RefID,VRD_RefID,Qty) VALUES(" & _
                 lEntryID & "," & Val(.Tag) & "," & dQty & ")"
                
                'con.Execute "UPDATE Items SET InHand=ISNULL(InHand,0)+" & dQty & " WHERE ItemID='" & .SubItems(5) & "'"
                con.Execute "UPDATE VendRcvdDetail SET IssQty=IssQty+" & dQty & " WHERE EntryID=" & Val(.Tag)
                
                strPOs = strPOs & .SubItems(2) & ","
                strLots = strLots & .SubItems(8) & ","
                
            End If
        End With
    Next
    
    If strPOs <> "" Then
        strPOs = Left(strPOs, Len(strPOs) - 1)
    End If
    If strLots <> "" Then
        strLots = Left(strLots, Len(strLots) - 1)
    End If
    
    '''''''''''''''''''''''
    Dim strRcvNo As String, strRcvdFrom As String
    strRcvNo = getRcvNo
    strRcvdFrom = "From Production Lot No.: " & strLots
    con.Execute "INSERT INTO RcvItemsSimple(RcvNo,DT,OrderNo,RcvdBy,RcvdFrom,UserName,ComputerName) VALUES('" & _
     strRcvNo & "','" & Date & " " & ServerTime & "','" & strPOs & "','" & cmbLocation_QEL & "','" & strRcvdFrom & "','" & CurrentUserName & "','" & _
     strComputerName & "')"
    
    lRefID = con.Execute("SELECT MAX(EntryID) FROM RcvItemsSimple").Fields(0).Value & ""
        
    Dim strMaterialID As String
    Dim lRISD_EntryID As Long
    
    strMaterialID = cmbItems.ID
    dQty = Val(txtQty_QEL)
    dBal = getItemBalance(strMaterialID, Date)
    dBal = dBal + dQty
    
    theDesc = "Transferred to Ready Finish"
    Dim strLocation As String
    If cmbLocation_QEL.MatchFound Then
        strLocation = cmbLocation_QEL.List(cmbLocation_QEL.ListIndex, 1)
    Else
        strLocation = cmbLocation_QEL.Text
    End If
    con.Execute "INSERT INTO ItemLedgerSimple(ItemID,DT,Description,Qty,Balance,OrderNo,Location) VALUES('" & _
     strMaterialID & "','" & Date & "','" & txtRemarks_QEL & "'," & dQty & "," & dBal & ",'" & strLots & "','" & strLocation & "')"
             
    lRefID = Val(con.Execute("SELECT MAX(EntryID) FROM ItemLedgerSimple").Fields(0).Value & "")
            
    Call adjustItemLedger(strMaterialID, Date, dQty)

    con.Execute "INSERT INTO RcvItemsSimpleDetail(RcvNo,ItemID,Qty,InHand,LedgerRefID,OrderNo,Location) VALUES('" & _
     strRcvNo & "','" & strMaterialID & "'," & dQty & "," & dBal & "," & lRefID & _
     ",'" & strLots & "','" & strLocation & "')"
    
    con.CommitTrans
    Call RefreshLV
    '-------------------------------------------------------------------------------------------------------------------------------
    FAcc.Enabled = True
    FraLocation_QEL.Visible = False

End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub


Private Sub cmdTransfer_Click()

    Dim lTotalQty As Long
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            lTotalQty = lTotalQty + Val(LV.ListItems(i).SubItems(7))
        End If
    Next
    
    FAcc.Enabled = False
    If strCompany = "QEL" Then
        txtQty_QEL = lTotalQty
        
        FraLocation_QEL.Visible = True
        Call AddToCombo(cmbLocation_QEL, "Location + ' (' + CAST(NetQty AS VARCHAR(50)) + ')'", "VItemsWithLocationwiseStock", " WHERE ItemID='" & cmbItems.ID & "' AND NetQty>0", , "Location,NetQty")
        'cmbLocation_QEL.SetFocus
    Else
        txtQty = lTotalQty
        FraLocation.Visible = True
        cmbStore.SetFocus
    End If
    
    
End Sub

Private Sub Command2_Click()

End Sub

Private Sub DtFrom_Change()
    'Call setDateRange(DtFrom)
End Sub

Private Sub DtTo_Change()
    'Call setDateRange(DtTo)
End Sub


Private Sub Form_Load()
    
    mnuPop.Visible = False
    
    If Not (strCompany = "QEL" Or (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai") Or strCompany = "Towne" Then
        mnuPrintPTC.Visible = False
    End If
    
    cmbCust.ListHeight = 3800
    cmbCust.AddItem "<All Makers>", "0"
    Call cmbCust.AddVals(con, "'{ ' + VendID1 + ' } ' + VenderName ", "Makers", "VendID")
    cmbCust.ID = "0"
    DtFrom = Date
    DtTo = Date
    cmbRange.ListIndex = 0

    cmbItems.DropDownOffLine = True
    'cmbItems.AddItem "<All Articles>", "0"
    Call cmbItems.AddVals(con, "ItemID + ' ' + ItemName + ' { ' + ISNULL(Attributes,'') + ' }'", "Items", "ItemID", " WHERE InActive=0 ORDER BY ItemName")
    'cmbItems.ID = "0"
    
    With cmbfield
        .Clear
        For i = 1 To LV.ColumnHeaders.count
            cmbfield.AddItem LV.ColumnHeaders(i).Text
        Next
    End With
        
    cmbLocation.ListHeight = 2400
    cmbStore.ListHeight = 2400
    cmbStore.AddVals con, "StoreName", "Stores", "EntryID", " ORDER BY StoreName"
    
    cmbCustCode.AddVals con, "CustCode", "ForeignCustomers", "CustCode"
    
End Sub

Private Sub RefreshLV()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim Cond As String
    Dim DTRange As String
    Dim ITM As ListItem
    
    Screen.MousePointer = vbHourglass
    
    If cmbRange.ListIndex = 0 Then
        DTRange = " OnlyDT Between '" & DateAdd("d", 0, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 1 Then
        DTRange = " OnlyDT Between '" & DateAdd("d", -15, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 2 Then
        DTRange = " OnlyDT Between '" & DateAdd("d", -30, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 3 Then
        DTRange = " OnlyDT Between '" & DateAdd("d", -60, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 4 Then
        DTRange = " OnlyDT Between '" & DateAdd("d", -90, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 5 Then
        DTRange = " OnlyDT BETWEEN '" & DtFrom & "' AND '" & DtTo & "'"
    End If
       
    Cond = " WHERE " & DTRange
    If cmbCust.ID <> "0" Then
        Cond = Cond & " AND VendID=" & cmbCust.ID & " "
    End If
    
    If cmbItems.ID <> "0" Then
        Cond = Cond & " AND ItemCode='" & cmbItems.ID & "' "
    End If
    
    If Trim(txtLotNo) <> "" Then
        Cond = Cond & " AND LotNo='" & txtLotNo & "' "
    End If
        
    Cond = Cond & " AND CustCode='" & cmbCustCode.ID & "'"
    
    Dim lProcessID As Long
    lProcessID = Val(getGeneralDataValue("ProcessIDForReadyFinish"))
    
    
    'Cond = Cond & " AND ProcessID=" & lProcessID & " AND (RcvdQty-ISNULL(Wastage,0)-ISNULL(LostQty,0))>IssQty AND ReqAuth=0"
    Cond = Cond & " AND NextProcessID IS NULL AND (RcvdQty-ISNULL(Wastage,0)-ISNULL(LostQty,0))>IssQty AND ReqAuth=0"
    
    With rs
        
        .Open "SELECT * FROM VVendReceivingList " & Cond & " ORDER BY DT", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        'LV.Sorted = False
        Do Until .EOF
'            If strCompany = "Tecno" Or (strCompany = "Dr-Frgz" And strSubCompany = "Electro") Then  'Check If this is the Process Just Before Scanning If Yes, THen Show otherwise Skip.
'                If IsNextProcessScanning(Val(!ProcessID & ""), !ItemCode) = False Then
'                    GoTo SKIPEntry
'                End If
'            End If
            Set ITM = LV.ListItems.add(, , .AbsolutePosition)
            ITM.Tag = Val(![VRD_EntryID] & "")
            ITM.ListSubItems.add(, , "{ " & !VendID1 & " } " & !VenderName).Tag = "" '![ImpMaterial] & ""
            ITM.ListSubItems.add , , ![RecieptID] & ""
            ITM.ListSubItems.add(, , Format(![DT] & "", "dd-MMM-yyyy")).Tag = ""
            ITM.ListSubItems.add(, , !Description & "").Tag = ""
            ITM.ListSubItems.add(, , ![ItemCode] & "").Tag = Val(!Issuance_RefID & "")
            
            ITM.ListSubItems.add , , Val(!RcvdQty & "") - (Val(!Wastage & "") + Val(!LostQty & "")) - Val(!IssQty & "") - Val(!ReWorkQty & "")
            ITM.ListSubItems.add , , Val(!RcvdQty & "") - (Val(!Wastage & "") + Val(!LostQty & "")) - Val(!IssQty & "") - Val(!ReWorkQty & "")
'            ITM.ListSubItems.add , , Val(!Wastage & "") + Val(!LostQty & "")
'            ITM.ListSubItems.add(, , IIf(!ReqAuth, "Under Auth.", "Authorized")).Tag = !LotNo & ""
            ITM.ListSubItems.add , , Val(!LotNo & "")
SKIPEntry:
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Function IsNextProcessScanning(lProcessID As Long, strItemID As String) As Boolean
    Dim iSNo As Integer
    iSNo = GetSingleLongValue("SNo", "ItemProcesses", " WHERE ItemID='" & strItemID & "' AND ProcessID=" & lProcessID)
    'Now Get The Next Process
    Dim bScanning As Boolean
    bScanning = GetSingleBooleanValue("TOP 1 Scanning", "ItemProcesses", " WHERE ItemID='" & strItemID & "' AND SNo>" & iSNo & " ORDER BY SNo")
    IsNextProcessScanning = bScanning
End Function

Private Sub Form_Resize()

    On Error Resume Next
    
    
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    FAcc.Top = (Me.ScaleHeight - FAcc.Height) / 2
   
    FraLocation.Move (Me.ScaleWidth - FraLocation.Width) / 2, (Me.ScaleHeight - FraLocation.Height) / 2
    FraLocation_QEL.Move (Me.ScaleWidth - FraLocation_QEL.Width) / 2, (Me.ScaleHeight - FraLocation_QEL.Height) / 2
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If LV.SortKey + 1 = ColumnHeader.Index Then
        If LV.SortOrder = lvwAscending Then
            LV.SortOrder = lvwDescending
            ColumnHeader.Icon = "Down"
        Else
            LV.SortOrder = lvwAscending
            ColumnHeader.Icon = "Up"
        End If
    Else
        LV.ColumnHeaderIcons = Nothing
        Set LV.ColumnHeaderIcons = ImageList2
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        ColumnHeader.Icon = "Down"
    End If
    Call RefreshSnos
    'If LV.Sorted = False Then LV.Sorted = True
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(8).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(8).Width
        .Text = Val(LV.SelectedItem.SubItems(7))
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Exit Sub
    If Button <> 2 Then Exit Sub
    If LV.ListItems.count = 0 Then
        Call ShowMenus(False)
    Else
        Call ShowMenus(True)
    End If
'    If bShipping Then
'        ShowMenusForShipping
'    ElseIf bCourierInvoice Then
'        ShowMenusforCourier
'    Else
'        If LV.ListItems.Count = 0 Then
'            Call ShowMenus(False)
'        Else
'            Call ShowMenus(True)
'        End If
'    End If
    'If LV.ListItems.Count = 0 Or Button <> 2 Or OutOfFinancialYear Then Exit Sub

    'if a Voucher Entry Show Delete Menu
    'If InStr(LV.SelectedItem, "-") > 0 Then Me.PopupMenu mnuPop
    
End Sub

Private Sub ShowMenusForShipping()

    On Error Resume Next
    Dim c As Control
    For Each c In Controls
        If TypeOf c Is Menu Then
            If c.Name <> "mnuPrintPList" Then
                c.Visible = False
            End If
        End If
    Next
    
    On Error GoTo 0
    
    If LV.ListItems.count > 0 Then
        mnuPrintPList.Visible = True
        mnuCustomPList.Visible = True
        mnuCustomPListOther.Visible = True
        mnuCommercialPList.Visible = True
        mnuCommercialPListOther.Visible = True
        mnuDashPackingLabels.Visible = True
        mnuPrintPackingLabels.Visible = True
    End If
'    mnuPrintDecDocs.Visible = False
    Me.PopUpMenu mnuPop
    
End Sub

Private Sub ShowMenus(vis As Boolean)
        
    
    Me.PopUpMenu mnuPop
    
End Sub

Private Sub ShowMenusforCourier()
         
    On Error Resume Next
    Dim c As Control
    For Each c In Controls
        If TypeOf c Is Menu Then
            c.Visible = False
        End If
    Next
    
    On Error GoTo 0
    mnuAdd.Visible = True
    If LV.ListItems.count > 0 Then
        dd.Visible = True
        mnuEdit.Visible = True
        mnuComDocs.Visible = True
        mnuPrintComInv.Visible = True
    End If
    mnuPrintDecDocs.Visible = False
    Me.PopUpMenu mnuPop
    
End Sub


Private Sub mnuAdd_Click()
    'Load frmNewCustomInvoice
    'frmNewCustomInvoice.Show
    'ShowForm frmNewCustomInvoice, MainForm, False, "New Custom Invoice"
    'frmNewCustomInvoice.AddNewProforma
    Load frmNewCustomInvoice
    frmNewCustomInvoice.AddNewInvoice True, bCourierInvoice
End Sub

Private Sub mnuAddShippInfo_Click()

    Load frmEditShippInfo
    'ShowForm f, MainForm, , "Shipp. Info"
    Call frmEditShippInfo.EditEntry(LV.SelectedItem.SubItems(3))
    
End Sub

Private Sub mnuBillOfExchange_Click()
    Call PrintBillOfExchange
End Sub

Private Sub PrintBillOfExchange(Optional Firstrpt As Boolean = True)

    On Error GoTo err
    Me.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptBillOfExchange.rpt")
    rpt.FormulaFields(3).Text = Firstrpt
    Dim frm As New frmPrevRpt
    frm.ShowReport "{CustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    Me.MousePointer = vbNormal
    
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description

End Sub

    
Private Sub mnuBillOfExchange2_Click()
    Call PrintBillOfExchange(False)
End Sub

Private Sub mnuCommercialPList_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3), False, True, LV.SelectedItem.SubItems(1) = "DECATHLON", LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2))
End Sub

Private Sub mnuCommercialPListOther_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3), True, True, LV.SelectedItem.SubItems(1) = "DECATHLON", LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2))
End Sub

Private Sub mnuCoveringLetter_Click()
    
    On Error GoTo err
    Me.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptComCoveringLetter.rpt")
    If LCase(CurrentUserName) = "mnaeem" Then
        rpt.FormulaFields(3).Text = "'Naeem Ahmed'"
    End If
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "{CustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub mnuCustomPList_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3))
End Sub

Private Sub mnuCustomPListOther_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3), True)
End Sub

Private Sub mnuDelPackingInfo_Click()
    
    If LV.ListItems.count = 0 Then Exit Sub
    Dim iAffect As Integer
    If MsgBox("Are you sure to Delete Packing List of Invoice # " & LV.SelectedItem.SubItems(3) & " ?", vbQuestion + vbYesNo) = vbYes Then
        Call StartTrans(con)
        con.Execute "Delete From CustomPList Where CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'", iAffect
        con.CommitTrans
        If iAffect > 0 Then
            LV.SelectedItem.ListSubItems(2).Tag = 0
        End If
    End If
    
End Sub

Private Sub mnuedit_Click()
    Call LV_DblClick
End Sub

Private Sub mnuEForm_Click()

    Load frmPrintFormE
    If LV.ListItems.count > 0 Then
        frmPrintFormE.ShowForm LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
    Else
        frmPrintFormE.Show , MainForm
    End If
    
End Sub

Private Sub mnuGSP_Click()

'    Load frmGSP
'    If LV.ListItems.Count > 0 Then
'        frmGSP.ShowForm LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
'    Else
'        frmGSP.Show , MainForm
'    End If
    
End Sub

Private Sub mnuPacking_Click()

'    If Left(mnuPacking.Caption, 3) = "Add" Then
'
'        Load frmNewPList
'        frmNewPList.ShowMe LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
'    Else
'        Load frmNewPList
'        Call frmNewPList.EditPList(LV.SelectedItem.SubItems(3))
'        ShowForm frmNewPList, MainForm, , "Packing List"
'    End If
    
End Sub

Private Sub mnuPrint_Click()
    Call PrintCustomInvoice
End Sub

Private Sub PrintCustomInvoice(Optional bImpMaterialDetail As Boolean = False)

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
    frmPrevRpt.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt
    
    Screen.MousePointer = vbDefault


    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub mnuPrint10_Click()
    Call PrintCommercialInvoice(True, False)
End Sub

Private Sub mnuPrint5_Click()
    Call PrintCommercialInvoice(True, True)
End Sub

Private Sub mnuPrintBillofExchange_Click()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptBillOfExchange.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VRptProformas.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintComInv_Click()
    Call PrintCommercialInvoice
End Sub

Private Sub PrintCommercialInvoice(Optional bPer As Boolean = False, Optional bFivePer As Boolean = True)
    
    On Error GoTo err

    Screen.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt
    Dim rpt As CRAXDDRT.Report
    
    If UCase(LV.SelectedItem.SubItems(1)) = "MIKASA" And LV.SelectedItem.SubItems(2) = "JAPAN" Then
        Set rpt = rptApp.OpenReport(rptexportpath & "\rptCommercialInvoiceMikasaJP.rpt")
    ElseIf bPer Then
        Set rpt = rptApp.OpenReport(rptexportpath & "\rptCommercialInvoiceFairTrade.rpt")
        rpt.FormulaFields(22).Text = bFivePer
    ElseIf LV.SelectedItem.SubItems(1) = "DECATHLON" Then
        Set rpt = rptApp.OpenReport(rptexportpath & "\rptDecathlonCommercialInvoice.rpt")
    Else
        Set rpt = rptApp.OpenReport(rptexportpath & "\rptCommercialInvoice.rpt")
    End If
    f.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt
    
    If Not bPer Then
        If LV.SelectedItem.SubItems(2) = "Canada" Then
            Dim rpt1 As CRAXDDRT.Report
            Set rpt1 = rptApp.OpenReport(rptexportpath & "\rptCommercialInvoice.rpt")
            Dim f1 As New frmPrevRpt
            Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
            Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
            
            Set FormulaFields = rpt1.FormulaFields
            
            For Each FormulaField In FormulaFields
                If FormulaField.Name = "{@forHeading}" Then
                    FormulaField.Text = "'C A N A D I A N   C U S T O M S   I N V O I C E'"
                    Exit For
                End If
            Next
            f1.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt1
        End If
    End If
    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub mnuPrintSummary_Click()

    On Error GoTo err
    Me.MousePointer = vbHourglass
    
    Call StartTrans(con)
    con.Execute "Delete From PrintFOrderList"
    
    For i = 1 To LV.ListItems.count
        con.Execute "Insert Into PrintFOrderList(Sno,Customer,Country," & _
         "OrderNo,OrderDT,OrderAmt,Status,Curr) Values(" & LV.ListItems(i) & _
         ",'" & LV.ListItems(i).SubItems(1) & "','" & LV.ListItems(i).SubItems(2) & _
         "','" & LV.ListItems(i).SubItems(3) & "','" & _
         LV.ListItems(i).SubItems(4) & "'," & Val(LV.ListItems(i).ListSubItems(5).Tag & "") & _
         ",'" & LV.ListItems(i).SubItems(6) & "','" & LV.ListItems(i).Tag & "')"
    Next i
    
    con.CommitTrans
    
    Load frmPrevRpt
    
    Dim strFromTo As String
    Dim iNumber As Integer
    If cmbRange.ListIndex = 0 Then
        iNumber = -15
    ElseIf cmbRange.ListIndex = 1 Then
        iNumber = -30
    ElseIf cmbRange.ListIndex = 2 Then
        iNumber = -60
    ElseIf cmbRange.ListIndex = 3 Then
        iNumber = -90
    End If
    
    If cmbRange.ListIndex = 4 Then
        strFromTo = Format(DtFrom, "dd-MMM-yyyy") & " To  " & Format(DtTo, "dd-MMM-yyyy")
    Else
        strFromTo = Format(DateAdd("d", iNumber, ServerDate), "dd-MMM-yyyy") & " To  " & Format(ServerDate, "dd-MMM-yyyy")
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptFProformaList.rpt")
    rpt.FormulaFields(3).Text = "'" & cmbCust.Text & "'"
    rpt.FormulaFields(4).Text = "'" & strFromTo & "'"
    
    frmPrevRpt.ShowReport "", rpt
    
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub mnuPrintCoverLetter_Click()

    On Error GoTo err

    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptDecathlonCoveringLetter.rpt")
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VCustomInvoice.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt

    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub mnuPrintDecDocs_Click()
    Load frmPrintDecathlonDocs
    frmPrintDecathlonDocs.ShowMe (LV.SelectedItem.SubItems(3))
End Sub

Private Sub mnuPrintInvWDetail_Click()

    Call PrintCustomInvoice(True)
    
End Sub

Private Sub mnuPrintPackingLabels_Click()

    On Error GoTo err
    
    If LV.SelectedItem.SubItems(1) <> "DECATHLON" Then
        MsgBox "Packing Labels are available only for DECATHLON at the moment.", vbInformation
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report, f As New frmPrevRpt
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptDecathlonPackingLabels.rpt")
    f.ShowReport "{VCustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    
    Me.MousePointer = vbNormal
    
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub mnuShippingInstructions_Click()
'    Load frmShippingInstructions
'    If LV.ListItems.Count = 0 Then
'        frmShippingInstructions.Show , MainForm
'    Else
'        frmShippingInstructions.ShowForm LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
'    End If
End Sub

Private Sub mnuUpdateCustomPrice_Click()
    If LV.ListItems.count = 0 Then Exit Sub
    If MsgBox("Are you Sure, You want to Update Custom Prices?", vbInformation + vbYesNo) = vbNo Then Exit Sub
    
    Dim dCustomPrice As Double
    dCustomPrice = Val(con.Execute("Select SUM(Qty*Price)/SUM(Qty) From CustomInvoiceItems Where CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'").Fields(0).Value & "")
    con.Execute "Update CustomInvoiceItems SET CustomPrice=" & dCustomPrice & " WHERE CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'"
    
End Sub

Private Sub txtDesc_GotFocus()
    cmdRefresh.Default = False
    txtDesc.SelStart = 0
    txtDesc.SelLength = Len(txtDesc)
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
  
    If KeyAscii = 13 Then
    
        If Trim(txtDesc = "") Then
            MsgBox "Invalid Description", vbInformation
            Exit Sub
        End If
    
        con.Execute "Update Vouchers Set [Description]='" & txtDesc & "' Where Sno=" & Val(LV.SelectedItem.key)
        LV.SelectedItem.ListSubItems(3) = txtDesc & txtDesc.Tag
        txtDesc.Visible = False
        cmdRefresh.Default = False
    ElseIf KeyAscii = 27 Then
        txtDesc.Visible = False
        cmdRefresh.Default = False
    End If

End Sub

Private Sub txtDesc_LostFocus()
    txtDesc.Visible = False
End Sub

Private Sub RefreshSnos()
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Text = i
    Next i
End Sub

Private Sub mnuPrintPTC_Click()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Dim f As New frmPrevRpt, strSelection As String
    
    Set rpt = rptApp.OpenReport(RptPath & "\PTCQEL.rpt")
    strSelection = "{VFOrderItemPTC.LotNo}='" & LV.SelectedItem.ListSubItems(8).Tag & "'"
    
    f.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub mnuPrintPurchaseOrder_Click()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\IssSlipWithRcving.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VendIssued.EntryID}=" & LV.SelectedItem.ListSubItems(5).Tag, rpt
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txtEdit) < 0 Or Val(txtEdit) > Val(LV.SelectedItem.SubItems(6)) Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        End If
        LV.SelectedItem.SubItems(7) = Val(txtEdit)
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub

Private Sub txtFind_Change()
    cmdFind.Enabled = Len(Trim(txtFind)) > 0
End Sub


Private Sub txtLotNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call RefreshLV
    End If
End Sub
