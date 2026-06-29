VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmCreateDispatchList 
   ClientHeight    =   8385
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
   ScaleHeight     =   8385
   ScaleMode       =   0  'User
   ScaleWidth      =   10721.69
   WindowState     =   2  'Maximized
   Begin VB.Frame FAcc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   8400
      Left            =   60
      TabIndex        =   0
      Top             =   -45
      Width           =   11865
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   75
         TabIndex        =   8
         Top             =   435
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
      Begin ComboList.Usercontrol1 cmbItems 
         Height          =   285
         Left            =   75
         TabIndex        =   3
         Top             =   945
         Width           =   8970
         _ExtentX        =   15822
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
      Begin VB.TextBox txtOrderNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9075
         TabIndex        =   32
         Top             =   945
         Width           =   975
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9090
         TabIndex        =   31
         Top             =   7935
         Visible         =   0   'False
         Width           =   1335
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6600
         Left            =   75
         TabIndex        =   16
         Top             =   1290
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   11642
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Lots Finalyzed"
         TabPicture(0)   =   "frmCreateDispatchList.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "LV"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtEdit"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "PicAddCartons"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Dispatch List"
         TabPicture(1)   =   "frmCreateDispatchList.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label1(2)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "lblTotalCartons"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Label1(3)"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "lblTotalPcs"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "LVPList"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "txtDispatchListNo"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).ControlCount=   7
         Begin VB.TextBox txtDispatchListNo 
            Height          =   345
            Left            =   1470
            TabIndex        =   26
            Top             =   495
            Width           =   4245
         End
         Begin VB.PictureBox PicAddCartons 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   525
            Left            =   -70680
            ScaleHeight     =   495
            ScaleWidth      =   3540
            TabIndex        =   19
            Top             =   1830
            Visible         =   0   'False
            Width           =   3570
            Begin VB.TextBox txtTotalQty 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2355
               TabIndex        =   22
               Top             =   225
               Width           =   1200
            End
            Begin VB.TextBox txtCartonQty 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1170
               TabIndex        =   21
               Top             =   225
               Width           =   1200
            End
            Begin VB.TextBox txtCartonFrom 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   -15
               TabIndex        =   20
               Top             =   225
               Width           =   1200
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00C5D1DC&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "    Carton #        Qty / Crtn         Total Qty"
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
               Index           =   30
               Left            =   -15
               TabIndex        =   23
               Top             =   -15
               Width           =   3570
            End
         End
         Begin VB.TextBox txtEdit 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   -71865
            TabIndex        =   17
            Top             =   1260
            Visible         =   0   'False
            Width           =   780
         End
         Begin MSComctlLib.ListView LV 
            Height          =   6165
            Left            =   -74940
            TabIndex        =   18
            Top             =   360
            Width           =   11550
            _ExtentX        =   20373
            _ExtentY        =   10874
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
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
            NumItems        =   10
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "S. #"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Maker"
               Object.Width           =   2117
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
               Text            =   "Qty"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "Done"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   8
               Text            =   "Available"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   9
               Text            =   "Lot #"
               Object.Width           =   1764
            EndProperty
         End
         Begin MSComctlLib.ListView LVPList 
            Height          =   5520
            Left            =   90
            TabIndex        =   24
            Top             =   930
            Width           =   11505
            _ExtentX        =   20294
            _ExtentY        =   9737
            SortKey         =   1
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
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "CFrom"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "CTo"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Qty"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Item ID"
               Object.Width           =   4233
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Item Name"
               Object.Width           =   7056
            EndProperty
         End
         Begin VB.Label lblTotalPcs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E7EBEF&
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
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   10080
            TabIndex        =   30
            Top             =   675
            Width           =   1515
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Total Pcs :"
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
            Index           =   3
            Left            =   8745
            TabIndex        =   29
            Top             =   675
            Width           =   1350
         End
         Begin MSForms.TextBox lblTotalCartons 
            Height          =   270
            Left            =   7245
            TabIndex        =   28
            Top             =   675
            Width           =   1515
            VariousPropertyBits=   746604575
            BackColor       =   15199215
            BorderStyle     =   1
            Size            =   "2672;476"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Total Cartons :"
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
            Left            =   5910
            TabIndex        =   27
            Top             =   675
            Width           =   1350
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Dispatch List No. :"
            Height          =   195
            Left            =   105
            TabIndex        =   25
            Top             =   540
            Width           =   1305
         End
      End
      Begin VB.TextBox txtLotNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10065
         TabIndex        =   4
         Top             =   945
         Width           =   825
      End
      Begin VB.Frame Frame1 
         Height          =   480
         Left            =   2640
         TabIndex        =   11
         Top             =   7890
         Width           =   5475
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   1965
            TabIndex        =   14
            Top             =   135
            Width           =   2430
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4395
            TabIndex        =   13
            Top             =   135
            Width           =   1020
         End
         Begin VB.ComboBox cmbfield 
            Height          =   315
            ItemData        =   "frmCreateDispatchList.frx":0038
            Left            =   60
            List            =   "frmCreateDispatchList.frx":0051
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   135
            Width           =   1875
         End
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
         ItemData        =   "frmCreateDispatchList.frx":0098
         Left            =   5265
         List            =   "frmCreateDispatchList.frx":00AE
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   435
         Width           =   1995
      End
      Begin VB.CommandButton cmdTransfer 
         Caption         =   "Save"
         Height          =   345
         Left            =   60
         TabIndex        =   2
         Top             =   7950
         Width           =   2205
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   9075
         TabIndex        =   6
         Top             =   435
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
         Format          =   163512323
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   7260
         TabIndex        =   7
         Top             =   435
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
         Format          =   163512323
         CurrentDate     =   37055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmCreateDispatchList.frx":00FE
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
         TabIndex        =   15
         Top             =   735
         Width           =   10785
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmCreateDispatchList.frx":01DD
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
         TabIndex        =   10
         Top             =   210
         Width           =   11130
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   810
         Left            =   10905
         TabIndex        =   9
         Top             =   435
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
         TabIndex        =   1
         Top             =   7935
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
            Picture         =   "frmCreateDispatchList.frx":0298
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCreateDispatchList.frx":0454
            Key             =   "Up"
         EndProperty
      EndProperty
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
Attribute VB_Name = "frmCreateDispatchList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lDPListEntryID As Long
Dim strDelList As String
Public Sub ShowMe(Optional p_bCustomInvoice As Boolean = True)
    bCustomInvoice = p_bCustomInvoice
    Me.Show
End Sub

Public Sub ShowForShipping()
    bShipping = True
    Me.Show
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


Private Sub cmdPrint_Click()

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\DispatchList.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub



Private Sub cmdTransfer_Click()
    
    If LVPList.ListItems.count = 0 Then Exit Sub
    
    If txtDispatchListNo = "" Then
        MsgBox "Please Enter Dispatch List No.", vbInformation
        Exit Sub
    End If
    '-----------------------------------------------
    Dim i As Integer, lEntryID As Long, lQty As Long, j As Integer
    Call StartTrans(con)
    If lDPListEntryID = 0 Then
        con.Execute "INSERT INTO DispatchList(UserName,MachineName,DT,DispatchListNo) VALUES('" & CurrentUserName & _
         "','" & strComputerName & "','" & GetServerDate & "','" & txtDispatchListNo & "')"
     
        lEntryID = GetSingleLongValue("MAX(EntryID)", "DispatchList", "WHERE MachineName='" & strComputerName & "'")
    Else
        con.Execute "UPDATE DispatchList SET DispatchListNo='" & txtDispatchListNo & "' WHERE EntryID=" & lDPListEntryID
        lEntryID = lDPListEntryID
        If strDelList <> "" Then
            strDelList = Left(strDelList, Len(strDelList) - 1)
            Dim rs As New ADODB.Recordset
            Dim l_Temp_Qty As Long
            With rs
                .Open "SELECT VRD_RefID,CartonFrom,CartonTo,DispatchListDetail_Cartons.Qty FROM DispatchListDetail INNER JOIN DispatchListDetail_Cartons ON DispatchListDetail.EntryID=DispatchListDetail_Cartons.RefID WHERE DispatchListDetail_Cartons.EntryID IN(" & strDelList & ")", con, adOpenForwardOnly, adLockReadOnly
                Do Until .EOF
                    l_Temp_Qty = (Val(!CartonTo & "") - Val(!CartonFrom & "") + 1) * Val(!Qty & "")
                    con.Execute "UPDATE VendRcvdDetail SET IssQty=IssQty-" & l_Temp_Qty & " WHERE EntryID=" & Val(!VRD_RefID & "")
                    .MoveNEXT
                Loop
                .Close
            End With
            Set rs = Nothing
            con.Execute "DELETE FROM DispatchListDetail_Cartons WHERE EntryID IN(" & strDelList & ")"
            con.Execute "DELETE FROM DispatchListDetail WHERE EntryID NOT IN(SELECT RefID FROM DispatchListDetail_Cartons)"
        End If
    End If
    
    Dim lEntryIDDetail As Long
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If Val(.SubItems(7)) > 0 Then
                lQty = Val(.SubItems(7))
                con.Execute "INSERT INTO DispatchListDetail(RefID,VRD_RefID,Qty) VALUES(" & _
                 lEntryID & "," & Val(.Tag) & "," & lQty & ")"
                        
                lEntryIDDetail = GetSingleLongValue("MAX(EntryID)", "DispatchListDetail", " WHERE RefID=" & lEntryID)
                
                'con.Execute "UPDATE Items SET InHand=ISNULL(InHand,0)+" & lQty & " WHERE ItemID='" & .SubItems(5) & "'"
                con.Execute "UPDATE VendRcvdDetail SET IssQty=IssQty+" & lQty & " WHERE EntryID=" & Val(.Tag)
                
            End If
        End With
        For j = 1 To LVPList.ListItems.count
            With LVPList.ListItems(j)
                If Val(LV.ListItems(i).Tag) = Val(.Tag) Then
                    con.Execute "INSERT INTO DispatchListDetail_Cartons(RefID,CartonFrom,CartonTo,Qty) VALUES(" & _
                     lEntryIDDetail & "," & Val(.Text) & "," & Val(.SubItems(1)) & "," & Val(.SubItems(2)) & ")"
                End If
            End With
        Next
    Next
    con.CommitTrans
    
    Unload Me
    'Call RefreshLV
    
End Sub

Private Sub DtFrom_Change()
    'Call setDateRange(DtFrom)
End Sub

Private Sub DtTo_Change()
    'Call setDateRange(DtTo)
End Sub


Private Sub Form_Load()
        
    lDPListEntryID = 0
    mnuPop.Visible = False
    
    If Not (strCompany = "QEL" Or (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai" Or strCompany = "Towne") Then
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
    cmbItems.AddItem "<All Articles>", "0"
    Call cmbItems.AddVals(con, "ItemID + ' ' + ItemName + ' { ' + ISNULL(Attributes,'') + ' }'", "Items", "ItemID", " WHERE InActive=0 ORDER BY ItemName")
    cmbItems.ID = "0"
    
    With cmbfield
        .Clear
        For i = 1 To LV.ColumnHeaders.count
            cmbfield.AddItem LV.ColumnHeaders(i).Text
        Next
    End With
    
    strDelList = ""
    
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
        Cond = Cond & " AND LotNo='" & Trim(txtLotNo) & "' "
    End If
    If Trim(txtOrderNo) <> "" Then
        Cond = Cond & " AND OrderNo='" & Trim(txtOrderNo) & "' "
    End If
    
    Cond = Cond & " AND NextProcessID IS NULL AND (RcvdQty-ISNULL(Wastage,0)-ISNULL(LostQty,0))>IssQty AND ReqAuth=0"
    'Cond = Cond & " AND Left(CustomInvoice,4)='VTL '"
    With rs
        
        .Open "SELECT * FROM VVendReceivingList " & Cond & " ORDER BY DT", con, adOpenForwardOnly, adLockReadOnly
        LVPList.ListItems.Clear
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
            ITM.ListSubItems.add , , !OrderNo & ""
            ITM.ListSubItems.add(, , Format(![DT] & "", "dd-MMM-yyyy")).Tag = ""
            ITM.ListSubItems.add(, , !Description & "").Tag = ""
            ITM.ListSubItems.add(, , ![ItemCode] & "").Tag = Val(!Issuance_RefID & "")
            
            ITM.ListSubItems.add , , Val(!RcvdQty & "") - (Val(!Wastage & "") + Val(!LostQty & "")) - Val(!IssQty & "")
            ITM.ListSubItems.add , , "0"
            ITM.ListSubItems.add , , Val(!RcvdQty & "") - (Val(!Wastage & "") + Val(!LostQty & "")) - Val(!IssQty & "")
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
End Sub


Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
'    With txtEdit
'        .Move LV.Left + LV.ColumnHeaders(8).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(8).Width
'        .Text = Val(LV.SelectedItem.SubItems(7))
'        .Visible = True
'        .SelStart = 0
'        .SelLength = Len(.Text)
'        .SetFocus
'    End With
    
    With PicAddCartons
        .Move LV.Left + LV.Width - .Width - 200, LV.Top + LV.SelectedItem.Top + LV.SelectedItem.Height
        .Visible = True
        If LVPList.ListItems.count > 0 Then
            txtCartonFrom.Text = Val(LVPList.ListItems(LVPList.ListItems.count).SubItems(1)) + 1
            txtCartonQty.Text = Val(LVPList.ListItems(LVPList.ListItems.count).SubItems(2))
        Else
            txtCartonFrom.Text = "1"
        End If
        txtTotalQty.Text = Val(LV.SelectedItem.SubItems(8))
        txtCartonFrom.SetFocus
    End With
        
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
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



Private Sub LVPList_KeyDown(KeyCode As Integer, Shift As Integer)

    If LVPList.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo) = vbYes Then
            If Val(LVPList.SelectedItem.ListSubItems(2).Tag) > 0 Then
                strDelList = strDelList & Val(LVPList.SelectedItem.ListSubItems(2).Tag) & ","
            End If
            LVPList.ListItems.Remove LVPList.SelectedItem.Index
        End If
    End If
    
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

Private Sub txtTotalQty_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Dim dResult As Double
        dResult = Val(txtTotalQty) / Val(txtCartonQty)
        If CLng(dResult) <> dResult Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        End If
        Dim lCrtns As Long
        iQty = Val(txtCartonQty)
        iTotalCartons = CLng(Val(txtTotalQty) / Val(txtCartonQty))
        If iTotalCartons <> Val(txtTotalQty) / Val(txtCartonQty) Then iTotalCartons = iTotalCartons + 1
        iCartonTo = Val(txtCartonFrom) + iTotalCartons - 1
        lCrtns = iTotalCartons 'Val(txtCrtnTo) - Val(txtCrtnFrom) + 1
        
        txtCartonQty = iCartonTo
        
        Set ITM = LVPList.ListItems.add(, , Val(txtCartonFrom))
        ITM.Tag = LV.SelectedItem.Tag
        ITM.ListSubItems.add , , Val(txtCartonQty)
        ITM.ListSubItems.add , , iQty

        LV.SelectedItem.SubItems(7) = LV.SelectedItem.SubItems(7) + (iQty * lCrtns)
        LV.SelectedItem.SubItems(8) = LV.SelectedItem.SubItems(8) - (iQty * lCrtns)
        
        ITM.ListSubItems.add , , LV.SelectedItem.SubItems(5) 'Article No.
        ITM.ListSubItems.add , , GetSingleStringValue("ItemName", "Items", " WHERE ItemID='" & LV.SelectedItem.SubItems(5) & "'") 'LV.SelectedItem.SubItems(3) 'Article Name.
        
        PicAddCartons.Visible = False
        
        Call CalculateTotals
    ElseIf KeyAscii = 27 Then
        PicAddCartons.Visible = False
    End If

End Sub

Private Sub CalculateTotals()


    Dim dTotal As Double
    Dim DPcsTotal As Long
    Dim dblTotalCartons As Long
    Dim Cartons As Long
    Dim PrevCartonNo As Long
    Dim dCurrCartons As Long
    
    With LVPList
        For i = 1 To .ListItems.count
            'dQty = dQty + Val(.TextMatrix(i, 4))
            
            If PrevCartonNo <> Val(.ListItems(i)) Then
                Cartons = Val(.ListItems(i).SubItems(1) & "") - Val(.ListItems(i)) + 1
            End If
            dCurrCartons = Val(.ListItems(i).SubItems(1) & "") - Val(.ListItems(i)) + 1
            
            Pcs = Pcs + Val(.ListItems(i).SubItems(2) & "")
            
            dblTotalCartons = dblTotalCartons + Cartons
            DPcsTotal = DPcsTotal + (Pcs * dCurrCartons)
            
            dTotal = dTotal + (Val(Pcs * dCurrCartons) * Val(.ListItems(i).ListSubItems(4).Tag))
            Pcs = 0
            Cartons = 0
            PrevCartonNo = Val(.ListItems(i).SubItems(1) & "")
        Next i
        
    End With
    
    'txttQty = dQty
    'lblTotalAmount = DTotal
    'lblTotWords = ConvertInWords(CCur(DTotal), txtCurrency.Text)
    
    lblTotalPcs = DPcsTotal
    lblTotalCartons = dblTotalCartons
    
End Sub

Public Sub EditDispatchList(p_lEntryID As Long)

    lDPListEntryID = p_lEntryID
    txtDispatchListNo = GetSingleStringValue("DispatchListNo", "DispatchList", " WHERE EntryID=" & lDPListEntryID)
    
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT CartonFrom,CartonTo,DispatchListDetail_Cartons.Qty,DispatchListDetail_Cartons.EntryID,DispatchListDetail_Cartons.RefID," & _
         "ItemID,ItemName FROM DispatchList INNER JOIN DispatchListDetail ON DispatchList.EntryID=DispatchListDetail.RefID " & _
         "INNER JOIN DispatchListDetail_Cartons ON DispatchListDetail.EntryID=DispatchListDetail_Cartons.RefID INNER JOIN " & _
         "VendRcvdDetail ON DispatchListDetail.VRD_RefID=VendRcvdDetail.EntryID INNER JOIN Items " & _
         "ON VendRcvdDetail.ItemCode=Items.ItemID WHERE DispatchList.EntryID=" & lDPListEntryID & _
         " ORDER BY DispatchListDetail.EntryID,DispatchListDetail_Cartons.EntryID", con, adOpenForwardOnly, adLockReadOnly
         
         LVPList.ListItems.Clear
         Do Until .EOF
            Set ITM = LVPList.ListItems.add(, , Val(!CartonFrom & ""))
            ITM.Tag = "0"
            ITM.ListSubItems.add(, , Val(!CartonTo & "")).Tag = Val(!RefID & "")
            ITM.ListSubItems.add(, , Val(!Qty & "")).Tag = Val(!EntryID & "")
            ITM.ListSubItems.add , , !ITemID & ""
            ITM.ListSubItems.add , , !ItemName & ""
            .MoveNEXT
         Loop
         .Close
    End With
    Set rs = Nothing
    
    Me.Show
    
End Sub
