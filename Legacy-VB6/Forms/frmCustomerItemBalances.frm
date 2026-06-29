VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmCustomerItemBalances 
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   14385
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
   ScaleHeight     =   9885
   ScaleMode       =   0  'User
   ScaleWidth      =   12884.84
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
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
      Left            =   15
      TabIndex        =   4
      Top             =   0
      Width           =   14325
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Balances"
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
         TabIndex        =   5
         Top             =   165
         Width           =   14175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Balances"
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
         TabIndex        =   6
         Top             =   180
         Width           =   14160
      End
   End
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
      Height          =   9225
      Left            =   15
      TabIndex        =   0
      Top             =   615
      Width           =   14325
      Begin VB.TextBox txtEdit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1965
         TabIndex        =   33
         Top             =   2460
         Visible         =   0   'False
         Width           =   795
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   45
         TabIndex        =   2
         Top             =   345
         Width           =   13170
         _ExtentX        =   23230
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
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10260
         TabIndex        =   32
         Top             =   330
         Visible         =   0   'False
         Width           =   1395
      End
      Begin ComboList.Usercontrol1 cmbItemGroup 
         Height          =   285
         Left            =   7965
         TabIndex        =   28
         Top             =   330
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
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
      Begin VB.CheckBox chkUrgent 
         Caption         =   "Urgent"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   12375
         TabIndex        =   26
         Top             =   375
         Visible         =   0   'False
         Width           =   885
      End
      Begin MSComCtl2.DTPicker DTDelivery 
         Height          =   300
         Left            =   3615
         TabIndex        =   23
         Top             =   1980
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
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
         CustomFormat    =   "dd-MMM-yy"
         Format          =   169869315
         CurrentDate     =   40642
      End
      Begin VB.CheckBox chkBoth 
         Caption         =   "Both"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   11700
         TabIndex        =   22
         Top             =   375
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.CheckBox chkNoDelivery 
         Caption         =   "No Delivery Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   11700
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   45
         TabIndex        =   15
         Top             =   8265
         Width           =   5475
         Begin VB.TextBox txtFind 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1965
            TabIndex        =   18
            Top             =   135
            Width           =   2430
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4395
            TabIndex        =   17
            Top             =   135
            Width           =   1020
         End
         Begin VB.ComboBox cmbfield 
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
            ItemData        =   "frmCustomerItemBalances.frx":0000
            Left            =   60
            List            =   "frmCustomerItemBalances.frx":0016
            Style           =   2  'Dropdown List
            TabIndex        =   16
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
         ItemData        =   "frmCustomerItemBalances.frx":005C
         Left            =   3150
         List            =   "frmCustomerItemBalances.frx":006C
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   330
         Visible         =   0   'False
         Width           =   1725
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   6435
         TabIndex        =   13
         Top             =   330
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   169869315
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   4875
         TabIndex        =   14
         Top             =   330
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   169869315
         CurrentDate     =   37055
      End
      Begin VB.TextBox txtTotalAmt 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5595
         TabIndex        =   11
         Top             =   8355
         Visible         =   0   'False
         Width           =   1785
      End
      Begin MSComctlLib.ListView LV 
         Height          =   7605
         Left            =   45
         TabIndex        =   7
         Top             =   660
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   13414
         View            =   3
         LabelEdit       =   1
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Order #"
            Object.Width           =   2088
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Int. Ord"
            Object.Width           =   1799
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Item Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Item Name"
            Object.Width           =   11818
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Delivery DT"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Bal. Qty"
            Object.Width           =   1508
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Delivery"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSForms.CommandButton cmdPrintUrgent 
         Height          =   375
         Left            =   10560
         TabIndex        =   34
         Top             =   8775
         Width           =   1200
         ForeColor       =   0
         Caption         =   "Print Urgent"
         PicturePosition =   327683
         Size            =   "2117;661"
         Accelerator     =   85
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrintListWithValue 
         Height          =   375
         Left            =   3735
         TabIndex        =   31
         Top             =   8775
         Visible         =   0   'False
         Width           =   1200
         ForeColor       =   0
         Caption         =   "P Value"
         PicturePosition =   327683
         Size            =   "2117;661"
         Accelerator     =   86
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrintR2 
         Height          =   375
         Left            =   4965
         TabIndex        =   30
         Top             =   8775
         Visible         =   0   'False
         Width           =   1050
         ForeColor       =   0
         Caption         =   "R2"
         PicturePosition =   327683
         Size            =   "1852;661"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrintR1 
         Height          =   375
         Left            =   6075
         TabIndex        =   29
         Top             =   8775
         Visible         =   0   'False
         Width           =   1050
         ForeColor       =   0
         Caption         =   "R1"
         PicturePosition =   327683
         Size            =   "1852;661"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrintOrderStatus 
         Height          =   375
         Left            =   7155
         TabIndex        =   27
         Top             =   8775
         Visible         =   0   'False
         Width           =   1380
         ForeColor       =   0
         Caption         =   "Order Status"
         PicturePosition =   327683
         Size            =   "2434;661"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrintItemBalance 
         Height          =   375
         Left            =   1950
         TabIndex        =   25
         Top             =   8775
         Visible         =   0   'False
         Width           =   1575
         ForeColor       =   0
         Caption         =   "Balance Report"
         PicturePosition =   327683
         Size            =   "2778;661"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPPRSelected 
         Height          =   375
         Left            =   8580
         TabIndex        =   24
         Top             =   8775
         Visible         =   0   'False
         Width           =   1575
         ForeColor       =   0
         Caption         =   "Selected PPR"
         PicturePosition =   327683
         Size            =   "2778;661"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrintPlanningReport 
         Height          =   375
         Left            =   8580
         TabIndex        =   20
         Top             =   8775
         Visible         =   0   'False
         Width           =   1575
         ForeColor       =   0
         Caption         =   "Planning Report"
         PicturePosition =   327683
         Size            =   "2778;661"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   375
         Left            =   11805
         TabIndex        =   19
         Top             =   8775
         Width           =   1200
         ForeColor       =   0
         Caption         =   "Print"
         PicturePosition =   327683
         Size            =   "2117;661"
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Amount (Rs.):"
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
         Left            =   7455
         TabIndex        =   10
         Top             =   8400
         Visible         =   0   'False
         Width           =   2130
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   13305
         TabIndex        =   9
         Top             =   225
         Width           =   915
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "1614;635"
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
         Left            =   13050
         TabIndex        =   8
         Top             =   8775
         Width           =   1200
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "2117;661"
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   1485
         TabIndex        =   3
         Top             =   345
         Visible         =   0   'False
         Width           =   1650
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2910;503"
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
         BackColor       =   &H00AFBDCB&
         Caption         =   "  Customer            "
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
         TabIndex        =   1
         Top             =   120
         Width           =   13140
      End
   End
End
Attribute VB_Name = "frmCustomerItemBalances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SQLString As String, SQLStringForR1R2 As String
Dim bForProduction As Boolean
Dim strReportSQL As String
Dim bAuthorized As Boolean
Dim iColNo As Integer
Private Sub mnuProdPlan_Click()
    Load frmProdPlan
    Call frmProdPlan.AddPlanning(Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1))
End Sub

Public Sub ShowMe(p_bForProduction As Boolean)
    bForProduction = p_bForProduction
    If bForProduction Then
        Label2.Visible = False
        txtTotalAmt.Visible = False
    End If
    Me.Show
    Call RefreshLV
End Sub

Private Sub cmbCust_matched()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    SQL = "Select Distinct Country From ForeignCustomers "
    If cmbCust.ID <> "0" Then
        SQL = SQL & " Where CustCode='" & cmbCust.ID & "'"
    End If
    
    With rs
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        cmbCountry.AddItem "<All Countries>"
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            'cmbCountry.List(cmbCountry.ListCount - 1, 1) = ![Curr] & ""
            .MoveNEXT
        Loop
        .Close
    End With
    cmbCountry.ListIndex = 0
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbRange_Click()

    DtTo = ServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
        Case 0
            DtTo = DateAdd("d", 15, ServerDate)
        Case 1
            DtTo = DateAdd("d", 30, ServerDate)
        Case 2
            DtTo = DateAdd("d", 45, ServerDate)
        Case 3
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
       
        fldValue = LV.ListItems(i).SubItems(cmbfield.ListIndex + 1)
       
        If LCase(fldValue) Like "*" & LCase(Trim(txtFind)) & "*" Then
            LV.ListItems(i).Selected = True
            LV.ListItems(i).EnsureVisible
            If i <> StartIdx Then Exit For
        End If
       
    Next
    Dim rpt As New CRAXDDRT.Report
    
    If i = LV.ListItems.count + 1 Then
        MsgBox "Record Not Found.", vbInformation
        Set LV.SelectedItem = Nothing
    End If

End Sub

Private Sub cmdPPRSelected_Click()
    
    Dim strItems As String
    Dim i As Integer
    strItems = ""
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            strItems = strItems & "'" & LV.ListItems(i).SubItems(5) & "',"
        End If
    Next
    
    If strItems <> "" Then
        strItems = Left(strItems, Len(strItems) - 1)
    Else
        MsgBox "No Item Selected.", vbInformation
        Exit Sub
    End If
    
    If LV.SelectedItem Is Nothing Then
        MsgBox "Please Select Item.", vbInformation
        Exit Sub
    End If
    
    Dim rpt As CRAXDDRT.Report
    
    Set rpt = rptApp.OpenReport(RptPath & "\OrderPlanningReport.rpt")
    
    'rpt.ParameterFields.GetItemByName("@OrderNo").SetCurrentValue strOrderNo
    rpt.ParameterFields.GetItemByName("@FromItems").SetCurrentValue True
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom.Value
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo.Value
    rpt.ParameterFields.GetItemByName("@NoDelivery").SetCurrentValue IIf(chkNoDelivery.Value = vbChecked, True, False)
    rpt.ParameterFields.GetItemByName("@Both").SetCurrentValue IIf(chkBoth.Value = vbChecked, True, False)
    
    If cmbCust.ID <> "0" Then
        rpt.ParameterFields.GetItemByName("@CustCode").SetCurrentValue cmbCust.ID
    End If
    
    If cmbCountry.ListIndex > 0 Then
        rpt.ParameterFields.GetItemByName("@Country").SetCurrentValue cmbCountry.Text
    End If
    
    rpt.ParameterFields.GetItemByName("@ItemCode").SetCurrentValue strItems
    rpt.EnableParameterPrompting = False
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, , False
    
End Sub

Private Sub cmdPrint_Click()

    Call PrintList
    
End Sub

Private Sub PrintList(Optional bUrgentOnly As Boolean = False)

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\OrderItemBalances.rpt")
    Dim f As New frmPrevRpt
    Dim strSelection As String
     
    Dim strSQL As String
    strSQL = SQLString
    
    If bUrgentOnly Then
        strSQL = strSQL & " AND DeliveryStatus=2"
    End If
    
    strSQL = strSQL & " ORDER BY VFOrderItems.DeliveryDT DESC"
'    If bUrgentOnly Then
'        strSelection = "{VFOrderItems.DeliveryStatus}=2"
'    End If
    
    f.ShowReport "", rpt, strSQL
    
End Sub

Private Sub cmdPrintItemBalance_Click()
    
    Dim strItems As String
    Dim i As Integer
    strItems = ""
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            strItems = strItems & "'" & LV.ListItems(i).SubItems(5) & "',"
        End If
    Next
    If strItems <> "" Then
        strItems = Left(strItems, Len(strItems) - 1)
    Else
        MsgBox "No Item Selected.", vbInformation
        Exit Sub
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptPartiallyShippedOrdersItemwise.rpt")
    Dim strSelection As String
    strSelection = "{FOrderItems.CompItemCode} IN[" & strItems & "] AND {VFOrderItemswithShippedQty.Qty}>{VFOrderItemswithShippedQty.ShippedQty}"
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub cmdPrintListWithValue_Click()
    If getDBPassword("ViewPValueReport") = False Then
        Exit Sub
    End If
    Call PrintList(True)
End Sub

Private Sub cmdPrintOrderStatus_Click()
    
    Dim i As Integer, strIDs As String
    For i = 1 To LV.ListItems.count
        strIDs = strIDs & Val(LV.ListItems(i).key) & ","
    Next
    
    If strIDs <> "" Then
        strIDs = Left(strIDs, Len(strIDs) - 1)
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\OrderStatusReport.rpt")
    rpt.ParameterFields.GetItemByName("@FOI_IDs").SetCurrentValue strIDs
    
    rpt.EnableParameterPrompting = False
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, "", False
    
End Sub

Private Sub cmdPrintPlanningReport_Click()

    Dim rpt As CRAXDDRT.Report
    
    Set rpt = rptApp.OpenReport(RptPath & "\OrderPlanningReport.rpt")
    
    'rpt.ParameterFields.GetItemByName("@OrderNo").SetCurrentValue strOrderNo
    rpt.ParameterFields.GetItemByName("@FromItems").SetCurrentValue True
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom.Value
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo.Value
    rpt.ParameterFields.GetItemByName("@NoDelivery").SetCurrentValue IIf(chkNoDelivery.Value = vbChecked, True, False)
    rpt.ParameterFields.GetItemByName("@Both").SetCurrentValue IIf(chkBoth.Value = vbChecked, True, False)
    rpt.ParameterFields.GetItemByName("@Urgent").SetCurrentValue IIf(chkUrgent.Value = vbChecked, True, False)
    rpt.ParameterFields.GetItemByName("@ItemCode").SetCurrentValue "0"
    
    rpt.ParameterFields.GetItemByName("@GroupID").SetCurrentValue Val(cmbItemGroup.ID)
    
    If cmbCust.ID <> "0" Then
        rpt.ParameterFields.GetItemByName("@CustCode").SetCurrentValue cmbCust.ID
    End If
    If cmbCountry.ListIndex > 0 Then
        rpt.ParameterFields.GetItemByName("@Country").SetCurrentValue cmbCountry.Text
    End If
    
    rpt.EnableParameterPrompting = False
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "", rpt, , False
    
End Sub

Private Sub PrintR1R2PlanningReport(Optional bR1 As Boolean)

    On Error GoTo err
    'Prepare Data for the Report...
    
    Dim strCondition As String
    strCondition = getCondition()
    If LV.SelectedItem Is Nothing Then
        MsgBox "Please Select Item.", vbInformation
        Exit Sub
    End If
     
    Dim cmd As New ADODB.Command
    With cmd
        Set .ActiveConnection = con_ServerSide
        .CommandType = adCmdStoredProc
        .CommandTimeout = 300
        .CommandText = "PreparePlanningReport_SP"
        .Parameters("@OrderNo").Value = "0"
        .Parameters("@R1").Value = bR1
        .Parameters("@ItemCodes").Value = strCondition 'strItems
        .Execute
    End With
     
    Dim rpt As CRAXDDRT.Report
    If bR1 Then
        Set rpt = rptApp.OpenReport(RptPath & "\OrderPlanningR1.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\OrderPlanningR2.rpt")
    End If
    rpt.FormulaFields.GetItemByName("OrderWise").Text = False
    rpt.FormulaFields.GetItemByName("OrderNo").Text = "'0'"
    Dim f As New frmPrevRpt
    Dim strConditionReport As String
    strConditionReport = getConditionitionForReport
    f.ShowReport strConditionReport, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdPrintR1_Click()
    Call PrintR1R2PlanningReport(True)
End Sub

Private Sub cmdPrintR2_Click()
    Call PrintR1R2PlanningReport(False)
End Sub

Private Sub cmdPrintUrgent_Click()
    Call PrintUrgentItems
    'Call PrintList(True)
End Sub

Private Sub PrintUrgentItems()
    
    Dim strCondition As String
    strCondition = "{VrptOrders.CustCode}='" & cmbCust.ID & "'"
    strCondition = strCondition & " AND {VrptOrders.DeliveryStatus}=2"
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\UrgentRequest.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport strCondition, rpt
    

End Sub
Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub


Private Sub DTDelivery_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
    
        Dim lID As Long, iStatus As Integer
        lID = Val(LV.SelectedItem.key)
        iStatus = Val(LV.SelectedItem.ListSubItems(7).Tag)
        Call StartTrans(con)
        
        'con.Execute "UPDATE FOrderItems SET DeliveryStatus=1,DeliveryDT='" & DTDelivery.Value & "' WHERE ID=" & lID
        con.Execute "UPDATE FOrderItems SET DeliveryDT='" & DTDelivery.Value & "' WHERE ID=" & lID
        con.Execute "INSERT INTO FOrderItems_DeliveryRevisions(FOI_RefID,DeliveryStatus,DeliveryDT,UserName,MachineName)" & _
         "VALUES(" & lID & "," & iStatus & ",'" & DTDelivery.Value & "','" & CurrentUserName & "','" & strComputerName & _
         "')"
        
        If strCompany = "Dr-Frgz" And strSubCompany = "Intl" Or strCompany = "Towne" Then
            con.Execute "INSERT INTO FOrderItems_DeliveryDates_PTW(FOI_RefID) VALUES(" & lID & ")"
        End If
        
        con.CommitTrans
        
        DTDelivery.Visible = False
        Call RefreshLV
        
    ElseIf KeyCode = 27 Then
        DTDelivery.Visible = False
    End If
    
End Sub

Private Sub DTDelivery_LostFocus()
    DTDelivery.Visible = False
End Sub

Private Sub Form_Load()
    
    'mnuPop.Visible = False
    bForProduction = False
    
    cmbCust.ListHeight = 2500
    'cmbCust.AddItem "<All Customers>", "0"
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode")
    
    cmbCust.ID = "0"
'    cmbCountry.ListIndex = 0
'
'    cmbItemGroup.AddItem "<All Groups>", "0"
'    cmbItemGroup.AddVals con, "Description", "ItemGroups", "ID", " ORDER BY Description"
'    cmbItemGroup.ID = "0"
'
'    Call RefreshLV
'
'    DtFrom.Value = Date
'    DtTo.Value = Date
'
'    cmbRange.ListIndex = 0
    cmbfield.ListIndex = 2
'
'    bAuthorized = False
'
'    If Not (strCompany = "Dr-Frgz") Then
'        cmdPrintR1.Visible = False
'        cmdPrintR2.Visible = False
'        cmdPrintOrderStatus.Visible = False
'    End If
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Dim TheHeight As Long
    Frame1.Left = (Me.ScaleWidth - Frame1.Width) / 2
    
    
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    
End Sub


Private Sub RefreshLV()
    
    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim Cond As String
    Dim DTRange As String
    Dim ITM As ListItem
    
    Screen.MousePointer = vbHourglass
    
    Dim strCondition As String
    strCondition = getCondition()
    
    SQLString = "SELECT VFOrderItems.* FROM VFOrderItems " & strCondition
    
    strReportSQL = "SELECT VFOrderItems.* FROM VFOrderItems " & strCondition & " ORDER BY VFOrderItems.DeliveryDT DESC"
    
    Dim sStatus As String
    Dim iRecNo As Integer
    Dim dTotalAmt As Double, dTotalAmtPKR As Double
    Dim strRemarks As String
    dTotalAmt = 0
    
    Dim strDeliveryStatus As String, lForeColor As Long, i As Integer
    With rs
        
        .Open SQLString & " ORDER BY VFOrderItems.DeliveryDT DESC", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        iRecNo = .RecordCount
        'LV.Sorted = False
        
        Do Until .EOF
            
            Set ITM = LV.ListItems.add(, !ID & "'", iRecNo - .AbsolutePosition + 1)
            
            ITM.ListSubItems.add , , ![OrderNo] & ""
            ITM.ListSubItems.add , , ![InternalRefNo] & ""
            ITM.ListSubItems.add , , ![CompItemCode] & ""
            ITM.ListSubItems.add , , ![ItemName] & ""
            ITM.ListSubItems.add(, , IIf(!DeliveryStatus = 0, "-", Format(![DeliveryDT] & "", "dd-MMM-yyyy"))).Tag = Val(!DeliveryStatus & "")
            
            ITM.ListSubItems.add(, , Val(![Qty] & "") - Val(!ShippedQty & "")).Tag = Val(![Qty] & "") - Val(!ShippedQty & "")
            
            'ITM.ListSubItems.add , , IIf(!DeliveryStatus = 0, "-", DateDiff("d", !DeliveryDT, Date))
            
            If Val(!DeliveryStatus & "") = 0 Then
                strDeliveryStatus = "No Date"
                lForeColor = LV.ForeColor
            ElseIf Val(!DeliveryStatus & "") = 1 Then
                strDeliveryStatus = "Regular"
                lForeColor = LV.ForeColor
            ElseIf Val(!DeliveryStatus & "") = 2 Then
                strDeliveryStatus = "Urgent"
                lForeColor = vbRed
            End If
            
            ITM.ListSubItems.add(, , strDeliveryStatus).Tag = Val(!DeliveryStatus & "")
            
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ForeColor = lForeColor
            Next
            .MoveNEXT
            
        Loop
        
        .Close
        
    End With
    
    Set rs = Nothing
    'txtTotalAmt = Format(dTotalAmtPKR, "###,##0.00")
    
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Function getCondition() As String

    Dim Cond As String
   
    
    Cond = " WHERE CustCode='" & cmbCust.ID & "'"
    
'
'    If cmbCountry.ListIndex <> 0 Then
'        Cond = Cond & " AND (VFOrderItems.Country='" & cmbCountry & "')"
'    End If
'
'    If cmbItemGroup.ID <> "0" Then
'        Cond = Cond & " AND (GroupID=" & cmbItemGroup.ID & ")"
'    End If
    
    Cond = Cond & " AND (Qty>ShippedQty)"
    
    Cond = Cond & " AND VFOrderItems.OrderNo IN (SELECT OrderNo FROM VUnshippedOrderList)"
    
    Cond = Cond & " AND VFOrderItems.CustCode NOT IN('Stock')"
    
'    If Trim(txtOrderNo) <> "" Then
'        Cond = Cond & " AND VFOrderItems.OrderNo='" & txtOrderNo & "'"
'    End If
    
    getCondition = Cond
    
End Function

Private Function getConditionitionForReport() As String

    Dim strstrConditionition As String
    
    If chkUrgent.Value = vbChecked Then
        strCondition = " ({VFOrderItems.DeliveryStatus}=2 OR {VFOrderItems.Remarks}<>'')"
    Else
        If chkBoth.Value = vbChecked Then
            strCondition = " ({VFOrderItems.DeliveryStatus}=0 OR ({VFOrderItems.DeliveryDT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#))"
        Else
            If chkNoDelivery.Value = vbChecked Then
                strCondition = " {VFOrderItems.DeliveryStatus}=0"
            Else
                strCondition = " {VFOrderItems.DeliveryDT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#"
            End If
        End If
    End If
    
    If cmbCust.ID <> "0" Then
        strCondition = strCondition & " AND ({VFOrderItems.CustCode}='" & cmbCust.ID & "') "
    End If
    
    If cmbCountry.ListIndex <> 0 Then
        strCondition = strCondition & " AND ({VFOrderItems.Country}='" & cmbCountry & "')"
    End If
    
    If cmbItemGroup.ID <> "0" Then
        strCondition = strCondition & " AND ({VFOrderItems.GroupID}=" & cmbItemGroup.ID & ")"
    End If
    
    strCondition = strCondition & " AND ({VFOrderItems.Qty}>{VFOrderItems.ShippedQty})"
    
    'strCondition = strCondition & " AND {VFOrderItems.OrderNo} IN (SELECT OrderNo FROM VUnshippedOrderList)"
    
    getConditionitionForReport = strCondition
    
End Function


Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If LV.SortKey + 1 = ColumnHeader.Index Then
        LV.ColumnHeaderIcons = Nothing
        'Set LV.ColumnHeaderIcons = ImageList2
        If LV.SortOrder = lvwAscending Then
            LV.SortOrder = lvwDescending
            'ColumnHeader.Icon = "Down"
        Else
            LV.SortOrder = lvwAscending
            'ColumnHeader.Icon = "Up"
        End If
    Else
        LV.ColumnHeaderIcons = Nothing
        'Set LV.ColumnHeaderIcons = ImageList2
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        'ColumnHeader.Icon = "Down"
    End If
    Call RefreshSnos
    If LV.Sorted = False Then LV.Sorted = True
    
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
            
'    If bAuthorized = False Then
'        If getDBPassword("EditDeliveryDT") = False Then
'            Exit Sub
'        Else
'            bAuthorized = True
'        End If
'    End If
'
'    If Val(LV.SelectedItem.ListSubItems(7).Tag) = 0 Then 'No Delivery
'        DTDelivery.Value = Date
'    Else
'        DTDelivery.Value = LV.SelectedItem.ListSubItems(7).Text
'    End If
'

    If iColNo = 8 Then
        Dim iStatus As Integer, strStatus As String, lForeColor As Long
        
        If Val(LV.SelectedItem.ListSubItems(7).Tag) = 2 Then        'If Urgent, Change to Normal
            strStatus = "Regular"
            iStatus = 1
            lForeColor = LV.ForeColor
        Else
            strStatus = "Urgent"
            iStatus = 2
            lForeColor = vbRed
        End If
        LV.SelectedItem.ListSubItems(7).Tag = iStatus
        LV.SelectedItem.ListSubItems(7).Text = strStatus
        con.Execute "UPDATE FOrderItems SET DeliveryStatus=" & iStatus & " WHERE ID=" & Val(LV.SelectedItem.key)
        
        
        Dim i As Integer
        For i = 1 To LV.SelectedItem.ListSubItems.count
            LV.SelectedItem.ListSubItems(i).ForeColor = lForeColor
        Next
    Else
        With txtEdit
            .Move LV.Left + LV.ColumnHeaders(7).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(7).Width
            .Text = Val(LV.SelectedItem.SubItems(6))
            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
        End With
    End If
    
    
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If X > Val(LV.ColumnHeaders(i).Left) And X < (Val(LV.ColumnHeaders(i).Left) + Val(LV.ColumnHeaders(i).Width)) Then
            iColNo = i
            Exit For
        End If
    Next
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
    Exit Sub
    If LV.ListItems.count = 0 Or Button <> 2 Then Exit Sub
 
    Dim strRet As String
'    If UserHasAccess("AddProdPlan") Then
'        strRet = ShowPopUpMenu1("&Edit This Order", "-", "&Add Production Planning", "-", "History Of The Order (Summary)", "History Of The Order (Detail)", "-", "Print Order (&Summary)", "Print Order (&Detail)", "-", "&Print Order", "-", "Print Order &Status")
'    Else
'        strRet = ShowPopUpMenu("&Edit This Order", "-", "History Of The Order (Summary)", "History Of The Order (Detail)", "-", "Print Order (&Summary)", "Print Order (&Detail)", "-", "&Print Order", "-", "Print Order &Status")
'    End If
    
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(OrderNo)", "FCustomerFinalOrders", " WHERE OrderNo='" & LV.SelectedItem.SubItems(3) & "'")
    
        
    If strCompany = "Sunlike" And bForProduction Then
        If lCount > 0 Then
            strRet = ShowPopUpMenu1("&Edit This Order", "-", "&Print Order", "-", "Print Order List", "-", "Print Balance Order", "Print PTC", "-", "Material Consumption Sheet", "Material Consumption Sheet (Articlewise)", "Un-Cancel Order")
        Else
            strRet = ShowPopUpMenu1("&Edit This Order", "-", "&Print Order", "-", "Print Order List", "-", "Print Balance Order", "Print PTC", "-", "Material Consumption Sheet", "Material Consumption Sheet (Articlewise)", "Mark As Shipped", "Cancel Order")
        End If
    ElseIf strCompany = "QEL" Then
        If lCount > 0 Then
            strRet = ShowPopUpMenu1("&Edit This Order", "-", "&Print Order", "-", "Print Order List", "-", "Print Order (Balance Quantities)", "-", "Print Balance Order", "Print PTC", "-", "Material Consumption Sheet", "Material Consumption Sheet (Articlewise)", "Print Customer Order", "-", "Stock Order Adjustment", "Un-Cancel Order")
        Else
            strRet = ShowPopUpMenu1("&Edit This Order", "-", "&Print Order", "-", "Print Order List", "-", "Print Order (Balance Quantities)", "-", "Print Balance Order", "Print PTC", "-", "Material Consumption Sheet", "Material Consumption Sheet (Articlewise)", "Print Customer Order", "-", "Stock Order Adjustment", "Mark As Shipped", "Cancel Order")
        End If
    Else
        If lCount > 0 Then
            strRet = ShowPopUpMenu1("&Edit This Order", "-", "&Print Order", "-", "Print Order List", "Print Order &Makerwise", "-", "Print Order (Balance Quantities)", "-", "Print Balance Order", "Print PTC", "-", "Material Consumption Sheet", "Material Consumption Sheet (Articlewise)", "Print Customer Order", "Print Order Planning Report", "-", "Stock Order Adjustment", "Un-Cancel Order")
        Else
            strRet = ShowPopUpMenu1("&Edit This Order", "-", "&Print Order", "-", "Print Order List", "Print Order &Makerwise", "-", "Print Order (Balance Quantities)", "-", "Print Balance Order", "Print PTC", "-", "Material Consumption Sheet", "Material Consumption Sheet (Articlewise)", "Print Customer Order", "Print Order Planning Report", "-", "Stock Order Adjustment", "Mark As Shipped", "Cancel Order")
        End If
    End If
    
    Select Case strRet
        Case "&Edit This Order"
            Call EditOrder
        Case "&Add Production Planning"
            Call AddProdPlanning
        Case "History Of The Order (Summary)"
            
        Case "History Of The Order (Detail)"
            
        Case "Print Order (&Summary)"
            Call PrintSummary
        Case "Print Order List"
            Call PrintOrderList
        Case "Print Order (&Detail)"
        Case "Print Order (Balance Quantities)"
            Call PrintOrderBalanceReport(LV.SelectedItem.SubItems(3))
        Case "&Print Order"
            Call PrintOrder
        Case "Print Order &Makerwise"
            Call PrintOrder(False, True)
        Case "Print Balance Order"
            Call PrintOrder(True)
        Case "Print PTC"
            Call PrintPTC
        Case "Material Consumption Sheet"
            Call PrintMaterialConsumptionSheet(LV.SelectedItem.SubItems(4))
        Case "Material Consumption Sheet (Articlewise)"
            Call PrintMaterialConsumptionSheet(LV.SelectedItem.SubItems(4), True)
        Case "Print Customer Order"
            Call PrintCustomerOrder
        Case "Stock Order Adjustment"
            Set f = New frmStockOrderAdjustment
            Load f
            f.ShowMe (LV.SelectedItem.SubItems(4))
        Case "Mark As Shipped", "Cancel Order", "Un-Cancel Order"

            If MsgBox("Are you sure ?", vbQuestion + vbYesNo) = vbYes Then
                
                If strRet = "Un-Cancel Order" Then
                    con.Execute "DELETE FROM FCustomerFinalOrders WHERE OrderNo='" & LV.SelectedItem.SubItems(3) & "'"
                Else
                    Dim iCancelled As Integer
                    If strRet = "Cancel Order" Then
                        iCancelled = 1
                    Else
                        iCancelled = 0
                    End If
                    con.Execute "INSERT INTO FCustomerFinalOrders(CustCode,Country,OrderNo,UserName,MachineName,Cancelled) VALUES('" & _
                     LV.SelectedItem.SubItems(2) & "','" & LV.SelectedItem.SubItems(3) & "','" & LV.SelectedItem.SubItems(4) & _
                     "','" & CurrentUserName & "','" & strComputerName & "'," & iCancelled & ")"
                End If
            End If
        Case "Print Order Planning Report"
            Call PrintOrderPlanningReport(LV.SelectedItem.SubItems(4))
    End Select
    'Me.PopupMenu mnuPop
    'If LV.ListItems.Count = 0 Or Button <> 2 Or OutOfFinancialYear Then Exit Sub

    'if a Voucher Entry Show Delete Menu
    'If InStr(LV.SelectedItem, "-") > 0 Then Me.PopupMenu mnuPop
    
End Sub

Private Sub PrintOrderList()

    On Error GoTo err
    Me.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptFOrderList.rpt")
        
    rpt.FormulaFields(3).Text = "'" & cmbCust.Text & " " & cmbCountry.Text & "'"
    
    frmPrevRpt.ShowReport " ", rpt, SQLString
    
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
End Sub
Private Sub PrintPTC()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Dim f As New frmPrevRpt, strSelection As String
    If strCompany = "QEL" Then
        Set rpt = rptApp.OpenReport(RptPath & "\PTCQEL.rpt")
        strSelection = "{VFOrderItemPTC.OrderNo}='" & LV.SelectedItem.SubItems(4) & "'"
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\PTC.rpt")
        strSelection = "{FCustomerOrders.OrderNo}='" & LV.SelectedItem.SubItems(4) & "'"
    End If
    f.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub PrintOrderBalanceReportOld()

    On Error GoTo err
    Me.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report, f As New frmPrevRpt
    Set rpt = rptApp.OpenReport(RptPath & "\rptPartiallyShippedOrders.rpt")
    'f.ShowReport "{FCustomerOrders.OrderNo}='" & LV.SelectedItem.SubItems(4) & "' AND {VFOrderItemswithShippedQty.Qty}>{VFOrderItemswithShippedQty.ShippedQty}", rpt
    f.ShowReport "{FCustomerOrders.OrderNo}='" & LV.SelectedItem.SubItems(4) & "'", rpt
    Me.MousePointer = vbNormal
    
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub PrintOrderStatus()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report 'New rptCustOrderStatus
    Set rpt = rptApp.OpenReport(RptPath & "\rptCustOrderStatus.rpt") '
    Dim f As New frmPrevRpt
    f.ShowReport "{VFCustOrders.OrderNo}='" & LV.SelectedItem.SubItems(3) & "' ", rpt
        
    Set rpt = Nothing
    Set f = Nothing
    
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub
Private Sub AddProdPlanning()
    
    Load frmProdPlan
    If Not frmProdPlan.CheckForPlanning(Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1)) Then
        Unload frmProdPlan
        MsgBox "This Order Has Been Planned.", vbInformation
        Exit Sub
    End If
    
End Sub

Private Sub EditOrder()
    Call LV_DblClick
End Sub

Private Sub mnuedit_Click()
    Call LV_DblClick
End Sub

Private Sub PrintOrder(Optional bPrintBalance As Boolean = False, Optional bMakerWise As Boolean = False)

    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    Dim strSelection As String
    Dim rpt As CRAXDDRT.Report
    If Not bPrintBalance Then
        If bMakerWise Then
            Set rpt = rptApp.OpenReport(RptPath & "\rptProductionOrderMakerwise.rpt") '
        Else
            Set rpt = rptApp.OpenReport(RptPath & "\rptProductionOrder.rpt") '
        End If
        rpt.FormulaFields.GetItemByName("ProductionOrder").Text = bForProduction
        strSelection = "{CustOrders.OrderNo}='" & Left(LV.SelectedItem.key & "'", Len(LV.SelectedItem.key) - 1) & "' "
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\rptProductionOrderBalanceQty.rpt") '
        rpt.FormulaFields.GetItemByName("ProductionOrder").Text = bForProduction
        strSelection = "{CustOrders.OrderNo}='" & Left(LV.SelectedItem.key & "'", Len(LV.SelectedItem.key) - 1) & "' AND {@Balance}>0"
    End If
    rpt.FormulaFields.GetItemByName("UserName").Text = "'" & CurrentUserName & "'"
    Load frmPrevRpt
    frmPrevRpt.ShowReport strSelection, rpt
    
    Screen.MousePointer = vbDefault

    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub PrintSummary()
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
    
    Dim f As New frmPrevRpt
    Dim rpt As CRAXDDRT.Report 'New rptFOrderList
    Set rpt = rptApp.OpenReport(RptPath & "\rptFOrderList.rpt") '
    rpt.SQLQueryString = SQLString
        
    rpt.FormulaFields(3).Text = "'" & cmbCust.Text & " " & cmbCountry.Text & "'"
    'rptFOrderList.FormulaFields(4).Text = "'" & strFromTo & "'"
    
    frmPrevRpt.ShowReport " ", rpt
    
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub mnuSummary_Click()

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
    
        con.Execute "Update Vouchers Set [Desc]='" & txtDesc & "' Where Sno=" & Val(LV.SelectedItem.key)
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

Private Sub PrintCustomerOrder()
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\CustomerPO.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{FCustomerOrders.OrderNo}='" & LV.SelectedItem.SubItems(4) & "'", rpt
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(txtEdit) < 0 Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        End If
        'Now Update Order & Proforma Qty....
        Dim lID As Long, lQtyDiff As Long
        lQtyDiff = Val(txtEdit) - Val(LV.SelectedItem.ListSubItems(6).Tag)
        
        lID = Val(LV.SelectedItem.key)
        
        Call StartTrans(con)
        con.Execute "UPDATE FOrderItems SET Qty=Qty+" & lQtyDiff & " WHERE ID=" & lID
        con.Execute "UPDATE FProformaOrders SET Qty=Qty+" & lQtyDiff & " WHERE OrderEntryID=" & lID
        con.CommitTrans
        
        LV.SelectedItem.ListSubItems(6).Text = Val(txtEdit)
        LV.SelectedItem.ListSubItems(6).Tag = Val(LV.SelectedItem.ListSubItems(6).Text)
        
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

Private Sub txtFind_GotFocus()
   txtFind.SelStart = 0
   txtFind.SelLength = Len(txtFind)
   cmdFind.Default = True
End Sub

Private Sub txtFind_LostFocus()
   cmdFind.Default = False
End Sub

