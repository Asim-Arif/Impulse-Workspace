VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmArticlewiseShippedStatus 
   ClientHeight    =   9960
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
   ScaleHeight     =   9960
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
      Height          =   9990
      Left            =   60
      TabIndex        =   1
      Top             =   -45
      Width           =   11865
      Begin VB.Frame Frame4 
         Caption         =   "Quotations :"
         Height          =   1980
         Left            =   120
         TabIndex        =   20
         Top             =   7575
         Width           =   11700
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   5550
            TabIndex        =   21
            Top             =   780
            Visible         =   0   'False
            Width           =   1000
         End
         Begin MSComctlLib.ListView LVQ 
            Height          =   1710
            Left            =   45
            TabIndex        =   22
            Top             =   210
            Width           =   11550
            _ExtentX        =   20373
            _ExtentY        =   3016
            View            =   3
            LabelEdit       =   1
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
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Sr. No."
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Customer"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Quot #"
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Qty"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Price"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Amount"
               Object.Width           =   2117
            EndProperty
         End
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   330
         Left            =   120
         TabIndex        =   15
         Top             =   330
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   582
         BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   -1
         Appearance      =   0
      End
      Begin VB.Frame Frame3 
         Caption         =   "Customer Prices :"
         Height          =   1935
         Left            =   120
         TabIndex        =   12
         Top             =   5640
         Width           =   11700
         Begin VB.TextBox txtEditPrice 
            Alignment       =   2  'Center
            Height          =   300
            Left            =   5550
            TabIndex        =   14
            Top             =   780
            Visible         =   0   'False
            Width           =   1000
         End
         Begin MSComctlLib.ListView LVCust 
            Height          =   1695
            Left            =   90
            TabIndex        =   13
            Top             =   180
            Width           =   11490
            _ExtentX        =   20267
            _ExtentY        =   2990
            View            =   3
            LabelEdit       =   1
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
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Cust Code"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Object.Tag             =   "FOB"
               Text            =   "FOB-KHI"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Object.Tag             =   "ExWorks"
               Text            =   "FOB-SKT"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Object.Tag             =   "CnISea"
               Text            =   "C&I Sea"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Object.Tag             =   "CnIAir"
               Text            =   "C&I Air"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Object.Tag             =   "CIFSea"
               Text            =   "CIF Sea"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Object.Tag             =   "CIFAir"
               Text            =   "CIF Air"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Object.Tag             =   "CnFSea"
               Text            =   "CnF Sea"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   8
               Object.Tag             =   "CnFAir"
               Text            =   "CnF Air"
               Object.Width           =   1764
            EndProperty
         End
      End
      Begin ComboList.Usercontrol1 cmbArticle 
         Height          =   285
         Left            =   3960
         TabIndex        =   2
         Top             =   345
         Width           =   6795
         _ExtentX        =   11986
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
      Begin VB.Frame Frame2 
         Caption         =   "Unshipped Status :"
         Height          =   2220
         Left            =   105
         TabIndex        =   6
         Top             =   3375
         Width           =   11700
         Begin VB.CommandButton cmdPrintUnShipped 
            Caption         =   "Print"
            Height          =   345
            Left            =   90
            TabIndex        =   11
            Top             =   1815
            Width           =   960
         End
         Begin VB.TextBox txtTotalUnshipped 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   8580
            TabIndex        =   9
            Top             =   1830
            Width           =   1215
         End
         Begin MSComctlLib.ListView LV1 
            Height          =   1530
            Left            =   90
            TabIndex        =   7
            Top             =   255
            Width           =   11490
            _ExtentX        =   20267
            _ExtentY        =   2699
            View            =   3
            LabelEdit       =   1
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
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Sr. No."
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "Customer"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Order No."
               Object.Width           =   2822
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Proforma"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Qty Ordered"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Qty Shipped"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "Balance"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "Price"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   8
               Text            =   "Amount"
               Object.Width           =   2117
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Shipped Status :"
         Height          =   2685
         Left            =   105
         TabIndex        =   4
         Top             =   690
         Width           =   11700
         Begin VB.CommandButton cmdPrintShipped 
            Caption         =   "Print"
            Height          =   285
            Left            =   90
            TabIndex        =   10
            Top             =   2325
            Width           =   960
         End
         Begin VB.TextBox txtTotalShipped 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   5175
            TabIndex        =   8
            Top             =   2325
            Width           =   1215
         End
         Begin MSComctlLib.ListView LV 
            Height          =   2055
            Left            =   90
            TabIndex        =   5
            Top             =   255
            Width           =   11490
            _ExtentX        =   20267
            _ExtentY        =   3625
            View            =   3
            LabelEdit       =   1
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
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Sr. No."
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Invoice #"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Date"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "Qty"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Price"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   5
               Text            =   "Amount"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Order No."
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "Days"
               Object.Width           =   1764
            EndProperty
         End
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Article"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3975
         TabIndex        =   19
         Top             =   120
         Width           =   540
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   330
         Left            =   2145
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   330
         Width           =   1800
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3175;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Calibri"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   17
         Top             =   105
         Width           =   810
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Country"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2160
         TabIndex        =   16
         Top             =   105
         Width           =   660
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   345
         Left            =   10440
         TabIndex        =   3
         Top             =   9570
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "2355;609"
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   390
         Left            =   10800
         TabIndex        =   0
         Top             =   240
         Width           =   960
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "1693;688"
         Accelerator     =   82
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopMenu"
   End
End
Attribute VB_Name = "frmArticlewiseShippedStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bFromPPDashBoard As Boolean
Dim iColNo As Integer

Private Sub cmbArticle_matched()

    On Error GoTo err
    
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim lTotalQty As Long
    Dim lBalanceQty As Long, lTotalUnshipped As Long
    
    Dim bRestrictedAccess As Boolean
    If CurrentUserName <> "Administrator" Then
        bRestrictedAccess = GetSingleBooleanValue("RestrictedItemProfile", "Users", " WHERE UserName='" & CurrentUserName & "'")
    Else
        bRestrictedAccess = False
    End If
    Dim dPrice As Double, lDays As Long
    Dim strCondition As String
    strCondition = " WHERE CompItemCode='" & cmbArticle.id & "' AND GatePassDT IS NOT NULL"
    If cmbCust.id <> "0" Then
        strCondition = strCondition & " AND CustCode='" & cmbCust.id & "'"
    End If
    strCondition = strCondition & " ORDER BY DT"
    
    Dim strShipped As String
    strShipped = " WHERE CompItemCode='" & cmbArticle.id & "' AND GatePassDT IS NOT NULL"
    If cmbCust.id <> "0" Then
        strShipped = strShipped & " AND CustCode='" & cmbCust.id & "'"
    End If
    strShipped = strShipped & "  ORDER BY DT DESC"
    With rs
        .Open "SELECT CustomInvoice,DT,Qty,Price,GatePassDT,OrderNo,DeliveryDTItem,OrderDT FROM VCustomInvoice" & strShipped, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            ITM.ListSubItems.add , , !CustomInvoice & ""
            ITM.ListSubItems.add , , Format(!GatePassDT, "dd-MMM-yyyy")
            ITM.ListSubItems.add , , Val(!Qty & "")
            If bRestrictedAccess Or bFromPPDashBoard Then
                dPrice = 0
            Else
                dPrice = Val(!Price & "")
            End If
            ITM.ListSubItems.add , , dPrice
            ITM.ListSubItems.add , , Val(!Qty & "") * dPrice
            lTotalQty = lTotalQty + Val(!Qty & "")
            ITM.ListSubItems.add , , !OrderNo & ""
'            If IsNull(!DeliveryDTItem) Or !DeliveryDTItem = #1/1/1900# Then
'                lDays = 0
'            Else
'                lDays = DateDiff("d", !DeliveryDTItem, !GatePassDT)
'            End If
            lDays = DateDiff("d", !OrderDT, !GatePassDT)
            ITM.ListSubItems.add , , lDays
            .MoveNext
        Loop
        .Close
        Set ITM = LV.ListItems.add
        ITM.ListSubItems.add , , ""
        ITM.ListSubItems.add , , ""
        ITM.ListSubItems.add , , lTotalQty
        ITM.ListSubItems(3).Bold = True
'        Dim lTotalQtyM As Long
'        .Open "SELECT * FROM ItemPriceHistory_Imported WHERE ItemCode='" & cmbArticle.id & "'", con, adOpenForwardOnly, adLockUnspecified
'        LVM.ListItems.Clear
'        Do Until .EOF
'            Set ITM = LVM.ListItems.add(, , LVM.ListItems.count + 1)
'            ITM.ListSubItems.add , , !InvoiceNo & ""
'            ITM.ListSubItems.add , , Format(!DT, "dd-MMM-yyyy")
'            ITM.ListSubItems.add , , Val(!Qty & "")
'            If bRestrictedAccess Or bFromPPDashBoard Then
'                dPrice = 0
'            Else
'                dPrice = Val(!Price & "")
'            End If
'            ITM.ListSubItems.add , , dPrice
'            ITM.ListSubItems.add , , Val(!Qty & "") * dPrice
'            lTotalQtyM = lTotalQtyM + Val(!Qty & "")
'            .MoveNext
'        Loop
'        .Close
'        Set ITM = LVM.ListItems.add
'        ITM.ListSubItems.add , , ""
'        ITM.ListSubItems.add , , ""
'        ITM.ListSubItems.add , , lTotalQtyM
'        ITM.ListSubItems(3).Bold = True
        
        txtTotalShipped = lTotalQty '+ lTotalQtyM
         
        Dim strUnShipped As String
        strUnShipped = " WHERE CompItemCode='" & cmbArticle.id & "' AND OrderNo IN(SELECT OrderNo FROM VUnshippedOrderList) AND Qty>ShippedQty AND OrderNo NOT IN('Stock-Order','Stock-OrderIss')"
        If cmbCust.id <> "0" Then
            strUnShipped = strUnShipped & " AND CustCode='" & cmbCust.id & "'"
        End If
        strUnShipped = strUnShipped & " ORDER BY OrderNo"
        
        .Open "SELECT * FROM VFOrderItemswithShippedQty" & strUnShipped, con, adOpenForwardOnly, adLockReadOnly
        LV1.ListItems.Clear
        Do Until .EOF
            Set ITM = LV1.ListItems.add(, , LV1.ListItems.count + 1)
            ITM.ListSubItems.add , , !CustCode & ""
            ITM.ListSubItems.add , , !OrderNo & ""
            ITM.ListSubItems.add , , !Proformas & ""
            ITM.ListSubItems.add , , Val(!Qty & "")
            ITM.ListSubItems.add , , Val(!ShippedQty & "")
            lBalanceQty = Val(!Qty & "") - Val(!ShippedQty & "")
            lTotalUnshipped = lTotalUnshipped + lBalanceQty
            ITM.ListSubItems.add , , lBalanceQty
            If bRestrictedAccess Or bFromPPDashBoard Then
                dPrice = 0
            Else
                dPrice = Val(!Price & "")
            End If
            ITM.ListSubItems.add , , dPrice
            ITM.ListSubItems.add , , dPrice * lBalanceQty
            .MoveNext
        Loop
        .Close
        txtTotalUnshipped = lTotalUnshipped
        
        If bRestrictedAccess = False Then
            .Open "SELECT * FROM FCustomerCatalog WHERE CompItemID='" & cmbArticle.id & "'", con, adOpenForwardOnly, adLockReadOnly
            LVCust.ListItems.Clear
            Do Until .EOF
                Set ITM = LVCust.ListItems.add(, !EntryID & "'", !CustCode & "")
                ITM.ListSubItems.add , , Val(!FOB & "")
                ITM.ListSubItems.add , , Val(!ExWorks & "")
                ITM.ListSubItems.add , , Val(![CnISea] & "")
                ITM.ListSubItems.add , , Val(![CnIAir] & "")
                ITM.ListSubItems.add , , Val(![CIFSea] & "")
                ITM.ListSubItems.add , , Val(![CIFAir] & "")
                ITM.ListSubItems.add , , Val(![CnFSea] & "")
                ITM.ListSubItems.add , , Val(![CnFAir] & "")
                .MoveNext
            Loop
            .Close
        End If
        
        .Open "SELECT * FROM VCustomerQuotations WHERE ItemCode='" & cmbArticle.id & "'", con, adOpenForwardOnly, adLockReadOnly
        LVQ.ListItems.Clear
        Do Until .EOF
            Set ITM = LVQ.ListItems.add(, , LV1.ListItems.count + 1)
            ITM.ListSubItems.add , , !CustCode & ""
            ITM.ListSubItems.add , , !QuotationNo & ""
            ITM.ListSubItems.add , , Val(!Qty & "")
            
            ITM.ListSubItems.add , , Val(!Price & "")
            ITM.ListSubItems.add , , Val(!Price & "") * Val(!Qty & "")
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()

    cmbArticle.ClearVals
    cmbArticle.DropDownOffLine = True
    If cmbCust.id = "0" Then
        cmbArticle.AddVals con, "ItemName + ' {' + ItemID + '}'", "Items", "ItemID", " ORDER BY ItemID"
    Else
        cmbArticle.AddVals con, "'{' + CompItemID + '} ' + ItemID + ' {' + ISNULL(OCR,'') + '} {' + ISNULL(OCR2,'') + '} ' + CompDesc", "VFCustomerCatalog", "CompItemID", " WHERE CustCode='" & cmbCust.id & "' AND Country='" & cmbCountry.Text & "'"
    End If
    
End Sub

Private Sub cmbCust_matched()

    If cmbCust.id = "0" Then
        cmbCountry.Clear
        cmbCountry.AddItem "<Comp>"
        cmbCountry.ListIndex = 0
        cmbCountry.Enabled = False
        Exit Sub
    Else
        cmbCountry.Enabled = True
    End If
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT Country,Curr FROM ForeignCustomers WHERE CustCode='" & cmbCust.id & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            cmbCountry.List(cmbCountry.ListCount - 1, 1) = ![Curr] & ""
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    If cmbCountry.ListCount > 0 Then cmbCountry.ListIndex = 0
    
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
    FAcc.Enabled = True
End Sub

Private Sub cmdPrintShipped_Click()
    Dim rpt As CRAXDDRT.Report, rptManual As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptArticlewiseSales.rpt")
'    Set rptManual = rpt.OpenSubreport("ManualSalesData")
'    rptManual.RecordSelectionFormula = "{ItemPriceHistory_Imported.ItemCode}='" & cmbArticle.id & "'"
    Dim f As New frmPrevRpt
    rpt.FormulaFields.GetItemByName("HideDetail").Text = False
    f.ShowReport "{VArticlewiseSales.ItemID}='" & cmbArticle.id & "'", rpt
End Sub

Private Sub cmdPrintUnShipped_Click()
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptPartiallyShippedOrders.rpt")
    Dim f As New frmPrevRpt
    
    f.ShowReport "{Items.ItemID}='" & cmbArticle.id & "' AND NOT ({FCustomerOrders.OrderNo} IN['Stock-Order','Stock-OrderIss'])", rpt
    
End Sub

Private Sub cmdRefresh_Click()
    Call cmbArticle_matched
End Sub

Private Sub Form_Load()

    iColNo = 0
    mnuPop.Visible = False
    
    cmbCust.ListHeight = 4000
    cmbCust.AddItem "<Comp>", "0"
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " WHERE Active=1")
    cmbCust.id = "0"
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    FAcc.Top = (Me.ScaleHeight - FAcc.Height) / 2
    
End Sub


Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If x > LV.ColumnHeaders(1).Left And x < LV.ColumnHeaders(i).Left + LV.ColumnHeaders(i).Width Then
            iColNo = i
            Exit For
        End If
    Next
    
End Sub

Private Sub LVCust_DblClick()

    If LVCust.ListItems.count = 0 Then Exit Sub
    
    With txtEditPrice
        .Move LVCust.Left + LVCust.ColumnHeaders(iColNo).Left, LVCust.Top + LVCust.SelectedItem.Top
        .Visible = True
        .Text = Val(LVCust.SelectedItem.SubItems(iColNo - 1))
        .Tag = Val(.Text)
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub

Private Sub LVCust_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    For i = 1 To LVCust.ColumnHeaders.count
        If x > LVCust.ColumnHeaders(1).Left And x < LVCust.ColumnHeaders(i).Left + LVCust.ColumnHeaders(i).Width Then
            iColNo = i
            Exit For
        End If
    Next
End Sub

Private Sub txtEditPrice_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        If Val(txtEditPrice) <> Val(txtEditPrice.Tag) Then  'Price Changed
    
            Dim strField As String
            
            strField = LVCust.ColumnHeaders(iColNo).Tag
            
            con.Execute "UPDATE FCustomerCatalog SET " & strField & "=" & Val(txtEditPrice) & " WHERE EntryID=" & Val(LVCust.SelectedItem.key)
            
            con.Execute "INSERT INTO FCustCatalRevs(RefID,Description,FOB,CIFSea,CIFAir,CnFSea,CnFAir,FOBSialkot,CnISea," & _
             "CnIAir,DT,UserName) SELECT EntryID,Description,FOB,CIFSea,CIFAir,CnFSea,CnFAir,ExWorks,CnISea," & _
             "CnIAir,getDate(),'" & CurrentUserName & "' FROM FCustomerCatalog WHERE EntryID=" & Val(LVCust.SelectedItem.key)
                     
            Dim lEntryID1 As Long
            
            lEntryID1 = GetSingleLongValue("MAX(EntryID)", "FCustCatalRevs")
            con.Execute "UPDATE FCustCatalRevs SET " & IIf(strField = "ExWorks", "FOBSialkot", strField) & "=" & Val(txtEditPrice.Tag) & " WHERE EntryID=" & lEntryID1
            
            LV.SelectedItem.SubItems(iColNo - 1) = Val(txtEditPrice)
        End If
        txtEditPrice.Visible = False
        
    ElseIf KeyAscii = 27 Then
        txtEditPrice.Visible = False
    End If
    
End Sub

Private Sub txtEditPrice_LostFocus()
    txtEditPrice.Visible = False
End Sub

Public Sub ShowMe(p_strItemCode As String, Optional p_bFromPPDashBoard As Boolean = False)
    
    bFromPPDashBoard = p_bFromPPDashBoard
    cmbArticle.ListHeight = 5800
    
'    If p_strItemCode = "" Then
'        cmbArticle.DropDownOffLine = True
'        Call cmbArticle.AddVals(con, "'{ ' + ItemID + ' } ' + ItemName ", "Items", "ItemID", " WHERE InActive=0")
'    Else
    If p_strItemCode <> "" Then
        cmbArticle.DropDownOffLine = False
        Call cmbArticle.AddVals(con, "'{ ' + ItemID + ' } ' + ItemName ", "Items", "ItemID", " WHERE ItemID='" & p_strItemCode & "'")
        cmbArticle.id = p_strItemCode
        Call cmbArticle_matched
    End If
'   End If
    Me.Show
    
End Sub
