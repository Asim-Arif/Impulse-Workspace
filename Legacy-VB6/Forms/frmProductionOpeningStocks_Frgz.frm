VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmProductionOpeningStocks_Frgz 
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   45
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
   ScaleHeight     =   9015
   ScaleMode       =   0  'User
   ScaleWidth      =   10721.69
   WindowState     =   2  'Maximized
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
      Height          =   9045
      Left            =   60
      TabIndex        =   7
      Top             =   -45
      Width           =   11865
      Begin ComboList.Usercontrol1 cmbProcess 
         Height          =   285
         Left            =   7515
         TabIndex        =   1
         Top             =   960
         Width           =   3240
         _ExtentX        =   5715
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
      Begin VB.TextBox txtMilNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9945
         TabIndex        =   25
         Top             =   1500
         Width           =   1080
      End
      Begin VB.TextBox txtBatchNo_Manual 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8835
         TabIndex        =   24
         Top             =   1500
         Width           =   1080
      End
      Begin VB.TextBox txtLotNo_Manual 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7725
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1500
         Width           =   1080
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Left            =   6285
         TabIndex        =   22
         Top             =   4275
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5730
         TabIndex        =   21
         Top             =   8580
         Width           =   1185
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
         Height          =   570
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   11880
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "S/F Receive Items"
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
            Index           =   2
            Left            =   4560
            TabIndex        =   19
            Top             =   135
            Width           =   2730
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "S/F Receive Items"
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
            Index           =   3
            Left            =   4560
            TabIndex        =   20
            Top             =   150
            Width           =   2730
         End
      End
      Begin ComboList.Usercontrol1 cmbArticle 
         Height          =   285
         Left            =   45
         TabIndex        =   0
         Top             =   960
         Width           =   7455
         _ExtentX        =   13150
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
         Left            =   2205
         TabIndex        =   4
         Top             =   1500
         Width           =   3690
         _ExtentX        =   6509
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
      Begin ComboList.Usercontrol1 cmbStore 
         Height          =   285
         Left            =   45
         TabIndex        =   3
         Top             =   1500
         Width           =   2145
         _ExtentX        =   3784
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
         Height          =   285
         Left            =   5910
         TabIndex        =   5
         Top             =   1500
         Width           =   1830
      End
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   10770
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin MSComctlLib.ListView LV 
         Height          =   6630
         Left            =   45
         TabIndex        =   8
         Top             =   1875
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   11695
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Process"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Location"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Remarks"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Date"
            Object.Width           =   2364
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Lot No."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Batch No."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "Mil No."
            Object.Width           =   1499
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Mil No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   10050
         TabIndex        =   28
         Top             =   1290
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Batch No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   8940
         TabIndex        =   27
         Top             =   1290
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Lot No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   7725
         TabIndex        =   26
         Top             =   1305
         Width           =   570
      End
      Begin MSForms.CommandButton cmdPintLedger 
         Height          =   375
         Left            =   8640
         TabIndex        =   17
         Top             =   8550
         Width           =   1755
         ForeColor       =   0
         Caption         =   "Print Ledger"
         PicturePosition =   327683
         Size            =   "3096;661"
         Accelerator     =   76
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdPostMissing 
         Height          =   375
         Left            =   45
         TabIndex        =   16
         Top             =   8550
         Visible         =   0   'False
         Width           =   1755
         ForeColor       =   0
         Caption         =   "Post Mmissing"
         PicturePosition =   327683
         Size            =   "3096;661"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   5940
         TabIndex        =   15
         Top             =   1290
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2220
         TabIndex        =   14
         Top             =   1305
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Store"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   45
         TabIndex        =   13
         Top             =   1290
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   10770
         TabIndex        =   12
         Top             =   750
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Process"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   7515
         TabIndex        =   11
         Top             =   765
         Width           =   660
      End
      Begin VB.Label lbl 
         Caption         =   "Article"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   45
         TabIndex        =   10
         Top             =   765
         Width           =   1155
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   375
         Left            =   10440
         TabIndex        =   9
         Top             =   8550
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
      Begin MSForms.CommandButton cmdAdd 
         Height          =   360
         Left            =   11055
         TabIndex        =   6
         Top             =   1440
         Width           =   690
         Caption         =   "Add"
         PicturePosition =   327683
         Size            =   "1217;635"
         Accelerator     =   65
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
End
Attribute VB_Name = "frmProductionOpeningStocks_Frgz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iColNo As Integer

Private Sub cmbArticle_matched()

    cmbProcess.ClearVals
    Call cmbProcess.AddVals(con, "Description", "VItemProcesses", "ProcessID", " WHERE IsExist='" & cmbArticle.ID & "' ORDER BY ItemSNo")
    
    Call RefreshList
    
    If cmdAdd.Enabled = False Then
        cmdAdd.Enabled = True
    End If
    
End Sub

Private Sub cmbStore_matched()
    
    cmbLocation.ClearVals
    cmbLocation.AddVals con, "RackNo + ' {' + ShelfNo + '}'", "VStoreShelfs", "EntryID", " WHERE Store_RefID=" & cmbStore.ID & " ORDER BY StoreName,RackNo,ShelfNo"
    
End Sub

Private Sub cmdAdd_Click()

    On Error GoTo err
    If cmbArticle.MatchFound = False Then
        MsgBox "Please Select Article.", vbInformation
        Exit Sub
    ElseIf cmbProcess.MatchFound = False Then
        MsgBox "Please Select Process.", vbInformation
        Exit Sub
    ElseIf Val(txtQty) <= 0 Then
        MsgBox "Invalid Qty.", vbInformation
        Exit Sub
    ElseIf cmbLocation.MatchFound = False Then
        MsgBox "Invalid Location.", vbInformation
        Exit Sub
    ElseIf txtBatchNo_Manual = "" Then
        MsgBox "Please Enter Batch No.", vbInformation
        Exit Sub
    End If
'    If txtLotNo_Manual = "" Then
'        MsgBox "Please Enter Lot No.", vbInformation
'        Exit Sub
'    End If
    'Check If Opening for Selected Article for Selected Process is feeded ?
    
    Dim lCount As Long
'    lCount = GetSingleLongValue("COUNT(LotNo_Manual)", "StockOrderOpening", " WHERE LotNo_Manual='" & txtLotNo_Manual & "'")
'    If lCount > 0 Then
'        MsgBox "Lot No. Exists.", vbInformation
'        Exit Sub
'    End If
    If txtBatchNo_Manual <> "" Then
        lCount = GetSingleLongValue("COUNT(BatchNo_Manual)", "StockOrderOpening", " WHERE BatchNo_Manual='" & txtBatchNo_Manual & "'")
        If lCount > 0 Then
            MsgBox "Batch No. Exists.", vbInformation
            Exit Sub
        End If
    End If
    txtLotNo_Manual = getNextLotNo
'    lCount = GetSingleLongValue("COUNT(ItemID)", "StockOrderOpening", " WHERE ItemID='" & cmbArticle.ID & "' AND ProcessID=" & cmbProcess.ID & " AND Shelf_RefID=" & cmbLocation.ID)
'
    Call StartTrans(con)
'
'    If lCount > 0 Then
'        Dim strRemarks As String
'        strRemarks = GetSingleStringValue("Remarks", "StockOrderOpening", " WHERE ItemID='" & cmbArticle.ID & "' AND ProcessID=" & cmbProcess.ID & " AND Shelf_RefID=" & cmbLocation.ID)
'        If strRemarks <> txtRemarks Then
'            strRemarks = strRemarks & "," & txtRemarks
'        End If
'        con.Execute "UPDATE StockOrderOpening SET Qty=Qty+" & Val(txtQty) & ",Remarks='" & strRemarks & "' WHERE ItemID='" & cmbArticle.ID & "' AND ProcessID=" & cmbProcess.ID
'
'    Else
'        con.Execute "INSERT INTO StockOrderOpening(ItemID,ProcessID,Qty,Location,UserName,MachineName,Shelf_RefID,Remarks) VALUES('" & cmbArticle.ID & _
'         "'," & cmbProcess.ID & "," & Val(txtQty) & ",'" & cmbLocation.Text & "','" & CurrentUserName & "','" & strComputerName & "'," & cmbLocation.ID & ",'" & txtRemarks & "')"
'    End If
    con.Execute "INSERT INTO StockOrderOpening(ItemID,ProcessID,Qty,Location,UserName,MachineName,Shelf_RefID,Remarks,LotNo_Manual,BatchNo_Manual,Mill_Certificate_No) VALUES('" & cmbArticle.ID & _
         "'," & cmbProcess.ID & "," & Val(txtQty) & ",'" & cmbLocation.Text & "','" & CurrentUserName & "','" & strComputerName & "'," & _
         cmbLocation.ID & ",'" & txtRemarks & "','" & txtLotNo_Manual & "','" & txtBatchNo_Manual & "','" & txtMilNo & "')"
         
    con.Execute "INSERT INTO Lots_List(LotNo,ItemID,Lot_Type,Reference_LotNo,Batch_No,Mill_Certificate_No) VALUES('" & txtLotNo_Manual & "','" & cmbArticle.ID & "',1,'','" & txtBatchNo_Manual & "','" & txtMilNo & "')"
    'Check Customer...
    lCount = 0
    lCount = GetSingleLongValue("COUNT(CustCode)", "ForeignCustomers", " WHERE CustCode='Stock' AND Country='PK'")
    If lCount = 0 Then
        con.Execute "INSERT INTO ForeignCustomers(Custcode,Country) VALUES('Stock','PK')"
    End If
    
    'Now Check to see if Stock Order is feeded.
    lCount = 0
    lCount = GetSingleLongValue("COUNT(OrderNo)", "FCustomerOrders", " WHERE OrderNo='Stock-Order'")
    If lCount = 0 Then
        con.Execute "INSERT INTO FCustomerOrders(CustCode,Country,OrderNo,DT) VALUES('Stock','PK','Stock-Order','" & Date & "')"
    End If
    
    'Now Check If Article is Assigned to Customer
    lCount = 0
    lCount = GetSingleLongValue("COUNT(CompItemID)", "FCustomerCatalog", " WHERE Custcode='Stock' AND Country='PK' AND CompItemID='" & cmbArticle.ID & "'")
    If lCount = 0 Then
        con.Execute "INSERT INTO FCustomerCatalog(Custcode,Country,ItemID,CompItemID) VALUES('Stock','PK','" & cmbArticle.ID & "','" & cmbArticle.ID & "')"
    End If
    'Now Check If Article is feeded against this Article
    lCount = 0
    lCount = GetSingleLongValue("COUNT(OrderNo)", "FOrderItems", " WHERE OrderNo='Stock-Order' AND CompItemCode='" & cmbArticle.ID & "'")
    If lCount = 0 Then
        con.Execute "INSERT INTO FOrderItems(OrderNo,ItemCode,CompItemCode,Qty,DeliveryDT,Stamps,Quality) VALUES('Stock-Order','" & cmbArticle.ID & _
         "','" & cmbArticle.ID & "'," & Val(txtQty) & ",'" & Date & "','','')"
    Else
        con.Execute "UPDATE FOrderItems SET Qty=Qty+" & Val(txtQty) & " WHERE OrderNo='Stock-Order' AND ItemCode='" & cmbArticle.ID & "'"
    End If
    
    Dim lEntryID As Long, lRcvEntryID As Long, lNextProcessID As Long
    lEntryID = GetSingleLongValue("MAX(EntryID)", "StockOrderOpening")
    con.Execute "INSERT INTO VendReceived(VendID,DT,RecieptID,UserID,ProcessID,Issuance_RefID,UserName,MachineName) VALUES(" & _
     "0,'" & Date & "','',0," & cmbProcess.ID & ",0,'" & CurrentUserName & "','" & strComputerName & "')"
    
    lRcvEntryID = GetSingleLongValue("MAX(EntryID)", "VendReceived", " WHERE MachineName='" & strComputerName & "'")
    
    lNextProcessID = getNextProcessID(cmbArticle.ID, cmbProcess.ID)
    
    con.Execute "INSERT INTO VendRcvdDetail(RefID,ItemCode,RecieptID,RcvdQty,Rate,LotNo,NextProcessID,OrderNo" & _
     ",CountedBy,Issue_RefID,ProcessID,RcvdWeight,Opening_RefID) VALUES(" & lRcvEntryID & ",'" & cmbArticle.ID & _
     "',''," & Val(txtQty) & ",0,'0'," & lNextProcessID & ",'Stock-Order','',0," & cmbProcess.ID & _
     ",0," & lEntryID & ")"
     
    con.CommitTrans
    
    
    
    Call RefreshList
    Call Print_SF_Opening_Slip(lEntryID)
    cmdAdd.Enabled = False
    cmbArticle.ID = ""
    txtLotNo_Manual = getNextLotNo
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdPintLedger_Click()

    If cmbArticle.MatchFound = False Then
        MsgBox "Please Select Item.", vbInformation
        Exit Sub
    End If
    Call PrintSFLedger(cmbArticle.ID)
    
End Sub

Private Sub cmdPostMissing_Click()
    
    Dim rs As New ADODB.Recordset
    Dim lEntryID As Long, lProcessID As Long, DT As Date, lRcvEntryID As Long, lNextProcessID As Long
    Dim strItemID As String, iCount As Integer
    
    Call StartTrans(con)
    With rs
        .Open "SELECT * FROM StockOrderOpening WHERE EntryID NOT IN(SELECT Opening_RefID FROM VendRcvdDetail WHERE Opening_RefID IS NOT NULL) ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        iCount = .RecordCount
        Do Until .EOF
            lEntryID = Val(!EntryID & "")
            lProcessID = Val(!ProcessID & "")
            DT = Format(!DTEntry, "dd-MMM-yyyy")
            strItemID = !ITemID & ""
            
            con.Execute "INSERT INTO VendReceived(VendID,DT,RecieptID,UserID,ProcessID,Issuance_RefID) VALUES(" & _
             "0,'" & DT & "','',0," & lProcessID & ",0)"
             
            lRcvEntryID = GetSingleLongValue("MAX(EntryID)", "VendReceived")
            lNextProcessID = getNextProcessID(strItemID, lProcessID)
            
            con.Execute "INSERT INTO VendRcvdDetail(RefID,ItemCode,RecieptID,RcvdQty,Rate,LotNo,NextProcessID,OrderNo" & _
             ",CountedBy,Issue_RefID,ProcessID,RcvdWeight,Opening_RefID) VALUES(" & lRcvEntryID & ",'" & strItemID & _
             "',''," & Val(!Qty & "") & ",0,'0'," & lNextProcessID & ",'Stock-Order','',0," & lProcessID & _
             ",0," & lEntryID & ")"
            
            .MoveNEXT
        Loop
    End With
    
    con.CommitTrans
    
    MsgBox "Done " & iCount & " Entries."
    
End Sub

Private Sub Form_Load()

    iColNo = 0
    cmbArticle.ListHeight = 5800
    cmbArticle.DropDownOffLine = True
    Call cmbArticle.AddVals(con, "'{ ' + ItemID + ' } ' + ItemName ", "Items", "ItemID", " WHERE InActive=0")
        
    cmbProcess.ListHeight = 4000
    
    'Call AddToCombo(cmbLocation, "StoreName + '(' + RackNo + ') (' + ShelfNo + ')'", "VStoreShelfs", , , "EntryID")
    cmbStore.ListHeight = 4000
    cmbLocation.ListHeight = 4000
    'cmbStore.AddVals con, "StoreName", "Stores", "EntryID", " WHERE EntryID=3 ORDER BY StoreName"
    cmbStore.AddVals con, "StoreName", "Stores", "EntryID", " WHERE EntryID IN(SELECT Store_RefID FROM Users_Stores WHERE UserID=" & UserID & ") ORDER BY StoreName"
    'cmbLocation.AddVals con, "StoreName + ' {' + RackNo + '} {' + ShelfNo + '}'", "VStoreShelfs", "EntryID", " ORDER BY StoreName,RackNo,ShelfNo"
    txtLotNo_Manual = getNextLotNo
'    If CurrentUserName = "Administrator" Then
'        cmdPostMissing.Visible = True
'    End If
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    FAcc.Top = (Me.ScaleHeight - FAcc.Height) / 2
   
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Dim myLV As ListView
    Set myLV = LV
    
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

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(6).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(6).Width
        .Text = LV.SelectedItem.SubItems(5)
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If LV.ListItems.count = 0 Then Exit Sub
    If strCompany = "Dr-Frgz" Then Exit Sub
    
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Delete ?", vbQuestion + vbYesNo) = vbYes Then
            If getDBPassword("Delete_Semi_Finish_Opening") = False Then Exit Sub
            'First Check If Issuance Have been made against it...
            Dim strItemID As String, lProcessID As Long, lEntryID As Long
            strItemID = LV.SelectedItem.Text
            lProcessID = Val(LV.SelectedItem.ListSubItems(2).Tag)
            'Now Get The EntryID FROM StockOrderOpening...
            lEntryID = Val(LV.SelectedItem.Tag) 'GetSingleLongValue("EntryID", "StockOrderOpening", " WHERE ProcessID=" & lProcessID & " AND ItemID='" & strItemID & "'")
            Dim lIssdQty As Long
            lIssdQty = GetSingleLongValue("SUM(IssQty)", "VendRcvdDetail", " WHERE Opening_RefID=" & lEntryID)
            If lIssdQty > 0 Then
                MsgBox "Can't Delete. Issuance has been made.", vbInformation
                Exit Sub
            End If
            Call StartTrans(con)
            con.Execute "DELETE FROM StockOrderOpening WHERE EntryID=" & lEntryID
            con.Execute "DELETE FROM VendReceived WHERE EntryID IN(SELECT RefID FROM VendRcvdDetail WHERE Opening_RefID=" & lEntryID & ")"
            con.Execute "DELETE FROM VendRcvdDetail WHERE Opening_RefID=" & lEntryID
            'con.Execute "DELETE FROM FOrderItems WHERE CompItemCode='" & strItemID & "' AND OrderNo='Stock-Order'"
            con.CommitTrans
            
            Call RefreshList
        End If
    End If
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

Private Sub RefreshList()

    On Error GoTo err
    Dim rs As New ADODB.Recordset, ITM As ListItem
    Dim lTotal As Long
    With rs
        .Open "SELECT * FROM VStockOrderOpening WHERE ItemID='" & cmbArticle.ID & "' AND Qty>ISNULL(IssdQty,0)", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , !ITemID)
            ITM.Tag = Val(!EntryID & "")
            ITM.ListSubItems.add , , !ItemName & ""
            ITM.ListSubItems.add(, , !Description & "").Tag = Val(!ProcessID & "")
            ITM.ListSubItems.add(, , Val(!Qty & "") - Val(!IssdQty & "")).Bold = True
            ITM.ListSubItems.add , , !StoreName & " (" & !Location & ")"
            ITM.ListSubItems.add , , !Remarks & ""
            ITM.ListSubItems.add , , Format(!DTEntry, "dd-MMM-yyyy")
            ITM.ListSubItems.add , , !LotNo_Manual & ""
            ITM.ListSubItems.add , , !BatchNo_Manual & ""
            ITM.ListSubItems.add , , !Mill_Certificate_No & ""
            lTotal = lTotal + (Val(!Qty & "") - Val(!IssdQty & ""))
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    txtTotal = Format(lTotal, "#,##")
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If LV.ListItems.count = 0 Then Exit Sub
    Dim strRet As String
    strRet = ShowPopUpMenu1("Print Tray Card")
    If strRet = "Print Tray Card" Then
        Dim rpt As CRAXDDRT.Report
        Set rpt = rptApp.OpenReport(RptPath & "\SF_Tray_Card.rpt")
    
        Dim frm As New frmPrevRpt
        frm.ShowReport "{VStockOrderOpening.EntryID}=" & Val(LV.SelectedItem.Tag), rpt
    End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LV.SelectedItem.SubItems(5) = txtEdit.Text
        con.Execute "UPDATE StockOrderOpening SET Remarks='" & txtEdit.Text & "' WHERE EntryID=" & Val(LV.SelectedItem.Tag)
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub
