VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmProductionOpeningStocks 
   ClientHeight    =   8385
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
   ScaleHeight     =   8385
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
      Height          =   8400
      Left            =   60
      TabIndex        =   1
      Top             =   -45
      Width           =   11865
      Begin ComboList.Usercontrol1 cmbArticle 
         Height          =   285
         Left            =   45
         TabIndex        =   2
         Top             =   345
         Width           =   5535
         _ExtentX        =   9763
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
      Begin ComboList.Usercontrol1 cmbProcess 
         Height          =   285
         Left            =   5610
         TabIndex        =   6
         Top             =   345
         Width           =   2955
         _ExtentX        =   5212
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
         Left            =   8595
         TabIndex        =   9
         Top             =   345
         Width           =   975
      End
      Begin MSComctlLib.ListView LV 
         Height          =   7140
         Left            =   45
         TabIndex        =   3
         Top             =   735
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   12594
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item Name"
            Object.Width           =   7779
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Process"
            Object.Width           =   5239
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
            Object.Width           =   2222
         EndProperty
      End
      Begin MSForms.CommandButton cmdPintLedger 
         Height          =   375
         Left            =   8625
         TabIndex        =   12
         Top             =   7935
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
         Left            =   9585
         TabIndex        =   11
         Top             =   135
         Width           =   720
      End
      Begin MSForms.ComboBox cmbLocation 
         Height          =   285
         Left            =   9585
         TabIndex        =   10
         Tag             =   "PaymentTerms"
         Top             =   345
         Width           =   1305
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2302;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   178
         FontPitchAndFamily=   2
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
         Left            =   8595
         TabIndex        =   8
         Top             =   135
         Width           =   300
      End
      Begin VB.Label Label1 
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
         Height          =   240
         Left            =   5610
         TabIndex        =   7
         Top             =   150
         Width           =   1155
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
         TabIndex        =   5
         Top             =   150
         Width           =   1155
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   375
         Left            =   10440
         TabIndex        =   4
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
      Begin MSForms.CommandButton cmdAdd 
         Height          =   360
         Left            =   10950
         TabIndex        =   0
         Top             =   315
         Width           =   795
         Caption         =   "Add"
         PicturePosition =   327683
         Size            =   "1402;635"
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
Attribute VB_Name = "frmProductionOpeningStocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iColNo As Integer

Private Sub cmbArticle_matched()

    cmbProcess.ClearVals
    Call cmbProcess.AddVals(con, "Description", "VItemProcesses", "ProcessID", " WHERE IsExist='" & cmbArticle.ID & "' ORDER BY SNo")
    
    Call RefreshList
    
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
    ElseIf cmbLocation.Text = "" Then
        MsgBox "Invalid Location.", vbInformation
        Exit Sub
    End If
    'Check If Opening for Selected Article for Selected Process is feeded ?
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(ItemID)", "StockOrderOpening", " WHERE ItemID='" & cmbArticle.ID & "' AND ProcessID=" & cmbProcess.ID)
    
    Call StartTrans(con)
    
    If lCount > 0 Then
        con.Execute "UPDATE StockOrderOpening SET Qty=Qty+" & Val(txtQty) & " WHERE ItemID='" & cmbArticle.ID & "' AND ProcessID=" & cmbProcess.ID
        'MsgBox "This Item has already Opening feeded at this Process.", vbInformation
        'Exit Sub
    Else
        con.Execute "INSERT INTO StockOrderOpening(ItemID,ProcessID,Qty,Location,UserName,MachineName) VALUES('" & cmbArticle.ID & _
         "'," & cmbProcess.ID & "," & Val(txtQty) & ",'" & cmbLocation & "','" & CurrentUserName & "','" & strComputerName & "')"
    End If
     
    
    
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
        If strCompany = "Banzai" Then
            con.Execute "INSERT INTO FOrderItems(OrderNo,ItemCode,CompItemCode,Qty,DeliveryDT,Stamps,Quality,DeliveryStatus) VALUES('Stock-Order','" & cmbArticle.ID & _
             "','" & cmbArticle.ID & "'," & Val(txtQty) & ",'" & Date & "','','',0)"
        Else
            con.Execute "INSERT INTO FOrderItems(OrderNo,ItemCode,CompItemCode,Qty) VALUES('Stock-Order','" & cmbArticle.ID & _
             "','" & cmbArticle.ID & "'," & Val(txtQty) & ")"
        End If
    Else
        con.Execute "UPDATE FOrderItems SET Qty=Qty+" & Val(txtQty) & " WHERE OrderNo='Stock-Order' AND ItemCode='" & cmbArticle.ID & "'"
    End If
    
    Dim lEntryID As Long, lRcvEntryID As Long, lNextProcessID As Long
    lEntryID = GetSingleLongValue("MAX(EntryID)", "StockOrderOpening")
    con.Execute "INSERT INTO VendReceived(VendID,DT,RecieptID,UserID,ProcessID,Issuance_RefID) VALUES(" & _
     "0,'" & Date & "','',0," & cmbProcess.ID & ",0)"
    
    lRcvEntryID = GetSingleLongValue("MAX(EntryID)", "VendReceived")
    
    lNextProcessID = getNextProcessID(cmbArticle.ID, cmbProcess.ID)
    
    con.Execute "INSERT INTO VendRcvdDetail(RefID,ItemCode,RecieptID,RcvdQty,Rate,LotNo,NextProcessID,OrderNo" & _
     ",CountedBy,Issue_RefID,ProcessID,RcvdWeight,Opening_RefID) VALUES(" & lRcvEntryID & ",'" & cmbArticle.ID & _
     "',''," & Val(txtQty) & ",0,'0'," & lNextProcessID & ",'Stock-Order','',0," & cmbProcess.ID & _
     ",0," & lEntryID & ")"
     
    con.CommitTrans
    
    Call RefreshList
    
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

Private Sub Form_Load()

    iColNo = 0
    cmbArticle.ListHeight = 5800
    
    Call cmbArticle.AddVals(con, "'{ ' + ItemID + ' } ' + ItemName ", "Items", "ItemID", " WHERE InActive=0")
        
    cmbProcess.ListHeight = 4000
    
    Call AddToCombo(cmbLocation, "DISTINCT Location", "IssItemsSimpleDetail")
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    FAcc.Top = (Me.ScaleHeight - FAcc.Height) / 2
   
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If LV.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Delete ?", vbQuestion + vbYesNo) = vbYes Then
            'First Check If Issuance Have been made against it...
            Dim strItemID As String, lProcessID As Long, lEntryID As Long
            strItemID = LV.SelectedItem.Text
            lProcessID = Val(LV.SelectedItem.ListSubItems(2).Tag)
            'Now Get The EntryID FROM StockOrderOpening...
            lEntryID = GetSingleLongValue("EntryID", "StockOrderOpening", " WHERE ProcessID=" & lProcessID & " AND ItemID='" & strItemID & "'")
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
            con.Execute "DELETE FROM FOrderItems WHERE CompItemCode='" & strItemID & "' AND OrderNo='Stock-Order'"
            
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
    
    With rs
        .Open "SELECT * FROM VStockOrderOpening WHERE ItemID='" & cmbArticle.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , !ITemID)
            ITM.ListSubItems.add , , !ItemName & ""
            ITM.ListSubItems.add(, , !Description & "").Tag = Val(!ProcessID & "")
            ITM.ListSubItems.add(, , Val(!Qty & "")).Bold = True
            ITM.ListSubItems.add , , !Location & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
