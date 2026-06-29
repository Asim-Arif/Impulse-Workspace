VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmStockStatusFinalOld 
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   13035
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleMode       =   0  'User
   ScaleWidth      =   11675.62
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImgList 
      Left            =   5460
      Top             =   180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStockStatusFinalOld.frx":0000
            Key             =   "Yes"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStockStatusFinalOld.frx":0452
            Key             =   "No"
         EndProperty
      EndProperty
   End
   Begin ComboList.Usercontrol1 cmbCust 
      Height          =   285
      Left            =   900
      TabIndex        =   6
      Top             =   405
      Width           =   2310
      _ExtentX        =   4075
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
   Begin VB.Frame FLV 
      Height          =   6255
      Left            =   30
      TabIndex        =   0
      Top             =   780
      Width           =   12615
      Begin VB.Frame fButs1 
         Height          =   585
         Left            =   8295
         TabIndex        =   3
         Top             =   5475
         Width           =   3045
         Begin MSForms.CommandButton cmdClose 
            Height          =   360
            Left            =   1530
            TabIndex        =   5
            Top             =   165
            Width           =   1425
            ForeColor       =   0
            Caption         =   " Close      "
            PicturePosition =   327683
            Size            =   "2514;635"
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
            Left            =   60
            TabIndex        =   4
            Top             =   165
            Width           =   1425
            ForeColor       =   0
            Caption         =   "Print         "
            PicturePosition =   327683
            Size            =   "2514;635"
            Picture         =   "frmStockStatusFinalOld.frx":0564
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
      Begin MSComctlLib.ListView LV 
         Height          =   5265
         Left            =   105
         TabIndex        =   1
         Top             =   195
         Width           =   12390
         _ExtentX        =   21855
         _ExtentY        =   9287
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Order No."
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item Code"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item Description"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Order Qty"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Completed"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Balance"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSForms.ComboBox cmbOrders 
      Height          =   285
      Left            =   3195
      TabIndex        =   7
      Top             =   405
      Width           =   1860
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3281;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   885
      TabIndex        =   2
      Top             =   165
      Width           =   4155
      BackColor       =   11517387
      Caption         =   " Customer                                   Orders"
      Size            =   "7329;423"
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
Attribute VB_Name = "frmStockStatusFinalOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PrintSQL As String
Dim ProcessID As Long
Dim CustAccNo As String

Private Sub cmbCust_matched()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    
    With rs
        
        cmbOrders.Clear
        If cmbCust.ID = "0" Then
            cmbOrders.AddItem "<All Hico Orders>"
            Sql = "Select OrderNo From FCustomerOrders Where CustCode='HICO'"
        Else
            Sql = "Select OrderNo From FCustomerOrders Where CustCode='" & cmbCust.ID & "'"
        End If
        
        
        .Open Sql, con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            cmbOrders.AddItem ![OrderNo] & ""
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
        
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbOrders_Change()
    Call cmbOrders_Click
End Sub

Private Sub cmbOrders_Click()
    Call RefreshLV
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    Screen.MousePointer = vbHourglass
    
    Dim rpt As craxddrt.Report 'New rptStockStatus
    Set rpt = rptApp.OpenReport(RptPath & "\rptStockStatus.rpt") '
    rpt.SQLQueryString = PrintSQL
    
    Dim F As New frmPrevRpt
    Load F
    
    F.ShowReport " ", rpt
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
err:
    MsgBox err.Description
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Load()

    
    cmbCust.ListHeight = 4400
    cmbCust.AddItem "<Free Stock>", "0"
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where CustCode<>'HICO'")
    
    
    
End Sub
Private Sub Form_Resize()

    On Error Resume Next
     
    FLV.Move 50, 800, Me.ScaleWidth - 100, Me.ScaleHeight - 900
     
    LV.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
     
    fButs1.Move LV.Left + LV.Width - fButs1.Width, LV.Top + LV.Height
    
End Sub

Private Sub RefreshLV()
    
    On Error GoTo err
    Dim rsLedger As New ADODB.Recordset
    
    Dim Cond As String
    
    Dim ITM As ListItem
    
    If ProcessID = 13 Then
        Cond = " And ProcessID=12"
    Else
        Cond = " And ProcessID=" & ProcessID
    End If
    
    If cmbOrders.Text = "<All Hico Orders>" Then
        PrintSQL = "Select 'HICO Orders' As OrderNo,CompItemCode,ItemDescription,Sum(NetOrderQty) As NetOrderQty" & _
        ",Sum(NetQty) As NetQty,Sum(InStock) As InStock From VStockStatusRpt " & _
        "Where Left(OrderNo,10)='HICO Order' " & Cond & " Group By CompItemCode,ItemDescription Order By CompItemCode"
    Else
        PrintSQL = "Select * From VStockStatusRpt Where OrderNo='" & cmbOrders & "' " & Cond & " Order By CompItemCode"
    End If
    
    Dim lColor As Long
    
    With rsLedger
        .Open PrintSQL, con, adOpenForwardOnly, adLockReadOnly
        
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
                
            
            
            Set ITM = LV.ListItems.Add(, , ![OrderNo] & "")

            
            'itm.Tag = ![NotIssued] & ""
            ITM.ListSubItems.Add , , ![CompItemCode] & ""
            ITM.ListSubItems.Add , , ![ItemDescription] & ""
            ITM.ListSubItems.Add(, , Val(![NetOrderQty] & "")).Tag = Val(![NetQty] & "")
            ITM.ListSubItems.Add , , Val(![InStock] & "")
            
            ITM.ListSubItems.Add(, , Val(![NetOrderQty] & "") - Val(![InStock] & "")).ForeColor = vbRed
                        
            If Val(con.Execute("Select Count(*) From VStockStatusRpt Where OrderNo='" & ![OrderNo] & "' And ItemID='" & ![CompItemCode] & "'").Fields(0).value & "") = 1 Then
                ITM.ListSubItems.Add , , "Final"
            Else
                ITM.ListSubItems.Add , , "In Process"
            End If
                        
            .MoveNext
        Next
        .Close
    End With
    LV.Refresh
    Set rsLedger = Nothing
        
    Cond = ""
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Public Sub ShowStockStatus(lProcessID As Long)
    
    On Error GoTo err
    Me.Hide
    
    ProcessID = lProcessID
    
    
    PrintSQL = ""
    

    
    Me.Show
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo err
    Dim strRet As String
    If LV.ListItems.Count = 0 Then Exit Sub
    If Button <> 2 Then Exit Sub
    
    If Not (Val(LV.SelectedItem.SubItems(4)) > 0 And LV.SelectedItem.SubItems(6) = "Final") Then Exit Sub
    If cmbOrders.MatchFound = False Then Exit Sub
    If cmbCust.MatchFound = False Or cmbCust.ID = "0" Then Exit Sub
    
    strRet = ShowPopUpMenu1("&Complete Order Quantity")
    If strRet = "&Complete Order Quantity" Then
        Dim F As New frmTransQty
        Load F
        F.PrepareToShow LV.SelectedItem.Text, LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), Val(LV.SelectedItem.SubItems(5))
    End If
    
    Exit Sub
    'Another Check wil go here<<
    '<<<<<<<<<<<<<<<<<<<<<<<<<
    'Another Check wil go here<<
    
    If Val(LV.SelectedItem.SubItems(3)) < Val(LV.SelectedItem.SubItems(4)) Then Exit Sub
    
    
    
    strRet = ShowPopUpMenu1("Ship To Consignee")
    
    If strRet = "Ship To Consignee" Then
    
        If Val(LV.SelectedItem.Tag) <> 0 Then
        
            MsgBox "Not All Quantity Of This Item Has Been Issued.", vbInformation
            Exit Sub
            
        End If
            
        Dim OrderNo As String, ItemCode As String
        
        Dim lRefID As Long
        
        OrderNo = LV.SelectedItem.Text & ""
        ItemCode = LV.SelectedItem.SubItems(1) & ""
        
        Dim Rcvable As Long
        Rcvable = Val(con.Execute("Select Sum(Wastage+isnull(RcvRepairQty,0)) from VendRcvdDetail Where OrderNo='" & OrderNo & "' and ItemCode='" & ItemCode & "' Group By OrderNo,ItemCode").Fields(0).value & "")
        If Val(LV.SelectedItem.SubItems(3)) + Rcvable <> Val(LV.SelectedItem.ListSubItems(4).Tag) Then
            MsgBox "Can't Ship, Qty Not Completed Yet.", vbInformation
            Exit Sub
        End If
        
        
        Call StartTrans(con)
        
        con.Execute "Update VendRcvdDetail Set IssQty=RcvdQty Where OrderNo='" & _
         OrderNo & "' and ItemCode='" & ItemCode & "'"
            
        con.Execute "Insert Into VendReceived(DT,VendID) Values('" & _
        ServerDate & "',0)"
        
        lRefID = Val(con.Execute("Select Max(EntryID) From VendReceived").Fields(0).value & "")
        
        con.Execute "Insert Into VendRcvdDetail(RefID,RecieptID,OrderNo,ItemCode" & _
        ",NextProcessID,ProcessID,RcvdQty,Rate,ReqAuth) Values(" & lRefID & ",'','Free Stock','" & _
        ItemCode & "',13,12," & _
        Val(LV.SelectedItem.SubItems(3)) - Val(LV.SelectedItem.SubItems(4)) & ",0,0)"
            
        'Accounts Entries...
        
        'Export Entries....
            
        con.CommitTrans
        
    End If
    
    Exit Sub
    
err:
    MsgBox err.Description
End Sub
