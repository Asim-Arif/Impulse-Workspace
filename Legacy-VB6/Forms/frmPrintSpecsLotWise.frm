VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmPrintSpecsLotWise 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Specification Sheets Lotwise"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpenTechDrawing 
      Caption         =   "Open Tech. Drawing"
      Height          =   465
      Left            =   4590
      TabIndex        =   12
      Top             =   7635
      Width           =   1860
   End
   Begin VB.TextBox txtLotQty 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4305
      TabIndex        =   10
      Top             =   135
      Width           =   1065
   End
   Begin VB.CommandButton cmdAddLots 
      Caption         =   "Generate Lots"
      Height          =   465
      Left            =   5445
      TabIndex        =   9
      Top             =   75
      Width           =   1530
   End
   Begin VB.TextBox txtEdit 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   5055
      TabIndex        =   8
      Top             =   2190
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton cmdIssue 
      Caption         =   "&Issue"
      Height          =   465
      Left            =   9480
      TabIndex        =   7
      Top             =   75
      Width           =   1530
   End
   Begin VB.CommandButton cmdUnCheckAll 
      Caption         =   "Un-Check All"
      Height          =   465
      Left            =   1620
      TabIndex        =   6
      Top             =   7635
      Width           =   1530
   End
   Begin VB.CommandButton cmdCheckAll 
      Caption         =   "Check All"
      Height          =   465
      Left            =   60
      TabIndex        =   5
      Top             =   7635
      Width           =   1530
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   465
      Left            =   7890
      TabIndex        =   4
      Top             =   7635
      Width           =   1530
   End
   Begin VB.TextBox txtIPONo 
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
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   735
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   135
      Width           =   1980
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   465
      Left            =   9480
      TabIndex        =   1
      Top             =   7635
      Width           =   1530
   End
   Begin MSComctlLib.ListView LV 
      Height          =   6945
      Left            =   60
      TabIndex        =   0
      Top             =   645
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   12250
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sr #"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Item Code"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Lot #"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Print"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Issue To"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Issuance"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Rate"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "RepairLot"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   7425
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Please Select Picture"
      Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Quantity :"
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
      Left            =   3420
      TabIndex        =   11
      Top             =   210
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "IPO # : "
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
      Left            =   135
      TabIndex        =   2
      Top             =   210
      Width           =   615
   End
End
Attribute VB_Name = "frmPrintSpecsLotWise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------
Dim lIPONo As Long
Dim iColNo As Integer

Public Sub ShowMe(p_lIPONo As Long)

    On Error GoTo err
    
    txtIPONo = p_lIPONo
    lIPONo = p_lIPONo
     
    Call FillLots
    
    Me.Show , MainForm
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub FillLots()

    Dim rs As New ADODB.Recordset, ITM As ListItem
    Dim lForeColor As Long
    With rs
        .Open "SELECT * FROM VInternalProductionOrderDetail WHERE IPONo=" & lIPONo & " ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, Val(!EntryID & "") & "'", LV.ListItems.count + 1)
            ITM.Tag = Val(!EntryID & "")
            ITM.ListSubItems.add(, , !CompItemCode & "").Tag = Val(!ProcessID & "")
            ITM.ListSubItems.add(, , Val(!EntryID & "")).Tag = Val(!Rate & "")
            ITM.ListSubItems.add(, , Val(!LotQty & "")).Tag = !CustomerPO & ""
           
            ITM.ListSubItems.add(, , IIf(!PrintDone, "Yes", "No")).Tag = Val(!Operation & "")
            
            If !PrintDone = False Then
                ITM.Checked = True
            End If
            
            'Fix Maker
            Dim lMakerID As Long, strMakerName As String
            lMakerID = Val(!Fix_Maker_RefID & "")
            strMakerName = !VenderName & ""
            ITM.ListSubItems.add(, , strMakerName).Tag = lMakerID
            'Check If Already Issued....
            Dim lCount As Long
            lCount = GetSingleLongValue("COUNT(IPODL_RefID)", "VendIssdDetail", " WHERE IPODL_RefID=" & Val(!EntryID & ""))
            If lCount > 0 Then
                ITM.ListSubItems.add(, , "Done").Tag = 0
                lForeColor = vbBlue
                ITM.ListSubItems(6).ForeColor = lForeColor
            Else
                ITM.ListSubItems.add(, , "No").Tag = 0
                lForeColor = vbBlack
            End If
            ITM.ListSubItems.add(, , Val(!Rate & "")).Tag = ""
            ITM.ListSubItems.add().Tag = !RepairLot
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub cmdAddLots_Click()

    Dim lMaxLotQty As Long, lRefID As Long, lQty As Long
    
    lMaxLotQty = GetSingleLongValue("MaxLotSize", "ItemCatagories INNER JOIN Items ON Items.CatID=ItemCatagories.CatID INNER JOIN InternalProductionOrderDetail ON InternalProductionOrderDetail.CompItemCode=Items.ItemID INNER JOIN InternalProductionOrder ON InternalProductionOrder.EntryID=InternalProductionOrderDetail.RefID", "WHERE IPONo=" & lIPONo)
    lQty = Val(txtLotQty)
    lRefID = GetSingleLongValue("InternalProductionOrderDetail.EntryID", "InternalProductionOrderDetail INNER JOIN InternalProductionOrder ON InternalProductionOrder.EntryID=InternalProductionOrderDetail.RefID", "WHERE IPONo=" & lIPONo)
     
    'We Need to Last Lot If It's Qty is Less than Max. Lot & it has not been Issued...
    Dim lLastLotEntryID As Long, lLastLotQty As Long
   
    lLastLotEntryID = GetSingleLongValue("MAX(EntryID)", "InternalProductionOrderDetailLots", " WHERE IPOD_RefID=" & lRefID)
    lLastLotQty = GetSingleLongValue("LotQty", "InternalProductionOrderDetailLots", " WHERE EntryID=" & lLastLotEntryID)
     
    If lLastLotQty < lMaxLotQty Then
        'Check If it has been Issued..
        Dim lIssCount As Long
        lIssCount = GetSingleLongValue("COUNT(IPODL_RefID)", "VendIssdDetail", "WHERE IPODL_RefID=" & lLastLotEntryID)
        If lIssCount = 0 Then
            con.Execute "UPDATE InternalProductionOrderDetailLots SET LotQty=LotQty+" & (lMaxLotQty - lLastLotQty) & " WHERE EntryID=" & lLastLotEntryID
            lQty = lQty - (lMaxLotQty - lLastLotQty)
        End If
    End If
    
    Dim lLots As Long, lExtraLotQty As Long
    lExtraLotQty = 0
    lLots = Floor(lQty / lMaxLotQty)
    If lQty > (lLots * lMaxLotQty) Then
        lExtraLotQty = lQty - (lLots * lMaxLotQty)
    End If
    
    Dim i As Integer
    For i = 1 To lLots
        con.Execute "INSERT INTO InternalProductionOrderDetailLots(IPOD_RefID,LotQty) VALUES(" & _
         lRefID & "," & lMaxLotQty & ")"
    Next
    
    If lExtraLotQty > 0 Then
        con.Execute "INSERT INTO InternalProductionOrderDetailLots(IPOD_RefID,LotQty) VALUES(" & _
         lRefID & "," & lExtraLotQty & ")"
    End If
    Call FillLots

End Sub

Private Sub cmdCheckAll_Click()
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdClose_Click()
   
    Unload Me
    
End Sub

Private Sub cmdIssue_Click()

    Dim i As Integer
    Dim bIssuanceFound As Boolean
    bIssuanceFound = False
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .SubItems(6) = "Yes" Then
                If Val(.ListSubItems(2).Tag) <= 0 And Val(.ListSubItems(4).Tag) > 0 Then
                    MsgBox .SubItems(1) & " is not assigned to " & .SubItems(5) & " or the Rate is Zero.", vbInformation
                    Exit Sub
                End If
                bIssuanceFound = True
            End If
        End With
    Next
    
    If bIssuanceFound = False Then
        MsgBox "No Issuance Found.", vbInformation
        Exit Sub
    End If
    
    Call Saved
    
End Sub

Private Sub cmdOpenTechDrawing_Click()
    Call ShowTechDrawing("Items", "TechnicalDrawing", " WHERE ItemID='" & LV.SelectedItem.SubItems(1) & "'", LV.SelectedItem.SubItems(1), CD1)
End Sub

Private Sub cmdPrint_Click()
    
    Dim strEntryIDs As String
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                strEntryIDs = strEntryIDs & .Tag & ","
                .ListSubItems(1).Tag = True
            End If
        End With
    Next
    If strEntryIDs = "" Then
        MsgBox "No Lot Selected.", vbInformation
        Exit Sub
    Else
        strEntryIDs = Left(strEntryIDs, Len(strEntryIDs) - 1)
    End If
     
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ProductionSpecification.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{InternalProductionOrderDetailLots.EntryID} IN [" & strEntryIDs & "]", rpt
    
    DoEvents
    con.Execute "UPDATE InternalProductionOrderDetailLots SET PrintDone=1 WHERE EntryID IN(" & strEntryIDs & ")"
    Call FillLots
    
End Sub

Private Sub cmdUnCheckAll_Click()
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = False
    Next
End Sub

Private Sub LV_DblClick()
    
    If LV.ListItems.count = 0 Then Exit Sub
    
    If iColNo = 7 Then  'Issuance Yes or No.
        If CBool(LV.SelectedItem.ListSubItems(8).Tag) Then Exit Sub  'If Repair Lot, Don't Allow Issuance from this form...
        If LV.SelectedItem.SubItems(6) = "Yes" Then
            LV.SelectedItem.SubItems(6) = "No"
            LV.SelectedItem.ListSubItems(6).Tag = "0"
        ElseIf LV.SelectedItem.SubItems(6) = "No" Then
            LV.SelectedItem.SubItems(6) = "Yes"
            LV.SelectedItem.ListSubItems(6).Tag = "1"
        End If
        
'        Dim dRate As Double
'        If Val(LV.ListItems(1).ListSubItems(4).Tag) = 0 Then        'Ist Process is Factroy Employee
'            dRate = InputBox("Please Enter Rate", "Rate")
'            If dRate = 0 Then
'                MsgBox "Invalid Rate.", vbInformation
'                Exit Function
'            End If
'        Else
'            dRate = Val(LV.ListItems(1).ListSubItems(2).Tag)
'        End If
        
    ElseIf iColNo = 8 Then
        With txtEdit
            .Move LV.Left + LV.ColumnHeaders(8).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(8).Width
            .Text = Val(LV.SelectedItem.SubItems(7))
            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
        End With
    ElseIf iColNo = 4 And LV.ListItems.count = LV.SelectedItem.Index Then   'If Last, Allow Edit Qty.
        With txtEdit
            .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(iColNo).Width
            .Text = Val(LV.SelectedItem.SubItems(iColNo - 1))
            .SelStart = 0
            .SelLength = Len(.Text)
            .Visible = True
            .SetFocus
        End With
    End If
    
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
    If LV.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If LV.SelectedItem.SubItems(6) = "Done" Then
            MsgBox "Issuance Done. Can't Delete Now.", vbInformation
            Exit Sub
        End If
        If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo) = vbYes Then
            con.Execute "DELETE FROM InternalProductionOrderDetailLots WHERE EntryID=" & Val(LV.SelectedItem.key)
            LV.ListItems.Remove LV.SelectedItem.Index
        End If
    End If
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If x > Val(LV.ColumnHeaders(i).Left) And x < Val(LV.ColumnHeaders(i).Left) + Val(LV.ColumnHeaders(i).Width) Then
            iColNo = i
            Exit For
        End If
    Next
    
End Sub

Private Function Saved() As Boolean

    On Error GoTo err
    
    
    Dim SaveIt As Boolean
    
    Dim VendID As New Collection
    Dim VIssID As New Collection
     
    Dim myPOSavedCol As New Collection
    
    Dim cmd As New ADODB.Command
    Dim TempIndex As Integer
    TempIndex = 1
    
    Dim lRefID As Long
    
    Dim IssLot As Boolean
    Dim iAuth As Integer 'Authorize
    
    IssLot = False 'RcvingLot 'RcvLot(cmbProcess.ID)
    
    Dim UserReqAuth As Integer
    UserReqAuth = 0
    Call StartTrans(con)
    
    Dim myDT As Date, strSpecialInstructions As String, lExcess As Long, lRcvings As Long
    Dim iPriority As Integer, lProcessID As Long
    myDT = Date
    lExcess = 100
    lRcvings = 50
    iPriority = 0
    Dim i As Integer, j As Integer
    Dim lLotNo As String
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            SaveIt = True
            Dim myPOObjTemp As clsPOEntry
            For j = 1 To myPOSavedCol.count
                Set myPOObjTemp = myPOSavedCol(j)
                If myPOObjTemp.iIssueProcess = Val(.ListSubItems(7).Tag) And myPOObjTemp.iReturnProcess = Val(.ListSubItems(3).Tag) And myPOObjTemp.iVendID = Val(.ListSubItems(1).Tag) And myPOObjTemp.strItemCode = .SubItems(5) Then
                    SaveIt = False
                    lRefID = myPOObjTemp.lRefID
                    Set myPOObjTemp = Nothing
                    Exit For
                End If
                Set myPOObjTemp = Nothing
            Next
            
            If .SubItems(6) = "Yes" Then
                iPriority = iPriority + 1
                lProcessID = Val(.ListSubItems(1).Tag)
                Dim lMakerIssNo As Long, strMakerIssNo As String
                lMakerIssNo = Val(con.Execute("SELECT MAX(CAST(RIGHT(RecieptID,Len(RecieptID)-12) AS INT)) FROM VendIssued WHERE CONVERT(varchar,DT,6)='" & Format(myDT, "dd MMM yy") & "'").Fields(0).Value & "")
                lMakerIssNo = lMakerIssNo + 1
                'IssID = "ISU-" & Format(myDT, "ddMMyy") & lMakerIssNo
                strMakerIssNo = "M-ISU-" & Format(myDT, "ddMMyy") & lMakerIssNo
                
                con.Execute "INSERT INTO VendIssued(VendID,DT,RecieptID,UserID,ProcessID,ItemID," & _
                "UserName,MachineName,SpecialInstructions,ExcessQtyPercentage," & _
                "MaximumRcvingsAgainstPO) VALUES(" & Val(.ListSubItems(5).Tag) & ",'" & _
                 myDT & "','" & strMakerIssNo & "'," & UserID & "," & _
                 lProcessID & ",'" & .SubItems(1) & "','" & CurrentUserName & "','" & _
                 strComputerName & "','" & strSpecialInstructions & "'," & lExcess & "," & _
                 lRcvings & ")"
                 
                lRefID = Val(con.Execute("SELECT MAX(EntryID) FROM VendIssued").Fields(0).Value & "")
                
                'Add to Collection After Assigning to Object
                Dim myPOObj As New clsPOEntry
'                myPOObj.iIssueProcess = Val(.ListSubItems(7).Tag)
'                myPOObj.iReturnProcess = Val(.ListSubItems(3).Tag)
'                myPOObj.iVendID = Val(.ListSubItems(1).Tag)
'                myPOObj.strItemCode = .SubItems(5)
'                myPOObj.lRefID = lRefID
'                myPOSavedCol.add myPOObj
'                Set myPOObj = Nothing
'                VendID.add Val(LV.ListItems(i).ListSubItems(5).Tag)
'                VIssID.add lRefID
                lLotNo = GetSingleLongValue("MAX(LotNo)", "VendIssdDetail")
                lLotNo = lLotNo + 1
                 con.Execute "INSERT INTO VendIssdDetail(RefID,RecieptID,ItemCode" & _
                    ",Rate,IssQty,ReqAuth,OrderNo,RcvProcessID,ReturnDT,Priority,IPODL_RefID,LotNo) VALUES(" & lRefID & ",'" & _
                    GetIssuanceIDDetail(myDT) & "','" & .SubItems(1) & "'," & Val(.ListSubItems(2).Tag) & "," & _
                    Val(.SubItems(3)) & "," & UserReqAuth & ",'" & .ListSubItems(3).Tag & "'," & lProcessID & _
                    ",'" & DateAdd("d", 7, myDT) & "'," & iPriority & "," & Val(.key) & "," & lLotNo & ")"
                
                'Authentication Info...
                If UserReqAuth = 0 Then
                    Dim DetailRefID As Long
                    DetailRefID = Val(con.Execute("SELECT MAX(EntryID) FROM VendIssdDetail").Fields(0).Value & "")
                    con.Execute "INSERT INTO VendIssAuthDetail(UserID,IssID,DT) VALUES(" & _
                     UserID & "," & DetailRefID & ",'" & ServerDate & " " & ServerTime & "')"
                End If
                .SubItems(6) = "Done"
                .ListSubItems(6).ForeColor = vbBlue
            Else
                'lRefID = VIssID(j)
                'lRefID = 'strVendnItem(j, 2)
            End If
        
            '1. Insert Into VendIssuedItems
            
           
                 

            
         
        End With
        'con.Execute "INSERT INTO PrintSession(RecieptNo) Values('" & LVIssued.ListItems(i).ListSubItems(2).Tag & "')"
        'Debug.Print i & " of " & LVIssued.ListItems.count
    Next i
                       
    
    con.CommitTrans
    
    Saved = True
    Exit Function
err:
    MsgBox err.Description
    
End Function


Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Dim strItemID As String, lMakerID As Long, lProcessID As Long, iRecAffect As Integer
        strItemID = LV.SelectedItem.SubItems(1)
        lMakerID = Val(LV.SelectedItem.ListSubItems(5).Tag)
        lProcessID = Val(LV.SelectedItem.ListSubItems(1).Tag)
        iRecAffect = 0
        If iColNo = 4 Then
            con.Execute "UPDATE InternalProductionOrderDetailLots SET LotQty=" & Val(txtEdit) & " WHERE EntryID=" & Val(LV.SelectedItem.Tag)
            LV.SelectedItem.SubItems(3) = Val(txtEdit)
        Else
            con.Execute "UPDATE VendAssItems SET Rate=" & Val(txtEdit) & " WHERE ItemID='" & strItemID & _
             "' AND VendID=" & lMakerID, iRecAffect
            
            If iRecAffect = 0 Then
                con.Execute "INSERT INTO VendAssItems(VendID,ProcessID,ItemID,Rate,PlantRate,StampRate,SnaffRate,Unit)" & _
                 "VALUES(" & lMakerID & ",0,'" & strItemID & "'," & Val(txtEdit) & _
                 ",0,0,0,'" & GetSingleStringValue("Unit", "Items", " WHERE ItemID='" & strItemID & "'") & "')"
            End If
            LV.SelectedItem.SubItems(7) = Val(txtEdit)
            LV.SelectedItem.ListSubItems(2).Tag = Val(txtEdit)
            Dim i As Integer
            For i = 1 To LV.ListItems.count
                LV.ListItems(i).SubItems(7) = Val(txtEdit)
            Next
        End If
        
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
    
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub
