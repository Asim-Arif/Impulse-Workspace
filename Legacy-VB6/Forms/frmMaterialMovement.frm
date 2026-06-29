VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmMaterialMovement 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6630
   ClientLeft      =   1095
   ClientTop       =   435
   ClientWidth     =   11565
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   11565
   Begin ComboList.Usercontrol1 cmbMaterial 
      Height          =   285
      Left            =   405
      TabIndex        =   8
      Top             =   900
      Width           =   10740
      _ExtentX        =   18944
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
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10275
      TabIndex        =   19
      Top             =   4350
      Width           =   780
   End
   Begin VB.Frame Frame2 
      Caption         =   "Move To :"
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
      Height          =   1260
      Left            =   405
      TabIndex        =   13
      Top             =   2910
      Width           =   10770
      Begin VB.TextBox txtInHandTo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9855
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   585
         Width           =   810
      End
      Begin MSForms.ComboBox cmbStore 
         Height          =   315
         Left            =   90
         TabIndex        =   18
         Top             =   585
         Width           =   3705
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "6535;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbRack 
         Height          =   315
         Left            =   3810
         TabIndex        =   17
         Top             =   585
         Width           =   2565
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4524;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbShelf 
         Height          =   315
         Left            =   6390
         TabIndex        =   16
         Top             =   585
         Width           =   3450
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "6085;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label4 
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   15
         Top             =   345
         Width           =   10575
         BackColor       =   11517387
         Caption         =   $"frmMaterialMovement.frx":0000
         Size            =   "18653;423"
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
   Begin VB.Frame Frame1 
      Caption         =   "Move From :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1200
      Left            =   405
      TabIndex        =   9
      Top             =   1485
      Width           =   10770
      Begin VB.TextBox txtInHand 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9855
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   585
         Width           =   810
      End
      Begin MSForms.Label Label4 
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   12
         Top             =   345
         Width           =   10575
         BackColor       =   11517387
         Caption         =   $"frmMaterialMovement.frx":00D6
         Size            =   "18653;423"
         BorderColor     =   16512
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbLocation 
         Height          =   285
         Left            =   90
         TabIndex        =   11
         Top             =   585
         Width           =   9750
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "17198;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame6 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   5850
      TabIndex        =   3
      Top             =   5430
      Width           =   5280
      Begin MSForms.CommandButton cmdSave 
         Height          =   360
         Left            =   75
         TabIndex        =   6
         Top             =   195
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
         Left            =   3510
         TabIndex        =   5
         Top             =   195
         Width           =   1665
         Caption         =   "Cancel            "
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
         Left            =   1785
         TabIndex        =   4
         Top             =   180
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
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Finished Stock Change Location"
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
         Left            =   2730
         TabIndex        =   1
         Top             =   135
         Width           =   4980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Finished Stock Change Location"
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
         Left            =   2745
         TabIndex        =   2
         Top             =   150
         Width           =   4980
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Qty to Move :"
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
      Height          =   195
      Left            =   9105
      TabIndex        =   20
      Top             =   4395
      Width           =   1125
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   405
      TabIndex        =   7
      Top             =   645
      Width           =   10740
      BackColor       =   11517387
      Caption         =   $"frmMaterialMovement.frx":01B4
      Size            =   "18944;423"
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
Attribute VB_Name = "frmMaterialMovement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCafeIssuance As Boolean


Private Sub cmbLocation_Change()
    Call cmbLocation_Click
End Sub

Private Sub cmbLocation_Click()
    If cmbLocation.MatchFound = False Then Exit Sub
    txtInHand = Val(cmbLocation.List(cmbLocation.ListIndex, 2))
End Sub


Private Sub cmbMaterial_matched()

    Dim rs As New ADODB.Recordset
    
    With rs
'        .Open "SELECT * FROM VMaterialLocationWiseBalances WHERE MaterialID='" & cmbMaterial.ID & "'", con, adOpenForwardOnly, adLockReadOnly
'        cmbLocation.Clear
'        Do Until .EOF
'            cmbLocation.AddItem !Store & " [" & !ShelfNo & "] (" & Val(!ShelfQty & "") & ")"
'            cmbLocation.List(.AbsolutePosition - 1, 1) = !Shelf_RefID
'            cmbLocation.List(.AbsolutePosition - 1, 2) = Val(!ShelfQty & "")
'            .MoveNEXT
'        Loop
'        .Close
        
        '.Open "SELECT VItemsWithShelfWiseStock.*,ShelfNo,StoreName FROM VItemsWithShelfWiseStock " & _
         "WHERE ItemID='" & cmbMaterial.ID & "' AND NetQty>0", con, adOpenForwardOnly, adLockReadOnly
            
        .Open "SELECT VItemsWithLotWiseStock.* FROM VItemsWithLotWiseStock " & _
         "WHERE ItemID='" & cmbMaterial.ID & "' AND NetQty>0", con, adOpenForwardOnly, adLockReadOnly
         
        cmbLocation.Clear
        
        Do Until .EOF
            cmbLocation.AddItem "Lot:" & !LotNo & ", Batch No.:" & !Batch_No & ", Location:" & !StoreName & "[" & !RackNo & "] [" & !ShelfNo & "] (" & Round(Val(!NetQty & "")) & ")"
            'cmbLocation.List(.AbsolutePosition - 1, 1) = !EntryID
            cmbLocation.List(.AbsolutePosition - 1, 1) = !Shelf_RefID
            cmbLocation.List(.AbsolutePosition - 1, 2) = Val(!NetQty & "")
            cmbLocation.List(.AbsolutePosition - 1, 3) = !Remarks & ""
            cmbLocation.List(.AbsolutePosition - 1, 4) = !EntryID & ""
            cmbLocation.List(.AbsolutePosition - 1, 5) = !LotNo & ""
            cmbLocation.List(.AbsolutePosition - 1, 6) = !Batch_No & ""
            cmbLocation.List(.AbsolutePosition - 1, 7) = !Mill_Certificate_No & ""
            .MoveNEXT
        Loop
        .Close
        
    End With
    Set rs = Nothing
    
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbPrinter_Change()
    txtAttn = cmbPrinter.Text
End Sub

Private Sub cmdAdd_Click()

    On Error GoTo err
    
    If cmbMaterial.MatchFound = False Then Exit Sub
    If cmbLocation.MatchFound = False Then Exit Sub
    
    If Val(txtQty) <= 0 Then Exit Sub
    If Val(txtQty) > Val(txtInHand) Then
        MsgBox "Can't Issue This Quantity." & vbNewLine & "Only " & txtInHand & " " & txtUnit & " Is Available In The Stock.", vbInformation
        Exit Sub
    ElseIf Val(txtQty) > Val(txtUnit.Tag) Then
        MsgBox "Can't Issue This Quantity." & vbNewLine & "Only " & txtUnit.Tag & " " & txtUnit & " Is Available In The Stock.", vbInformation
        Exit Sub
'    ElseIf txtOrderNo = "" Then
'        MsgBox "Please Enter Order No.", vbInformation
'        Exit Sub
    End If
    
    Dim ITM As ListItem
    Set ITM = LV.ListItems.add(, cmbMaterial.ID, LV.ListItems.count + 1)
    ITM.Tag = Val(txtUnit.Tag)
    ITM.ListSubItems.add(, , cmbMaterial.ID).Tag = cmbMaterial.Tag 'Rate For Printers
    ITM.ListSubItems.add(, , cmbMaterial.Text).Tag = Val(txtInHand.Tag)
    ITM.ListSubItems.add , , txtUnit
    ITM.ListSubItems.add , , Val(txtQty)
    ITM.ListSubItems.add , , txtOrderNo.Text
    ITM.ListSubItems.add(, , cmbLocation.Text).Tag = Val(cmbLocation.List(cmbLocation.ListIndex, 1))
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbRack_Change()
    Call cmbRack_Click
End Sub

Private Sub cmbRack_Click()

    On Error GoTo err
    
    If cmbRack.MatchFound = False Then Exit Sub
    Call AddToCombo(cmbShelf, "ShelfNo", "StoreShelfs", " WHERE Rack_RefID=" & cmbRack.List(cmbRack.ListIndex, 1), , "EntryID")
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbShelf_Change()
    Call cmbShelf_Click
End Sub

Private Sub cmbShelf_Click()

    Dim dQty As Double
    If cmbShelf.MatchFound = False Then Exit Sub
    dQty = GetSingleDoubleValue("NetQty", "VItemsWithShelfWiseStock", " WHERE Shelf_RefID=" & cmbShelf.List(cmbShelf.ListIndex, 1) & " AND ItemID='" & cmbMaterial.ID & "'")
    txtInHandTo = Round(dQty, 2)
    
End Sub

Private Sub cmbStore_Change()
    Call cmbStore_Click
End Sub

Private Sub cmbStore_Click()

    On Error GoTo err
    
    If cmbStore.MatchFound = False Then Exit Sub
    Call AddToCombo(cmbRack, "RackNo", "StoreRacks", " WHERE Store_RefID=" & cmbStore.List(cmbStore.ListIndex, 1), , "EntryID")
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Saved Then
       cmbLocation.ListIndex = -1
       cmbStore.ListIndex = -1
       cmbShelf.ListIndex = -1
       cmbRack.ListIndex = -1
       txtInHand = ""
       txtInHandTo = ""
       txtQty = ""
       cmbMaterial.ID = ""
       cmbMaterial.ID = ""
    End If
End Sub

Private Sub cmdSaveClose_Click()

    If Saved Then Unload Me
    
End Sub


Public Function Saved() As Boolean

    On Error GoTo err
    
    If cmbMaterial.MatchFound = False Then
        MsgBox "Invalid Material.", vbInformation
        Exit Function
    ElseIf cmbLocation.MatchFound = False Then
        MsgBox "Invalid From Location.", vbInformation
        Exit Function
    ElseIf cmbShelf.MatchFound = False Then
        MsgBox "Invalid Shelf.", vbInformation
        Exit Function
    ElseIf Val(txtQty) > Val(txtInHand) Or Val(txtQty) <= 0 Then
        MsgBox "Invalid Qty.", vbInformation
        Exit Function
    End If
    
    Dim strIssNo As String, strRcvNo As String, dInHand As Double, dQty As Double
    strIssNo = getIssuanceNo(Date)
    strRcvNo = GetRcvingNo(Date)
    
    Dim p_lEntryID As Long, strLotNo As String, strBatchNo As String, strMilNo As String
    p_lEntryID = cmbLocation.List(cmbLocation.ListIndex, 4)
    strLotNo = cmbLocation.List(cmbLocation.ListIndex, 5)
    strBatchNo = cmbLocation.List(cmbLocation.ListIndex, 6)
    strMilNo = cmbLocation.List(cmbLocation.ListIndex, 7)
    
    'Dim rs As New ADODB.Recordset
    '    With rs
'        .Open "SELECT T1.EntryID,TP.RcvdQty-ISNULL(TIssd.Qty,0) AS NetQty FROM RcvItemsSimpleDetail T1" & _
'        " LEFT JOIN (SELECT Rcv_ISD_RefID,SUM(Qty) AS Qty FROM IssItemsSimpleDetail GROUP BY Rcv_ISD_RefID) TIssd ON T1.EntryID=TIssd.RCV_ISD_RefID " & _
'        " LEFT JOIN RcvItemsSimpleDetail_Placement TP ON T1.EntryID=TP.RISD_RefID WHERE T1.ItemID='" &
'        , con, adOpenForwardOnly, adLockReadOnly
'        .Close
'    End With
    
    Call StartTrans(con)
    'Material Issuance
    
    con.Execute "INSERT INTO IssItemsSimple(IssNo,DT,OrderNo,Department,AttnPerson,UserName,ComputerName,SampleIssuance) Values('" & _
     strIssNo & "','" & Date & " " & ServerTime & "','For Location Change','','','" & CurrentUserName & _
     "','" & strComputerName & "',0)"
    
    lRefID = con.Execute("SELECT Max(EntryID) From IssItemsSimple").Fields(0).Value & ""
    
    Dim strMaterialID As String
    Dim lIISD_EntryID As Long
    
    strMaterialID = cmbMaterial.ID
    dQty = Val(txtQty)
    dBal = getItemBalance(strMaterialID, Date)
    dBal = dBal - dQty
    theDesc = "For Material Location Change"
    
    con.Execute "INSERT INTO ItemLedgerSimple(ItemID,DT,Description,Qty,Balance,OrderNo,Location,Remarks) Values('" & _
     strMaterialID & "','" & Dated & "','" & theDesc & "'," & -dQty & "," & dBal & ",'','" & cmbLocation.Text & "','')"
     
    lRefID = Val(con.Execute("Select Max(EntryID) From ItemLedgerSimple").Fields(0).Value & "")
            
    Call adjustItemLedger(strMaterialID, Date, -dQty)
    
    
    con.Execute "INSERT INTO IssItemsSimpleDetail(IssNo,ItemID,Qty,InHand,LedgerRefID,OrderNo,Location,Remarks,LotNo_Manual,BatchNo_Manual,RCV_ISD_RefID) Values('" & _
     strIssNo & "','" & strMaterialID & "'," & dQty & "," & getItemBalance(cmbMaterial.ID, Date) - dQty & "," & lRefID & _
     ",'','" & cmbLocation.Text & "','','" & strLotNo & "','" & strBatchNo & "'," & p_lEntryID & ")"
             
    lIISD_EntryID = GetSingleLongValue("MAX(EntryID)", "IssItemsSimpleDetail")
    
    
    con.Execute "INSERT INTO IssItemsSimpleDetail_More(IISD_RefID,IssdQty,Shelf_RefID,Remarks) VALUES(" & _
     lIISD_EntryID & "," & dQty & "," & Val(cmbLocation.List(cmbLocation.ListIndex, 1)) & ",'')"
            
    'Update Remarks in RcvItemsSimpleDetail_Placement; Last Entry of selected material @ selected Location...
    Dim lEntryID As Long
    lEntryID = GetSingleLongValue("MAX(EntryID)", "VRcvItemsSimpleDetail_Placement", " WHERE ItemID='" & strMaterialID & "' AND Shelf_RefID=" & Val(cmbLocation.List(cmbLocation.ListIndex, 1)))
    'con.Execute "UPDATE RcvItemsSimpleDetail_Placement SET Remarks='" & .SubItems(7) & "' WHERE EntryID=" & lEntryID
            
    Call UpdateItemLocation(strMaterialID, Val(cmbLocation.List(cmbLocation.ListIndex, 1)), dQty)
            'con.Execute "Update Material Set InHand=InHand-" & Val(.SubItems(4)) & " Where MaterialID='" & .key & "'"
            
    'Now Receiving...
    con.Execute "INSERT INTO RcvItemsSimple(RcvNo,DT,OrderNo,RcvdBy,RcvdFrom,UserName,ComputerName) VALUES('" & _
     strRcvNo & "','" & Date & " " & ServerTime & "','Material Movement','','','" & CurrentUserName & _
     "','" & strComputerName & "')"
    
    lRefID = con.Execute("Select Max(EntryID) From RcvItemsSimple").Fields(0).Value & ""
    
    Dim lRISD_EntryID As Long
    
    dBal = getItemBalance(strMaterialID, Date)
    dBal = dBal + dQty
            
    theDesc = "Material Movement"
            
    con.Execute "INSERT INTO ItemLedgerSimple(ItemID,DT,Description,Qty,Balance,OrderNo,Location,Remarks) VALUES('" & _
     strMaterialID & "','" & Date & "','" & theDesc & "'," & dQty & "," & dBal & ",'','" & cmbRack.Text & " {" & cmbShelf.Text & "}','')"
             
    lRefID = Val(con.Execute("Select Max(EntryID) From ItemLedgerSimple").Fields(0).Value & "")
            
    Call adjustItemLedger(strMaterialID, Date, dQty)

    con.Execute "INSERT INTO RcvItemsSimpleDetail(RcvNo,ItemID,Qty,InHand,LedgerRefID,OrderNo,Location,Remarks,LotNo_Manual,BatchNo_Manual,Mill_Certificate_No) Values('" & _
     strRcvNo & "','" & strMaterialID & "'," & dQty & "," & getItemBalance(cmbMaterial.ID, Date) - dQty & "," & lRefID & _
     ",'','','','" & strLotNo & "','" & strBatchNo & "','" & strMilNo & "')"
            
    lRISD_EntryID = GetSingleLongValue("MAX(EntryID)", "RcvItemsSimpleDetail")
            
    con.Execute "INSERT INTO RcvItemsSimpleDetail_Placement(RISD_RefID,Shelf_RefID,RcvdQty,Remarks) VALUES(" & _
     lRISD_EntryID & "," & cmbShelf.List(cmbShelf.ListIndex, 1) & "," & dQty & ",'" & cmbLocation.List(cmbLocation.ListIndex, 3) & "')"
            
    con.CommitTrans
    
'    con.Execute "INSERT INTO MaterialMovements(MLS_From_RefID,MLS_To_RefID,UserName,MachineName) VALUES(" & _
'     cmbLocation.List(cmbLocation.ListIndex, 1) & "," & lRefID & ",'" & CurrentUserName & "','" & strComputerName & _
'     "')"
'
    
    Saved = True
    
    Exit Function
err:
    MsgBox err.Description
End Function
Private Sub Form_Load()

    'DT = Date
    'Call DT_Change

'    Call RestrictIssRcvDates(DT)
    DT = Date
    
    cmbMaterial.ListHeight = 3000
    
    Dim strCondition As String
    If CurrentUserName <> "Administrator" Then
        strCondition = " WHERE EntryID IN(SELECT Store_RefID FROM Users_Stores WHERE UserID=" & UserID & ")"
    End If
    Call AddToCombo(cmbStore, "StoreName", "Stores", strCondition & " ORDER BY StoreName", , "EntryID")
    
    cmbMaterial.DropDownOffLine = True
    cmbMaterial.AddVals con, "'{' + ItemID + '} ' + ItemName", "Items", "ItemID"
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
    
    Dim TitleBottom As Long

    TitleBottom = cmbMaterial.Top + cmbMaterial.Height + 50 '+ fTitle.Height
 
End Sub


Public Sub ShowMe(Optional p_bCafeIssuance As Boolean = False)

    bCafeIssuance = p_bCafeIssuance
    If bCafeIssuance Then
        cmbDept.ID = "VCAFE"
        cmbDept.Enabled = False
        'txtAttn.Text = "Shabana"
        'txtAttn.Locked = True
    End If
    
    Me.Show 1
    
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)

    If LV.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to remove this entry?", vbQuestion + vbYesNo) = vbYes Then
            LV.ListItems.Remove LV.SelectedItem.Index
        End If
    End If
End Sub
