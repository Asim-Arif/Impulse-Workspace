VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmRMMovement 
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
   Begin ComboList.Usercontrol1 cmbMaterialID 
      Height          =   285
      Left            =   405
      TabIndex        =   7
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
         Caption         =   $"frmRMMovement.frx":0000
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
         Caption         =   $"frmRMMovement.frx":00D6
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
         Caption         =   "Change Material Location"
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
         Left            =   3330
         TabIndex        =   1
         Top             =   135
         Width           =   4080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change Material Location"
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
         Left            =   3345
         TabIndex        =   2
         Top             =   150
         Width           =   4080
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
      TabIndex        =   8
      Top             =   645
      Width           =   10740
      BackColor       =   11517387
      Caption         =   " Material"
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
Attribute VB_Name = "frmRMMovement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCafeIssuance As Boolean

Private Sub cmbdept_matched()
    
    If cmbDept.ID = "CNT" Or Left(cmbDept.ID, 4) = "PRNT" Then 'Printing Contractors
        cmbPrinter.Visible = True
        txtAttn.Visible = False
        cmbMaterialID.AddVals con, "MaterialID", "Material", "MaterialID", " WHERE ISNULL(RateForPrinters,0)>0 AND CatID<>'14' AND (LEFT(MaterialID,2)<>'32') AND InActive=0 Order By MaterialID"
        cmbMaterial.AddVals con, "MaterialName", "Material", "MaterialID", " WHERE ISNULL(RateForPrinters,0)>0 AND CatID<>'14' AND (LEFT(MaterialID,2)<>'32') AND InActive=0 Order By MaterialID"
        If cmbDept.ID = "CNT" Then
            cmbPrinter.Text = "Nadeem Khaliq Butt"
        ElseIf cmbDept.ID = "PRNT" Then
            cmbPrinter.Text = "Sohail Ikram"
        ElseIf cmbDept.ID = "PRNT1" Then
            cmbPrinter.Text = "Hafiz Afzal"
        ElseIf cmbDept.ID = "PRNT2" Then
            cmbPrinter.Text = "Hafiz Shafique"
        End If
    Else
        cmbPrinter.Visible = False
        txtAttn.Visible = True
        Dim strCondition As String
        If bCafeIssuance Then
            strCondition = " WHERE LEFT(MaterialID,2)='32' AND InActive=0"
        Else
            strCondition = " WHERE CatID<>'14' AND InActive=0"
        End If
        strCondition = strCondition & " AND (LEFT(MaterialID,2)<>'32')" '" AND (LEFT(MaterialID,2) IN('06','07','08') OR CatID IN('05-01','05-02','05-03','05-04','05-05','05-08'))"
        cmbMaterialID.AddVals con, "MaterialID", "Material", "MaterialID", strCondition & "  Order By MaterialID"
        cmbMaterial.AddVals con, "MaterialName", "Material", "MaterialID", strCondition & " Order By MaterialID"
    End If
    
End Sub

Private Sub cmbLocation_Change()
    Call cmbLocation_Click
End Sub

Private Sub cmbLocation_Click()
    If cmbLocation.MatchFound = False Then Exit Sub
    txtInHand = Val(cmbLocation.List(cmbLocation.ListIndex, 3))
End Sub

Private Sub cmbMaterial_matched()
    cmbMaterialID.ID = cmbMaterialID.ID
End Sub

Private Sub cmbMaterialID_matched()

    Dim rs As New ADODB.Recordset
        
    With rs
    
        .Open "SELECT MaterialLocationwiseStatus.*,ShelfNo,VStoreShelfs.StoreName FROM MaterialLocationwiseStatus INNER JOIN VendRcvdDetailPO ON " & _
            "MaterialLocationwiseStatus.Rcvd_RefID=VendRcvdDetailPO.EntryID INNER JOIN VStoreShelfs ON " & _
            "MaterialLocationwiseStatus.Shelf_RefID=VStoreShelfs.EntryID WHERE MaterialID='" & cmbMaterialID.ID & _
            "' AND ROUND(QtyPlaced,2)>ROUND(QtyIssued,2)", con, adOpenForwardOnly, adLockReadOnly
            
        cmbLocation.Clear
        Do Until .EOF
            cmbLocation.AddItem !StoreName & " [" & !ShelfNo & "] (" & Round(Val(!QtyPlaced & "") - Val(!QtyIssued & ""), 2) & ")" & " Batch: " & !BatchNo & ")"
            cmbLocation.List(.AbsolutePosition - 1, 1) = !EntryID
            cmbLocation.List(.AbsolutePosition - 1, 2) = !Shelf_RefID
            cmbLocation.List(.AbsolutePosition - 1, 3) = Round(Val(!QtyPlaced & "") - Val(!QtyIssued & ""), 2)
            cmbLocation.List(.AbsolutePosition - 1, 4) = !Rcvd_RefID
            cmbLocation.List(.AbsolutePosition - 1, 5) = !BatchNo & ""
            cmbLocation.List(.AbsolutePosition - 1, 6) = !LotNo & ""
            .MoveNEXT
        Loop
        .Close
        
'        .Open "Select Unit,InHand From RM WHERE RMID1='" & cmbMaterialID.ID & "'", con, adOpenForwardOnly, adLockReadOnly
'        txtUnit = !Unit & ""
'        txtUnit.Tag = Val(!InHand & "")
'        txtInHand = Val(!InHand & "")
'        'txtInHand.Tag = Val(!AvgRate & "")
'        .Close
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
    Set ITM = LV.ListItems.add(, cmbMaterialID.ID, LV.ListItems.count + 1)
    ITM.Tag = Val(txtUnit.Tag)
    ITM.ListSubItems.add(, , cmbMaterialID.ID).Tag = cmbMaterial.Tag 'Rate For Printers
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
    dQty = GetSingleDoubleValue("ShelfQty", "VMaterialLocationWiseBalances", " WHERE Shelf_RefID=" & cmbShelf.List(cmbShelf.ListIndex, 1) & " AND MaterialID='" & cmbMaterialID.ID & "'")
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
       cmbMaterialID.ID = ""
       cmbMaterialID.ID = ""
    End If
End Sub

Private Sub cmdSaveClose_Click()

    If Saved Then Unload Me
    
End Sub


Public Function Saved() As Boolean
    On Error GoTo err
    
    If cmbMaterialID.MatchFound = False Then
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
     
    Call StartTrans(con)
    
    con.Execute "INSERT INTO MaterialLocationwiseStatus(Rcvd_RefID,Shelf_RefID,QtyPlaced,UserName,MachineName,BatchNo,LotNo) VALUES(" & _
     cmbLocation.List(cmbLocation.ListIndex, 4) & "," & cmbShelf.List(cmbShelf.ListIndex, 1) & "," & Val(txtQty) & _
     ",'" & CurrentUserName & "','" & strComputerName & "','" & cmbLocation.List(cmbLocation.ListIndex, 5) & "','" & cmbLocation.List(cmbLocation.ListIndex, 6) & "')"
     
    con.Execute "UPDATE MaterialLocationwiseStatus SET QtyIssued=QtyIssued+" & Val(txtQty) & " WHERE EntryID=" & _
     cmbLocation.List(cmbLocation.ListIndex, 1)
    
    Dim lRefID As Long
    lRefID = GetSingleLongValue("MAX(EntryID)", "MaterialLocationwiseStatus")
    
    con.Execute "INSERT INTO MaterialMovements(MLS_From_RefID,MLS_To_RefID,UserName,MachineName,QtyMoved) VALUES(" & _
     cmbLocation.List(cmbLocation.ListIndex, 1) & "," & lRefID & ",'" & CurrentUserName & "','" & strComputerName & _
     "'," & Val(txtQty) & ")"
     
    con.CommitTrans
    
    Saved = True
    
    Exit Function
err:
    MsgBox err.Description
End Function

Private Sub Form_Load()

    DT = Date
    
    cmbMaterialID.ListHeight = 2400
        
    Dim strCondition As String
    If CurrentUserName <> "Administrator" Then
        strCondition = " WHERE EntryID IN(SELECT Store_RefID FROM Users_Stores WHERE UserID=" & UserID & ")"
    End If
    Call AddToCombo(cmbStore, "StoreName", "Stores", strCondition & " ORDER BY StoreName", , "EntryID")
    cmbMaterialID.AddVals con, "'{' + RMID1 + '} ' + RMName", "RM", "RMID1", " ORDER BY RMName"
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
    
    Dim TitleBottom As Long

    TitleBottom = cmbMaterialID.Top + cmbMaterialID.Height + 50 '+ fTitle.Height
 
End Sub


Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)

    If LV.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to remove this entry?", vbQuestion + vbYesNo) = vbYes Then
            LV.ListItems.Remove LV.SelectedItem.Index
        End If
    End If
End Sub
