VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmMaterialMovementSF 
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
      Width           =   8655
      _ExtentX        =   15266
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
      Left            =   9075
      TabIndex        =   16
      Top             =   900
      Width           =   1935
      _ExtentX        =   3413
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
   Begin ComboList.Usercontrol1 cmbFromStore 
      Height          =   285
      Left            =   420
      TabIndex        =   14
      Top             =   1965
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   556
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextBackColor   =   16777215
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
      Appearance      =   0
   End
   Begin ComboList.Usercontrol1 cmbFromLocation 
      Height          =   285
      Left            =   2850
      TabIndex        =   15
      Top             =   1965
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   556
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextBackColor   =   16777215
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
      Appearance      =   0
   End
   Begin ComboList.Usercontrol1 cmbToStore 
      Height          =   285
      Left            =   420
      TabIndex        =   19
      Top             =   3525
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   556
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextBackColor   =   16777215
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
      Appearance      =   0
   End
   Begin ComboList.Usercontrol1 cmbToLocation 
      Height          =   285
      Left            =   2865
      TabIndex        =   20
      Top             =   3525
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   556
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextBackColor   =   16777215
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
      Appearance      =   0
   End
   Begin VB.TextBox txtInHandTo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10185
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3525
      Width           =   810
   End
   Begin VB.TextBox txtRemarks 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   420
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2445
      Width           =   10590
   End
   Begin VB.TextBox txtInHand 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10170
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1965
      Width           =   840
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10275
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   4350
      Width           =   780
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
         Caption         =   "Semi Finish Change Locations"
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
         Left            =   3090
         TabIndex        =   1
         Top             =   135
         Width           =   4650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Semi Finish Change Locations"
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
         Left            =   3105
         TabIndex        =   2
         Top             =   150
         Width           =   4650
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Move To :"
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
      Index           =   1
      Left            =   360
      TabIndex        =   11
      Top             =   2880
      Width           =   1035
   End
   Begin VB.Shape Shape2 
      Height          =   795
      Left            =   375
      Top             =   3150
      Width           =   10695
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Move From :"
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
      Index           =   0
      Left            =   360
      TabIndex        =   23
      Top             =   1410
      Width           =   1350
   End
   Begin VB.Shape Shape1 
      Height          =   1110
      Left            =   375
      Top             =   1680
      Width           =   10695
   End
   Begin MSForms.Label Label4 
      Height          =   240
      Index           =   1
      Left            =   420
      TabIndex        =   22
      Top             =   3300
      Width           =   10575
      BackColor       =   11517387
      Caption         =   $"frmMaterialMovementSF.frx":0000
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
   Begin MSForms.Label Label4 
      Height          =   240
      Index           =   2
      Left            =   420
      TabIndex        =   18
      Top             =   2235
      Width           =   10575
      BackColor       =   11517387
      Caption         =   " Remarks"
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
   Begin MSForms.Label Label4 
      Height          =   240
      Index           =   0
      Left            =   420
      TabIndex        =   13
      Top             =   1740
      Width           =   10575
      BackColor       =   11517387
      Caption         =   $"frmMaterialMovementSF.frx":00D9
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
      TabIndex        =   10
      Top             =   4395
      Width           =   1125
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   420
      TabIndex        =   7
      Top             =   645
      Width           =   10575
      BackColor       =   11517387
      Caption         =   $"frmMaterialMovementSF.frx":01B2
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
Attribute VB_Name = "frmMaterialMovementSF"
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


Private Sub cmbFromLocation_matched()

    txtInHand = GetSingleLongValue("Qty-IssdQty", "VendRcvdDetail INNER JOIN VStockOrderOpening ON VendRcvdDetail.Opening_RefID=VStockOrderOpening.EntryID" _
         , " WHERE VStockOrderOpening.EntryID=" & cmbFromLocation.ID)
    
    txtQty = Val(txtInHand)
    
    txtRemarks = GetSingleStringValue("Remarks", "VendRcvdDetail INNER JOIN VStockOrderOpening ON VendRcvdDetail.Opening_RefID=VStockOrderOpening.EntryID" _
        , " WHERE VStockOrderOpening.EntryID=" & cmbFromLocation.ID)
End Sub

Private Sub cmbFromStore_matched()

    cmbFromLocation.ClearVals
    cmbFromLocation.AddVals con, "RackNo + ' {' + ShelfNo + '} ' + '{' + CAST((Qty-IssdQty) AS VARCHAR(50)) + '}' + '(' + LotNo_Manual + ')'", "VendRcvdDetail INNER JOIN VStockOrderOpening ON VendRcvdDetail.Opening_RefID=VStockOrderOpening.EntryID" _
         , "VStockOrderOpening.EntryID", " WHERE ItemID='" & cmbMaterial.ID & "' AND Qty-IssdQty>0 AND VStockOrderOpening.ProcessID=" & cmbProcess.ID & " AND Store_RefID=" & cmbFromStore.ID & ""
         
End Sub

Private Sub cmbMaterial_matched()

    cmbProcess.ClearVals
    cmbProcess.AddVals con, "Description", "VItemProcesses", "ProcessID", " WHERE IsExist='" & cmbMaterial.ID & "'"
    
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

Private Sub cmbProcess_matched()

    cmbFromStore.ClearVals
    cmbFromStore.AddVals con, "DISTINCT StoreName", "VendRcvdDetail INNER JOIN VStockOrderOpening ON VendRcvdDetail.Opening_RefID=VStockOrderOpening.EntryID" _
             , "Store_RefID", "WHERE ItemCode='" & cmbMaterial.ID & "' AND (Qty-ISNULL(IssdQty,0))>0 AND VStockOrderOpening.ProcessID=" & cmbProcess.ID


    
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

Private Sub cmbToLocation_matched()
    txtInHandTo = GetSingleLongValue("SUM(Qty-IssdQty)", "VendRcvdDetail INNER JOIN VStockOrderOpening ON VendRcvdDetail.Opening_RefID=VStockOrderOpening.EntryID" _
         , " WHERE VStockOrderOpening.Shelf_RefID=" & cmbToLocation.ID & " AND ItemID='" & cmbMaterial.ID & "' AND VStockOrderOpening.ProcessID=" & cmbProcess.ID)
End Sub

Private Sub cmbToStore_matched()
    cmbToLocation.ClearVals
    cmbToLocation.AddVals con, "RackNo + ' {' + ShelfNo + '}'", "VStoreShelfs", "EntryID", " WHERE Store_RefID=" & cmbToStore.ID & " ORDER BY StoreName,RackNo,ShelfNo"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    If Saved Then
       cmbFromLocation.ID = ""
       cmbFromStore.ID = ""
       cmbToStore.ID = ""
       cmbToLocation.ID = ""
       
       txtInHand = ""
       txtInHandTo = ""
       txtQty = ""
    End If
    
End Sub

Private Sub cmdSaveClose_Click()

    If Saved Then Unload Me
    
End Sub


Public Function Saved() As Boolean
    On Error GoTo err
    
    If cmbMaterial.MatchFound = False Then
        MsgBox "Invalid Item.", vbInformation
        Exit Function
    ElseIf cmbFromLocation.MatchFound = False Then
        MsgBox "Invalid From Location.", vbInformation
        Exit Function
    ElseIf cmbToLocation.MatchFound = False Then
        MsgBox "Invalid To Location.", vbInformation
        Exit Function
    ElseIf Val(txtQty) > Val(txtInHand) Or Val(txtQty) <= 0 Then
        MsgBox "Invalid Qty.", vbInformation
        Exit Function
    End If
    
    Dim lFromEntryID As Long, lToEntryID As Long
    Call StartTrans(con)
    
'    con.Execute "INSERT INTO StockOrderOpening(ItemID,ProcessID,Qty,Location,UserName,MachineName,Shelf_RefID,Remarks) " & _
'     "VALUES('" & cmbMaterial.ID & "'," & cmbProcess.ID & "," & Val(txtQty) & ",'" & cmbToLocation.Text & "','" & _
'     curruentusername & "','" & strComputerName & "'," & cmbToLocation.ID & ",'" & txtRemarks & "')"
'
'    lToEntryID = GetSingleLongValue("MAX(EntryID)", "StockOrderOpening", " WHERE MachineName='" & strComputerName & "'")
'
'    lFromEntryID = cmbFromLocation.ID
'
'    con.Execute "INSERT INTO StockOrderOpening_LocationChanges(SOO_From_RefID,SOO_To_RefID,Qty) VALUES(" & lFromEntryID & _
'     "," & lToEntryID & "," & Val(txtQty) & ")"
'
'    con.Execute "UPDATE StockOrderOpening SET Qty=Qty-" & Val(txtQty) & " WHERE EntryID=" & cmbFromLocation.ID
    
    con.Execute "UPDATE StockOrderOpening SET Location='" & cmbToLocation.Text & "',Shelf_RefID=" & cmbToLocation.ID & " WHERE EntryID=" & _
     cmbFromLocation.ID
     
    con.CommitTrans
    
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
    
    cmbMaterial.DropDownOffLine = True
    cmbMaterial.AddVals con, "'{' + ItemID + '} ' + ItemName", "Items", "ItemID"
    
    
    cmbProcess.ListHeight = 3000
    cmbFromStore.ListHeight = 3000
    cmbFromLocation.ListHeight = 3000
    
    
    cmbToStore.ListHeight = 2400
    cmbToLocation.ListHeight = 2400
    Dim strCondition As String
    If CurrentUserName <> "Administrator" Then
        strCondition = " WHERE EntryID IN(SELECT Store_RefID FROM Users_Stores WHERE UserID=" & UserID & ")"
    End If
    cmbToStore.AddVals con, "StoreName", "Stores", "EntryID", strCondition & " ORDER BY StoreName"
    
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
