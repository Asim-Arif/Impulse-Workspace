VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmNewIPO 
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11700
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11700
   WindowState     =   2  'Maximized
   Begin VB.Frame FAll 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8520
      Left            =   45
      TabIndex        =   11
      Top             =   30
      Width           =   11640
      Begin VB.TextBox txtDept 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   9165
         TabIndex        =   5
         Top             =   990
         Width           =   1860
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   285
         TabIndex        =   0
         Top             =   990
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
      Begin VB.TextBox txtEdit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5385
         TabIndex        =   23
         Top             =   3030
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.PictureBox PicProduct 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2865
         Left            =   285
         ScaleHeight     =   2835
         ScaleWidth      =   2760
         TabIndex        =   21
         Top             =   2715
         Visible         =   0   'False
         Width           =   2790
         Begin VB.Image ImgProduct 
            Height          =   2835
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2760
         End
      End
      Begin VB.CheckBox chkShowPicture 
         Caption         =   "Show Picture"
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
         Height          =   255
         Left            =   270
         TabIndex        =   20
         Top             =   8100
         Width           =   1560
      End
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
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   11640
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "New IPO"
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
            TabIndex        =   15
            Top             =   165
            Width           =   11505
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "New IPO"
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
            TabIndex        =   16
            Top             =   180
            Width           =   11490
         End
      End
      Begin VB.CommandButton cmdDelete 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11055
         Picture         =   "frmNewIPO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Remove This Entry From Voucher"
         Top             =   2715
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   540
      End
      Begin MSComctlLib.ListView LV 
         Height          =   6120
         Left            =   270
         TabIndex        =   6
         Top             =   1350
         Width           =   10740
         _ExtentX        =   18944
         _ExtentY        =   10795
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         TextBackground  =   -1  'True
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
            Text            =   "S No."
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   4233
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Ord. Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Stock Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Alr. Res."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Reserved"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Balance"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Prod. Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Special Instructions"
            Object.Width           =   4233
         EndProperty
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
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
         Left            =   3555
         TabIndex        =   12
         Top             =   7785
         Width           =   7470
         Begin MSForms.CommandButton cmdPrint 
            Height          =   360
            Left            =   105
            TabIndex        =   24
            Top             =   195
            Width           =   2085
            Caption         =   "Print Contract Review"
            PicturePosition =   327683
            Size            =   "3678;635"
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdCancel 
            Height          =   360
            Left            =   5655
            TabIndex        =   10
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
         Begin MSForms.CommandButton cmdSClear 
            Height          =   360
            Left            =   2235
            TabIndex        =   8
            Top             =   195
            Width           =   1665
            Caption         =   "Save & New   "
            PicturePosition =   327683
            Size            =   "2937;635"
            Accelerator     =   78
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdSClose 
            Height          =   360
            Left            =   3945
            TabIndex        =   9
            Top             =   195
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
      Begin MSComCtl2.DTPicker DTIPO 
         Height          =   300
         Left            =   6075
         TabIndex        =   3
         Top             =   990
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   529
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
         Format          =   100794371
         CurrentDate     =   37250
      End
      Begin MSComCtl2.DTPicker DTTarget 
         Height          =   300
         Left            =   7620
         TabIndex        =   4
         Top             =   990
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   529
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
         Format          =   97386499
         CurrentDate     =   37250
      End
      Begin MSForms.ComboBox cmbOrderNo 
         Height          =   285
         Left            =   4140
         TabIndex        =   2
         Top             =   990
         Width           =   1935
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3413;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   2220
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   990
         Width           =   1935
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3413;503"
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
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmNewIPO.frx":0532
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
         Index           =   7
         Left            =   285
         TabIndex        =   22
         Top             =   750
         Width           =   10740
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Amount :"
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
         Index           =   1
         Left            =   8655
         TabIndex        =   19
         Top             =   7440
         Width           =   885
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Amount (In Words) :"
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
         Left            =   285
         TabIndex        =   18
         Top             =   7440
         Width           =   1830
      End
      Begin VB.Label lblTotWords 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2100
         TabIndex        =   17
         Top             =   7440
         Width           =   6570
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   9525
         TabIndex        =   13
         Top             =   7440
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frmNewIPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lIPONo As Long
Dim strDelList As String
Dim iColNo As Integer

Public Sub EditIPO(p_lIPONo As Long)
    
    On Error GoTo err
    
    lIPONo = p_lIPONo
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "SELECT * FROM VInternalProductionOrder WHERE IPONo=" & lIPONo, con, adOpenForwardOnly, adLockReadOnly
        cmbCust.ID = !CustCode & ""
        cmbCountry.Text = !Country & ""
        
        cmbOrderNo.Clear
        cmbOrderNo.AddItem !CustomerPO
        cmbOrderNo.ListIndex = 0
         
        DTIPO = !IPODT
        DTTarget = !TargetDT
        txtDept = !Department & ""
        
        .Close
        
        .Open "SELECT * FROM VInternalProductionOrderDetail1 WHERE IPONo=" & lIPONo, con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
        
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            ITM.ListSubItems.add(, , !CompItemCode & "").Tag = ""
            ITM.ListSubItems.add , , !ItemName & ""
            ITM.ListSubItems.add , , Val(!Qty & "")
            ITM.ListSubItems.add , , Val(!InHand & "")
            lAlreadyReservedQty = GetSingleLongValue("TotalReservedQty", "VIPOTotalReserved", " WHERE CompItemCode='" & !CompItemCode & "'")
            lAlreadyReservedQty = lAlreadyReservedQty - Val(!ReservedQty & "")
            ITM.ListSubItems.add , , lAlreadyReservedQty
            ITM.ListSubItems.add , , Val(!ReservedQty & "")
            lBalanceQty = Val(!Qty & "") - Val(!InHand & "") + lAlreadyReservedQty
            If lBalanceQty < 0 Then lBalanceQty = 0
            ITM.ListSubItems.add , , lBalanceQty
            'ITM.ListSubItems.add , , lBalanceQty
            ITM.ListSubItems.add , , Val(!ProductionQty & "")
            ITM.ListSubItems.add(, , !SpecialInstructions).Tag = Val(!DetailEntryID & "")
            
            .MoveNEXT
        Loop
        .Close
        
    End With
     
    Set rs = Nothing
    
    Me.Show
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub chkCust_Click()

    If cmbCatID.MatchFound Then
        Call FillItems
        For i = 1 To LV.ListItems.count
            If chkCust Then
                LV.ListItems(i).SubItems(2) = LV.ListItems(i).key
                LV.ListItems(i).ListSubItems(3) = LV.ListItems(i).ListSubItems(3).key
            Else
                LV.ListItems(i).SubItems(2) = LV.ListItems(i).Tag
                LV.ListItems(i).ListSubItems(3) = LV.ListItems(i).ListSubItems(3).Tag
            End If
        Next i
    End If
    
End Sub

Private Sub chkShowPicture_Click()
    If chkShowPicture.Value = vbChecked Then
        Call LV_ItemClick(LV.SelectedItem)
    Else
        PicProduct.Visible = False
    End If
End Sub

Private Sub cmbCatID_Change()
    Call cmbCatID_Click
End Sub

Private Sub cmbCatID_Click()
    If cmbCatID.MatchFound = False Then Exit Sub
    Call FillItems
End Sub

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    
    On Error GoTo err
    If cmbCountry.MatchFound = False Then Exit Sub

    Call AddToCombo(cmbOrderNo, "OrderNo", "FCustomerOrders", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "' AND OrderNo NOT IN(SELECT CustomerPO FROM InternalProductionOrder)")
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbCust_matched()
    
    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Country,Curr From ForeignCustomers Where CustCode='" & cmbCust.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            cmbCountry.List(cmbCountry.ListCount - 1, 1) = ![Curr] & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    If cmbCountry.ListCount > 0 Then cmbCountry.ListIndex = 0
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbCust_UnMatched()
    'cmbCountry.ClearVals
End Sub

Private Sub cmbCustItemID_Change()
    Call cmbCustItemID_Click
End Sub

Private Sub cmbCustItemID_Click()
    cmbItems.ListIndex = cmbCustItemID.ListIndex
End Sub

Private Sub cmbCustItemID_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode.Value = 13 Then
        txtQty.SetFocus
    End If
End Sub

Private Sub cmbDesc_Change()
    Call cmbDesc_Click
End Sub

Private Sub cmbDesc_Click()
    If cmbDesc.MatchFound Then
        cmbItems.ListIndex = cmbDesc.ListIndex
    End If
End Sub

Private Sub cmbDesc_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode.Value = 13 Then
        txtQty.SetFocus
    End If
End Sub

Private Sub cmbItems_Change()
    Call cmbItems_Click
    ''''''''Show Picture
    If chkShowPicture.Value = vbChecked Then
        Dim rs As New ADODB.Recordset
        With rs
            .Open "SELECT ItemPic FROM Items WHERE ItemID='" & cmbItems.Text & "'", con, adOpenForwardOnly, adLockReadOnly
            If Not IsNull(![ItemPic]) Then
                Set ImgProduct.DataSource = rs
                ImgProduct.DataField = "ItemPic"
            Else
                ImgProduct.Picture = LoadPicture("")
            End If
            .Close
        End With
        Set rs = Nothing
        PicProduct.Visible = True
    End If
End Sub

Private Sub cmbItems_Click()

    If cmbItems.MatchFound Then
        
        cmbCustItemID.ListIndex = cmbItems.ListIndex
        cmbDesc.ListIndex = cmbItems.ListIndex
        txtUnit = cmbItems.List(cmbItems.ListIndex, 1)
        If cmbQuality.ListIndex <= 0 Then
            txtPrice = Val(cmbItems.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2) & "")
        Else    'Top Quality
            txtPrice = Val(cmbDesc.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2) & "") 'Val(cmbDesc.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2) & "")
        End If
        
    Else
        txtUnit = ""
        txtPrice = ""
        txtPrice.Tag = ""
    End If
    
End Sub

Private Sub cmbItems_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode.Value = 13 Then
        txtQty.SetFocus
    End If
End Sub

Private Sub cmbQuality_Change()
    Call cmbQuality_Click
End Sub

Private Sub cmbQuality_Click()
    If cmbItems.MatchFound = False Then Exit Sub
    If cmbQuality.ListIndex <= 0 Then
        txtPrice = Val(cmbItems.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2) & "")
    Else    'Top Quality
        txtPrice = Val(cmbDesc.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2) & "") 'Val(cmbDesc.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2) & "")
    End If
End Sub

Private Sub cmbTradeTerms_Change()
    Call cmbTradeTerms_Click
End Sub

Private Sub cmbTradeTerms_Click()

    If cmbTradeTerms.MatchFound Then
        cmbCatID.Locked = False
        cmbItems.Locked = False
        cmbCustItemID.Locked = False
        cmbDesc.Locked = False
        If cmbItems.MatchFound Then
            txtPrice = cmbItems.List(cmbItems.ListIndex, cmbTradeTerms.ListIndex + 2)
        End If
    Else
        cmbCatID.Locked = True
        cmbItems.Locked = True
        cmbCustItemID.Locked = True
        cmbDesc.Locked = True
    End If
    
End Sub

Private Sub cmbOrderNo_Change()
    Call cmbOrderNo_Click
End Sub

Private Sub cmbOrderNo_Click()

    On Error GoTo err
    If lIPONo <> 0 Then Exit Sub
    LV.ListItems.Clear
        
    If cmbOrderNo.MatchFound = False Then Exit Sub
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim lAlreadyReservedQty As Long
    Dim lBalanceQty As Long
    With rs
        .Open "SELECT * FROM VrptOrders WHERE OrderNo='" & cmbOrderNo.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            ITM.ListSubItems.add(, , !CompITemID & "").Tag = Val(!MaxLotSize & "")
            ITM.ListSubItems.add , , !ItemName & ""
            ITM.ListSubItems.add , , Val(!Qty & "")
            ITM.ListSubItems.add , , Val(!InHand & "")
            lAlreadyReservedQty = GetSingleLongValue("TotalReservedQty", "VIPOTotalReserved", " WHERE CompItemCode='" & !CompITemID & "'")
            ITM.ListSubItems.add , , lAlreadyReservedQty
            ITM.ListSubItems.add , , "0"
            lBalanceQty = Val(!Qty & "") - Val(!InHand & "") + lAlreadyReservedQty
            If lBalanceQty < 0 Then lBalanceQty = 0
            ITM.ListSubItems.add , , lBalanceQty
            ITM.ListSubItems.add , , lBalanceQty
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ContractReview.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{FCustomerOrders.OrderNo}='" & cmbOrderNo.Text & "'", rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    'If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub cmdAddItem_Click()
    
    On Error GoTo err
    
    If Not cmbItems.MatchFound Or cmbItems = "" Then
        MsgBox "Invalid Item Selected", vbInformation
        Exit Sub
    'ElseIf Val(txtPrice) = 0 Then
    '    MsgBox "Price Cannot Be Zero", vbInformation
    '    Exit Sub
    ElseIf Val(txtQty) = 0 Then
        MsgBox "Quantity To Order Cannot Be Zero", vbInformation
        Exit Sub
    End If

    Dim ITM As ListItem
    Dim SNo As Integer
    If LV.ListItems.count > 0 Then
        SNo = Val(LV.ListItems(LV.ListItems.count).Text) + 1
    Else
        SNo = 1
    End If
    
    Set ITM = LV.ListItems.add(, cmbCustItemID.Text & "'", SNo)
    ITM.Tag = cmbItems.Text 'cmbItems.List(cmbItems.ListIndex, 2)
    
    ITM.ListSubItems.add , , cmbCustItemID.Text
    ITM.ListSubItems.add , , cmbItems
    ITM.ListSubItems.add(, , cmbDesc).Tag = cmbCustItemID.Text
    
    ITM.ListSubItems.add , , txtUnit
    
    ITM.ListSubItems.add(, , Format(txtPrice, "#0.0000")).Tag = Val(txtPrice.Tag)
    ITM.ListSubItems.add(, , txtQty).Tag = "New"
    ITM.ListSubItems.add , , lblAmt
    ITM.ListSubItems.add , , SNo
   
    'cmbItemID.RemoveItem cmbItemID.ListIndex
    'cmbItem.RemoveItem cmbItem.ListIndex
    
    Call GetTotal

    txtQty = 0
    cmbCustItemID.SetFocus
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    
    If LV.SelectedItem.ListSubItems(6).Tag <> "New" Then
        'If Val(LV.SelectedItem.ListSubItems(6).Tag) <> Val(LV.SelectedItem.ListSubItems(6)) Then
        If Val(LV.SelectedItem.ListSubItems(6).Tag) <> 0 Then
            MsgBox "Can't Remove This Item.This Item Is In Use In Some Proforma.", vbInformation
            Exit Sub
        End If
    End If
    
    If Val(LV.SelectedItem.ListSubItems(7).Tag) <> 0 Then
        strDelList = strDelList & Val(LV.SelectedItem.ListSubItems(7).Tag) & ","
    End If
    
    LV.ListItems.Remove LV.SelectedItem.Index
    
End Sub

Private Sub cmdSClear_Click()
    If Saved Then
        Unload Me
        Me.Show
    End If
End Sub

Private Sub cmdSClose_Click()
    If Saved Then
        Unload Me
    End If
End Sub


Private Sub Form_Load()
   
    On Error GoTo err
    lIPONo = 0

    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " WHERE Active=1")
    DTIPO = Date 'ServerDate
    DTTarget = DateAdd("m", 2, DTIPO)
    Me.KeyPreview = True
    cmbCust.ListHeight = 2500
    
    Exit Sub
err:
    MsgBox err.Description
   
End Sub

Public Sub AddNew()
    txtInternalRefNo = GetInternalRefNo
    Me.Show
End Sub

Private Function GetInternalRefNo() As Long
 
    On Error GoTo err
    
    Dim lRefNo As Long
    lRefNo = GetSingleLongValue("MAX(InternalRefNo)", "FCustomerOrders")
     
    If lRefNo = 0 Then
        lRefNo = Val(InputBox("Please Enter the Starting Number for Internal Ref. #", "Internal Reference No.", "596"))
    Else
        lRefNo = lRefNo + 1
    End If
    
    GetInternalRefNo = lRefNo
    
    Exit Function
err:
    MsgBox err.Description

    
End Function

Private Sub Form_Resize()
On Error Resume Next
   FAll.Move ((Me.ScaleWidth - FAll.Width) / 2), ((Me.ScaleHeight - FAll.Height) / 2)
End Sub

Private Sub lblPymtTerms_Click()
    Dim f As New frmMiscS
    Load f
    Call f.ShowForm("PaymentTerms", "Payment Terms", "Payment Terms")
End Sub

Private Sub LV_Click()
    If LV.ListItems.count > 0 Then
        Call LV_ItemClick(LV.SelectedItem)
    End If
End Sub

Private Sub LV_DblClick()
    If Not (iColNo = 10 Or iColNo = 9 Or iColNo = 7) Then Exit Sub
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(iColNo).Width
        .Text = LV.SelectedItem.SubItems(iColNo - 1)
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
End Sub

Private Function UsedInPInvoice() As String

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim lRefID As Long
    lRefID = Val(LV.SelectedItem.ListSubItems(7).Tag)
    With rs
        .Open "Select Top 1 CustomInvoice From CustomInvoiceItems Where RefID In(Select EntryID From FProformaOrders Where OrderEntryID=" & lRefID & ")", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            UsedInPInvoice = ""
        Else
            UsedInPInvoice = .Fields(0)
        End If
        .Close
    End With
    
    Exit Function
    
err:
    MsgBox err.Description
    
End Function

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)

    With cmdDelete
        .Move .Left, LV.Top + Item.Top
        .Visible = True
    End With
    
    If chkShowPicture.Value = vbChecked Then
        Dim rs As New ADODB.Recordset
        With rs
            .Open "SELECT ItemPic FROM Items WHERE ItemID='" & Item.SubItems(1) & "'", con, adOpenForwardOnly, adLockReadOnly
            If Not IsNull(![ItemPic]) Then
                Set ImgProduct.DataSource = rs
                ImgProduct.DataField = "ItemPic"
            Else
                ImgProduct.Picture = LoadPicture("")
            End If
            .Close
        End With
        Set rs = Nothing
        Dim lTop As Long
        lTop = LV.Top + LV.SelectedItem.Top
        If (lTop + PicProduct.Height) > LV.Top + LV.Height Then
            lTop = LV.Top + LV.Height - PicProduct.Height
        End If
        PicProduct.Move LV.Left + LV.Width - PicProduct.Width, lTop
        PicProduct.Visible = True
    Else
'        If PicProduct.Visible = True Then PicProduct.Visible = False
    End If
    
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If x > LV.ColumnHeaders(i).Left And x < LV.ColumnHeaders(i).Left + LV.ColumnHeaders(i).Width Then
            iColNo = i
            Exit For
        End If
    Next
    
End Sub

Private Sub OrderDT_Change()
    DeliveryDT.Value = DateAdd("m", 2, OrderDT.Value)
End Sub

Private Sub txtEditPrice_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        txtEditPrice.Visible = False
    ElseIf KeyAscii = 13 Then
        LV.SelectedItem.SubItems(5) = Val(txtEditPrice)
        LV.SelectedItem.SubItems(7) = Val(txtEditPrice) * Val(LV.SelectedItem.SubItems(6))
        txtEditPrice.Visible = False
    End If

End Sub

Private Sub txtEditQty_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        txtEditQty.Visible = False
    ElseIf KeyAscii = 13 Then
        If iColNo = 9 Then
            If Val(txtEditQty) <= 0 Then
                MsgBox "Invalid Sort No.", vbInformation
                txtEditQty.SetFocus
                Exit Sub
            End If
            LV.SelectedItem.SubItems(8) = Val(txtEditQty)
            If LV.SelectedItem.Index < LV.ListItems.count Then
                Dim iIndex As Integer
                iIndex = LV.SelectedItem.Index + 1
                LV.ListItems(iIndex).Selected = True
                With txtEditQty
                    .Top = LV.SelectedItem.Top + LV.Top
                    .Text = LV.SelectedItem.SubItems(8)
                    .SelStart = 0
                    .SelLength = Len(.Text)
                    .SetFocus
                End With
            Else
                txtEditQty.Visible = False
            End If
            'txtEditQty.Visible = False
        Else
            If Val(txtEditQty) <= 0 Or Val(txtEditQty) < Val(LV.SelectedItem.ListSubItems(6).Tag) Then
                MsgBox "Invalid Qty.", vbInformation
                txtEditQty.SetFocus
                Exit Sub
            End If
            LV.SelectedItem.SubItems(6) = Val(txtEditQty)
            LV.SelectedItem.SubItems(7) = Val(txtEditQty) * Val(LV.SelectedItem.SubItems(5))
            txtEditQty.Visible = False
        End If
    End If
    
End Sub

Private Sub txtPrice_Change()
   lblAmt = Format(Val(txtQty) * Val(txtPrice), "#0.00")
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
   KeyAscii = OnlyNumber(KeyAscii, True)
End Sub

Private Sub txtQty_Change()
   lblAmt = Format(Val(txtQty) * Val(txtPrice), "#0.00")
End Sub

Private Sub txtQty_GotFocus()
    cmdAddItem.Default = True
    With txtQty
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
Private Sub txtQty_LostFocus()
    cmdAddItem.Default = False
End Sub
Private Sub txtQty_KeyPress(KeyAscii As Integer)
   KeyAscii = OnlyNumber(KeyAscii)
End Sub


Private Sub GetTotal()

    Dim i As Integer
    Dim total As Double
    total = 0

    For i = 1 To LV.ListItems.count
        total = total + Val(LV.ListItems(i).ListSubItems(7).Text)
    Next i

    lblTotal = Format(total, "#0.00")
    lblTotWords = ConvertInWords(lblTotal, txtCurrency.Text)
    
End Sub



Private Function Saved() As Boolean
            
    On Error GoTo err
        
    Saved = False
    
    If LV.ListItems.count = 0 Then
        MsgBox "No Items Added to This Order", vbInformation
        Exit Function
    End If
    
    Call StartTrans(con)
    '----------------------------------------------------------------
    'If strOrderNo <> "" Then
    '    con.Execute "Delete from FCustomerOrders Where OrderNo='" & _
         strOrderNo & "'"
    '     'Now Deleting is giving a problem...
    'End If
    
    'Save Order
    Dim lEntryID As Long
    
    If lIPONo <> 0 Then
            con.Execute "UPDATE InternalProductionOrder SET IPODT='" & DTIPO & "',TargetDT='" & DTTarget & _
             "',Department='" & txtDept & "' WHERE IPONo=" & lIPONo
             
'             If strDelList <> "" Then
'                strDelList = Left(strDelList, Len(strDelList) - 1)
'
'                con.Execute "DELETE FROM FOrderItems Where ID In(" & strDelList & ")"
'            End If
        lEntryID = GetSingleLongValue("EntryID", "InternalProductionOrder", " WHERE IPONo=" & lIPONo)
    Else
        Dim lNewIPONo As Long
        lNewIPONo = GetSingleLongValue("MAX(IPONo)", "InternalProductionOrder") + 1
        con.Execute "INSERT INTO InternalProductionOrder(IssuedTo,IPONo,IPODT,TargetDT,CustomerPO,Department) " & _
         "VALUES('MSD'," & lNewIPONo & ",'" & DTIPO & "','" & DTTarget & "','" & cmbOrderNo.Text & "','" & txtDept & "')"
            
        lEntryID = GetSingleLongValue("MAX(EntryID)", "InternalProductionOrder")
    End If
    
    'Now Save Items
    Dim strItemCode As String, lIPOD_EntryID As Long
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If lIPONo = 0 Then
                
                con.Execute "INSERT INTO InternalProductionOrderDetail(RefID,CompItemCode,Qty,ReservedQty,SpecialInstructions,ProductionQty" & _
                 ") VALUES(" & lEntryID & ",'" & .SubItems(1) & "'," & Val(.SubItems(3)) & "," & Val(.SubItems(6)) & ",'" & _
                 .SubItems(9) & "'," & Val(.SubItems(8)) & ")"
                
                lIPOD_EntryID = GetSingleLongValue("MAX(EntryID)", "InternalProductionOrderDetail", " WHERE RefID=" & lEntryID)
            Else
                con.Execute "UPDATE InternalProductionOrderDetail SET ReservedQty=" & Val(.SubItems(6)) & _
                 ",SpecialInstructions='" & .SubItems(9) & "',ProductionQty=" & Val(.SubItems(8)) & " WHERE EntryID=" & Val(.ListSubItems(9).Tag)
            End If
                'Commented the code below; don't generate lots at saving...
                'If Val(.SubItems(8)) > 0 Then   'If Production Qty >0, Generate Lots.
                '    Call SaveLots(lIPOD_EntryID, Val(.ListSubItems(1).Tag), Val(.SubItems(8)))
                'End If
            'Else
                'con.Execute "UPDATE FOrderItems SET Qty=" & Val(LV.ListItems(i).SubItems(6)) & _
                 ",InvQty=" & Val(LV.ListItems(i).SubItems(6)) - Val(LV.ListItems(i).ListSubItems(6).Tag) & _
                 ",Price=" & Val(LV.ListItems(i).SubItems(5)) & ",CustomPrice=" & Val(LV.ListItems(i).SubItems(5)) & _
                 ",SortNo=" & Val(LV.ListItems(i).SubItems(8)) & " WHERE ID=" & Val(LV.ListItems(i).ListSubItems(7).Tag)
                       
            'End If
            End With
    Next
      
    con.CommitTrans
    
    Saved = True
'    If chkPrint.Value = True Then
'        Call PrintOrder(txtOrderNo)
'    End If
    
    Exit Function

err:
    MsgBox err.Description
    
End Function

Private Sub SaveLots(lRefID As Long, lMaxLotQty As Long, lQty As Long)

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
    
End Sub

Private Sub PrintOrder(strOrderNo As String)

    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptFCustOrders.rpt") '
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{CustOrders.OrderNo}='" & strOrderNo & "' ", rpt
    
    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub FillItems()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    With rs
        
        If cmbCatID.ListIndex = 0 Then
            If strCompany = "All Seas" Then
                .Open "SELECT ItemID,ItemID AS CompItemID,ItemName AS Description,0 AS PackingMode,1 AS FOB,1 AS ExWorks" & _
                 ",1 AS CIFAir,1 AS CIFSea,1 AS CnFAir,1 AS CnFSea,1 AS CnIAir,1 AS CnISea,1 AS CustomFOB,1 AS CustomExWorks,1 AS CustomCIFAir,1 AS CustomCIFSea,1 AS CustomCnFAir,1 AS CustomCnFSea,1 AS CustomCnIAir,1 AS CustomCnISea FROM Items", con, adOpenForwardOnly, adLockReadOnly
            Else
                .Open "SELECT * FROM VFCustomerCatalog WHERE CustCode+Country='" & cmbCust.Text & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
            End If
        Else
            If strCompany = "All Seas" Then
                .Open "SELECT ItemID,ItemID AS CompItemID,ItemName AS Description,Unit AS PackingMode,1 AS FOB,1 AS ExWorks" & _
                 ",1 AS CIFAir,1 AS CIFSea,1 AS CnFAir,1 AS CnFSea,1 AS CnIAir,1 AS CnISea,1 AS CustomFOB,1 AS CustomExWorks,1 AS CustomCIFAir,1 AS CustomCIFSea,1 AS CustomCnFAir,1 AS CustomCnFSea,1 AS CustomCnIAir,1 AS CustomCnISea FROM Items WHERE CatID=" & cmbCatID.List(cmbCatID.ListIndex, 1), con, adOpenForwardOnly, adLockReadOnly
            Else
                .Open "SELECT * FROM VFCustomerCatalog WHERE CatID=" & cmbCatID.List(cmbCatID.ListIndex, 1) & " AND CustCode+Country='" & cmbCust.Text & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
            End If
        End If
        cmbItems.Clear
        cmbDesc.Clear
        cmbCustItemID.Clear
        Do Until .EOF
            cmbCustItemID.AddItem !ITemID & ""
            If chkCust Then
                cmbItems.AddItem ![ITemID] & ""
                cmbDesc.AddItem ![Description] & ""
            Else
                cmbItems.AddItem ![CompITemID] & ""
                'cmbDesc.AddItem ![CompDesc] & ""
                cmbDesc.AddItem ![Description] & ""
            End If
            
            'cmbItems.List(cmbItems.ListCount - 1, 1) = ![Unit] & ""
            cmbItems.List(cmbItems.ListCount - 1, 1) = Val(![PackingMode] & "")
            
            cmbItems.List(cmbItems.ListCount - 1, 2) = Val(![FOB] & "")
            cmbItems.List(cmbItems.ListCount - 1, 3) = Val(![ExWorks] & "")
            
            cmbItems.List(cmbItems.ListCount - 1, 4) = Val(![CIFAir] & "")
            cmbItems.List(cmbItems.ListCount - 1, 5) = Val(![CIFSea] & "")
            
            cmbItems.List(cmbItems.ListCount - 1, 6) = Val(![CnFAir] & "")
            cmbItems.List(cmbItems.ListCount - 1, 7) = Val(![CnFSea] & "")
            
            cmbItems.List(cmbItems.ListCount - 1, 8) = Val(![CnIAir] & "")
            cmbItems.List(cmbItems.ListCount - 1, 9) = Val(![CnISea] & "")
            
            'cmbDesc.List(cmbDesc.ListCount - 1, 1) = ![Description] & ""
            'cmbDesc.List(cmbDesc.ListCount - 1, 2) = ![CompDesc] & ""
            
            cmbDesc.List(cmbDesc.ListCount - 1, 2) = Val(![FOBTop] & "")
            cmbDesc.List(cmbDesc.ListCount - 1, 3) = Val(![ExWorksTop] & "")
            cmbDesc.List(cmbDesc.ListCount - 1, 4) = Val(![CIFAirTop] & "")
            cmbDesc.List(cmbDesc.ListCount - 1, 5) = Val(![CIFSeaTop] & "")
            
            cmbDesc.List(cmbDesc.ListCount - 1, 6) = Val(![CnFAirTop] & "")
            cmbDesc.List(cmbDesc.ListCount - 1, 7) = Val(![CnFSeaTop] & "")
            
            cmbDesc.List(cmbDesc.ListCount - 1, 8) = Val(![CnIAirTop] & "")
            cmbDesc.List(cmbDesc.ListCount - 1, 9) = Val(![CnISeaTop] & "")
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If iColNo = 7 Then  'Reserve....
            If Val(txtEdit) > Val(LV.SelectedItem.SubItems(4)) - Val(LV.SelectedItem.SubItems(5)) Then
                MsgBox "Invalid Qty.", vbInformation
                Exit Sub
            End If
        End If
        LV.SelectedItem.SubItems(iColNo - 1) = txtEdit
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
    
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub
