VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmFNewStockOrder 
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10635
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
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   10635
   WindowState     =   2  'Maximized
   Begin VB.Frame FAll 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7155
      Left            =   75
      TabIndex        =   5
      Top             =   60
      Width           =   10560
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   285
         TabIndex        =   11
         Top             =   1050
         Width           =   3000
         _ExtentX        =   5292
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
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6465
         MaxLength       =   150
         TabIndex        =   21
         Top             =   1065
         Width           =   1650
      End
      Begin VB.TextBox txtCurrency 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5250
         MaxLength       =   150
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1065
         Width           =   1185
      End
      Begin VB.Frame FNew 
         BorderStyle     =   0  'None
         Height          =   570
         Left            =   285
         TabIndex        =   12
         Top             =   1365
         Width           =   10230
         Begin VB.TextBox txtQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   8625
            MaxLength       =   150
            TabIndex        =   20
            Top             =   255
            Width           =   930
         End
         Begin VB.TextBox txtUnit 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   7665
            MaxLength       =   150
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   255
            Width           =   930
         End
         Begin VB.CommandButton cmdAddItem 
            Caption         =   "&Add"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   9630
            Picture         =   "frmFNewStockOrder.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Add This Entry To Voucher"
            Top             =   15
            UseMaskColor    =   -1  'True
            Width           =   540
         End
         Begin MSForms.ComboBox cmbCatID 
            Height          =   285
            Left            =   0
            TabIndex        =   16
            Top             =   255
            Width           =   2385
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4207;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbDesc 
            Height          =   285
            Left            =   4260
            TabIndex        =   18
            Top             =   255
            Width           =   3375
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "5953;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbItems 
            Height          =   285
            Left            =   2415
            TabIndex        =   17
            Top             =   255
            Width           =   1815
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "3201;503"
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
            Caption         =   $"frmFNewStockOrder.frx":0532
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
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   9555
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   10560
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "New Stock Order"
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
            Left            =   105
            TabIndex        =   8
            Top             =   165
            Width           =   10605
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "New Stock Order"
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
            TabIndex        =   9
            Top             =   180
            Width           =   10590
         End
      End
      Begin VB.CommandButton cmdDelete 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9915
         Picture         =   "frmFNewStockOrder.frx":05DA
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Remove This Entry From Voucher"
         Top             =   1965
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   540
      End
      Begin MSComctlLib.ListView LV 
         Height          =   4470
         Left            =   285
         TabIndex        =   0
         Top             =   1965
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   7885
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S No."
            Object.Width           =   1014
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cat"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Description"
            Object.Width           =   5997
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Unit"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Qty"
            Object.Width           =   1499
         EndProperty
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   4665
         TabIndex        =   6
         Top             =   6405
         Width           =   5280
         Begin MSForms.CommandButton cmdCancel 
            Height          =   360
            Left            =   3510
            TabIndex        =   4
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
            Left            =   90
            TabIndex        =   2
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
            Left            =   1800
            TabIndex        =   3
            Top             =   195
            Width           =   1665
            Caption         =   "Save & Close  "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmFNewStockOrder.frx":0B0C
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
      Begin MSComCtl2.DTPicker OrderDT 
         Height          =   300
         Left            =   8145
         TabIndex        =   22
         Top             =   1050
         Width           =   1725
         _ExtentX        =   3043
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
         Format          =   60293123
         CurrentDate     =   37250
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   3315
         TabIndex        =   24
         Top             =   1065
         Width           =   1905
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3360;503"
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
         Caption         =   $"frmFNewStockOrder.frx":0F5A
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
         TabIndex        =   23
         Top             =   810
         Width           =   9555
      End
      Begin MSForms.CheckBox chkPrint 
         Height          =   345
         Left            =   270
         TabIndex        =   10
         Top             =   6450
         Width           =   1545
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   4
         Size            =   "2725;609"
         Value           =   "0"
         Caption         =   "Print On Save"
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
End
Attribute VB_Name = "frmFNewStockOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrderNo As String
Dim strDelList As String
Public Sub EditOrder(TempOrdNo As String)
    
    On Error GoTo err
    strOrderNo = TempOrdNo
    strDelList = ""
    Dim rs As New ADODB.Recordset
    Dim itm As ListItem
    
    With rs
        .Open "Select * From VrptOrders Where OrderNo='" & strOrderNo & "'", con, adOpenForwardOnly, adLockReadOnly
        
        cmbCust.Text = ![CustCode] & ""
        cmbCountry = ![Country] & ""
        
        txtOrderNo = ![OrderNo] & ""
        OrderDT = ![DT] & ""
        
        
        Do Until .EOF
        
            Set itm = LV.ListItems.Add(, ![ITemID] & "'", .AbsolutePosition)
            
            itm.Tag = ![CompItemID] & "" 'cmbItems.List(cmbItems.ListIndex, 2)
            
            itm.ListSubItems.Add , , Left(![ITemID] & "", 2) '![CatID] & ""
            itm.ListSubItems.Add , , ![ITemID] & ""
            itm.ListSubItems.Add(, ![ItemName] & "", ![ItemName] & "").Tag = ![Description] & ""
            
            itm.ListSubItems.Add(, , txtUnit).Tag = ![ID] & ""
            
            itm.ListSubItems.Add(, , Val(![Qty] & "")).Tag = Val(![InvQty] & "")
            
            .MoveNext
        Loop
        .Close
    End With
    
    txtQty = 0
    cmbItems.SetFocus
    
    Set rs = Nothing
    Me.Show
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub cmbCatID_Change()
    Call cmbCatID_Click
End Sub

Private Sub cmbCatID_Click()
    Call FillItems
End Sub

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    
    On Error GoTo err
    If cmbCountry.MatchFound = False Then Exit Sub
    Dim rs As New ADODB.Recordset
    txtCurrency = cmbCountry.List(cmbCountry.ListIndex, 1)
    With rs
        '.Open "Select Distinct CatID,CatDesc From VFCustomerCatalog Where CustCode+Country='" & cmbCust.Text & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        .Open "Select Distinct CatID,CatDesc From VFCustomerCatalog ", con, adOpenForwardOnly, adLockReadOnly
        cmbCatID.Clear
        Do Until .EOF
            cmbCatID.AddItem .Fields(1) & ""
            cmbCatID.List(cmbCatID.ListCount - 1, 1) = .Fields(0) & ""
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbCust_matched()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    If cmbCust.Text = "HICO" Then
    
        Dim OrdNo As Long
        With rs
            .Open "Select Max(Cast(SubString(OrderNo,12,Len(OrderNo)-11) As int)) From FCustomerOrders Where Left(OrderNo,11)='HICO Order-'", con, adOpenForwardOnly, adLockReadOnly
            If .RecordCount > 0 Then
                OrdNo = Val(.Fields(0) & "") + 1
            Else
                OrdNo = 1
            End If
            
            .Close
        End With
        Set rs = Nothing
        txtOrderNo = "HICO Order-" & OrdNo
        txtOrderNo.Locked = True
        
    Else
        txtOrderNo = ""
        txtOrderNo.Locked = False
    End If
    
    
    With rs
        .Open "Select Country,Curr From ForeignCustomers Where CustCode='" & cmbCust.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            cmbCountry.List(cmbCountry.ListCount - 1, 1) = ![Curr] & ""
            .MoveNext
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

Private Sub cmbDesc_Change()
    Call cmbDesc_Click
End Sub

Private Sub cmbDesc_Click()
    If cmbDesc.MatchFound Then
        cmbItems.ListIndex = cmbDesc.ListIndex
    End If
End Sub

Private Sub cmbItems_Change()
    Call cmbItems_Click
End Sub

Private Sub cmbItems_Click()

    If cmbItems.MatchFound Then
        cmbDesc.ListIndex = cmbItems.ListIndex
        txtUnit = cmbItems.List(cmbItems.ListIndex, 3)
        
    Else
        txtUnit = ""
    End If
    
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{Tab}"
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

    Dim itm As ListItem
    Dim Sno As Integer
    If LV.ListItems.Count > 0 Then
        Sno = Val(LV.ListItems(LV.ListItems.Count).Text) + 1
    Else
        Sno = 1
    End If
    
    Set itm = LV.ListItems.Add(, cmbItems.List(cmbItems.ListIndex, 1) & "'", Sno)
    itm.Tag = cmbItems.List(cmbItems.ListIndex, 2)
    
    itm.ListSubItems.Add , , cmbCatID
    itm.ListSubItems.Add , , cmbItems
    itm.ListSubItems.Add(, cmbDesc.List(cmbDesc.ListIndex, 1), cmbDesc).Tag = cmbDesc.List(cmbDesc.ListIndex, 2)
    
    itm.ListSubItems.Add , , txtUnit
    
    itm.ListSubItems.Add(, , txtQty).Tag = "New"
    
   
    'cmbItemID.RemoveItem cmbItemID.ListIndex
    'cmbItem.RemoveItem cmbItem.ListIndex
    
    txtQty = 0
    cmbItems.SetFocus
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdDelete_Click()

    If LV.ListItems.Count = 0 Then Exit Sub
    
    If Val(LV.SelectedItem.ListSubItems(5).Tag) <> Val(LV.SelectedItem.ListSubItems(5)) Then
        MsgBox "Can't Remove This Item.This Item Is In Use In Some Proforma.", vbInformation
        Exit Sub
    End If
    
    If Val(LV.SelectedItem.ListSubItems(4).Tag) <> 0 Then
        strDelList = strDelList & Val(LV.SelectedItem.ListSubItems(4).Tag) & ","
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
      
    strOrderNo = ""
    cmbCust.ListHeight = 2500
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where CustCode='HICO'")
    cmbCust.Text = "HICO"
    cmbCust.Enabled = False
    
    OrderDT = ServerDate
    DeliveryDT = DateAdd("m", 1, ServerDate)
    Me.KeyPreview = True
    
    Exit Sub
err:
    MsgBox err.Description
   
End Sub

Private Sub Form_Resize()
On Error Resume Next
   FAll.Move ((Me.ScaleWidth - FAll.Width) / 2), ((Me.ScaleHeight - FAll.Height) / 2)
End Sub



Private Sub LV_Click()
    If LV.ListItems.Count > 0 Then
        Call LV_ItemClick(LV.SelectedItem)
    End If
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With cmdDelete
        .Move .Left, LV.Top + Item.Top
        .Visible = True
    End With
End Sub


Private Sub txtQty_Change()
   lblAmt = Format(Val(txtQty) * Val(txtPrice), "#0.00")
End Sub

Private Sub txtQty_GotFocus()
cmdAddItem.Default = True
End Sub
Private Sub txtQty_LostFocus()
cmdAddItem.Default = False
End Sub
Private Sub txtQty_KeyPress(KeyAscii As Integer)
   KeyAscii = OnlyNumber(KeyAscii)
End Sub






Private Function Saved() As Boolean
    
    Saved = False
    If txtOrderNo = "" Then
        MsgBox "Please Enter Order No.", vbInformation
        txtOrderNo.SetFocus
        Exit Function
    End If
    
    If LV.ListItems.Count = 0 Then
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
    
    If strOrderNo <> "" Then
        con.Execute "Update FCustomerOrders Set OrderNo='" & txtOrderNo & "',DT='" & _
         OrderDT & "' Where OrderNo='" & strOrderNo & "'"
         
        If strDelList <> "" Then
            strDelList = Left(strDelList, Len(strDelList) - 1)
         
            con.Execute "Delete From FOrderItems Where ID In(" & strDelList & ")"
        End If
    Else
        con.Execute "Insert Into FCustomerOrders (CustCode,Country," & _
            "OrderNo,DT) values('" & _
            cmbCust.ID & "','" & cmbCountry & "','" & txtOrderNo & _
             "','" & OrderDT & "')"
    End If
    
    'Now Save Items
    Dim strItemCode As String
    For i = 1 To LV.ListItems.Count
        strItemCode = Left(LV.ListItems(i).key, Len(LV.ListItems(i).key) - 1)
        
SaveNewItem:
            If Val(LV.ListItems(i).ListSubItems(5).Tag) = 0 Then 'Means that it's a new Entry
                
                con.Execute "Insert into FOrderItems (OrderNo,ItemCode,CompItemCode," & _
                 "Price,Qty,InvQty,CustomPrice) Values ('" & txtOrderNo & "','" & _
                 Left(LV.ListItems(i).key, Len(LV.ListItems(i).key) - 1) & _
                 "','" & LV.ListItems(i).Tag & "',0," & _
                 Val(LV.ListItems(i).SubItems(5)) & "," & _
                 Val(LV.ListItems(i).SubItems(5)) & ",0)"
                 
            End If
    Next
      
    con.CommitTrans
    
    Saved = True
    Exit Function

err:
    MsgBox err.Description
    
End Function

Private Sub FillItems()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    With rs
        '.Open "Select * From VFCustomerCatalog Where CatID=" & cmbCatID.List(cmbCatID.ListIndex, 1) & " and CustCode+Country='" & cmbCust.Text & cmbCountry.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        .Open "Select * From VFCustomerCatalog Where CatID=" & cmbCatID.List(cmbCatID.ListIndex, 1), con, adOpenForwardOnly, adLockReadOnly
        cmbItems.Clear
        cmbDesc.Clear
        
        Do Until .EOF
        
            If chkCust Then
                cmbItems.AddItem ![ITemID] & ""
                cmbDesc.AddItem ![Description] & ""
            Else
                cmbItems.AddItem ![CompItemID] & ""
                cmbDesc.AddItem ![CompDesc] & ""
            End If
            
            cmbItems.List(cmbItems.ListCount - 1, 1) = ![ITemID] & ""
            cmbItems.List(cmbItems.ListCount - 1, 2) = ![CompItemID] & ""
            'cmbItems.List(cmbItems.ListCount - 1, 2) = ![Description] & ""
            cmbItems.List(cmbItems.ListCount - 1, 3) = ![Unit] & ""
            
            cmbItems.List(cmbItems.ListCount - 1, 4) = Val(![FOB] & "")
            cmbItems.List(cmbItems.ListCount - 1, 5) = Val(![CIFAir] & "")
            cmbItems.List(cmbItems.ListCount - 1, 6) = Val(![CIFSea] & "")
            
            cmbItems.List(cmbItems.ListCount - 1, 7) = Val(![CnFAir] & "")
            cmbItems.List(cmbItems.ListCount - 1, 8) = Val(![CnFSea] & "")
            
            cmbDesc.List(cmbDesc.ListCount - 1, 1) = ![Description] & ""
            cmbDesc.List(cmbDesc.ListCount - 1, 2) = ![CompDesc] & ""
            
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

