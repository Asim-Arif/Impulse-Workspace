VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmRMPurchOrd 
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
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7155
      Left            =   45
      TabIndex        =   14
      Top             =   30
      Width           =   10560
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   285
         TabIndex        =   0
         Top             =   1050
         Width           =   2490
         _ExtentX        =   4392
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
      Begin VB.Frame FNew 
         BorderStyle     =   0  'None
         Height          =   570
         Left            =   285
         TabIndex        =   21
         Top             =   1350
         Width           =   10230
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
            Picture         =   "frmRMPurchOrd.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Add This Entry To Voucher"
            Top             =   15
            UseMaskColor    =   -1  'True
            Width           =   540
         End
         Begin VB.TextBox txtUnit 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   5265
            MaxLength       =   150
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   255
            Width           =   930
         End
         Begin VB.TextBox txtPrice 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   6225
            MaxLength       =   150
            TabIndex        =   22
            Top             =   255
            Width           =   930
         End
         Begin VB.TextBox txtQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   7185
            MaxLength       =   150
            TabIndex        =   7
            Top             =   255
            Width           =   930
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"frmRMPurchOrd.frx":0532
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
            TabIndex        =   24
            Top             =   0
            Width           =   9555
         End
         Begin VB.Label lblAmt 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   8145
            TabIndex        =   23
            Top             =   255
            Width           =   1410
         End
         Begin MSForms.ComboBox cmbCatID 
            Height          =   285
            Left            =   0
            TabIndex        =   3
            Top             =   255
            Width           =   1545
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "2725;503"
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
            Left            =   1575
            TabIndex        =   4
            Top             =   255
            Width           =   1215
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "2143;503"
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
            Left            =   2820
            TabIndex        =   5
            Top             =   255
            Width           =   2415
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4260;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.TextBox txtOrderNo 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2790
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   1
         Top             =   1050
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   10560
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Issue Raw Material"
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
            TabIndex        =   18
            Top             =   165
            Width           =   10605
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Issue Raw Material"
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
            TabIndex        =   19
            Top             =   180
            Width           =   10590
         End
      End
      Begin VB.CommandButton cmdDelete 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9915
         Picture         =   "frmRMPurchOrd.frx":05CE
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Remove This Entry From Voucher"
         Top             =   1965
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   540
      End
      Begin MSComCtl2.DTPicker OrderDT 
         Height          =   300
         Left            =   4260
         TabIndex        =   2
         Top             =   1050
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
         Format          =   61800451
         CurrentDate     =   37250
      End
      Begin MSComctlLib.ListView LV 
         Height          =   4095
         Left            =   285
         TabIndex        =   9
         Top             =   1965
         Width           =   9540
         _ExtentX        =   16828
         _ExtentY        =   7223
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
         NumItems        =   8
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
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Price"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Qty"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Amount"
            Object.Width           =   1852
         EndProperty
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   4665
         TabIndex        =   15
         Top             =   6405
         Width           =   5280
         Begin MSForms.CommandButton cmdCancel 
            Height          =   360
            Left            =   3510
            TabIndex        =   13
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
            TabIndex        =   11
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
            TabIndex        =   12
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
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Vender                                         Order No                 Order Date"
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
         TabIndex        =   28
         Top             =   810
         Width           =   5520
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
         Left            =   7470
         TabIndex        =   27
         Top             =   6090
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
         TabIndex        =   26
         Top             =   6090
         Width           =   1830
      End
      Begin MSForms.CheckBox chkPrint 
         Height          =   345
         Left            =   270
         TabIndex        =   25
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
      Begin VB.Label lblTotWords 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   2100
         TabIndex        =   20
         Top             =   6090
         Width           =   5385
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   8340
         TabIndex        =   16
         Top             =   6090
         Width           =   1500
      End
   End
End
Attribute VB_Name = "frmRMPurchOrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrderNo As String
Dim FromCat As Boolean
Public Sub AddPurchOrd(RMID As Long, Qty As Long)
    On Error GoTo err
    FromCat = True
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select GroupID,RMName,GroupName From VRM Where RMID=" & RMID, con, adOpenForwardOnly, adLockReadOnly
    
        'For i = 0 To cmbCatID.ListCount - 1
    
        '    If cmbCatID.List(i, 1) = ![GroupID] Then
        '        cmbCatID.ListIndex = i
        '        Exit For
        '    End If
    
        'Next i
        cmbCatID.AddItem ![GroupName] & ""
        cmbCatID.List(0, 1) = ![GroupID] & ""
        cmbCatID.ListIndex = 0
        'cmbItems.AddItem ![RMName] & ""
        'cmbItems.List(0, 1) = RMID
        'cmbItems.ListIndex = 0
        
        .Close
    End With
    Set rs = Nothing
    
    For i = 0 To cmbItems.ListCount - 1
    
            If cmbItems.List(i, 1) = RMID Then
                cmbItems.ListIndex = i
                Exit For
            End If
    
    Next i
    
    cmbCust.ClearVals
    Call cmbCust.AddVals(con, "VenderName", "RawMaterialVenders", "VendID", " Where VendID In (Select VendID From RMVendAssItems Where RMID=" & RMID & ")")
    
    Me.Show
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
Public Sub EditOrder(TempOrdNo As String)
    
    On Error GoTo err
    strOrderNo = TempOrdNo
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "Select * From VRMPOrders Where OrderNo='" & strOrderNo & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCust.Text = ![VenderName] & ""
        'cmbCountry = ![Country] & ""
        txtOrderNo = ![OrderNo] & ""
        OrderDT = ![OrderDT] & ""
        
        Do Until .EOF
            Set ITM = LV.ListItems.add(, ![RMID] & "'", .AbsolutePosition)
            'ITM.Tag = ![CompItemID] & "" 'cmbItems.List(cmbItems.ListIndex, 2)
            
            ITM.ListSubItems.add , , ![GroupDesc] & "" '![CatID] & ""
            ITM.ListSubItems.add , , ![RMID1] & ""
            'ITM.ListSubItems.Add(, ![ItemName] & "", ![ItemName] & "").Tag = ![Description] & ""
            ITM.ListSubItems.add , , ![RMName] & ""
            
            ITM.SubItems(4) = ![Unit] & ""
            
            ITM.SubItems(5) = Format(![Price] & "", "#0.00")
            ITM.SubItems(6) = Val(![Qty] & "")
            ITM.SubItems(7) = Val(![Qty] & "") * Val(![Price])
   
            .MoveNext
        Loop
        .Close
    End With
    Call GetTotal

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



Private Sub cmbCust_matched()
    On Error GoTo err
    If FromCat Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select Distinct GroupID,GroupDesc From VRMVendAssItems Where VendID=" & cmbCust.ID, con, adOpenForwardOnly, adLockReadOnly
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
        txtUnit = cmbItems.List(cmbItems.ListIndex, 2)
        txtPrice = cmbItems.List(cmbItems.ListIndex, 3)
    Else
        txtUnit = ""
        txtPrice = ""
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
    ElseIf Val(txtPrice) = 0 Then
        MsgBox "Price Cannot Be Zero", vbInformation
        Exit Sub
    ElseIf Val(txtQty) = 0 Then
        MsgBox "Quantity To Order Cannot Be Zero", vbInformation
        Exit Sub
    End If

    Dim ITM As ListItem
    Dim Sno As Integer
    If LV.ListItems.count > 0 Then
        Sno = Val(LV.ListItems(LV.ListItems.count).Text) + 1
    Else
        Sno = 1
    End If
    
    Set ITM = LV.ListItems.add(, cmbItems.List(cmbItems.ListIndex, 1) & "'", Sno)
    ITM.Tag = cmbItems.List(cmbItems.ListIndex, 2)
    
    ITM.ListSubItems.add , , cmbCatID
    ITM.ListSubItems.add , , cmbItems
    ITM.ListSubItems.add , , cmbDesc
    
    ITM.SubItems(4) = txtUnit
    
    ITM.SubItems(5) = Format(txtPrice, "#0.00")
    ITM.SubItems(6) = txtQty
    ITM.SubItems(7) = lblAmt
   
    'cmbItemID.RemoveItem cmbItemID.ListIndex
    'cmbItem.RemoveItem cmbItem.ListIndex
    
    Call GetTotal

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
    FromCat = False
    Call cmbCust.AddVals(con, "VenderName", "RawMaterialVenders", "VendID")
    
    OrderDT = ServerDate
    
    Me.KeyPreview = True
    cmbCust.ListHeight = 2500
    
    txtOrderNo = "R-" & Val(con.Execute("Select Max(Cast(SubString(OrderNo,3,50) as int)) From RMPOrders").Fields(0).value & "") + 1
    
    Exit Sub
err:
    MsgBox err.Description
   
End Sub
Private Sub Form_Resize()
On Error Resume Next
   FALL.Move ((Me.ScaleWidth - FALL.Width) / 2), ((Me.ScaleHeight - FALL.Height) / 2)
End Sub

Private Sub LV_Click()
    If LV.ListItems.count > 0 Then
        Call LV_ItemClick(LV.SelectedItem)
    End If
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With cmdDelete
        .Move .Left, LV.Top + Item.Top
        .Visible = True
    End With
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
    lblTotWords = ConvertInWords(lblTotal, "Rs")
    
End Sub



Private Function Saved() As Boolean
    
    Saved = False
    If txtOrderNo = "" Then
        MsgBox "Please Enter Order No.", vbInformation
        txtOrderNo.SetFocus
        Exit Function
    End If
    
    If LV.ListItems.count = 0 Then
        MsgBox "No Items Added to This Order", vbInformation
        Exit Function
    End If
    
    Call StartTrans(con)
    '----------------------------------------------------------------
    If strOrderNo <> "" Then
        con.Execute "Delete from RMPOrders Where OrderNo='" & _
         strOrderNo & "'"
    End If
    
    'Save Order
    con.Execute "Insert Into RMPOrders (OrderNo," & _
     "OrderDT,VendID) values('" & txtOrderNo & "','" & _
     OrderDT & "','" & cmbCust.ID & "')"
     
    'Now Save Items
    For i = 1 To LV.ListItems.count
        con.Execute "Insert into RMPOrderItems (OrderNo,RMID," & _
         "Price,Qty) Values ('" & txtOrderNo & "','" & _
         Left(LV.ListItems(i).key, Len(LV.ListItems(i).key) - 1) & _
         "'," & Val(LV.ListItems(i).SubItems(5)) & _
         "," & Val(LV.ListItems(i).SubItems(6)) & ")"
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
    
        If FromCat Then
            .Open "Select * From VRM Where GroupID=" & cmbCatID.List(cmbCatID.ListIndex, 1), con, adOpenForwardOnly, adLockReadOnly
            'Exit Sub
        Else
            .Open "Select * From VRMVendAssItems Where GroupID=" & cmbCatID.List(cmbCatID.ListIndex, 1) & " and VendID=" & cmbCust.ID, con, adOpenForwardOnly, adLockReadOnly
            
        End If
        
            
        
        cmbItems.Clear
        cmbDesc.Clear
        
        
        Do Until .EOF
        
            cmbItems.AddItem ![RMID1] & ""
            cmbDesc.AddItem ![RMName] & ""
            
            cmbItems.List(cmbItems.ListCount - 1, 1) = ![RMID] & ""
            cmbItems.List(cmbItems.ListCount - 1, 2) = ![Unit] & ""
            cmbItems.List(cmbItems.ListCount - 1, 3) = ![Rate] & ""
            
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
