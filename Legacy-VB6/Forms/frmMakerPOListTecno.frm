VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmMakerPOListTecno 
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
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
      TabIndex        =   3
      Top             =   -45
      Width           =   11865
      Begin VB.Frame Frame1 
         Height          =   480
         Left            =   2820
         TabIndex        =   13
         Top             =   7875
         Width           =   5475
         Begin VB.ComboBox cmbfield 
            Height          =   315
            ItemData        =   "frmMakerPOListTecno.frx":0000
            Left            =   60
            List            =   "frmMakerPOListTecno.frx":0019
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   135
            Width           =   1875
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4395
            TabIndex        =   15
            Top             =   135
            Width           =   1020
         End
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   1965
            TabIndex        =   14
            Top             =   135
            Width           =   2430
         End
      End
      Begin ComboList.Usercontrol1 cmbItems 
         Height          =   285
         Left            =   7155
         TabIndex        =   12
         Top             =   360
         Width           =   3720
         _ExtentX        =   6562
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
      Begin VB.ComboBox cmbRange 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMakerPOListTecno.frx":0060
         Left            =   2400
         List            =   "frmMakerPOListTecno.frx":0076
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   345
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   5640
         TabIndex        =   1
         Top             =   345
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
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
         Format          =   220987395
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   4110
         TabIndex        =   0
         Top             =   345
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
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
         Format          =   220987395
         CurrentDate     =   37055
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   195
         TabIndex        =   6
         Top             =   345
         Width           =   2190
         _ExtentX        =   3863
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
      Begin MSComctlLib.ListView LV 
         Height          =   7110
         Left            =   75
         TabIndex        =   7
         Top             =   765
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   12541
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Maker"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "PO #"
            Object.Width           =   2302
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Process"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Item ID"
            Object.Width           =   1667
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Value"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Status"
            Object.Width           =   1931
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Lot #"
            Object.Width           =   2469
         EndProperty
      End
      Begin MSForms.CommandButton cmdAuth 
         Height          =   375
         Left            =   8370
         TabIndex        =   11
         Top             =   7935
         Width           =   2040
         ForeColor       =   0
         Caption         =   "Authorize Selected"
         PicturePosition =   327683
         Size            =   "3598;661"
         Accelerator     =   65
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdCheckAll 
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   10
         Top             =   7920
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Un-Check All"
         PicturePosition =   327683
         Size            =   "2355;661"
         Accelerator     =   85
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdCheckAll 
         Height          =   375
         Index           =   0
         Left            =   60
         TabIndex        =   9
         Top             =   7920
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Check All"
         PicturePosition =   327683
         Size            =   "2355;661"
         Accelerator     =   104
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   375
         Left            =   10455
         TabIndex        =   8
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
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   10920
         TabIndex        =   2
         Top             =   315
         Width           =   840
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "1482;635"
         Accelerator     =   82
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMakerPOListTecno.frx":00C6
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
         Left            =   165
         TabIndex        =   4
         Top             =   120
         Width           =   11130
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   10
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakerPOListTecno.frx":0156
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMakerPOListTecno.frx":0312
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopMenu"
      Begin VB.Menu mnuPrintPurchaseOrder 
         Caption         =   "Print Purchase Order"
      End
      Begin VB.Menu mnuPrintGatePass 
         Caption         =   "Print Gate-Pass"
      End
      Begin VB.Menu mnuPrintList 
         Caption         =   "Print List"
      End
      Begin VB.Menu mnuPrintProcessOrder 
         Caption         =   "Print Process Order"
      End
      Begin VB.Menu mnuDashDelete 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Order"
      End
   End
End
Attribute VB_Name = "frmMakerPOListTecno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCourierInvoice As Boolean
Dim bShipping As Boolean
Dim bCustomInvoice As Boolean

Public Sub ShowMe(Optional p_bCustomInvoice As Boolean = True)
    bCustomInvoice = p_bCustomInvoice
    Me.Show
End Sub

Public Sub ShowForShipping()
    bShipping = True
    Me.Show
End Sub

Private Sub cmbRange_Click()

    DtTo = GetServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
        Case 0
            DtFrom = Date
        Case 1
            DtFrom = DateAdd("d", -15, Date)
        Case 2
            DtFrom = DateAdd("d", -30, Date)
        Case 3
            DtFrom = DateAdd("d", -60, Date)
        Case 4
            DtFrom = DateAdd("d", -90, Date)
        Case 5
            DtFrom.Enabled = True
            DtTo.Enabled = True
            DtFrom.SetFocus
            Exit Sub
    End Select
    
End Sub

Private Sub cmdAuth_Click()

    Dim i As Integer
 
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked And CBool(.ListSubItems(7).Tag) = False Then
                con.Execute "UPDATE VendIssued SET Authorized=1 WHERE EntryID=" & .Tag
                .ListSubItems(6).Text = "Authorized"
                .ListSubItems(6).Tag = True
            End If
        End With
    Next
    Call RefreshLV
    
End Sub

Private Sub cmdCheckAll_Click(Index As Integer)
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = (Index = 0)
    Next
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdFind_Click()

    Dim StartIdx As Integer

    If LV.SelectedItem Is Nothing Then
       StartIdx = 1
    Else
       If LV.SelectedItem.Index = LV.ListItems.count Then
          StartIdx = 1
       Else
          StartIdx = LV.SelectedItem.Index
       End If
    End If

    Dim fldValue As String




    For i = StartIdx To LV.ListItems.count
       
       
       If cmbField.ListIndex = 0 Then
          fldValue = Trim(Replace(LV.ListItems(i), "»", ""))
       Else
          fldValue = Trim(Replace(LV.ListItems(i).SubItems(cmbField.ListIndex), "»", ""))
       End If
    
       If LCase(fldValue) Like "*" & LCase(Trim(txtFind)) & "*" Then
          LV.ListItems(i).Selected = True
          LV.ListItems(i).EnsureVisible
          If i <> StartIdx Then Exit For
       End If
    Next
    
    If i = LV.ListItems.count + 1 Then
        MsgBox "No Mathes Found.", vbInformation
        Set LV.SelectedItem = Nothing
    End If

    
    
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub


Private Sub Command1_Click()

End Sub

Private Sub DtFrom_Change()
    'Call setDateRange(DtFrom)
End Sub

Private Sub DtTo_Change()
    'Call setDateRange(DtTo)
End Sub


Private Sub Form_Load()
    
    mnupop.Visible = False
    cmbCust.ListHeight = 3800
    cmbItems.ListHeight = 6000
    cmbCust.AddItem "<All Makers>", "0"
    
    Call cmbCust.AddVals(con, "'{ ' + VendID1 + ' } ' + VenderName ", "Makers", "VendID")
    
    cmbCust.ID = "0"
    DtFrom = Date
    DtTo = Date
    cmbRange.ListIndex = 0
     
    cmbItems.AddItem "<All Articles>", "0"
    Call cmbItems.AddVals(con, "ItemID + ' ' + ItemName + ' { ' + ISNULL(Attributes,'') + ' }'", "Items", "ItemID", " WHERE InActive=0 ORDER BY ItemName")
    cmbItems.ID = "0"
     
    With cmbField
        .Clear
        For i = 1 To LV.ColumnHeaders.count
            cmbField.AddItem LV.ColumnHeaders(i).Text
        Next
    End With
    
End Sub

Private Sub RefreshLV()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim Cond As String
    Dim DTRange As String
    Dim ITM As ListItem
    
    Screen.MousePointer = vbHourglass
    
    If cmbRange.ListIndex = 0 Then
        DTRange = " OnlyDT Between '" & DateAdd("d", 0, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 1 Then
        DTRange = " OnlyDT Between '" & DateAdd("d", -15, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 2 Then
        DTRange = " OnlyDT Between '" & DateAdd("d", -30, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 3 Then
        DTRange = " OnlyDT Between '" & DateAdd("d", -60, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 4 Then
        DTRange = " OnlyDT Between '" & DateAdd("d", -90, GetServerDate) & "' and '" & GetServerDate & "'"
    ElseIf cmbRange.ListIndex = 5 Then
        DTRange = " OnlyDT Between '" & DtFrom & "' and '" & DtTo & "'"
    End If
    Cond = " WHERE " & DTRange
    If cmbCust.ID <> "0" Then
        Cond = Cond & " AND VendID=" & cmbCust.ID & " "
    End If
    If cmbItems.ID <> "0" Then
        Cond = Cond & " AND ItemID='" & cmbItems.ID & "' "
    End If
    'Cond = Cond & " AND Left(CustomInvoice,4)='VTL '"
    With rs
        
        .Open "Select * From VVendIssued " & Cond & " ORDER BY DT", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        'LV.Sorted = False
        Do Until .EOF
            Set ITM = LV.ListItems.add(, Val(![EntryID] & "") & "'", .AbsolutePosition)
            ITM.Tag = Val(![EntryID] & "")
            ITM.ListSubItems.add(, , "{ " & !VendID1 & " } " & !VenderName).Tag = "" '![ImpMaterial] & ""
            ITM.ListSubItems.add(, , ![RecieptID] & "").Tag = !code & ""
            ITM.ListSubItems.add(, , Format(![DT] & "", "dd-MMM-yyyy")).Tag = ""
          
            ITM.ListSubItems.add , , !Description & ""
            ITM.ListSubItems.add , , ![ITemID] & ""
                     
            ITM.ListSubItems.add , , Val(!TotalValue & "")
            ITM.ListSubItems.add(, , IIf(!Authorized, "Authorized", "Un-Auth.")).Tag = !Authorized
            
            ITM.ListSubItems.add(, , Val(!IPODL_RefID & "")).Tag = ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    FAcc.Top = (Me.ScaleHeight - FAcc.Height) / 2
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If LV.SortKey + 1 = ColumnHeader.Index Then
        If LV.SortOrder = lvwAscending Then
            LV.SortOrder = lvwDescending
            ColumnHeader.Icon = "Down"
        Else
            LV.SortOrder = lvwAscending
            ColumnHeader.Icon = "Up"
        End If
    Else
        LV.ColumnHeaderIcons = Nothing
        Set LV.ColumnHeaderIcons = ImageList2
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        ColumnHeader.Icon = "Down"
    End If
    Call RefreshSnos
    'If LV.Sorted = False Then LV.Sorted = True
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 2 Then Exit Sub
    If LV.ListItems.count = 0 Then
        Call ShowMenus(False)
    Else
        Call ShowMenus(True)
    End If
'    If bShipping Then
'        ShowMenusForShipping
'    ElseIf bCourierInvoice Then
'        ShowMenusforCourier
'    Else
'        If LV.ListItems.Count = 0 Then
'            Call ShowMenus(False)
'        Else
'            Call ShowMenus(True)
'        End If
'    End If
    'If LV.ListItems.Count = 0 Or Button <> 2 Or OutOfFinancialYear Then Exit Sub

    'if a Voucher Entry Show Delete Menu
    'If InStr(LV.SelectedItem, "-") > 0 Then Me.PopupMenu mnuPop
    
End Sub

Private Sub ShowMenusForShipping()

    On Error Resume Next
    Dim c As Control
    For Each c In Controls
        If TypeOf c Is Menu Then
            If c.Name <> "mnuPrintPList" Then
                c.Visible = False
            End If
        End If
    Next
    
    On Error GoTo 0
    
    If LV.ListItems.count > 0 Then
        mnuPrintPList.Visible = True
        mnuCustomPList.Visible = True
        mnuCustomPListOther.Visible = True
        mnuCommercialPList.Visible = True
        mnuCommercialPListOther.Visible = True
        mnuDashPackingLabels.Visible = True
        mnuPrintPackingLabels.Visible = True
    End If
'    mnuPrintDecDocs.Visible = False
    Me.PopUpMenu mnupop
    
End Sub

Private Sub ShowMenus(vis As Boolean)
    
    mnuPrintProcessOrder.Visible = False
    If LV.ListItems.count > 0 Then
        If Left(LV.SelectedItem.ListSubItems(2).Tag, 2) = "SC" Then
            Me.mnuPrintProcessOrder.Visible = True
        End If
    End If
    Me.PopUpMenu mnupop
    
End Sub

Private Sub ShowMenusforCourier()
         
    On Error Resume Next
    Dim c As Control
    For Each c In Controls
        If TypeOf c Is Menu Then
            c.Visible = False
        End If
    Next
    
    On Error GoTo 0
    mnuAdd.Visible = True
    If LV.ListItems.count > 0 Then
        dd.Visible = True
        mnuEdit.Visible = True
        mnuComDocs.Visible = True
        mnuPrintComInv.Visible = True
    End If
    mnuPrintDecDocs.Visible = False
    Me.PopUpMenu mnupop
    
End Sub


Private Sub mnuAdd_Click()
    'Load frmNewCustomInvoice
    'frmNewCustomInvoice.Show
    'ShowForm frmNewCustomInvoice, MainForm, False, "New Custom Invoice"
    'frmNewCustomInvoice.AddNewProforma
    Load frmNewCustomInvoice
    frmNewCustomInvoice.AddNewInvoice True, bCourierInvoice
End Sub

Private Sub mnuAddShippInfo_Click()

    Load frmEditShippInfo
    'ShowForm f, MainForm, , "Shipp. Info"
    Call frmEditShippInfo.EditEntry(LV.SelectedItem.SubItems(3))
    
End Sub

Private Sub mnuBillOfExchange_Click()
    Call PrintBillOfExchange
End Sub

Private Sub PrintBillOfExchange(Optional Firstrpt As Boolean = True)

    On Error GoTo err
    Me.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptBillOfExchange.rpt")
    rpt.FormulaFields(3).Text = Firstrpt
    Dim frm As New frmPrevRpt
    frm.ShowReport "{CustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    Me.MousePointer = vbNormal
    
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description

End Sub

    
Private Sub mnuBillOfExchange2_Click()
    Call PrintBillOfExchange(False)
End Sub

Private Sub mnuCommercialPList_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3), False, True, LV.SelectedItem.SubItems(1) = "DECATHLON", LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2))
End Sub

Private Sub mnuCommercialPListOther_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3), True, True, LV.SelectedItem.SubItems(1) = "DECATHLON", LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2))
End Sub

Private Sub mnuCoveringLetter_Click()
    
    On Error GoTo err
    Me.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptComCoveringLetter.rpt")
    If LCase(CurrentUserName) = "mnaeem" Then
        rpt.FormulaFields(3).Text = "'Naeem Ahmed'"
    End If
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "{CustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub mnuCustomPList_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3))
End Sub

Private Sub mnuCustomPListOther_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3), True)
End Sub

Private Sub mnuDelPackingInfo_Click()
    
    If LV.ListItems.count = 0 Then Exit Sub
    Dim iAffect As Integer
    If MsgBox("Are you sure to Delete Packing List of Invoice # " & LV.SelectedItem.SubItems(3) & " ?", vbQuestion + vbYesNo) = vbYes Then
        Call StartTrans(con)
        con.Execute "Delete From CustomPList Where CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'", iAffect
        con.CommitTrans
        If iAffect > 0 Then
            LV.SelectedItem.ListSubItems(2).Tag = 0
        End If
    End If
    
End Sub


Private Sub mnuEForm_Click()

    Load frmPrintFormE
    If LV.ListItems.count > 0 Then
        frmPrintFormE.ShowForm LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
    Else
        frmPrintFormE.Show , MainForm
    End If
    
End Sub

Private Sub mnuGSP_Click()

'    Load frmGSP
'    If LV.ListItems.Count > 0 Then
'        frmGSP.ShowForm LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
'    Else
'        frmGSP.Show , MainForm
'    End If
    
End Sub

Private Sub mnuPacking_Click()

'    If Left(mnuPacking.Caption, 3) = "Add" Then
'
'        Load frmNewPList
'        frmNewPList.ShowMe LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
'    Else
'        Load frmNewPList
'        Call frmNewPList.EditPList(LV.SelectedItem.SubItems(3))
'        ShowForm frmNewPList, MainForm, , "Packing List"
'    End If
    
End Sub

Private Sub mnuPrint_Click()
    Call PrintCustomInvoice
End Sub

Private Sub PrintCustomInvoice(Optional bImpMaterialDetail As Boolean = False)

    On Error GoTo err

    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    If bCustomInvoice Then
        Set rpt = rptApp.OpenReport(RptPath & "\rptCustomInvoice.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\rptCommercialInvoice.rpt")
    End If
    'rpt.FormulaFields(15).Text = bImpMaterialDetail
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt
    
    Screen.MousePointer = vbDefault


    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub mnuPrint10_Click()
    Call PrintCommercialInvoice(True, False)
End Sub

Private Sub mnuPrint5_Click()
    Call PrintCommercialInvoice(True, True)
End Sub

Private Sub mnuPrintBillofExchange_Click()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptBillOfExchange.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{VRptProformas.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintComInv_Click()
    Call PrintCommercialInvoice
End Sub

Private Sub PrintCommercialInvoice(Optional bPer As Boolean = False, Optional bFivePer As Boolean = True)
    
    On Error GoTo err

    Screen.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt
    Dim rpt As CRAXDDRT.Report
    
    If UCase(LV.SelectedItem.SubItems(1)) = "MIKASA" And LV.SelectedItem.SubItems(2) = "JAPAN" Then
        Set rpt = rptApp.OpenReport(rptexportpath & "\rptCommercialInvoiceMikasaJP.rpt")
    ElseIf bPer Then
        Set rpt = rptApp.OpenReport(rptexportpath & "\rptCommercialInvoiceFairTrade.rpt")
        rpt.FormulaFields(22).Text = bFivePer
    ElseIf LV.SelectedItem.SubItems(1) = "DECATHLON" Then
        Set rpt = rptApp.OpenReport(rptexportpath & "\rptDecathlonCommercialInvoice.rpt")
    Else
        Set rpt = rptApp.OpenReport(rptexportpath & "\rptCommercialInvoice.rpt")
    End If
    f.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt
    
    If Not bPer Then
        If LV.SelectedItem.SubItems(2) = "Canada" Then
            Dim rpt1 As CRAXDDRT.Report
            Set rpt1 = rptApp.OpenReport(rptexportpath & "\rptCommercialInvoice.rpt")
            Dim f1 As New frmPrevRpt
            Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
            Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
            
            Set FormulaFields = rpt1.FormulaFields
            
            For Each FormulaField In FormulaFields
                If FormulaField.Name = "{@forHeading}" Then
                    FormulaField.Text = "'C A N A D I A N   C U S T O M S   I N V O I C E'"
                    Exit For
                End If
            Next
            f1.ShowReport "{VrptCustomInvoiceDetail.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt1
        End If
    End If
    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub mnuPrintSummary_Click()

    On Error GoTo err
    Me.MousePointer = vbHourglass
    
    Call StartTrans(con)
    con.Execute "Delete From PrintFOrderList"
    
    For i = 1 To LV.ListItems.count
        con.Execute "Insert Into PrintFOrderList(Sno,Customer,Country," & _
         "OrderNo,OrderDT,OrderAmt,Status,Curr) Values(" & LV.ListItems(i) & _
         ",'" & LV.ListItems(i).SubItems(1) & "','" & LV.ListItems(i).SubItems(2) & _
         "','" & LV.ListItems(i).SubItems(3) & "','" & _
         LV.ListItems(i).SubItems(4) & "'," & Val(LV.ListItems(i).ListSubItems(5).Tag & "") & _
         ",'" & LV.ListItems(i).SubItems(6) & "','" & LV.ListItems(i).Tag & "')"
    Next i
    
    con.CommitTrans
    
    Load frmPrevRpt
    
    Dim strFromTo As String
    Dim iNumber As Integer
    If cmbRange.ListIndex = 0 Then
        iNumber = -15
    ElseIf cmbRange.ListIndex = 1 Then
        iNumber = -30
    ElseIf cmbRange.ListIndex = 2 Then
        iNumber = -60
    ElseIf cmbRange.ListIndex = 3 Then
        iNumber = -90
    End If
    
    If cmbRange.ListIndex = 4 Then
        strFromTo = Format(DtFrom, "dd-MMM-yyyy") & " To  " & Format(DtTo, "dd-MMM-yyyy")
    Else
        strFromTo = Format(DateAdd("d", iNumber, ServerDate), "dd-MMM-yyyy") & " To  " & Format(ServerDate, "dd-MMM-yyyy")
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptFProformaList.rpt")
    rpt.FormulaFields(3).Text = "'" & cmbCust.Text & "'"
    rpt.FormulaFields(4).Text = "'" & strFromTo & "'"
    
    frmPrevRpt.ShowReport "", rpt
    
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub mnuPrintCoverLetter_Click()

    On Error GoTo err

    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptDecathlonCoveringLetter.rpt")
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VCustomInvoice.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt

    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub mnuPrintDecDocs_Click()
    Load frmPrintDecathlonDocs
    frmPrintDecathlonDocs.ShowMe (LV.SelectedItem.SubItems(3))
End Sub

Private Sub mnuPrintInvWDetail_Click()

    Call PrintCustomInvoice(True)
    
End Sub

Private Sub mnuPrintPackingLabels_Click()

    On Error GoTo err
    
    If LV.SelectedItem.SubItems(1) <> "DECATHLON" Then
        MsgBox "Packing Labels are available only for DECATHLON at the moment.", vbInformation
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report, f As New frmPrevRpt
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptDecathlonPackingLabels.rpt")
    f.ShowReport "{VCustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    
    Me.MousePointer = vbNormal
    
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub mnuShippingInstructions_Click()
'    Load frmShippingInstructions
'    If LV.ListItems.Count = 0 Then
'        frmShippingInstructions.Show , MainForm
'    Else
'        frmShippingInstructions.ShowForm LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
'    End If
End Sub

Private Sub mnuUpdateCustomPrice_Click()
    If LV.ListItems.count = 0 Then Exit Sub
    If MsgBox("Are you Sure, You want to Update Custom Prices?", vbInformation + vbYesNo) = vbNo Then Exit Sub
    
    Dim dCustomPrice As Double
    dCustomPrice = Val(con.Execute("Select SUM(Qty*Price)/SUM(Qty) From CustomInvoiceItems Where CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'").Fields(0).Value & "")
    con.Execute "Update CustomInvoiceItems SET CustomPrice=" & dCustomPrice & " WHERE CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'"
    
End Sub

Private Sub txtDesc_GotFocus()
    cmdRefresh.Default = False
    txtDesc.SelStart = 0
    txtDesc.SelLength = Len(txtDesc)
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
  
    If KeyAscii = 13 Then
    
        If Trim(txtDesc = "") Then
            MsgBox "Invalid Description", vbInformation
            Exit Sub
        End If
    
        con.Execute "Update Vouchers Set [Description]='" & txtDesc & "' Where Sno=" & Val(LV.SelectedItem.key)
        LV.SelectedItem.ListSubItems(3) = txtDesc & txtDesc.Tag
        txtDesc.Visible = False
        cmdRefresh.Default = False
    ElseIf KeyAscii = 27 Then
        txtDesc.Visible = False
        cmdRefresh.Default = False
    End If

End Sub

Private Sub txtDesc_LostFocus()
    txtDesc.Visible = False
End Sub
Private Sub RefreshSnos()
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Text = i
    Next i
End Sub

Private Sub mnuDelete_Click()

    On Error GoTo err
    If LV.ListItems.count = 0 Then Exit Sub
    'Check for Receivings...
    If LV.SelectedItem.SubItems(6) = "Un-Auth." Then
        If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo) = vbYes Then
            con.Execute "DELETE FROM VendIssued WHERE RecieptID='" & LV.SelectedItem.SubItems(2) & "'"
            Call RefreshLV
        End If
    Else
        If MsgBox("Are you sure to Delete this Aurhorized Order ?", vbQuestion + vbYesNo) = vbYes Then
            'Check Rcving Against this Order.
            Dim lCount As Long
            lCount = GetSingleLongValue("COUNT(Issuance_RefID)", "VendReceived", "WHERE Issuance_RefID=" & Val(LV.SelectedItem.Tag))
            If lCount > 0 Then
                MsgBox "Can't Delete, It has been Received.", vbInformation
                Exit Sub
            Else
                con.Execute "DELETE FROM VendIssued WHERE RecieptID='" & LV.SelectedItem.SubItems(2) & "'"
                Call RefreshLV
            End If
        End If
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintGatePass_Click()
    Call PrintIssSlip(True)
End Sub

Private Sub mnuPrintList_Click()

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\IssList.rpt")
    Dim strList As String
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                strList = strList & .Tag & ","
            End If
        End With
    Next
    strList = Left(strList, Len(strList) - 1)
    Dim f As New frmPrevRpt
    
    Dim strSelection As String
    strSelection = "{VendIssued.EntryID} IN[" & strList & "]"
    
    Dim rs As New ADODB.Recordset
    Dim strOrders As String
    strOrders = ""
    With rs
        .Open "SELECT DISTINCT OrderNo FROM VendIssdDetail WHERE RefID IN(" & strList & ")", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            strOrders = strOrders & !OrderNo & ","
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    strOrders = Left(strOrders, Len(strOrders) - 1)
     
    rpt.FormulaFields.GetItemByName("Orders").Text = "'" & strOrders & "'"
    
    f.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintProcessOrder_Click()
   
    
    Dim f As New frmPrintProcessOrder
    Load f
    f.ShowMe Val(LV.SelectedItem.Tag)
    
    Exit Sub
    
    
End Sub

Private Sub mnuPrintPurchaseOrder_Click()
    Call PrintIssSlip
End Sub

Private Sub PrintIssSlip(Optional bPrintGP As Boolean = False)

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    If bPrintGP Then
        Set rpt = rptApp.OpenReport(RptPath & "\GatePassAndDuplicate.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\IssSlip.rpt")
    End If
    Dim strList As String
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                strList = strList & .Tag & ","
            End If
        End With
    Next
    
    Dim f As New frmPrevRpt
    If strList = "" Then
        f.ShowReport "{VendIssued.EntryID} IN[" & LV.SelectedItem.Tag & "]", rpt
    Else
        f.ShowReport "{VendIssued.EntryID} IN[" & Left(strList, Len(strList) - 1) & "]", rpt
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub txtFind_Change()
    cmdFind.Enabled = Len(Trim(txtFind)) > 0
End Sub

