VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmVendRcvItemsHnD 
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleMode       =   0  'User
   ScaleWidth      =   10641.08
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10560
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Order Status"
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
         TabIndex        =   5
         Top             =   165
         Width           =   10605
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Order Status"
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
         TabIndex        =   6
         Top             =   180
         Width           =   10590
      End
   End
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   6555
      Left            =   15
      TabIndex        =   0
      Top             =   630
      Width           =   11850
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   195
         TabIndex        =   2
         Top             =   345
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
      Begin MSComctlLib.ListView LV 
         Height          =   5355
         Left            =   195
         TabIndex        =   9
         Top             =   660
         Width           =   11565
         _ExtentX        =   20399
         _ExtentY        =   9446
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   1014
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item ID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Order Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Shipped"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Shippable"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Pack Mode"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   375
         Left            =   8385
         TabIndex        =   10
         Top             =   6105
         Width           =   1650
         ForeColor       =   0
         Caption         =   "Print     "
         PicturePosition =   327683
         Size            =   "2910;661"
         Picture         =   "frmVendRcvItemsHnD.frx":0000
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbOrders 
         Height          =   285
         Left            =   4425
         TabIndex        =   8
         Top             =   345
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
      Begin MSForms.CommandButton cmdClose 
         Height          =   375
         Left            =   10110
         TabIndex        =   7
         Top             =   6105
         Width           =   1650
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "2910;661"
         Picture         =   "frmVendRcvItemsHnD.frx":0112
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbCountry 
         Height          =   285
         Left            =   2535
         TabIndex        =   3
         Top             =   345
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
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         Caption         =   " Customer                                   Country                           Order No"
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
         Left            =   180
         TabIndex        =   1
         Top             =   120
         Width           =   6105
      End
   End
End
Attribute VB_Name = "frmVendRcvItemsHnD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    On Error GoTo err
    
    If cmbCust.ID <> "0" Then
        Cond = " Where CustCode='" & cmbCust.ID & "' "
    Else
        Cond = ""
    End If
    
    If cmbCountry.ListIndex <> 0 Then
        If Cond = "" Then
            Cond = " Where Country='" & cmbCountry & "'"
        Else
            Cond = Cond & " and Country='" & cmbCountry & "'"
        End If
    End If
    
    'Now Get Rid Of HICO Order
    If Cond = "" Then
        Cond = " Where OrderNo<>'HICO Order'"
    Else
        Cond = Cond & " and OrderNo<>'HICO Order'"
    End If
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select OrderNo From FCustomerOrders " & Cond, con, adOpenForwardOnly, adLockReadOnly
        cmbOrders.Clear
        Do Until .EOF
            cmbOrders.AddItem .Fields(0) & ""
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

    On Error GoTo err
        
    If cmbOrders.MatchFound Then
        Call ShowOrderDetail(cmbOrders.Text)
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub cmbCust_matched()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Sql = "Select Distinct Country From ForeignCustomers "
    If cmbCust.ID <> "0" Then
        Sql = Sql & " Where CustCode='" & cmbCust.ID & "'"
    End If
    
    With rs
        .Open Sql, con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        cmbCountry.AddItem "<All Countries>"
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            'cmbCountry.List(cmbCountry.ListCount - 1, 1) = ![Curr] & ""
            .MoveNext
        Loop
        .Close
    End With
    If cmbCountry.ListCount = 2 Then
        cmbCountry.ListIndex = 1
    Else
        cmbCountry.ListIndex = 0
    End If
    
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdClose_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub Form_Load()
        

    cmbCust.ListHeight = 2500
    cmbCust.AddItem "<All Customers>", "0"
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where CustCode<>'HICO'")
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Dim TheHeight As Long
    Frame1.Left = (Me.ScaleWidth - Frame1.Width) / 2
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    
End Sub

Private Sub ShowOrderDetail(TempOrdNo As String)
    
    On Error GoTo err
    strOrderNo = TempOrdNo
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    LV.ListItems.Clear
    With rs
        .Open "Select * From VVendRcvHnD Where OrderNo='" & TempOrdNo & "'", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            Set ITM = LV.ListItems.Add(, ![ItemCode] & "'", .AbsolutePosition)
            'itm.Tag = ![CompItemID] & "" 'cmbItems.List(cmbItems.ListIndex, 2)
            ITM.ListSubItems.Add , , ![ItemCode] & ""
            ITM.ListSubItems.Add(, , ![ItemName] & "").ToolTipText = ![ItemName] & ""
            ITM.ListSubItems.Add , , ![Qty] & ""
            ITM.ListSubItems.Add , , ![ShippedQty] & ""
            ITM.ListSubItems.Add , , ![ShippableQty] & ""
            ITM.ListSubItems.Add , , ![PackMode] & ""
            .MoveNext
        Loop
        .Close
    End With
    'Call GetTotal

    txtQty = 0
    'cmbItems.SetFocus
    
    Set rs = Nothing
    'Me.Show
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

