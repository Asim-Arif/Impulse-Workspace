VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmVendIssItems 
   ClientHeight    =   9615
   ClientLeft      =   1110
   ClientTop       =   60
   ClientWidth     =   11805
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9615
   ScaleWidth      =   11805
   WindowState     =   2  'Maximized
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
      Left            =   4710
      TabIndex        =   14
      Top             =   8685
      Width           =   5280
      Begin MSForms.CommandButton cmdSave 
         Height          =   360
         Left            =   75
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   195
         Width           =   1665
         Caption         =   "Cancel            "
         PicturePosition =   327683
         Size            =   "2937;635"
         Picture         =   "frmVendIssItems.frx":0000
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
         TabIndex        =   15
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
   Begin ComboList.Usercontrol1 cmbOrders 
      Height          =   285
      Left            =   5070
      TabIndex        =   3
      Top             =   810
      Width           =   1755
      _ExtentX        =   3096
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
      Left            =   795
      TabIndex        =   12
      Top             =   810
      Width           =   2355
      _ExtentX        =   4154
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   3675
      Left            =   405
      TabIndex        =   9
      Top             =   4935
      Width           =   10950
      _ExtentX        =   19315
      _ExtentY        =   6482
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Makers/Suppliers"
      TabPicture(0)   =   "frmVendIssItems.frx":0112
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LVMaker"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtRate"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Issued Items"
      TabPicture(1)   =   "frmVendIssItems.frx":012E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LVIssued"
      Tab(1).Control(1)=   "txtRate1"
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtRate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72840
         TabIndex        =   18
         Top             =   1590
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4950
         TabIndex        =   13
         Top             =   645
         Visible         =   0   'False
         Width           =   1350
      End
      Begin MSComctlLib.ListView LVMaker 
         Height          =   3180
         Left            =   90
         TabIndex        =   10
         Top             =   390
         Width           =   10770
         _ExtentX        =   18997
         _ExtentY        =   5609
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S No."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Vender Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Rate"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Qty To Issue"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView LVIssued 
         Height          =   3195
         Left            =   -74925
         TabIndex        =   11
         Top             =   390
         Width           =   10800
         _ExtentX        =   19050
         _ExtentY        =   5636
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
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
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Vender"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Iss ID"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Order No"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Item Code"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Description"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Rate"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Issued Qty"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin ComboList.Usercontrol1 cmbCust 
      Height          =   285
      Left            =   3180
      TabIndex        =   5
      Top             =   810
      Width           =   1875
      _ExtentX        =   3307
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   6825
      TabIndex        =   8
      Top             =   810
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   60817411
      CurrentDate     =   37559
   End
   Begin VB.CommandButton cmdUnAssign 
      Height          =   240
      Left            =   1095
      Picture         =   "frmVendIssItems.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "UnAssign Selected Item..."
      Top             =   6105
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   435
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   -15
      Top             =   -300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   555
      Top             =   -300
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
            Picture         =   "frmVendIssItems.frx":024C
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVendIssItems.frx":0408
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -15
      Top             =   -300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVendIssItems.frx":05C4
            Key             =   "Active"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVendIssItems.frx":0A18
            Key             =   "Inactive"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV 
      Height          =   3675
      Left            =   420
      TabIndex        =   7
      Top             =   1215
      Width           =   10860
      _ExtentX        =   19156
      _ExtentY        =   6482
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
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
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Customer"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Order No"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Item Code"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Description"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Issueable Qty"
         Object.Width           =   2822
      EndProperty
   End
   Begin VB.Frame fTitle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   795
      TabIndex        =   0
      Top             =   0
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Items"
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
         Left            =   4440
         TabIndex        =   1
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Items"
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
         Left            =   4455
         TabIndex        =   2
         Top             =   135
         Width           =   1695
      End
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   795
      TabIndex        =   4
      Top             =   585
      Width           =   7410
      BackColor       =   11517387
      Caption         =   "  Process                                    Customer Code              Order No                      Date"
      Size            =   "13070;423"
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
Attribute VB_Name = "frmVendIssItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyCol As New Collection
Dim MyCol1 As New Collection
Dim IsItemsFirst As Boolean
Dim IssIDs As New Collection
Private Sub cmbCust_matched()
    On Error GoTo Err
    'cmbCountry.ClearVals
    'cmbCountry.AddVals con, "Country", "ForeignCustomers", "Country", " Where CustCode='" & cmbCust.ID & "'"
    'Call FillCountries(cmbCust.Text, cmbCountry)
    cmbOrders.ClearVals
    cmbOrders.AddItem "<All Orders>", "0"
    If cmbCust.ID = "0" Then
        cmbOrders.AddVals con, "OrderNo", "FCustomerOrders", "OrderNo"
    Else
        cmbOrders.AddVals con, "OrderNo", "FCustomerOrders", "OrderNo", " Where CustCode='" & cmbCust.ID & "'"
    End If
    'cmbCustName.ID = cmbCust.ID
    LV.ListItems.Clear
    
    Exit Sub
Err:
    MsgBox Err.Description
End Sub

Private Sub cmbCust_UnMatched()
    LV.ListItems.Clear
    cmbCountry.ClearVals
End Sub

Private Sub cmbOrders_matched()
    
    On Error GoTo Err
    LV.ListItems.Clear
    
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    
    For i = 1 To MyCol.Count
        MyCol.Remove 1
    Next i
    
    With cmd
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_IssueItems"
        
        If rs.State = 1 Then rs.Close
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        
        .Parameters("@ProcessID").value = cmbProcess.ID
        If cmbOrders.ID = "0" Then
            If cmbCust.ID <> "0" Then
                .Parameters("@CustCode").value = cmbCust.ID
            End If
        Else
            .Parameters("@OrderNo").value = cmbOrders.ID
        End If
            
        Set rs = .Execute
    End With
    
    Dim ITM As ListItem
    With rs
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.Add(, ![OrderNo] & "|" & ![CompItemCode], LV.ListItems.Count + 1)
            'ITM.Tag = ![EntryID]
            ITM.ListSubItems.Add , , ![CustCode] & ""
            ITM.ListSubItems.Add , , ![OrderNo] & ""
            
            ITM.ListSubItems.Add , , ![CompItemCode] & ""
            ITM.ListSubItems.Add , , ![ItemDescription] & ""
            ITM.ListSubItems.Add(, , Val(![IssQty] & "")).Tag = Val(![IssQty] & "")
            
            Dim Obj As New IssueItems
            
            Obj.LVIndex = LV.ListItems.Count + 1
                
            Obj.OrderQty = Val(![Qty] & "")
            Obj.FreeStock = Val(![FreeStock] & "")
            Obj.IssueableQty = Val(![IssQty] & "")
            
            MyCol.Add Obj, ![OrderNo] & "|" & ![CompItemCode]
            Set Obj = Nothing
            
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Call AddItems
    
    Exit Sub
Err:
    MsgBox Err.Description
End Sub


Private Sub AddItems()
    
    On Error GoTo Err
    Exit Sub
    Dim OrCond As String
    'Create In List
    
    'For i = 1 To MyCol.Count
    '    OrCond = OrCond & "'" & MyCol(i).CompItemCode & "',"
    'Next i
    
    If Len(InList) > 0 Then
        InList = Mid$(InList, 1, Len(InList) - 1)
    End If
    'InList = InList & ")"
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select * From VVendAssItems", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            For i = 1 To MyCol.Count
                If MyCol(i).CompItemCode = ![ItemID] Then
                    cmbItems.AddItem ![ItemID] 'MyCol(i).ItemCode
                    cmbItems.List(cmbItems.ListCount - 1, 1) = MyCol(i).ItemCode
                    cmbItems.List(cmbItems.ListCount - 1, 2) = MyCol(i).lQty
                    cmbItems.List(cmbItems.ListCount - 1, 3) = ![Rate]
                    cmbItemDesc.AddItem MyCol(i).ItemDescription
                    Exit For
                End If
            Next i
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Exit Sub
Err:
    MsgBox Err.Description

End Sub

Private Sub cmbProcess_matched()
    'Call cmbCust.ClearVals
    Call cmbOrders.ClearVals
    Call cmbCust_matched
    'Call cmbVend.AddVals(con, "VenderName + ' (' + NickName + ')'", "VVenders", "EntryID", " Where PhaseID=" & cmbProcess.ID)
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    If Saved Then
        Unload Me
        Me.Show
    End If
    
End Sub

Private Sub cmdSaveClose_Click()
    
    If Saved Then
        Unload Me
    End If
End Sub

Private Sub cmdUnAssign_Click()

    On Error GoTo Err
    If LV.ListItems.Count = 0 Then Exit Sub
    Dim iAffect As Integer
    
    con.Execute "Delete From VendAssItems Where EntryID=" & Val(LV.SelectedItem.key), iAffect
    If iAffect > 0 Then
        'Call cmbNick_Click
        cmdUnAssign.Visible = False
    End If
    Exit Sub
Err:
    MsgBox Err.Description
    
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Exit Sub 'For Now
    LV.Sorted = True
    If LV.SortKey = ColumnHeader.Index - 1 Then
        If LV.SortOrder = lvwAscending Then
            LV.SortOrder = lvwDescending
            ColumnHeader.Icon = "Up"
        Else
            LV.SortOrder = lvwAscending
            ColumnHeader.Icon = "Down"
        End If
    Else
        LV.ColumnHeaderIcons = Nothing
        Set LV.ColumnHeaderIcons = ImageList2
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        ColumnHeader.Icon = "Down"
    End If

    On Error Resume Next
    LV.SelectedItem.EnsureVisible

End Sub


Private Sub LVIssued_DblClick()

    If LVIssued.ListItems.Count = 0 Then Exit Sub
    'If cmbCust.MatchFound = False Then Exit Sub

    
    With txtRate1
        .Move LVIssued.Left + LVIssued.ColumnHeaders(8).Left, LVIssued.Top + LVIssued.SelectedItem.Top, LVIssued.ColumnHeaders(8).Width
        .Text = LVIssued.SelectedItem.SubItems(7) 'LVMaker.SelectedItem.SubItems(3)
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
End Sub

Private Sub LVIssued_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo Err
    
    If KeyCode = 46 Then
        MyCol(LVIssued.SelectedItem.Tag).IssueableQty = MyCol(LVIssued.SelectedItem.Tag).IssueableQty + Val(LVIssued.SelectedItem.SubItems(7))
        LV.ListItems(LVIssued.SelectedItem.Tag).SubItems(5) = MyCol(LVIssued.SelectedItem.Tag).IssueableQty
        LVIssued.ListItems.Remove LVIssued.SelectedItem.Index
        Call LV_ItemClick(LV.SelectedItem)
    End If
    
    Exit Sub
Err:
    MsgBox Err.Description
    
End Sub

Private Sub LvMaker_DblClick()

    If LVMaker.ListItems.Count = 0 Then Exit Sub
    'If cmbCust.MatchFound = False Then Exit Sub

    
    With txtRate
        .Move LVMaker.Left + LVMaker.ColumnHeaders(4).Left, LVMaker.Top + LVMaker.SelectedItem.Top, LVMaker.ColumnHeaders(4).Width
        '.Text = MyCol(LV.SelectedItem.SubItems(2) & "|" & LV.SelectedItem.SubItems(3)).IssueableQty 'LVMaker.SelectedItem.Tag 'LVMaker.SelectedItem.SubItems(3)
        .Text = MyCol(LV.SelectedItem.key).IssueableQty 'LVMaker.SelectedItem.Tag 'LVMaker.SelectedItem.SubItems(3)
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With

    

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    
    On Error GoTo Err
    
    cmbCust.ListHeight = 2600
    'cmbCustName.ListHeight = 2600
    cmbOrders.ListHeight = 2600
    cmbProcess.ListHeight = 2600
    cmbProcess.AddVals con, "Description", "Processes", "ProcessID", " Order By SNo"
    
    DTPicker1.value = ServerDate
    
    cmbCust.AddItem "<All Customers>", "0"
    cmbCust.AddVals con, "CustCode", "ForeignCustomers", "CustCode", " Order By CustCode"
    'cmbCustName.AddVals con, "Name", "ForeignCustomers", "CustCode", " Order By CustCode"
    
    SSTab1.Tab = 0
    
    Exit Sub
Err:
    MsgBox Err.Description
    
End Sub
Public Sub RefreshLV()
    
    On Error GoTo Err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "Select *  from FCustomerCatalog Where CustCode='" & cmbCust.ID & "' and Country='" & cmbCountry.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.Add(, ![CustCode] & ![Country] & ![ItemID], ![ItemID] & "")
            ITM.SubItems(1) = ![Description] & ""
            ITM.SubItems(2) = ![Unit] & ""
            ITM.SubItems(3) = ![FOB] & ""
            ITM.SubItems(4) = ![CIFSea] & ""
            ITM.SubItems(5) = ![CIFAir] & ""
            ITM.SubItems(6) = ![CnFSea] & ""
            ITM.SubItems(7) = ![CnFAir] & ""
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Exit Sub
Err:
    MsgBox Err.Description
End Sub
Private Sub Form_Resize()

    On Error Resume Next
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
    
    Dim TitleBottom As Long

    TitleBottom = DTPicker1.Top + DTPicker1.Height + 50 '+ fTitle.Height
 
    'LV.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + Frame6.Height + 100)
    'LV.Move 50, TitleBottom ', Me.ScaleWidth - 100 , Me.ScaleHeight - (TitleBottom + Frame6.Height + 100)
    LV.Move Label2.Left, TitleBottom
    SSTab1.Move Label2.Left, LV.Top + LV.Height + 50
    
    'Picture1.Move LV.Left + LV.Width + 40, TitleBottom - 50
    
    'Frame6.Move Me.ScaleWidth - Frame6.Width - 50, Me.ScaleHeight - Frame6.Height - 50
    
 
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)

    On Error GoTo Err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    
    'With MyCol(Item.Index)
    '    txtOrdQty = .OrderQty
    '    txtFree = .FreeStock
    '    txtNet = .OrderQty + .FreeStock '.IssueableQty
    '    txtIss = .IssQty
    '    txtIssBal = .IssueableQty
    '    txtRcv = .RcvQty
    '    txtRcvBal = .RcvBal
    'End With
    
    Dim VendIDList As String
    VendIDList = ""
    For i = 1 To LVIssued.ListItems.Count
        With LVIssued.ListItems(i)
            If LV.SelectedItem.key = .Tag Then
                VendIDList = VendIDList & Val(.ListSubItems(1).key) & ","
            End If
        End With
    Next i
    
    If VendIDList <> "" Then
        VendIDList = "(" & Left(VendIDList, Len(VendIDList) - 1) & ")"
    Else
        VendIDList = "(0)"
    End If
    
    With rs
        .Open "Select VendID,Rate,VenderName From VVendAssItems Where ItemID='" & LV.SelectedItem.SubItems(3) & "' and PhaseID=" & cmbProcess.ID & " and VendID Not In" & VendIDList, con, adOpenForwardOnly, adLockReadOnly
        LVMaker.ListItems.Clear
        Do Until .EOF
            'Set ITM = LVMaker.ListItems.Add(, ![VendID] & "'", LVMaker.ListItems.Count + 1)
            Set ITM = LVMaker.ListItems.Add(, , LVMaker.ListItems.Count + 1)
            ITM.Tag = Val(MyCol(Item.Index).IssueableQty & "")
            ITM.ListSubItems.Add , ![VendID] & "'", .Fields(2) & ""
            ITM.ListSubItems.Add , , Val(![Rate] & "")
            ITM.ListSubItems.Add , , 0
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
Err:
    MsgBox Err.Description
    
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Err
    
    If KeyAscii = 13 Then
        'LVMaker.SelectedItem.SubItems(3) = Val(txtRate)
        
        Dim IssID As String
        For i = 1 To LVIssued.ListItems.Count
            If Val(LVMaker.SelectedItem.ListSubItems(1).key) = Val(LVIssued.ListItems(i).ListSubItems(2).key) Then
                IssID = LVIssued.ListItems(i).ListSubItems(2)
                GoTo GotID:
            End If
        Next i
        
        IssID = Val(con.Execute("Select Max(right(RecieptID,3)) From VendIssued Where Convert(varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "'").Fields(0).value & "")
        
        IssID = "ISS-" & Val(LVMaker.SelectedItem.ListSubItems(1).key) & "-" & _
         Format(ServerDate, "ddMMyy") & "-" & Format(Val(IssID) + 1, "000")
GotID:
        
        Dim ITM As ListItem
        With ITM
            'Set ITM = LVIssued.ListItems.Add(, Val(LV.SelectedItem.Key) & Val(LVMaker.SelectedItem.Key) & "'", LVIssued.ListItems.Count + 1)
            'Set ITM = LVIssued.ListItems.Add(, LV.SelectedItem.Key & "|" & LVIssued.ListItems.Count + 1, LVIssued.ListItems.Count + 1)
            Set ITM = LVIssued.ListItems.Add(, , LVIssued.ListItems.Count + 1)
            ITM.Tag = LV.SelectedItem.key
            ITM.ListSubItems.Add , LVMaker.SelectedItem.ListSubItems(1).key, LVMaker.SelectedItem.SubItems(1)
            ITM.ListSubItems.Add , , IssID & ""
            ITM.ListSubItems.Add , , LV.SelectedItem.SubItems(2)
            ITM.ListSubItems.Add , , LV.SelectedItem.SubItems(3)
            ITM.ListSubItems.Add , , LV.SelectedItem.SubItems(4)
            ITM.ListSubItems.Add , , LVMaker.SelectedItem.ListSubItems(2)
            ITM.ListSubItems.Add , , Val(txtRate)
        End With
        
        MyCol(LV.SelectedItem.Index).IssueableQty = MyCol(LV.SelectedItem.Index).IssueableQty - Val(txtRate)
        LV.SelectedItem.SubItems(5) = Val(LV.SelectedItem.SubItems(5)) - Val(txtRate)
        txtRate.Visible = False
        Call LV_ItemClick(LV.SelectedItem)
        'End If
    ElseIf KeyAscii = 27 Then
        txtRate.Visible = False
    End If
    Call OnlyNums(KeyAscii)
    
    Exit Sub
Err:
    MsgBox Err.Description
    
End Sub
Private Function Saved() As Boolean

    On Error GoTo Err
    Dim SaveIt As Boolean
    
    Dim VendID As New Collection
    Dim VIssID As New Collection
    
    Dim TempIndex As Integer
    TempIndex = 1
    
    Dim lRefID As Long
    Call StartTrans(con)
    
    For i = 1 To LVIssued.ListItems.Count
        With LVIssued.ListItems(i)
            
            
            SaveIt = True
            For j = 1 To VendID.Count
                If Val(LVIssued.ListItems(i).ListSubItems(1).key) = VendID(j) Then
                    'It Means That It's Been Saved Already.
                    SaveIt = False
                    Exit For
                End If
            Next j
        
        
            If SaveIt Then
                con.Execute "Insert Into VendIssued(DT,VendID,RecieptID) Values('" & _
                 DTPicker1 & "'," & Val(LVIssued.ListItems(i).ListSubItems(1).key) & ",'" & .SubItems(2) & "')"
                 
                lRefID = Val(con.Execute("Select Max(EntryID) From VendIssued").Fields(0).value & "")
        
                VendID.Add Val(LVIssued.ListItems(i).ListSubItems(1).key)
                VIssID.Add lRefID
            Else
                lRefID = VIssID(j)
            End If
        
            '1. Insert Into VendIssuedItems
            con.Execute "Insert Into VendIssdDetail(RefID,OrderNo,ItemCode" & _
             ",Rate,ProcessID,IssQty) Values('" & lRefID & _
             "','" & .SubItems(3) & "','" & .SubItems(4) & "'," & _
             Val(.SubItems(6)) & "," & cmbProcess.ID & "," & _
             Val(.SubItems(7)) & ")"
        End With
        
    Next i
                       
    
    '2. Now If Issued After Received Then
    'Update The IssQty In VendRecievedItems....
    '-----------------------------------------------------------------
    For i = 1 To LV.ListItems.Count
        Dim lQty As Long
        lQty = Val(LV.ListItems(i).ListSubItems(5).Tag) - Val(LV.ListItems(i).ListSubItems(5))
        If Val(LV.ListItems(i).ListSubItems(5)) <> Val(LV.ListItems(i).ListSubItems(5).Tag) Then
            Dim cmd As New ADODB.Command
            With cmd
                .ActiveConnection = con
                .CommandType = adCmdStoredProc
                .CommandText = "SP_UpdateForIAR"
                .Parameters("@OrderNo").value = LV.ListItems(i).SubItems(2)
                .Parameters("@ItemCode").value = LV.ListItems(i).SubItems(3)
                .Parameters("@ProcessID").value = cmbProcess.ID
                .Parameters("@Qty").value = lQty
                .Execute
            End With
        End If
    Next
    '-----------------------------------------------------------------
    
    con.CommitTrans
    
    Saved = True
    Exit Function
Err:
    MsgBox Err.Description
    
End Function

Private Sub txtRate1_KeyPress(KeyAscii As Integer)

    On Error GoTo Err
    
    If KeyAscii = 13 Then
        'LVMaker.SelectedItem.SubItems(3) = Val(txtRate)
        Dim ITM As ListItem
        
        MyCol(LVIssued.SelectedItem.Tag).IssueableQty = MyCol(LVIssued.SelectedItem.Tag).IssueableQty + Val(LVIssued.SelectedItem.SubItems(7)) - Val(txtRate1)
        LV.ListItems(LVIssued.SelectedItem.Tag).SubItems(5) = MyCol(LVIssued.SelectedItem.Tag).IssueableQty
        LVIssued.SelectedItem.SubItems(7) = Val(txtRate1)
        txtRate1.Visible = False
        Call LV_ItemClick(LV.SelectedItem)
        'End If
        If Val(txtRate1) = 0 Then
            Call LVIssued_KeyDown(46, 0)
        End If
        
    ElseIf KeyAscii = 27 Then
        txtRate1.Visible = False
    End If
    Call OnlyNums(KeyAscii)
    
    Exit Sub
Err:
    MsgBox Err.Description
    
End Sub

Private Sub txtRate1_LostFocus()
    txtRate1.Visible = False
End Sub
Private Sub txtRate_LostFocus()
    txtRate.Visible = False
End Sub
