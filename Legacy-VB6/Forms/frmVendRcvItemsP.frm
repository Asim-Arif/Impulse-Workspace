VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmVendRcvItemsP 
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11040
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
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   11040
   WindowState     =   2  'Maximized
   Begin VB.Frame FALL 
      Height          =   9090
      Left            =   105
      TabIndex        =   4
      Top             =   45
      Width           =   10635
      Begin ComboList.Usercontrol1 cmbVend 
         Height          =   285
         Left            =   3390
         TabIndex        =   11
         Top             =   960
         Width           =   3300
         _ExtentX        =   5821
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
         Left            =   150
         TabIndex        =   12
         Top             =   960
         Width           =   3195
         _ExtentX        =   5636
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
      Begin VB.Frame Frame1 
         Height          =   645
         Left            =   0
         TabIndex        =   7
         Top             =   75
         Width           =   10635
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Receive Items"
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
            Left            =   60
            TabIndex        =   8
            Top             =   165
            Width           =   10605
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Receive Items"
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
            Index           =   1
            Left            =   75
            TabIndex        =   9
            Top             =   180
            Width           =   10590
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   6690
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   960
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   59834371
         CurrentDate     =   37250
      End
      Begin VB.Frame FraCash 
         BorderStyle     =   0  'None
         Height          =   7785
         Left            =   120
         TabIndex        =   5
         Top             =   1275
         Width           =   10395
         Begin VB.PictureBox PicEdit 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3300
            ScaleHeight     =   285
            ScaleWidth      =   1695
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   765
            Visible         =   0   'False
            Width           =   1695
            Begin VB.TextBox txtWastage 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   840
               TabIndex        =   19
               Top             =   0
               Width           =   850
            End
            Begin VB.TextBox txtRepair 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   0
               TabIndex        =   18
               Top             =   0
               Width           =   850
            End
         End
         Begin VB.PictureBox Pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            CausesValidation=   0   'False
            ClipControls    =   0   'False
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1350
            ScaleHeight     =   255
            ScaleWidth      =   945
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   1110
            Visible         =   0   'False
            Width           =   945
            Begin VB.TextBox txtQty 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   255
               Left            =   0
               TabIndex        =   16
               Top             =   0
               Width           =   950
            End
         End
         Begin VB.Frame Frame6 
            Height          =   660
            Left            =   5220
            TabIndex        =   10
            Top             =   7050
            Width           =   5175
            Begin MSForms.CommandButton cmdSClose 
               Height          =   360
               Left            =   1770
               TabIndex        =   2
               Top             =   195
               Width           =   1620
               Caption         =   "Save & Print"
               PicturePosition =   327683
               Size            =   "2857;635"
               Accelerator     =   83
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
               Left            =   3435
               TabIndex        =   3
               Top             =   195
               Width           =   1620
               Caption         =   "Cancel            "
               PicturePosition =   327683
               Size            =   "2857;635"
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
               Left            =   75
               TabIndex        =   1
               Top             =   195
               Width           =   1620
               Caption         =   "Save & New   "
               PicturePosition =   327683
               Size            =   "2857;635"
               Accelerator     =   78
               FontName        =   "Tahoma"
               FontEffects     =   1073741825
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
               ParagraphAlign  =   3
               FontWeight      =   700
            End
         End
         Begin MSComctlLib.ListView LV 
            Height          =   7050
            Left            =   45
            TabIndex        =   14
            Top             =   0
            Width           =   10305
            _ExtentX        =   18177
            _ExtentY        =   12435
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   8388608
            BackColor       =   15199215
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "SNo"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Key             =   "ItemCode"
               Text            =   "Item ID"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Key             =   "ItemDescription"
               Text            =   "Item Description"
               Object.Width           =   5997
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Key             =   "Quantity"
               Text            =   "Qty"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Repair"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Wastage"
               Object.Width           =   1499
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Key             =   "Amt"
               Text            =   "Rcvd. Qty"
               Object.Width           =   1676
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Key             =   "RecieptID"
               Text            =   "Receive ID"
               Object.Width           =   2469
            EndProperty
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
                  Picture         =   "frmVendRcvItemsP.frx":0000
                  Key             =   "Down"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmVendRcvItemsP.frx":01BC
                  Key             =   "Up"
               EndProperty
            EndProperty
         End
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   6
         Top             =   735
         Width           =   8055
         ForeColor       =   4194304
         BackColor       =   12504280
         Caption         =   $"frmVendRcvItemsP.frx":0378
         Size            =   "14208;397"
         BorderColor     =   10860487
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox txtRecNo 
         Height          =   285
         Left            =   8220
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   2235
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         Size            =   "3942;503"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmVendRcvItemsP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dWastage As Double
Dim dRcvdQty As Double
Dim Col_OrdQty As New Collection
Dim MyCol1 As New Collection
Dim IsItemsFirst As Boolean
Dim InitLotProcess As Boolean
Dim lLotNo As Long
Dim GotRcvID As Long
Dim RcvingLot As Boolean
Dim UserReqAuth As Integer
Dim strOrder As String

Private Sub cmbProcess_matched()

    On Error GoTo err
    
    Dim cmd As New ADODB.Command
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "SP_ReqAuthForRcv"
        Set .ActiveConnection = con
        .Parameters("@UserID").value = UserID
        .Parameters("@ProcessID").value = cmbProcess.ID
        .Execute
        UserReqAuth = .Parameters("@ReqAuth").value
    End With
    
    If RcvLot(cmbProcess.ID) Then
        RcvingLot = True
    Else
        RcvingLot = False
    End If
    
    LV.ListItems.Clear
    GotRcvID = 0
    Call cmbVend.ClearVals
    
    If cmbProcess.ID = 12 Then
        Call cmbVend.AddVals(con, "VenderName", "VVenders", "VendID", " Where PhaseID=12 OR VendID=129")
    Else
        If HicoVisible Then
            Call cmbVend.AddVals(con, "VenderName", "VVenders", "VendID", " Where PhaseID=" & cmbProcess.ID & " Or PhaseID=-1")
        Else
            Call cmbVend.AddVals(con, "VenderName", "VVenders", "VendID", " Where PhaseID=" & cmbProcess.ID)
        End If
    End If
    
    Dim lSNo1 As Long, lSNo2 As Long
    lSNo1 = Val(con.Execute("Select SNo From VInitialLotProcess").Fields(0).value & "")
    lSNo2 = Val(con.Execute("Select SNo From Processes Where ProcessID=" & cmbProcess.ID).Fields(0).value & "")
    
    If lSNo2 >= lSNo1 Then
        'If LV.ColumnHeaders.Count = 5 Then
        '    LV.ColumnHeaders.Add(, , "Lot #", "2000").Tag = "2000"
        'End If
        cmbProcess.Tag = "Lot No"
    Else
        'If LV.ColumnHeaders.Count = 6 Then
        '    LV.ColumnHeaders.Remove 6
        'End If
        cmbProcess.Tag = ""
    End If
    
    cmbVend.SetFocus
    SendKeys "a"
    SendKeys Chr(8)
    
    
    cmbVend.ID = 129
    Exit Sub
    
err:
    MsgBox err.Description
End Sub

Private Sub cmbVend_GotFocus()
    Me.KeyPreview = True
End Sub

Private Sub cmbVend_matched()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim Sql As String
    
    InitLotProcess = False
    'If LV.ColumnHeaders.Count = 7 Then LV.ColumnHeaders.Remove "OrderNo"
    
    
    With rs
        If cmbProcess.Tag = "Lot No" Then
            'SQL = "Select LotNo,RecieptID,ItemCode,ItemDescription,Sum(IssQty-RcvdQty) as Quantity From VVendIssItems1 Where IssQty-RcvdQty>0 and VendID=" & cmbVend.ID & " and ReqAuth=0 Group By LotNo,RecieptID,ItemCode,ItemDescription"
            Sql = "Select * From VVendRcvLots Where VendID=" & cmbVend.ID & " And ProcessID=" & cmbProcess.ID & " Order By Cast(SubString(RecieptID,4,6) As Int),Cast(SubString(RecieptID,4,Len(RecieptID)-3) As Int)"
        End If
        
        .Open Sql, con, adOpenForwardOnly, adLockReadOnly
        
        LV.ListItems.Clear
        GotRcvID = 0
        
        Do Until .EOF
            Set ITM = LV.ListItems.Add(, , LV.ListItems.Count + 1)
            
            'If cmbProcess.Tag = "Lot No" Then itm.Tag = ![DeliveryDT] & ""
            
            'If InitLotProcess Then itm.ListSubItems.Add , , ![OrderNo] & ""
            ITM.ListSubItems.Add , , ![ItemCode] & ""
            ITM.ListSubItems.Add , , ![ItemDescription] & ""
            ITM.ListSubItems.Add , , ![Quantity] & ""
            
            ITM.ListSubItems.Add
            
            ITM.ListSubItems.Add
            ITM.ListSubItems.Add
            
            If cmbProcess.Tag = "Lot No" Then
                'itm.ListSubItems.Add , , ![LotNo] & ""
                ITM.ListSubItems(4).Tag = ![LotNo] & ""
            Else
                ITM.ListSubItems(4).Tag = ""
            End If
            
            If RcvingLot Then
                ITM.ListSubItems.Add , , ![RecieptID] & ""
            Else
                ITM.ListSubItems.Add
            End If
            
            
            .MoveNext
        Loop
        .Close
        
    End With
    
    Set rs = Nothing
    
    Dim myDT As String
    If cmbProcess.Tag = "Lot No" Then
        If LV.ListItems.Count > 0 Then
            myDT = LV.ListItems(1).Tag
            For i = 1 To LV.ListItems.Count
                Set ITM = LV.ListItems(i)
                If myDT = ITM.Tag Then
                    'Set The ForeColor Of all The ListSubItems To Red.
                    ITM.ForeColor = vbRed
                    For j = 1 To ITM.ListSubItems.Count
                        ITM.ListSubItems(j).ForeColor = vbRed
                    Next j
                End If
            Next i
        End If
    End If
    
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbVend_Validate(Cancel As Boolean)
    Me.KeyPreview = False
End Sub

Private Sub cmdSClear_Click()

    If Saved Then
        FormProcessID = cmbProcess.ID
        Unload Me
        'Me.Show
        'Me.SetFocus
        Call ShowRcvVForm(FormProcessID)
    End If
    'Call ClearFields
    
End Sub

Private Sub cmdSClose_Click()

    'If Not AskForAutoApply Then Exit Sub

    If Saved Then
        'Call PrintIssRcvRpts("Original", False)
        Unload Me
        con.Execute "Delete From PrintSession"
    End If

End Sub

Private Sub cmdCancel_Click()
    Unload Me
    con.Execute "Delete From PrintSession"
End Sub


Private Sub Form_Activate()
    'Check For User Access
    If PermissionDenied Then
        Me.Hide
        MsgBox "You Donot Have Permissions To Use This Option. Please Contact Your Administrator", vbInformation
        Unload Me
        Exit Sub
    End If


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    
    On Error GoTo err

    'Me.KeyPreview = True
    DTPicker1.value = ServerDate
    
    'Call DTPicker1_Click
    'Call FillAccCmbs
    cmbProcess.Enabled = True
    cmbProcess.ListHeight = 2600
    cmbVend.ListHeight = 2600
    Call cmbProcess.AddVals(con, "Description", "Processes", "ProcessID", "Order By SNo")
    
    lLotNo = 0
    Call SaveLV(LV)
    
    GotRcvID = 0
    
    If UserID = 1 Then DTPicker1.Enabled = True
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
Public Sub LockProc(ProcID As Long)
    Call Form_Resize
    cmbProcess.ID = ProcID
    cmbProcess.Enabled = False
    Me.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Set Col_OrdQty = Nothing
    Set frmVendRcvItems = Nothing
End Sub

Private Sub Form_Resize()
   FALL.Move (Me.ScaleWidth - FALL.Width) / 2, (Me.ScaleHeight - FALL.Height) / 2
End Sub



Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    Exit Sub
    If ColumnHeader.Key = "" Then Exit Sub
    strOrder = " Order By " & ColumnHeader.Key
    Call cmbVend_matched
'    If LV.SortKey + 1 = ColumnHeader.Index Then
'        If LV.SortOrder = lvwAscending Then
'            LV.SortOrder = lvwDescending
'            ColumnHeader.Icon = "Down"
'        Else
'            LV.SortOrder = lvwAscending
'            ColumnHeader.Icon = "Up"
'        End If
'    Else
'        LV.ColumnHeaderIcons = Nothing
'        Set LV.ColumnHeaderIcons = ImageList2
'        LV.SortKey = ColumnHeader.Index - 1
'        LV.SortOrder = lvwAscending
'        ColumnHeader.Icon = "Down"
'    End If

End Sub

Private Sub LV_DblClick()
    
    'If RcvingLot Then Exit Sub 'Then CheckBoxes Will Work
    If LV.ListItems.Count = 0 Then Exit Sub
    If LV.SelectedItem.Checked = False Then Exit Sub
    Call ResetLV(LV)
    
    With PicEdit
        .Move LV.Left + 10 + LV.ColumnHeaders(5).Left, LV.SelectedItem.Top + LV.Top
        .Visible = True
    End With
    
    With txtRepair
        .Text = LV.SelectedItem.SubItems(4)
        txtWastage.Text = LV.SelectedItem.SubItems(5)
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
    
End Sub


Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    If Item.Checked Then
        Item.SubItems(4) = 0
        Item.SubItems(5) = 0
        Item.SubItems(6) = Item.SubItems(3)
        
        Call ResetLV(LV)
        
        'LV.ListItems(cmbItem.ID & "'").EnsureVisible
        'LV.FindItem "     »  " & cmbitem.ID
        'LV.ListItems(Item).Selected = True
        Item.Selected = True
        With PicEdit
            .Move LV.Left + 10 + LV.ColumnHeaders(5).Left, Item.Top + LV.Top
            .Visible = True
        End With
        
        With txtRepair
            .Text = Item.SubItems(4)
            txtWastage.Text = Item.SubItems(5)
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
        
        
    Else
        Item.SubItems(4) = 0
        Item.SubItems(5) = 0
        Item.SubItems(6) = 0
    End If
   
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'If PicEdit.Visible Then PicEdit.Visible = False
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
   ' Exit Sub
   ' If LV.ListItems.Count = 0 Then Exit Sub
    'If Val(LV.SelectedItem.SubItems(6)) = 0 Then Exit Sub
    'Dim iCol As Integer
    'iCol = GetColNo(LV, X)
    
    'If iCol = 8 Then
    '    Load frmVendRcvItems1
    '    LV.SelectedItem.ListSubItems(7).Text = frmVendRcvItems1.FillList(Col_OrdQty, LV.SelectedItem.SubItems(1))
    '    LV.SelectedItem.ListSubItems(7).Tag = LV.SelectedItem.ListSubItems(7).Text
    'End If
    
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    If KeyAscii = 13 Then
        
        If Val(txtQty) = 0 Then
            If InitLotProcess Then
                LV.SelectedItem.SubItems(5) = Val(txtQty)
            Else
                LV.SelectedItem.SubItems(4) = Val(txtQty)
            End If
            LV.SelectedItem.SubItems(LV.ColumnHeaders.Count - 1) = ""
            'Reset The IssID's
            Dim TempID As Long
            TempID = 0
            
            For i = LV.SelectedItem.Index + 1 To LV.ListItems.Count
                TempID = Val(Right(LV.ListItems(i).SubItems(LV.ColumnHeaders.Count - 1), 3)) - 1
                If TempID > 0 Then
                    LV.ListItems(i).SubItems(LV.ColumnHeaders.Count - 1) = "RCV-" & _
                     Format(ServerDate, "ddMMyy") & "-" & Format(TempID, "000")
                End If
            Next i
            If GotRcvID > 0 Then GotRcvID = GotRcvID - 1
            
            
            
            Pic.Visible = False
            Exit Sub
        End If
        
        If InitLotProcess Then
            If Val(txtQty) > Val(LV.SelectedItem.SubItems(4)) Then
                MsgBox "Invalid Quantities!", vbInformation
                Exit Sub
            End If
        Else
            If Val(txtQty) > Val(LV.SelectedItem.SubItems(3)) Then
                MsgBox "Invalid Quantities!", vbInformation
                Exit Sub
            End If
        End If
        
        If Not txtQty.Locked Then
            If cmbProcess.Tag = "Lot No" Then
                If lLotNo = 0 Then
                    lLotNo = Val(con.Execute("Select Max(Cast(LotNo As int)) From VendRcvdDetail").Fields(0).value & "") + 1
                Else
                    lLotNo = lLotNo + 1
                End If
                
                If InitLotProcess Then
                    LV.SelectedItem.ListSubItems(4).Tag = lLotNo
                Else
                    LV.SelectedItem.ListSubItems(4).Tag = lLotNo
                End If
            End If
        End If
        If InitLotProcess Then
            LV.SelectedItem.SubItems(5) = Val(txtQty)
        Else
            LV.SelectedItem.SubItems(4) = Val(txtQty)
        End If
        
        '-----------------------
        
        'Dim RcvdID As String
        
            If cmbProcess.Tag = "Lot No" Then
                If GotRcvID = 0 Then
                    GotRcvID = Val(con.Execute("Select Max(right(RecieptID,3)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "' and UPPER(Left(RecieptID,2))='HS'").Fields(0).value & "") + 1
                Else
                    GotRcvID = GotRcvID + 1
                End If
                LV.SelectedItem.SubItems(LV.ColumnHeaders.Count - 1) = "HS-" & _
                 Format(DTPicker1, "ddMMyy") & "-" & Format(GotRcvID, "000")
            Else
                If GotRcvID = 0 Then
                    GotRcvID = Val(con.Execute("Select Max(right(RecieptID,3)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "' and UPPER(Left(RecieptID,3))='RCV'").Fields(0).value & "") + 1
                Else
                    GotRcvID = GotRcvID + 1
                End If
                 LV.SelectedItem.SubItems(LV.ColumnHeaders.Count - 1) = "RCV-" & _
                 Format(DTPicker1, "ddMMyy") & "-" & Format(GotRcvID, "000")
            End If
            
        '-----------------------
        
        'Call AdjustQty
        Pic.Visible = False
    ElseIf KeyAscii = 27 Then
        Pic.Visible = False
    End If
    
    Call OnlyNums(KeyAscii, False)
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub AdjustQty()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim lQty As Long, TotQty As Long, WasteQty As Long
    Dim TempQty As Long, TempWasteQty As Long
    Dim StrDetail As String
    
    TempQty = 0
    TotQty = Val(LV.SelectedItem.SubItems(6))
    WasteQty = Val(LV.SelectedItem.SubItems(5))
    
    With rs
        .Open "Select OrderNo,ItemCode,TotQty From VTotIssdByOrder Where ItemCode='" & LV.SelectedItem.SubItems(1) & "'", con, adOpenForwardOnly
        
        Do Until .EOF
            Dim Obj As New RcvAdjOrderQty
            Obj.OrderNo = ![OrderNo] & ""
            Obj.ItemCode = ![ItemCode] & ""
            Obj.lOrdQty = ![TotQty] & ""
            Obj.lWastage = 0
            Obj.lQty = 0
            
            Col_OrdQty.Add Obj, ![OrderNo] & "|" & ![ItemCode]
            Set Obj = Nothing
            .MoveNext
        Loop
        .MoveFirst
        
        Do While lQty < TotQty
            
            TempQty = TotQty - lQty
            TempWasteQty = 0
            If TempQty > ![TotQty] Then
                TempQty = ![TotQty]
            End If
            
            If WasteQty > 0 Then
                If TempQty > WasteQty Then
                    TempQty = TempQty - WasteQty
                    TempWasteQty = WasteQty
                    WasteQty = 0
                Else
                    TempWasteQty = WasteQty
                    TempQty = 0
                    WasteQty = WasteQty - TempQty
                End If
            End If
            
            
            Col_OrdQty(![OrderNo] & "|" & ![ItemCode]).lQty = TempQty
            Col_OrdQty(![OrderNo] & "|" & ![ItemCode]).lWastage = TempWasteQty
            
            lQty = lQty + TempQty + TempWasteQty
            
            StrDetail = StrDetail & ![OrderNo] & "=" & TempQty + TempWasteQty & ","
            
            .MoveNext
        Loop
        .Close
    End With
    
    StrDetail = Left(StrDetail, Len(StrDetail) - 1)
    LV.SelectedItem.ListSubItems(7) = StrDetail
    LV.SelectedItem.ListSubItems(7).ToolTipText = StrDetail
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Function OLDSaved() As Boolean

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim Col As New Collection
    
    Dim cmd As New ADODB.Command
    
    Dim iRefID As Long
    
    Call StartTrans(con)
    
    con.Execute "Insert Into VendReceived(DT,RecieptNo,VendID) Values('" & DTPicker1 & "','" & txtRecNo & "'," & cmbVend.ID & ")"
    
    iRefID = Val(con.Execute("Select Max(EntryID) From VendReceived").Fields(0).value & "")
    
    'Get The Next Process ID For Current Item...
    For i = 1 To LV.ListItems.Count
        
            If Val(LV.ListItems(i).SubItems(6)) > 0 Then
                Dim iNextProcID As Integer
                'Dim cmd As New ADODB.Command
                With cmd
                    .ActiveConnection = con
                    .CommandType = adCmdStoredProc
                    .CommandText = "SP_GetNextProcID"
                    .Parameters("@ItemCode").value = LV.ListItems(i).SubItems(1)
                    .Parameters("@ProcID").value = cmbProcess.ID
                    .Execute
                    iNextProcID = Val(.Parameters("@NextProcID").value & "")
                End With
                
                For j = 1 To Col_OrdQty.Count
                    With Col_OrdQty(j)
                        If .ItemCode = LV.ListItems(i).SubItems(1) And .lQty + .lWastage > 0 Then
                        
                            con.Execute "Insert Into VendRcvdDetail(RefID,OrderNo,ItemCode,NextProcessID,ProcessID,RcvdQty,Wastage) Values(" & _
                            iRefID & ",'" & .OrderNo & "','" & .ItemCode & "'," & iNextProcID & "," & cmbProcess.ID & "," & .lQty & "," & .lWastage & ")"
                            
                            '2. Now If Rcvd After Issued Then
                            'Update The RcvdQty In VendIssdDetails....
                            '-----------------------------------------------------------------
                            With cmd
                                .ActiveConnection = con
                                .CommandType = adCmdStoredProc
                                .CommandText = "SP_UpdateIssued"
                                .Parameters("@OrderNo").value = Col_OrdQty(j).OrderNo
                                .Parameters("@ItemCode").value = Col_OrdQty(j).ItemCode
                                .Parameters("@ProcessID").value = cmbProcess.ID
                                .Parameters("@Qty").value = Col_OrdQty(j).lQty + Col_OrdQty(j).lWastage
                                .Execute
                            End With
                            '-----------------------------------------------------------------
                                
                                                    
                            'con.Execute "Update VendIssdDetail Set RcvdQty=RcvdQty+" & .lQty + .lWastage & " Where OrderNo='" & .OrderNo & "' and ItemCode='" & .ItemCode & "'"
                        End If
                    End With
                Next j
                
                'For j = 1 To Col.Count
                '    Col.Remove 1
                'Next
                
                
                
                
                '---------------------------------------------------------------------------------------------------------------------------------------
                'Update RcvdQty In The VendIssdDetail Table....
                
                'With rs
                '    .Open "Select EntryID,IssQty-RcvdQty From VendIssdDetail Where ItemCode='" & LV.ListItems(i).SubItems(1) & "' and IssQty-RcvdQty>0 and ProcessID=" & cmbProcess.ID & " Order By EntryID", con, adOpenForwardOnly, adLockReadOnly
                    
                '    Do Until .EOF
                '        Dim Obj As New AdjIssQty
                '        Obj.EntryID = ![EntryID]
                '        Obj.lQty = .Fields(1)
                '        Col.Add Obj
                '        Set Obj = Nothing
                '        .MoveNext
                '    Loop
                '    .Close
                'End With
                
                'Dim TempQty As Long, TotQty As Long
                'Dim TempTot As Long
                'TotQty = Val(LV.ListItems(i).SubItems(6))
                'TempQty = 0
                'TempTotQty = 0
                
                'For j = 1 To Col.Count
                
                '    If TotQty - TempTot >= Col(j).lQty Then
                '        TempQty = Col(j).lQty
                '    Else
                '        TempQty = TotQty - TempTot
                '    End If
                '    TempTot = TempTot + TempQty
                '
                '    con.Execute "Update VendIssdDetail Set RcvdQty=" & TempQty & " Where EntryID=" & Col(j).EntryID
                '
                '    If TempTot >= TotQty Then Exit For
                '
                'Next j
                '---------------------------------------------------------------------------------------------------------------------------------------
            End If
        
    Next i
    
    Set rs = Nothing
    con.CommitTrans
    OLDSaved = True
    Exit Function
err:
    MsgBox err.Description
    
End Function
    
Private Function Saved() As Boolean

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim Col As New Collection
    
    Dim cmd As New ADODB.Command
    Dim Qty As Long
    
    Dim iRefID As Long
    
    Call StartTrans(con)
    
    'con.Execute "Insert Into VendReceived(DT,VendID) Values('" & DTPicker1 & " " & ServerTime & "'," & cmbVend.ID & ")"
    con.Execute "Insert Into VendReceived(VendID,DT,RecieptID,UserID,ProcessID) Values(" & cmbVend.ID & _
     ",'" & DTPicker1 & " " & ServerTime & "',''," & UserID & "," & cmbProcess.ID & ")"
    
    
    iRefID = Val(con.Execute("Select Max(EntryID) From VendReceived").Fields(0).value & "")
    
    'Get The Next Process ID For Current Item...
    Dim MyCol As New RcvAdjOrderQty
    
    For i = 1 To LV.ListItems.Count
            
            
            Qty = Val(LV.ListItems(i).SubItems(6))
            lLotNo = Val(LV.ListItems(i).ListSubItems(4).Tag)
            
            If Qty > 0 Then
                '2. Now If Rcvd After Issued Then
                'Update The RcvdQty In VendIssdDetails....
                '-----------------------------------------------------------------
                
                With cmd
                    .ActiveConnection = con
                    .CommandType = adCmdStoredProc
                
                    .CommandText = "SP_InsertVendRcvdLotsP"
                    
                    .Parameters("@ProcessID").value = cmbProcess.ID
                    .Parameters("@Qty").value = Val(LV.ListItems(i).SubItems(6) & "")
                    .Parameters("@Wastage").value = Val(LV.ListItems(i).SubItems(5) & "")
                    .Parameters("@Repair").value = Val(LV.ListItems(i).SubItems(4) & "")
                    .Parameters("@RefID").value = iRefID
                    .Parameters("@VendID").value = cmbVend.ID
                    .Parameters("@LotNo").value = lLotNo
                    .Parameters("@RecID").value = LV.ListItems(i).SubItems(LV.ColumnHeaders.Count - 1)
                    .Parameters("@ReqAuth").value = UserReqAuth
                    
                    .Execute
                End With
                
                '-----------------------------------------------------------------
            
'                If UserReqAuth = 0 Then
'                    Dim DetailRefID As Long
'                    DetailRefID = Val(con.Execute("Select Max(EntryID) From VendRcvdDetail").Fields(0).value & "")
'                    con.Execute "Insert Into VendRcvAuthDetail(UserID,RcvID,DT) Values(" & _
'                     UserID & "," & DetailRefID & ",'" & DTPicker1 & " " & ServerTime & "')"
'                End If
                con.Execute "Insert Into PrintSession(RecieptNo) Values('" & LV.ListItems(i).SubItems(LV.ColumnHeaders.Count - 1) & "')"
            End If
            
    Next i
    
    Set rs = Nothing
    con.CommitTrans
    Saved = True
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Private Sub txtRepair_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        With txtWastage
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
        End With
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub


Private Sub txtWastage_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        With LV.SelectedItem
            If Val(txtRepair) + Val(txtWastage) > Val(.SubItems(3)) Then
                MsgBox "Invalid Qty.", vbInformation
                Exit Sub
            End If
            .SubItems(4) = Val(txtRepair)
            .SubItems(5) = Val(txtWastage)
            .SubItems(6) = Val(.SubItems(3)) - Val(txtRepair) - Val(txtWastage)
            PicEdit.Visible = False
            
            If LV.SelectedItem.Index < LV.ListItems.Count Then
                LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
                LV.SetFocus
            End If
            
        End With
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub
