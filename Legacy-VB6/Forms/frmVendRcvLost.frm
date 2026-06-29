VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmVendRcvLost 
   ClientHeight    =   9945
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11880
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
   MDIChild        =   -1  'True
   ScaleHeight     =   9945
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame FALL 
      Height          =   9855
      Left            =   105
      TabIndex        =   4
      Top             =   45
      Width           =   11715
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
         TabIndex        =   13
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
            Caption         =   "Receive Lost"
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
            Caption         =   "Receive Lost"
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
         Format          =   44367875
         CurrentDate     =   37250
      End
      Begin VB.Frame FraCash 
         BorderStyle     =   0  'None
         Height          =   8370
         Left            =   120
         TabIndex        =   5
         Top             =   1275
         Width           =   11445
         Begin VB.PictureBox Pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            CausesValidation=   0   'False
            ClipControls    =   0   'False
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   840
            ScaleHeight     =   285
            ScaleWidth      =   2475
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   540
            Visible         =   0   'False
            Width           =   2475
            Begin VB.TextBox txtRate 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1290
               TabIndex        =   17
               Top             =   0
               Width           =   1170
            End
            Begin VB.TextBox txtQty 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   0
               TabIndex        =   16
               Top             =   0
               Width           =   1300
            End
         End
         Begin VB.Frame Frame6 
            Height          =   660
            Left            =   6195
            TabIndex        =   10
            Top             =   7710
            Width           =   5175
            Begin MSForms.CommandButton cmdSClose 
               Height          =   360
               Left            =   1770
               TabIndex        =   2
               Top             =   195
               Width           =   1620
               Caption         =   "Save & Close  "
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
            Height          =   7695
            Left            =   60
            TabIndex        =   12
            Top             =   0
            Width           =   11280
            _ExtentX        =   19897
            _ExtentY        =   13573
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
               Text            =   "SNo"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Item ID"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Item Description"
               Object.Width           =   6350
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Qty"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Key             =   "Amt"
               Text            =   "Rcvd. Lost"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Rate"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Date"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Receive ID"
               Object.Width           =   2646
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
         Caption         =   $"frmVendRcvLost.frx":0000
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
         TabIndex        =   14
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
Attribute VB_Name = "frmVendRcvLost"
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
    
    If HicoVisible Then
        If cmbProcess.ID = 7 Then
            Call cmbVend.AddVals(con, "VenderName", "VVenders", "VendID", " Where (PhaseID=" & cmbProcess.ID & " Or PhaseID=-1 OR PhaseID=0) And VendID In(Select VendID From VVendIssItems Where ProcessID=" & cmbProcess.ID & " Group By VendID Having Sum(RemQty)>0)")
        Else
            Call cmbVend.AddVals(con, "VenderName", "VVenders", "VendID", " Where (PhaseID=" & cmbProcess.ID & " Or PhaseID=-1 OR PhaseID=0) And VendID In(Select VendID From VVendIssItems Where ProcessID=" & cmbProcess.ID & " Group By VendID Having Sum(RemQty)>0)")
        End If
    Else
        If cmbProcess.ID = 7 Then
            Call cmbVend.AddVals(con, "VenderName", "VVenders", "VendID", " Where (PhaseID=" & cmbProcess.ID & " OR PhaseID=0) And VendID In(Select VendID From VVendIssItems Where ProcessID=" & cmbProcess.ID & " Group By VendID Having Sum(RemQty)>0)")
        Else
            Call cmbVend.AddVals(con, "VenderName", "VVenders", "VendID", " Where PhaseID=" & cmbProcess.ID & " And VendID In(Select VendID From VVendIssItems Where ProcessID=" & cmbProcess.ID & " Group By VendID Having Sum(RemQty)>0)")
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
    
    With rs
        '.Open "Select ItemCode,ItemDescription,Sum(IssQty-RcvdQty) as Quantity From VVendIssItems1 Where IssQty-RcvdQty>0 and VendID=" & cmbVend.ID & " Group By ItemCode,ItemDescription", con, adOpenForwardOnly, adLockReadOnly
        .Open "Select ItemCode,ItemDescription,Sum(IssQty-RcvdQty)-dbo.F_AuthQtyForIss(ItemCode,null,null,null," & cmbVend.ID & ") as Quantity From VVendIssItems1 Where IssQty-RcvdQty>0 and VendID=" & cmbVend.ID & " And IsNull(ReqAuth,0)=0 Group By ItemCode,ItemDescription ", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.Add(, , LV.ListItems.Count + 1)
            ITM.ListSubItems.Add , , ![ItemCode] & ""
            ITM.ListSubItems.Add , , ![ItemDescription] & ""
            ITM.ListSubItems.Add , , ![Quantity] & ""
            ITM.ListSubItems.Add
            ITM.ListSubItems.Add
            ITM.ListSubItems.Add
            ITM.ListSubItems.Add
            .MoveNext
        Loop
        .Close
        
    End With
    
    'Dim RcvdID As String
    
    'RcvdID = Val(con.Execute("Select Max(right(RecieptID,3)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "' and UPPER(Left(RecieptID,3))='RCV'").Fields(0).value & "")
    '    RcvdID = "RCV-" & _
         Format(DTPicker1, "ddMMyy") & "-" & Format(Val(RcvdID) + 1, "000")
         
    
     
    'txtRecNo = RcvdID
    Set rs = Nothing
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
        Call ShowLostRcvForm(FormProcessID)
    End If
    'Call ClearFields

End Sub

Private Sub cmdSClose_Click()

    'If Not AskForAutoApply Then Exit Sub

    If Saved Then
        Unload Me
    End If

End Sub

Private Sub cmdCancel_Click()
    Unload Me
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

Private Sub Form_Load()
    
    On Error GoTo err
    If UserID = 1 Then DTPicker1.Enabled = True
    'Me.KeyPreview = True
    DTPicker1.value = ServerDate
    
    cmbProcess.Enabled = True
    cmbProcess.ListHeight = 2600
    cmbVend.ListHeight = 2600
    Call cmbProcess.AddVals(con, "Description", "Processes", "ProcessID", "Order By SNo")
    
    lLotNo = 0
    Call SaveLV(LV)
    
    GotRcvID = 0
    
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

Private Function GetLastRate() As Double
    
    Dim dRetVal As Double
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Top 1 Rate From VenderLedger Where VendID=" & cmbVend.ID & " AND ItemCode='" & LV.SelectedItem.SubItems(1) & "' AND Rate<0 Order By EntryID Desc", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            dRetVal = 0
        Else
            dRetVal = Abs(Val(.Fields(0) & ""))
        End If
        .Close
    End With
    Set rs = Nothing
    GetLastRate = dRetVal
    
End Function
Private Sub LV_DblClick()
    
    
    If RcvingLot Then Exit Sub 'Then CheckBoxes Will Work
    Call ResetLV(LV)
    
    With Pic
        If InitLotProcess Then
            .Move LV.Left + 10 + LV.ColumnHeaders(6).Left, LV.SelectedItem.Top + LV.Top
        Else
            .Move LV.Left + 10 + LV.ColumnHeaders(5).Left, LV.SelectedItem.Top + LV.Top
        End If
        .Visible = True
    End With
    
    txtRate.Text = ""
    
    With txtQty
    
        If Val(LV.SelectedItem.SubItems(4)) = 0 Then
            .Text = Val(LV.SelectedItem.SubItems(3))
            txtRate.Text = GetLastRate
        Else
            .Text = Val(LV.SelectedItem.SubItems(4))
            txtRate.Text = Val(LV.SelectedItem.SubItems(5))
        End If
           
        'txtWastage = Val(LV.SelectedItem.SubItems(5))
        If cmbProcess.Tag = "Lot No" And RcvingLot Then
            .Locked = True
            If Val(LV.SelectedItem.SubItems(4)) = 0 Then
                .Text = Val(LV.SelectedItem.SubItems(3))
                txtRate.Text = GetLastRate
            Else
                .Text = Val(LV.SelectedItem.SubItems(4))
                txtRate.Text = Val(LV.SelectedItem.SubItems(5))
            End If
        Else
            .Locked = False
        End If
        
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
    
End Sub

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        Item.SubItems(4) = Item.SubItems(3)
    Else
        Item.SubItems(4) = 0
    End If
End Sub

Private Sub LV_KeyPress(KeyAscii As Integer)
    Call LV_DblClick
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
                         Format(ServerDate, "ddMMyy") & TempID
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
            
            With txtRate
                .SelStart = 0
                .SelLength = Len(.Text)
                .SetFocus
            End With
            
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
    
    Dim iRefID As Long
    
    Call StartTrans(con)
    
    'con.Execute "Insert Into VendReceived(DT,VendID) Values('" & ServerDate & " " & ServerTime & "'," & cmbVend.ID & ")"
    con.Execute "Insert Into VendReceived(VendID,DT,RecieptID,UserID,ProcessID) Values(" & cmbVend.ID & _
     ",'" & DTPicker1 & " " & ServerTime & "',''," & UserID & "," & cmbProcess.ID & ")"
     
    iRefID = Val(con.Execute("Select Max(EntryID) From VendReceived").Fields(0).value & "")
    
    'Get The Next Process ID For Current Item...
    For i = 1 To LV.ListItems.Count
        
            If Val(LV.ListItems(i).SubItems(4)) > 0 Then
                '2. Now If Rcvd After Issued Then
                'Update The RcvdQty In VendIssdDetails....
                '-----------------------------------------------------------------
                With cmd
                    .ActiveConnection = con
                    .CommandType = adCmdStoredProc
                    .CommandText = "SP_InsertVendRcvdLost"
                    .Parameters("@ItemCode").value = LV.ListItems(i).SubItems(1)
                    .Parameters("@ProcessID").value = cmbProcess.ID
                    .Parameters("@Qty").value = 0
                    .Parameters("@WasteQty").value = Val(LV.ListItems(i).SubItems(4) & "")
                    .Parameters("@iRefID").value = iRefID
                    .Parameters("@VendID").value = cmbVend.ID
                    .Parameters("@LotNo").value = 0
                    .Parameters("@RcvDT").value = DTPicker1.value
                    .Parameters("@Rate").value = Val(LV.ListItems(i).SubItems(5) & "")
                    .Parameters("@RecID").value = LV.ListItems(i).SubItems(LV.ColumnHeaders.Count - 1)
                    .Parameters("@ReqAuth").value = UserReqAuth
                    .Parameters("@UserID").value = UserID
                    
                    .Execute
                End With
                '-----------------------------------------------------------------
                If UserReqAuth = 0 Then
                    Dim DetailRefID As Long
                    DetailRefID = Val(con.Execute("Select Max(EntryID) From VendIssdDetail").Fields(0).value & "")
                    con.Execute "Insert Into VendIssAuthDetail(UserID,IssID,DT) Values(" & _
                     UserID & "," & DetailRefID & ",'" & ServerDate & " " & ServerTime & "')"
                End If
            End If
    Next i
    
    Set rs = Nothing
    con.CommitTrans
    Saved = True
    
    Exit Function
err:
    MsgBox err.Description
    
End Function
    
Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Pic.Visible = False
    ElseIf KeyCode = 13 Then
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
                         Format(ServerDate, "ddMMyy") & TempID
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
                LV.SelectedItem.SubItems(6) = Format(DTPicker2.value, "dd-MMM-yyyy")
            Else
                LV.SelectedItem.SubItems(4) = Val(txtQty)
                LV.SelectedItem.SubItems(5) = Format(DTPicker2.value, "dd-MMM-yyyy")
            End If
            
            '-----------------------
            
            'Dim RcvdID As String
            
            If cmbProcess.Tag = "Lot No" Then
                
                If GotRcvID = 0 Then
                    GotRcvID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-10) As Int)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "' and UPPER(Left(RecieptID,3))='WST'").Fields(0).value & "") + 1
                Else
                    GotRcvID = GotRcvID + 1
                End If
                'LV.SelectedItem.SubItems(LV.ColumnHeaders.Count - 1) = "RCV-" & _
                 Format(DTPicker1, "ddMMyy") & "-" & Format(GotRcvID, "000")
                 LV.SelectedItem.SubItems(LV.ColumnHeaders.Count - 1) = "WST-" & _
                  Format(DTPicker2, "ddMMyy") & GotRcvID
            End If
                
            '-----------------------
            
            'Call AdjustQty
            Pic.Visible = False
            
            If LV.SelectedItem.Index < LV.ListItems.Count Then
                'If Val(LV.SelectedItem.SubItems(3)) = 0 Then
                 LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
                'End If
                LV.SetFocus
            End If
            
        ElseIf KeyAscii = 27 Then
            Pic.Visible = False
        End If
End Sub


Private Sub txtRate_KeyPressOld(KeyAscii As Integer)

    On Error GoTo err
    
    If KeyAscii = 13 Then
    
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
                         Format(ServerDate, "ddMMyy") & TempID
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
                LV.SelectedItem.SubItems(6) = Format(DTPicker1.value, "dd-MMM-yyyy")
            Else
                LV.SelectedItem.SubItems(4) = Val(txtQty)
                LV.SelectedItem.SubItems(5) = Format(DTPicker1.value, "dd-MMM-yyyy")
            End If
            
            '-----------------------
            
            'Dim RcvdID As String
            
            If cmbProcess.Tag = "Lot No" Then
                
                If GotRcvID = 0 Then
                    GotRcvID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-10) As Int)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "' and UPPER(Left(RecieptID,3))='LST'").Fields(0).value & "") + 1
                Else
                    GotRcvID = GotRcvID + 1
                End If
                'LV.SelectedItem.SubItems(LV.ColumnHeaders.Count - 1) = "RCV-" & _
                 Format(DTPicker1, "ddMMyy") & "-" & Format(GotRcvID, "000")
                 LV.SelectedItem.SubItems(LV.ColumnHeaders.Count - 1) = "LST-" & _
                  Format(DTPicker1, "ddMMyy") & GotRcvID
            Else
                If GotRcvID = 0 Then
                    GotRcvID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-10) As Int)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(DTPicker1, "dd MMM yy") & "' and UPPER(Left(RecieptID,3))='LST'").Fields(0).value & "") + 1
                Else
                    GotRcvID = GotRcvID + 1
                End If
                'LV.SelectedItem.SubItems(LV.ColumnHeaders.Count - 1) = "RCV-" & _
                 Format(DTPicker1, "ddMMyy") & "-" & Format(GotRcvID, "000")
                 LV.SelectedItem.SubItems(LV.ColumnHeaders.Count - 1) = "LST-" & _
                  Format(DTPicker1, "ddMMyy") & GotRcvID
            

            End If
                
            '-----------------------
            
            'Call AdjustQty
            Pic.Visible = False
            
            If LV.SelectedItem.Index < LV.ListItems.Count Then
                'If Val(LV.SelectedItem.SubItems(3)) = 0 Then
                 LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
                'End If
                LV.SetFocus
            End If
            DTPicker1.Enabled = False
            
    ElseIf KeyAscii = 27 Then
        Pic.Visible = False
    End If
    
    Call OnlyNums(KeyAscii, True)
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)

    On Error GoTo err
    
    If KeyAscii = 13 Then
            
'            If Val(txtRate) <= 0 Then
'                MsgBox "Can't Accept Zero.", vbInformation
'                txtRate.SetFocus
'                Exit Sub
'            End If
            
            If InitLotProcess Then
                LV.SelectedItem.SubItems(5) = Val(txtQty)
                LV.SelectedItem.SubItems(6) = Val(txtRate)
                LV.SelectedItem.SubItems(7) = Format(DTPicker1.value, "dd-MMM-yyyy")
            Else
                LV.SelectedItem.SubItems(4) = Val(txtQty)
                LV.SelectedItem.SubItems(5) = Val(txtRate)
                LV.SelectedItem.SubItems(6) = Format(DTPicker1.value, "dd-MMM-yyyy")
            End If
            
            '-----------------------
            
            'Dim RcvdID As String
            
            If cmbProcess.Tag = "Lot No" Then
                
                If GotRcvID = 0 Then
                    GotRcvID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-10) As Int)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "' and UPPER(Left(RecieptID,3))='LST'").Fields(0).value & "") + 1
                Else
                    GotRcvID = GotRcvID + 1
                End If
                'LV.SelectedItem.SubItems(LV.ColumnHeaders.Count - 1) = "RCV-" & _
                 Format(DTPicker1, "ddMMyy") & "-" & Format(GotRcvID, "000")
                 LV.SelectedItem.SubItems(LV.ColumnHeaders.Count - 1) = "LST-" & _
                  Format(DTPicker1, "ddMMyy") & GotRcvID
            Else
                If GotRcvID = 0 Then
                    GotRcvID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-10) As Int)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(DTPicker1, "dd MMM yy") & "' and UPPER(Left(RecieptID,3))='LST'").Fields(0).value & "") + 1
                Else
                    GotRcvID = GotRcvID + 1
                End If
                'LV.SelectedItem.SubItems(LV.ColumnHeaders.Count - 1) = "RCV-" & _
                 Format(DTPicker1, "ddMMyy") & "-" & Format(GotRcvID, "000")
                 LV.SelectedItem.SubItems(LV.ColumnHeaders.Count - 1) = "LST-" & _
                  Format(DTPicker1, "ddMMyy") & GotRcvID
            

            End If
                
            '-----------------------
            
            'Call AdjustQty
            Pic.Visible = False
            
            If LV.SelectedItem.Index < LV.ListItems.Count Then
                'If Val(LV.SelectedItem.SubItems(3)) = 0 Then
                 LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
                'End If
                LV.SetFocus
            End If
            DTPicker1.Enabled = False
            
    ElseIf KeyAscii = 27 Then
        Pic.Visible = False
    End If
    
    Call OnlyNums(KeyAscii, True)
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
