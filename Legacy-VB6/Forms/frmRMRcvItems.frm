VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmRMRcvItems 
   ClientHeight    =   7425
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
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   11040
   WindowState     =   2  'Maximized
   Begin VB.Frame FALL 
      Height          =   7260
      Left            =   105
      TabIndex        =   10
      Top             =   45
      Width           =   10635
      Begin ComboList.Usercontrol1 cmbVend 
         Height          =   285
         Left            =   150
         TabIndex        =   0
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
      Begin VB.Frame Frame1 
         Height          =   645
         Left            =   0
         TabIndex        =   13
         Top             =   75
         Width           =   10635
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Receive Raw Material"
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
            TabIndex        =   14
            Top             =   165
            Width           =   10605
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Receive Raw Material"
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
            TabIndex        =   15
            Top             =   180
            Width           =   10590
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   3480
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   960
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   44695555
         CurrentDate     =   37250
      End
      Begin VB.Frame FraCash 
         BorderStyle     =   0  'None
         Height          =   5955
         Left            =   120
         TabIndex        =   11
         Top             =   1275
         Width           =   10395
         Begin ComboList.Usercontrol1 cmbItem 
            Height          =   285
            Left            =   75
            TabIndex        =   2
            Top             =   240
            Width           =   2265
            _ExtentX        =   3995
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
         End
         Begin ComboList.Usercontrol1 cmbItemName 
            Height          =   285
            Left            =   2325
            TabIndex        =   3
            Top             =   240
            Width           =   5355
            _ExtentX        =   9446
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
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   585
            Left            =   9495
            TabIndex        =   6
            Top             =   -30
            Width           =   825
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
            ScaleWidth      =   960
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   1110
            Visible         =   0   'False
            Width           =   960
            Begin VB.TextBox txtWastage 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   255
               Left            =   0
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   0
               Width           =   950
            End
         End
         Begin VB.Frame Frame6 
            Height          =   660
            Left            =   5205
            TabIndex        =   16
            Top             =   5235
            Width           =   5175
            Begin MSForms.CommandButton cmdSClose 
               Height          =   360
               Left            =   1770
               TabIndex        =   8
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
               TabIndex        =   9
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
               TabIndex        =   7
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
            Height          =   4635
            Left            =   60
            TabIndex        =   17
            Top             =   600
            Width           =   10305
            _ExtentX        =   18177
            _ExtentY        =   8176
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
               Text            =   "SNo"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "RM ID"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "RM Description"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Rate"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Qty"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Amount"
               Object.Width           =   2117
            EndProperty
         End
         Begin MSForms.TextBox txtAddQty 
            Height          =   285
            Left            =   8580
            TabIndex        =   5
            Top             =   240
            Width           =   870
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "1535;503"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.Label Label2 
            Height          =   225
            Index           =   0
            Left            =   15
            TabIndex        =   20
            Top             =   0
            Width           =   9435
            ForeColor       =   4194304
            BackColor       =   12504280
            Caption         =   $"frmRMRcvItems.frx":0000
            Size            =   "16642;397"
            BorderColor     =   10860487
            BorderStyle     =   1
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.TextBox txtRate 
            Height          =   285
            Left            =   7695
            TabIndex        =   4
            Top             =   240
            Width           =   870
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            Size            =   "1535;503"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   12
         Top             =   735
         Width           =   4845
         ForeColor       =   4194304
         BackColor       =   12504280
         Caption         =   "  RM Vender                                                    Date"
         Size            =   "8546;397"
         BorderColor     =   10860487
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
End
Attribute VB_Name = "frmRMRcvItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dWastage As Double
Dim dRcvdQty As Double
Dim Col_OrdQty As New Collection
Dim MyCol1 As New Collection
Dim IsItemsFirst As Boolean

Private Sub cmbItem_matched()
    
    On Error GoTo err
    cmbItemName.ID = cmbItem.ID
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Rate From RMVendAssItems Where RMID=" & cmbItem.ID, con, adOpenForwardOnly, adLockReadOnly
        txtRate = Val(.Fields(0) & "")
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbItemName_matched()
    cmbItem.ID = cmbItemName.ID
End Sub

Private Sub cmbVend_matched()
    
    LV.ListItems.Clear
    Call cmbItem.ClearVals
    Call cmbItem.AddVals(con, "RMID1", "VRMVendAssItems", "RMID", " Where VendID=" & cmbVend.ID)
    
    Call cmbItemName.ClearVals
    Call cmbItemName.AddVals(con, "RMName", "VRMVendAssItems", "RMID", " Where VendID=" & cmbVend.ID)
    
End Sub

Private Sub cmdAdd_Click()
        
    On Error GoTo err
    If cmbVend.MatchFound = False Then
        MsgBox "Invalid Vender.", vbInformation
        Exit Sub
    ElseIf cmbItem.MatchFound = False Or cmbItemName.MatchFound = False Then
        MsgBox "Invalid Item.", vbInformation
        Exit Sub
    ElseIf Val(txtRate) <= 0 Then
        MsgBox "Invalid Rate.", vbInformation
        txtRate.SetFocus
        Exit Sub
    ElseIf Val(txtAddQty) <= 0 Then
        MsgBox "Invalid Qty.", vbInformation
        txtAddQty.SetFocus
        Exit Sub
    End If
    
    
    Dim itm As ListItem
    Set itm = LV.ListItems.Add(, cmbItem.ID & "'", LV.ListItems.Count + 1)
    itm.SubItems(1) = cmbItem.Text
    itm.SubItems(2) = cmbItemName.Text
    itm.SubItems(3) = Val(txtRate)
    itm.SubItems(4) = Val(txtAddQty)
    itm.SubItems(5) = Val(txtRate) * Val(txtAddQty)
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub cmdSClose_Click()

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


    Me.KeyPreview = True
    DTPicker1.value = ServerDate
    
    cmbVend.ListHeight = 2600
    cmbItem.ListHeight = 2600
    cmbItemName.ListHeight = 2600
    
    Call cmbVend.ClearVals
    Call cmbVend.AddVals(con, "VenderName", "RawMaterialVenders", "VendID")
    
    
    
    Call SaveLV(LV)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Set Col_OrdQty = Nothing
    Set frmVendRcvItems = Nothing
End Sub

Private Sub Form_Resize()
   FALL.Move (Me.ScaleWidth - FALL.Width) / 2, (Me.ScaleHeight - FALL.Height) / 2
End Sub



Private Sub LV_DblClick()
    
    If LV.ListItems.Count = 0 Then Exit Sub
    Call ResetLV(LV)
    
    With Pic
        .Move LV.Left + 10 + LV.ColumnHeaders(7).Left, LV.SelectedItem.Top + LV.Top
        .Visible = True
    End With
    
    With txtWastage
        .Text = Val(LV.SelectedItem.SubItems(6))
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
    
End Sub


Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'Exit Sub
    'If LV.ListItems.Count = 0 Then Exit Sub
    'If Val(LV.SelectedItem.SubItems(6)) = 0 Then Exit Sub
    'Dim iCol As Integer
    'iCol = GetColNo(LV, X)
    
    'If iCol = 8 Then
    '    Load frmVendRcvItems1
    '    LV.SelectedItem.ListSubItems(7).Text = frmVendRcvItems1.FillList(Col_OrdQty, LV.SelectedItem.SubItems(1))
    '    LV.SelectedItem.ListSubItems(7).Tag = LV.SelectedItem.ListSubItems(7).Text
    'End If
    
End Sub



Private Sub txtWastage_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Val(txtWastage) > Val(LV.SelectedItem.SubItems(5)) Then
            MsgBox "Invalid Quantities!", vbInformation
            Exit Sub
        End If
        
        LV.SelectedItem.SubItems(6) = Val(txtWastage)
        'LV.SelectedItem.SubItems(6) = Val(txtQty) + Val(txtWastage)
        'Call AdjustQty
        Pic.Visible = False
    ElseIf KeyAscii = 27 Then
        Pic.Visible = False
    End If
    
    Call OnlyNums(KeyAscii, False)
    
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
    Dim ReqAuth As Boolean
    Dim iRefID As Long
    
    If CurrentUserName = "Administrator" Then
        ReqAuth = False
    Else
        ReqAuth = True
    End If
    
    Call StartTrans(con)
    
    con.Execute "Insert Into RMRcvd(VendID,DT) Values(" & cmbVend.ID & ",'" & DTPicker1 & "')"
    
    iRefID = Val(con.Execute("Select Max(ID) From RMRcvd").Fields(0).value & "")
    
    'Get The Next Process ID For Current Item...
    For i = 1 To LV.ListItems.Count
        With LV.ListItems(i)
            
                '-----------------------------------------------------------------
                 
                'con.Execute "Update RM Set QtyInStock=IsNull(QtyInStock,0)+" & _
                 Val(.SubItems(6)) & " Where RMID=" & .ListSubItems(2).Tag
                
                con.Execute "Insert Into RMRcvdItems(RefID,RMID,Rate,Qty,ReqAuth)" & _
                " Values(" & iRefID & "," & Val(.Key) & "," & Val(.SubItems(3)) & _
                "," & Val(.SubItems(4)) & "," & Abs(ReqAuth) & ")"
                
                
                '-----------------------------------------------------------------
            
        End With
    Next i
    
    Set rs = Nothing
    con.CommitTrans
    Saved = True
    
    Exit Function
err:
    MsgBox err.Description
    
End Function
    


