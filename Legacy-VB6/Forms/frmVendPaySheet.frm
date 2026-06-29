VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmVendPaySheet 
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   360
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
      TabIndex        =   1
      Top             =   45
      Width           =   10635
      Begin ComboList.Usercontrol1 cmbVend 
         Height          =   285
         Left            =   3390
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   4
         Top             =   75
         Width           =   10635
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Sheet"
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
            TabIndex        =   5
            Top             =   165
            Width           =   10605
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Sheet"
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
            TabIndex        =   6
            Top             =   180
            Width           =   10590
         End
      End
      Begin VB.Frame FraCash 
         BorderStyle     =   0  'None
         Height          =   5955
         Left            =   120
         TabIndex        =   2
         Top             =   1275
         Width           =   10395
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
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   1110
            Visible         =   0   'False
            Width           =   945
            Begin VB.TextBox txtQty 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   255
               Left            =   0
               TabIndex        =   12
               Top             =   0
               Width           =   950
            End
         End
         Begin VB.Frame Frame6 
            Height          =   645
            Left            =   7200
            TabIndex        =   7
            Top             =   5235
            Width           =   3165
            Begin MSForms.CommandButton cmdPrint 
               Height          =   360
               Left            =   120
               TabIndex        =   14
               Top             =   180
               Width           =   1425
               ForeColor       =   0
               Caption         =   "Print         "
               PicturePosition =   327683
               Size            =   "2514;635"
               Accelerator     =   80
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
               Left            =   1590
               TabIndex        =   0
               Top             =   180
               Width           =   1425
               Caption         =   "Cancel            "
               PicturePosition =   327683
               Size            =   "2514;635"
               Accelerator     =   67
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
            Height          =   5205
            Left            =   60
            TabIndex        =   10
            Top             =   0
            Width           =   10305
            _ExtentX        =   18177
            _ExtentY        =   9181
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
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Vend ID"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Vender Name"
               Object.Width           =   5292
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Balance"
               Object.Width           =   2469
            EndProperty
         End
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   6735
         TabIndex        =   13
         Top             =   885
         Width           =   1305
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "2302;635"
         Accelerator     =   82
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   3
         Top             =   735
         Width           =   6540
         ForeColor       =   4194304
         BackColor       =   12504280
         Caption         =   $"frmVendPaySheet.frx":0000
         Size            =   "11536;397"
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
   Begin VB.Menu mnuPop 
      Caption         =   "Pop Menu"
      Begin VB.Menu mnuPrintPaySheet 
         Caption         =   "Print Payment Sheet"
      End
      Begin VB.Menu mnuPrintVendLedger 
         Caption         =   "Print Vender Ledger"
      End
   End
End
Attribute VB_Name = "frmVendPaySheet"
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
Private Sub cmbProcess_matched()

    On Error GoTo err
    
    LV.ListItems.Clear
    Call cmbVend.ClearVals
    cmbVend.AddItem "<All Venders>", "0"
    
    If cmbProcess.ID = "0" Then
        Call cmbVend.AddVals(con, "VenderName", "VVenders", "VendID")
    Else
        Call cmbVend.AddVals(con, "VenderName", "VVenders", "VendID", " Where PhaseID=" & cmbProcess.ID)
    End If
    
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbVend_matched()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim Sql As String
    
    InitLotProcess = False
    If LV.ColumnHeaders.Count = 7 Then LV.ColumnHeaders.Remove "OrderNo"
    With rs
        If cmbProcess.Tag = "Lot No" Then
            If RcvLot(cmbProcess.ID) Then
                Sql = "Select LotNo,ItemCode,ItemDescription,CompItemID,Sum(IssQty-RcvdQty) as Quantity From VVendIssItems1 Where IssQty-RcvdQty>0 and VendID=" & cmbVend.ID & " Group By LotNo,ItemCode,ItemDescription,CompItemID"
            Else
                InitLotProcess = True
                LV.ColumnHeaders.Add(2, "OrderNo", "Order No", 1200).Tag = "1200"
                Sql = "Select LotNo,OrderNo,ItemCode,ItemDescription,CompItemID,Sum(IssQty-RcvdQty) as Quantity From VVendIssItems1 Where IssQty-RcvdQty>0 and VendID=" & cmbVend.ID & " Group By LotNo,OrderNo,ItemCode,ItemDescription,CompItemID"
            End If
        Else
            Sql = "Select ItemCode,ItemDescription,CompItemID,Sum(IssQty-RcvdQty) as Quantity From VVendIssItems1 Where IssQty-RcvdQty>0 and VendID=" & cmbVend.ID & " Group By ItemCode,ItemDescription,CompItemID"
        End If
        
        .Open Sql, con, adOpenForwardOnly, adLockReadOnly
        
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.Add(, , LV.ListItems.Count + 1)
            If InitLotProcess Then ITM.ListSubItems.Add , , ![OrderNo] & ""
            ITM.ListSubItems.Add(, , ![ItemCode] & "").Tag = ![CompItemID] & ""
            ITM.ListSubItems.Add , , ![ItemDescription] & ""
            ITM.ListSubItems.Add , , ![Quantity] & ""
            ITM.ListSubItems.Add
            If cmbProcess.Tag = "Lot No" Then
                ITM.ListSubItems.Add , , ![LotNo] & ""
            Else
                ITM.ListSubItems.Add
            End If
            ITM.ListSubItems.Add
            ITM.ListSubItems.Add
            .MoveNext
        Loop
        .Close
        
    End With
    
    Dim RcvdID As String
    
        
    'RcvdID = "RCV-" & Val(cmbVend.ID) & "-" & _
     Format(ServerDate, "ddMMyy") & "-" & Format(Val(RcvdID) + 1, "000")
     
     If cmbProcess.Tag = "Lot No" Then
        RcvdID = Val(con.Execute("Select Max(right(RecieptNo,3)) From VendReceived Where DT='" & Format(ServerDate, "dd-MMM-yyyy") & "' and UPPER(Left(RecieptNo,2))='HS'").Fields(0).value & "")
        RcvdID = "HS-" & _
         Format(DTPicker1, "ddMMyy") & "-" & Format(Val(RcvdID) + 1, "000")
     Else
        RcvdID = Val(con.Execute("Select Max(right(RecieptNo,3)) From VendReceived Where DT='" & Format(ServerDate, "dd-MMM-yyyy") & "' and UPPER(Left(RecieptNo,3))='RCV'").Fields(0).value & "")
        RcvdID = "RCV-" & _
         Format(DTPicker1, "ddMMyy") & "-" & Format(Val(RcvdID) + 1, "000")
     End If
     
    txtRecNo = RcvdID
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdRefresh_Click()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    If cmbProcess.ID = "0" Then
        If cmbVend.ID = "0" Then
            Sql = "Select * From VVendCurrentBal"
        Else
            Sql = "Select * From VVendCurrentBal Where VendID=" & cmbVend.ID
        End If
    Else
        If cmbVend.ID = "0" Then
            Sql = "Select * From VVendCurrentBal Where PhaseID=" & cmbProcess.ID
        Else
            Sql = "Select * From VVendCurrentBal Where VendID=" & cmbVend.ID
        End If
    End If
    
    With rs
        .Open Sql, con, adOpenForwardOnly, adLockReadOnly
        
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.Add(, ![VendID] & "'", ![VendID1] & "")
            ITM.SubItems(1) = ![VenderName] & ""
            ITM.SubItems(2) = Format(0 - Val(![CurrentBal] & ""), "0.00")
            
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdSClear_Click()

    If Not AskForAutoApply Then Exit Sub
    If cmbPymtmode.ListIndex = 0 Then
        If Not SaveTransForCash Then Exit Sub
    Else
        If Not SaveTransForCheque Then Exit Sub
    End If

    Unload Me
    Me.Show
    Me.SetFocus
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

    Me.KeyPreview = True
    'DTPicker1.value = ServerDate
    
    'Call DTPicker1_Click
    'Call FillAccCmbs
    cmbProcess.Enabled = True
    cmbProcess.ListHeight = 2600
    cmbVend.ListHeight = 2600
    
    cmbProcess.AddItem "<All Processes>", "0"
    
    Call cmbProcess.AddVals(con, "Description", "Processes", "ProcessID", "Order By SNo")
    
    'lLotNo = 0
    Call SaveLV(LV)
    
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



Private Sub LV_DblClick()

    Call ResetLV(LV)
    
    With Pic
        If InitLotProcess Then
            .Move LV.Left + 10 + LV.ColumnHeaders(6).Left, LV.SelectedItem.Top + LV.Top
        Else
            .Move LV.Left + 10 + LV.ColumnHeaders(5).Left, LV.SelectedItem.Top + LV.Top
        End If
        .Visible = True
    End With
    
    With txtQty
        If InitLotProcess Then
            .Text = Val(LV.SelectedItem.SubItems(5))
        Else
            .Text = Val(LV.SelectedItem.SubItems(4))
        End If
            'txtWastage = Val(LV.SelectedItem.SubItems(5))
        If cmbProcess.Tag = "Lot No" And RcvLot(cmbProcess.ID) Then
            .Locked = True
            .Text = Val(LV.SelectedItem.SubItems(3))
        Else
            .Locked = False
        End If
        
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
    
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

Private Sub mnuPrintPaySheet_Click()

    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    
    Load frmPrevRpt
    
    
    'Dim rpt As New rptVendPaySheet
    Dim rpt As craxddrt.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptVendPaySheet.rpt") '
    'Dim FFDefs As FormulaFieldDefinitions
    'Dim FFDef As CRAXDDT.FormulaFieldDefinition
    'Set FFDefs = rpt.FormulaFields
    
    'For Each FFDef In FFDefs
    '    If FFDef.Name = "{@FromTo}" Then
    '        FFDef.Text = "' From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
    '    ElseIf FFDef.Name = "{@EmpName}" Then
    '        FFDef.Text = "'" & cmbEmp.Text & "'"
    '    End If
    'Next
    
    
    frmPrevRpt.ShowReport " ", rpt
    
    Screen.MousePointer = vbDefault

    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description

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
        
        If Not txtQty.Locked Then
            If cmbProcess.Tag = "Lot No" Then
                If lLotNo = 0 Then
                    lLotNo = Val(con.Execute("Select Max(Cast(LotNo As int)) From VendRcvdDetail").Fields(0).value & "") + 1
                Else
                    lLotNo = lLotNo + 1
                End If
                
                If InitLotProcess Then
                    LV.SelectedItem.SubItems(6) = lLotNo
                Else
                    LV.SelectedItem.SubItems(5) = lLotNo
                End If
            End If
        End If
        If InitLotProcess Then
            LV.SelectedItem.SubItems(5) = Val(txtQty)
        Else
            LV.SelectedItem.SubItems(4) = Val(txtQty)
        End If
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
    
    con.Execute "Insert Into VendReceived(DT,RecieptNo,VendID) Values('" & DTPicker1 & "','" & txtRecNo & "'," & cmbVend.ID & ")"
    
    iRefID = Val(con.Execute("Select Max(EntryID) From VendReceived").Fields(0).value & "")
    
    'Get The Next Process ID For Current Item...
    Dim ReceivingLot As Boolean
    ReceivingLot = RcvLot(cmbProcess.ID)
    
    For i = 1 To LV.ListItems.Count
            
            If InitLotProcess Then
                Qty = Val(LV.ListItems(i).SubItems(5))
                lLotNo = LV.ListItems(i).SubItems(6)
            Else
                Qty = Val(LV.ListItems(i).SubItems(4))
                If LV.ColumnHeaders.Count = 6 Then
                    lLotNo = LV.ListItems(i).SubItems(5)
                Else
                    lLotNo = 0
                End If
            End If
            
            If Qty > 0 Then
                '2. Now If Rcvd After Issued Then
                'Update The RcvdQty In VendIssdDetails....
                '-----------------------------------------------------------------
                If InitLotProcess Then
                    With cmd
                        .ActiveConnection = con
                        .CommandType = adCmdStoredProc
                    
                        .CommandText = "SP_InsertVendRcvdILP"
                        
                        .Parameters("@OrderNo").value = LV.ListItems(i).SubItems(1) & ""
                        .Parameters("@ItemCode").value = LV.ListItems(i).SubItems(2) & ""
                        .Parameters("@ProcessID").value = cmbProcess.ID
                        .Parameters("@Qty").value = Val(LV.ListItems(i).SubItems(5) & "")
                        
                        .Parameters("@RefID").value = iRefID
                        .Parameters("@VendID").value = cmbVend.ID
                        .Parameters("@LotNo").value = lLotNo
                        .Execute
                    End With
                Else
                    If ReceivingLot Then
                        With cmd
                            .ActiveConnection = con
                            .CommandType = adCmdStoredProc
                        
                            .CommandText = "SP_InsertVendRcvdLots"
                            
                            .Parameters("@ProcessID").value = cmbProcess.ID
                            .Parameters("@Qty").value = Val(LV.ListItems(i).SubItems(4) & "")
                            
                            .Parameters("@RefID").value = iRefID
                            .Parameters("@VendID").value = cmbVend.ID
                            .Parameters("@LotNo").value = lLotNo
                            .Execute
                        End With
                    Else
                        With cmd
                            .ActiveConnection = con
                            .CommandType = adCmdStoredProc
                        
                            .CommandText = "SP_InsertVendRcvd"
                            .Parameters("@ItemCode").value = LV.ListItems(i).SubItems(1)
                            .Parameters("@ProcessID").value = cmbProcess.ID
                            .Parameters("@Qty").value = Val(LV.ListItems(i).SubItems(4) & "")
                            .Parameters("@WasteQty").value = 0 'Val(LV.ListItems(i).SubItems(5) & "")
                            .Parameters("@iRefID").value = iRefID
                            .Parameters("@VendID").value = cmbVend.ID
                            .Parameters("@LotNo").value = lLotNo
                            .Execute
                        End With
                    End If
                '-----------------------------------------------------------------
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
    


