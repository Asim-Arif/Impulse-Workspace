VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmVendRcvItemsManual 
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
      Height          =   9870
      Left            =   105
      TabIndex        =   1
      Top             =   45
      Width           =   11685
      Begin ComboList.Usercontrol1 cmbProcess 
         Height          =   285
         Left            =   165
         TabIndex        =   8
         Top             =   945
         Width           =   4035
         _ExtentX        =   7117
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
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "Clear All"
         Height          =   255
         Left            =   1260
         TabIndex        =   17
         Top             =   9180
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdCheckAll 
         Caption         =   "Check &All"
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   9180
         Visible         =   0   'False
         Width           =   975
      End
      Begin ComboList.Usercontrol1 cmbVend 
         Height          =   285
         Left            =   4170
         TabIndex        =   7
         Top             =   945
         Width           =   5790
         _ExtentX        =   10213
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
         Width           =   11685
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
            TabIndex        =   5
            Top             =   165
            Width           =   11565
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
            TabIndex        =   6
            Top             =   180
            Width           =   11550
         End
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   9990
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   945
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   149946371
         CurrentDate     =   37250
      End
      Begin VB.Frame FraCash 
         BorderStyle     =   0  'None
         Height          =   8505
         Left            =   120
         TabIndex        =   2
         Top             =   1275
         Width           =   11415
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
            Height          =   600
            Left            =   6150
            TabIndex        =   15
            Top             =   7890
            Width           =   5280
            Begin VB.CommandButton cmdSave 
               Caption         =   "Save && &New   "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   60
               TabIndex        =   10
               Top             =   165
               Width           =   1665
            End
            Begin VB.CommandButton cmdSaveClose 
               Caption         =   "&Save && Print"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   1770
               TabIndex        =   11
               Top             =   150
               Width           =   1665
            End
            Begin VB.CommandButton cmdCancel 
               Caption         =   "&Cancel"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   3480
               TabIndex        =   12
               Top             =   165
               Width           =   1665
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
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   1110
            Visible         =   0   'False
            Width           =   945
            Begin VB.TextBox txtQty 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   255
               Left            =   0
               TabIndex        =   14
               Top             =   0
               Width           =   950
            End
         End
         Begin MSComctlLib.ListView LV 
            Height          =   7830
            Left            =   45
            TabIndex        =   9
            Top             =   0
            Width           =   11355
            _ExtentX        =   20029
            _ExtentY        =   13811
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
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "SNo"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Iss No."
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Order No."
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Item ID"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Item Description"
               Object.Width           =   6703
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Qty"
               Object.Width           =   1587
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
               Text            =   "Date"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Receive ID"
               Object.Width           =   2646
            EndProperty
         End
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   3
         Top             =   735
         Width           =   11325
         ForeColor       =   4194304
         BackColor       =   12504280
         Caption         =   $"frmVendRcvItemsManual.frx":0000
         Size            =   "19976;397"
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
Attribute VB_Name = "frmVendRcvItemsManual"
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
Dim lLotNo As String
Dim GotRcvID As Long
Dim RcvingLot As Boolean
Dim UserReqAuth As Integer

Private Sub cmbProcess_matched()
   
    On Error GoTo err
        
'    If cmbProcess.ID = 9 Then
'        cmdCheckAll.Visible = True
'        cmdClearAll.Visible = True
'    Else
'        cmdCheckAll.Visible = False
'        cmdClearAll.Visible = False
'    End If
    Dim iVendID As Integer
    
    If cmbVend.MatchFound Then
        iVendID = cmbVend.ID
    End If
    
    If cmbProcess.ID = "0" Then
        cmbVend.AddVals con, "'{ ' + VendID1 + ' } ' + VenderName", "VMakers", "VendID", " WHERE ISNULL(PhaseID,0)<>0"
        Exit Sub
    End If
    
    Dim cmd As New ADODB.Command
    
    Call cmbVend.ClearVals
    
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "SP_ReqAuthForRcv"
        Set .ActiveConnection = con
        .Parameters("@UserID").Value = UserID
        .Parameters("@ProcessID").Value = cmbProcess.ID
        .Execute
        UserReqAuth = Val(.Parameters("@ReqAuth").Value & "")
        
        '.CommandText = "SP_VendersForRcving"
        '.Parameters("@ProcessID").value = cmbProcess.ID
        '.Parameters("@HicoVisible").value = HicoVisible
        
        'Set rs = .Execute
    End With
    
    If RcvLot(cmbProcess.ID) Then
        RcvingLot = True
    Else
        RcvingLot = False
    End If
    
    LV.ListItems.Clear
    GotRcvID = 0
    
    
    'If HicoVisible Then
    '    Call cmbVend.AddVals(con, "VenderName", "VVendIssItems", "VendID", " Where PhaseID=" & cmbProcess.ID & " Or PhaseID=-1")
    'Else
    '    Call cmbVend.AddVals(con, "VenderName", "VVendIssItems", "VendID", " Where PhaseID=" & cmbProcess.ID)
    'End If
    
    
'    If HicoVisible Then
'        If cmbProcess.ID = 7 Then
'            Call cmbVend.AddVals(con, "VenderName", "Venders", "VendID", " Where (PhaseID=" & cmbProcess.ID & " Or PhaseID=-1 ) And VendID In(Select VendID From VMakerIssItems Where ProcessID=" & cmbProcess.ID & "  Group By VendID Having Sum(RemQty)>0)")
'        Else
'            Call cmbVend.AddVals(con, "VenderName", "Makers", "VendID", " Where (PhaseID=" & cmbProcess.ID & " Or PhaseID=-1) And VendID In(Select VendID From VMakerIssItems Where ProcessID=" & cmbProcess.ID & " Group By VendID Having Sum(RemQty)>0)")
'        End If
'    Else
'        If cmbProcess.ID = 7 Then
'            Call cmbVend.AddVals(con, "VenderName", "Venders", "VendID", " Where PhaseID=" & cmbProcess.ID & "  And VendID In(Select VendID From VVendIssItems Where ProcessID=" & cmbProcess.ID & " Group By VendID Having Sum(RemQty)>0)")
'        Else
'            Call cmbVend.AddVals(con, "VenderName", "Venders", "VendID", " Where PhaseID=" & cmbProcess.ID & " And VendID In(Select VendID From VVendIssItems Where ProcessID=" & cmbProcess.ID & " Group By VendID Having Sum(RemQty)>0)")
'        End If
'    End If
    
    'Call cmbVend.AddVals(con, "'{ ' + VendID1 + ' } ' + VenderName", "VMakers", "VendID", " WHERE VendID In(SELECT VendID From VMakerIssItems WHERE RcvProcessID=" & cmbProcess.ID & " GROUP BY VendID HAVING SUM(RemQty)>0)")
    
    Dim lSNo1 As Long, lSNo2 As Long
    'lSNo1 = Val(con.Execute("Select SNo From VInitialLotProcess").Fields(0).value & "")
    
    lSNo2 = Val(con.Execute("SELECT SNo FROM Processes WHERE ProcessID=" & cmbProcess.ID).Fields(0).Value & "")
    
'    If lSNo2 >= lSNo1 Then
'        cmbProcess.Tag = "Lot No"
'    Else
'        cmbProcess.Tag = ""
'    End If
    'cmbProcess.Tag = ""
    'Set Focus To CmbVenders And Also Drop It Down!
'    cmbVend.SetFocus

    'SendKeys "a"
    'SendKeys Chr(8)
    
    'Call cmbVend.AddVals(con, "VenderName", "VVenders", "VendID", " Where ProcessID=" & cmbProcess.ID)
    If cmbProcess.ID = "0" Then
        cmbVend.AddVals con, "'{ ' + VendID1 + ' } ' + VenderName", "VMakers", "VendID", " WHERE ISNULL(PhaseID,0)<>0"
    Else
        cmbVend.AddVals con, "'{ ' + VendID1 + ' } ' + VenderName", "VMakers", "VendID", " WHERE VendID IN(SELECT VendID FROM MakerProcesses WHERE ProcessID=" & cmbProcess.ID & ")"
    End If
    If iVendID <> 0 Then
        cmbVend.ID = iVendID
    End If
    Exit Sub
    
err:
    MsgBox err.Description
End Sub

Private Sub cmbVend_matched()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim SQL As String
    
    InitLotProcess = False
    'If LV.ColumnHeaders.count = 8 Then LV.ColumnHeaders.Remove "OrderNo"
    LV.Checkboxes = False
    
    With rs
        .Open "SELECT * FROM VMakerInitialProcess WHERE VendID=" & cmbVend.ID, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            cmbProcess.ID = Val(!FirstProcessID & "")
        End If
        .Close
        
        If cmbProcess.Tag = "Lot No" Then
            If RcvingLot Then
                SQL = "Select * From VVendRcvLots Where VendID=" & cmbVend.ID & " And ProcessID=" & cmbProcess.ID & " "
                'LV.ColumnHeaders(6).Text = "Receive ID"
                LV.ColumnHeaders.Remove 6
                'LV.ColumnHeaders(7).Text = "Order No"
                LV.Checkboxes = True
            Else
                
                InitLotProcess = True
                'LV.ColumnHeaders.Add(2, "OrderNo", "Order No", 1200).Tag = "1200"
                'SQL = "Select LotNo,OrderNo,ItemCode,ItemDescription,Sum(IssQty-RcvdQty) as Quantity From VVendIssItems1 Where IssQty-RcvdQty>0 and VendID=" & cmbVend.ID & " and ReqAuth=0 Group By LotNo,OrderNo,ItemCode,ItemDescription"
                SQL = "Select * From VVendIssItems3 Where  VendID=" & cmbVend.ID & " And ProcessID=" & cmbProcess.ID & "  "
            End If
        Else
            'Sql = "SELECT OrderNo,ItemCode,ItemDescription,SUM(IssQty-RcvdQty)-dbo.F_AuthQtyForIss(ItemCode,null,null,null," & cmbVend.ID & ") AS Quantity FROM VMakerIssItems1 WHERE IssQty-RcvdQty>0 AND VendID=" & cmbVend.ID & " AND RcvProcessID=" & cmbProcess.ID & " AND ISNULL(ReqAuth,0)=0 AND RefID=" & txtID.Tag & " GROUP BY OrderNo,ItemCode,ItemDescription"
            SQL = "SELECT OrderNo,ItemCode,ItemDescription,IssQty-RcvdQty AS Quantity ,MainRecieptID,MainEntryID FROM VMakerIssItems WHERE IssQty-RcvdQty>0 AND VendID=" & cmbVend.ID & " AND ProcessID=" & cmbProcess.ID
        End If
        
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        .Filter = "Quantity<>0"
        LV.ListItems.Clear
        GotRcvID = 0
        Dim FColor As Long
        Do Until .EOF
        
            If cmbProcess.Tag = "Lot No" And Not RcvingLot And .AbsolutePosition = 1 Then
                FColor = vbRed
            'ElseIf cmbProcess.Tag = "Lot No" And RcvingLot Then
            '    If Not IsNull(![Early]) Then
            '        FColor = vbRed
            '    Else
            '        FColor = LV.ForeColor
            '    End If
            Else
                FColor = LV.ForeColor
            End If
            
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            ITM.ForeColor = FColor
            
            'If cmbProcess.Tag = "Lot No" Then ITM.Tag = ![DeliveryDT] & ""
            
            'If InitLotProcess Then ITM.ListSubItems.Add(, , ![OrderNo] & "").ForeColor = FColor
            ITM.ListSubItems.add(, , ![MainRecieptID] & "").ForeColor = FColor
            ITM.ListSubItems.add(, , ![OrderNo] & "").ForeColor = FColor
            ITM.ListSubItems.add(, , ![ItemCode] & "").ForeColor = FColor
            ITM.ListSubItems.add(, , ![ItemDescription] & "").ForeColor = FColor
            ITM.ListSubItems.add(, , ![Quantity] & "").ForeColor = FColor
            ITM.ListSubItems(4).Bold = True
            ITM.ListSubItems.add
            
            ITM.ListSubItems(1).Tag = Val(!MainEntryID & "")
            If cmbProcess.Tag = "Lot No" Then
                'itm.ListSubItems.Add , , ![LotNo] & ""
                ITM.ListSubItems(5).Tag = ![LotNo] & ""
            Else
                ITM.ListSubItems(5).Tag = ""
            End If
            
            If RcvingLot Then
                
                ITM.ListSubItems.add , , ![RecieptID] & ""
                'itm.ListSubItems.Add , , ![OrderNo] & ""
            Else
                ITM.ListSubItems.add
            End If
            
            
            .MoveNEXT
        Loop
        .Close
        
    End With
    
    Set rs = Nothing
    
    Dim myDT As String
    If cmbProcess.Tag = "Lot No" Then
        If LV.ListItems.count > 0 Then
            myDT = LV.ListItems(1).Tag
            For i = 1 To LV.ListItems.count
                Set ITM = LV.ListItems(i)
                If myDT = ITM.Tag Then
                    'Set The ForeColor Of all The ListSubItems To Red.
                    ITM.ForeColor = vbRed
                    For j = 1 To ITM.ListSubItems.count
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

Private Sub cmdSClear_Click()

    If Saved Then
        FormProcessID = cmbProcess.ID
        Unload Me
        Call ShowRcvVForm(FormProcessID)
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
    con.Execute "Delete From PrintSession"
End Sub


Private Sub cmdCheckAll_Click()
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = True
        Call LV_ItemCheck(LV.ListItems(i))
    Next
End Sub

Private Sub cmdClearAll_Click()
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdSave_Click()

    If Saved Then
        FormProcessID = cmbProcess.ID
        Unload Me
        Call ShowRcvVForm(FormProcessID, True)
    End If
    
End Sub

Private Sub cmdSaveClose_Click()

    If Saved Then
'        If cmbProcess.ID <= 7 Then
'            If Not UserReqAuth Then Call PrintIssRcvRpts("Original", False)
'        End If
        Call PrintRcvingSlips
        DoEvents
        
        Unload Me
        con.Execute "Delete From PrintSession"
    End If
    
End Sub

Private Sub PrintRcvingSlips()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report, strSelection As String
    Set rpt = rptApp.OpenReport(RptPath & "\RcvSlip.rpt")
    
    strSelection = "ISNULL({PrintSession.RecieptNo})=False"
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

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
                LV.SelectedItem.SubItems(LV.ColumnHeaders.count - 1) = ""
                'Reset The IssID's
                Dim TempID As Long
                TempID = 0
                
                For i = LV.SelectedItem.Index + 1 To LV.ListItems.count
                    TempID = Val(Right(LV.ListItems(i).SubItems(LV.ColumnHeaders.count - 1), 3)) - 1
                    If TempID > 0 Then
                        LV.ListItems(i).SubItems(LV.ColumnHeaders.count - 1) = "RCV-" & _
                         Format(ServerDate, "ddMMyy") & TempID
                    End If
                Next i
                If GotRcvID > 0 Then GotRcvID = GotRcvID - 1
                
                
                
                Pic.Visible = False
                Exit Sub
            End If
            
            If InitLotProcess Then
                If Val(txtQty) > Val(LV.SelectedItem.SubItems(3)) Then
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
                        lLotNo = Val(con.Execute("Select Max(Cast (LotNo As Int)) From VendRcvdDetail").Fields(0).Value & "") + 1
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
                LV.SelectedItem.SubItems(4) = Val(txtQty)
                LV.SelectedItem.SubItems(5) = Format(DTPicker2.Value, "dd-MMM-yyyy")
            Else
                LV.SelectedItem.SubItems(4) = Val(txtQty)
                LV.SelectedItem.SubItems(5) = Format(DTPicker2.Value, "dd-MMM-yyyy")
            End If
            
            '-----------------------
            
            'Dim RcvdID As String
            
            If cmbProcess.Tag = "Lot No" Then
                If GotRcvID = 0 Then
                    GotRcvID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-9) As Int)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(DTPicker2, "dd MMM yy") & "' and UPPER(Left(RecieptID,2))='HS'").Fields(0).Value & "") + 1  'and UPPER(Left(RecieptID,3))<>'HSR'
                Else
                    GotRcvID = GotRcvID + 1
                End If
                LV.SelectedItem.SubItems(LV.ColumnHeaders.count - 1) = "HS-" & _
                Format(DTPicker2, "ddMMyy") & GotRcvID
            Else
                If GotRcvID = 0 Then
                    GotRcvID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-10) As Int)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(DTPicker2, "dd MMM yy") & "' and UPPER(Left(RecieptID,3))='RCV'").Fields(0).Value & "") + 1
                Else
                    GotRcvID = GotRcvID + 1
                End If
                'LV.SelectedItem.SubItems(LV.ColumnHeaders.Count - 1) = "RCV-" & _
                 Format(DTPicker1, "ddMMyy") & "-" & Format(GotRcvID, "000")
                 LV.SelectedItem.SubItems(LV.ColumnHeaders.count - 1) = "RCV-" & _
                  Format(DTPicker2, "ddMMyy") & GotRcvID
            End If
                
            '-----------------------
            
            'Call AdjustQty
            Pic.Visible = False
            
            If LV.SelectedItem.Index < LV.ListItems.count Then
                'If Val(LV.SelectedItem.SubItems(3)) = 0 Then
                 LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
                'End If
                LV.SetFocus
            End If
            
        ElseIf KeyAscii = 27 Then
            Pic.Visible = False
        End If
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

    'Me.KeyPreview = True
    DTPicker1.Value = ServerDate

    cmbProcess.Enabled = True
    cmbProcess.ListHeight = 2600
    cmbVend.ListHeight = 2600
    cmbProcess.AddItem "<All Processes>", "0"
    Call cmbProcess.AddVals(con, "Description", "Processes", "ProcessID", "Order By SNo")
    cmbProcess.ID = "0"
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

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
    
    Set Col_OrdQty = Nothing
    Set frmVendRcvItems = Nothing
End Sub

Private Sub Form_Resize()
   FAll.Move (Me.ScaleWidth - FAll.Width) / 2, (Me.ScaleHeight - FAll.Height) / 2
End Sub



Private Sub LV_DblClick()
    
    If RcvingLot Then Exit Sub 'Then CheckBoxes Will Work
    Call ResetLV(LV)
    
    With Pic
        If InitLotProcess Then
            .Move LV.Left + 10 + LV.ColumnHeaders(6).Left, LV.SelectedItem.Top + LV.Top
        Else
            .Move LV.Left + 10 + LV.ColumnHeaders(7).Left, LV.SelectedItem.Top + LV.Top
        End If
        .Visible = True
    End With
    
    With txtQty
        If InitLotProcess Then
            .Text = Val(LV.SelectedItem.SubItems(5))
        Else
            .Text = Val(LV.SelectedItem.SubItems(6))
        End If
            'txtWastage = Val(LV.SelectedItem.SubItems(5))
        If cmbProcess.Tag = "Lot No" And RcvingLot Then
            .Locked = True
            .Text = Val(LV.SelectedItem.SubItems(4))
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
    If KeyAscii = 13 Then
        Call LV_DblClick
    End If
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

Private Sub txtQtyOld_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    If KeyAscii = 13 Then
        
        If Val(txtQty) = 0 Then
            If InitLotProcess Then
                LV.SelectedItem.SubItems(5) = Val(txtQty)
            Else
                LV.SelectedItem.SubItems(4) = Val(txtQty)
            End If
            LV.SelectedItem.SubItems(LV.ColumnHeaders.count - 1) = ""
            'Reset The IssID's
            Dim TempID As Long
            TempID = 0
            
            For i = LV.SelectedItem.Index + 1 To LV.ListItems.count
                TempID = Val(Right(LV.ListItems(i).SubItems(LV.ColumnHeaders.count - 1), 3)) - 1
                If TempID > 0 Then
                    LV.ListItems(i).SubItems(LV.ColumnHeaders.count - 1) = "RCV-" & _
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
                    lLotNo = Val(con.Execute("Select Max(Cast(LotNo As int)) From VendRcvdDetail").Fields(0).Value & "") + 1
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
                    GotRcvID = Val(con.Execute("Select Max(right(RecieptID,3)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "' and UPPER(Left(RecieptID,2))='HS'").Fields(0).Value & "") + 1  'and UPPER(Left(RecieptID,3))<>'HSR'
                Else
                    GotRcvID = GotRcvID + 1
                End If
                LV.SelectedItem.SubItems(LV.ColumnHeaders.count - 1) = "HS-" & _
                 Format(DTPicker1, "ddMMyy") & "-" & Format(GotRcvID, "000")
            Else
                If GotRcvID = 0 Then
                    GotRcvID = Val(con.Execute("Select Max(right(RecieptID,3)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "' and UPPER(Left(RecieptID,3))='RCV'").Fields(0).Value & "") + 1
                Else
                    GotRcvID = GotRcvID + 1
                End If
                 LV.SelectedItem.SubItems(LV.ColumnHeaders.count - 1) = "RCV-" & _
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
    Dim strDetail As String
    
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
            
            Col_OrdQty.add Obj, ![OrderNo] & "|" & ![ItemCode]
            Set Obj = Nothing
            .MoveNEXT
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
            
            strDetail = strDetail & ![OrderNo] & "=" & TempQty + TempWasteQty & ","
            
            .MoveNEXT
        Loop
        .Close
    End With
    
    strDetail = Left(strDetail, Len(strDetail) - 1)
    LV.SelectedItem.ListSubItems(7) = strDetail
    LV.SelectedItem.ListSubItems(7).ToolTipText = strDetail
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
    
    iRefID = Val(con.Execute("Select Max(EntryID) From VendReceived").Fields(0).Value & "")
    
    'Get The Next Process ID For Current Item...
    For i = 1 To LV.ListItems.count
        
            If Val(LV.ListItems(i).SubItems(6)) > 0 Then
                Dim iNextProcID As Integer
                'Dim cmd As New ADODB.Command
                With cmd
                    .ActiveConnection = con
                    .CommandType = adCmdStoredProc
                    .CommandText = "SP_GetNextProcID"
                    .Parameters("@ItemCode").Value = LV.ListItems(i).SubItems(1)
                    .Parameters("@ProcID").Value = cmbProcess.ID
                    .Execute
                    iNextProcID = Val(.Parameters("@NextProcID").Value & "")
                End With
                
                For j = 1 To Col_OrdQty.count
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
                                .Parameters("@OrderNo").Value = Col_OrdQty(j).OrderNo
                                .Parameters("@ItemCode").Value = Col_OrdQty(j).ItemCode
                                .Parameters("@ProcessID").Value = cmbProcess.ID
                                .Parameters("@Qty").Value = Col_OrdQty(j).lQty + Col_OrdQty(j).lWastage
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
    
    If LV.ListItems.count = 0 Then Exit Function
    Dim rs As New ADODB.Recordset
    Dim Col As New Collection
    
    Dim cmd As New ADODB.Command
    Dim Qty As Long
    
    Dim iRefID As Long
    Dim iIssRefID As Long
    
    Call StartTrans(con)
    
    con.Execute "Insert Into VendReceived(VendID,DT,RecieptID,UserID,ProcessID) Values(" & cmbVend.ID & _
     ",'" & DTPicker1 & " " & ServerTime & "',''," & UserID & "," & cmbProcess.ID & ")"
    
    iRefID = Val(con.Execute("Select Max(EntryID) From VendReceived").Fields(0).Value & "")
    
    'Get The Next Process ID For Current Item...
    
    For i = 1 To LV.ListItems.count
            
            If InitLotProcess Then
                Qty = Val(LV.ListItems(i).SubItems(4))
                lLotNo = Val(LV.ListItems(i).ListSubItems(4).Tag & "")
            Else
                Qty = Val(LV.ListItems(i).SubItems(6))
                'If LV.ColumnHeaders.Count = 7 Then
                '    lLotNo = LV.ListItems(i).ListSubItems(4).Tag
                'Else
                '    lLotNo = 0
                'End If
                If Left(LV.ListItems(i).SubItems(6), 2) = "HS" Then
                    lLotNo = LV.ListItems(i).ListSubItems(4).Tag
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
                        
                        .Parameters("@ItemCode").Value = LV.ListItems(i).SubItems(1) & ""
                        .Parameters("@ProcessID").Value = cmbProcess.ID
                        .Parameters("@Qty").Value = Val(LV.ListItems(i).SubItems(4) & "")
                        
                        .Parameters("@RefID").Value = iRefID
                        .Parameters("@VendID").Value = cmbVend.ID
                        .Parameters("@LotNo").Value = lLotNo
                        
                        .Parameters("@RecID").Value = LV.ListItems(i).SubItems(LV.ColumnHeaders.count - 1)
                        .Parameters("@ReqAuth").Value = UserReqAuth
                        
                        
                        .Execute
                    End With
                Else
                    If RcvingLot Then
                        With cmd
                            .ActiveConnection = con
                            .CommandType = adCmdStoredProc
                            'if
                            .CommandText = "SP_InsertVendRcvdLots"
                            
                            .Parameters("@ProcessID").Value = cmbProcess.ID
                            .Parameters("@Qty").Value = Val(LV.ListItems(i).SubItems(4) & "")
                            
                            .Parameters("@RefID").Value = iRefID
                            .Parameters("@VendID").Value = cmbVend.ID
                            .Parameters("@LotNo").Value = lLotNo
                            '.Parameters("@RcvDT").value = DTPicker1 'LV.ListItems(i).SubItems(LV.ColumnHeaders.Count - 2)
                            .Parameters("@RecID").Value = LV.ListItems(i).SubItems(LV.ColumnHeaders.count - 1)
                            .Parameters("@ReqAuth").Value = UserReqAuth
                            
                            .Execute
                            
                        End With
                    Else
                        With cmd
                            .ActiveConnection = con
                            .CommandType = adCmdStoredProc
                        
                            .CommandText = "SP_InsertVendRcvdNew"
                            .Parameters("@ItemCode").Value = LV.ListItems(i).SubItems(3)
                            .Parameters("@ProcessID").Value = cmbProcess.ID
                            .Parameters("@Qty").Value = Val(LV.ListItems(i).SubItems(6) & "")
                            .Parameters("@WasteQty").Value = 0 'Val(LV.ListItems(i).SubItems(5) & "")
                            .Parameters("@iRefID").Value = iRefID
                            .Parameters("@VendID").Value = cmbVend.ID
                            .Parameters("@LotNo").Value = lLotNo
                            .Parameters("@RcvDT").Value = LV.ListItems(i).SubItems(LV.ColumnHeaders.count - 2)
                            .Parameters("@RecID").Value = LV.ListItems(i).SubItems(LV.ColumnHeaders.count - 1)
                            .Parameters("@OrderNo").Value = LV.ListItems(i).SubItems(2)
                            .Parameters("@DeliveryRefID").Value = Val(LV.ListItems(i).ListSubItems(1).Tag)
                            If cmbProcess.ID = 4 Then
                                .Parameters("@ReqAuth").Value = 0 '1
                            Else
                                .Parameters("@ReqAuth").Value = UserReqAuth
                            End If
                            .Parameters("@UserID").Value = UserID
                            
                            .Execute
                        End With
                    End If
                '-----------------------------------------------------------------
                End If
                If UserReqAuth = 0 And cmbProcess.ID <> 4 Then
                    Dim DetailRefID As Long
                    DetailRefID = Val(con.Execute("Select Max(EntryID) From VendRcvdDetail").Fields(0).Value & "")
                    con.Execute "Insert Into VendRcvAuthDetail(UserID,RcvID,DT) Values(" & _
                     UserID & "," & DetailRefID & ",'" & ServerDate & " " & ServerTime & "')"
                End If
                con.Execute "Insert Into PrintSession(RecieptNo) Values('" & LV.ListItems(i).SubItems(LV.ColumnHeaders.count - 1) & "')"
            End If
            
            ''Now Issue This Item If processID=9 and User Doesn't Require Authentication.
            If cmbProcess.ID = 9 And UserReqAuth = 0 Then
            
                con.Execute "Insert Into VendIssued(VendID,DT,RecieptID,UserID,ProcessID) " & _
                 "Values(129,'" & ServerDate & " " & ServerTime & "',''," & _
                 UserID & ",12)"
                 
                iIssRefID = Val(con.Execute("Select Max(EntryID) From VendIssued").Fields(0).Value & "")
                
                With cmd
                    .ActiveConnection = con
                    .CommandType = adCmdStoredProc
                    .CommandText = "SP_UpdateForIARLots"
                    
                    .Parameters("@ProcessID").Value = 12
                    .Parameters("@Qty").Value = Val(LV.ListItems(i).SubItems(4) & "") 'Val(LVIssued.ListItems(i).SubItems(6))
                    .Parameters("@LotNo").Value = lLotNo
                    .Parameters("@RefiD").Value = iIssRefID
                    .Parameters("@Rate").Value = 0 'Val(LVIssued.ListItems(i).SubItems(5))
                    .Parameters("@RecID").Value = LV.ListItems(i).SubItems(LV.ColumnHeaders.count - 1)
                    .Parameters("@ReqAuth").Value = UserReqAuth
                    
                    .Execute
                End With
                
                DetailRefID = Val(con.Execute("Select Max(EntryID) From VendIssdDetail").Fields(0).Value & "")
                con.Execute "Insert Into VendIssAuthDetail(UserID,IssID,DT) Values(" & _
                 UserID & "," & DetailRefID & ",'" & ServerDate & " " & ServerTime & "')"
                
                con.Execute "Update VendRcvdDetail Set IssQty=RcvdQty Where RecieptID='" & LV.ListItems(i).SubItems(LV.ColumnHeaders.count - 1) & "'"
                 
            End If
            'Debug.Print i; " of " & LV.ListItems.count
    Next i
    
    Set rs = Nothing
    con.CommitTrans
    Saved = True
    
    Exit Function
err:
    MsgBox err.Description
    
End Function
    

Private Sub txtID_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Dim strIssID As String
        Dim lEntryID As Long, lVendID As Long, lProcessID As Long, strItemID As String
        strIssID = txtID1.Text & txtID.Text
        Dim rs As New ADODB.Recordset
        With rs
            .Open "SELECT EntryID,VendID,ProcessID,ItemID FROM VendIssued WHERE RecieptID='" & strIssID & "'", con, adOpenForwardOnly, adLockReadOnly
            If .EOF = False Then
                lEntryID = Val(!EntryID & "")
                lVendID = Val(!VendID & "")
                lProcessID = Val(!ProcessID & "")
                strItemID = !ITemID & ""
            End If
            .Close
            cmbProcess.ID = lProcessID
            cmbVend.ID = lVendID
            txtID.Tag = lEntryID
        End With
        Set rs = Nothing
    End If
    
    cmbVend.ClearVals
    Call cmbVend.AddVals(con, "'{ ' + VendID1 + ' } ' + VenderName", "VMakers", "VendID", " WHERE VendID=" & lVendID)
    cmbVend.ID = lVendID
    
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)

    On Error GoTo err
    Dim dRate As Double
    If KeyAscii = 13 Then
        
        If InitLotProcess Then
            If Val(txtQty) > Val(LV.SelectedItem.SubItems(4)) Then
                MsgBox "Invalid Quantities!", vbInformation
                Exit Sub
            End If
        Else
            If Val(txtQty) > Val(LV.SelectedItem.SubItems(5)) Then
                MsgBox "Invalid Quantities!", vbInformation
                Exit Sub
            End If
        End If
        dRate = Val(con.Execute("Select Rate From VendAssItems Where VendID=" & cmbVend.ID & " AND ItemID='" & LV.SelectedItem.SubItems(3) & "'").Fields(0).Value & "")
        If dRate = 0 Then
            MsgBox "Rate Is Zero in Vender Profile.", vbInformation
            Exit Sub
        End If
'        dtpicker1.SetFocus
        If Val(txtQty) = 0 Then
                If InitLotProcess Then
                    LV.SelectedItem.SubItems(6) = Val(txtQty)
                Else
                    LV.SelectedItem.SubItems(5) = Val(txtQty)
                End If
                LV.SelectedItem.SubItems(LV.ColumnHeaders.count - 1) = ""
                'Reset The IssID's
                Dim TempID As Long
                TempID = 0
                
                For i = LV.SelectedItem.Index + 1 To LV.ListItems.count
                    TempID = Val(Right(LV.ListItems(i).SubItems(LV.ColumnHeaders.count - 1), 3)) - 1
                    If TempID > 0 Then
                        LV.ListItems(i).SubItems(LV.ColumnHeaders.count - 1) = "RCV-" & _
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
                If Val(txtQty) > Val(LV.SelectedItem.SubItems(5)) Then
                    MsgBox "Invalid Quantities!", vbInformation
                    Exit Sub
                End If
            End If
            
            If Not txtQty.Locked Then
                If cmbProcess.Tag = "Lot No" Then
                    If lLotNo = 0 Then
                        lLotNo = Val(con.Execute("Select Max(Cast (LotNo As Int)) From VendRcvdDetail").Fields(0).Value & "") + 1
                    Else
                        lLotNo = lLotNo + 1
                    End If
                    
                    If InitLotProcess Then
                        LV.SelectedItem.ListSubItems(5).Tag = lLotNo
                    Else
                        LV.SelectedItem.ListSubItems(5).Tag = lLotNo
                    End If
                End If
            End If
            
            If InitLotProcess Then
                LV.SelectedItem.SubItems(5) = Val(txtQty)
                LV.SelectedItem.SubItems(6) = Format(DTPicker1.Value, "dd-MMM-yyyy")
            Else
                LV.SelectedItem.SubItems(6) = Val(txtQty)
                LV.SelectedItem.SubItems(7) = Format(DTPicker1.Value, "dd-MMM-yyyy")
            End If
            
            '-----------------------
            
            'Dim RcvdID As String
            
            If cmbProcess.Tag = "Lot No" Then
                If GotRcvID = 0 Then
                    GotRcvID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-9) As Int)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(DTPicker1, "dd MMM yy") & "' AND ProcessID=7 and UPPER(Left(RecieptID,2))='HS' AND Left(RecieptID,3)<>'HSR' AND Left(RecieptID,5)<>'HSPOL' AND Left(RecieptID,5)<>'HSPAC'").Fields(0).Value & "") + 1
                    'GotRcvID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-9) As Int)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(DTPicker1, "dd MMM yy") & "' AND ProcessID=7 and UPPER(Left(RecieptID,2))='HS' AND Left(RecieptID,3)<>'HSR'").Fields(0).value & "") + 1
                    'GotRcvID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-9) As Int)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(DTPicker1, "dd MMM yy") & "' AND Replace(convert(varchar,GetDate(),3),'/','')='" & Format(DTPicker1, "ddMMyy") & "' and UPPER(Left(RecieptID,2))='HS'").Fields(0).value & "") + 1
                    'GotRcvID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-9) As Int)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(DTPicker1, "dd MMM yy") & "' and UPPER(Left(RecieptID,2))='HS'").Fields(0).value & "") + 1
                Else
                    GotRcvID = GotRcvID + 1
                End If
                LV.SelectedItem.SubItems(LV.ColumnHeaders.count - 1) = "HS-" & _
                Format(DTPicker1, "ddMMyy") & GotRcvID
            Else
                If GotRcvID = 0 Then
                    GotRcvID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-10) As Int)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(DTPicker1, "dd MMM yy") & "' and UPPER(Left(RecieptID,3))='RCV'").Fields(0).Value & "") + 1
                Else
                    GotRcvID = GotRcvID + 1
                End If
                'LV.SelectedItem.SubItems(LV.ColumnHeaders.Count - 1) = "RCV-" & _
                 Format(DTPicker1, "ddMMyy") & "-" & Format(GotRcvID, "000")
                 LV.SelectedItem.SubItems(LV.ColumnHeaders.count - 1) = "RCV-" & _
                  Format(DTPicker1, "ddMMyy") & GotRcvID
            End If
                
            '-----------------------
            
            'Call AdjustQty
            Pic.Visible = False
            
            If LV.SelectedItem.Index < LV.ListItems.count Then
                'If Val(LV.SelectedItem.SubItems(3)) = 0 Then
                 LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
                'End If
                LV.SetFocus
            End If
            DTPicker1.Enabled = False
    ElseIf KeyAscii = 27 Then
        Pic.Visible = False
    End If
    
    Call OnlyNums(KeyAscii, False)
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
