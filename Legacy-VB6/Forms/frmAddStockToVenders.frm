VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmAddStockToVenders 
   ClientHeight    =   9915
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10755
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
   ScaleHeight     =   9915
   ScaleWidth      =   10755
   WindowState     =   2  'Maximized
   Begin VB.Frame FALL 
      Height          =   9855
      Left            =   105
      TabIndex        =   1
      Top             =   45
      Width           =   10635
      Begin ComboList.Usercontrol1 cmbVend 
         Height          =   285
         Left            =   3405
         TabIndex        =   21
         Top             =   960
         Width           =   5505
         _ExtentX        =   9710
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
         Left            =   195
         TabIndex        =   7
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
      Begin ComboList.Usercontrol1 cmbItem 
         Height          =   285
         Left            =   195
         TabIndex        =   14
         Top             =   1485
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
         Left            =   2460
         TabIndex        =   13
         Top             =   1485
         Width           =   5925
         _ExtentX        =   10451
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
         Left            =   9630
         TabIndex        =   16
         Top             =   1275
         Width           =   825
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
            Caption         =   "Issue ItemsTo Venders"
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
            Caption         =   "Issue ItemsTo Venders"
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
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   8925
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   960
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   54001667
         CurrentDate     =   37250
      End
      Begin VB.Frame FraCash 
         BorderStyle     =   0  'None
         Height          =   7815
         Left            =   120
         TabIndex        =   2
         Top             =   1920
         Width           =   10395
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
            Height          =   600
            Left            =   5085
            TabIndex        =   17
            Top             =   7200
            Width           =   5280
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
               TabIndex        =   20
               Top             =   165
               Width           =   1665
            End
            Begin VB.CommandButton cmdSaveClose 
               Caption         =   "&Save && Close"
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
               TabIndex        =   19
               Top             =   165
               Width           =   1665
            End
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
               TabIndex        =   18
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
            Left            =   600
            ScaleHeight     =   255
            ScaleWidth      =   2355
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   840
            Visible         =   0   'False
            Width           =   2355
            Begin VB.TextBox txtQty 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   255
               Left            =   0
               TabIndex        =   10
               Top             =   0
               Width           =   950
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   285
               Left            =   960
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   0
               Width           =   1400
               _ExtentX        =   2461
               _ExtentY        =   503
               _Version        =   393216
               CustomFormat    =   "dd-MMM-yyyy"
               Format          =   54001667
               CurrentDate     =   37250
            End
         End
         Begin MSComctlLib.ListView LV 
            Height          =   7140
            Left            =   60
            TabIndex        =   8
            Top             =   15
            Width           =   10305
            _ExtentX        =   18177
            _ExtentY        =   12594
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
               Text            =   "Item ID"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Item Description"
               Object.Width           =   5644
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Vender"
               Object.Width           =   3881
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Rcvd. Qty"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Rcv ID"
               Object.Width           =   2117
            EndProperty
         End
      End
      Begin MSForms.TextBox txtAddQty 
         Height          =   285
         Left            =   8415
         TabIndex        =   15
         Top             =   1485
         Width           =   1170
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         Size            =   "2064;503"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   12
         Top             =   1245
         Width           =   9435
         ForeColor       =   4194304
         BackColor       =   12504280
         Caption         =   $"frmAddStockToVenders.frx":0000
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
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   3
         Top             =   735
         Width           =   10305
         ForeColor       =   4194304
         BackColor       =   12504280
         Caption         =   $"frmAddStockToVenders.frx":00AE
         Size            =   "18177;397"
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
Attribute VB_Name = "frmAddStockToVenders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dWastage As Double
Dim dRcvdQty As Double
Dim Col_OrdQty As New Collection
Dim MyCol1 As New Collection
Dim IsItemsFirst As Boolean
Dim UserReqAuth As Integer
Dim GotRcvID As Long
Dim lLotNo As Long
Private Sub cmbItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAddQty.SetFocus
    End If
End Sub

Private Sub cmbItem_matched()
    cmbItemName.ID = cmbItem.ID
End Sub

Private Sub cmbItemName_matched()
    cmbItem.ID = cmbItemName.ID
End Sub

Private Sub cmbProcess_matched()
    
    On Error GoTo err
    
    LV.ListItems.Clear
    Dim cmd As New ADODB.Command
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "SP_ReqAuthForRcv"
        Set .ActiveConnection = con
        .Parameters("@UserID").value = UserID
        .Parameters("@ProcessID").value = cmbProcess.ID
        .Execute
        UserReqAuth = IIf(IsNull(.Parameters("@ReqAuth").value), 0, .Parameters("@ReqAuth").value)
    End With
    
    'txtSupervisor.Text = con.Execute("Select Supervisor From Processes Where ProcessID=" & cmbProcess.ID).Fields(0).value & ""
    If cmbProcess.ID = 7 Then
        Call cmbVend.AddVals(con, "VenderName", "Venders", "VendID", " Where PhaseID=" & cmbProcess.ID & " OR RepairVend=1")
    Else
        Call cmbVend.AddVals(con, "VenderName", "Venders", "VendID", " Where PhaseID=" & cmbProcess.ID)
    End If
    
    Dim RcvdID As String
    
     
    txtRecNo = RcvdID

    Set rs = Nothing
    
    GotRcvID = 0

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
    
    
    'Set Focus To CmbVenders And Also Drop It Down!
    cmbVend.SetFocus
    SendKeys "a"
    SendKeys Chr(8)
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbVend_matched()
    
    On Error GoTo err
    LV.ListItems.Clear
    
    cmbItem.AddVals con, "ItemID", "VVendAssItems", "ItemID", " Where VendID=" & Val(cmbVend.ID) & "  Order By ItemID"
    cmbItemName.AddVals con, "ItemName", "VVendAssItems", "ItemID", " Where VendID=" & Val(cmbVend.ID) & "  Order By ItemID"
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdAdd_Click()

    On Error GoTo err
    
    Dim ITM As ListItem
    
    If GotRcvID = 0 Then
        If cmbProcess.ID > 7 Then 'Lot Process
            GotRcvID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-9) As Int)) From VVendIssdDetail Where Convert(varchar,DT,6)='" & Format(DTPicker1, "dd MMM yy") & "' and UPPER(Left(RecieptID,2))='HS'").Fields(0).value & "") + 1
        Else
            GotRcvID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-10) As Int)) From VVendIssdDetail Where Convert(varchar,DT,6)='" & Format(DTPicker1, "dd MMM yy") & "' and UPPER(Left(RecieptID,3))='ISS'").Fields(0).value & "") + 1
        End If
    Else
        GotRcvID = GotRcvID + 1
    End If
    
    If cmbProcess.ID > 7 Then
        If lLotNo = 0 Then
            lLotNo = Val(con.Execute("Select Max(Cast (LotNo As Int)) From VendRcvdDetail").Fields(0).value & "") + 1
        Else
            lLotNo = lLotNo + 1
        End If
    End If
    
    Set ITM = LV.ListItems.Add(, , LV.ListItems.Count + 1)
    ITM.ListSubItems.Add(, , cmbItem.Text).Tag = lLotNo
    ITM.ListSubItems.Add , , cmbItemName.Text
    ITM.ListSubItems.Add(, , cmbVend.Text).Tag = cmbVend.ID
    ITM.ListSubItems.Add , , Val(txtAddQty)
    If cmbProcess.ID > 7 Then
        ITM.SubItems(LV.ColumnHeaders.Count - 1) = "HS-" & _
         Format(DTPicker1, "ddMMyy") & GotRcvID
    Else
        ITM.SubItems(LV.ColumnHeaders.Count - 1) = "ISS-" & _
         Format(DTPicker1, "ddMMyy") & GotRcvID
    End If
        
    txtAddQty = 0
    cmbItem.SetFocus
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdSClear_Click()

    'If Not AskForAutoApply Then Exit Sub
    'If cmbPymtmode.ListIndex = 0 Then
    '    If Not SaveTransForCash Then Exit Sub
    'Else
    '    If Not SaveTransForCheque Then Exit Sub
    'End If
    If Saved Then
        FormProcessID = cmbProcess.ID
        Unload Me
        Call ShowRcvSForm(FormProcessID)
    End If
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



Private Sub cmdSave_Click()
    If Saved Then
        'FormProcessID = cmbProcess.ID
        Unload Me
        'Call ShowRcvSForm(FormProcessID)
        frmAddStockToVenders.Show
    End If
End Sub

Private Sub cmdSaveClose_Click()
    
    If Saved Then
        Unload Me
    End If
    
End Sub

Private Sub DTPicker2_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo err
    
    If KeyCode = 13 Then
        If Val(txtQty) > Val(LV.SelectedItem.SubItems(3)) Then
            MsgBox "Invalid Quantities!", vbInformation
            Exit Sub
        End If
        
        If GotRcvID = 0 Then
            GotRcvID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-10) As Int)) From VVendRcvdItems Where Convert(varchar,RcvDT,6)='" & Format(DTPicker2, "dd MMM yy") & "' and UPPER(Left(RecieptID,3))='RCV'").Fields(0).value & "") + 1
        Else
            GotRcvID = GotRcvID + 1
        End If
        'LV.SelectedItem.SubItems(LV.ColumnHeaders.Count - 1) = "RCV-" & _
         Format(DTPicker1, "ddMMyy") & "-" & Format(GotRcvID, "000")
         LV.SelectedItem.SubItems(LV.ColumnHeaders.Count - 1) = "RCV-" & _
          Format(DTPicker2, "ddMMyy") & GotRcvID
        
        
        LV.SelectedItem.SubItems(4) = Val(txtQty)
        LV.SelectedItem.SubItems(5) = Format(DTPicker2.value, "dd-MMM-yyyy")
        
        
        
        'Call AdjustQty
        Pic.Visible = False
        
        If LV.SelectedItem.Index < LV.ListItems.Count Then
            'If Val(LV.SelectedItem.SubItems(3)) = 0 Then
             LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
            'End If
            LV.SetFocus
        End If
        
    ElseIf KeyCode = 27 Then
        Pic.Visible = False
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Activate()

    'Check For User Access
    If PermissionDenied Then
        Me.Hide
        MsgBox "You Donot Have Permissions To Use This Option. Please Contact Your Administrator", vbInformation
        Unload Me
        Exit Sub
    End If
    
    
'    cmbItem.AddVals con, "ItemID", "VItems1", "ItemID", " Order By ItemID"
'    cmbItemName.AddVals con, "CompleteItemName", "VItems1", "ItemID", " Order By ItemID"
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo err
    lLotNo = 0
    Me.KeyPreview = True
    DTPicker1.value = ServerDate
    DTPicker2.value = ServerDate
    
    cmbProcess.Enabled = True
    cmbProcess.ListHeight = 2600
    
    Call cmbProcess.AddVals(con, "Description", "Processes", "ProcessID", "Where ProcessID<=7 Order By SNo")
    cmbItem.ListHeight = 2200
    cmbItemName.ListHeight = 2200
    cmbVend.ListHeight = 2200
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
        .Move LV.Left + 10 + LV.ColumnHeaders(5).Left, LV.SelectedItem.Top + LV.Top
        .Visible = True
    End With
    
    With txtQty
        .Text = Val(LV.SelectedItem.SubItems(4))
        'txtWastage = Val(LV.SelectedItem.SubItems(5))
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
    
    
End Sub


Private Sub LV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call LV_DblClick
    End If
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

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    
    If KeyAscii = 13 Then
        
        If GotRcvID = 0 Then
            GotRcvID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-10) As Int)) From VVendRcvdItems Where Convert(varchar,RcvDT,6)='" & Format(DTPicker2, "dd MMM yy") & "' and UPPER(Left(RecieptID,3))='RCV'").Fields(0).value & "") + 1
        Else
            GotRcvID = GotRcvID + 1
        End If
        'LV.SelectedItem.SubItems(LV.ColumnHeaders.Count - 1) = "RCV-" & _
         Format(DTPicker1, "ddMMyy") & "-" & Format(GotRcvID, "000")
         LV.SelectedItem.SubItems(LV.ColumnHeaders.Count - 1) = "RCV-" & _
          Format(DTPicker2, "ddMMyy") & GotRcvID
        
        
        LV.SelectedItem.SubItems(4) = Val(txtQty)
        LV.SelectedItem.SubItems(5) = Format(DTPicker2.value, "dd-MMM-yyyy")
        
        
        
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
    
    'con.Execute "Insert Into VendReceived(DT,RecieptNo,VendID) Values('" & DTPicker1 & "','" & txtRecNo & "'," & cmbVend.ID & ")"
    con.Execute "Insert Into VendIssued(VendID,DT,RecieptID," & _
     "ProcessID,UserID) Values(" & cmbVend.ID & ",'" & DTPicker1 & " " & ServerTime & "','" & txtRecNo & "'," & _
     cmbProcess.ID & "," & UserID & ")"
    
    iRefID = Val(con.Execute("Select Max(EntryID) From VendIssued").Fields(0).value & "")
    
    'Get The Next Process ID For Current Item...
    For i = 1 To LV.ListItems.Count
        
            If Val(LV.ListItems(i).SubItems(4)) > 0 Then
                '2. Now If Rcvd After Issued Then
                'Update The RcvdQty In VendIssdDetails....
                '-----------------------------------------------------------------
                With cmd
                    .ActiveConnection = con
                    .CommandType = adCmdStoredProc
                    .CommandText = "SP_AddStockVenders"
                    .Parameters("@ItemCode").value = LV.ListItems(i).SubItems(1)
                    .Parameters("@ProcessID").value = cmbProcess.ID
                    .Parameters("@Qty").value = Val(LV.ListItems(i).SubItems(4) & "")
                    
                    .Parameters("@iRefID").value = iRefID
                    
                    .Parameters("@RecID").value = LV.ListItems(i).SubItems(LV.ColumnHeaders.Count - 1)
                    .Parameters("@ReqAuth").value = UserReqAuth
                    .Parameters("@VendID").value = LV.ListItems(i).ListSubItems(3).Tag
                    If cmbProcess.ID > 7 Then
                        .Parameters("@LotNo").value = LV.ListItems(i).ListSubItems(1).Tag
                    End If
                    '.Parameters("@UserID").value = UserID
                    .Execute
                End With
                '-----------------------------------------------------------------
                If UserReqAuth = 0 Then
                    Dim DetailRefID As Long
                    DetailRefID = Val(con.Execute("Select Max(EntryID) From VendIssdDetail").Fields(0).value & "")
                    con.Execute "Insert Into VendIssAuthDetail(UserID,IssID,DT) Values(" & _
                     UserID & "," & DetailRefID & ",'" & DTPicker1 & " " & ServerTime & "')"
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
    


