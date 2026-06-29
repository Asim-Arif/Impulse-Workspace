VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmVendNewPO 
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
      Begin VB.TextBox txtOrderNo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6810
         TabIndex        =   22
         Top             =   960
         Width           =   1320
      End
      Begin ComboList.Usercontrol1 cmbVend 
         Height          =   285
         Left            =   2820
         TabIndex        =   7
         Top             =   960
         Width           =   3960
         _ExtentX        =   6985
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
         Left            =   135
         TabIndex        =   8
         Top             =   960
         Width           =   2640
         _ExtentX        =   4657
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
            Caption         =   "Purchase Order"
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
            Caption         =   "Purchase Order"
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
         Left            =   8175
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   960
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   46202883
         CurrentDate     =   37250
      End
      Begin VB.Frame FraCash 
         BorderStyle     =   0  'None
         Height          =   8505
         Left            =   120
         TabIndex        =   2
         Top             =   1275
         Width           =   11415
         Begin ComboList.Usercontrol1 cmbItem 
            Height          =   285
            Left            =   60
            TabIndex        =   16
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
            Left            =   2310
            TabIndex        =   17
            Top             =   240
            Width           =   5835
            _ExtentX        =   10292
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
            Left            =   9960
            TabIndex        =   20
            Top             =   -30
            Width           =   825
         End
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
            Left            =   6150
            TabIndex        =   12
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
               TabIndex        =   15
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
               TabIndex        =   14
               Top             =   165
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
               TabIndex        =   13
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
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1110
            Visible         =   0   'False
            Width           =   945
            Begin VB.TextBox txtQty 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Height          =   255
               Left            =   0
               TabIndex        =   11
               Top             =   0
               Width           =   950
            End
         End
         Begin MSComctlLib.ListView LV 
            Height          =   7230
            Left            =   30
            TabIndex        =   9
            Top             =   570
            Width           =   11355
            _ExtentX        =   20029
            _ExtentY        =   12753
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
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "SNo"
               Object.Width           =   1058
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Item ID"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Item Description"
               Object.Width           =   10583
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Rate"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Key             =   "Amt"
               Text            =   "Qty"
               Object.Width           =   1676
            EndProperty
         End
         Begin MSForms.TextBox txtRate 
            Height          =   285
            Left            =   8160
            TabIndex        =   21
            Top             =   240
            Width           =   870
            VariousPropertyBits=   746604575
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
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   9915
            ForeColor       =   4194304
            BackColor       =   12504280
            Caption         =   $"frmVendNewPO.frx":0000
            Size            =   "17489;397"
            BorderColor     =   10860487
            BorderStyle     =   1
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.TextBox txtAddQty 
            Height          =   285
            Left            =   9045
            TabIndex        =   18
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
      Begin MSComCtl2.DTPicker DTDelivery 
         Height          =   285
         Left            =   9555
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   945
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   46202883
         CurrentDate     =   37250
      End
      Begin MSForms.Label Label2 
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   3
         Top             =   735
         Width           =   10755
         ForeColor       =   4194304
         BackColor       =   12504280
         Caption         =   $"frmVendNewPO.frx":00B0
         Size            =   "18971;397"
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
Attribute VB_Name = "frmVendNewPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrderNo As String

Private Sub cmbProcess_matched()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Max(Cast(SubString(OrderNo,8,Len(OrderNo)-1) As Int)) From VendPurchaseOrder", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            txtOrderNo = "HS-" & Left(cmbProcess.Text, 3) & "-1"
        Else
            txtOrderNo = "HS-" & Left(cmbProcess.Text, 3) & "-" & Val(.Fields(0) & "") + 1
        End If
        .Close
    End With
    Set rs = Nothing
    
    Dim cmd As New ADODB.Command
    
    Call cmbVend.ClearVals
    
    With cmd
        .CommandType = adCmdStoredProc
        .CommandText = "SP_ReqAuthForRcv"
        Set .ActiveConnection = con
        .Parameters("@UserID").value = UserID
        .Parameters("@ProcessID").value = cmbProcess.ID
        .Execute
        UserReqAuth = Val(.Parameters("@ReqAuth").value & "")
        
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
'        Call cmbVend.AddVals(con, "VenderName", "Venders", "VendID", " Where (PhaseID=" & cmbProcess.ID & " Or PhaseID=-1) And VendID In(Select VendID From VVendIssItems Where ProcessID=" & cmbProcess.ID & " Group By VendID Having Sum(RemQty)>0)")
'    Else
    Call cmbVend.AddVals(con, "VenderName", "VVenders", "VendID", " Where ProcessID=" & cmbProcess.ID & " and VendType=0")
'    End If
    
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
    
    cmbItem.AddVals con, "ItemID", "VVendAssItems", "ItemID", " Where VendID=" & Val(cmbVend.ID) & "  Order By ItemID"
    cmbItemName.AddVals con, "ItemName", "VVendAssItems", "ItemID", " Where VendID=" & Val(cmbVend.ID) & "  Order By ItemID"
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdSClear_Click()
'
'    If Saved Then
'        FormProcessID = cmbProcess.ID
'        Unload Me
'        Call ShowRcvVForm(FormProcessID)
'    End If
'    'Call ClearFields

End Sub

Private Sub cmdSClose_Click()

    'If Not AskForAutoApply Then Exit Sub

    If Saved Then
        Unload Me
    End If

End Sub

Private Sub cmdAdd_Click()

    On Error GoTo err
    If Val(txtAddQty) <= 0 Then
        MsgBox "Invalid Qty.", vbInformation
        txtAddQty.SetFocus
        Exit Sub
    End If
    
    If Val(txtRate) <= 0 Then
        MsgBox "Invalid Rate", vbInformation
        txtRate.SetFocus
        Exit Sub
    End If
    
    Dim ITM As ListItem
    
    Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
    ITM.ListSubItems.add , , cmbItem.Text
    ITM.ListSubItems.add , , cmbItemName.Text
    ITM.ListSubItems.add , , Val(txtRate)
    ITM.ListSubItems.add , , Val(txtAddQty)
    
    txtRate = 0
    txtAddQty = 0
    cmbItem.SetFocus
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub
Private Sub cmbItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txtRate) = 0 Then
            txtRate.SetFocus
        Else
            txtAddQty.SetFocus
        End If
    End If
End Sub

Private Sub cmbItem_matched()

    On Error GoTo err
    
    cmbItemName.ID = cmbItem.ID
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Rate From VendAssItems Where VendID=" & cmbVend.ID & " AND ItemID='" & cmbItem.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        If Not .EOF Then txtRate = Val(.Fields(0) & "")
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

Private Sub cmdCancel_Click()
    Unload Me
    con.Execute "Delete From PrintSession"
End Sub


Private Sub cmdSave_Click()
'
'    If Saved Then
'        FormProcessID = cmbProcess.ID
'        Unload Me
'        Call ShowIssPOForm(FormProcessID)
'    End If
    
End Sub

Private Sub cmdSaveClose_Click()

    If Saved Then
        Unload Me
    End If
    
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
                LV.SelectedItem.SubItems(4) = Val(txtQty)
                LV.SelectedItem.SubItems(5) = Format(DTPicker2.value, "dd-MMM-yyyy")
            Else
                LV.SelectedItem.SubItems(4) = Val(txtQty)
                LV.SelectedItem.SubItems(5) = Format(DTPicker2.value, "dd-MMM-yyyy")
            End If
            
            '-----------------------
            
            'Dim RcvdID As String
            
            If cmbProcess.Tag = "Lot No" Then
                If GotRcvID = 0 Then
                    GotRcvID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-9) As Int)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(DTPicker2, "dd MMM yy") & "' and UPPER(Left(RecieptID,2))='HS'").Fields(0).value & "") + 1  'and UPPER(Left(RecieptID,3))<>'HSR'
                Else
                    GotRcvID = GotRcvID + 1
                End If
                LV.SelectedItem.SubItems(LV.ColumnHeaders.count - 1) = "HS-" & _
                Format(DTPicker2, "ddMMyy") & GotRcvID
            Else
                If GotRcvID = 0 Then
                    GotRcvID = Val(con.Execute("Select Max(Cast(right(RecieptID,Len(RecieptID)-10) As Int)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(DTPicker2, "dd MMM yy") & "' and UPPER(Left(RecieptID,3))='RCV'").Fields(0).value & "") + 1
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

Private Sub DTPicker1_Change()
    DTDelivery.value = DateAdd("m", 1, DTPicker1)
End Sub

Private Sub DTPicker1_Click()
    Call DTPicker1_Change
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
    
    strOrderNo = ""
    'Me.KeyPreview = True
    DTPicker1.value = ServerDate
    'DTPicker2.value = ServerDate
    If UserID = 1 Then DTPicker1.Enabled = True
    
    'Call DTPicker1_Click
    'Call FillAccCmbs
    cmbProcess.Enabled = True
    cmbProcess.ListHeight = 2600
    cmbVend.ListHeight = 2600
    Call cmbProcess.AddVals(con, "Description", "Processes", "ProcessID", "Order By SNo")
    
    lLotNo = 0
    Call SaveLV(LV)
    
    GotRcvID = 0
        
    cmbItem.ListHeight = 2200
    cmbItemName.ListHeight = 2200
    
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
    
    If RcvingLot Then Exit Sub 'Then CheckBoxes Will Work
    Call ResetLV(LV)
    
    With Pic
        If InitLotProcess Then
            .Move LV.Left + 10 + LV.ColumnHeaders(5).Left, LV.SelectedItem.Top + LV.Top
        Else
            .Move LV.Left + 10 + LV.ColumnHeaders(5).Left, LV.SelectedItem.Top + LV.Top
        End If
        .Visible = True
    End With
    
    With txtQty
        If InitLotProcess Then
            .Text = Val(LV.SelectedItem.SubItems(4))
        Else
            .Text = Val(LV.SelectedItem.SubItems(4))
        End If
            'txtWastage = Val(LV.SelectedItem.SubItems(5))
        If cmbProcess.Tag = "Lot No" And RcvingLot Then
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
                    GotRcvID = Val(con.Execute("Select Max(right(RecieptID,3)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "' and UPPER(Left(RecieptID,2))='HS'").Fields(0).value & "") + 1  'and UPPER(Left(RecieptID,3))<>'HSR'
                Else
                    GotRcvID = GotRcvID + 1
                End If
                LV.SelectedItem.SubItems(LV.ColumnHeaders.count - 1) = "HS-" & _
                 Format(DTPicker1, "ddMMyy") & "-" & Format(GotRcvID, "000")
            Else
                If GotRcvID = 0 Then
                    GotRcvID = Val(con.Execute("Select Max(right(RecieptID,3)) From VVendRcvdItems Where Convert(varchar,DT,6)='" & Format(ServerDate, "dd MMM yy") & "' and UPPER(Left(RecieptID,3))='RCV'").Fields(0).value & "") + 1
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
'
'    On Error GoTo err
'    Dim rs As New ADODB.Recordset
'    Dim lQty As Long, TotQty As Long, WasteQty As Long
'    Dim TempQty As Long, TempWasteQty As Long
'    Dim StrDetail As String
'
'    TempQty = 0
'    TotQty = Val(LV.SelectedItem.SubItems(6))
'    WasteQty = Val(LV.SelectedItem.SubItems(5))
'
'    With rs
'        .Open "Select OrderNo,ItemCode,TotQty From VTotIssdByOrder Where ItemCode='" & LV.SelectedItem.SubItems(1) & "'", con, adOpenForwardOnly
'
'        Do Until .EOF
'            Dim Obj As New RcvAdjOrderQty
'            Obj.OrderNo = ![OrderNo] & ""
'            Obj.ItemCode = ![ItemCode] & ""
'            Obj.lOrdQty = ![TotQty] & ""
'            Obj.lWastage = 0
'            Obj.lQty = 0
'
'            Col_OrdQty.add Obj, ![OrderNo] & "|" & ![ItemCode]
'            Set Obj = Nothing
'            .MoveNext
'        Loop
'        .MoveFirst
'
'        Do While lQty < TotQty
'
'            TempQty = TotQty - lQty
'            TempWasteQty = 0
'            If TempQty > ![TotQty] Then
'                TempQty = ![TotQty]
'            End If
'
'            If WasteQty > 0 Then
'                If TempQty > WasteQty Then
'                    TempQty = TempQty - WasteQty
'                    TempWasteQty = WasteQty
'                    WasteQty = 0
'                Else
'                    TempWasteQty = WasteQty
'                    TempQty = 0
'                    WasteQty = WasteQty - TempQty
'                End If
'            End If
'
'
'            Col_OrdQty(![OrderNo] & "|" & ![ItemCode]).lQty = TempQty
'            Col_OrdQty(![OrderNo] & "|" & ![ItemCode]).lWastage = TempWasteQty
'
'            lQty = lQty + TempQty + TempWasteQty
'
'            StrDetail = StrDetail & ![OrderNo] & "=" & TempQty + TempWasteQty & ","
'
'            .MoveNext
'        Loop
'        .Close
'    End With
'
'    StrDetail = Left(StrDetail, Len(StrDetail) - 1)
'    LV.SelectedItem.ListSubItems(7) = StrDetail
'    LV.SelectedItem.ListSubItems(7).ToolTipText = StrDetail
'    Set rs = Nothing
'    Exit Sub
'err:
'    MsgBox err.Description
'
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
    For i = 1 To LV.ListItems.count
        
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
    
    If strOrderNo = "" Then
        con.Execute "Insert Into VendPurchaseOrder(OrderNo,DT,VendID,ProcessID,DeliveryDT,UserID) Values('" & txtOrderNo & "','" & DTPicker1 & "'," & cmbVend.ID & "," & cmbProcess.ID & ",'" & DTDelivery.value & "'," & UserID & ")"
    Else
        'con.Execute "Update Vend"
    End If
    
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            con.Execute "Insert Into VendPurchOrdDetail(OrderNo,ItemCode,Rate,Qty) Values('" & txtOrderNo & "','" & .SubItems(1) & "'," & Val(.SubItems(3)) & "," & Val(.SubItems(4)) & ")"
        End With
    Next
    
    con.CommitTrans
    
    'Print Report
    Call PrintPO
    Saved = True
    
    Exit Function
err:
    MsgBox err.Description
    
End Function
    
Private Sub PrintPO()
    
    On Error GoTo err
    Screen.MousePointer = vbHourglass
    Dim rpt As craxddrt.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptVendPO.rpt") '
    
        
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VendPurchaseOrder.OrderNo}='" & txtOrderNo & "' ", rpt
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub txtQty_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    If KeyAscii = 13 Then
        LV.SelectedItem.SubItems(4) = Val(txtQty)
        Pic.Visible = False
    ElseIf KeyAscii = 27 Then
        Pic.Visible = False
    End If
    
    Call OnlyNums(KeyAscii, False)
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
