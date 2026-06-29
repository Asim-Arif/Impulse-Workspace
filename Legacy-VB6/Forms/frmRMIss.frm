VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmRMIss 
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10635
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   10635
   WindowState     =   2  'Maximized
   Begin VB.Frame FAll 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7155
      Left            =   45
      TabIndex        =   14
      Top             =   30
      Width           =   10560
      Begin VB.TextBox txtIssID 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   7470
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   24
         Top             =   1035
         Width           =   2250
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2040
         MaxLength       =   150
         TabIndex        =   23
         Top             =   6705
         Visible         =   0   'False
         Width           =   930
      End
      Begin ComboList.Usercontrol1 cmbCust 
         Height          =   285
         Left            =   285
         TabIndex        =   0
         Top             =   1050
         Width           =   2805
         _ExtentX        =   4948
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
      Begin VB.Frame FNew 
         BorderStyle     =   0  'None
         Height          =   570
         Left            =   285
         TabIndex        =   19
         Top             =   1350
         Width           =   10230
         Begin VB.CommandButton cmdAddItem 
            Caption         =   "&Add"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   9630
            Picture         =   "frmRMIss.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Add This Entry To Voucher"
            Top             =   15
            UseMaskColor    =   -1  'True
            Width           =   540
         End
         Begin VB.TextBox txtUnit 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   7665
            MaxLength       =   150
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   255
            Width           =   930
         End
         Begin VB.TextBox txtQty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "RptTime"
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   8625
            MaxLength       =   150
            TabIndex        =   7
            Top             =   255
            Width           =   930
         End
         Begin MSForms.ComboBox cmbDesc 
            Height          =   285
            Left            =   4140
            TabIndex        =   5
            Top             =   255
            Width           =   3495
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "6165;503"
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
            BackColor       =   &H00C5D1DC&
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"frmRMIss.frx":0532
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
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   9555
         End
         Begin MSForms.ComboBox cmbCatID 
            Height          =   285
            Left            =   0
            TabIndex        =   3
            Top             =   255
            Width           =   1965
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "3466;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin MSForms.ComboBox cmbItems 
            Height          =   285
            Left            =   1980
            TabIndex        =   4
            Top             =   255
            Width           =   2130
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "3757;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.TextBox txtOrderNo 
         Appearance      =   0  'Flat
         DataField       =   "RptTime"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   1
         Top             =   1050
         Width           =   2775
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   10560
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Issue Raw Material"
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
            TabIndex        =   17
            Top             =   165
            Width           =   10605
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00AFBDCB&
            BackStyle       =   0  'Transparent
            Caption         =   "Issue Raw Material"
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
            TabIndex        =   18
            Top             =   180
            Width           =   10590
         End
      End
      Begin VB.CommandButton cmdDelete 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9915
         Picture         =   "frmRMIss.frx":05DF
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Remove This Entry From Voucher"
         Top             =   1965
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   540
      End
      Begin MSComCtl2.DTPicker OrderDT 
         Height          =   300
         Left            =   5910
         TabIndex        =   2
         Top             =   1035
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   529
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
         Format          =   56360963
         CurrentDate     =   37250
      End
      Begin MSComctlLib.ListView LV 
         Height          =   4470
         Left            =   285
         TabIndex        =   9
         Top             =   1965
         Width           =   9540
         _ExtentX        =   16828
         _ExtentY        =   7885
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
            Text            =   "S No."
            Object.Width           =   1014
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Process"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cat"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Item ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Description"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Unit"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Qty"
            Object.Width           =   1499
         EndProperty
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
         Height          =   660
         Left            =   4665
         TabIndex        =   15
         Top             =   6405
         Width           =   5280
         Begin MSForms.CommandButton cmdCancel 
            Height          =   360
            Left            =   3510
            TabIndex        =   13
            Top             =   195
            Width           =   1665
            Caption         =   "Cancel            "
            PicturePosition =   327683
            Size            =   "2937;635"
            Picture         =   "frmRMIss.frx":0B11
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
            Left            =   90
            TabIndex        =   11
            Top             =   195
            Width           =   1665
            Caption         =   "Save & New   "
            PicturePosition =   327683
            Size            =   "2937;635"
            Accelerator     =   78
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdSClose 
            Height          =   360
            Left            =   1800
            TabIndex        =   12
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
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C5D1DC&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmRMIss.frx":0C23
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
         Index           =   1
         Left            =   285
         TabIndex        =   22
         Top             =   810
         Width           =   9435
      End
      Begin MSForms.CheckBox chkPrint 
         Height          =   345
         Left            =   270
         TabIndex        =   21
         Top             =   6450
         Width           =   1545
         VariousPropertyBits=   1015023643
         BackColor       =   -2147483633
         ForeColor       =   12582912
         DisplayStyle    =   4
         Size            =   "2725;609"
         Value           =   "0"
         Caption         =   "Print On Save"
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
End
Attribute VB_Name = "frmRMIss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrderNo As String

Public Sub AddPurchOrd(RMID As Long, Qty As Long)
    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select GroupID,RMName,GroupName From VRM Where RMID=" & RMID, con, adOpenForwardOnly, adLockReadOnly
    
        'For i = 0 To cmbCatID.ListCount - 1
    
        '    If cmbCatID.List(i, 1) = ![GroupID] Then
        '        cmbCatID.ListIndex = i
        '        Exit For
        '    End If
    
        'Next i
        cmbCatID.AddItem ![GroupName] & ""
        cmbCatID.List(0, 1) = ![GroupID] & ""
        cmbCatID.ListIndex = 0
        'cmbItems.AddItem ![RMName] & ""
        'cmbItems.List(0, 1) = RMID
        'cmbItems.ListIndex = 0
        
        .Close
    End With
    Set rs = Nothing
    
    For i = 0 To cmbItems.ListCount - 1
    
            If cmbItems.List(i, 1) = RMID Then
                cmbItems.ListIndex = i
                Exit For
            End If
    
    Next i
    
    Me.Show
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
Public Sub EditOrder(TempOrdNo As String)
    
    On Error GoTo err
    strOrderNo = TempOrdNo
    Dim rs As New ADODB.Recordset
    Dim itm As ListItem
    
    With rs
        .Open "Select * From VRMPOrders Where OrderNo='" & strOrderNo & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCust.Text = ![VenderName] & ""
        'cmbCountry = ![Country] & ""
        txtOrderNo = ![OrderNo] & ""
        OrderDT = ![OrderDT] & ""
        
        Do Until .EOF
            Set itm = LV.ListItems.Add(, ![RMID] & "'", .AbsolutePosition)
            'ITM.Tag = ![CompItemID] & "" 'cmbItems.List(cmbItems.ListIndex, 2)
            
            itm.ListSubItems.Add , , ![GroupDesc] & "" '![CatID] & ""
            itm.ListSubItems.Add , , ![RMID1] & ""
            'ITM.ListSubItems.Add(, ![ItemName] & "", ![ItemName] & "").Tag = ![Description] & ""
            itm.ListSubItems.Add , , ![RMName] & ""
            
            itm.SubItems(4) = ![Unit] & ""
            
            itm.SubItems(5) = Format(![Price] & "", "#0.00")
            itm.SubItems(6) = Val(![Qty] & "")
            itm.SubItems(7) = Val(![Qty] & "") * Val(![Price])
   
            .MoveNext
        Loop
        .Close
    End With
    Call GetTotal

    txtQty = 0
    cmbItems.SetFocus
    
    Set rs = Nothing
    Me.Show
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub cmbCatID_Change()
    Call cmbCatID_Click
End Sub

Private Sub cmbCatID_Click()
    Call FillItems
End Sub



Private Sub cmbCust_matched()
    On Error GoTo err
    
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select Supervisor From Processes Where ProcessID=" & cmbCust.ID, con, adOpenForwardOnly, adLockReadOnly
        
        txtOrderNo = .Fields(0) & ""
        .Close
    End With
    
    Dim IssID As String
    IssID = Val(con.Execute("Select Max(right(IssID,3)) From RMVendIss Where IssDT='" & Format(ServerDate, "dd-MMM-yyyy") & "'").Fields(0).value & "")
        
        IssID = "ISRM-" & _
         Format(ServerDate, "ddMMyy") & "-" & Format(Val(IssID) + 1, "000")
    txtIssID = IssID
    
    Set rs = Nothing

    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbCust_UnMatched()
    'cmbCountry.ClearVals
End Sub

Private Sub cmbDesc_Change()
    Call cmbDesc_Click
End Sub

Private Sub cmbDesc_Click()
    If cmbDesc.MatchFound Then
        cmbItems.ListIndex = cmbDesc.ListIndex
    End If
End Sub

Private Sub cmbItems_Change()
    Call cmbItems_Click
End Sub

Private Sub cmbItems_Click()

    If cmbItems.MatchFound Then
        cmbDesc.ListIndex = cmbItems.ListIndex
        txtUnit = cmbItems.List(cmbItems.ListIndex, 2)
        txtPrice = cmbItems.List(cmbItems.ListIndex, 3)
    Else
        txtUnit = ""
        txtPrice = ""
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub cmdAddItem_Click()
    
    On Error GoTo err
    
    If Not cmbItems.MatchFound Or cmbItems = "" Then
        MsgBox "Invalid Item Selected", vbInformation
        Exit Sub
    ElseIf cmbCust.MatchFound = False Then
        MsgBox "Select Process.", vbInformation
        Exit Sub
    ElseIf Val(txtQty) = 0 Then
        MsgBox "Quantity To Order Cannot Be Zero", vbInformation
        Exit Sub
    End If

    Dim itm As ListItem
    Dim Sno As Integer
    If LV.ListItems.Count > 0 Then
        Sno = Val(LV.ListItems(LV.ListItems.Count).Text) + 1
    Else
        Sno = 1
    End If
    
    'Set ITM = LV.ListItems.Add(, cmbItems.List(cmbItems.ListIndex, 1) & "'", Sno)
    
    Set itm = LV.ListItems.Add(, cmbItems.List(cmbItems.ListIndex, 1) & "'", Sno)
    itm.Tag = cmbItems.List(cmbItems.ListIndex, 2)
    
    itm.ListSubItems.Add(, , cmbCust.Text).Tag = cmbCust.ID
    itm.ListSubItems.Add , , cmbCatID
    itm.ListSubItems.Add(, , cmbItems).Tag = cmbItems.List(cmbItems.ListIndex, 1)
    itm.ListSubItems.Add , , cmbDesc
    
    itm.SubItems(5) = txtUnit
    
    'ITM.SubItems(6) = Format(txtPrice, "#0.00")
    itm.SubItems(6) = txtQty
    'ITM.SubItems(7) = lblAmt
   
    'cmbItemID.RemoveItem cmbItemID.ListIndex
    'cmbItem.RemoveItem cmbItem.ListIndex
    
    'Call GetTotal

    txtQty = 0
    cmbItems.SetFocus
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdDelete_Click()
    LV.ListItems.Remove LV.SelectedItem.Index
End Sub

Private Sub cmdSClear_Click()
    If Saved Then
        Unload Me
        Me.Show
    End If
End Sub

Private Sub cmdSClose_Click()
    If Saved Then
        Unload Me
    End If
End Sub


Private Sub Form_Load()
   
    On Error GoTo err

    Call cmbCust.AddVals(con, "Description", "Processes", "ProcessID", " Order BY SNo")
    OrderDT = ServerDate
    
    Me.KeyPreview = True
    cmbCust.ListHeight = 2500
    
    Call FillCmbs
    
    Exit Sub
err:
    MsgBox err.Description
   
End Sub
Private Sub FillCmbs()

    On Error GoTo err
    
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select Distinct GroupID,GroupName From VRM", con, adOpenForwardOnly, adLockReadOnly
        cmbCatID.Clear
        Do Until .EOF
            cmbCatID.AddItem .Fields(1) & ""
            cmbCatID.List(cmbCatID.ListCount - 1, 1) = .Fields(0) & ""
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing

    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Resize()
On Error Resume Next
   FALL.Move ((Me.ScaleWidth - FALL.Width) / 2), ((Me.ScaleHeight - FALL.Height) / 2)
End Sub

Private Sub LV_Click()
    If LV.ListItems.Count > 0 Then
        Call LV_ItemClick(LV.SelectedItem)
    End If
End Sub

Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With cmdDelete
        .Move .Left, LV.Top + Item.Top
        .Visible = True
    End With
End Sub

Private Sub txtPrice_Change()
   lblAmt = Format(Val(txtQty) * Val(txtPrice), "#0.00")
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
   KeyAscii = OnlyNumber(KeyAscii, True)
End Sub

Private Sub txtQty_Change()
   lblAmt = Format(Val(txtQty) * Val(txtPrice), "#0.00")
End Sub

Private Sub txtQty_GotFocus()
cmdAddItem.Default = True
End Sub
Private Sub txtQty_LostFocus()
cmdAddItem.Default = False
End Sub
Private Sub txtQty_KeyPress(KeyAscii As Integer)
   KeyAscii = OnlyNumber(KeyAscii)
End Sub


Private Sub GetTotal()

    Dim i As Integer
    Dim total As Double
    total = 0

    For i = 1 To LV.ListItems.Count
        total = total + Val(LV.ListItems(i).ListSubItems(7).Text)
    Next i

    lblTotal = Format(total, "#0.00")
    lblTotWords = ConvertInWords(lblTotal, "Rs")
    
End Sub



Private Function Saved() As Boolean
    
    Saved = False
    'If txtOrderNo = "" Then
    '    MsgBox "Please Enter Order No.", vbInformation
    '    txtOrderNo.SetFocus
    '    Exit Function
    'End If
    
    If LV.ListItems.Count = 0 Then
        MsgBox "No Items Added to This Order", vbInformation
        Exit Function
    End If
    
    Call StartTrans(con)
    '----------------------------------------------------------------
    'If strOrderNo <> "" Then
    '    con.Execute "Delete from RMPOrders Where OrderNo='" & _
    '     strOrderNo & "'"
    'End If
    
    'Save Order
    
    con.Execute "Insert Into RMVendIss (IssDT,IssID)" & _
     " values('" & OrderDT & "','" & txtIssID & "')"
    
    Dim iRefID As Long
    iRefID = Val(con.Execute("Select Max(ID) From RMVendIss").Fields(0).value & "")
    'Now Save Items
    For i = 1 To LV.ListItems.Count
        With LV.ListItems(i)
            con.Execute "Insert into RMVendIssItems (RefID,ProcessID," & _
             "RMID,Qty) Values (" & iRefID & "," & _
             Val(.ListSubItems(1).Tag) & "," & _
             Val(LV.ListItems(i).ListSubItems(3).Tag) & _
             "," & Val(LV.ListItems(i).SubItems(6)) & ")"
             
             con.Execute "Update RM Set QtyInStock=QtyInStock-" & _
              Val(LV.ListItems(i).SubItems(6)) & _
              " Where RMID=" & Val(.ListSubItems(3).Tag)
        End With
    Next
      
    con.CommitTrans
    
    Saved = True
    Exit Function

err:
    MsgBox err.Description
End Function
Private Sub FillItems()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    With rs
    
        .Open "Select * From VRM Where GroupID=" & cmbCatID.List(cmbCatID.ListIndex, 1), con, adOpenForwardOnly, adLockReadOnly
        cmbItems.Clear
        cmbDesc.Clear
        
        
        Do Until .EOF
        
            cmbItems.AddItem ![RMID1] & ""
            cmbDesc.AddItem ![RMName] & ""
            
            cmbItems.List(cmbItems.ListCount - 1, 1) = ![RMID] & ""
            cmbItems.List(cmbItems.ListCount - 1, 2) = ![Unit] & ""
            cmbItems.List(cmbItems.ListCount - 1, 3) = ![Rate] & ""
            
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
