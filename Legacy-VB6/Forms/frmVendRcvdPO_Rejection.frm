VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmVendRcvdPO_Rejection 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rejection"
   ClientHeight    =   6645
   ClientLeft      =   1095
   ClientTop       =   285
   ClientWidth     =   11910
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00E7EBEF&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   7860
      ScaleHeight     =   285
      ScaleWidth      =   2355
      TabIndex        =   13
      Top             =   2085
      Visible         =   0   'False
      Width           =   2385
      Begin VB.TextBox txtQty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -15
         TabIndex        =   15
         Top             =   -15
         Width           =   1200
      End
      Begin VB.TextBox txtSheets 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1170
         TabIndex        =   14
         Top             =   -15
         Width           =   1200
      End
   End
   Begin VB.Frame Frame6 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7845
      TabIndex        =   7
      Top             =   5610
      Width           =   3630
      Begin MSForms.CommandButton cmdCancel 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   1830
         TabIndex        =   2
         Top             =   210
         Width           =   1665
         Caption         =   "Close"
         PicturePosition =   327683
         Size            =   "2937;635"
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
         Left            =   105
         TabIndex        =   1
         Top             =   210
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
   Begin MSComctlLib.ListView LV 
      Height          =   4050
      Left            =   420
      TabIndex        =   0
      Top             =   1455
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   7144
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Material"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Rcvd Qty"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Sheet/Bar"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Rej. Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Rej. Bars"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Rej."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Sheet/Bar"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Frame fTitle 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   15
      TabIndex        =   3
      Top             =   30
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rejection of Received RM"
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
         Left            =   3525
         TabIndex        =   4
         Top             =   120
         Width           =   4005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rejection of Received RM"
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
         Left            =   3540
         TabIndex        =   5
         Top             =   135
         Width           =   4005
      End
   End
   Begin MSComCtl2.DTPicker DT 
      Height          =   300
      Left            =   9825
      TabIndex        =   8
      Top             =   1170
      Width           =   1665
      _ExtentX        =   2937
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
      Format          =   41746435
      CurrentDate     =   -8461
   End
   Begin MSForms.TextBox txtPONo 
      Height          =   285
      Left            =   5145
      TabIndex        =   12
      Top             =   1185
      Width           =   1695
      VariousPropertyBits=   746604575
      BorderStyle     =   1
      Size            =   "2990;503"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtVender 
      Height          =   285
      Left            =   2145
      TabIndex        =   11
      Top             =   1185
      Width           =   3015
      VariousPropertyBits=   746604575
      BorderStyle     =   1
      Size            =   "5318;503"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtDesc 
      Height          =   285
      Left            =   6825
      TabIndex        =   10
      Top             =   1185
      Width           =   3015
      VariousPropertyBits=   746604575
      BorderStyle     =   1
      Size            =   "5318;503"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox txtRcvNo 
      Height          =   285
      Left            =   420
      TabIndex        =   9
      Top             =   1185
      Width           =   1755
      VariousPropertyBits=   746604575
      BorderStyle     =   1
      Size            =   "3096;503"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label3 
      Height          =   240
      Left            =   420
      TabIndex        =   6
      Top             =   960
      Width           =   11055
      BackColor       =   11517387
      Caption         =   $"frmVendRcvdPO_Rejection.frx":0000
      Size            =   "19500;423"
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
Attribute VB_Name = "frmVendRcvdPO_Rejection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strRcvNo As String

Private Sub cmbOrderNo_matched()

    On Error GoTo err
    Dim LastGroup As String
    Dim rs As New ADODB.Recordset
    
    With rs
        
        .Open "SELECT * FROM VVendOrdersToRcv WHERE PONo='" & cmbOrderNo.id & "' ORDER BY GroupID,MaterialName", con, adOpenForwardOnly, adLockReadOnly
        Dim ITM As ListItem
        LV.ListItems.Clear
        
        Do Until .EOF
            
            Set ITM = LV.ListItems.add(, , "")
            ITM.Tag = 0 'Val(![WORef] & "")
            
            'Display the Customer Order No if chkOrderNo is checked Else Show the Company Order No
            ITM.ListSubItems.add(, , ![MaterialID]).Tag = Val(!Rate & "")
            ITM.ListSubItems.add , , ![MaterialName] & ""
            ITM.ListSubItems.add(, , Val(![QtyOrdered]) & " " & ![Unit]).Tag = Abs(!CCItem)
            ITM.ListSubItems.add , , Val(![QtyToRcv]) & " " & ![Unit]
            ITM.ListSubItems.add().Tag = ![Unit] & ""
            ITM.ListSubItems.add().Tag = ""
            ITM.ListSubItems.add(, , Val(!Rate & "")).Tag = Val(!Rate & "")
            'Set Tool Tips
            
            ITM.ListSubItems(4).Tag = Val(!OtherUnitEntryID & "")
            
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ToolTipText = ITM.ListSubItems(i).Text
            Next
            
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    cmbMaterial.ClearVals
    cmbMaterial.AddVals con, "'{' + MaterialID + '} ' + MaterialName", "VVendOrdersToRcv", "MaterialID", " WHERE PONo='" & cmbOrderNo.id & "' ORDER BY GroupID,MaterialName"
    
    
Exit Sub

err:
   MsgBox err.Description
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

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub

    For i = 1 To LV.ColumnHeaders.count
        LV.ColumnHeaders(i).Width = Val(LV.ColumnHeaders(i).Tag)
    Next
    
    If Val(LV.SelectedItem.SubItems(2)) > 0 Then
        txtSheets.Enabled = True
    Else
        txtSheets.Enabled = False
    End If
    
    With PicEdit
        .Move LV.Left + LV.ColumnHeaders(6).Left, LV.Top + LV.SelectedItem.Top
        .Visible = True
    End With
    
    txtSheets.Text = Val(LV.SelectedItem.SubItems(6))
    
    With txtQty
        .Text = Val(LV.SelectedItem.SubItems(5))
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With

End Sub

Private Sub Form_Load()

    DT = GetServerDate
    
    
    For i = 1 To LV.ColumnHeaders.count
        LV.ColumnHeaders(i).Tag = LV.ColumnHeaders(i).Width
    Next
    
    
    strRcvNo = ""
        
    'lShelfNoForAutoPlacement = GetSingleLongValue("MAX(EntryID)", "StoreShelfs")
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
 
    Dim TitleBottom As Long

    TitleBottom = txtVender.Top + txtVender.Height 'fTitle.Top + fTitle.Height
 
    LV.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + Frame6.Height + 100)
 
    Frame6.Move Me.ScaleWidth - Frame6.Width - 50, Me.ScaleHeight - Frame6.Height - 50

    'chkPrint.Move LV.Left, Frame6.Top + 50

 
End Sub

Private Function Saved() As Boolean

    On Error GoTo err
    Dim i As Integer, lEntryID As Long
    Call StartTrans(con)
    
    con.Execute "INSERT INTO VendRcvdDetailPO_Rejection(DT,UserName,MachineName,OrderNo) VALUES('" & DT.Value & "','" & _
     CurrentUserName & "','" & strComputerName & "','" & txtPONo.Text & "')"
    
    lEntryID = GetSingleLongValue("MAX(EntryID)", "VendRcvdDetailPO_Rejection", " WHERE MachineName='" & strComputerName & "'")
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If Val(.SubItems(5)) > 0 Then
                con.Execute "INSERT INTO VendRcvdDetailPO_Rejection_Detail(VRDPO_Rej_RefID,VRDPO_RefID,QtyRejected,SheetsRejected)" & _
                 " VALUES(" & lEntryID & "," & Val(.Tag) & "," & Val(.SubItems(5)) & "," & Val(.SubItems(6)) & ")"
            End If
        End With
    Next
    con.CommitTrans
    
    Unload Me
    Exit Function
err:
    MsgBox err.Description
End Function

Public Sub ShowMe(p_strRcvID As String)

    strRcvNo = p_strRcvID
    txtRcvNo = strRcvNo
    
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
    
        strSQL = "SELECT * FROM VMaterialsForPlacement WHERE RcvID='" & strRcvNo & "' ORDER BY EntryID"
        .Open strSQL, con, adOpenForwardOnly, adLockReadOnly
            
        txtVender = !AccTitle & ""
        txtPONo = !PORefNo & ""
        
        LV.ListItems.Clear
        
        Do Until .EOF
            
            Set ITM = LV.ListItems.add(, , !MaterialID & " " & !RMName & "")
            ITM.Tag = !EntryID & ""
            
            ITM.ListSubItems.add , , Round(Val(![QtyPassed] & "") - Val(![QtyPlaced] & ""), 2)
            
            ITM.ListSubItems.add , , Round(Val(![Sheets] & "") - Val(![SheetsPlaced] & ""), 2)
            
            ITM.ListSubItems.add , , Round(Val(![QtyRejected] & ""), 2)
            ITM.ListSubItems.add , , Round(Val(![SheetsRejected] & ""), 2)
            
            ITM.ListSubItems.add , , ""
            
            ITM.ListSubItems.add , , ""
            
            .MoveNext
        Loop
        .Close
        
    End With
    
    Me.Show
    
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(txtQty) > Val(LV.SelectedItem.SubItems(1)) Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        End If
        
        If txtSheets.Enabled = False Then
            LV.SelectedItem.SubItems(5) = Val(txtQty)
            
            LV.SetFocus
            
            PicEdit.Visible = False
        Else
        
            With txtSheets
                .SelStart = 0
                .SelLength = Len(.Text)
                .SetFocus
            End With
            
        End If
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub

Private Sub txtSheets_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(txtSheets) > Val(LV.SelectedItem.SubItems(2)) Then
            MsgBox "Invalid Sheets.", vbInformation
            Exit Sub
        End If
        If txtSheets.Enabled = False Then
            ITM.SubItems(5) = Val(txtQty)
            ITM.SubItems(6) = Val(txtSheets)
            LV.SetFocus
            PicEdit.Visible = False
        Else
            txtSheets.SetFocus
        End If
        
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub
