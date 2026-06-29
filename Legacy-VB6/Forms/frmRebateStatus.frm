VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRebateStatus 
   ClientHeight    =   10530
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
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   10530
   ScaleMode       =   0  'User
   ScaleWidth      =   10641.08
   WindowState     =   2  'Maximized
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
      TabIndex        =   7
      Top             =   0
      Width           =   10560
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rebate Status"
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
         TabIndex        =   8
         Top             =   165
         Width           =   10605
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Rebate Status"
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
         TabIndex        =   9
         Top             =   180
         Width           =   10590
      End
   End
   Begin VB.Frame FAcc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   9765
      Left            =   15
      TabIndex        =   3
      Top             =   615
      Width           =   11850
      Begin VB.TextBox txtTotalBalance 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   8820
         Width           =   1185
      End
      Begin VB.TextBox txtTotalAmt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   330
         Left            =   4230
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   8820
         Width           =   1185
      End
      Begin VB.ComboBox cmbRange 
         Height          =   315
         ItemData        =   "frmRebateStatus.frx":0000
         Left            =   2070
         List            =   "frmRebateStatus.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   345
         Width           =   2010
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   5820
         TabIndex        =   1
         Top             =   345
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   556
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
         Format          =   107937795
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   4095
         TabIndex        =   0
         Top             =   345
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   556
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
         Format          =   107937795
         CurrentDate     =   37055
      End
      Begin MSComctlLib.ListView LV 
         Height          =   8085
         Left            =   60
         TabIndex        =   10
         Top             =   750
         Width           =   11715
         _ExtentX        =   20664
         _ExtentY        =   14261
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
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
            Text            =   "S. #"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Invoice #"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Appd. DT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Balance"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3480
         TabIndex        =   12
         Top             =   8850
         Width           =   525
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   375
         Left            =   10425
         TabIndex        =   11
         Top             =   9285
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "2355;661"
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbStatus 
         Height          =   285
         Left            =   180
         TabIndex        =   6
         Top             =   345
         Width           =   1860
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3281;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   7395
         TabIndex        =   2
         Top             =   285
         Width           =   1500
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "2646;635"
         Accelerator     =   82
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         Caption         =   " Status                              Range                                 From                            To"
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
         Left            =   165
         TabIndex        =   4
         Top             =   120
         Width           =   7230
      End
   End
End
Attribute VB_Name = "frmRebateStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InvsForPrint As String

Private Sub cmbRange_Click()

    DtTo = ServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
        Case 0
            
        Case 1
            DtFrom = ServerDate
        Case 2
            DtFrom = DateAdd("d", -7, ServerDate)
        Case 3
            DtFrom = DateAdd("d", -15, ServerDate)
        Case 4
            DtFrom = DateAdd("d", -30, ServerDate)
        Case 5
            DtFrom.Enabled = True
            DtTo.Enabled = True
            DtFrom.SetFocus
            Exit Sub
    End Select
    
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub



Private Sub DtFrom_Change()
    'Call setDateRange(DtFrom)
End Sub

Private Sub DtTo_Change()
    'Call setDateRange(DtTo)
End Sub


Private Sub Form_Load()
    
    
    
    cmbRange.ListIndex = 0
    
    
    With cmbStatus
        .AddItem "<Receiveables>"
        .AddItem "<Received>"
        .AddItem "<Not Applied>"
        .AddItem "<All Rebates>"
    End With
    
    cmbStatus.ListIndex = 0
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Dim TheHeight As Long
    Frame1.Left = (Me.ScaleWidth - Frame1.Width) / 2
    
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    
End Sub


Private Sub RefreshLV()
    
    On Error GoTo err
    
    Dim Cond As String
    Dim DTRange As String
    Dim ITM As ListItem
    
    
    Screen.MousePointer = vbHourglass
            
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_RebateStatus"
        
        .Parameters("@Status").Value = cmbStatus.ListIndex
        
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        Set rs = .Execute
    End With
    
    If cmbRange.ListIndex = 0 Then
        DTRange = " DT Between '" & DateAdd("d", -15, ServerDate) & "' and '" & ServerDate & "'"
    ElseIf cmbRange.ListIndex = 1 Then
        DTRange = " DT Between '" & DateAdd("d", -30, ServerDate) & "' and '" & ServerDate & "'"
    ElseIf cmbRange.ListIndex = 2 Then
        DTRange = " DT Between '" & DateAdd("d", -60, ServerDate) & "' and '" & ServerDate & "'"
    ElseIf cmbRange.ListIndex = 3 Then
        DTRange = " DT Between '" & DateAdd("d", -90, ServerDate) & "' and '" & ServerDate & "'"
    ElseIf cmbRange.ListIndex = 4 Then
        DTRange = " DT Between '" & DtFrom & "' and '" & DtTo & "'"
    End If
    
    If Cond = "" Then
        Cond = " Where "
    Else
        Cond = Cond & " and "
    End If
    Cond = Cond & DTRange
    
    Dim dTotalAmt As Double, dTotalBalance As Double
    Dim BalAmt As Double
    With rs
        LV.ListItems.Clear
        InvsForPrint = "In["
        Do Until .EOF
            If ![EntryType] = 0 Then
                
                BalAmt = Round(Val(![Balance] & ""), 3)
                
                Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
                ITM.ListSubItems.add , , ![CInvoice] & ""
                
                If IsNull(![DT]) Then
                    ITM.ListSubItems.add , , "N/A"
                    ITM.ListSubItems.add , , "N/A"
                    ITM.ListSubItems.add , , "N/A"
                Else
                    ITM.ListSubItems.add , , Format(![DT] & "", "dd-MMM-yyyy")
                    ITM.ListSubItems.add , , Round(Val(![TotalAmt] & ""))
                    ITM.ListSubItems.add , , ![Balance] & ""
                    dTotalAmt = dTotalAmt + Val(!TotalAmt & "")
                    dTotalBalance = dTotalBalance + Val(!Balance & "")
                End If
                LColor = LV.ForeColor
                InvsForPrint = InvsForPrint & "'" & ![CustomInvoice] & "',"
            Else
                Set ITM = LV.ListItems.add(, , "")
                ITM.ListSubItems.add , , ""
                
                ITM.ListSubItems.add , , Format(![DT] & "", "dd-MMM-yyyy")
                ITM.ListSubItems(2).Tag = ![EntryID] & ""
                ITM.ListSubItems.add , , Round(Val(![TotalAmt] & ""))
                ITM.ListSubItems.add , , ![Description] & ""
                
                LColor = vbBlue
            End If
            
            ITM.ListSubItems(1).Tag = ![CustomInvoice] & ""
            ITM.ListSubItems(4).Tag = BalAmt
            
            ITM.ForeColor = LColor
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ToolTipText = ITM.ListSubItems(i).Text
                ITM.ListSubItems(i).ForeColor = LColor
                
                If ITM.Text <> "" Then
                    ITM.ListSubItems(4).ForeColor = vbRed
                End If
            Next i
            
            .MoveNEXT
        Loop
        .Close
    End With
        
    txtTotalAmt = Format(dTotalAmt, "#,##")
    txtTotalBalance = Format(dTotalBalance, "#,##")
    
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub


Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button <> 2 Then Exit Sub
    
    Dim strRet As String
        
    If LV.ListItems.count = 0 Then
        strRet = ShowPopUpMenu1("&Refresh")
    ElseIf Val(LV.SelectedItem.ListSubItems(4).Tag) > 0 Then
        strRet = ShowPopUpMenu1("&Refresh", "-", "&Receive Payments For This Rebate", "-", "&Print Payment Status For This Rebate", "-", "&Print Payment Status Of All")
    Else
        strRet = ShowPopUpMenu1("&Refresh", "-", "&Print Payment Status For This Rebate", "-", "&Print Payment Status Of All")
    End If

    'strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "&Foreign Customers", "-", "Customer Catalog", "-", "Customer &Ledger", "-", "Assign &FDBC No.", "-", "&Receive Payments")
    Dim f As Form
    
    Select Case strRet
        Case "&Refresh"
            Call RefreshLV
        Case "&Receive Payments For This Rebate"
        
            Set f = New frmRcvRebates
            Load f
            Call f.ShowForm(LV.SelectedItem.ListSubItems(1).Tag)
            
        Case "&Print Payment Status For This Rebate"
        
            Screen.MousePointer = vbHourglass
            Dim rpt As CRAXDDRT.Report 'New rptRebateStatus
            Set rpt = rptApp.OpenReport(RptPath & "\rptRebateStatus.rpt") '
            Load frmPrevRpt
            frmPrevRpt.ShowReport "{CustomInvoice.CustomInvoice}='" & LV.SelectedItem.ListSubItems(1).Tag & "' ", rpt
            'frmPrevRpt.ShowReport
            Screen.MousePointer = vbDefault
            
        Case "&Print Payment Status Of All"
        
            Screen.MousePointer = vbHourglass
            
            Load frmPrevRpt
            frmPrevRpt.ShowReport "{CustomInvoice.CustomInvoice} " & Left(InvsForPrint, Len(InvsForPrint) - 1) & "] ", rpt
            'frmPrevRpt.ShowReport
            Screen.MousePointer = vbDefault
            
    End Select
    
End Sub
