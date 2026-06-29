VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExportStats 
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   10530
   ScaleMode       =   0  'User
   ScaleWidth      =   10641.08
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10560
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Export Statistics"
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
         Caption         =   "Export Statistics"
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
      ForeColor       =   &H00800000&
      Height          =   9765
      Left            =   15
      TabIndex        =   3
      Top             =   615
      Width           =   11850
      Begin VB.TextBox txtFinished 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5085
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   8970
         Width           =   1725
      End
      Begin VB.TextBox txtInVis 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1395
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   8940
         Width           =   1725
      End
      Begin VB.TextBox txtVis 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5100
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   8580
         Width           =   1725
      End
      Begin VB.TextBox txtTotQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10035
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   8505
         Width           =   1725
      End
      Begin VB.TextBox txtForg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   8550
         Width           =   1725
      End
      Begin VB.ComboBox cmbRange 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmExportStats.frx":0000
         Left            =   2070
         List            =   "frmExportStats.frx":0016
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
         Format          =   45678595
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
         Format          =   45678595
         CurrentDate     =   37055
      End
      Begin MSComctlLib.ListView LV 
         Height          =   7695
         Left            =   60
         TabIndex        =   10
         Top             =   750
         Width           =   11715
         _ExtentX        =   20664
         _ExtentY        =   13573
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
            Text            =   "Cust Code"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item Code"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Item Name"
            Object.Width           =   10478
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Qty"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   375
         Left            =   9015
         TabIndex        =   22
         Top             =   9285
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Print      "
         PicturePosition =   327683
         Size            =   "2355;661"
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Total Finished :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   3930
         TabIndex        =   20
         Top             =   9015
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Waste (Invisible) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   75
         TabIndex        =   19
         Top             =   8970
         Width           =   1320
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Total Qty :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   15
         Left            =   9180
         TabIndex        =   16
         Top             =   8550
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Total Waste(Visible) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   3495
         TabIndex        =   14
         Top             =   8625
         Width           =   1530
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Total Forging :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   345
         TabIndex        =   13
         Top             =   8580
         Width           =   1050
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
      Begin MSForms.ComboBox cmbCust 
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
         Caption         =   " Customer                        Range                                 From                            To"
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
Attribute VB_Name = "frmExportStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InvsForPrint As String
Dim SqlQuery As String

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

Private Sub cmdPrint_Click()
    
    On Error GoTo err
    Dim F As New frmPrevRpt
    Dim rpt As craxddrt.Report 'New rptExportStats
    Set rpt = rptApp.OpenReport(RptPath & "\rptExportStats.rpt") '
    Dim Formulas As craxddrt.FormulaFieldDefinitions
    Dim rptFormulas As craxddrt.FormulaFieldDefinition
    
    Set Formulas = rpt.FormulaFields
    For Each rptFormulas In Formulas
        If rptFormulas.Name = "{@TotalForging}" Then
            rptFormulas.Text = Val(txtForg)
        ElseIf rptFormulas.Name = "{@TotalWasteInVis}" Then
            rptFormulas.Text = Val(txtInVis)
        ElseIf rptFormulas.Name = "{@TotalWasteVis}" Then
            rptFormulas.Text = Val(txtVis)
        ElseIf rptFormulas.Name = "{@TotalFinished}" Then
            rptFormulas.Text = Val(txtFinished)
        End If
        Debug.Print rptFormulas.Name
    Next
    rpt.SQLQueryString = SqlQuery
    F.ShowReport " ", rpt
    Exit Sub
err:
    MsgBox err.Description
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
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Distinct CustCode From ForeignCustomers", con, adOpenForwardOnly, adLockReadOnly
        cmbCust.Clear
        cmbCust.AddItem "<All Customers>"
        Do Until .EOF
            cmbCust.AddItem .Fields(0) & ""
            .MoveNext
        Loop
        .Close
    End With
    If cmbCust.ListCount > 0 Then cmbCust.ListIndex = 0
    Set rs = Nothing
    cmbRange.ListIndex = 0
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Dim TheHeight As Long
    Frame1.Left = (Me.ScaleWidth - Frame1.Width) / 2
    
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    
End Sub


Private Sub RefreshLV()
    
    On Error GoTo err
    
    If cmbCust.MatchFound = False Then Exit Sub
    Dim rs As New ADODB.Recordset
    
    Dim Cond As String
    Dim DTRange As String
    Dim itm As ListItem
    
    
    Screen.MousePointer = vbHourglass
    
    If cmbCust.ListIndex <> 0 Then
        Cond = " Where CustCode='" & cmbCust.Text & "' "
    End If
    
    DTRange = ""
    If cmbRange.ListIndex = 1 Then
        'DTRange = " DT Between '" & DateAdd("d", -15, ServerDate) & "' and '" & ServerDate & "'"
    ElseIf cmbRange.ListIndex = 2 Then
        DTRange = " DT Between '" & DateAdd("d", -30, ServerDate) & "' and '" & ServerDate & "'"
    ElseIf cmbRange.ListIndex = 3 Then
        DTRange = " DT Between '" & DateAdd("d", -60, ServerDate) & "' and '" & ServerDate & "'"
    ElseIf cmbRange.ListIndex = 4 Then
        DTRange = " DT Between '" & DateAdd("d", -90, ServerDate) & "' and '" & ServerDate & "'"
    ElseIf cmbRange.ListIndex = 5 Then
        DTRange = " DT Between '" & DtFrom & "' and '" & DtTo & "'"
    End If
    
    If DTRange <> "" Then
        If Cond = "" Then
            Cond = " Where "
        Else
            Cond = Cond & " and "
        End If
    
        Cond = Cond & DTRange
    End If
        
    Dim TotQty As Long
    Dim TotForg As Double
    Dim TotVis As Double
    Dim TotFin As Double
    Dim TotInVis As Double
    TotQty = 0
    Dim lQty As Long
    
    SqlQuery = "Select CustCode,ItemCode,CompleteItemName,Sum(Qty) As TotalQty," & _
         "Sum(ForgingWeight) As TotalForging,Sum(WasteVisible) AS TotalVisible," & _
         "Sum(FinishedWeight) As TotalFinished,Sum(WasteInvisible) As TotalInVisible" & _
         " From VExportStats " & Cond & " Group By CustCode,ItemCode,CompleteItemName Order By CustCode,ItemCode"
         
    With rs
        .Open SqlQuery, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        
        Do Until .EOF
            
            Set itm = LV.ListItems.Add(, , LV.ListItems.Count + 1)
            itm.SubItems(1) = ![CustCode] & ""
            itm.SubItems(2) = ![ItemCode] & ""
            itm.SubItems(3) = ![CompleteItemName] & ""
            itm.SubItems(4) = ![TotalQty]
            lQty = Val(![TotalQty])
            TotQty = TotQty + Val(![TotalQty] & "")
            TotForg = TotForg + (Val(![TotalForging] & "") * lQty)
            TotVis = TotVis + (Val(![TotalVisible] & "") * lQty)
            TotFin = TotFin + (Val(![TotalFinished] & "") * lQty)
            TotInVis = TotInVis + (Val(![TotalInVisible] & "") * lQty)
            .MoveNext
        Loop
        .Close
    End With
    txtTotQty = TotQty
    txtForg = (TotForg) / 1000
    txtVis = (TotVis) / 1000
    txtInVis = (TotInVis) / 1000
    txtFinished = (TotFin) / 1000
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub


Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button <> 2 Then Exit Sub
    
    Dim strRet As String
        
    If LV.ListItems.Count = 0 Then
        strRet = ShowPopUpMenu1("&Refresh")
    ElseIf Val(LV.SelectedItem.ListSubItems(4).Tag) > 0 Then
        strRet = ShowPopUpMenu1("&Refresh", "-", "&Receive Payments For This Rebate", "-", "&Print Payment Status For This Rebate", "-", "&Print Payment Status Of All")
    Else
        strRet = ShowPopUpMenu1("&Refresh", "-", "&Print Payment Status For This Rebate", "-", "&Print Payment Status Of All")
    End If

    'strRet = ShowPopUpMenu(CoolBar1.Width, MenuTop, "&Foreign Customers", "-", "Customer Catalog", "-", "Customer &Ledger", "-", "Assign &FDBC No.", "-", "&Receive Payments")
    Dim F As Form
    
    Select Case strRet
        Case "&Refresh"
            Call RefreshLV
        Case "&Receive Payments For This Rebate"
        
            Set F = New frmRcvRebates
            Load F
            Call F.ShowForm(LV.SelectedItem.ListSubItems(1).Tag)
            
        Case "&Print Payment Status For This Rebate"
        
            Screen.MousePointer = vbHourglass
            Dim rpt As craxddrt.Report 'New rptRebateStatus
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
