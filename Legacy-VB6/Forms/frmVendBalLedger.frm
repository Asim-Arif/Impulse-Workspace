VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmVendBalLedger 
   ClientHeight    =   10965
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11880
   ClipControls    =   0   'False
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
   ScaleHeight     =   10965
   ScaleMode       =   0  'User
   ScaleWidth      =   10641.08
   WindowState     =   2  'Maximized
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   75
      TabIndex        =   2
      Top             =   945
      Width           =   2220
      _ExtentX        =   3916
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
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Top             =   945
      Width           =   3090
      _ExtentX        =   5450
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   9135
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   16113
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Issue Detail"
      TabPicture(0)   =   "frmVendBalLedger.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LVIssd"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtIssTotal"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Recieve Detail"
      TabPicture(1)   =   "frmVendBalLedger.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LVRcvd"
      Tab(1).Control(1)=   "txtRcvdTotal"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Balances"
      TabPicture(2)   =   "frmVendBalLedger.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtBalTotal"
      Tab(2).Control(1)=   "LV"
      Tab(2).Control(2)=   "LVPrint"
      Tab(2).ControlCount=   3
      Begin VB.TextBox txtBalTotal 
         Height          =   375
         Left            =   -64740
         TabIndex        =   16
         Top             =   8655
         Width           =   1380
      End
      Begin VB.TextBox txtRcvdTotal 
         Height          =   375
         Left            =   -64740
         TabIndex        =   15
         Top             =   8640
         Width           =   1380
      End
      Begin VB.TextBox txtIssTotal 
         Height          =   375
         Left            =   10260
         TabIndex        =   14
         Top             =   8670
         Width           =   1380
      End
      Begin MSComctlLib.ListView LV 
         Height          =   8175
         Left            =   -74940
         TabIndex        =   10
         Top             =   390
         Width           =   11580
         _ExtentX        =   20426
         _ExtentY        =   14420
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15199215
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S #"
            Object.Width           =   2363
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item ID"
            Object.Width           =   3151
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item Name"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Balance"
            Object.Width           =   3151
         EndProperty
      End
      Begin MSComctlLib.ListView LVRcvd 
         Height          =   8175
         Left            =   -74940
         TabIndex        =   11
         Top             =   390
         Width           =   11580
         _ExtentX        =   20426
         _ExtentY        =   14420
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15199215
         Appearance      =   1
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
            Text            =   "Rec No"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Item Name"
            Object.Width           =   11719
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Qty"
            Object.Width           =   1411
         EndProperty
      End
      Begin MSComctlLib.ListView LVIssd 
         Height          =   8175
         Left            =   60
         TabIndex        =   12
         Top             =   390
         Width           =   11580
         _ExtentX        =   20426
         _ExtentY        =   14420
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15199215
         Appearance      =   1
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
            Text            =   "Rec No"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Item Name"
            Object.Width           =   11719
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Qty"
            Object.Width           =   1411
         EndProperty
      End
      Begin MSComctlLib.ListView LVPrint 
         Height          =   1995
         Left            =   -74040
         TabIndex        =   17
         Top             =   6930
         Visible         =   0   'False
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   3519
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   15199215
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S #"
            Object.Width           =   2363
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item ID"
            Object.Width           =   3151
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item Name"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Balance"
            Object.Width           =   3151
         EndProperty
      End
   End
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11520
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vender Ledger"
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
         Index           =   1
         Left            =   4440
         TabIndex        =   7
         Top             =   120
         Width           =   2340
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vender Ledger"
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
         Index           =   2
         Left            =   4455
         TabIndex        =   8
         Top             =   135
         Width           =   2340
      End
   End
   Begin MSComCtl2.DTPicker DTMonth 
      Height          =   300
      Left            =   5385
      TabIndex        =   13
      Top             =   930
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "MMMM, yyyy"
      Format          =   60227587
      CurrentDate     =   38334
   End
   Begin MSForms.CommandButton cmdPrintSlip 
      Height          =   360
      Left            =   7455
      TabIndex        =   18
      Top             =   10545
      Width           =   1425
      ForeColor       =   0
      Caption         =   "Print Reciepts"
      PicturePosition =   327683
      Size            =   "2514;635"
      Accelerator     =   101
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   360
      Left            =   8925
      TabIndex        =   5
      Top             =   10545
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
   Begin MSForms.CommandButton cmdClose 
      Height          =   360
      Left            =   10395
      TabIndex        =   4
      Top             =   10545
      Width           =   1425
      ForeColor       =   0
      Caption         =   " Close      "
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
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   90
      TabIndex        =   1
      Top             =   675
      Width           =   7065
      BackColor       =   11517387
      Caption         =   "  Process                                Venders"
      Size            =   "12462;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   7275
      TabIndex        =   0
      Top             =   960
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
End
Attribute VB_Name = "frmVendBalLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
    NewContEntry.Show 1
End Sub

Private Sub PrintBalRpt()
    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    
    Load frmPrevRpt
    
    'rptPurchOrdForSteel.FormulaFields(5).Text = Val(LV.SelectedItem.SubItems(5) & "")
    'rptPurchOrdForSteel.FormulaFields(2).Text = "'" & strFromTo & "'"
    Dim rpt As craxddrt.Report 'New rptVendLedger
    Set rpt = rptApp.OpenReport(RptPath & "\rptVendLedger.rpt") '
    'rpt.FormulaFields(4).Text = "' From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
    Dim FFDefs As FormulaFieldDefinitions
    Dim FFDef As craxddrt.FormulaFieldDefinition
    Set FFDefs = rpt.FormulaFields
    
    For Each FFDef In FFDefs
        If FFDef.Name = "{@FromTo}" Then
            FFDef.Text = "' From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
            Exit For
        End If
    Next
    
    'rpt.FormulaFields("@FromTo").Text = "' From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
    'rpt.SetReportVariableValue
    frmPrevRpt.ShowReport "{VLedger.VendID}=" & cmbEmp.ID & " and ({VLedger.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & "))" & " ", rpt
    
    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub PrintIssRcvStatus()
    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    
    Load frmPrevRpt
    
    
    Dim rpt As craxddrt.Report 'New rptVendIssRcvStatus
    Set rpt = rptApp.OpenReport(RptPath & "\rptVendIssRcvStatus.rpt") '
    
    Dim FFDefs As FormulaFieldDefinitions
    Dim FFDef As craxddrt.FormulaFieldDefinition
    Set FFDefs = rpt.FormulaFields
    
    For Each FFDef In FFDefs
        If FFDef.Name = "{@FromTo}" Then
            FFDef.Text = "' From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
        ElseIf FFDef.Name = "{@EmpName}" Then
            FFDef.Text = "'" & cmbEmp.Text & "'"
        End If
    Next
    
    Dim subRpt As Report
    Set subRpt = rpt.OpenSubreport("ContRcvdItems")
    subRpt.RecordSelectionFormula = "{VLedger.VendID}=" & cmbEmp.ID & " and ({VLedger.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & "))" & " "
    
    frmPrevRpt.ShowReport "{VLedger.VendID}=" & cmbEmp.ID & " and ({VLedger.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & "))" & " ", rpt
    
    Screen.MousePointer = vbDefault

    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Sub cmdPrint_Click()

'    Dim strRet As String
'    strRet = ShowPopUpMenu1("Print &Balance Report", "-", "Print &Issue Receive Status")
'
'    Select Case strRet
'        Case "Print &Balance Report"
'            Call PrintBalRpt
'        Case "Print &Issue Receive Status"
'            Call PrintIssRcvStatus
'
'    End Select
    Call StartTrans(con)
    con.Execute "Delete From PrintVendBalLedger"
    For i = 1 To LVPrint.ListItems.Count
        With LVPrint.ListItems(i)
            con.Execute "Insert Into PrintVendBalLedger(VendID,ItemID,Balance) Values(" & .Tag & ",'" & .SubItems(1) & "'," & Val(.SubItems(3)) & ")"
        End With
    Next
    con.CommitTrans
    
    
    Dim rpt As craxddrt.Report
    Dim frm As New frmPrevRpt
    Set rpt = rptApp.OpenReport(RptPath & "\rptVendBalLedger.rpt")
    frm.ShowReport "{Venders.VendID}<>129 AND Left({Venders.VendID1},3)<>'HEA' AND Left({Venders.VendID1},3)<>'FOR' AND Left({Venders.VendID1},3)<>'FUN' ", rpt
    
End Sub

Private Sub PrintRpts()

    On Error GoTo err
    
    Screen.MousePointer = vbHourglass
    
    Call StartTrans(con)
    'Insert Into PrintIssRcvRptsTemp.
    con.Execute "Delete From PrintIssRcvRptsTemp"
    con.Execute "Delete From PrintIssRcvRpts"
    
    con.Execute "Insert Into PrintIssRcvRptsTemp Select RecieptID From VVendRcvdDetail Where VendID=" & cmbEmp.ID & " AND Month(DT)=" & Month(DTMonth) & " AND Year(DT)=" & year(DTMonth)
    
    con.Execute "Insert Into PrintIssRcvRpts Select CompItemDesc,DT,RecieptID,ItemCode,VenderName" & _
                ",VendID1,ProcessID,Process,RcvdQty,Wastage,LostQty,UserID,AuthUserID,AuthDT,0 As IssEntry " & _
                "From VVendRcvItemsrpt Where VendID=" & cmbEmp.ID & " AND Month(DT)=" & Month(DTMonth) & " AND Year(DT)=" & year(DTMonth)
                
    con.CommitTrans
    
    Dim rpt As craxddrt.Report  'New rptVendIssDetail
    Set rpt = rptApp.OpenReport(RptPath & "\rptVendIssDetail.rpt") '
    Dim ForTxt As String
'    If TableNAme = "VVendIssItemsrpt" Then
'        ForTxt = "GATEPASS ISSUED"
'    Else
'        ForTxt = "GOODS RECEIVED"
'    End If
    
    Dim FFDefs As FormulaFieldDefinitions
    Dim FFDef As craxddrt.FormulaFieldDefinition
    Set FFDefs = rpt.FormulaFields
    
    For Each FFDef In FFDefs
        'If FFDef.Name = "{@Heading}" Then
        '    FFDef.Text = "'" & ForTxt & "'"
        If FFDef.Name = "{@ForOriginal}" Then
            FFDef.Text = "'Duplicate'"
        End If
    Next
        
    Load frmPrevRpt
    frmPrevRpt.ShowReport " ", rpt
    
    
    Screen.MousePointer = vbDefault

    
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
End Sub

Private Sub cmdPrintSlip_Click()
    Call PrintRpts
End Sub

Public Sub cmdRefresh_Click()
    Me.MousePointer = vbHourglass
    Call RefreshLedger
    Me.MousePointer = vbNormal
    'Call Form_Resize
End Sub

Private Sub CommandButton1_Click()
    frmVendPymt.Show 1
End Sub

Private Sub Form_Load()
    
    On Error GoTo err
    
    DTMonth = ServerDate
    
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1600
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Processes>", "0"
    cmbDept.AddVals con, "Description", "Processes", "ProcessID"
    cmbDept.ID = "0"
    'cmbEmp.Text = "<All Contractors>"
    'Set LV.SmallIcons = ImageList1

    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub cmbdept_matched()
    
    On Error GoTo err
    cmbEmp.ClearVals
    'cmbEmp.AddItem "<All Contractors>", "0"
    'cmbEmp.AddItem "<All Contractors>", "0"
    If cmbDept.ID = "0" Then
        'cmbEmp.AddVals con, "'{' +  EmpID + '}  ' + Name", "Contractors", "EmpID"
        Call cmbEmp.AddVals(con, "VenderName", "VVenders", "VendID")
    Else
        Call cmbEmp.AddVals(con, "VenderName", "VVenders", "VendID", " Where PhaseID=" & cmbDept.ID)
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub RefreshLedger()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim lColor As Long
    Dim ITM As ListItem
    Dim lTotal As Long
    Dim lSubTotal As Long
    Dim strPrevItem As String
    
    If cmbEmp.ID = "0" Then
        'If cmbDept.ID = "0" Then
            Cond = " Where "
        'Else
        '    Cond = " Where DeptID='" & cmbDept.ID & "' and"
        'End If
    Else
        Cond = " Where VendID=" & cmbEmp.ID & " and"
    End If
    
    LV.Visible = False
    LVIssd.Visible = False
    LVRcvd.Visible = False
    
    
    LV.ListItems.Clear
    LVIssd.ListItems.Clear
    LVRcvd.ListItems.Clear
    lTotal = 0
    lSubTotal = 0
    
    
    Cond = Cond & "  Month(DT)=" & Month(DTMonth) & " AND Year(DT)=" & year(DTMonth)
    
    Sql = "select * From VVendIssItems " & Cond & " Order By ItemCode,DT Asc"
        
    Dim FirstDT As Date, ItemCodes As String
    ItemCodes = ""
    FirstDT = DateSerial(DTMonth.year, DTMonth.Month, 1)
        
    
        rs1.Open "Select IsNull(VendBalIssd.ItemCode,VendBalRcvd.ItemCode) As ItemCode,VendBalIssd.ItemDescription,VendBalIssd.IssdQty,VendBalRcvd.RcvdQty From " & _
         " (Select VendID,ItemCode,Itemdescription,Sum(IssQty) As IssdQty From VVendIssItems1 Where DT<'" & FirstDT & "' and vendid=" & cmbEmp.ID & _
         " Group By VendID,ItemCode,Itemdescription) VendBalIssd FULL OUTER JOIN (Select VendID,ItemCode,Sum(RcvdQty+Wastage+LostQty) As RcvdQty " & _
         " From VVendRcvdDetail Where DT<'" & FirstDT & "' " & _
         " And vendid=" & cmbEmp.ID & " AND ReqAuth=0 Group By VendID,ItemCode) VendBalRcvd ON VendBalIssd.VendID=VendBalRcvd.VendID " & _
         " AND VendBalIssd.ItemCode=VendBalRcvd.ItemCode Where IsNull(VendBalIssd.IssdQty,0)-IsNull(VendBalRcvd.RcvdQty,0)<>0", con, adOpenForwardOnly, adLockReadOnly
    
    With rs
    
        .Open Sql, con, adOpenForwardOnly, adLockReadOnly
        
        strPrevItem = ""
        
            
        For i = 0 To .RecordCount - 1
            If strPrevItem = "" Then
                rs1.Find "ItemCode='" & ![CompItemCode] & "'", 0, adSearchForward, 1
                Set ITM = LVIssd.ListItems.Add()
                ITM.ListSubItems.Add
                ITM.ListSubItems.Add , , !ItemCode & ""
                ITM.ListSubItems.Add(, , "********************B R O U G H T    F O R W A R D*******************").ForeColor = RGB(0, 150, 0)
                ITM.ListSubItems.Add(, , 0).Bold = True
                ITM.ListSubItems(4).ForeColor = RGB(0, 150, 0)
                If rs1.EOF = False Then
                    ITM.ListSubItems(4) = Val(rs1!IssdQty & "") - Val(rs1!RcvdQty & "")
                End If
                lSubTotal = lSubTotal + Val(ITM.SubItems(4))
                lTotal = lTotal + Val(ITM.SubItems(4))
            End If
            
            If ![CompItemCode] & "" <> strPrevItem And strPrevItem <> "" Then
                ItemCodes = ItemCodes & "'" & strPrevItem & "',"
                Set ITM = LVIssd.ListItems.Add()
                ITM.ListSubItems.Add
                ITM.ListSubItems.Add , , strPrevItem
                ITM.ListSubItems.Add(, , "********************T O T A L*******************").ForeColor = RGB(150, 0, 150)
                ITM.ListSubItems.Add(, , lSubTotal).Bold = True
                ITM.ListSubItems(4).ForeColor = RGB(150, 0, 150)
                lSubTotal = 0
                
                Set ITM = LVIssd.ListItems.Add()
                ITM.ListSubItems.Add
                ITM.ListSubItems.Add , , ""
                ITM.ListSubItems.Add , , ""
                ITM.ListSubItems.Add , , ""
                
                
                
                rs1.Find "ItemCode='" & ![CompItemCode] & "'", 0, adSearchForward, 1
                Set ITM = LVIssd.ListItems.Add()
                ITM.ListSubItems.Add
                ITM.ListSubItems.Add , , !ItemCode & ""
                ITM.ListSubItems.Add(, , "********************B R O U G H T    F O R W A R D*******************").ForeColor = RGB(0, 150, 0)
                ITM.ListSubItems.Add(, , 0).Bold = True
                ITM.ListSubItems(4).ForeColor = RGB(0, 150, 0)
                If rs1.EOF = False Then
                    ITM.ListSubItems(4) = Val(rs1!IssdQty & "") - Val(rs1!RcvdQty & "")
                End If
                lSubTotal = lSubTotal + Val(ITM.SubItems(4))
                lTotal = lTotal + Val(ITM.SubItems(4))
            End If
            
            strPrevItem = ![CompItemCode] & ""
                
            If !ReqAuth Then lColor = vbBlue Else lColor = LVIssd.ForeColor
            'Set itm = LV.ListItems.Add(, "'" & ![ItemID] & "", Format(![DT], "ddd-MMM-yyyy"))
            Set ITM = LVIssd.ListItems.Add(, , ![RecieptID] & "")
            'itm.SmallIcon = "Issue"
            ITM.Tag = ![VendID] & Format(![DT], "dd-MMM-yyyy") & ![CompItemCode] & ""
            ITM.ForeColor = lColor
            ITM.ListSubItems.Add(, , Format(![DT], "dd-MMM-yyyy")).ForeColor = lColor
            ITM.ListSubItems.Add(, , ![CompItemCode] & "").ForeColor = lColor
            ITM.ListSubItems.Add(, , ![ItemDescription] & "").ForeColor = lColor
            ITM.ListSubItems.Add(, , ![IssQty] & "").ForeColor = lColor
            'ITM.ListSubItems.Add(, , ![RcvdQty] & "").ForeColor = lColor
            lSubTotal = lSubTotal + Val(ITM.SubItems(4))
            lTotal = lTotal + Val(ITM.SubItems(4))
            .MoveNext
            
        Next
        If .RecordCount > 0 Then
            ItemCodes = ItemCodes & "'" & strPrevItem & "',"
            Set ITM = LVIssd.ListItems.Add()
            ITM.ListSubItems.Add
            ITM.ListSubItems.Add , , strPrevItem
            ITM.ListSubItems.Add(, , "********************T O T A L*******************").ForeColor = RGB(150, 0, 150)
            ITM.ListSubItems.Add(, , lSubTotal).Bold = True
            ITM.ListSubItems(4).ForeColor = RGB(150, 0, 150)
            
            Set ITM = LVIssd.ListItems.Add()
            ITM.ListSubItems.Add
            ITM.ListSubItems.Add , , ""
            ITM.ListSubItems.Add , , ""
            ITM.ListSubItems.Add , , ""
                
        End If
        .Close
        
        If ItemCodes <> "" Then
            rs1.Close
            ItemCodes = Left(ItemCodes, Len(ItemCodes) - 1)
            rs1.Open "Select IsNull(VendBalIssd.ItemCode,VendBalRcvd.ItemCode) As ItemCode,VendBalIssd.Itemdescription,VendBalIssd.IssdQty,VendBalRcvd.RcvdQty From " & _
             " (Select VendID,ItemCode,Itemdescription,Sum(IssQty) As IssdQty From VVendIssItems1 Where DT<'" & FirstDT & "' and vendid=" & cmbEmp.ID & _
             " And ItemCode Not In(" & ItemCodes & ") Group By VendID,ItemCode,Itemdescription) VendBalIssd FULL OUTER JOIN (Select VendID,ItemCode,Sum(RcvdQty+Wastage+LostQty) As RcvdQty " & _
             " From VVendRcvdDetail Where DT<'" & FirstDT & "' and ItemCode Not In(" & ItemCodes & ")" & _
             " and vendid=" & cmbEmp.ID & " AND ReqAuth=0 Group By VendID,ItemCode) VendBalRcvd ON VendBalIssd.VendID=VendBalRcvd.VendID " & _
             " AND VendBalIssd.ItemCode=VendBalRcvd.ItemCode", con, adOpenForwardOnly, adLockReadOnly
        Else
            
        End If
        For i = 0 To rs1.RecordCount - 1
            Set ITM = LVIssd.ListItems.Add()
            ITM.ListSubItems.Add
            ITM.ListSubItems.Add , , rs1!ItemCode & ""
            ITM.ListSubItems.Add(, , "********************B R O U G H T    F O R W A R D*******************").ForeColor = RGB(0, 150, 0)
            ITM.ListSubItems.Add(, , 0).Bold = True
            ITM.ListSubItems(4).ForeColor = RGB(0, 150, 0)
            If rs1.EOF = False Then
                ITM.ListSubItems(4) = Val(rs1!IssdQty & "") - Val(rs1!RcvdQty & "")
            End If
            lTotal = lTotal + Val(ITM.SubItems(4))
            rs1.MoveNext
        Next
        rs1.Close
        
        txtIssTotal = lTotal
        
        Dim TotalRate As Double
        
        '.Open "Select * From VVendRcvdItems  " & Cond & " and TransEntry Is Null and RcvdQty>0 Order By DT", con, adOpenForwardOnly, adLockReadOnly
        '.Open "Select * From VVenderLedger " & Cond & " Order By DT ", con, adOpenForwardOnly, adLockReadOnly
        .Open "Select * From VVendRcvdItems " & Cond & " Order By ItemCode,DT ", con, adOpenForwardOnly, adLockReadOnly
        lTotal = 0
        lSubTotal = 0
        strPrevItem = ""
        For i = 0 To .RecordCount - 1
        
             If ![ItemCode] & "" <> strPrevItem And strPrevItem <> "" Then
                Set ITM = LVRcvd.ListItems.Add()
                ITM.ListSubItems.Add
                ITM.ListSubItems.Add
                ITM.ListSubItems.Add(, , "********************T O T A L*******************").ForeColor = RGB(150, 0, 150)
                ITM.ListSubItems.Add(, , lSubTotal).Bold = True
                ITM.ListSubItems(4).ForeColor = RGB(150, 0, 150)
                lSubTotal = 0
                
                Set ITM = LVRcvd.ListItems.Add()
                ITM.ListSubItems.Add
                ITM.ListSubItems.Add , , ""
                ITM.ListSubItems.Add , , ""
                ITM.ListSubItems.Add , , ""
                
            End If
            
            strPrevItem = ![ItemCode] & ""
            
            
            If !ReqAuth Then
                lColor = vbBlue
            ElseIf Left(![RecieptID] & "", 3) = "WST" Then
                lColor = vbRed
            Else
                lColor = LV.ForeColor
            End If
            
            Set ITM = LVRcvd.ListItems.Add(, , ![RecieptID] & "")
            ITM.Tag = ![EntryID] & "'"
            ITM.ForeColor = lColor
            
            ITM.ListSubItems.Add(, , Format(![DT], "dd-MMM-yyyy")).ForeColor = lColor
            
            ITM.ListSubItems.Add(, , ![ItemCode] & "").ForeColor = lColor
            ITM.ListSubItems.Add(, , ![ItemDescription] & "").ForeColor = lColor
            
            If Left(ITM.Text, 3) = "WST" Then
                ITM.ListSubItems.Add(, , Val(![Wastage] & "")).ForeColor = lColor
                'ITM.ListSubItems.Add(, , 0).ForeColor = lColor
                TotalRate = 0
                lSubTotal = lSubTotal + Val(![Wastage] & "")
            ElseIf Left(ITM.Text, 3) = "LST" Then
                ITM.ListSubItems.Add(, , Val(![LostQty] & "")).ForeColor = lColor
                'ITM.ListSubItems.Add(, , 0).ForeColor = lColor
                TotalRate = 0
                lSubTotal = lSubTotal + Val(![LostQty] & "")
            Else
                ITM.ListSubItems.Add(, , Val(![RcvdQty] & "")).ForeColor = lColor
                'ITM.ListSubItems.Add(, , Round(Val(![Rate] & ""), 3)).ForeColor = lColor
                TotalRate = Round(Val(![Rate] & ""), 3)
                lSubTotal = lSubTotal + Val(![RcvdQty] & "")
            End If
            
            'ITM.SubItems(6) = Val(![PlantRate] & "")
            'ITM.SubItems(7) = Val(![StampRate] & "")
            'ITM.SubItems(8) = Val(![SnaffingRate] & "")
            'TotalRate = Round(Val(![Rate] & ""), 3) '+ Val(![PlantRate] & "") + Val(![StampRate] & "") + Val(![SnaffingRate] & "")
            
            'ITM.ListSubItems.Add(, , TotalRate * Val(ITM.SubItems(4))).ForeColor = lColor
            'dWork = dWork + Val(ITM.SubItems(6))
            lTotal = lTotal + Val(ITM.SubItems(4))
            
            .MoveNext
        Next
        .Close
        
        
        Set ITM = LVRcvd.ListItems.Add()
        ITM.ListSubItems.Add
        ITM.ListSubItems.Add
        ITM.ListSubItems.Add(, , "********************T O T A L*******************").ForeColor = RGB(150, 0, 150)
        ITM.ListSubItems.Add(, , lSubTotal).Bold = True
        ITM.ListSubItems(4).ForeColor = RGB(150, 0, 150)
        lSubTotal = 0
            
        txtRcvdTotal = lTotal
        
     
        .Open "Select ItemCode,ItemDescription,Sum(IssQty-RcvdQty)-dbo.F_AuthQtyForIss(ItemCode,null,null,null," & cmbEmp.ID & ") as Quantity From VVendIssItems1 Where IssQty-RcvdQty>0 and VendID=" & cmbEmp.ID & " And IsNull(ReqAuth,0)=0 Group By ItemCode,ItemDescription ", con, adOpenForwardOnly, adLockReadOnly
        .Filter = "Quantity<>0"
        lTotal = 0
        Do Until .EOF
            Set ITM = LV.ListItems.Add(, , LV.ListItems.Count + 1)
            ITM.SubItems(1) = ![ItemCode] & ""
            ITM.SubItems(2) = ![ItemDescription] & ""
            ITM.SubItems(3) = Val(![Quantity] & "")
            'ITM.SubItems(3) = Val(![Issdqty] & "") - Val(!rcvdqty & "")
            lTotal = lTotal + Val(ITM.SubItems(3))
            .MoveNext
        Loop
        .Close
        txtBalTotal = lTotal
        
        .Open "Select VendID,ItemCode,ItemDescription,Sum(IssQty-RcvdQty)-dbo.F_AuthQtyForIss(ItemCode,null,null,null,VendID) as Quantity From VVendIssItems1 Where IssQty-RcvdQty>0 And IsNull(ReqAuth,0)=0 Group By VendID,ItemCode,ItemDescription ", con, adOpenForwardOnly, adLockReadOnly
        .Filter = "Quantity<>0"
        lTotal = 0
        LVPrint.ListItems.Clear
        Do Until .EOF
            Set ITM = LVPrint.ListItems.Add(, , LVPrint.ListItems.Count + 1)
            ITM.Tag = ![VendID] & ""
            ITM.SubItems(1) = ![ItemCode] & ""
            ITM.SubItems(2) = ![ItemDescription] & ""
            'ITM.SubItems(3) = Val(![RemQty] & "")
            ITM.SubItems(3) = Val(![Quantity] & "")
            lTotal = lTotal + Val(ITM.SubItems(3))
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    LV.Visible = True
    LVIssd.Visible = True
    LVRcvd.Visible = True
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LV_DblClick()
    Exit Sub
    'If LV.ListItems.Count = 0 Or UserHasAccess("EditOverTime") = False Then Exit Sub
    
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(5).Left + 35, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(5).Width
        .Text = LV.SelectedItem.SubItems(4)
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With

End Sub



Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    
    On Error GoTo err
    If KeyAscii = 13 Then
        con.Execute "Update ContLedger Set Rate=" & Val(txtEdit) & " Where Empid+format(DT,'dd-MMM-yyyy')+ItemId='" & LV.SelectedItem.Tag & "'"
        LV.SelectedItem.SubItems(4) = Val(txtEdit)
        LV.SetFocus
    ElseIf KeyAscii = 27 Then
        LV.SetFocus
    Else
        KeyAscii = OnlyNumber(KeyAscii, True)
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub

Private Sub LVIssd_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    LVIssd.SortKey = "2"
    LVIssd.Sorted = True
    
End Sub
