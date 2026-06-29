VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmTrialBalance 
   ClientHeight    =   8715
   ClientLeft      =   165
   ClientTop       =   60
   ClientWidth     =   12870
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8715
   ScaleWidth      =   12870
   WindowState     =   2  'Maximized
   Begin VB.Frame fTotals 
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   1590
      TabIndex        =   14
      Top             =   7635
      Width           =   11220
      Begin VB.Label lblTotalOpenBalDebit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   30
         TabIndex        =   23
         Top             =   285
         Width           =   1875
      End
      Begin VB.Label lblTotalBalanceCredit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   9315
         TabIndex        =   22
         Top             =   285
         Width           =   1875
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmTrialBalance.frx":0000
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   7
         Left            =   30
         TabIndex        =   19
         Top             =   45
         Width           =   11160
      End
      Begin VB.Label lblTotalBalanceDebit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   7455
         TabIndex        =   18
         Top             =   285
         Width           =   1875
      End
      Begin VB.Label lblTotalCredit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   5595
         TabIndex        =   17
         Top             =   285
         Width           =   1875
      End
      Begin VB.Label lblTotalDebit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3735
         TabIndex        =   16
         Top             =   285
         Width           =   1875
      End
      Begin VB.Label lblTotalOpenBalCredit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1890
         TabIndex        =   15
         Top             =   285
         Width           =   1860
      End
   End
   Begin VB.PictureBox PicStatus 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   45
      ScaleHeight     =   420
      ScaleWidth      =   2625
      TabIndex        =   10
      Top             =   8220
      Visible         =   0   'False
      Width           =   2625
      Begin MSComctlLib.ProgressBar PBar 
         Height          =   210
         Left            =   30
         TabIndex        =   11
         Top             =   210
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   370
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Calculating Trial Balance..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   12
         Top             =   0
         Width           =   1905
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2760
      Top             =   3735
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrialBalance.frx":00A0
            Key             =   "ACC"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrialBalance.frx":04F4
            Key             =   "PARIN"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrialBalance.frx":0948
            Key             =   "ACCIN"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrialBalance.frx":0D9C
            Key             =   "PAR"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrialBalance.frx":11F0
            Key             =   "TYPE"
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   915
      Top             =   1095
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Trial Balance"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSComctlLib.ListView LV 
      Height          =   6630
      Left            =   15
      TabIndex        =   5
      Top             =   990
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   11695
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   15199215
      BorderStyle     =   1
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Account #"
         Object.Width           =   1799
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Account Title"
         Object.Width           =   4180
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Opening Bal(Dr)"
         Object.Width           =   3228
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Opening Bal.(Cr)"
         Object.Width           =   3228
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Debit"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Credit"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Last Balance(Dr)"
         Object.Width           =   3228
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Last Balance(Cr)"
         Object.Width           =   3228
      EndProperty
   End
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   960
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11760
      Begin VB.CheckBox chkNonZero 
         Caption         =   "Show Accounts With Zero Transaction"
         Height          =   210
         Left            =   6930
         TabIndex        =   3
         Top             =   285
         Value           =   1  'Checked
         Width           =   3540
      End
      Begin VB.CheckBox chkDetail 
         Caption         =   "Detailed Report"
         Height          =   210
         Left            =   6930
         TabIndex        =   2
         Top             =   660
         Width           =   1635
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   4110
         TabIndex        =   1
         Top             =   555
         Width           =   2745
         _ExtentX        =   4842
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
         CustomFormat    =   "dddd, MMM dd, yyyy"
         Format          =   120193027
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   765
         TabIndex        =   0
         Top             =   555
         Width           =   2715
         _ExtentX        =   4789
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
         CustomFormat    =   "dddd, MMM dd, yyyy"
         Format          =   120127491
         CurrentDate     =   37055
      End
      Begin MSForms.ComboBox cmbAccHead 
         Height          =   360
         Left            =   1485
         TabIndex        =   24
         Top             =   150
         Width           =   5370
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "9472;635"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Account Head : "
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   21
         Top             =   240
         Width           =   1305
      End
      Begin MSForms.CommandButton cmdRefresh 
         Default         =   -1  'True
         Height          =   360
         Left            =   10455
         TabIndex        =   4
         Top             =   510
         Width           =   1230
         Caption         =   "Refresh "
         PicturePosition =   327683
         Size            =   "2170;635"
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
         AutoSize        =   -1  'True
         Caption         =   "To :"
         Height          =   195
         Index           =   1
         Left            =   3750
         TabIndex        =   9
         Top             =   615
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From :"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   615
         Width           =   525
      End
   End
   Begin MSForms.CommandButton cmdClosingBalances 
      Height          =   375
      Left            =   5880
      TabIndex        =   25
      Top             =   8250
      Width           =   2130
      Caption         =   "Print With Closing Bal."
      PicturePosition =   327683
      Size            =   "3757;661"
      Accelerator     =   79
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdPrintOP 
      Height          =   375
      Left            =   8055
      TabIndex        =   20
      Top             =   8250
      Width           =   1980
      Caption         =   "Print With Op.Bal"
      PicturePosition =   327683
      Size            =   "3492;661"
      Accelerator     =   79
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   375
      Left            =   11475
      TabIndex        =   13
      Top             =   8250
      Width           =   1335
      ForeColor       =   64
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
   Begin MSForms.CommandButton cmdPrint 
      Height          =   375
      Left            =   10080
      TabIndex        =   6
      Top             =   8250
      Width           =   1335
      Caption         =   "Print        "
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
End
Attribute VB_Name = "frmTrialBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PermissionDenied As Boolean

Dim Rec As New ADODB.Recordset
Dim rsAccounts As New ADODB.Recordset
Dim i As Integer
Dim Bal As Double
Dim AccList As String
Dim SubAccList As String
Dim TotalDebit As Double, TotalCredit As Double

Dim ShowAll As Boolean
Dim dTotalOpeningDebit As Double, dTotalOpeningCredit As Double

Private Sub RefreshList()

    Screen.MousePointer = vbHourglass
    On Error GoTo err
    LV.Sorted = False

    Dim PrevIndex As Double

    If LV.ListItems.count > 0 Then PrevIndex = LV.SelectedItem.Index

    'LV.Visible = False
    LV.ListItems.Clear
    Dim i As Integer
    Dim Balance As Double, dLastBalDr As Double, dLastBalCr As Double

    Screen.MousePointer = vbHourglass

    With rsAccounts
        Dim OBal As Double, dOBalDr As Double, dOBalCr As Double
        .MoveFirst
   
        pbar.Min = 0
        pbar.Max = .RecordCount + 20
        PicStatus.Visible = .RecordCount > 0
        PicStatus.Refresh
        Dim dParentOpeningBalance As Double
        For i = 0 To .RecordCount - 1
            TotalDebit = 0
            TotalCredit = 0
            pbar.Value = i + 1
            AccList = ""
            If ![Parent] Then
                OBal = CalculateBalance(.Fields("AccNo"), True, dOBalDr, dOBalCr, dLastBalDr, dLastBalCr)
                'dLastBalDr = dOBalDr + TotalDebit
                'dLastBalCr = dOBalCr + TotalCredit
            Else
                Call CalculateBalance(.Fields("AccNo"))
                OBal = Val((![OpeningBalance] & ""))
                If OBal >= 0 Then
                    dOBalDr = OBal
                    dOBalCr = 0
                Else
                    dOBalDr = 0
                    dOBalCr = Abs(OBal)
                End If
                Balance = OBal + TotalDebit - TotalCredit
                If Balance >= 0 Then
                    dLastBalDr = Balance
                    dLastBalCr = 0
                Else
                    dLastBalCr = Abs(Balance)
                    dLastBalDr = 0
                End If
                If OBal < 0 Then
                    dTotalOpeningCredit = dTotalOpeningCredit + Abs(OBal)
                Else
                    dTotalOpeningDebit = dTotalOpeningDebit + OBal
                End If
            End If
            
            'Balance = (![OpeningBalance] & "") + TotalDebit - TotalCredit
            Balance = OBal + TotalDebit - TotalCredit
            
            Call AddToList(i, .Fields("SubaccOf") & "", .Fields("AccNo"), .Fields("AccTitle"), dOBalCr, dOBalDr, TotalDebit, TotalCredit, dLastBalCr, dLastBalDr, .Fields("AccNo"), ![Parent], True)
            .MoveNEXT
        Next i
    End With
    Call AddHeads
    PicStatus.Visible = False
    ''''Very Important To Sort
    LV.Sorted = True

    'LV.FindItem

    LV.Visible = True

    If PrevIndex > 0 And PrevIndex <= LV.ListItems.count Then
        LV.ListItems(PrevIndex).Selected = True
        LV.ListItems(PrevIndex).EnsureVisible
    End If


    Screen.MousePointer = vbDefault

    Exit Sub
err:
    If err.Number = 380 Then Resume Next Else MsgBox "ERROR :" & err.Number & vbNewLine & err.Description
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub chkDetail_Click()
   'Call cmdRefresh_Click
End Sub

Private Sub cmdOpt_Click()

End Sub

Private Sub chkNonZero_Click()
    'Call cmdRefresh_Click
End Sub

Private Sub cmdClose_Click()
    Set rsAccounts = Nothing
    Unload Me
End Sub

Private Sub cmdClosingBalances_Click()
    Call PrintTrialBalanceOpenging(True)
End Sub

Private Sub PrintTrialBalanceOpenging(Optional bClosing As Boolean = False)

    Screen.MousePointer = vbHourglass
    con.BeginTrans
    con.Execute "DELETE FROM TempTrialBal"
    Dim Parent As Boolean
    For i = 1 To LV.ListItems.count - 1
        With LV.ListItems(i)
            con.Execute "Insert into TempTrialBal(AccNo,AccTitle,OpenBal,Debit,Credit,LastBal,Bold) " & _
             " VALUES ('" & .Text & "','" & .SubItems(1) & "'," & Val(.ListSubItems(2).Tag) & "," & Val(.ListSubItems(3).Tag) & "," & Val(.ListSubItems(4).Tag) & "," & Val(.ListSubItems(5).Tag) & "," & Abs(.Bold) & ")"
        End With
    Next

    con.CommitTrans

    'Sleep (500)
    Dim rpt As CRAXDDRT.Report
    If bClosing Then
        Set rpt = rptApp.OpenReport(RptPath & "\TrialBalWCl.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\TrialBalWOp.rpt")
    End If
    
    Dim frm As New frmPrevRpt
    
    rpt.FormulaFields.GetItemByName("Company").Text = "'" & strCompany & "'"
    rpt.FormulaFields.GetItemByName("Caption").Text = "'" & IIf(chkDetail = vbChecked, "Trial Balance (Detail)", "Trial Balance (Summary)") & "'"
    rpt.FormulaFields.GetItemByName("FromTo").Text = "'From " & Format(DtFrom, "dd/MM/yyyy") & " To " & Format(DtTo, "dd/MM/yyyy") & "'"
    rpt.FormulaFields.GetItemByName("TotalDebit").Text = Val(Format(LV.ListItems(LV.ListItems.count).SubItems(3), "0.00"))
    rpt.FormulaFields.GetItemByName("TotalCredit").Text = Val(Format(LV.ListItems(LV.ListItems.count).SubItems(4), "0.00"))
    rpt.FormulaFields.GetItemByName("Address").Text = "'" & COMPANYADDRESS & "'"
    
    frm.ShowReport "", rpt
    
    Screen.MousePointer = vbDefault
    
End Sub
Private Sub cmdPrint_Click()
   Call PrintTrialBal
End Sub

Private Sub cmdPrintOP_Click()
    Call PrintTrialBalanceOpenging(False)
End Sub

Private Sub cmdRefresh_Click()

    Set rsAccounts = Nothing
    'Call Form_Activate
    dTotalOpeningDebit = 0
    dTotalOpeningCredit = 0
    Call RefreshTrial
    
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub DtFrom_Change()
    'Call setDateRange(DtFrom)
    DtTo.MinDate = DtFrom
End Sub


Private Sub DtTo_Change()
    'Call setDateRange(DtTo)
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

Private Sub RefreshTrial()

    On Error GoTo err
    If rsAccounts.State <> 0 Then Exit Sub
   
    Screen.MousePointer = vbHourglass

    Me.KeyPreview = True
    Dim Condition As String
    'If ShowAll Then
    '   Condition = ""
    'Else
    '   Condition = "Where Active "
    'End If
    'Condition = ""
    rsAccounts.CursorType = adOpenForwardOnly
    rsAccounts.LockType = adLockReadOnly
    
    Set LV.SmallIcons = ImageList1
   
    Dim cmd As New ADODB.Command
    With cmd
        .CommandTimeout = 120
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "TBBase"
        .CommandType = adCmdStoredProc
        .Parameters("@FromDate").Value = DtFrom
        If cmbAccHead.ListIndex > 0 Then
            .Parameters("@Code").Value = cmbAccHead.List(cmbAccHead.ListIndex, 1) 'CStr(cmbAccHead.ItemData(cmbAccHead.ListIndex))
        End If
        .Parameters("@Summary").Value = IIf(chkDetail.Value = vbChecked, 0, 1)
        Set rsAccounts = cmd.Execute
    End With
       
    Set cmd = Nothing
        
    con.Execute "Delete From TempTBBase"

    Dim TheBalance As Double, TheOpeningBalance As Double
    Dim dOpBalDr As Double, dOpBalCr As Double
    Dim dLastBalDr As Double, dLastBalCr As Double
    Dim TotalDebit As Double
    Dim TotalCredit As Double
      
    With rsAccounts
        For i = 0 To .RecordCount - 1
            TheBalance = Val(![Balance] & "")    'Added....
            If TheBalance >= 0 Then
                dLastBalDr = TheBalance
            Else
                dLastBalCr = Abs(TheBalance)
            End If
            TheOpeningBalance = Val(!OpeningBalance & "")
            If TheOpeningBalance >= 0 Then
                dOpBalDr = TheOpeningBalance
            Else
                dOpBalCr = Abs(TheOpeningBalance)
            End If
            
            con.Execute "Insert Into TempTBBase(AccNo,TheBalanceCr,TheBalanceDr,OpeningBalanceCr,OpeningBalanceDr) Values('" & ![AccNo] & _
            "'," & dLastBalCr & "," & dLastBalDr & "," & dOpBalCr & "," & dOpBalDr & ")"
            
            .MoveNEXT
        Next
         
        '.Close
        
'   If chkDetail = vbChecked Then
'    .Open "Select * From VTBDetail", con, adOpenForwardOnly, adLockReadOnly
'   Else
'    .Open "Select * From VTBSummary", con, adOpenForwardOnly, adLockReadOnly
'   End If
   
   End With
    DtFrom.Tag = DtFrom
    DtTo.Tag = DtTo
   
    Call RefreshList
'    Dim CrdTotal As Double, DbtTotal As Double
'
'    With con.Execute("Select Sum(debit) as TotalDebit,Sum(Credit) as totalCredit From vouchers where (VDate Between '" & DTFrom & "' AND '" & DTTo & "')")
'        CrdTotal = Val(![TotalCredit] & "")
'        DbtTotal = Val(![TotalDebit] & "")
'        .Close
'    End With
'
'    With LV.ListItems.add(, , "Totals :")
'        .ListSubItems.add(, , "                           Totals : ").Bold = True
'        .ListSubItems.add
'        .ListSubItems.add(, , Format(DbtTotal, "#,##0.00")).Bold = True
'        .ListSubItems.add(, , Format(CrdTotal, "#,##0.00")).Bold = True
'    End With
'    If strCompany = "Banzai" Then
'        Call SetMasterHeadColors
'    End If
    Screen.MousePointer = vbDefault
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub SetMasterHeadColors()
    Dim lForeColor As Long, j As Integer
    lForeColor = RGB(255, 50, 50)
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Bold Then
                .ForeColor = lForeColor
                For j = 1 To .ListSubItems.count
                    .ListSubItems(j).ForeColor = lForeColor
                Next
            End If
        End With
    Next
End Sub
Private Sub AddToList(ByVal No As Integer, Parent As String, code As String, AccTitle As String, OpenBalCr As Double, OpenBalDr As Double, Debit As Double, Credit As Double, LastBalCr As Double, LastBalDr As Double, key As String, ParentAccount As Boolean, Active As Boolean)

    Dim dOpBalDr As Double, dOpBalCr As Double
    Dim dLastBalDr As Double, dLastBalCr As Double

    If chkNonZero = vbUnchecked And (Debit + Credit) = 0 Then Exit Sub

    Dim LColor As Long
    Dim NumSpaces As Integer
    Dim TheTag As String

    Dim Openbalance As String
    dOpBalDr = OpenBalDr
    dOpBalCr = OpenBalCr
'    If OpenBal < 0 Then
'        Openbalance = Format(Abs(OpenBal), "#,##0.00") & " CR"
'    ElseIf OpenBal > 0 Then
'        Openbalance = Format(Abs(OpenBal), "#,##0.00") & " DR"
'    Else
'        Openbalance = Format(Abs(OpenBal), "#,##0.00") & "      "
'    End If
    
    Dim CurrBalance As String
    dLastBalDr = LastBalDr
    dLastBalCr = LastBalCr
'    If Bal < 0 Then
'        CurrBalance = Format(Abs(Bal), "#,##0.00") & " CR"
'    ElseIf Bal > 0 Then
'        CurrBalance = Format(Abs(Bal), "#,##0.00") & " DR"
'    Else
'        CurrBalance = Format(Abs(Bal), "#,##0.00") & "      "
'    End If

    If Not Active Then
        LColor = &H808080
        TheTag = "Inactive"
    Else
        LColor = &H800000
        TheTag = "Active"
    End If

    Dim img As String
    If ParentAccount Then
        If Active Then
            img = "PAR"
        Else
            img = "PARIN"
        End If
    Else
        If Active Then
            img = "ACC"
        Else
            img = "ACCIN"
        End If
    End If
   
    Dim i As Integer

    For i = 1 To LV.ListItems.count
        If Parent = Replace(LV.ListItems(i).key, "[", "") Then
            NumSpaces = Len(LV.ListItems(i).ListSubItems(1).Text) - Len(Trim(LV.ListItems(i).ListSubItems(1).Text))
            LV.ListItems.add i + 1, "[" & key, code, , img
            LV.ListItems(i + 1).Bold = ParentAccount
            LV.ListItems(i + 1).Tag = TheTag
            LV.ListItems(i + 1).ForeColor = LColor
            LV.ListItems(i + 1).ListSubItems.add(, , Space$(5 + NumSpaces) & "»  " & AccTitle).ToolTipText = AccTitle
            LV.ListItems(i + 1).ListSubItems(1).ForeColor = LColor
            LV.ListItems(i + 1).ListSubItems(1).Bold = ParentAccount
            'Start Opening Balance
            LV.ListItems(i + 1).ListSubItems.add , , Format(dOpBalDr, "#,##0.00")
            LV.ListItems(i + 1).ListSubItems(2).Tag = dOpBalDr
            LV.ListItems(i + 1).ListSubItems(2).ForeColor = LColor
            LV.ListItems(i + 1).ListSubItems(2).Bold = ParentAccount
            LV.ListItems(i + 1).ListSubItems.add , , Format(dOpBalCr, "#,##0.00")
            LV.ListItems(i + 1).ListSubItems(3).Tag = dOpBalCr
            LV.ListItems(i + 1).ListSubItems(3).ForeColor = LColor
            LV.ListItems(i + 1).ListSubItems(3).Bold = ParentAccount
            'Finish Opening Balance
            
            LV.ListItems(i + 1).ListSubItems.add , , Format(Debit, "#,##0.00")
            LV.ListItems(i + 1).ListSubItems(4).ForeColor = LColor
            LV.ListItems(i + 1).ListSubItems(4).Tag = Debit
            LV.ListItems(i + 1).ListSubItems.add , , Format(Credit, "#,##0.00")
            LV.ListItems(i + 1).ListSubItems(5).ForeColor = LColor
            LV.ListItems(i + 1).ListSubItems(5).Tag = Credit
            LV.ListItems(i + 1).ListSubItems(5).Bold = ParentAccount
            
            'Start Last Balance
            LV.ListItems(i + 1).ListSubItems.add , , Format(dLastBalDr, "#,##0.00")
            LV.ListItems(i + 1).ListSubItems(6).Tag = dLastBalDr
            LV.ListItems(i + 1).ListSubItems(6).Bold = ParentAccount
            LV.ListItems(i + 1).ListSubItems.add , , Format(dLastBalCr, "#,##0.00")
            LV.ListItems(i + 1).ListSubItems(7).Tag = dLastBalCr
            LV.ListItems(i + 1).ListSubItems(7).Bold = ParentAccount
            'Finish Last Balance
            Exit Sub
        End If
    Next i


    No = LV.ListItems.add(, "[" & key, code, , img).Index - 1
    LV.ListItems(No + 1).Bold = ParentAccount
    LV.ListItems(No + 1).Tag = TheTag
    
    LV.ListItems(No + 1).ForeColor = LColor
    LV.ListItems(No + 1).ListSubItems.add(, , "      »  " & AccTitle).ToolTipText = AccTitle
    LV.ListItems(No + 1).ListSubItems(1).ForeColor = LColor
    LV.ListItems(No + 1).ListSubItems(1).Bold = ParentAccount
    
    'Start Opening Balance
    LV.ListItems(No + 1).ListSubItems.add , , Format(dOpBalDr, "#,##0.00")
    LV.ListItems(No + 1).ListSubItems(2).Tag = dOpBalDr
    LV.ListItems(No + 1).ListSubItems(2).ForeColor = LColor
    LV.ListItems(No + 1).ListSubItems(2).Bold = ParentAccount
    LV.ListItems(No + 1).ListSubItems.add , , Format(dOpBalCr, "#,##0.00")
    LV.ListItems(No + 1).ListSubItems(3).Tag = dOpBalCr
    LV.ListItems(No + 1).ListSubItems(3).ForeColor = LColor
    LV.ListItems(No + 1).ListSubItems(3).Bold = ParentAccount
    'Finish Opening Balance
    
    LV.ListItems(No + 1).ListSubItems.add , , Format(Debit, "#,##0.00")
    LV.ListItems(No + 1).ListSubItems(4).Bold = ParentAccount
    LV.ListItems(No + 1).ListSubItems(4).ForeColor = LColor
    LV.ListItems(No + 1).ListSubItems(4).Tag = Debit
    LV.ListItems(No + 1).ListSubItems.add , , Format(Credit, "#,##0.00")
    LV.ListItems(No + 1).ListSubItems(5).Bold = ParentAccount
    LV.ListItems(No + 1).ListSubItems(5).ForeColor = LColor
    LV.ListItems(No + 1).ListSubItems(5).Tag = Credit
    
    'Start Last Balance
    LV.ListItems(No + 1).ListSubItems.add , , Format(dLastBalDr, "#,##0.00")
    LV.ListItems(No + 1).ListSubItems(6).Tag = dLastBalDr
    LV.ListItems(No + 1).ListSubItems(6).Bold = ParentAccount
    LV.ListItems(No + 1).ListSubItems.add , , Format(dLastBalCr, "#,##0.00")
    LV.ListItems(No + 1).ListSubItems(7).Tag = dLastBalCr
    LV.ListItems(No + 1).ListSubItems(7).Bold = ParentAccount
    'Finish Last Balance
    
End Sub

Private Sub CreateLists(PRec As ADODB.Recordset, AccNo As String)

'If Len(AccList) > 0 Then AccList = AccList & "'" & AccNo & "',"
AccList = AccList & "'" & AccNo & "',"
SubAccList = ""
Dim i As Integer
With PRec
    For i = 0 To .RecordCount - 1
        AccList = AccList & "'" & .Fields("AccNo") & "',"
        SubAccList = SubAccList & "'" & .Fields("AccNo") & "',"
        .MoveNEXT
    Next i
    
End With
If Right$(SubAccList, 1) = "," Then SubAccList = Left$(SubAccList, Len(SubAccList) - 1)

End Sub

Private Function CalculateBalance(AccNo As String, Optional Parent As Boolean, Optional ByRef dOpeningBalanceDr As Double, Optional ByRef dOpeningBalanceCr As Double, Optional ByRef dLastBalDr As Double, Optional ByRef dLastBalCr As Double) As Double

    Dim i As Integer
    
    With Rec
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        .Open "Select AccNo from Accounts Where SubAccOf='" & AccNo & "'", con, adOpenForwardOnly, adLockReadOnly
        Do While .RecordCount > 0
            Call CreateLists(Rec, AccNo)
            If .State = 1 Then .Close
            .CursorLocation = adUseClient
            .Open "Select AccNo from Accounts Where SubAccOf in (" & SubAccList & ")", con, adOpenForwardOnly, adLockReadOnly
        Loop
        If Len(AccList) > 0 Then
            AccList = Left$(AccList, Len(AccList) - 1)
        Else
            AccList = "'" & AccNo & "'"
        End If
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
         .Open "Select Sum(Debit),Sum(Credit) from Vouchers Where (AccNo in (" & AccList & ")) and (Credit is not null and Debit is not null)" & _
         " AND (VDate Between '" & DtFrom & "' AND '" & DtTo & "')", con, adOpenForwardOnly, adLockReadOnly
        
        TotalDebit = Val(.Fields(0) & "")
        TotalCredit = Val(.Fields(1) & "")
    
        .Close
        
        'If Parent Then CalculateBalance = Val(con.Execute("Select Sum(OpeningBalance) From VTBDetail Where Accno In (" & AccList & ")").Fields(0) & "")
        If Parent Then
            .Open "SELECT SUM(ISNULL(TBase.TheBalance, dbo.Accounts.openbal)) AS OpeningBalance," & _
            "SUM(CASE WHEN ISNULL(TBase.TheBalance, dbo.Accounts.openbal)>=0 THEN ISNULL(TBase.TheBalance, dbo.Accounts.openbal) ELSE 0 END) AS OpeningBalanceDr," & _
            "SUM(CASE WHEN ISNULL(TBase.TheBalance, dbo.Accounts.openbal)<0 THEN ABS(ISNULL(TBase.TheBalance, dbo.Accounts.openbal)) ELSE 0 END) AS OpeningBalanceCr," & _
            "SUM(CASE WHEN ISNULL(TBaseBal.TheBalance, dbo.Accounts.OpenBal)>=0 THEN ISNULL(TBaseBal.TheBalance, dbo.Accounts.OpenBal) ELSE 0 END) AS LastBalanceBalanceDr," & _
            "SUM(CASE WHEN ISNULL(TBaseBal.TheBalance, dbo.Accounts.OpenBal)<0 THEN ABS(ISNULL(TBaseBal.TheBalance, dbo.Accounts.OpenBal)) ELSE 0 END) AS LastBalanceBalanceCr " & _
            "FROM dbo.Accounts LEFT OUTER JOIN " & _
             "(SELECT Accno,balance AS TheBalance From dbo.Vouchers WHERE (SNo IN (SELECT MAX(Sno) FROM Vouchers " & _
             "WHERE VDATE<'" & DtFrom.Value & "' GROUP BY Accno))) TBase ON dbo.Accounts.AccNo = TBase.Accno LEFT OUTER JOIN " & _
             "(SELECT Accno,balance AS TheBalance From dbo.Vouchers WHERE (SNo IN (SELECT MAX(Sno) FROM Vouchers WHERE VDATE<='" & _
             DtTo.Value & "' GROUP BY Accno))) TBaseBal ON dbo.Accounts.AccNo = TBaseBal.Accno " & _
             "WHERE Accounts.AccNo In(" & AccList & ")", con, adOpenForwardOnly, adLockReadOnly
            
            
            dOpeningBalanceDr = Val(.Fields(1).Value & "")
            dOpeningBalanceCr = Val(.Fields(2).Value & "")
            dLastBalDr = Val(.Fields(3).Value & "")
            dLastBalCr = Val(.Fields(4).Value & "")
            CalculateBalance = Val(.Fields(0).Value & "")
            .Close
        End If
        
        
    End With
    
    
End Function

Private Sub Form_Load()
    
    'cmbAccHead.AddVals con, "Head", "Heads", "Code", " Order By Code"
    cmbAccHead.AddItem "< All Heads >"
    cmbAccHead.List(0, 1) = "0"
    'cmbAccHead.ItemData(0) = 0
    Dim rs As New ADODB.Recordset
    With rs
        '.Open "Select Head,Code From Heads Order By Code", con, adOpenForwardOnly, adLockReadOnly
        .Open "SELECT AccNo,AccTitle FROM Accounts ORDER BY AccNo", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            cmbAccHead.AddItem !AccTitle & "" '!Head & ""
            cmbAccHead.List(.AbsolutePosition, 1) = !AccNo & ""
            'cmbAccHead.ItemData(.AbsolutePosition) = !AccNo & "" 'Val(!code & "")
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    cmbAccHead.ListIndex = 0
    
    'PermissionDenied = Not UserHasAccess("TrialBalance")
    
    DtFrom = DateAdd("M", -1, Date)
    DtTo = Date
   
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    Dim TitleBottom As Long
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    TitleBottom = FAcc.Top + FAcc.Height
    LV.Move 50, TitleBottom, Me.ScaleWidth - 100, Me.ScaleHeight - (TitleBottom + cmdClose.Height + fTotals.Height + 100)
    fTotals.Move LV.Left + LV.Width - fTotals.Width, LV.Top + LV.Height
'   cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, LV.Top + LV.Height + 50
'   cmdPrint.Move cmdClose.Left - cmdPrint.Width - 50, cmdClose.Top
    cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, fTotals.Top + fTotals.Height + 50
    cmdPrint.Move cmdClose.Left - cmdPrint.Width - 50, cmdClose.Top
    cmdPrintOP.Move cmdPrint.Left - cmdPrintOP.Width - 50, cmdClose.Top
    cmdClosingBalances.Move cmdPrintOP.Left - cmdClosingBalances.Width - 50, cmdClose.Top
    PicStatus.Move LV.Left, LV.Top + LV.Height + 25
   
End Sub


Private Sub Form_Unload(cancel As Integer)
    Set rsAccounts = Nothing
    
End Sub


Private Sub PrintTrialBal()

    Screen.MousePointer = vbHourglass
    con.BeginTrans
    con.Execute "DELETE FROM TempTrialBalanceNew"
    Dim Parent As Boolean
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            'If .Text = "24-003" Then Stop
            con.Execute "Insert into TempTrialBalanceNew(AccNo,AccTitle,OpenBalDr,OpenBalCr,Debit,Credit,LastBalDr,LastBalCr,Bold) " & _
             " VALUES ('" & .Text & "','" & .SubItems(1) & "'," & Val(.ListSubItems(2).Tag) & "," & Val(.ListSubItems(3).Tag) & "," & _
             Val(.ListSubItems(4).Tag) & "," & Val(.ListSubItems(5).Tag) & "," & Val(.ListSubItems(6).Tag) & "," & _
             Val(.ListSubItems(7).Tag) & "," & Abs(.Bold) & ")"
        End With
    Next

    con.CommitTrans

    'Sleep (500)
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\TrialBalanceNew.rpt")
    Dim frm As New frmPrevRpt
    
    rpt.FormulaFields.GetItemByName("Company").Text = "'" & strCompany & "'"
    rpt.FormulaFields.GetItemByName("Caption").Text = "'" & IIf(chkDetail = vbChecked, "Trial Balance (Detail)", "Trial Balance (Summary)") & "'"
    rpt.FormulaFields.GetItemByName("FromTo").Text = "'From " & Format(DtFrom, "dd/MM/yyyy") & " To " & Format(DtTo, "dd/MM/yyyy") & "'"
    rpt.FormulaFields.GetItemByName("TotalDebit").Text = Val(Format(LV.ListItems(LV.ListItems.count).SubItems(3), "0.00"))
    rpt.FormulaFields.GetItemByName("TotalCredit").Text = Val(Format(LV.ListItems(LV.ListItems.count).SubItems(4), "0.00"))
    rpt.FormulaFields.GetItemByName("Address").Text = "'" & COMPANYADDRESS & "'"
    rpt.FormulaFields.GetItemByName("SummaryReport").Text = Not (chkDetail.Value = vbChecked)
    
    frm.ShowReport " ", rpt
'    CR1.ReportFileName = RptPath & "\TrialBal.rpt"
'    CR1.Connect = ConnectStr
'    CR1.Formulas(0) = "Company='" & strCompany & "'"
'    CR1.Formulas(1) = "Caption='" & IIf(chkDetail = vbChecked, "Trial Balance (Detail)", "Trial Balance (Summary)") & "'"
'    CR1.Formulas(2) = "FromTo='From " & Format(DtFrom, "dd/MM/yyyy") & " To " & Format(DtTo, "dd/MM/yyyy") & "'"
'    CR1.Formulas(3) = "TotalDebit=" & Val(Format(LV.ListItems(LV.ListItems.Count).SubItems(3), "0.00"))
'    CR1.Formulas(4) = "TotalCredit=" & Val(Format(LV.ListItems(LV.ListItems.Count).SubItems(4), "0.00"))
'    CR1.Formulas(5) = "Address='" & COMPANYADDRESS & "'"
'    CR1.Action = 1
'    CR1.PageZoomNext

    Screen.MousePointer = vbDefault
    
End Sub

Private Sub AddHeads()

    Dim rs As New ADODB.Recordset
    Dim TotalOpenBalDebit As Double, TotalOpenBalCredit As Double
    Dim TotalCredit As Double, TotalDebit As Double
    Dim TotalBalanceDebit As Double, TotalBalanceCredit As Double
    
    Dim dOpBal As Double, dOpBalDr As Double, dOpBalCr As Double
    Dim dLastBal As Double, dLastBalDr As Double, dLastBalCr As Double
    rs.CursorLocation = adUseClient
   
    Dim cmd As New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "TBHeadsNewWithCrDr"
    
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters("@FromDate").Value = Format(DtFrom, "dd/MMM/yyyy")
    cmd.Parameters("@ToDate").Value = Format(DtTo, "dd/MMM/yyyy")
    cmd.Parameters("@Code").Value = cmbAccHead.List(cmbAccHead.ListIndex, 1) 'CStr(cmbAccHead.ItemData(cmbAccHead.ListIndex))
    rs.CursorType = adOpenForwardOnly
    rs.LockType = adLockReadOnly
    Set rs = cmd.Execute
    Set cmd = Nothing
    
    With rs
        Dim ITM As ListItem, OpenBal As String, BalanceStr As String, Balance As Double
        Dim Credit As Double, Debit As Double
        dOpBalDr = 0
        dOpBalCr = 0
        For i = 0 To .RecordCount - 1
            If Val(![OpeningBalance] & "") > 0 Then
                OpenBal = Format(Abs(Val(![OpeningBalance] & "")), "#,##0.00") & " DR"
            ElseIf Val(![OpeningBalance] & "") < 0 Then
                OpenBal = Format(Abs(Val(![OpeningBalance] & "")), "#,##0.00") & " CR"
            Else
                OpenBal = Format(Abs(Val(![OpeningBalance] & "")), "#,##0.00") & "      "
            End If
            dOpBalDr = Val(![OpeningBalanceDr] & "")
            dOpBalCr = Val(![OpeningBalanceCr] & "")
            dLastBalDr = Val(!LastBalanceBalanceDr & "")
            dLastBalCr = Val(!LastBalanceBalanceCr & "")
            
            If Rec.State = 1 Then Rec.Close
      
            '''''''''''''''''''''''''''''''''
            '      Debug.Print ![Code]      '
            '''''''''''''''''''''''''''''''''
      
            Rec.Open "Select Sum(Debit) as TotalDebit,Sum(Credit) as TotalCredit from Vouchers WHERE left(LTrim(accno),2)='" & Trim(![code]) & "' and (Credit is not null and Debit is Not null)" & _
             " AND (VDate Between '" & DtFrom & "' AND '" & DtTo & "')", con, adOpenForwardOnly, adLockReadOnly
      
            Debit = Val(Rec![TotalDebit] & "")
            Credit = Val(Rec![TotalCredit] & "")
      
            '''''''''''''''''''' Non Zero Entries  ''''''''''''''''''''''''
      
            If chkNonZero = vbUnchecked And (Debit + Credit) = 0 Then GoTo EndOfLoop
            'dLastBalDr = Val(![OpeningBalanceDr] & "") + Debit
            'dLastBalCr = Val(![OpeningBalanceCr] & "") + Credit
            Balance = Val(![OpeningBalance] & "") - Credit + Debit
            
            If Balance > 0 Then
                BalanceStr = Format(Abs(Balance), "#,##0.00") & " DR"
            ElseIf Val(Balance & "") < 0 Then
                BalanceStr = Format(Abs(Balance), "#,##0.00") & " CR"
            Else
                BalanceStr = Format(Abs(Balance), "#,##0.00") & "      "
            End If
      
            Set ITM = LV.ListItems.add(, ![code] & "'", ![code] & "", , "TYPE")
            ITM.Bold = True
            ITM.ListSubItems.add(, , ![Head] & "").Bold = True
            ITM.ListSubItems(1).ToolTipText = ![Head] & ""
               
            'Start Opening Balance
            ITM.ListSubItems.add(, , Format(Abs(dOpBalDr), "#,##0.00")).Bold = True
            ITM.ListSubItems(2).Tag = dOpBalDr
            ITM.ListSubItems.add(, , Format(Abs(dOpBalCr), "#,##0.00")).Bold = True
            ITM.ListSubItems(3).Tag = dOpBalCr
            'Finish Opening Balance
            
            ITM.ListSubItems.add(, , Format(Debit, "#,##0.00")).Bold = True
            ITM.ListSubItems(4).Tag = Debit
            ITM.ListSubItems.add(, , Format(Credit, "#,##0.00")).Bold = True
            ITM.ListSubItems(5).Tag = Credit
            
            'Start Last Balance
            ITM.ListSubItems.add(, , Format(Abs(dLastBalDr), "#,##0.00")).Bold = True
            ITM.ListSubItems(6).Tag = dLastBalDr
            ITM.ListSubItems.add(, , Format(Abs(dLastBalCr), "#,##0.00")).Bold = True
            ITM.ListSubItems(7).Tag = dLastBalCr
            'Finish Last Balance
            
            'Calculate Totals
            TotalOpenBalDebit = TotalOpenBalDebit + dOpBalDr
            TotalOpenBalCredit = TotalOpenBalCredit + dOpBalCr

            TotalCredit = TotalCredit + Credit
            TotalDebit = TotalDebit + Debit
            
            TotalBalanceDebit = TotalBalanceDebit + dLastBalDr
            TotalBalanceCredit = TotalBalanceCredit + dLastBalCr
EndOfLoop:
            .MoveNEXT
            pbar.Value = pbar.Value + 1
        Next
        .Close
    End With
    Set rs = Nothing
    
    lblTotalOpenBalDebit = Format(TotalOpenBalDebit, "#,##0.00")
    lblTotalOpenBalCredit = Format(TotalOpenBalCredit, "#,##0.00")
 
    lblTotalCredit = Format(TotalCredit, "#,##0.00")
    lblTotalDebit = Format(TotalDebit, "#,##0.00")
    
    lblTotalBalanceDebit = Format(TotalBalanceDebit, "#,##0.00")
    lblTotalBalanceCredit = Format(TotalBalanceCredit, "#,##0.00")

End Sub
