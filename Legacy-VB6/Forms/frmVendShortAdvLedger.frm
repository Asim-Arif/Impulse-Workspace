VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmVendShortAdvLedger 
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11550
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   6975
   ScaleMode       =   0  'User
   ScaleWidth      =   10345.49
   WindowState     =   2  'Maximized
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   330
      Left            =   855
      TabIndex        =   6
      Top             =   405
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Height          =   330
      Left            =   3075
      TabIndex        =   7
      Top             =   405
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin VB.Frame FLV 
      Height          =   6045
      Left            =   30
      TabIndex        =   0
      Top             =   780
      Width           =   11385
      Begin VB.Frame fButs2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   510
         Left            =   165
         TabIndex        =   12
         Top             =   5460
         Width           =   5535
         Begin MSForms.CommandButton cmdNewLT 
            Height          =   360
            Left            =   2760
            TabIndex        =   14
            Top             =   90
            Width           =   2700
            ForeColor       =   4194304
            Caption         =   "New Long Term Advance          "
            PicturePosition =   327683
            Size            =   "4762;635"
            Accelerator     =   78
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdNewST 
            Height          =   360
            Left            =   30
            TabIndex        =   13
            Top             =   90
            Width           =   2700
            ForeColor       =   4194304
            Caption         =   "New Short Term Advance"
            PicturePosition =   327683
            Size            =   "4762;635"
            Picture         =   "frmVendShortAdvLedger.frx":0000
            Accelerator     =   78
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   465
         Left            =   8295
         TabIndex        =   9
         Top             =   5475
         Width           =   3030
         Begin MSForms.CommandButton cmdClose 
            Height          =   360
            Left            =   1530
            TabIndex        =   11
            Top             =   75
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
         Begin MSForms.CommandButton cmdPrint 
            Height          =   360
            Left            =   60
            TabIndex        =   10
            Top             =   75
            Width           =   1425
            ForeColor       =   0
            Caption         =   "Print         "
            PicturePosition =   327683
            Size            =   "2514;635"
            Picture         =   "frmVendShortAdvLedger.frx":0112
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
      Begin Crystal.CrystalReport cr1 
         Left            =   3840
         Top             =   195
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5265
         Left            =   105
         TabIndex        =   1
         Top             =   195
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   9287
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Emp ID"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emp Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   4322
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Loan Type"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Balance"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Ded/Month"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin MSComCtl2.DTPicker DtTo 
      Height          =   300
      Left            =   8220
      TabIndex        =   2
      Top             =   435
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   22609923
      CurrentDate     =   37081
   End
   Begin MSComCtl2.DTPicker DtFrom 
      Height          =   300
      Left            =   6840
      TabIndex        =   3
      Top             =   435
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   22609923
      CurrentDate     =   37055
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVendShortAdvLedger.frx":0224
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVendShortAdvLedger.frx":0676
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   885
      TabIndex        =   5
      Top             =   165
      Width           =   8685
      BackColor       =   11517387
      Caption         =   $"frmVendShortAdvLedger.frx":0788
      Size            =   "15319;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cmbShift 
      Height          =   285
      Left            =   5340
      TabIndex        =   8
      Top             =   435
      Width           =   1455
      VariousPropertyBits=   679495705
      DisplayStyle    =   7
      Size            =   "2566;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontEffects     =   1073750016
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   9735
      TabIndex        =   4
      Top             =   405
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
End
Attribute VB_Name = "frmVendShortAdvLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql As String
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNewLT_Click()
    With NewAdvanceShort
      .LongTerm = True
      .Show 1
      Call cmdRefresh_Click
   End With
    
End Sub

Private Sub cmdNewST_Click()
With NewAdvanceShort
      .LongTerm = False
      .Show 1
      Call cmdRefresh_Click
   End With
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo ERR
    Screen.MousePointer = vbHourglass
    Dim Cond As String
    
    
    
    If cmbEmp.ID = "0" Then
        If cmbDept.ID = "0" Then
            Cond = "({VAdvances.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & "))"
        Else
            Cond = "({VAdvances.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & ")) and {VAdvances.DeptID}='" & cmbDept.ID & "'"
        End If
    Else
        Cond = "{VAdvances.empid}='" & cmbEmp.ID & "' AND ({VAdvances.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & "))"
    End If
    
    
    If cmbShift.ListIndex > 0 Then
      Cond = Cond & " and {VADvances.Type}=" & cmbShift.ListIndex
         
    End If
With cr1
    .ReportFileName = RptPRPath & "\LoanLedger.rpt"
    '.DataFiles(0) = rptPRPath & "\Payroll.mdb"
    .Connect = ConnectStr
    .SelectionFormula = Cond
    
    .Formulas(0) = "FromTo='From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
    
    .Action = 1
    .PageZoomNext
End With

Screen.MousePointer = vbDefault
Exit Sub
ERR:
    MsgBox ERR.Description
    Screen.MousePointer = vbDefault
End Sub

Public Sub cmdRefresh_Click()
    

    Call RefreshLedger

    'Call Form_Resize
End Sub

Private Sub Form_Load()

'cmdNewST.Enabled = UserHasAccess("ShortTermLoan")
'cmdNewLT.Enabled = UserHasAccess("LongTermLoan")

    DtTo = ServerDate
    DtFrom = DateAdd("M", -1, ServerDate)
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    
    'cmbShift.AddItem "<All>"
    cmbShift.AddItem "Short Term"
    cmbShift.AddItem "Long Term"
    cmbShift.ListIndex = 0
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Processes>", "0"
    cmbDept.AddVals con, "Description", "Processes", "ProcessID"
    cmbDept.Text = "<All Processes>"
    
    'Set LV.SmallIcons = ImageList1
    
End Sub
Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    cmbEmp.AddItem "<All Venders>", "0"
    If cmbDept.Text = "<All Processes>" Then
        cmbEmp.AddVals con, "VenderName", "Venders", "VendID"
    Else
        cmbEmp.AddVals con, "VenderName", "Venders", "VendID", "Where PhaseID=" & cmbDept.ID
    End If
    cmbEmp.Text = "<All Venders>"
End Sub

Private Sub Form_Resize()

On Error Resume Next
 
 FLV.Move 50, 800, Me.ScaleWidth - 100, Me.ScaleHeight - 900
 
 LV.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
 
 fButs1.Move LV.Left + LV.Width - fButs1.Width, LV.Top + LV.Height
 
 FButs2.Move LV.Left, LV.Top + LV.Height

 
End Sub



Private Sub RefreshLedger()

    If cmbDept.MatchFound = False Or cmbEmp.MatchFound = False Then Exit Sub
    Dim rsLedger As New ADODB.Recordset
    Dim ITM As ListItem
    Dim Cond As String
    
    Dim ICO As String
    
    
    Cond = " Where Type=" & Val(cmbShift.ListIndex)
    
    
    If cmbEmp.ID = "0" Then
        If cmbDept.ID = "0" Then
            Cond = Cond & " and DT between '" & DtFrom & "' and '" & DtTo & "' "
        Else
            Cond = Cond & " and DT between '" & DtFrom & "' and '" & DtTo & "' and PhaseID=" & cmbDept.ID
        End If
    Else
        Cond = Cond & " and VendID=" & cmbEmp.ID & " and DT between '" & DtFrom & "' and '" & DtTo & "' "
    End If
    
    Sql = "select * from VVenderAdvances " & Cond & " Order By VendID,DT,Type Asc"

    With rsLedger
        
        .Open Sql, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            'If ![Attendance] Then
            '    ICO = "Present"
            'Else
            '    ICO = "Absent"
            'End If
            Set ITM = LV.ListItems.Add(, "'" & ![EntryID] & "", ![VendID1] & "")
            ITM.SubItems(1) = ![VenderName] & ""
            ITM.SubItems(2) = Format(![DT], "dddd dd MMMM ,yyyy")
            ITM.SubItems(3) = IIf(![Type] = 0, "Short Term", "Long Term")
            ITM.SubItems(4) = ![Amount] & ""
            ITM.SubItems(5) = Val(![Amount] & "") - Val(![AmountCleared] & "")
            ITM.SubItems(6) = Val(![dAmount] & "")
            .MoveNext
        Next
        .Close
    End With
    If cmbShift.ListIndex = 0 Then
        LV.ColumnHeaders(7).Width = 0
    Else
            LV.ColumnHeaders(7).Width = 1200
    End If
    Set rsLedger = Nothing
    Sql = ""
    Cond = ""
    ICO = ""
End Sub




