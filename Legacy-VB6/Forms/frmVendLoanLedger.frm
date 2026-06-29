VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmVendLoanLedger 
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11550
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   7515
   ScaleMode       =   0  'User
   ScaleWidth      =   10345.49
   WindowState     =   2  'Maximized
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11520
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vender Loan Ledger"
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
         Left            =   3975
         TabIndex        =   9
         Top             =   120
         Width           =   3225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vender Loan Ledger"
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
         Left            =   3990
         TabIndex        =   10
         Top             =   135
         Width           =   3225
      End
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   330
      Left            =   690
      TabIndex        =   6
      Top             =   870
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
      Left            =   2910
      TabIndex        =   7
      Top             =   870
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
      Top             =   1365
      Width           =   11385
      Begin VB.Frame fButs2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   525
         Left            =   60
         TabIndex        =   14
         Top             =   5475
         Width           =   5505
         Begin MSForms.CommandButton cmdNewLT 
            Height          =   360
            Left            =   2730
            TabIndex        =   16
            Top             =   105
            Width           =   2700
            ForeColor       =   4194304
            Caption         =   "New Long Term Advance          "
            PicturePosition =   327683
            Size            =   "4762;635"
            Picture         =   "frmVendLoanLedger.frx":0000
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
            Left            =   0
            TabIndex        =   15
            Top             =   105
            Width           =   2700
            ForeColor       =   4194304
            Caption         =   "New Short Term Advance"
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
      End
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   540
         Left            =   7080
         TabIndex        =   11
         Top             =   5445
         Width           =   3000
         Begin MSForms.CommandButton cmdClose 
            Height          =   360
            Left            =   1500
            TabIndex        =   13
            Top             =   135
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
            Left            =   30
            TabIndex        =   12
            Top             =   135
            Width           =   1425
            ForeColor       =   0
            Caption         =   "Print         "
            PicturePosition =   327683
            Size            =   "2514;635"
            Picture         =   "frmVendLoanLedger.frx":0452
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
         Left            =   60
         TabIndex        =   1
         Top             =   180
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   4322
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Balance"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Ded/Month"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin MSComCtl2.DTPicker DtTo 
      Height          =   300
      Left            =   6525
      TabIndex        =   2
      Top             =   870
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
      Format          =   3735555
      CurrentDate     =   37081
   End
   Begin MSComCtl2.DTPicker DtFrom 
      Height          =   300
      Left            =   5145
      TabIndex        =   3
      Top             =   870
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
      Format          =   3735555
      CurrentDate     =   37055
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   165
      Top             =   675
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
            Picture         =   "frmVendLoanLedger.frx":0564
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVendLoanLedger.frx":09B6
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   720
      TabIndex        =   5
      Top             =   630
      Width           =   7215
      BackColor       =   11517387
      Caption         =   "  Departement                      Employee                               From                    To"
      Size            =   "12726;423"
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
      Left            =   7965
      TabIndex        =   4
      Top             =   855
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
Attribute VB_Name = "frmVendLoanLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql As String
Dim Openbalance As Double
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNewLT_Click()

    With frmVendNewAdv
      .LongTerm = True
      .Show 1
      Call cmdRefresh_Click
   End With
    
End Sub

Private Sub cmdNewST_Click()

    With frmVendNewAdv
      .LongTerm = False
      .Show 1
      Call cmdRefresh_Click
    End With
    
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    Screen.MousePointer = vbHourglass
    Dim Cond As String
    
    Cond = "{VAdvances.empid}='" & cmbEmp.ID & "' AND ({VAdvances.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & "))"
    
    
    
    With Cr1
        .ReportFileName = RptPRPath & "\LTLoanLedger.rpt"
        '.DataFiles(0) = rptPRPath & "\Payroll.mdb"
        .Connect = ConnectStr
        .SelectionFormula = Cond
        
        .Formulas(0) = "FromTo='From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
        .Formulas(1) = "PrevBal=" & Openbalance
        
        .Action = 1
        .PageZoomNext
    End With

    Screen.MousePointer = vbDefault
Exit Sub
err:
    MsgBox err.Description
    Screen.MousePointer = vbDefault
End Sub

Public Sub cmdRefresh_Click()
    

    Call RefreshLedger

    'Call Form_Resize
End Sub

Private Sub Form_Load()

    DtTo = ServerDate
    DtFrom = DateAdd("M", -1, ServerDate)
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    

    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Processes>", "0"
    cmbDept.AddVals con, "Description", "Processes", "ProcessID"
    cmbDept.Text = "<All Processes>"
    
    
End Sub
Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    'cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Processes>" Then
        cmbEmp.AddVals con, "VenderName", "Venders", "VendID"
    Else
        cmbEmp.AddVals con, "VenderName", "Venders", "VendID", "Where PhaseID=" & cmbDept.ID
    End If
    'cmbEmp.Text = "<All Venders>"
End Sub

Private Sub Form_Resize()

    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
    
    FLV.Move 50, 1230, Me.ScaleWidth - 100, Me.ScaleHeight - 1400
    
    LV.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
    
    fButs1.Move LV.Left + LV.Width - fButs1.Width, LV.Top + LV.Height
    
    fButs2.Move LV.Left, LV.Top + LV.Height
    
End Sub



Private Sub RefreshLedger()

    If cmbDept.MatchFound = False Or cmbEmp.MatchFound = False Then Exit Sub
    Dim rsLedger As New ADODB.Recordset
    Dim ITM As ListItem

    
    Sql = "select * from VVendLoanLedger  Where VendID=" & cmbEmp.ID & " and Dt Between '" & Format(DtFrom, "dd-MMM-yyyy") & "' and  '" & Format(DtTo, "dd-MMM-yyyy") & "' Order By VendID,DT Asc"

    With rsLedger
        
        .Open Sql, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        
        Sql = "(Select IsNull(Sum(AmtClrd),0) balance From VendAmtCleared Where Dt<'" & Format(DtFrom, "dd-MMM-yyyy") & "' and VendID=" & cmbEmp.ID & ")"
        
           Balance = Val(con.Execute("Select Sum(Amount)-" & Sql & "  from VenderAdvances Where VendID=" & cmbEmp.ID & " and Type=1 and  Dt<'" & Format(DtFrom, "dd-MMM-yyyy") & "'").Fields(0) & "")
           Openbalance = Balance
        Set ITM = LV.ListItems.Add()
        ITM.SubItems(1) = "Previous Balance "
        ITM.SubItems(3) = Balance
        
        For i = 0 To .RecordCount - 1
                    
            Set ITM = LV.ListItems.Add(, , Format(![DT], "dddd dd MMMM ,yyyy"))
            
            ITM.SubItems(1) = ![Description] & ""
            ITM.SubItems(2) = Val(![Amount] & "")
            
            Balance = Balance + Val(![Amount] & "")
            ITM.SubItems(3) = Balance
            ITM.SubItems(4) = IIf(Val(![dAmount] & "") = 0, "", ![dAmount])
            
            .MoveNext
        Next
        .Close
    End With
'    If cmbShift.ListIndex = 0 Then
'        LV.ColumnHeaders(7).Width = 0
'    Else
'            LV.ColumnHeaders(7).Width = 1200
'    End If
    Set rsLedger = Nothing
    Sql = ""
    Cond = ""
    ICO = ""
End Sub
