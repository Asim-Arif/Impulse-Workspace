VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form ContLedger 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11910
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
   ScaleHeight     =   8490
   ScaleMode       =   0  'User
   ScaleWidth      =   10667.95
   WindowState     =   2  'Maximized
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
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
      Height          =   285
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
   Begin MSComCtl2.DTPicker DtTo 
      Height          =   300
      Left            =   6690
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
      Format          =   60293123
      CurrentDate     =   37081
   End
   Begin MSComCtl2.DTPicker DtFrom 
      Height          =   300
      Left            =   5310
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
      Format          =   60293123
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
            Picture         =   "ContLedger.frx":0000
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ContLedger.frx":0452
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin VB.Frame FLV 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5835
      Left            =   15
      TabIndex        =   0
      Top             =   795
      Width           =   11385
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
         Height          =   5055
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   8910
         _ExtentX        =   15716
         _ExtentY        =   8916
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
            Text            =   "Date"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item ID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Unit"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Rate"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Qty"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Amount"
            Object.Width           =   1411
         EndProperty
      End
      Begin MSComctlLib.ListView LV2 
         Height          =   5055
         Left            =   9000
         TabIndex        =   8
         Top             =   180
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   8916
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Payment"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   9855
         TabIndex        =   14
         Top             =   5400
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Balance :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   9000
         TabIndex        =   13
         Top             =   5475
         Width           =   750
      End
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   10440
      Picture         =   "ContLedger.frx":0564
      Top             =   240
      Width           =   465
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   360
      Left            =   4560
      TabIndex        =   12
      Top             =   6750
      Visible         =   0   'False
      Width           =   1425
      ForeColor       =   0
      Caption         =   "Print         "
      PicturePosition =   327683
      Size            =   "2514;635"
      Picture         =   "ContLedger.frx":0B7A
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
      Left            =   10155
      TabIndex        =   11
      Top             =   6750
      Width           =   1425
      ForeColor       =   0
      Caption         =   " Close      "
      PicturePosition =   327683
      Size            =   "2514;635"
      Picture         =   "ContLedger.frx":0C8C
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdNew 
      Height          =   360
      Left            =   8685
      TabIndex        =   10
      Top             =   6750
      Width           =   1425
      ForeColor       =   0
      Caption         =   "New Entry   "
      PicturePosition =   327683
      Size            =   "2514;635"
      Picture         =   "ContLedger.frx":0D9E
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   360
      Left            =   7200
      TabIndex        =   9
      Top             =   6750
      Width           =   1425
      ForeColor       =   0
      Caption         =   "Pay. Entry   "
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
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   885
      TabIndex        =   5
      Top             =   165
      Width           =   7215
      BackColor       =   11517387
      Caption         =   "  Department                         Contractor                              From                    To"
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
      Left            =   8175
      TabIndex        =   4
      Top             =   405
      Width           =   1500
      Caption         =   "Refresh        "
      PicturePosition =   327683
      Size            =   "2646;635"
      Picture         =   "ContLedger.frx":0EB0
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
Attribute VB_Name = "ContLedger"
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

Private Sub cmdPrint_Click()

    On Error GoTo ERR
    Call ShowWait
    Dim Cond As String
    
    If cmbShift.ListIndex = 0 Then
        'Cond = " {VAttendance.Shift} in('First','Second') "
        Cond = " {VAttendance.Shift}='First' Or {VAttendance.Shift}='Second'"
    Else
        Cond = " {VAttendance.Shift}='" & cmbShift & "' "
    End If
    
    If cmbEmp.ID = "0" Then
        If cmbDept.ID = "0" Then
            Cond = Cond & " AND ({VAttendance.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & "))"
        Else
            Cond = Cond & " AND ({VAttendance.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & ")) and {VAttendance.DeptID}='" & cmbDept.ID & "'"
        End If
    Else
        Cond = Cond & " and {VAttendance.empid}='" & cmbEmp.ID & "' AND ({VAttendance.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & "))"
    End If
    
With cr1
    .ReportFileName = RptPRPath & "\AttSheet.rpt"
    '.DataFiles(0) = App.Path & "\Payroll.mdb"
    .Connect = ConnectStr
    .SelectionFormula = Cond
    '.SelectionFormula = "{VAttendance.VendID}='" & cmbCustID & "' AND ({VOrdItems.Dated}=date(" & Format(DtFrom.Tag, "yyyy,MM,dd") & ") TO date(" & Format(DtTo.Tag, "yyyy,MM,dd") & "))"
    '.Formulas(0) = "CompanyName='" & getCompanyName & "'"
    .Formulas(0) = "FromTo='From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
    '.Formulas(2) = "Ledger='Ledger of " & cmbCustName & "'"
    '.Formulas(2) = "SubReport=" & LVItems.Visible
    'if Cheques Are Present then Show Cheque Detail Also
    'If LVItems.Visible Then
    '    .SubreportToChange = "ChqLedger"
    '    .DataFiles(0) = DBName
    '    .SelectionFormula = "{VChqLedger.AccNo}='" & cmbCustID & "' AND ({VChqLedger.CDate}=date(" & Format(DtFrom.Tag, "yyyy,MM,dd") & ") TO date(" & Format(DtTo.Tag, "yyyy,MM,dd") & "))"
    'End If
    .Action = 1
    .PageZoomNext
End With

Call HideWait
Exit Sub
ERR:
    MsgBox ERR.Description
    Call HideWait
End Sub

Public Sub cmdRefresh_Click()
    Call ShowWait
    Call RefreshLedger
    Call HideWait
    'Call Form_Resize
End Sub

Private Sub CommandButton1_Click()
    NewPayEntry.Show 1
End Sub

Private Sub Form_Load()

    DtTo = ServerDate
    DtFrom = DateAdd("M", -1, ServerDate)
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "ContDepartments", "DeptID"
    cmbDept.Text = "<All Departments>"
    cmbEmp.Text = "<All Contractors>"
    'Set LV.SmallIcons = ImageList1
    
End Sub
Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    cmbEmp.AddItem "<All Contractors>", "0"
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals con, "Name", "Contractors", "EmpID"
    Else
        cmbEmp.AddVals con, "Name", "Contractors", "EmpID", "Where DeptID='" & cmbDept.ID & "'"
    End If
    
End Sub

Private Sub Form_Resize()

On Error Resume Next
Dim TheHeight As Long


'FAcc.Move (Me.ScaleWidth - FAcc.Width) / 2
 FLV.Move (Me.ScaleWidth - FLV.Width) / 2
'cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 100
'cmdPrint.Move cmdClose.Left - cmdPrint.Width - 50, cmdClose.Top

'FDrag.Move 50, Me.ScaleHeight * 3 / 5, Me.ScaleWidth - 100
 
 'If LVItems.Visible Then
 '  TheHeight = FDrag.Top
 '  FDrag.Visible = True
 'Else
 '  TheHeight = Me.ScaleHeight - cmdClose.Height - 200
 '  FDrag.Visible = False
 'End If
 
 'LV.Move 50, FAcc.Top + FAcc.Height, Me.ScaleWidth - 100, TheHeight - (FAcc.Top + FAcc.Height)
 'LVItems.Move 50, FDrag.Top + FDrag.Height, Me.ScaleWidth - 100, Me.ScaleHeight - (FAcc.Height + cmdClose.Height + LV.Height + 250)
 'FDrag.ZOrder 1
 
End Sub



Private Sub RefreshLedger()
    
    On Error GoTo ERR
    If cmbDept.MatchFound = False Or cmbEmp.MatchFound = False Then Exit Sub
    Dim rsLedger As New ADODB.Recordset
    Dim ITM As ListItem
    Dim Cond As String
    Dim Sql As String
    Dim Balance As Double
    
    If cmbEmp.ID = "0" Then
        If cmbDept.ID = "0" Then
            Cond = " Where "
        Else
            Cond = " Where DeptID='" & cmbDept.ID & "' and"
        End If
    Else
        Cond = " Where empid='" & cmbEmp.ID & "' and"
    End If
    
    Cond = Cond & "  DT between '" & DtFrom & "' and '" & DtTo & "'"
    
    Sql = "select * from vContLedger " & Cond & " Order By EmpID,DT,ItemID Asc"

    With rsLedger
        
        .Open Sql, con, adOpenForwardOnly, adLockReadOnly
        LV.Visible = False
        LV2.Visible = False
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
    
            'Set itm = LV.ListItems.Add(, "'" & ![ItemID] & "", Format(![DT], "ddd-MMM-yyyy"))
            Set ITM = LV.ListItems.Add(, , Format(![DT], "dd-MMM-yyyy"))
            ITM.SubItems(1) = ![ItemID] & ""
            ITM.SubItems(2) = ![ItemName] & ""
            ITM.SubItems(3) = ![Unit] & ""
            'itm.SubItems(4) = ![Description] & ""
            ITM.SubItems(4) = ![Rate] & ""
            ITM.SubItems(5) = ![Qty] & ""
            ITM.SubItems(6) = Val(![Rate]) * Val(![Qty])
            Balance = Balance + Val(ITM.SubItems(6))
            'If IsNull(![Amt]) Then
            '    itm.SubItems(8) = 0
            'Else
            '    itm.SubItems(8) = ![Amt]
            'End If
            .MoveNext
        Next
        .Close
        .Open "Select * from ContPayment" & Cond & " Order by EmpID,DT", con, adOpenForwardOnly, adLockReadOnly
        LV2.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set ITM = LV2.ListItems.Add(, , Format(![DT], "dd-MMM-yyyy"))
            ITM.SubItems(1) = ![Amt] & ""
            Balance = Balance - Val(ITM.SubItems(1))
            .MoveNext
        Next i
        .Close
    End With
    lblBalance = Balance
    LV.Visible = True
    LV2.Visible = True
    Set rsLedger = Nothing
    Sql = ""
    Cond = ""
    Exit Sub
ERR:
    MsgBox ERR.Description
End Sub

