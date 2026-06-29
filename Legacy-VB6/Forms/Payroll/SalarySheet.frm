VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form SalarySheet 
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12525
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   7785
   ScaleMode       =   0  'User
   ScaleWidth      =   11218.81
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport CR 
      Left            =   1080
      Top             =   2580
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
   Begin VB.TextBox txtTotal 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7020
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   855
      Width           =   960
   End
   Begin VB.CheckBox chkExternal 
      Caption         =   "External Salary Sheet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9690
      TabIndex        =   12
      Top             =   915
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11670
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Sheet"
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
         Left            =   4680
         TabIndex        =   9
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Sheet"
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
         Left            =   4695
         TabIndex        =   10
         Top             =   135
         Width           =   1935
      End
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   45
      TabIndex        =   4
      Top             =   855
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
      Left            =   2265
      TabIndex        =   5
      Top             =   855
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
            Picture         =   "SalarySheet.frx":0000
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SalarySheet.frx":0452
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin VB.Frame FLV 
      Height          =   6120
      Left            =   75
      TabIndex        =   0
      Top             =   1215
      Width           =   12060
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   5070
         TabIndex        =   11
         Top             =   690
         Visible         =   0   'False
         Width           =   1155
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
         Height          =   5010
         Left            =   165
         TabIndex        =   1
         Top             =   135
         Width           =   11685
         _ExtentX        =   20611
         _ExtentY        =   8837
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
         NumItems        =   21
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Emp ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emp Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "B. Sal"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Rate"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "A.Days"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "A. Amt"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "S.days"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "S. Amt"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "O.Hrs"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "O.Amt"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "L.Hrs"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "L.Amt"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   12
            Text            =   "Total"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   13
            Text            =   "Tax %"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   14
            Text            =   "T. Amt"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   15
            Text            =   "Net Ttl"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   16
            Text            =   "ST Loan"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   17
            Text            =   "Loan Ded."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   18
            Text            =   "Adv Sal"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   19
            Text            =   "S. Sec."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   20
            Text            =   "Balance"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   945
         Left            =   8325
         TabIndex        =   14
         Top             =   5145
         Width           =   3585
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Total Salary Amount :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   120
            TabIndex        =   20
            Top             =   135
            Width           =   1875
         End
         Begin MSForms.CommandButton cmdCancel 
            Height          =   360
            Left            =   1995
            TabIndex        =   18
            Top             =   510
            Visible         =   0   'False
            Width           =   1425
            ForeColor       =   0
            Caption         =   " Cancel     "
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
         Begin MSForms.CommandButton cmdClose 
            Height          =   360
            Left            =   1995
            TabIndex        =   17
            Top             =   510
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
            Left            =   450
            TabIndex        =   16
            Top             =   510
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
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00 "
            Height          =   285
            Left            =   2025
            TabIndex        =   15
            Top             =   75
            Width           =   1425
         End
         Begin MSForms.CommandButton cmdSave 
            Height          =   360
            Left            =   450
            TabIndex        =   19
            Top             =   510
            Visible         =   0   'False
            Width           =   1425
            ForeColor       =   0
            Caption         =   "  Save       "
            PicturePosition =   327683
            Size            =   "2514;635"
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
      Begin VB.Frame FButs2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   825
         Left            =   150
         TabIndex        =   21
         Top             =   5205
         Width           =   5040
         Begin VB.CommandButton Command1 
            Caption         =   "Pay All"
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
            Left            =   2355
            TabIndex        =   23
            Top             =   405
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Unpay All"
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
            Left            =   3660
            TabIndex        =   22
            Top             =   405
            Visible         =   0   'False
            Width           =   1275
         End
         Begin MSForms.CommandButton cmdUpdate 
            Height          =   360
            Left            =   45
            TabIndex        =   24
            Top             =   420
            Visible         =   0   'False
            Width           =   2265
            ForeColor       =   0
            Caption         =   "Save Changes"
            PicturePosition =   327683
            Size            =   "3995;635"
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
   End
   Begin MSForms.ComboBox cmbMonth 
      Height          =   300
      Left            =   5685
      TabIndex        =   7
      Top             =   870
      Width           =   1320
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2328;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbYear 
      Height          =   300
      Left            =   4530
      TabIndex        =   6
      Top             =   870
      Width           =   1110
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "1958;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   75
      TabIndex        =   3
      Top             =   615
      Width           =   7935
      BackColor       =   11517387
      Caption         =   "  Department                         Employee                              Year                  Month                Total Days"
      Size            =   "13996;423"
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
      Left            =   8025
      TabIndex        =   2
      Top             =   855
      Width           =   1605
      Caption         =   "Refresh        "
      PicturePosition =   327683
      Size            =   "2831;635"
      Accelerator     =   82
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Menu mnupop 
      Caption         =   "mnupop"
      Visible         =   0   'False
      Begin VB.Menu mnuSalSheet 
         Caption         =   "Print Salary Sheet (Detail)"
      End
      Begin VB.Menu PrintSalsummary 
         Caption         =   "Print Salary Sheet (Summary)"
      End
      Begin VB.Menu ln1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCurrency 
         Caption         =   "Print Currency Denomination"
      End
      Begin VB.Menu dash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPaySlips 
         Caption         =   "Print PaySlips"
      End
   End
End
Attribute VB_Name = "SalarySheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Thousands As Double, FHundreds As Double, Hundreds As Double, _
Fifties As Double, Tens As Double, Fives As Double, TotalSal As Double, _
Twos As Double, Ones As Double

Dim EditAbsents As Boolean
Dim MakeFinal As Integer
Dim rsLedger As New ADODB.Recordset

Private Sub chkExternal_Click()
   Call cmdRefresh_Click
End Sub

Private Sub cmbMonth_Click()
    txtTotal = TotalMonthDays
End Sub

Private Sub cmbYear_Change()
    txtTotal = TotalMonthDays
End Sub

Private Sub cmdCancel_Click()
    cmdRefresh.Visible = True
    cmdClose.Visible = True
    cmdPrint.Visible = True
    cmdSave.Visible = False
    cmdCancel.Visible = False
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub PrintReport(PaySlips As Boolean, Optional Crpt As CrystalReport = Nothing, Optional Summary As Boolean)

On Error GoTo err
If LV.ListItems.Count = 0 Then Exit Sub
Call ShowWait
Dim DT As Date
DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)

Dim i As Integer
Dim j As Integer
con.Execute "Delete from PrintSalary"
Dim rs As New ADODB.Recordset
With rs
    .Open "select * from PrintSalary", con, adOpenStatic, adLockOptimistic
    For i = 1 To LV.ListItems.Count
        .AddNew
        ![EmpId] = LV.ListItems(i).Text & ""
        For j = 1 To 20
            .Fields(j) = LV.ListItems(i).SubItems(j) & ""
        Next j
        ![DT] = DT & ""
        ![PrevLtLoan] = Val(LV.ListItems(i).ListSubItems(1).Tag)
        .Update
    Next i
End With
Dim unloadflag As Boolean
'Sleep (1000)
'Dim TDate As Date
'Dim MonthDays As Integer
'TDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
'MonthDays = DateAdd("m", 1, TDate) - TDate
If Crpt Is Nothing Then
    Set Crpt = cr1
    unloadflag = False
Else
    unloadflag = True
End If
With Crpt
    
    If Not PaySlips Then
    If chkExternal = vbChecked Then
      .ReportFileName = RptPRPath & "\ExtSalSheet.rpt"
      .Formulas(1) = "Daysinmonth=" & Val(txtTotal)
    ElseIf Summary Then
      .ReportFileName = RptPRPath & "\SalarySummary.rpt"
      .Formulas(1) = ""
    Else
      .ReportFileName = RptPRPath & "\SalarySheet.rpt"
      .Formulas(1) = ""
    End If
    '.DataFiles(0) = DatabasePath
      .Connect = ConnectStr
      .Formulas(0) = "Fromto='For " & cmbMonth.Tag & "-" & cmbYear.Tag & IIf(chkExternal = vbChecked, "", " ( For Factory Use Only )") & "'"
      
    Else
        
            .ReportFileName = RptPRPath & "\PaySlip.rpt"
            
            '.DataFiles(0) = DatabasePath
            .Connect = ConnectStr
            .Formulas(0) = "ForMonth='For " & cmbMonth.Tag & "-" & cmbYear.Tag & "'"
            
    End If
    
    '.Formulas(1) = "ForMonthDays=" & MonthDays
    .Action = 1
    .PageZoomNext
End With

Call HideWait

If unloadflag Then Unload Me

Exit Sub
err:
    MsgBox err.Description
    Call HideWait
End Sub

Private Sub cmdPrint_Click()

Me.PopUpMenu mnupop, 2, cmdPrint.Left + FLV.Left + fButs1.Left, fButs1.Top + cmdPrint.Top + FLV.Top + cmdPrint.Height

End Sub

Public Sub cmdRefresh_Click()
    Call ShowWait
    Call RefreshLedger
     
    MakeFinal = 0
    
    Command1.Visible = LV.Checkboxes
    Command2.Visible = LV.Checkboxes
    Call HideWait
    'Call Form_Resize
End Sub



Private Sub cmdSave_Click()
    'Call SaveIt(rsLedger)
    cmdRefresh.Visible = True
    cmdClose.Visible = True
    cmdPrint.Visible = True
    cmdSave.Visible = False
    cmdCancel.Visible = False
End Sub

Private Sub cmdUpdate_Click()
    
If cmdUpdate.Caption = "Make This Sheet Final" Then

    With rsLedger
   If .RecordCount = 0 Then Exit Sub Else .MoveFirst
   
   con.BeginTrans
   Dim i As Integer
   For i = 1 To LV.ListItems.Count
      .Find "EMPID='" & LV.ListItems(i) & "'", 0, adSearchForward, 1
      Call SaveIt(rsLedger, 0, i)
      .MoveNext
   Next
   
    con.CommitTrans
   
    Call cmdRefresh_Click
    
    End With
    LV.Checkboxes = True
    Command1.Visible = True
    Command2.Visible = True
    cmdUpdate.Caption = "Save Changes"
Else
   
  
    
    
    
    Dim setVal As Integer
    Dim EmpId As String
    Dim TheMonth As Byte
    Dim TheYear As Integer
    
    con.BeginTrans
    
    For i = 1 To LV.ListItems.Count
        setVal = Abs(LV.ListItems(i).Checked)
        EmpId = Left(LV.ListItems(i).key, 7)
        TheMonth = Month(CDate(Right(LV.ListItems(i).key, Len(LV.ListItems(i).key) - 7)))
        TheYear = year(CDate(Right(LV.ListItems(i).key, Len(LV.ListItems(i).key) - 7)))
        
        con.Execute "Update MonthlySalaries Set Payed=" & setVal & ",NetSalary=" & Val(LV.ListItems(i).SubItems(20)) + Val(LV.ListItems(i).SubItems(18)) & _
        ",OTAmt=" & Val(LV.ListItems(i).SubItems(9)) & ",StAmt=" & Val(LV.ListItems(i).SubItems(16)) & _
        ",LTAmt=" & Val(LV.ListItems(i).SubItems(17)) & ",LateAmt=" & Val(LV.ListItems(i).SubItems(11)) & _
        ",salAmt=" & Val(LV.ListItems(i).SubItems(7)) & _
        " Where EmpID='" & EmpId & "' and month(dt)=" & TheMonth & " and Year(dt)=" & TheYear & " and AdvSalary=0"
        
        AmtClrd = Val(LV.ListItems(i).SubItems(17))
        con.Execute "Update AmtCleared Set AmtClrd=" & AmtClrd & " Where EmpID='" & EmpId & "' and month(dt)=" & TheMonth & " and Year(dt)=" & TheYear
        If LV.ListItems(i).Checked Then
            con.Execute "Update Advances Set amountcleared=amount Where EmpId='" & EmpId & "' and type =0 and month(dt)=" & TheMonth & " and Year(dt)=" & TheYear
            
        Else
            con.Execute "Update Advances Set amountcleared=0 Where EmpId='" & EmpId & "' and type =0 and month(dt)=" & TheMonth & " and Year(dt)=" & TheYear
        End If
    Next i
    
    con.CommitTrans
End If
       
End Sub

Private Sub Command1_Click()
For i = 1 To LV.ListItems.Count
    LV.ListItems(i).Checked = True
Next
End Sub

Private Sub Command2_Click()
For i = 1 To LV.ListItems.Count
    LV.ListItems(i).Checked = False
Next
End Sub

Private Sub Form_Load()
cmdUpdate.Visible = UserHasAccess("FinalSalSheet")
    'LV.ColumnHeaders(14).Width = 0
    'LV.ColumnHeaders(15).Width = 0

    DtTo = ServerDate
    DtFrom = DateAdd("M", -1, ServerDate)
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    
    
    cmbMonth.AddItem "Jan"
    cmbMonth.List(0, 1) = 1
    cmbMonth.AddItem "Feb"
    cmbMonth.List(1, 1) = 2
    cmbMonth.AddItem "Mar"
    cmbMonth.List(2, 1) = 3
    cmbMonth.AddItem "Apr"
    cmbMonth.List(3, 1) = 4
    cmbMonth.AddItem "May"
    cmbMonth.List(4, 1) = 5
    cmbMonth.AddItem "Jun"
    cmbMonth.List(5, 1) = 6
    cmbMonth.AddItem "Jul"
    cmbMonth.List(6, 1) = 7
    cmbMonth.AddItem "Aug"
    cmbMonth.List(7, 1) = 8
    cmbMonth.AddItem "Sep"
    cmbMonth.List(8, 1) = 9
    cmbMonth.AddItem "Oct"
    cmbMonth.List(9, 1) = "10"
    cmbMonth.AddItem "Nov"
    cmbMonth.List(10, 1) = 11
    cmbMonth.AddItem "Dec"
    cmbMonth.List(11, 1) = 12
    
    Dim CYear As Long
    CYear = Format(ServerDate, "yyyy")
    For i = 2002 To CYear
        cmbYear.AddItem i
    Next i
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID"
    cmbDept.Text = "<All Departments>"
    
    cmbEmp.Text = "<All Employees>"
    'Set LV.SmallIcons = ImageList1
    cmbMonth.ListIndex = Val(Format$(ServerDate, "MM")) - 1
    cmbYear = Format$(ServerDate, "yyyy")
    
    
    'Call cmdRefresh_Click
    
End Sub
Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals con, "Name", "Employees", "EmpID", "WHERE Vender=0"
    Else
        cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' AND Vender=0"
    End If
    cmbEmp.ID = 0
    
End Sub

Private Sub Form_Resize()

On Error Resume Next

 fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2

 FLV.Move 50, 1230, Me.ScaleWidth - 100, Me.ScaleHeight - 1400
 
 LV.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
 
 fButs1.Move LV.Left + LV.Width - fButs1.Width, LV.Top + LV.Height
 
 FButs2.Move LV.Left, LV.Top + LV.Height
 
End Sub



Private Sub RefreshLedger()
Dim GrandTotal As Double

    If cmbDept.MatchFound = False Or cmbEmp.MatchFound = False Then Exit Sub
    If cmbYear.MatchFound = False Or cmbMonth.MatchFound = False Then Exit Sub
    
    Dim SunDates() As Date
        
    Dim EmpId As String
    Dim TotalSalDays As Double
    Dim TempTotalSaldays As Double
    Dim LastDay As Integer
    
    Dim ITM As ListItem
    Dim Cond As String
    Dim Sql As String
    Dim cmd As New ADODB.Command
    
    
    
    
    
    If cmbEmp.ID = "0" Then
       EmpId = IIf(cmbDept.ID = "0", "", cmbDept.ID)
    Else
       EmpId = IIf(cmbEmp.ID = "0", "", cmbEmp.ID)
    End If
    
    Dim rsCheck As New ADODB.Recordset
    Dim TDate As Date
    Dim SDate As Date 'For Saving
    'TDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), Day(ServerDate))
    
    SDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    TDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), Day(ServerDate))
    
    TotalSalDays = DateAdd("m", 1, SDate) - SDate
    
    
    
    LastDay = TotalSalDays
    
    Dim JDate As Date
    If TDate = ServerDate Then
        TotalSalDays = Day(ServerDate)
    ElseIf ServerDate < TDate Then
        MsgBox "Can't Show Entries Of Future Months.", vbInformation
        Exit Sub
    End If
    
    TempTotalSaldays = TotalSalDays    'Store For Later Use
    
    JDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), TotalSalDays)
        
    'Sql = "Select * from VAdvanceSalary " & Cond & " JoinDate<='" & Format(JDate, "dd-MMM-yyyy") & "' Order By DeptID,EmpID"
    
    'cmd.CommandText = Sql
    
    
    
    
    cmd.CommandType = adCmdStoredProc
    
    cmd.CommandText = "VAdvanceSalary"
    
    Set cmd.ActiveConnection = con
    
    cmd.Parameters.Refresh
    
    cmd.Parameters(1).Type = adDBTimeStamp
    cmd.Parameters(1).value = JDate

    cmd.Parameters(2).Type = adVarChar
    cmd.Parameters(2).value = EmpId
    
    
    If rsLedger.State = 1 Then rsLedger.Close
    rsLedger.CursorType = adOpenForwardOnly
    rsLedger.LockType = adLockReadOnly
    Set rsLedger = cmd.Execute
    With rsLedger
        
        '.Open Sql, Con, adOpenForwardOnly, adLockReadOnly
        
        Dim TotalSuns As Integer
        Dim SalDays As Double
        Dim Rate As Double
        Dim AbsAllowed As Integer
        Dim AbsentDays As Double
        
        
        Dim Total As Double
        Dim OverTimeHours As Double
        Dim LateHours As Double
        Dim TaxPer As Double
        Dim DedAmt As Double
        Dim Payed As Double
        Dim LongLoan As Double
        Dim TaxSal As Double
        Dim IsComputed As Boolean
        
        Dim NTime As Integer
        Dim LTime As Double
        Dim DT As String
        Dim IsPayed As Boolean
        Dim IsTransfered As Boolean
        Dim Holidays As Integer
        Dim rs As New ADODB.Recordset
        Dim EmpSalary As Double
        Dim EmpExtSal As Double
        
        Dim SocialAmt As Double

        
       
        
        LV.ListItems.Clear
        LV.Visible = False
        For i = 0 To .RecordCount - 1
                    
                    'Reset Total Salary days For Each Employee
                    TotalSalDays = DateAdd("m", 1, SDate) - SDate
                    If TDate = ServerDate Then
                        TotalSalDays = Day(ServerDate)
                    End If
                    
                    
                    
                                        
                    'Check that Employee Discontinued In This Month
                    
               If Not IsNull(![DiscontinuedOn]) Then
                   If Month(![DiscontinuedOn]) = Month(JDate) And ![DiscontinuedOn] <= JDate Then
                     TotalSalDays = Day(![DiscontinuedOn])
                   End If
               Else
                  TotalSalDays = TempTotalSaldays
               End If
                    
                    
           
           'if Joined This Month Calculate From Join Date
           
           If Format(![JoinDate], "MMM-yyyy") = Format(JDate, "MMM-yyyy") Then
                rs.Open "Select count(DT) from Holidays Where DT Between '" & ![JoinDate] & "' and '" & JDate & "'", con, adOpenForwardOnly, adLockReadOnly
                  If rs.RecordCount > 0 Then
                     Holidays = rs.Fields(0) & ""
                  End If
                rs.Close
               Set rs = Nothing
               
               SundayAbsents = getSundayAbsents(![EmpId] & "", Val(cmbMonth.List(cmbMonth.ListIndex, 1)), cmbYear, ![WeeklyAbs], Int(TotalSalDays))
               
               TotalSuns = TotalSundays(cmbMonth.List(cmbMonth.ListIndex, 1), cmbYear, TotalSalDays, Day(![JoinDate]))
               
               Holidays = TotalSuns + Holidays - SundayAbsents
            Else
                 rs.Open "Select count(DT) from Holidays Where convert(char(3),DT,7)='" & cmbMonth & "' and Year(DT)=" & cmbYear & " and day(dt)<" & TotalSalDays, con, adOpenForwardOnly, adLockReadOnly
                  If rs.RecordCount > 0 Then
                     Holidays = rs.Fields(0) & ""
                  End If
                 rs.Close
               Set rs = Nothing
               
               
               SundayAbsents = getSundayAbsents(![EmpId], cmbMonth.List(cmbMonth.ListIndex, 1), cmbYear, Int(TotalSalDays), ![WeeklyAbs])
               
               TotalSuns = TotalSundays(cmbMonth.List(cmbMonth.ListIndex, 1), cmbYear, TotalSalDays)
                                              
               Holidays = TotalSuns + Holidays - SundayAbsents
               
               
               
            End If
                    
                    
            'Get Sundays Not To Be Paid
            
            
                    
                    
            SocialAmt = IIf((!Social = True), Val(!SocialAmt & ""), 0)
            
            Call Computed(![EmpId], IsComputed, TaxPer, AbsAllowed, OverTimeHours, NTime, LTime, LongLoan, DT, IsPayed, IsTransfered, JDate, EmpSalary, EmpExtSal, ![Active], SocialAmt)
            
            If Not IsComputed Then
                'Tax %
                If IsNull(![PerDeduct]) Then TaxPer = 0 Else TaxPer = Val(![PerDeduct]) & ""
                
                'Absent Days....
                AbsAllowed = Val(![AbsAllowed] & "")
                
                'OverTime Hours
                If IsNull(![OverTimeHours]) Then OverTimeHours = 0 Else OverTimeHours = ![OverTimeHours]
                
                'Late Hours
                NTime = Val(![NTimes] & "")
                LTime = Val(![TimeAllowed] & "")
                
                'Long Term Loan Deduction
                Dim lDedAmt As Long, lClsAmt As Long, lTotAmt As Long
                If IsNull(![DedAmt]) Then lDedAmt = 0 Else lDedAmt = ![DedAmt]
                If IsNull(![totamt]) Then lTotAmt = 0 Else lTotAmt = ![totamt]
                If IsNull(![clsamt]) Then lClsAmt = 0 Else lClsAmt = ![clsamt]
                
                If lDedAmt > (lTotAmt - lClsAmt) Then
                    LongLoan = (lTotAmt - lClsAmt)
                Else
                    LongLoan = lDedAmt
                End If
                                      
                'Get The Salary Depending On Salary Sheet Type
                'EmpSalary = IIf(chkExternal = vbUnchecked, Val(![StartingSalary] & ""), Val(![Salary] & ""))
                 EmpSalary = Val(![StartingSalary] & "")
                 EmpExtSal = Val(![Salary] & "")
            End If
            
            
            
            
            
            Dim theKey As String
            If IsComputed Then
                theKey = ![EmpId] & Format(SDate, "dd-MMM-yyyy")
            Else
                theKey = ![EmpId] & ""
            End If
            Set ITM = LV.ListItems.Add(, theKey, ![EmpId] & "")
            If IsComputed Then
                ITM.Checked = IsPayed
                ITM.Bold = IsTransfered
            End If
            
            ITM.Tag = Val(![EmpType] & "")
            
            
            
            ITM.ListSubItems.Add(, , ![Name] & "").Tag = Val(![totamt] & "") - Val(![clsamt] & "")
            
            ITM.SubItems(2) = IIf(chkExternal = vbChecked, EmpExtSal, EmpSalary & "")
            Rate = Format(Val(EmpSalary) / LastDay, "0.00")
            ITM.SubItems(3) = Rate
            
            
                
                
            
            
            If IsNull(![presentDays]) Then
                SalDays = 0
                'IF NO Presents Found Exclude This Employee
                LV.ListItems.Remove ITM.key
                GoTo NEXT_EMPLOYEE
            Else
                SalDays = ((Val(![presentDays]) / 2) + Holidays) + AbsAllowed
                'if Salary Days Exceed Total Salary Days Set Them To Salary Days
                If SalDays > TotalSalDays Then SalDays = TotalSalDays
            End If
            
            
            
            
            
            Dim DaysToExclude As Integer
            
            If Format(![JoinDate], "MMM-yyyy") = Format(JDate, "MMM-yyyy") Then
               DaysToExclude = Day(![JoinDate]) - 1
            Else
               DaysToExclude = 0
            End If
            
            
            'Deduct Penalty Dayss
            SalDays = SalDays - Val(![AbsentDays] & "")
            If SalDays < 0 Then SalDays = 0
            '''''''''''''''''''''''''''''''
            
            AbsentDays = TotalSalDays - SalDays - DaysToExclude
            
                        
            ' Make Absents 0 if in Negative
            
            If AbsentDays < 0 Then AbsentDays = 0
            
            
            'If IsNull(![AbsentDays]) Then AbsentDays = 0 Else AbsentDays = Val(![AbsentDays]) / 2
            'AbsentDays = IIf(AbsentDays > 0, AbsentDays - AbsAllowed, AbsentDays)
            ITM.SubItems(4) = AbsentDays
            ITM.SubItems(5) = Round(AbsentDays * Rate)
            
            
            'If SalDays = 0 Then SalDays = 0 Else SalDays = (SalDays - AbsentDays)
            
            ITM.SubItems(6) = SalDays
            ITM.SubItems(7) = Round(SalDays * Rate)
            
            
            
            ITM.SubItems(8) = getHrsMin(OverTimeHours)
            ITM.SubItems(9) = OverTimeHours * (Rate / Val(![OverTime] & ""))
            
            LateHours = GetLateHours(![EmpId], NTime, LTime, JDate)
            
            ITM.SubItems(10) = LateHours
            ITM.SubItems(11) = Round((LateHours * (Rate / 8)))
            
            Total = Round(Val(ITM.SubItems(7)) + Val(ITM.SubItems(9)) - Val(ITM.SubItems(11)))
            ITM.SubItems(12) = Total
            'TaxPer = Val(![PerDeduct])
            ITM.SubItems(13) = TaxPer
            'DedAmt = Round((Val(![Salary]) / TotalSalDays) * (TaxPer / 100))
            TaxSal = Round((Val(EmpSalary) / LastDay) * TotalSalDays)
            DedAmt = Round(TaxSal * (TaxPer / 100))
            ITM.SubItems(14) = DedAmt
            ITM.SubItems(15) = Total - DedAmt
            If IsNull(![Payed]) Then Payed = 0 Else Payed = ![Payed]
            ITM.SubItems(16) = Payed
            
            ITM.SubItems(17) = LongLoan
            
            ITM.SubItems(18) = Val(![AdvSalpaid] & "")
            
            ITM.SubItems(19) = SocialAmt
            
            ITM.SubItems(20) = Round((Val(ITM.SubItems(15)) - (Payed + SocialAmt + LongLoan + Val(ITM.SubItems(18)))) / 10, 0) * 10
            
            GrandTotal = GrandTotal + Val(ITM.SubItems(20))
            
            Dim LastDate As Date
            LastDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), LastDay)

            
            If (IsComputed) Then
                    cmdUpdate.Visible = UserHasAccess("FinalSalSheet")
            ElseIf Not (IsComputed) And (ServerDate > LastDate) Then
                    cmdUpdate.Visible = UserHasAccess("FinalSalSheet")
            Else
                    cmdUpdate.Visible = False
            End If
NEXT_EMPLOYEE:
            .MoveNext
        Next
        
        LV.Visible = True
        lblTotal.Tag = GrandTotal
        lblTotal = Format(GrandTotal, "#,##0.00")
        
    End With
    cmbMonth.Tag = cmbMonth.Text
    cmbYear.Tag = cmbYear.Text
    
    Sql = ""
    Cond = ""
    ICO = ""
End Sub
Private Function GetLateHours(EmpId As String, NTime As Integer, LTime As Double, ToDate As Date) As Double

    Dim LHours As Double
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    With cmd
        '.CommandText = "Select LHrs from VLHours Where EmpID='" & EmpID & "'"
        .CommandText = "VLHours"
        .CommandType = adCmdStoredProc
         Set .ActiveConnection = con
        .Parameters.Refresh
        .Parameters(1).Type = adNumeric
        .Parameters(1).value = LTime
        
        .Parameters(2).Type = adDBTimeStamp
        .Parameters(2).value = ToDate
        
        .Parameters(3).Type = adNumeric
        .Parameters(3).value = NTime
        
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        Set rs = .Execute
        
    End With
    
    If rs.RecordCount > 0 Then
        If IsNull(rs.Fields(1)) Then LHours = 0 Else LHours = rs.Fields(1)
    Else
        LHours = 0
    End If
    GetLateHours = LHours
    
End Function
Private Sub Computed(Emp As String, ByRef IsComputed As Boolean, ByRef TaxPer As Double, ByRef AbsentDays As Integer, ByRef OverTimeHours As Double, ByRef NTime As Integer, ByRef LTime As Double, ByRef LongLoan As Double, ByRef DT As String, ByRef Payed As Boolean, ByRef Transfered As Boolean, ToDate As Date, ByRef EmpSal As Double, Optional EmpExtSal As Double, Optional IsActive As Boolean, Optional SocialAmt As Double)

    Dim rsCheck As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    With cmd
        '.CommandText = "Select * from VSalarySettings where empid='" & Emp & "'"
        .CommandType = adCmdStoredProc
        .CommandText = "VSalarySettings"
        Set .ActiveConnection = con
        .Parameters.Refresh
        .Parameters(1).Type = adDBTimeStamp
        .Parameters(1).value = ToDate
        
        
        .Parameters(2).Type = adVarChar
        .Parameters(2).value = Emp
        
        .Parameters(3).Type = adBoolean
        .Parameters(3).value = 0
        
        rsCheck.CursorType = adOpenForwardOnly
        rsCheck.LockType = adLockReadOnly
        Set rsCheck = .Execute()
    End With
    
    With rsCheck
        
        If .RecordCount > 0 Then
            IsComputed = True
            TaxPer = ![PerDeduct]
            AbsentDays = ![absentallowed]
            OverTimeHours = ![OverTime]
            NTime = ![NTimes]
            LTime = ![TimeAllowed]
            DT = Format(![DT], "dd-MMM-yyyy") & ""
            ', AbsentDays, OverTimeHours, NTime, LTime)
            LV.Checkboxes = True
            cmdUpdate.Visible = True
            
            Payed = ![Payed] & ""
            Transfered = ![Transfered] & ""
            SocialAmt = Val(!SocialAmount & "")
            If IsNull(![LongLoan]) Then LongLoan = 0 Else LongLoan = ![LongLoan]
            
            'Get The Salary Depending On Salary Sheet Type
            'EmpSal = IIf(chkExternal = vbUnchecked, Val(![Salary] & ""), Val(![ExtSalary] & ""))
            EmpSal = Val(![Salary] & "")
            EmpExtSal = Val(![ExtSalary] & "")
            cmdUpdate.Caption = "Save Changes"
            
        Else
            If IsActive Then
                IsComputed = False
                LV.Checkboxes = False
                'cmdUpdate.Visible = False
                cmdUpdate.Caption = "Make This Sheet Final"
            End If
        End If
        
        
    End With
    
End Sub

Private Sub SaveIt(rs As ADODB.Recordset, PerDeduct As Double, Optional i As Integer)
    Dim DT As Date
    Dim LongLoan As Double
    Dim lDedAmt As Long, lClsAmt1 As Long, lTotAmt As Long
    If IsNull(rs![DedAmt]) Then lDedAmt = 0 Else lDedAmt = rs![DedAmt]
    If IsNull(rs![totamt]) Then lTotAmt = 0 Else lTotAmt = rs![totamt]
    If IsNull(rs![clsamt]) Then lClsAmt1 = 0 Else lClsAmt1 = rs![clsamt]
                
    If lDedAmt > (lTotAmt - lClsAmt1) Then
        LongLoan = (lTotAmt - lClsAmt1)
    Else
        LongLoan = lDedAmt
    End If
    
    'If IsNull(rs![LongLoan]) Then LongLoan = 0 Else LongLoan = rs![LongLoan]
    
    
    
    
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    
    'Get The Last Day of Month
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), DateAdd("m", 1, DT) - DT)
    
    con.Execute "Insert into MonthlySalaries (EmpID,DT,Salary," & _
     "ExtSalary,PerDeduct,AbsentAllowed,TimeAllowed,NTimes,OverTime,Payed,ShortTermPer,LongLoan,Transfered," & _
     "AdvSalary,NetSalary,OTAmt,SalAmt,LtAmt,StAmt,LateAmt,SocialAmount)" & " Values('" & rs![EmpId] & "','" & DT & _
     "'," & rs![StartingSalary] & "," & rs![Salary] & "," & PerDeduct & "," & Abs(rs![AbsAllowed]) & _
     "," & rs![TimeAllowed] & "," & Val(rs![NTimes] & "") & "," & Val(rs![OverTimeHours] & "") & ",0," & _
     Val(rs![ShortTermPer] & "") & "," & LongLoan & ",0,0," & Val(LV.ListItems(i).SubItems(20)) + Val(LV.ListItems(i).SubItems(18)) & _
     "," & Val(LV.ListItems(i).SubItems(9)) & "," & Val(LV.ListItems(i).SubItems(7)) & "," & Val(LV.ListItems(i).SubItems(17)) & _
     "," & Val(LV.ListItems(i).SubItems(16)) & "," & Val(LV.ListItems(i).SubItems(11)) & "," & Val(LV.ListItems(i).SubItems(19)) & ")"
    
    
    'Update Long Term ...
    
    Dim LongAmt As Double
    LongAmt = Val(LV.ListItems(i).SubItems(17))
    
    con.Execute "Insert Into AmtCleared (EmpID,DT,AmtClrd) Values('" & rs![EmpId] & "','" & DT & "'," & LongAmt & ")"
    
    
    If LV.Checkboxes = False Then
     LV.Checkboxes = True
        cmdUpdate.Visible = True
    End If
    
End Sub
Private Function TotalMonthDays() As Integer
    Dim TDate As Date
    Dim Tot As Integer
    Dim year As Integer
    If cmbYear = "" Then
        year = 2002
    Else
        year = cmbYear
    End If
    
    If Month(ServerDate) = cmbMonth.List(cmbMonth.ListIndex, 1) Then
        Tot = Day(ServerDate)
    Else
        TDate = DateSerial(year, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
        Tot = DateAdd("m", 1, TDate) - TDate
    End If
    TotalMonthDays = Tot
End Function

Private Sub LV_DblClick()
    If LV.ListItems.Count = 0 Then Exit Sub
    If Val(LV.SelectedItem.SubItems(17)) = 0 Then Exit Sub
    If LV.SelectedItem.Bold Then Exit Sub
    
    
    Dim Amt As String
    Amt = InputBox("Please Enter Amount To Deduct For Long Term." & vbNewLine & "The Current Deduction Amount Is " & LV.SelectedItem.SubItems(17) & ".", , LV.SelectedItem.SubItems(17)) & ""
    
    If Not IsNumeric(Amt) Then
        MsgBox "Invalid Amount.", vbInformation
    Else
        
        ' Update The Balance
        
        LV.SelectedItem.SubItems(20) = Val(LV.SelectedItem.SubItems(20)) + Val(LV.SelectedItem.SubItems(17)) - Abs(Val(Amt))
        LV.SelectedItem.SubItems(17) = Abs(Val(Amt))
        
        ' Update The Saved Entries
        
        con.Execute "Update AmtCleared Set AmtClrd=" & Abs(Val(Amt)) & "  Where convert(char(3),dt,7)='" & cmbMonth & "' and EmpID='" & LV.SelectedItem & "'"
        con.Execute "Update MonthlySalaries Set LongLoan=" & Abs(Val(Amt)) & " Where convert(char(3),dt,7)='" & cmbMonth & "' and EmpID='" & LV.SelectedItem & "'"
        
    End If
    
    ' With LV.SelectedItem
    '    txtEdit.Move LV.Left + LV.ColumnHeaders(18).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(18).Width
    '    txtEdit = .SubItems(17)
    ' End With
    '
    ' With txtEdit
    '     .Visible = True
    '     .SelStart = 0
    '     .SelLength = Len(.Text)
    '     .SetFocus
    ' End With
    ' MsgBox LV.FindItem (

End Sub

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If Item.Bold Then
      Item.Checked = Not Item.Checked
   MsgBox "Cannot Modify This Entry It Has Been Transfered To Accounts", vbInformation

End If
End Sub

Private Sub LV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If LV.ListItems.Count = 0 Then
    EditAbsents = False
    Exit Sub
End If
    If x >= LV.Left + LV.ColumnHeaders(5).Left And x < LV.Left + LV.ColumnHeaders(6).Left And _
        y >= LV.SelectedItem.Top And y < LV.SelectedItem.Top + LV.SelectedItem.Height Then
        EditAbsents = True
    Else
        EditAbsents = False
    End If
    
End Sub

Private Sub mnuPaySlips_Click()
   Call PrintReport(True)
End Sub

Private Sub mnuSalSheet_Click()
   Call PrintReport(False)
End Sub

Private Sub PrintSalsummary_Click()
 Call PrintReport(False, , True)
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LV.SelectedItem.SubItems(17) = txtEdit.Text
        txtEdit.Text = ""
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Text = ""
        txtEdit.Visible = False
    Else
        KeyAscii = OnlyNumber(KeyAscii)
    End If
End Sub






Private Function AddSalaryEntries() As Boolean

   Dim SSAdv As Double, SLadv As Double, SSal As Double, STax As Double     'Staff Member Vars
   Dim DSAdv As Double, DLadv As Double, DSal As Double, DTax As Double     'Directors Vars
   Dim PSAdv As Double, PLadv As Double, PSal As Double, PTax As Double     'Production Staff Vars
   
   
   
   Dim Dated As Date
   Dated = ServerDate
   
   '''''''''''''''''' A D V A N C E   C L E A R E N C E ''''''''''''''''''
   ConAcc.BeginTrans
   
   Call GetTotal(SSAdv, DSAdv, PSAdv, "SADV")        'Short Term Advances
   Call AddAdvClearVoucher(SSAdv, DSAdv, PSAdv, Dated, False, cmbMonth & ", " & cmbYear)
   
   Call GetTotal(SLadv, DLadv, PLadv, "LADV")        'Long  Term Advances
   Call AddAdvClearVoucher(SLadv, DLadv, PLadv, Dated, True, cmbMonth & ", " & cmbYear)
   
   
   '''''''''''''''''' S A L R A R Y   &   T A X    E N T R Y ''''''''''''''''''
   
   Call GetTotal(SSal, DSal, PSal, "SAL")            'Salaries
   Call GetTotal(STax, DTax, PTax, "TAX")            'Taxes
   
   Call addSalaryVoucher(SSal, DSal, PSal, Dated, STax, DTax, PTax, cmbMonth & ", " & cmbYear)
   
   ConAcc.CommitTrans
   
   
   
    For i = 1 To LV.ListItems.Count
    
        con.Execute "Update MonthlySalaries Set Transfered=" & CBool(LV.ListItems(i).Checked Or LV.ListItems(i).Bold) & " Where EmpID+Format(DT,'dd-MMM-yyyy')='" & LV.ListItems(i).key & "'"
    
    Next i
      
   
   AddSalaryEntries = True
   Exit Function
err:
   AddSalaryEntries = False
   MsgBox "Cannot Add Salary Vouchers Due To Error in " & err.Source & " :" & vbNewLine & err.Description, vbCritical
   ConAcc.RollbackTrans
End Function

Private Sub GetTotal(Stotal As Double, DTotal As Double, PTotal As Double, TotalType As String)

   Dim StaffTotal As Double, ProductionTotal As Double, DirectorTotal As Double, Amount As Double



   For i = 1 To LV.ListItems.Count
      'if Paid or Transfered then
      If Not LV.ListItems(i).Checked Or LV.ListItems(i).Bold Then GoTo NEXT_ITEM
         
   
         Select Case TotalType
            Case "LADV"
               'Long Term Advances
               Amount = Val(LV.ListItems(i).SubItems(17))
            Case "SADV"
               'Short Tem Advances
               Amount = Val(LV.ListItems(i).SubItems(16))
            Case "SAL"
               'Salary amount
               Amount = Val(LV.ListItems(i).SubItems(12))
            Case "TAX"
               Amount = Val(LV.ListItems(i).SubItems(14))
         End Select
   
   
   
   
    Select Case LV.ListItems(i).Tag
      Case "0"
         StaffTotal = StaffTotal + Amount
      Case "1"
         DirectorTotal = DirectorTotal + Amount
      Case "2"
         ProductionTotal = ProductionTotal + Amount
    End Select


NEXT_ITEM:
   Next
   
   PTotal = ProductionTotal
   Stotal = StaffTotal
   DTotal = DirectorTotal
   
End Sub

Private Function GetTotalTax() As Double
Dim Total As Double
   For i = 1 To LV.ListItems.Count
    Total = Total + Val(LV.ListItems(i).SubItems(14))
   Next
GetTotalTax = Total
End Function


Public Sub ShowExternalSalSheet(forMonth As String, forYear As String, Dept As String, Crpt As CrystalReport)
   Me.Visible = False
   cmbMonth.Text = forMonth
   cmbYear.Text = forYear
   cmbDept.Text = Dept
   cmbEmp.ID = 0
   chkExternal = vbChecked
   Call PrintReport(False, Crpt)
End Sub


Private Sub mnuCurrency_Click()

    Thousands = 0
    FHundreds = 0
    Hundreds = 0
    Fifties = 0
    Tens = 0
    Fives = 0
    Twos = 0
    Ones = 0
    TotalSal = 0


    For i = 1 To LV.ListItems.Count
        If Val(LV.ListItems(i).ListSubItems(20)) > 0 Then
            CountRupee Val(LV.ListItems(i).ListSubItems(20))
            TotalSal = TotalSal + Val(LV.ListItems(i).ListSubItems(20))
        End If
    Next
    
    

    With Cr
        .ReportFileName = RptPRPath & "\RupeeStructure.rpt"
        
            .Formulas(0) = "Thousands=" & Thousands
            .Formulas(1) = "FHundreds=" & FHundreds
            .Formulas(2) = "Hundreds=" & Hundreds
            .Formulas(3) = "Fifties=" & Fifties
            .Formulas(4) = "Tens=" & Tens
            .Formulas(5) = "Fives=" & Fives
            .Formulas(6) = "Tows=" & Twos
            .Formulas(7) = "Ones=" & Ones
            .Formulas(8) = "TotalSal=" & lblTotal.Tag
            .Formulas(9) = "Companyname='" & CompanyName & "'"
        
    
        .WindowState = crptMaximized
        .WindowTitle = "Print Rupee Structure"
        .Action = 1
        
    End With

End Sub

Private Sub CountRupee(Salary As Double)

    Thousands = Thousands + Int(Salary / 1000)
    If Salary >= 500 Then FHundreds = FHundreds + Int(Val(Right(Salary, 3)) / 500)
    Hundreds = Hundreds + Int(Val(Right(Salary, 3) Mod 500) / 100)
    Fifties = Fifties + Int(Val(Right(Salary, 2)) / 50)
    Tens = Tens + Int(Val(Right(Salary, 3) Mod 50) / 10)
    Fives = Fives + Int(Val(Right(Salary, 1)) / 5)
    Twos = Twos + Int(Val(Right(Salary, 1) Mod 5) / 2)
    Ones = Ones + Int(Val(Right(Salary, 1) Mod 5) Mod 2)


End Sub


