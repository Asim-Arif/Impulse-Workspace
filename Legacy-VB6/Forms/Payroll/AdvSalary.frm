VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form AdvSalary 
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleMode       =   0  'User
   ScaleWidth      =   10667.95
   WindowState     =   2  'Maximized
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
      Left            =   9750
      TabIndex        =   12
      Top             =   885
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   0
      TabIndex        =   8
      Top             =   15
      Width           =   11670
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Advance Salary Sheet"
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
         Width           =   3345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Advance Salary Sheet"
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
         Width           =   3345
      End
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   45
      TabIndex        =   4
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
      Height          =   285
      Left            =   2265
      TabIndex        =   5
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
      Height          =   6255
      Left            =   45
      TabIndex        =   0
      Top             =   1230
      Width           =   11835
      Begin VB.Frame fButs2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   900
         Left            =   45
         TabIndex        =   21
         Top             =   5340
         Width           =   4935
         Begin VB.CommandButton cmdAccounts 
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
            Left            =   2310
            TabIndex        =   23
            Top             =   450
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.CommandButton Command1 
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
            Left            =   3615
            TabIndex        =   22
            Top             =   450
            Visible         =   0   'False
            Width           =   1275
         End
         Begin MSForms.CommandButton cmdUpdate 
            Height          =   360
            Left            =   15
            TabIndex        =   24
            Top             =   450
            Visible         =   0   'False
            Width           =   2265
            ForeColor       =   0
            Caption         =   "Make This Sheet Final"
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
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   900
         Left            =   8250
         TabIndex        =   14
         Top             =   5310
         Width           =   3495
         Begin MSForms.CommandButton cmdCancel 
            Height          =   360
            Left            =   1995
            TabIndex        =   15
            Top             =   495
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
            Left            =   2010
            TabIndex        =   20
            Top             =   495
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
            Left            =   465
            TabIndex        =   19
            Top             =   495
            Width           =   1425
            ForeColor       =   0
            Caption         =   "Print         "
            PicturePosition =   327683
            Size            =   "2514;635"
            Picture         =   "AdvSalary.frx":0000
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
            Left            =   1980
            TabIndex        =   18
            Top             =   75
            Width           =   1425
         End
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
            Left            =   15
            TabIndex        =   17
            Top             =   90
            Width           =   1875
         End
         Begin MSForms.CommandButton cmdSave 
            Height          =   360
            Left            =   465
            TabIndex        =   16
            Top             =   495
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
         Height          =   5130
         Left            =   60
         TabIndex        =   1
         Top             =   195
         Width           =   11685
         _ExtentX        =   20611
         _ExtentY        =   9049
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
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emp Name"
            Object.Width           =   2117
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
            Text            =   "Balance"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   19
            Text            =   "%age"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   20
            Text            =   "Payable"
            Object.Width           =   1764
         EndProperty
      End
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
            Picture         =   "AdvSalary.frx":0112
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AdvSalary.frx":0564
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   8055
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
   Begin MSForms.ComboBox cmbDays 
      Height          =   300
      Left            =   7050
      TabIndex        =   13
      Top             =   900
      Width           =   990
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "1746;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbMonth 
      Height          =   300
      Left            =   5685
      TabIndex        =   7
      Top             =   885
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
      Top             =   885
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
   Begin VB.Menu mnupop 
      Caption         =   "mnupop"
      Visible         =   0   'False
      Begin VB.Menu mnuSalSheet 
         Caption         =   "Print Salary Sheet"
      End
      Begin VB.Menu mnuPaySlips 
         Caption         =   "Print PaySlips"
      End
   End
End
Attribute VB_Name = "AdvSalary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsLedger As New ADODB.Recordset
Private Sub chkExternal_Click()
    Call cmdRefresh_Click
End Sub

Private Sub cmbMonth_Click()
    Dim maxday As Integer
cmbDays = TotalMonthDays
If Format(ServerDate, "MMM-yyyy") = cmbMonth & "-" & cmbYear Then
    maxday = Day(ServerDate)
Else
    maxday = TotalMonthDays
End If
cmbDays.Clear
For i = 1 To maxday
cmbDays.AddItem Format(i, "00")

Next
cmbDays.ListIndex = maxday - 1
End Sub

Private Sub cmbYear_Change()
    txtTotal = TotalMonthDays
End Sub

Private Sub cmdAccounts_Click()
For i = 1 To LV.ListItems.Count
    LV.ListItems(i).Checked = True
Next

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

Private Sub PrintReport(PaySlips As Boolean)

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
     If LV.Checkboxes And LV.ListItems(i).Checked Then
        .AddNew
        ![EmpId] = LV.ListItems(i).Text & ""
        For j = 1 To 20
            
            If j = 19 Then
                .Fields(j) = LV.ListItems(i).SubItems(j + 1) & ""
            ElseIf j = 20 Then
                .Fields(j) = LV.ListItems(i).SubItems(j - 1) & ""
            Else
                .Fields(j) = LV.ListItems(i).SubItems(j) & ""
            End If
        Next j
        ![DT] = DT & ""
        .Update
     End If
    Next i
End With

'Sleep (1000)
'Dim TDate As Date
'Dim MonthDays As Integer
'TDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
'MonthDays = DateAdd("m", 1, TDate) - TDate

With cr1
    
    If Not PaySlips Then
      .ReportFileName = RptPRPath & "\AdvSalSheet.rpt"
      '.DataFiles(0) = DatabasePath
      .Connect = ConnectStr
      .Formulas(0) = "Fromto=' For " & cmbMonth.Tag & "-" & cmbYear.Tag & "'"
    Else
      .ReportFileName = RptPRPath & "\AdvPaySlip.rpt"
      '.DataFiles(0) = DatabasePath
      .Connect = ConnectStr
      .Formulas(0) = "formonth=' For " & cmbMonth.Tag & "-" & cmbYear.Tag & "'"
    End If
    
    '.Formulas(1) = "ForMonthDays=" & MonthDays
    .Action = 1
    .PageZoomNext
End With

Call HideWait
Exit Sub
err:
    MsgBox err.Description
    Call HideWait
End Sub

Private Sub cmdPrint_Click()
Me.PopUpMenu mnupop, 2, FLV.Left + fButs1.Left + cmdPrint.Left, FLV.Top + fButs1.Top + cmdPrint.Top + cmdPrint.Height
'Call PrintReport(False)
End Sub

Public Sub cmdRefresh_Click()
    Call ShowWait
    Call RefreshLedger
     
    'cmdAccounts.Visible = cmdUpdate.Visible And AccountsLinked
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
   
   For i = 1 To LV.ListItems.Count
   .Find "EMPID='" & LV.ListItems(i) & "'", 0, adSearchForward, 1
      Call SaveIt(rsLedger, 0, LV.ListItems(i).Checked, Val(LV.ListItems(i).SubItems(20)))
   
      .MoveNext
   Next
   
   con.CommitTrans
Call cmdRefresh_Click
End With
   LV.Checkboxes = True
   cmdAccounts.Visible = True
   Command1.Visible = True
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
        
        con.Execute "Update MonthlySalaries Set Payed=" & setVal & " Where EmpID='" & EmpId & "' and month(dt)=" & TheMonth & " and Year(dt)=" & TheYear & " and AdvSalary=1", a
    Next i
    
    con.CommitTrans
    
End If

End Sub

Private Sub Command1_Click()
For i = 1 To LV.ListItems.Count
    LV.ListItems(i).Checked = False
Next
End Sub

Private Sub Form_Load()

cmdUpdate.Visible = UserHasAccess("FinalAdvSal")
 
 'Hide Columns Which are Not required
 
    LV.ColumnHeaders(14).Width = 0
    LV.ColumnHeaders(15).Width = 0
    LV.ColumnHeaders(16).Width = 0
    
    'LV.ColumnHeaders(11).Width = 0
    'LV.ColumnHeaders(12).Width = 0
    
    LV.ColumnHeaders(5).Width = 0
    LV.ColumnHeaders(6).Width = 0
    
    

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
    
    
    
    
    
    
    Dim EmpId As String
    Dim TotalSalDays As Double
    Dim TempTotalSaldays As Double
    Dim LastDay As Integer
    
    Dim ITM As ListItem
    Dim Cond As String
    Dim Sql As String
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    'Dim rsLedger As New ADODB.Recordset
    
    Dim SundayAbsents As Double
    
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
    
    'TotalSalDays = DateAdd("m", 1, SDate) - SDate
    
    rs.Open "Select distinct(DT) from MonthlySalaries Where convert(char(3),dt,7)='" & cmbMonth & "' and Year(DT)=" & cmbYear & " and AdvSalary=1", con
    Dim AdvSalLocked As Boolean
    If rs.RecordCount > 0 Then
      cmbDays.Enabled = False
      cmbDays.Text = Format(Day(rs.Fields(0)), "00")
      LV.Checkboxes = True
      TDate = rs.Fields(0)
      TotalSalDays = Day(rs.Fields(0))
    Else
      cmbDays.Enabled = True
      LV.Checkboxes = False
      'TotalSalDays = DateAdd("m", 1, SDate) - SDate
      TotalSalDays = cmbDays
    End If
      
    rs.Close
    Set rs = Nothing
    
    
    LastDay = DateAdd("m", 1, SDate) - SDate
    
    Dim JDate As Date
    If TDate = ServerDate Then
        TotalSalDays = cmbDays
    ElseIf Month(ServerDate) < Month(TDate) Then
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
    
    'cmd.Parameters(4).Type = adDBTimeStamp
    'cmd.Parameters(4).value = JDate
    If rsLedger.State = 1 Then rsLedger.Close
    rsLedger.CursorType = adOpenKeyset
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
        
        Dim EmpSalary As Double
        
        
       
        
        LV.ListItems.Clear
        LV.Visible = False
        For i = 0 To .RecordCount - 1
                    
            Holidays = 0
            AbsentDays = 0
                    
                    
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
                    
                    
            
            
            
            Dim DaysToExclude As Integer
            If Format(![JoinDate], "MMM-yyyy") = Format(JDate, "MMM-yyyy") Then
               DaysToExclude = Day(![JoinDate]) - 1
            Else
               DaysToExclude = 0
            End If
            
            
            
                    
                    
                    
                    
                    
            Call Computed(![EmpId], IsComputed, TaxPer, AbsAllowed, OverTimeHours, NTime, LTime, LongLoan, DT, IsPayed, IsTransfered, JDate, EmpSalary, ![Active])
            
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
               EmpSalary = IIf(chkExternal = vbUnchecked, Val(![StartingSalary] & ""), Val(![Salary] & ""))
            
            End If
            
            
            
            'Dim theKey As String
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
            
            
            
            ITM.SubItems(1) = ![Name] & ""
            ITM.SubItems(2) = EmpSalary & ""
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
            
            
            
                     AbsentDays = TotalSalDays - SalDays - DaysToExclude
                     ' Make Absents 0 if in negative
            
                     If AbsentDays < 0 Then AbsentDays = 0
            
            
            
            
            'If IsNull(![AbsentDays]) Then AbsentDays = 0 Else AbsentDays = Val(![AbsentDays]) / 2
            'AbsentDays = IIf(AbsentDays > 0, AbsentDays - AbsAllowed, AbsentDays)
            ITM.SubItems(4) = AbsentDays
            ITM.SubItems(5) = Round(AbsentDays * Rate)
            
            
            'If SalDays = 0 Then SalDays = 0 Else SalDays = (SalDays - AbsentDays)
            
            ITM.SubItems(6) = SalDays
            ITM.SubItems(7) = Round(SalDays * Rate)
            
            
            
            ITM.SubItems(8) = OverTimeHours
            ITM.SubItems(9) = Round(OverTimeHours * (Rate / Val(![OverTime] & "")))
            
            LateHours = GetLateHours(![EmpId], NTime, LTime, JDate)
            
            ITM.SubItems(10) = "0" 'LateHours
            ITM.SubItems(11) = "0" ' Round((LateHours * (Rate / 8)))
            
            'Dont Deduct the Late Hours
            Total = Round(Val(ITM.SubItems(7)) + Val(ITM.SubItems(9)))  '- Val(itm.SubItems(11)))
            
            ITM.SubItems(12) = Total
            'TaxPer = Val(![PerDeduct])
            ITM.SubItems(13) = TaxPer
            'DedAmt = Round((Val(![Salary]) / TotalSalDays) * (TaxPer / 100))
            TaxSal = Round((Val(EmpSalary) / LastDay) * TotalSalDays)
            DedAmt = 0 'Round(TaxSal * (TaxPer / 100))
            ITM.SubItems(14) = DedAmt
            ITM.SubItems(15) = Total - DedAmt
            If IsNull(![Payed]) Then Payed = 0 Else Payed = ![Payed]
            ITM.SubItems(16) = Payed
             
            ITM.SubItems(17) = LongLoan
            ITM.SubItems(18) = Val(ITM.SubItems(15)) - Payed - LongLoan
            
            ITM.SubItems(19) = Val(![AdvSalPer] & "")
            Dim NetPayable As Long
            
            NetPayable = Round(Val(ITM.SubItems(18)) * Val(![AdvSalPer] & "") / 10000, 0) * 100
            
            ITM.SubItems(20) = Format(NetPayable, "0.00")
            'Calculate Grand Total Payable
            GrandTotal = GrandTotal + NetPayable
            Dim LastDate As Date
            LastDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), LastDay)
            'Save Tax,Settings for Selected Month...
            'If Not (IsComputed) And (ServerDate > LastDate) Then Call SaveIt(rsLedger, TaxPer, Val(itm.SubItems(19)))
            'If Not (IsComputed) And (ServerDate > LastDate) Then
            '    cmdRefresh.Visible = False
            '    cmdClose.Visible = False
            '    cmdPrint.Visible = False
            '    cmdSave.Visible = True
            '    cmdCancel.Visible = True
            'End If
NEXT_EMPLOYEE:
            .MoveNext
        Next
        '.Close
        LV.Visible = True
    End With
    cmbMonth.Tag = cmbMonth.Text
    cmbYear.Tag = cmbYear.Text
    'Set rsLedger = Nothing
    
    Command1.Visible = LV.Checkboxes
    cmdAccounts.Visible = LV.Checkboxes
    'Show Grand Total
    lblTotal = Format(GrandTotal, "#,##0.00")
    Sql = ""
    Cond = ""
    ICO = ""
End Sub
Private Function GetLateHours(EmpId As String, NTime As Integer, LTime As Double, ByVal ToDate As Date) As Double

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
Private Sub Computed(Emp As String, ByRef IsComputed As Boolean, ByRef TaxPer As Double, ByRef AbsentDays As Integer, ByRef OverTimeHours As Double, ByRef NTime As Integer, ByRef LTime As Double, ByRef LongLoan As Double, ByRef DT As String, ByRef Payed As Boolean, ByRef Transfered As Boolean, ToDate As Date, ByRef EmpSal As Double, Optional IsActive As Boolean)

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
        .Parameters(3).value = 1
        
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
            cmdUpdate.Visible = UserHasAccess("FinalAdvSal")
            
            Payed = ![Payed] & ""
            Transfered = ![Transfered] & ""
            
            If IsNull(![LongLoan]) Then LongLoan = 0 Else LongLoan = ![LongLoan]
                        
            'Get The Salary Depending On Salary Sheet Type
            EmpSal = IIf(chkExternal = vbUnchecked, Val(![Salary] & ""), Val(![ExtSalary] & ""))
            
            'EmpSal = Val(![StartingSalary] & "")
            
            cmdUpdate.Caption = "Save Changes"
            
        Else
            
            IsComputed = False
            If IsActive Then
                cmdUpdate.Caption = "Make This Sheet Final"
            End If
            
        End If
        
        
    End With
    
End Sub

Private Sub SaveIt(rs As ADODB.Recordset, PerDeduct As Double, Optional Paid As Boolean, Optional PaidAmt As Double)
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
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), cmbDays)
    con.Execute "Insert into MonthlySalaries (EmpID,DT,Salary,ExtSalary" & _
     ",PerDeduct,AbsentAllowed,TimeAllowed,NTimes,OverTime,Payed,ShortTermPer,LongLoan,Transfered,AdvSalary,PaidAmt,PerPay)" & _
     " Values('" & rs![EmpId] & "','" & DT & _
     "'," & rs![StartingSalary] & "," & rs![Salary] & "," & PerDeduct & "," & Abs(rs![AbsAllowed]) & _
     "," & rs![TimeAllowed] & "," & rs![NTimes] & "," & Val(rs![OverTimeHours] & "") & ",0," & Val(rs![ShortTermPer] & "") & "," & LongLoan & ",0,1," & IIf(Paid, PaidAmt, 0) & "," & Val(rs![AdvSalPer] & "") & ")"
     
    'Update Long Term ...
    'Dim LongAmt As Double
    'LongAmt = Val(LV.ListItems(LV.ListItems.Count).SubItems(17))
    
    'Con.Execute "Insert Into AmtCleared (EmpID,DT,AmtClrd) Values('" & rs![EmpID] & "','" & DT & "'," & LongAmt & ")"
    
    
    If LV.Checkboxes = False Then
        LV.Checkboxes = True
        cmdUpdate.Visible = UserHasAccess("FinalAdvSal")
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

Private Sub mnuPaySlips_Click()
   Call PrintReport(True)
End Sub

Private Sub mnuSalSheet_Click()
   Call PrintReport(False)
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

Public Function addSalaryVoucher(StaffAmt As Double, DirectorAmt As Double, ProductionAmt As Double, Dated As Date, STax As Double, DTax As Double, PTax As Double, Optional SalFor As String) As Boolean
    
    Dim i As Long, INextSno As Double, Bal As Double, Amt As Double
    Dim VchrNo As String, TheDesc As String, AccNo As String
    
      
    'Charge All Amount to Cash
    Amt = StaffAmt + ProductionAmt + DirectorAmt
    
    If Amt = 0 Then Exit Function
    
     'ConAcc.BeginTrans
     
     'Get The Required Values
     
     INextSno = getNextSno(Dated)
     VchrNo = getNextVchrNo(Dated, "PV")
     
     

     
     
     
     
    ''''''''''''''''Charge To  Cash Entry'''''''''''''
    
    AccNo = CASHINHAND
    
    Bal = getBalance(AccNo, Dated)
    
    Bal = Bal - Amt
    
    TheDesc = "Payment For " & SalFor
    
    'Credit Entry
    ConAcc.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & TheDesc & "',0," & Amt & "," & Bal & ",'',0)"
        
    
    
    
    Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
     
     
     INextSno = INextSno + 1
     
     
     
     
    '''''''''''''''''' D E B I T   E N T R I E S '''''''''''''
     
     
     
     
     
If StaffAmt = 0 Then GoTo DIRECTOR_ENTRY
      
      'Determine the Account no for Staff Member
      'Long / Short Tem Advance
      
      AccNo = SalaryAccNo
     
      Bal = getBalance(AccNo, Dated)
      
      Amt = StaffAmt - STax
      Bal = Bal + Amt
      
      
    
      'Debit To Staff Long/Short Advance Account
      
      ConAcc.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & TheDesc & "'," & Amt & ",0," & Bal & ",'',0)"
        
     
        Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
     
         INextSno = INextSno + 1
     
DIRECTOR_ENTRY:

If DirectorAmt = 0 Then GoTo PRODUCTION_ENTRY
      'Determine the Account no for Directors
      
      
      
      AccNo = DSalaryAccNo
     
      Bal = getBalance(AccNo, Dated)
      
      
      Amt = DirectorAmt - DTax
      Bal = Bal + Amt
    
    
      'Debit To Staff Long/Short Advance Account
      
      ConAcc.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & TheDesc & "'," & Amt & ",0," & Bal & ",'',0)"
        
     
        Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
        
         INextSno = INextSno + 1
         
         
PRODUCTION_ENTRY:
         'Determine the Account no for Directors
      
If ProductionAmt = 0 Then GoTo TAX_ENTRY

      
      
      AccNo = PSalaryAccNo
     
      Bal = getBalance(AccNo, Dated)
      
      
      Amt = ProductionAmt - PTax
      Bal = Bal + Amt
    
    
      'Debit To Staff Long/Short Advance Account
      
      ConAcc.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
        Dated & "','" & VchrNo & "','" & AccNo & _
        "','" & TheDesc & "'," & Amt & ",0," & Bal & ",'',0)"
        
     
        Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
         INextSno = INextSno + 1
         
TAX_ENTRY:

   'Calculate Total Tax
   Amt = STax + DTax + PTax
   
   If Amt = 0 Then GoTo END_OF_VOUCHER

      TheDesc = "Tax Deduction For " & SalFor
      
      AccNo = TaxAccNo
     
      Bal = getBalance(AccNo, Dated)
      
      
      
      Bal = Bal + Amt
    
    
      'TAX ACCOUNT ENTRY
      
      ConAcc.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
                     Dated & "','" & VchrNo & "','" & AccNo & _
                     "','" & TheDesc & "'," & Amt & ",0," & Bal & ",'',0)"
        
     
        Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
        

END_OF_VOUCHER:

    'ConAcc.CommitTrans
    addSalaryVoucher = True
    Exit Function
        
End Function



