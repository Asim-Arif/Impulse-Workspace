VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form BonusSheet 
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11400
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   7695
   ScaleMode       =   0  'User
   ScaleWidth      =   10641.08
   WindowState     =   2  'Maximized
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   11670
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bonus Sheet"
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
         TabIndex        =   8
         Top             =   120
         Width           =   1875
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bonus Sheet"
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
         TabIndex        =   9
         Top             =   135
         Width           =   1875
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5115
      Top             =   735
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
            Picture         =   "BonusSheet.frx":0000
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BonusSheet.frx":0452
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   5775
      Top             =   795
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
   Begin VB.Frame FLV 
      Height          =   5670
      Left            =   105
      TabIndex        =   0
      Top             =   1290
      Width           =   11670
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   945
         Left            =   8295
         TabIndex        =   11
         Top             =   4725
         Width           =   3030
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1515
            TabIndex        =   12
            Top             =   75
            Width           =   1410
         End
         Begin MSForms.CommandButton cmdSave 
            Height          =   360
            Left            =   45
            TabIndex        =   17
            Top             =   510
            Visible         =   0   'False
            Width           =   1425
            ForeColor       =   0
            Caption         =   "  Save       "
            PicturePosition =   327683
            Size            =   "2514;635"
            Picture         =   "BonusSheet.frx":0564
            Accelerator     =   83
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
            Left            =   1530
            TabIndex        =   16
            Top             =   510
            Width           =   1425
            ForeColor       =   0
            Caption         =   " Close    "
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
            Left            =   45
            TabIndex        =   15
            Top             =   510
            Visible         =   0   'False
            Width           =   1425
            ForeColor       =   0
            Caption         =   "Print         "
            PicturePosition =   327683
            Size            =   "2514;635"
            Picture         =   "BonusSheet.frx":0676
            Accelerator     =   80
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdCancel 
            Height          =   360
            Left            =   1530
            TabIndex        =   14
            Top             =   510
            Visible         =   0   'False
            Width           =   1425
            ForeColor       =   0
            Caption         =   " Cancel     "
            PicturePosition =   327683
            Size            =   "2514;635"
            Picture         =   "BonusSheet.frx":0788
            Accelerator     =   67
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Total Amount :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   195
            Left            =   225
            TabIndex        =   13
            Top             =   105
            Width           =   1245
         End
      End
      Begin VB.TextBox txtLateHours 
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
         Height          =   285
         Left            =   1485
         TabIndex        =   2
         Top             =   780
         Visible         =   0   'False
         Width           =   915
      End
      Begin MSComctlLib.ListView LV 
         Height          =   4515
         Left            =   90
         TabIndex        =   1
         Top             =   195
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   7964
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
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
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emp Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "F Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Department"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Joining Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Salary"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Bonus Amount"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSForms.CommandButton CommandButton1 
         Height          =   360
         Left            =   120
         TabIndex        =   10
         Top             =   5190
         Width           =   2010
         Caption         =   "Save Changes"
         PicturePosition =   327683
         Size            =   "3545;635"
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
   Begin VB.Image Image1 
      Height          =   375
      Left            =   6840
      Picture         =   "BonusSheet.frx":089A
      Top             =   720
      Width           =   465
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   2985
      TabIndex        =   6
      Top             =   885
      Width           =   1500
      Caption         =   "Refresh       "
      PicturePosition =   327683
      Size            =   "2646;635"
      Picture         =   "BonusSheet.frx":0EB0
      Accelerator     =   82
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
      Left            =   360
      TabIndex        =   5
      Top             =   675
      Width           =   2490
      BackColor       =   11517387
      Caption         =   "  Year                  Month"
      Size            =   "4392;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cmbYear 
      Height          =   300
      Left            =   390
      TabIndex        =   4
      Top             =   930
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
   Begin MSForms.ComboBox cmbMonth 
      Height          =   300
      Left            =   1530
      TabIndex        =   3
      Top             =   930
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
End
Attribute VB_Name = "BonusSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TDate As Date
Dim Loading As Boolean
Private Sub cmbShift_Change()
    Call FillList
End Sub

Private Sub cmdCheckAll_Click()
    With LV
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = True
        Next i
    End With
End Sub

Private Sub cmdClearAll_Click()
    With LV
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = False
        Next i
    End With
End Sub

Private Sub cmbMonth_Change()
    If Loading Then Exit Sub
    Dim iDays As Integer
    Dim DT As Date
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    iDays = DateAdd("m", 1, DT) - DT
    txtMax = iDays
    txtMin = "1"
End Sub

Private Sub cmbYear_Change()
    If Loading Then Exit Sub
    Dim iDays As Integer
    Dim DT As Date
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    iDays = DateAdd("m", 1, DT) - DT
    txtMax = iDays
    txtMin = "1"
End Sub

Private Sub cmdCancel_Click()
    cmdSave.Visible = False
    cmdCancel.Visible = False
    cmdRefresh.Visible = True
    cmdClose.Visible = True
    'cmdPrint.Visible = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

On Error GoTo ERR
Call ShowWait
con.BeginTrans
    con.Execute "Delete From PrintBonus"
    For i = 1 To LV.ListItems.Count
        con.Execute "Insert into PrintBonus values ('" & LV.ListItems(i) & "'," & Val(LV.ListItems(i).SubItems(6)) & ")"
    Next
con.CommitTrans
With cr1
    .ReportFileName = RptPRPath & "\BonusSheet.rpt"
    .Formulas(0) = "Month='" & cmbMonth & ", " & cmbYear & "'"
    .Connect = ConnectStr
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
    If Not (cmbYear.MatchFound Or cmbMonth.MatchFound) Then Exit Sub
    Call FillList
    
    'Call Form_Resize
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ERR
    
    
    con.BeginTrans
    With LV
        For i = 1 To .ListItems.Count
            
            con.Execute "insert into SSMonthly(EmpID," & _
             "DT,Salary,PerDeduct,DaysWorked) Values('" & _
             .ListItems(i) & "','" & TDate & "'," & _
             .ListItems(i).SubItems(7) & "," & _
             .ListItems(i).Tag & "," & .ListItems(i).SubItems(6) & ")"
        Next i
    End With
    con.CommitTrans
    
    cmdSave.Visible = False
    cmdCancel.Visible = False
    cmdRefresh.Visible = True
    cmdClose.Visible = True
    cmdPrint.Visible = True
    Exit Sub
ERR:
    MsgBox ERR.Description
End Sub



Private Sub DTPicker1_Change()
    Call FillList
End Sub

Private Sub CommandButton1_Click()
con.BeginTrans
con.Execute "Delete From BonusPaid Where Month(dated)=" & cmbMonth.List(cmbMonth.ListIndex, 1) & " and Year(Dated)=" & cmbYear
For i = 1 To LV.ListItems.Count
    If LV.ListItems(i).Checked Then con.Execute "insert into Bonuspaid(EmpID,Dated) values ('" & LV.ListItems(i) & "','" & DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), Day(ServerDate)) & "')"
Next
con.CommitTrans
End Sub

Private Sub Form_Load()
    Loading = True
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
    cmbMonth = Format(ServerDate, "MMM")
    cmbYear = CYear
    Loading = False
       
End Sub

Private Sub FillList()
    
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim Total As Long
    
    With cmd
      Sql = "(Select Avg(Salary) From SalaryLedger Where EmpID=VEmp.EmpID and DT between '" & DateSerial(Val(cmbYear) - 1, cmbMonth.List(cmbMonth.ListIndex, 1), 1) & "' and '" & DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1) & "') "
      rs.Open "Select EmpID,Name,FName,DeptName,JoinDate,StartingSalary," & Sql & " Bonus ," & _
      "(Select Dated From BonusPaid Where Empid=Vemp.empID and month(Dated)=" & cmbMonth.List(cmbMonth.ListIndex, 1) & " and Year(Dated)=" & cmbYear & ") Paid From VEmp Where convert(char(3),JoinDate,7)='" & cmbMonth & "' and Active=1 and Bonus=1 and Year(JoinDate)<" & cmbYear, con
    End With
    
    With rs
        If .RecordCount > 0 Then
            Dim MDays As Integer
            Dim ITM As ListItem
            LV.ListItems.Clear
            For i = 0 To .RecordCount - 1
                Set ITM = LV.ListItems.Add(, ![EmpId], ![EmpId])
                ITM.SubItems(1) = ![Name] & ""
                ITM.SubItems(2) = ![FName] & ""
                ITM.SubItems(3) = ![Deptname] & ""
                ITM.SubItems(4) = Format(![JoinDate], "dd-MMM-yyyy")
                ITM.SubItems(5) = Format(![StartingSalary] & "", "0.00")
                
                ITM.Checked = Not IsNull(![Paid])
                
                ITM.SubItems(6) = Format(IIf(IsNull(![Bonus]), ![StartingSalary], ![Bonus]), "0.00")
                
                Total = Total + Val(ITM.SubItems(6))
                .MoveNext
            Next
            txtTotal = Total
            cmdPrint.Visible = True
                  
        End If
        
    End With
    rs.Close
    Set rs = Nothing
    
End Sub
Private Sub ComputeIT()
    On Error GoTo ERR
    If MsgBox("Social Security For The Selected Month Has Not Been Computed." & vbNewLine & "Do You Want To Compute It...?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    Dim rsList As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim ITM As ListItem
    Dim Sql As String
    Dim Shift As String
    Dim Total As Long
    With cmd
        .CommandText = "Select * from Employees Where Social=1"
        Set .ActiveConnection = con
        .Parameters.Refresh
        '.Parameters(0).Type = adVarChar
        '.Parameters(0).value = cmbMonth
        '.Parameters(1).Type = adVarChar
        '.Parameters(1).value = cmbYear
        rsList.CursorType = adOpenForwardOnly
        rsList.LockType = adLockReadOnly
        Set rsList = .Execute
        LV.ListItems.Clear
    End With
    Dim PerSocial As Double
    Dim MonthDays As Integer
    
    Dim Sal As Double
    Dim iMax As Integer, iMin As Integer
    
    iMax = Val(txtMax)
    iMin = Val(txtMin)
    TDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    MonthDays = DateAdd("m", 1, TDate) - TDate
    
    Dim rsPer As New ADODB.Recordset
    With rsPer
        .Open "Select SocialSecurity from GeneralSettings", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            PerSocial = Val(.Fields(0)) / 100
        End If
        .Close
    End With
    Set rsPer = Nothing
    
    With rsList
        For i = 0 To .RecordCount - 1
            
            Set ITM = LV.ListItems.Add(, ![EmpId], ![EmpId])
            ITM.Tag = PerSocial
            ITM.SubItems(1) = ![Name] & ""
            ITM.SubItems(2) = ![FName] & ""
            ITM.SubItems(3) = ![SocialNo] & ""
            ITM.SubItems(4) = ![Designation] & ""
            ITM.SubItems(5) = ![Status] & ""
            Randomize
            
            ITM.SubItems(6) = Int(Rnd * (iMax - (iMin - 1)) + (iMin))
            Sal = Val(![SocialSal])
            ITM.SubItems(7) = Sal
            
            
            ITM.SubItems(8) = Round((Sal / MonthDays) * Val(ITM.SubItems(6)))
            ITM.SubItems(9) = Round(Val(ITM.SubItems(8)) * PerSocial)
            Total = Total + Val(ITM.SubItems(8))
            .MoveNext
        Next
        .Close
        txtTotal = Total
    End With
    Set rsLedger = Nothing
    Sql = ""
    
    cmdSave.Visible = True
    cmdCancel.Visible = True
    cmdClose.Visible = False
    cmdRefresh.Visible = False
    cmdPrint.Visible = False
    Exit Sub
ERR:
    MsgBox ERR.Description
End Sub
Private Sub ShowIt()
    On Error GoTo ERR
    Dim rsList As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    
    Dim ITM As ListItem
    
    With cmd
        .CommandText = "Select * from VSSMonthly"
        .ActiveConnection = con
        .Parameters.Refresh
        .Parameters(0).Type = adVarChar
        .Parameters(0).value = cmbMonth
        .Parameters(1).Type = adVarChar
        .Parameters(1).value = cmbYear
        rsList.CursorType = adOpenForwardOnly
        rsList.LockType = adLockReadOnly
        Set rsList = .Execute
    End With
    
    With rsList
        Dim MDays As Integer
        
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set ITM = LV.ListItems.Add(, ![EmpId], ![EmpId])
            ITM.ListSubItems(1) = ![Name] & ""
            ITM.ListSubItems(2) = ![FName] & ""
            ITM.ListSubItems(3) = ![SocialNo] & ""
            ITM.ListSubItems(4) = ![Designation] & ""
            ITM.ListSubItems(5) = ![DaysWorked] & ""
            ITM.ListSubItems(6) = ![Salary] & ""
            MDays = DateAdd("m", 1, ![DT]) - ![DT]
            ITM.ListSubItems(7) = (Val(![Salary]) / MDays) * Val(![DaysWorked])
            ITM.ListSubItems(8) = Val(ITM.ListSubItems(7)) * Val(![PerDeduct])
            .MoveNext
        Next
        .Close
    End With
    Set rsLedger = Nothing
    Sql = ""
    Exit Sub
ERR:
    MsgBox ERR.Description
End Sub
Private Sub AddEmps(ITM As ListItem, rs As ADODB.Recordset, Shift As String)
    With rs
        Set ITM = LV.ListItems.Add(, ![EmpId] & Shift, ![EmpId] & "", , ICO)
        ITM.SubItems(1) = ![Name]
        ITM.SubItems(2) = ![FName]
        ITM.SubItems(3) = Shift
        ITM.SubItems(4) = 0
    End With
    'Itm.SubItems(4) = Format(![DT], "dddd dd MMMM ,yyyy")
        
End Sub

Private Sub Form_Resize()
    
On Error Resume Next

 fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2

 FLV.Move 50, 1230, Me.ScaleWidth - 100, Me.ScaleHeight - 1400
 
 LV.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
 
 fButs1.Move LV.Left + LV.Width - fButs1.Width, LV.Top + LV.Height
 
CommandButton1.Move LV.Left, LV.Top + LV.Height + 50

End Sub

Private Sub LV_DblClick()
'    If LV.ListItems.Count = 0 Then Exit Sub
'    With txtLateHours
'        .Move LV.Left + LV.ColumnHeaders(5).Left + 35, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(5).Width
'        .Text = LV.SelectedItem.SubItems(4)
'        .SelStart = 0
'        .SelLength = Len(.Text)
'        .Visible = True
'        .SetFocus
'    End With
    
End Sub
Private Sub txtlatehours_KeyPress(KeyAscii As Integer)

On Error GoTo ERR
    If KeyAscii = 27 Then
        txtLateHours = ""
        txtLateHours.Visible = False
    ElseIf KeyAscii = 13 Then
         If Val(txtLateHours) > 8 Then
            MsgBox "Invalid Quantity", vbInformation
            txtLateHours.SelStart = 0
            txtLateHours.SelLength = Len(txtLateHours)
            txtLateHours.SetFocus
            Exit Sub
        Else
            LV.SelectedItem.SubItems(4) = Val(txtLateHours)
            'LV.SelectedItem.SubItems(5) = Val(txtLateHours) * Val(LV.SelectedItem.SubItems(2))
            txtLateHours.Visible = False
            'lblTotal = GetTotal
        End If
    Else
        OnlyNumber (KeyAscii)
    End If
    Exit Sub
ERR:
    MsgBox ERR.Description
    
End Sub

Private Sub txtMax_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumber(KeyAscii)
End Sub

Private Sub txtMax_Validate(cancel As Boolean)

    Dim iDays As Integer
    Dim DT As Date
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    iDays = DateAdd("m", 1, DT) - DT
    With txtMax
        If Val(.Text) > iDays Then
            MsgBox "Value Can't Be More Than The Days Of Selected Month.", vbInformation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            cancel = True
        ElseIf Val(.Text) < 2 Then
            MsgBox "Value Can't Be Less Than The 2.", vbInformation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            cancel = True
        ElseIf Val(.Text) < Val(txtMin) Then
            MsgBox "Value Can't Be Less Than Min Limit.", vbInformation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            cancel = True
        End If
    End With
    
End Sub

Private Sub txtMin_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumber(KeyAscii)
End Sub

Private Sub txtMin_Validate(cancel As Boolean)

    Dim iDays As Integer
    Dim DT As Date
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    iDays = DateAdd("m", 1, DT) - DT
    
    With txtMin
        If Val(.Text) < 1 Then
            MsgBox "Value Can't Be Less Than 1.", vbInformation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            cancel = True
        ElseIf Val(.Text) > (iDays - 1) Then
            MsgBox "Value Can't Be More Than The Days Of Selected Month.", vbInformation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            cancel = True
        ElseIf Val(.Text) > Val(txtMax.Text) Then
            MsgBox "Value Can't Be More Than Max Limit.", vbInformation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            cancel = True
        End If
    End With
    
End Sub
