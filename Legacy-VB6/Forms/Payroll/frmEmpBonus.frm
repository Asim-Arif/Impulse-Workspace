VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmEmpBonus 
   BorderStyle     =   0  'None
   ClientHeight    =   7110
   ClientLeft      =   390
   ClientTop       =   5565
   ClientWidth     =   11880
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleMode       =   0  'User
   ScaleWidth      =   10641.08
   ShowInTaskbar   =   0   'False
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Top             =   885
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
      Enabled         =   0   'False
   End
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   11670
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Bonuses"
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
         Left            =   4485
         TabIndex        =   5
         Top             =   120
         Width           =   2850
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Bonuses"
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
         Left            =   4500
         TabIndex        =   6
         Top             =   135
         Width           =   2850
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8250
      Top             =   690
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
            Picture         =   "frmEmpBonus.frx":0000
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpBonus.frx":0452
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin VB.Frame FLV 
      Height          =   5670
      Left            =   75
      TabIndex        =   0
      Top             =   1320
      Width           =   11670
      Begin VB.PictureBox PicEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1725
         Negotiate       =   -1  'True
         ScaleHeight     =   225
         ScaleWidth      =   945
         TabIndex        =   14
         Top             =   735
         Visible         =   0   'False
         Width           =   975
         Begin VB.TextBox txtedit 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
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
            Index           =   1
            Left            =   -15
            TabIndex        =   15
            Top             =   0
            Width           =   945
         End
      End
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   885
         Left            =   5160
         TabIndex        =   7
         Top             =   4725
         Width           =   6375
         Begin VB.TextBox txtTotal 
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
            Left            =   4860
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   75
            Width           =   1410
         End
         Begin MSForms.CommandButton cmdPrintPacket 
            Height          =   375
            Left            =   1650
            TabIndex        =   17
            Top             =   435
            Width           =   1500
            ForeColor       =   0
            Caption         =   "Print Packet"
            PicturePosition =   327683
            Size            =   "2646;661"
            Accelerator     =   116
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdSave 
            Height          =   375
            Left            =   75
            TabIndex        =   12
            Top             =   435
            Visible         =   0   'False
            Width           =   1500
            ForeColor       =   0
            Caption         =   "Save"
            PicturePosition =   327683
            Size            =   "2646;661"
            Accelerator     =   83
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
            Left            =   3210
            TabIndex        =   10
            Top             =   435
            Width           =   1500
            ForeColor       =   0
            Caption         =   "Print         "
            PicturePosition =   327683
            Size            =   "2646;661"
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
            Height          =   375
            Left            =   4785
            TabIndex        =   11
            Top             =   435
            Width           =   1500
            ForeColor       =   0
            Caption         =   " Close    "
            PicturePosition =   327683
            Size            =   "2646;661"
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
            Left            =   3510
            TabIndex        =   9
            Top             =   105
            Width           =   1245
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   4515
         Left            =   90
         TabIndex        =   1
         Top             =   195
         Width           =   11460
         _ExtentX        =   20214
         _ExtentY        =   7964
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
            Object.Tag             =   "1500.09"
            Text            =   "Emp ID"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "2200"
            Text            =   "Emp Name"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "2200"
            Text            =   "F Name"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Tag             =   "2000"
            Text            =   "Designation"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Balance"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Amount"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Paid"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin MSComCtl2.DTPicker DTUpto 
      Height          =   300
      Left            =   2400
      TabIndex        =   16
      Top             =   885
      Width           =   1755
      _ExtentX        =   3096
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
      CustomFormat    =   "ddd dd-MMM yyyy"
      Format          =   93061123
      CurrentDate     =   37384
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   4200
      TabIndex        =   3
      Top             =   900
      Width           =   1500
      Caption         =   "Refresh       "
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
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   630
      Width           =   4755
      BackColor       =   11517387
      Caption         =   " Department                           Date"
      Size            =   "8387;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmEmpBonus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DT As Date
Dim Loading As Boolean
Dim LoadDT As Date

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

Private Sub cmbdept_matched()
    If Loading Then Exit Sub
    Call cmdRefresh_Click
End Sub

Private Sub cmbMonth_Change()
    If Loading Or cmbMonth.MatchFound = False Then Exit Sub
    Dim iDays As Integer
    Dim DT As Date
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    iDays = DateAdd("m", 1, DT) - DT
    cmbMonth.Tag = iDays
    txtMax = iDays
    txtMin = "1"
End Sub

Private Sub cmbMonth_Click()

    Call cmdRefresh_Click
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

Private Sub cmbYear_Click()
Call cmdRefresh_Click
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Call PrintReport(False)
End Sub

Private Sub PrintReport(bSalaryPacket As Boolean)
    
    On Error GoTo err
    If LV.ListItems.Count = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    If bSalaryPacket Then
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptBonusPacket.rpt")
    Else
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpBonus.rpt")
    End If
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VEmpLunchAmt.DT}=#" & LoadDT & "# ", rpt
    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub cmdPrintPacket_Click()
    Call PrintReport(True)
End Sub

Public Sub cmdRefresh_Click()
    
    If Not ((Month(DTUpto) = 12 And Day(DTUpto) = 31) Or (Month(DTUpto) = 6 And Day(DTUpto) = 30)) Then
        MsgBox "Bonus Is Only Calculated For 31st Of December and 30th Of June.", vbInformation
        Exit Sub
    End If
    
'    Dim rs As New ADODB.Recordset
'    With rs
'        .Open "Select  Max(DT) From EmpBonus Where DT>='" & DTUpto & "'", con, adOpenForwardOnly, adLockReadOnly
'        If Not IsNull(.Fields(0)) Then
'            LoadDT = .Fields(0)
'            Call LoadList
'            cmdSave.Visible = False
'        Else
'            Call FillList
'            cmdSave.Visible = True
'        End If
'    End With
        
    Dim rs  As New ADODB.Recordset
    With rs
        .Open "Select DT From EmpBonus Where DT='" & DTUpto & "'", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount = 0 Then
            Call FillList
            cmdSave.Visible = True
        Else
            LoadDT = DTUpto
            Call LoadList
            cmdSave.Visible = False
        End If
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
    
    Call StartTrans(con)
    For i = 1 To LV.ListItems.Count
        With LV.ListItems(i)
            con.Execute "Insert Into EmpBonus(EmpID,DT,Balance,Amount,Paid) Values('" & .Text & _
             "','" & DTUpto & "'," & Val(.SubItems(4)) & "," & Val(.SubItems(5)) & "," & Val(.SubItems(6)) & ")"
        End With
    Next
    
    con.CommitTrans
    cmdSave.Visible = False
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub DTPicker1_Change()
    Call FillList
End Sub

Private Sub Form_Activate()
    Loading = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()

    For i = 1 To LV.ColumnHeaders.Count
        LV.ColumnHeaders(i).Tag = LV.ColumnHeaders(i).Width
    Next
    
    Loading = True
    DTUpto = Date
    
    
    
    Dim CYear As Long
    
    
    For i = 2002 To CYear
        cmbYear.AddItem i
    Next i
    
    cmbMonth = Format(GetServerDate, "MMM")
    cmbYear = CYear
    
    
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID", "Where Deptid<>'VDRS'"
    
    cmbDept.Text = "<All Departments>"
    
'    mnupop.Visible = False
 '   cmdSave.Visible = False
End Sub

Private Sub LoadList()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "Select * From VEmpBonus Where DT='" & LoadDT & "'", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, !EmpId, !EmpId)
            ITM.SubItems(1) = !Name & ""
            ITM.SubItems(2) = !FName & ""
            ITM.SubItems(3) = !Designation & ""
            ITM.SubItems(4) = Val(!Balance & "")
            ITM.SubItems(5) = Val(!Amount & "")
            ITM.SubItems(6) = Val(!Paid & "")
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub FillList()
    
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
    Dim cmd As New ADODB.command
    Dim Total As Long, Cond As String
    Dim StartDT As Date
    Dim JoinDT As Date
    Dim JoinMonths As Integer, JoinDays As Double
    Cond = IIf(cmbDept.ID = "0", "", " and DeptID='" & cmbDept.ID & "'")
    Dim Amt As Double
    DT = DTUpto
    Dim bCheckBal As Boolean
    Dim myMaxDT As Date
    Dim myStartDT As Date
    Dim LastDT As Date
    Dim JoinDT1 As Date
    
    Dim iTotalMonths As Integer
    Dim iTotalDays As Integer
    Dim SalaryPerMonth As Integer
    Dim SalaryPerDay As Integer
    
    With rs
        '.Open "Select * From  VEmpLunchAmt Where ((Month(DT)=" & (cmbMonth.ListIndex + 1) & _
        '" and Year(DT)=" & Val(cmbYear) & ") OR DT is null) and isnull(EmpID,'')<>'' " & Cond, con, adOpenForwardOnly, adLockReadOnly
            .Open "Select Max(DT) From EmpBonus", con, adOpenForwardOnly, adLockReadOnly
            
            If IsNull(.Fields(0)) Then
                bCheckBal = False
                If Month(DT) >= 7 Then
                    StartDT = DateSerial(Year(DT), 7, 1)
                Else
                    StartDT = DateSerial(Year(DT) - 1, 7, 1)
                End If
                If Month(DTUpto) = 12 Then
                    LastDT = DateSerial(Year(StartDT), 6, 30)
                    JoinDT1 = DateSerial(Year(StartDT), 3, 31)
                Else
                    LastDT = DateSerial(Year(StartDT), 1, 31)
                    JoinDT1 = DateSerial(Year(StartDT), 9, 30)
                End If
            Else
                myMaxDT = .Fields(0)
                StartDT = .Fields(0)
                If Month(DTUpto) = 12 Then
                    JoinDT1 = DateSerial(Year(StartDT), 3, 31)
                Else
                    JoinDT1 = DateSerial(Year(StartDT), 9, 30)
                End If
                iTotalDays = DateDiff("d", StartDT, DTUpto)
                bCheckBal = True
            End If
            .Close
            
            .Open "Select EmpID,Name,FName,Designation,JoinDate,StartingSalary,ProbationPeriod From Employees Where DeptID Not In(Select DeptID From Departments Where TempDept=1) AND Active=1 AND DeptID<>'VDRS' AND (DiscontinuedOn Is Null OR DiscontinuedOn>'" & DTUpto & "')" & Cond, con, adOpenForwardOnly, adLockReadOnly
            LV.ListItems.Clear
            
            For i = 0 To .RecordCount - 1
'            Debug.Assert !EmpId <> "VSVS-0067"
                Set ITM = LV.ListItems.add(, ![EmpId], ![EmpId])
                ITM.SubItems(1) = ![Name] & ""
                ITM.SubItems(2) = ![FName] & ""
                ITM.SubItems(3) = ![Designation] & ""
    

                ITM.SubItems(4) = 0
                
'                JoinDT = !JoinDate 'DateAdd("d", -1, !JoinDate)
                JoinDT = DateAdd("d", Val(!ProbationPeriod & ""), !JoinDate)
                If JoinDT > StartDT Then
                    myStartDT = JoinDT
                Else
                    If JoinDT > JoinDT1 Then
                        myStartDT = JoinDT
                    Else
                        myStartDT = StartDT
                    End If
                End If
                SalaryPerMonth = Val(!StartingSalary & "") / 12
                SalaryPerDay = Val(!StartingSalary & "") / 365
                JoinMonths = DateDiff("m", DateAdd("d", -1, JoinDT), DT)
                Call DateDiffInMonthsNDays(JoinDT, DTUpto, iTotalMonths, iTotalDays)
                
                If JoinMonths >= 3 Then
'                    If JoinMonths > 6 Then
'                        Amt = Round(Round(Val(!StartingSalary & "") / 2 / 10) * 10)
'                    Else
'                        JoinDays = DateDiff("d", myStartDT, DT) + 1
'                        If JoinDays >= iTotalDays Then
'                            Amt = Round(Round(Val(!StartingSalary & "") / 2 / 10) * 10)
'                        Else
'                            Amt = Round((Round(JoinDays * (Val(!StartingSalary & "") / 365)) / 10)) * 10
'                        End If
                        If iTotalMonths >= 6 Then
                            Amt = Round(Round(Val(!StartingSalary & "") / 2 / 10) * 10)
                        Else
                            Amt = (SalaryPerMonth * iTotalMonths) + (SalaryPerDay * iTotalDays)
                            Amt = Round(Amt / 10) * 10
                        End If
                    'End If
                Else
                    Amt = 0
                End If
                ITM.SubItems(5) = Amt
                ITM.SubItems(6) = Amt
                .MoveNext
            Next
            
        .Close
    End With
    
    Set rs = Nothing
    'Call GetTotal
    
End Sub

Private Sub ComputeIT()
'    On Error GoTo ERR
'    If MsgBox("Social Security For The Selected Month Has Not Been Computed." & vbNewLine & "Do You Want To Compute It...?", vbQuestion + vbYesNo) = vbNo Then
'        Exit Sub
'    End If
'
'    Dim rsList As New ADODB.Recordset
'    Dim cmd As New ADODB.command
'    Dim ITM As ListItem
'    Dim Sql As String
'    Dim Shift As String
'    Dim Total As Long
'    With cmd
'
'        .CommandText = "Select * from Employees Where Social=1"
'
'        Set .ActiveConnection = con
'        .Parameters.Refresh
'
'        '.Parameters(0).Type = adVarChar
'        '.Parameters(0).value = cmbMonth
'        '.Parameters(1).Type = adVarChar
'        '.Parameters(1).value = cmbYear
'
'        rsList.CursorType = adOpenForwardOnly
'        rsList.LockType = adLockReadOnly
'        Set rsList = .Execute
'        LV.ListItems.Clear
'
'    End With
'
'    Dim PerSocial As Double
'    Dim MonthDays As Integer
'
'    Dim Sal As Double
'    Dim iMax As Integer, iMin As Integer
'
'    iMax = Val(txtMax)
'    iMin = Val(txtMin)
'    TDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
'    MonthDays = DateAdd("m", 1, TDate) - TDate
'
'    Dim rsPer As New ADODB.Recordset
'
'    With rsPer
'        .Open "Select SocialSecurity from GeneralSettings", con, adOpenForwardOnly, adLockReadOnly
'        If .RecordCount > 0 Then
'            PerSocial = Val(.Fields(0)) / 100
'        End If
'        .Close
'    End With
'
'    Set rsPer = Nothing
'
'    With rsList
'        For i = 0 To .RecordCount - 1
'
'            Set ITM = LV.ListItems.Add(, ![EmpId], ![EmpId])
'
'            ITM.Tag = PerSocial
'            ITM.SubItems(1) = ![Name] & ""
'            ITM.SubItems(2) = ![FName] & ""
'            ITM.SubItems(3) = ![SocialNo] & ""
'            ITM.SubItems(4) = ![Designation] & ""
'            ITM.SubItems(5) = ![Status] & ""
'
'            Randomize
'
'            ITM.SubItems(6) = Int(Rnd * (iMax - (iMin - 1)) + (iMin))
'            Sal = Val(![SocialSal])
'            ITM.SubItems(7) = Sal
'
'            ITM.SubItems(8) = Round((Sal / MonthDays) * Val(ITM.SubItems(6)))
'            ITM.SubItems(9) = Round(Val(ITM.SubItems(8)) * PerSocial)
'            Total = Total + Val(ITM.SubItems(8))
'            .MoveNext
'        Next
'
'        .Close
'        txtTotal = Total
'
'    End With
'
'    Set rsLedger = Nothing
'    Sql = ""
'
'    cmdSave.Visible = True
'    cmdCancel.Visible = True
'    cmdClose.Visible = False
'    cmdRefresh.Visible = False
'    cmdPrint.Visible = False
'
'    Exit Sub
'ERR:
'    MsgBox ERR.Description
End Sub
Private Sub ShowIt()
'    On Error GoTo ERR
'    Dim rsList As New ADODB.Recordset
'    Dim cmd As New ADODB.command
'
'    Dim ITM As ListItem
'
'    With cmd
'        .CommandText = "Select * from VSSMonthly"
'        .ActiveConnection = con
'        .Parameters.Refresh
'        .Parameters(0).Type = adVarChar
'        .Parameters(0).value = cmbMonth
'        .Parameters(1).Type = adVarChar
'        .Parameters(1).value = cmbYear
'        rsList.CursorType = adOpenForwardOnly
'        rsList.LockType = adLockReadOnly
'        Set rsList = .Execute
'    End With
'
'    With rsList
'        Dim MDays As Integer
'
'        LV.ListItems.Clear
'        For i = 0 To .RecordCount - 1
'            Set ITM = LV.ListItems.Add(, ![EmpId], ![EmpId])
'            ITM.ListSubItems(1) = ![Name] & ""
'            ITM.ListSubItems(2) = ![FName] & ""
'            ITM.ListSubItems(3) = ![SocialNo] & ""
'            ITM.ListSubItems(4) = ![Designation] & ""
'            ITM.ListSubItems(5) = ![DaysWorked] & ""
'            ITM.ListSubItems(6) = ![Salary] & ""
'            MDays = DateAdd("m", 1, ![DT]) - ![DT]
'            ITM.ListSubItems(7) = (Val(![Salary]) / MDays) * Val(![DaysWorked])
'            ITM.ListSubItems(8) = Val(ITM.ListSubItems(7)) * Val(![PerDeduct])
'            .MoveNext
'        Next
'        .Close
'    End With
'    Set rsLedger = Nothing
'    Sql = ""
'    Exit Sub
err:
    MsgBox err.Description
End Sub
Private Sub AddEmps(ITM As ListItem, rs As ADODB.Recordset, Shift As String)
    With rs
        Set ITM = LV.ListItems.add(, ![EmpId] & Shift, ![EmpId] & "", , ICO)
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
 
 

    
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.Count = 0 Then Exit Sub

    For i = 1 To LV.ColumnHeaders.Count
        LV.ColumnHeaders(i).Width = Val(LV.ColumnHeaders(i).Tag)
    Next

    With PicEdit
        LV.SelectedItem.EnsureVisible
        .Move LV.Left + LV.ColumnHeaders(7).Left + 35, LV.Top + LV.SelectedItem.Top + 30, LV.ColumnHeaders(7).Width + 30, LV.SelectedItem.Height
         
        
        'txtedit(1).Move txtedit(1).Width, -15, LV.ColumnHeaders(7).Width
        txtedit(1).Text = LV.SelectedItem.SubItems(6)
        
        .Visible = True
        txtedit(1).SetFocus
    End With
    
    
End Sub




Private Sub mnuPrintSSForm_Click()
    Screen.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptSSForm.rpt")
    Dim f As New frmPrevRpt
    Dim SelFormula As String
    SelFormula = "Month({VSocialSecurity.DT})=" & cmbMonth.List(cmbMonth.ListIndex, 1) & _
    " and Year({VSocialSecurity.DT})=" & Val(cmbYear)
    
    f.ShowReport SelFormula & " ", rpt
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub muuPrintSSSheet_Click()
    
    Screen.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptSocialSecurity.rpt")
    Dim f As New frmPrevRpt
    Dim SelFormula As String
    SelFormula = "Month({VSocialSecurity.DT})=" & cmbMonth.List(cmbMonth.ListIndex, 1) & _
    " and Year({VSocialSecurity.DT})=" & Val(cmbYear)
    
    f.ShowReport SelFormula & " ", rpt
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub txtedit_GotFocus(Index As Integer)
    txtedit(Index).SelStart = 0
    txtedit(Index).SelLength = Len(txtedit(Index))
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)

    On Error GoTo err
    
    If KeyAscii = 27 Then
        PicEdit.Visible = False
    ElseIf KeyAscii = 13 And Index = 1 Then
        With LV.SelectedItem
            .SubItems(6) = Val(txtedit(1))
        End With
    Else
        OnlyNumber (KeyAscii)
    End If
    
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If LCase(Screen.ActiveControl.Name) <> "txtedit" Then PicEdit.Visible = False
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



Private Sub GetTotal()
    
    Dim Total As Double

    For i = 1 To LV.ListItems.Count
        Total = Total + Val(LV.ListItems(i).SubItems(5))
    Next
    txtTotal = Format(Total, "0.00")
    
End Sub
