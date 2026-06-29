VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmEmpLunch 
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
      TabIndex        =   15
      Top             =   885
      Width           =   2265
      _ExtentX        =   3995
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
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   11670
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Monthly Lunch Deduction"
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
         Left            =   3098
         TabIndex        =   7
         Top             =   120
         Width           =   5670
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Monthly Lunch Deduction"
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
         Left            =   3113
         TabIndex        =   8
         Top             =   135
         Width           =   5670
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
            Picture         =   "frmEmpLunch.frx":0000
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpLunch.frx":0452
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
         ScaleWidth      =   1905
         TabIndex        =   16
         Top             =   735
         Visible         =   0   'False
         Width           =   1935
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
            Left            =   945
            TabIndex        =   18
            Top             =   0
            Width           =   945
         End
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
            Index           =   0
            Left            =   -15
            TabIndex        =   17
            Top             =   -30
            Width           =   945
         End
      End
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   885
         Left            =   6885
         TabIndex        =   9
         Top             =   4725
         Width           =   4665
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
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   75
            Width           =   1410
         End
         Begin MSForms.CommandButton cmdSave 
            Height          =   375
            Left            =   45
            TabIndex        =   14
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
            Left            =   1590
            TabIndex        =   12
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
            Left            =   3165
            TabIndex        =   13
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
            Left            =   1890
            TabIndex        =   11
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
         NumItems        =   6
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
            Object.Tag             =   "700.15"
            Text            =   "Days"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Tag             =   "1000.06"
            Text            =   "Lunch Amount"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   4935
      TabIndex        =   5
      Top             =   795
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
      TabIndex        =   4
      Top             =   630
      Width           =   4755
      BackColor       =   11517387
      Caption         =   " Department                           Month                   Year"
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
   Begin MSForms.ComboBox cmbYear 
      Height          =   285
      Left            =   3780
      TabIndex        =   3
      Top             =   885
      Width           =   1110
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "1958;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbMonth 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Top             =   885
      Width           =   1365
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2408;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmEmpLunch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DT As Date
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

    On Error GoTo err

    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpLunchDetail.rpt")
    Load frmPrevRpt
    Dim strDeptCond As String
    If cmbDept.ID <> "0" Then
        strDeptCond = " AND {VEmpLunchAmt.DeptID}='" & cmbDept.ID & "'"
    Else
        strDeptCond = ""
    End If
    frmPrevRpt.ShowReport "Year({VEmpLunchAmt.DT})=" & cmbYear & " AND Month({VEmpLunchAmt.DT})=" & cmbMonth.ListIndex + 1 & strDeptCond, rpt

    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description

End Sub

Public Sub cmdRefresh_Click()
    
    If Not cmbYear.MatchFound Or Not cmbMonth.MatchFound Then Exit Sub
    Call FillList
    
End Sub

Private Sub cmdSave_Click()
'    On Error GoTo ERR
'
'
'    con.BeginTrans
'
'
'    Cond = IIf(cmbDept.ID = "0", "", " and EmpID Like '" & cmbDept.ID & "%'")
'
'    con.Execute "Delete From SSMonthly Where Month(dt)=" & Month(Dated) & " and Year(dt)=" & year(Dated) & Cond
'
'
'    With LV
'        For i = 1 To .ListItems.Count
'            If .ListItems(i).Checked Then
'                con.Execute "insert into SSMonthly(EmpID,DT,EmpStatus,Salary," & _
'                "DaysWorked,NetSalary,SSAmt) Values('" & _
'                .ListItems(i).key & "','" & Dated & "'," & _
'                IIf(.ListItems(i).SubItems(5) = "Permanent", 1, 0) & "," & _
'                Val(.ListItems(i).SubItems(6)) & "," & Val(.ListItems(i).SubItems(7)) & _
'                "," & Val(.ListItems(i).SubItems(8)) & "," & Val(.ListItems(i).SubItems(9)) & ")"
'            End If
'        Next i
'    End With
'
'    con.CommitTrans
'    MsgBox "Saved Successfully", vbInformation
'    cmdSave.Visible = False
'
'    Exit Sub
'ERR:
'    MsgBox ERR.Description

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
    CYear = Format(GetServerDate, "yyyy")
    
    For i = 2002 To CYear
        cmbYear.AddItem i
    Next i
    
    cmbMonth = Format(GetServerDate, "MMM")
    cmbYear = CYear
    
    With cmbDept
        .ListHeight = 2200
        .ClearVals
        .AddItem "<All Departments>", "0"
        .AddVals con, "Name", "Departments", "DeptID", "Where Deptid<>'VDRS'"
    
        .Text = "<All Departments>"
    End With
End Sub

Private Sub FillList()
    
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.command
    Dim Total As Long, Cond As String
    Dim SPCond As String
    Cond = IIf(cmbDept.ID = "0", "", " and DeptID='" & cmbDept.ID & "'")
    SPCond = IIf(cmbDept.ID = "0", "null", "'" & cmbDept.ID & "'")
    DT = DateSerial(Val(cmbYear), cmbMonth.ListIndex + 1, 1)
    
    With rs
        '.Open "Select * From  VEmpLunchAmt Where ((Month(DT)=" & (cmbMonth.ListIndex + 1) & _
        '" and Year(DT)=" & Val(cmbYear) & ") OR DT is null) and isnull(EmpID,'')<>'' " & Cond, con, adOpenForwardOnly, adLockReadOnly
        
            .Open "execute sp_EmpLunchAmt '" & Format(DT, "dd-MMM-yyyy") & "'," & SPCond, con, adOpenForwardOnly, adLockReadOnly
            LV.ListItems.Clear
            
            For i = 0 To .RecordCount - 1
                Set ITM = LV.ListItems.add(, ![EmpId], ![EmpId])
                ITM.SubItems(1) = ![Name] & ""
                ITM.SubItems(2) = ![FName] & ""
                ITM.SubItems(3) = ![Designation] & ""
                ITM.SubItems(4) = ![LunchDays] & ""
                ITM.SubItems(5) = IIf(IsNull(![LunchDays]), "< None >", Format(Val(![LunchAmt] & ""), "0.00"))
                .MoveNext
            Next
            
        .Close
    End With
    
    Set rs = Nothing
    Call GetTotal
    
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
        .Move LV.Left + LV.ColumnHeaders(5).Left + 35, LV.Top + LV.SelectedItem.Top + 30, LV.ColumnHeaders(5).Width + LV.ColumnHeaders(6).Width + 30, LV.SelectedItem.Height
         
        txtedit(0).Move 0, -15, LV.ColumnHeaders(5).Width
        txtedit(0).Text = LV.SelectedItem.SubItems(4)
 
        txtedit(1).Move txtedit(0).Width, -15, LV.ColumnHeaders(6).Width
        txtedit(1).Text = LV.SelectedItem.SubItems(5)
        
        .Visible = True
        txtedit(0).SetFocus
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

Private Sub txtEdit_Change(Index As Integer)

    If Index = 0 Then
        txtedit(1) = Val(txtedit(0)) * 10
    ElseIf Index = 1 Then
'        txtedit(0) = Val(txtedit(1)) / 10
    End If
    
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
            con.Execute "Update EmpLunchAmt Set LunchAmt=" & Val(txtedit(1)) & " Where EmpID='" & _
            .key & "' and DT='" & DT & "'", a
            
            'If not previously added and lunch amount is greater than zero then add it
            If a = 0 And Val(txtedit(1)) > 0 Then
                con.Execute "Insert Into EmpLunchAmt(EmpID,DT,LunchAmt) Values('" & .key & "','" & _
                Format(DT, "dd-MMM-yyyy") & "','" & Val(txtedit(1)) & "')"
            ElseIf a <> 0 And Val(txtedit(1)) = 0 Then
                con.Execute "Delete From EmpLunchAmt Where EmpID='" & .key & "' AND DT='" & DT & "'"
            End If
        End With
        Call FillList
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
