VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form AttSheet 
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11550
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   7560
   ScaleMode       =   0  'User
   ScaleWidth      =   10345.49
   WindowState     =   2  'Maximized
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11505
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Attendance Sheet"
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
         TabIndex        =   10
         Top             =   120
         Width           =   2685
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Attendance Sheet"
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
         TabIndex        =   11
         Top             =   135
         Width           =   2685
      End
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   0
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
      Height          =   285
      Left            =   2220
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
   Begin MSComCtl2.DTPicker DtTo 
      Height          =   300
      Left            =   7365
      TabIndex        =   2
      Top             =   900
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
      Format          =   22740995
      CurrentDate     =   37081
   End
   Begin MSComCtl2.DTPicker DtFrom 
      Height          =   300
      Left            =   5985
      TabIndex        =   3
      Top             =   900
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
      Format          =   22740995
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
            Picture         =   "AttSheet.frx":0000
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AttSheet.frx":0452
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin VB.Frame FLV 
      Height          =   6045
      Left            =   15
      TabIndex        =   0
      Top             =   1335
      Width           =   11385
      Begin VB.Frame FButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   570
         Left            =   8205
         TabIndex        =   12
         Top             =   5490
         Width           =   2940
         Begin MSForms.CommandButton cmdPrint 
            Height          =   360
            Left            =   0
            TabIndex        =   14
            Top             =   105
            Width           =   1425
            ForeColor       =   0
            Caption         =   "Print         "
            PicturePosition =   327683
            Size            =   "2514;635"
            Picture         =   "AttSheet.frx":0564
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
            Left            =   1485
            TabIndex        =   13
            Top             =   105
            Width           =   1425
            ForeColor       =   0
            Caption         =   " Close      "
            PicturePosition =   327683
            Size            =   "2514;635"
            Picture         =   "AttSheet.frx":0676
            Accelerator     =   67
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Emp ID"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Employee Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Father Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Shift (I)"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Late Hrs (I)"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Shift (II)"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "L Hrs (II)"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   10800
      Picture         =   "AttSheet.frx":0788
      Top             =   840
      Width           =   360
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   30
      TabIndex        =   5
      Top             =   630
      Width           =   8685
      BackColor       =   11517387
      Caption         =   $"AttSheet.frx":0CD0
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
      Left            =   4485
      TabIndex        =   8
      Top             =   900
      Width           =   1455
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "2566;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   8880
      TabIndex        =   4
      Top             =   870
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
Attribute VB_Name = "AttSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                                      
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrintOld_Click()

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
    '.DataFiles(0) = rptPRPath & "\Payroll.mdb"
    .Connect = ConnectStr
    .SelectionFormula = Cond
    
    .Formulas(0) = "FromTo='From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
    
    .Action = 1
    .PageZoomNext
End With

Call HideWait
Exit Sub
ERR:
    MsgBox ERR.Description
    Call HideWait
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo ERR
    If LV.ListItems.Count = 0 Then Exit Sub
    Call ShowWait
    
    con.Execute "Delete from PrintAttSheet"
    Dim rs As New ADODB.Recordset
    Dim DT As Date
    
    With rs
        For i = 1 To LV.ListItems.Count
        
            If LV.ListItems(i) = "" Then
                DT = CDate(LV.ListItems(i).key)
                GoTo EndFor
            End If
            
            con.Execute "Insert into PrintAttSheet(DT,EmpID,EmpName,FName" & _
             ",Shift1,LateHours1,Shift2,LateHours2) Values('" & DT & "','" & _
             LV.ListItems(i) & "','" & LV.ListItems(i).SubItems(1) & "','" & _
             LV.ListItems(i).SubItems(2) & "','" & LV.ListItems(i).SubItems(3) & _
             "','" & LV.ListItems(i).SubItems(4) & "','" & _
             LV.ListItems(i).SubItems(5) & "','" & LV.ListItems(i).SubItems(6) & "')"
EndFor:
             
        Next i
    End With

    Sleep (1000)
    'Dim TDate As Date
    'Dim MonthDays As Integer
    'TDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    'MonthDays = DateAdd("m", 1, TDate) - TDate

    With cr1
        .ReportFileName = RptPRPath & "\NewAttSheet.rpt"
        '.DataFiles(0) = DatabasePath
        .Connect = ConnectStr
        '.Formulas(0) = "From " & "Fromto='" & DtFrom & "  To  " & DtTo & "'"
        .Formulas(0) = "FromTo='From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
        '.Formulas(1) = "ForMonthDays=" & MonthDaysf
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
    Call RefreshLedger
    
    'Call Form_Resize
End Sub



Private Sub Form_Load()

    DtTo = ServerDate
    DtFrom = DateAdd("M", -1, ServerDate)
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    
    cmbShift.AddItem "<Both Shifts>"
    cmbShift.AddItem "First"
    cmbShift.AddItem "Second"
    cmbShift.ListIndex = 0
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID"
    cmbDept.Text = "<All Departments>"
    cmbEmp.Text = "<All Employees>"
    'Set LV.SmallIcons = ImageList1
    
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
 
 'FButs2.Move LV.Left, LV.Top + LV.Height
End Sub



Private Sub RefreshLedger()

    If cmbDept.MatchFound = False Or cmbEmp.MatchFound = False Then Exit Sub
    Dim rsLedger As New ADODB.Recordset
    Dim rsShifts As New ADODB.Recordset
    Dim ITM As ListItem
    Dim Cond As String
    Dim Sql As String
    Dim ICO As String
    Dim Dates As String
    
    If cmbShift.ListIndex = 0 Then
        Cond = " Where Shift in('First','Second') "
    Else
        Cond = " Where Shift='" & cmbShift & "' "
    End If
    
    If cmbEmp.ID = "0" Then
        If cmbDept.ID = "0" Then
            'Cond = Cond & " and DT between '" & DtFrom & "' and '" & DtTo & "' "
            Cond = Cond
        Else
            Cond = Cond & " and DeptID='" & cmbDept.ID & "'"
        End If
    Else
        Cond = Cond & " and empid='" & cmbEmp.ID & "'"
    End If
    Dates = " and DT between '" & DtFrom & "' and '" & DtTo & "' "
    
    Sql = "select DT from VAttendanceSheet " & Cond & Dates & " Group By DT Order By DT Asc"
    
    rsLedger.CursorLocation = adUseClient
    rsLedger.Open Sql, con, adOpenForwardOnly, adLockReadOnly
    Dim Pos As Integer
    Dim LateHours As String
    Dim DT As String
    LV.ListItems.Clear
    With rsShifts
        Do Until rsLedger.EOF
            
            Set ITM = LV.ListItems.Add(, Format(rsLedger![DT] & "", "dd-MMM-yyyy"))
            ITM.Bold = True
            ITM.ListSubItems.Add
            ITM.ListSubItems.Add(, , "<< " & Format(rsLedger![DT], "dddd  dd-MMM-yyyy") & " >>").Bold = True
            DT = rsLedger![DT] & ""
            .Open "Select * from VAttendance " & Cond & " and DT='" & rsLedger![DT] & "' Order by EmpID,Shift", con, adOpenForwardOnly, adLockReadOnly
            
            'For i = 0 To .RecordCount - 1
            Do Until .EOF
                Set ITM = LV.ListItems.Add(, ![EmpId] & Format(![DT] & "", "dd-MMM-yyyy") & ![Shift], ![EmpId] & "")
                ITM.SubItems(1) = ![Name] & ""
                ITM.SubItems(2) = ![FName] & ""
                If ![Shift] = "First" Then
                    ITM.SubItems(3) = ![Attend] & ""
                    ITM.SubItems(4) = IIf(IsNull(![LateHours]), 0, ![LateHours] & "")
                    
                    Pos = .AbsolutePosition
                    .MoveNext
                    If Not (.EOF) Then
                        If Trim(![EmpId]) = ITM.Text And Trim(![DT]) = DT Then
                            ITM.SubItems(5) = ![Attend] & ""
                            LateHours = IIf(IsNull(![LateHours]), 0, ![LateHours] & "")
                        Else
                            ITM.SubItems(5) = ""
                            .MovePrevious
                        End If
                    Else
                        ITM.SubItems(5) = ""
                        .MovePrevious
                    End If
                Else
                    ITM.SubItems(3) = ""
                    ITM.SubItems(4) = "0"
                    ITM.SubItems(5) = ![Attend] & ""
                    LateHours = IIf(IsNull(![LateHours]), 0, ![LateHours] & "")
                End If
                
                ITM.SubItems(6) = LateHours
                .MoveNext
            Loop
            .Close
            rsLedger.MoveNext
        Loop
    End With
    
    Set rsLedger = Nothing
    Sql = ""
    Cond = ""
    ICO = ""
End Sub




