VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vBalIml6.ocx"
Begin VB.Form frmEmpLedger 
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11985
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   8715
   ScaleMode       =   0  'User
   ScaleWidth      =   10735.12
   WindowState     =   2  'Maximized
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   0
      Pattern         =   "*.txt"
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   2340
   End
   Begin MSWinsockLib.Winsock tcpClientS 
      Left            =   5460
      Top             =   2865
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   90
      TabIndex        =   12
      Top             =   885
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextForeColor   =   12582912
      ListIndex       =   -1
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   5730
      Top             =   3285
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker DT 
      Height          =   300
      Left            =   7560
      TabIndex        =   10
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
      Format          =   116981763
      CurrentDate     =   37384
   End
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   11895
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
         Left            =   5055
         TabIndex        =   4
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
         Left            =   5070
         TabIndex        =   5
         Top             =   135
         Width           =   2685
      End
   End
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   3450
      TabIndex        =   2
      Top             =   885
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextForeColor   =   12582912
      ListIndex       =   -1
   End
   Begin VB.Frame FLV 
      Height          =   7485
      Left            =   60
      TabIndex        =   0
      Top             =   1200
      Width           =   11925
      Begin vbAcceleratorSGrid6.vbalGrid EmpGrid 
         Height          =   6765
         Left            =   60
         TabIndex        =   13
         Top             =   150
         Width           =   11745
         _ExtentX        =   20717
         _ExtentY        =   11933
         RowMode         =   -1  'True
         GridLines       =   -1  'True
         NoVerticalGridLines=   -1  'True
         GridLineMode    =   1
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         GridLineColor   =   15263976
         GridFillLineColor=   15263976
         GroupingAreaBackColor=   14737632
         AlternateRowBackColor=   -2147483648
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DisableIcons    =   -1  'True
      End
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   480
         Left            =   60
         TabIndex        =   6
         Top             =   6915
         Width           =   11805
         Begin MSForms.CommandButton CommandButton1 
            Height          =   375
            Left            =   0
            TabIndex        =   17
            Top             =   90
            Visible         =   0   'False
            Width           =   2790
            ForeColor       =   0
            Caption         =   "Abs. Leave Detail (Monthwise)"
            PicturePosition =   327683
            Size            =   "4921;661"
            Accelerator     =   68
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdPrintALSummary 
            Height          =   375
            Left            =   4725
            TabIndex        =   15
            Top             =   90
            Visible         =   0   'False
            Width           =   2370
            ForeColor       =   0
            Caption         =   "Absent / Leave Summary"
            PicturePosition =   327683
            Size            =   "4180;661"
            Accelerator     =   83
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdPrintALDetail 
            Height          =   375
            Left            =   7845
            TabIndex        =   14
            Top             =   90
            Width           =   1860
            ForeColor       =   0
            Caption         =   "Absent Leave Detail"
            PicturePosition =   327683
            Size            =   "3281;661"
            Accelerator     =   68
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
            Left            =   9735
            TabIndex        =   8
            Top             =   90
            Width           =   1020
            ForeColor       =   0
            Caption         =   "Print         "
            PicturePosition =   327683
            Size            =   "1799;661"
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
            Left            =   10800
            TabIndex        =   7
            Top             =   90
            Width           =   1005
            ForeColor       =   0
            Caption         =   " Close      "
            PicturePosition =   327683
            Size            =   "1773;661"
            Accelerator     =   67
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAttVerification 
            Height          =   375
            Left            =   6945
            TabIndex        =   18
            Top             =   90
            Visible         =   0   'False
            Width           =   2760
            ForeColor       =   0
            Caption         =   "Attendance Verification Sheet"
            PicturePosition =   327683
            Size            =   "4868;661"
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
   End
   Begin vbalIml6.vbalImageList ilsIcons 
      Left            =   7455
      Top             =   900
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   4592
      Images          =   "frmEmpLedger.frx":0000
      Version         =   131072
      KeyCount        =   4
      Keys            =   "ÿÿÿ"
   End
   Begin MSForms.CommandButton cmdImportFromTxts 
      Height          =   360
      Left            =   9405
      TabIndex        =   19
      Top             =   825
      Width           =   2265
      Caption         =   "Import from Text files"
      PicturePosition =   327683
      Size            =   "3995;635"
      Accelerator     =   73
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdRefreshRFID 
      Height          =   360
      Left            =   9375
      TabIndex        =   11
      Top             =   570
      Visible         =   0   'False
      Width           =   2610
      Caption         =   "Refresh Data From RFIDs"
      PicturePosition =   327683
      Size            =   "4604;635"
      Accelerator     =   82
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdRefreshSoca 
      Height          =   360
      Left            =   9375
      TabIndex        =   16
      Top             =   570
      Visible         =   0   'False
      Width           =   2610
      Caption         =   "Refresh Data From BioMetric"
      PicturePosition =   327683
      Size            =   "4604;635"
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmEmpLedger.frx":1210
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   105
      TabIndex        =   9
      Top             =   630
      Width           =   9180
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   9375
      TabIndex        =   1
      Top             =   570
      Visible         =   0   'False
      Width           =   2610
      Caption         =   "Refresh        "
      PicturePosition =   327683
      Size            =   "4604;635"
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
Attribute VB_Name = "frmEmpLedger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalPresents As Double, TotalAbsents As Double, TotalLeaves As Double
Dim Sql As String
Dim FirstRefresh As Boolean
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/
'Private Declare Sub GetLocalTime Lib "kernel32" (TT As SYSTEMTIME)

'Dim str, str2 As String
'Dim Week(6) As String
'Dim sLF As String
'Dim mf As HSF1000MachInfo
'Dim h As HSF1000HistoryBank
'Dim pr As HSF1000Parameter
'Dim le As HSF1000LastError
'Dim t As SYSTEMTIME
'Dim usr As HSF1000UserBank
'Dim tf As HSF1000TimeFrameBank
'Dim hc As HSF1000HolidayCtrl
'Dim hday As HSF1000Holiday
'Dim tz As HSF1000TimeZone
'Dim img As HSF1000FPUser
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/

Private Sub cmbdept_matched()
    'Call RefreshLedger
End Sub

Private Sub cmbEmp_matched()
    Call RefreshLedger
    Exit Sub
End Sub

Private Sub cmdAttVerification_Click()

    On Error GoTo err
    
    Me.MousePointer = vbHourglass
    
    Dim cmd As New ADODB.Command
    
    With cmd
        .CommandTimeout = 120
        .CommandType = adCmdStoredProc
        .CommandText = "SP_PrintDailyAbsentLeaveDetail"
        Set .ActiveConnection = con
        .Parameters("@DT").Value = DT.Value
        .Execute
    End With
    
    Dim rpt As CRAXDDRT.Report, f As New frmPrevRpt, strSelFor As String
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptDailyAttendance.rpt")
    strSelFor = "{Employees.Active}=True AND {VMainDepartments.TempDept}=False"
    If cmbDept.ID <> "0" Then
        strSelFor = strSelFor & " AND {VMainDepartments.DeptID}='" & cmbDept.ID & "'"
    End If
    rpt.FormulaFields(2).Text = "#" & Format(DT.Value, "dd-MMM-yyyy") & "#"
    
    f.ShowReport strSelFor, rpt
    Me.MousePointer = vbNormal
    
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPresentEmps_Click()

    Screen.MousePointer = vbHourglass

    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptPresentEmpList.rpt")
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

    Set FormulaFields = rpt.FormulaFields

    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & strCompany & "'"
        ElseIf FormulaField.Name = "{@ForDate}" Then
            FormulaField.Text = "'" & Format(DT, "dddd dd-MMM-yyyy") & "'"
        End If
    
    Next
    
'    On Error Resume Next
'    SrcReport.Database.Tables(1).ConnectBufferString = "DSN=DSN_Vision;;User ID=sa;;Database=Vision;;Password=visionasim"
'    rpt.SQLQueryString = "Select * from VEmp Where EmpID In (Select EmpID From AttendanceSheet Where Attendance=1 and DT='" & _
    Format(DT, "dd-MMM-yyyy") & "')"
    
    f.ShowReport " ", rpt, "Select * from VEmp Where EmpID In (Select EmpID From AttendanceSheet Where Attendance=1 and DT='" & _
     Format(DT, "dd-MMM-yyyy") & "')"

    Screen.MousePointer = vbDefault

End Sub



Private Sub cmdPrint_Click()
    
    Call cmdAttVerification_Click
    Exit Sub
    On Error GoTo err
    Screen.MousePointer = vbHourglass
    Dim Status As Integer
    
    Call StartTrans(con)
    con.Execute "DELETE FROM PrintActivitySheet"
    'For i = 1 To LV.ListItems.Count
    '
    '    With LV.ListItems(i)
    '        Status = IIf(.ListSubItems(2).Bold And Len(.Text) > 2, 0, 1)
    '        con.Execute "INSERT INTO PrintActivitySheet VALUES('" & .Tag & "','" & .SubItems(1) _
    '            & "','" & .SubItems(2) & "','" & .SubItems(3) & "','" & .SubItems(4) & _
    '            "','" & .SubItems(5) & "','" & .SubItems(6) & "','" & .SubItems(7) & "','" & Status & "')"
    '    End With
    'Next
    
    With EmpGrid
    
        For i = 1 To .Rows
            Status = IIf(.CellText(i, 11) <> "", 1, 0)
            
            con.Execute "Insert Into PrintActivitySheet(EmpID,Name,InTime,OutTime,Hrs," & _
            "LateInTime,ShortTime,ExcessTime,Status) Values('" & .CellText(i, 2) & "','" & _
            .CellText(i, 3) & "','" & .CellText(i, 4) & "','" & .CellText(i, 5) & "','" & _
            .CellText(i, 6) & "','" & .CellText(i, 7) & "','" & .CellText(i, 8) & "','" & _
            .CellText(i, 9) & "'," & Status & ")"
        Next
        
    End With
    con.CommitTrans
    
    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpActivitySheet.rpt")
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & strCompany & "'"
        ElseIf FormulaField.Name = "{@FromTo}" Then
            FormulaField.Text = "'" & Format(DT, "dd-MMM-yyyy") & "'"
        End If
    Next

    'SelFormula = "Year({Holidays.DT})=" & cmbYear
    'rpt.SQLQueryString = SQL

    f.ShowReport Cond & " ", rpt

    Screen.MousePointer = vbDefault
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdPrintAttComp_Click()

    On Error GoTo err
    
    Dim TotDays As Integer
    Dim StartDT As Date
    StartDT = DateAdd("d", -Day(DT) + 1, DT)
    
    TotDays = DateDiff("d", StartDT, DateAdd("M", 1, StartDT))
    
    Dim DayAtt(31) As Integer
    Dim CurrDT As Date
    Dim rs As New ADODB.Recordset
    rs.LockType = adLockReadOnly
    rs.CursorType = adOpenForwardOnly
    Set rs.ActiveConnection = con
    
    Call StartTrans(con)
    con.Execute "Delete From PrintAttComparison"
    
    For i = 1 To TotDays
        CurrDT = DateAdd("d", i - 1, StartDT)
        
        rs.Open "Select Count(Distinct EmpID) From EmpTimes Where Cast(InTime As DateTime)<='08:00' AND DT='" & CurrDT & "'"
        DayAtt(i) = rs.Fields(0) & ""
        rs.Close
    Next i
    
    
    Call SaveIt(DayAtt, "Before 08:00", 1)
    
    
    Dim STime As Date
    Dim ETime As Date
    STime = #8:00:00 AM#
    
    For j = 1 To 12
    
        For i = 1 To TotDays
        
            CurrDT = DateAdd("d", i - 1, StartDT)
            
            rs.Open "Select Count(Distinct EmpID) From EmpTimes Where Cast(InTime As DateTime)>'" & STime & "' AND Cast(InTime As DateTime)<='" & DateAdd("n", 5, STime) & "' AND DT='" & CurrDT & "'"
            DayAtt(i) = rs.Fields(0) & ""
            
            rs.Close
        Next i
        Call SaveIt(DayAtt, STime & " To " & DateAdd("n", 5, STime), j + 1)
        STime = DateAdd("n", 5, STime)
    Next j
    
    For i = 1 To TotDays
        CurrDT = DateAdd("d", i - 1, StartDT)
        
        rs.Open "Select Count(Distinct EmpID) From EmpTimes Where Cast(InTime As DateTime)>'09:00' AND Cast(InTime As DateTime)<'10:00'AND DT='" & CurrDT & "'"
        DayAtt(i) = rs.Fields(0) & ""
        rs.Close
    Next i
        
    Call SaveIt(DayAtt, "After 09:00", 14)
    
    
    For i = 1 To TotDays
        CurrDT = DateAdd("d", i - 1, StartDT)
        
        rs.Open "Select Count(Distinct EmpID) From EmpTimes Where OutTime Is Null AND DT='" & CurrDT & "'"
        DayAtt(i) = rs.Fields(0) & ""
        rs.Close
    Next i
        
    Call SaveIt(DayAtt, "No Outgoing", 15)
    
    con.CommitTrans
        
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptAttComparison.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport " ", rpt
    
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub SaveIt(DayAtt() As Integer, TimeDesc As String, SortID As Integer)

    Dim strSQL As String
    Dim strtemp As String
    strSQL = "Insert Into PrintAttComparison Values(" & SortID & ",'" & TimeDesc & "',"
    
    For i = 1 To 31
        strtemp = strtemp & DayAtt(i) & ","
    Next i
    strSQL = strSQL & Left(strtemp, Len(strtemp) - 1) & ")"
    
    con.Execute strSQL
    
End Sub

Private Sub cmdPrintExcessTimes_Click()

    Screen.MousePointer = vbHourglass
    Dim Status As Integer
    
    
    con.Execute "DELETE FROM PrintActivitySheet"
    For i = 1 To LV.ListItems.count
        
        With LV.ListItems(i)
            If .SubItems(7) <> "" Or Len(.Text) = 2 Then
                Status = IIf(.ListSubItems(2).Bold And Len(.Text) > 2, 0, 1)
                con.Execute "INSERT INTO PrintActivitySheet VALUES('" & .Tag & "','" & .SubItems(1) _
                & "','" & .SubItems(2) & "','" & .SubItems(3) & "','" & .SubItems(4) & _
                "','" & .SubItems(5) & "','" & .SubItems(6) & "','" & .SubItems(7) & "','" & Status & "')"
            End If
        End With
    Next


    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpActivitySheet.rpt")
    
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
    
    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & strCompany & "'"
        ElseIf FormulaField.Name = "{@FromTo}" Then
            FormulaField.Text = "'For " & cmbMonth & ", " & cmbYear & ".'"
        End If
    Next

    'SelFormula = "Year({Holidays.DT})=" & cmbYear
    'rpt.SQLQueryString = SQL
    
    f.ShowReport Cond & " ", rpt
    
    Screen.MousePointer = vbDefault


End Sub

Private Sub cmdPrintALDetail_Click()
    Call PrintDailyAbsentLeave(False)
End Sub

Private Sub cmdPrintALSummary_Click()
    Call PrintDailyAbsentLeave(True)
End Sub

Private Sub PrintDailyAbsentLeave(Optional bSummary As Boolean = False, Optional bDateRange As Boolean)

    On Error GoTo err
    
    Me.MousePointer = vbHourglass
    
    Dim cmd As New ADODB.Command
    
    With cmd
        .CommandTimeout = 120
        .CommandType = adCmdStoredProc
        If bSummary Then
            .CommandText = "SP_PrintDailyAbsentLeaveSummary"
        Else
            If bDateRange Then
                .CommandText = "SP_PrintDailyAbsentLeaveDetailDateRange"
            Else
                .CommandText = "SP_PrintDailyAbsentLeaveDetail"
            End If
        End If
        Set .ActiveConnection = con
        
        If bDateRange Then
            .Parameters("@DT").Value = DateSerial(DT.year, DT.Month, 1)
            .Parameters("@DTTo").Value = DateSerial(DT.year, DT.Month + 1, 0)
        Else
            .Parameters("@DT").Value = DT.Value
        End If
        .Execute
    End With
    
    Dim rpt As CRAXDDRT.Report, f As New frmPrevRpt, strSelFor As String
    If bSummary Then
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptAbsentLeaveSummary.rpt")
        rpt.FormulaFields(2).Text = "'" & Format(DT.Value, "dd-MMM-yyyy") & "'"
    Else
        If bDateRange Then
            Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptAbsentLeaveDetailDateRange.rpt")
            strSelFor = "{Employees.Active}=TRUE AND {VMainDepartments.TempDept}=False"
            rpt.FormulaFields(2).Text = "'" & Format(DateSerial(DT.year, DT.Month, 1), "dd-MMM-yyyy") & " to " & Format(DateSerial(DT.year, DT.Month + 1, 0), "dd-MMM-yyyy") & "'"
        Else
            Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptAbsentLeaveDetail.rpt")
            strSelFor = "{PrintDailyAbsentLeaveDetail.Present}=False"
            rpt.FormulaFields(2).Text = "'" & Format(DT.Value, "dd-MMM-yyyy") & "'"
        End If
       
    End If
    
    If cmbDept.ID <> "0" Then
        strSelFor = strSelFor & " AND {Employees.DeptID}='" & cmbDept.ID & "'"
    End If
    
    f.ShowReport strSelFor, rpt
    Me.MousePointer = vbNormal
    
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Public Sub cmdRefresh_Click()

    Call RefreshLedger
    
    'Call Form_Resize
End Sub

Private Sub cmdRefreshRFID_Click()
    
    MsgBox "Not Implemented!!!", vbInformation
    Exit Sub
    Screen.MousePointer = vbHourglass
    Dim Rec As New ADODB.Recordset
    With Rec
        .Open "SELECT CompName FROM Computers", con, adOpenForwardOnly, adLockReadOnly
        For i = 1 To .RecordCount
            If tcpClient.State <> sckClosed Then tcpClient.Close
            tcpClient.RemoteHost = !CompName & ""
            tcpClient.RemotePort = 1012
            tcpClient.Connect
        Next
        .Close
    End With
    Set Rec = Nothing
    Call RefreshLedger
   
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub cmdRefreshSocaOld_Click()
    
    Screen.MousePointer = vbHourglass
    Dim Rec As New ADODB.Recordset
    With Rec
        .Open "SELECT CompName FROM Computers", con, adOpenForwardOnly, adLockReadOnly
        For i = 1 To .RecordCount
            If tcpClientS.State <> sckClosed Then tcpClientS.Close
            tcpClientS.RemoteHost = !CompName & ""
            tcpClientS.RemotePort = 1011
            tcpClientS.Connect
        Next
        .Close
    End With
    Set Rec = Nothing
    Call RefreshLedger
   
    Screen.MousePointer = vbDefault

End Sub

'Private Sub cmdRefreshSoca_Click()
'
'    Dim DTStart As Date, iRecordsRead As Integer
'
'    frmWait.Show , Me
'    DTStart = Now
'    iRecordsRead = 0
'    Call ReadFromBiometric(iRecordsRead)
'    Unload frmWait
'
'    con.Execute "Insert Into BioMetricsRefreshHistory(RefreshStart,RefreshFinish,UserName,ComputerName," & _
'     "ValidRecordsRead) Values('" & DTStart & "','" & Now & "','" & CurrentUserName & "','" & _
'     strComputerName & "'," & iRecordsRead & ")"
'
'    Dim strHTMLBody As String, strBody As String
'    strHTMLBody = "<FONT Face='Tahoma' Size='2' Color='#000000'><BR>" & _
'                   "Machine Data Refreshed:<BR><BR>" & _
'                "<BR></Font><FONT Face='Tahoma' Size='2' Color='#CC3300'>" & _
'                "Start Time:               " & DTStart & "<BR>" & _
'                "Finish Time:  " & Now & "<BR>" & _
'                "Computer Name:     " & strComputerName & "<BR>" & _
'                "User Name:           " & CurrentUserName & "<BR><BR>" & _
'                "Valid Records Read:      " & iRecordsRead & "<BR><BR>" & _
'                "</Font><FONT Face='Tahoma' Size='2' Color='#000000'><BR><BR>Regards,</Font>"
'
''    Call SendEmail("asim@vision.com.pk,jamil@vision.com.pk", "Data Refresh", strHTMLBody, strBody)
'
'End Sub

Private Sub CommandButton1_Click()
    Call PrintDailyAbsentLeave(False, True)
End Sub

Private Sub DT_Change()
    Call RefreshLedger
End Sub

Private Sub Form_Load()

    FirstRefresh = True
    ConfigureGrid

'    Dim CYear As Long
'    CYear = Format(GetServerDate, "yyyy")
'    For i = 2002 To CYear
'        cmbYear.AddItem i
'    Next i
'
'
'    cmbMonth.ListIndex = Val(Format$(GetServerDate, "MM")) - 1
'    cmbYear = Format$(GetServerDate, "yyyy")
'
'
'
'    DTTo = GetServerDate
'    DTFrom = DateAdd("M", -1, GetServerDate)
'    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    cmbDept.ListHeight = 1250
    DT = Date
    
'    cmbDept.ClearVals
'    cmbDept.AddItem "<All Departments>", "0"
'    cmbDept.AddVals con, "Name", "Departments", "DeptID"
'    cmbDept.Text = "<All Departments>"
    'cmbEmp.Text = "<All Employees>"
    'Set LV.SmallIcons = ImageList1
    
    cmbEmp.ClearVals
    cmbEmp.AddItem "< All Employees >", "0"
    cmbEmp.AddVals con, "Name + ' {' + EmpID + '}'", "Employees", "EmpID", "Where Active=1"
    
    cmbDept.ClearVals
    cmbDept.AddItem "< All Departments >", "0"
    cmbDept.AddVals con, "Name + ' {' + DeptID + '}'", "Departments", "DeptID"
    
    cmbDept.ID = "0"
    cmbEmp.ID = "0"
    
End Sub
Private Sub Form_Resize()
    
    Exit Sub

    On Error Resume Next
    
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
    
    FLV.Move 50, 1230, Me.ScaleWidth - 100, Me.ScaleHeight - 1400
     
     
    'LV.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
     
    EmpGrid.Move 100, 200, Me.Width - 400, FLV.Height - FButs1.Height - 300
     
     
    FButs1.Move EmpGrid.Left + EmpGrid.Width - FButs1.Width, EmpGrid.Top + EmpGrid.Height
     
    FButs2.Move EmpGrid.Left, EmpGrid.Top + EmpGrid.Height
 
End Sub


Private Sub tcpClient_Connect()

    'MainForm.Enabled = False
    frmWait.Show
    tcpClient.SendData "RefreshList"


End Sub

Private Sub tcpClients_Connect()

    'MainForm.Enabled = False
    frmWait.Show
    tcpClientS.SendData "RefreshList"


End Sub

Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)
    
    Dim str As String
    tcpClient.getData str
    If str = "Completed" Then
        Unload frmWait
        'MainForm.Enabled = True
        Call cmdRefresh_Click
    End If
    
End Sub

Private Sub tcpClients_DataArrival(ByVal bytesTotal As Long)
    
    Dim str As String
    tcpClientS.getData str
    If str = "Completed" Then
        Unload frmWait
        'MainForm.Enabled = True
        Call cmdRefresh_Click
    End If
    
End Sub


Private Sub RefreshLedgerOLD()
    
    
    Dim StrText As String
    
    
    Dim Rec As New ADODB.Recordset, PrevEmp As String, ITM As ListItem, Cond As String
    Dim Color As Long, Invert As Boolean, Bold As Boolean
    
    Dim rsDept As New ADODB.Recordset, DeptCond As String
    Dim LateInTime As Long, FirstInTime As String, LastOutTime As String
    Dim TotalWorkingMins As Long, TotalWorkedMins As Long
    Dim LastItmIndex As Long

    'Cond = GetTreeValues
    'If Cond = "" Then LV.ListItems.Clear: Exit Sub
    'Cond = "WHERE EmpID IN(" & Cond & ")"
    'If Cond = "" Then

    
    Cond = "WHERE (DT='" & DT & "' OR ISNULL(DT,'')='') "
    
    
    If cmbDept.ID = "0" Then
        If cmbEmp.ID <> "0" Then Cond = Cond & " AND EmpID='" & cmbEmp.ID & "'"
    Else
        If cmbEmp.ID = "0" Then
            Cond = Cond & " AND DeptID='" & cmbDept.ID & "'"
        Else
            Cond = Cond & " AND  empid='" & cmbEmp.ID & "'"
        End If
        DeptCond = "Where DeptID='" & cmbDept.ID & "'"
    End If


    'SQL = "SELECT * FROM VEmpTimes " & Cond & " ORDER BY EmpID"


    Sql = "SELECT Employees.*, EmpTimes.DT AS DT, EmpTimes.EntryID AS EntryID, EmpTimes.InTime AS InTime," & _
        " EmpTimes.OutTime AS OutTime,EmpTimes.Hrs,EmpTimes.LateHours as lateHours,EmpTimes.ExtraHrs as ExtraHours FROM (SELECT * FROM dbo.VEmptimes1 " & Cond & ") AS EmpTimes RIGHT OUTER JOIN" & _
        " dbo.VEmp as Employees ON EmpTimes.EmpID = Employees.EmpID Where Employees.Active=1"
       
    
    
    LV.ListItems.Clear
    LV.Visible = False
    
    rsDept.Open "Select * from VDeptSettings " & DeptCond & " Order By Sno", con, adOpenForwardOnly, adLockReadOnly

    With Rec
    
    .Open Sql, con, adOpenStatic, adLockReadOnly
    
    For dpt = 0 To rsDept.RecordCount - 1
        
        With rsDept
            Color = ![Color]
            Set ITM = LV.ListItems.add(, , !DeptID & "")
            ITM.ForeColor = Color
            ITM.Bold = True
            ITM.Tag = ![DeptID]
            ITM.ListSubItems.add(, , ![Name]).Bold = True
            ITM.ListSubItems(1).ForeColor = Color
            ITM.ListSubItems.add(, , Format(!EnterTime & "", "HH:NN AMPM")).Bold = True
            ITM.ListSubItems(2).ForeColor = Color
            ITM.ListSubItems.add(, , Format(!ExitTime & "", "HH:NN AMPM")).Bold = True
            ITM.ListSubItems(3).ForeColor = Color
            
            ITM.ListSubItems.add(, , getHrsMin(![WorkingHrs])).Bold = True
            ITM.ListSubItems(4).ForeColor = Color
            
            
            TotalWorkingMins = ![WorkingHrs] * 60
            
            
        End With
        
        
        .Filter = "DeptID='" & rsDept![DeptID] & "'"
        
        For i = 1 To .RecordCount
            
            If PrevEmp <> !EmpId & "" Then
             
                'If !EmpId = "02-0001" Then Stop
                
                
                If LastOutTime <> "" And FirstInTime <> "" Then
                    Set ITM = LV.ListItems(LastItmIndex)
                    
                    TotalWorkedMins = DateDiff("n", CDate(FirstInTime), CDate(LastOutTime))
                    
                    ' Exclude Lunch Time From His Working Time
                    TotalWorkedMins = TotalWorkedMins - (Val(![LunchTime] & "") * 60)
                    
                    If TotalWorkedMins > TotalWorkingMins Then
                        ITM.ListSubItems.add
                        strtxt = MinsToTime(TotalWorkedMins - TotalWorkingMins)
                        ITM.ListSubItems.add(, , IIf(strtxt = "00:00", "", strtxt)).ForeColor = ITM.ForeColor
                    Else
                        strtxt = MinsToTime(TotalWorkingMins - TotalWorkedMins)
                        ITM.ListSubItems.add(, , IIf(strtxt = "00:00", "", strtxt)).ForeColor = ITM.ForeColor
                    End If
                    
                End If
                
               FirstInTime = ![InTime] & ""
                
                Color = rsDept![Color]
                
                Set ITM = LV.ListItems.add(, , !EmpId & "")
                ITM.Tag = ![EmpId]
                ITM.ListSubItems.add(, , !Name & "").ForeColor = Color
                
                If Not IsNull(![InTime]) Then
                    LateInTime = DateDiff("n", CDate(rsDept![EnterTime]), CDate(![InTime]))
                Else
                    LateInTime = 0
                End If
                
                PrevEmp = !EmpId & ""
                
                LastOutTime = ![OutTime] & ""
                LastItmIndex = ITM.Index
            Else
                Set ITM = LV.ListItems.add(, , "")
                ITM.Tag = PrevEmp
                ITM.ListSubItems.add , , ""
                LateInTime = 0
                LastOutTime = ![OutTime] & ""
                LastItmIndex = ITM.Index
            End If
            
            ITM.ForeColor = Color
            ITM.ListSubItems(1).ForeColor = Color
            
            ITM.ListSubItems.add(, , Format(!InTime & "", "HH:NN AMPM")).ForeColor = Color
            ITM.ListSubItems.add(, , Format(!OutTime & "", "HH:NN AMPM")).ForeColor = Color
            
            ITM.ListSubItems.add(, , getHrsMin(Val(!Hrs & ""))).ForeColor = Color
            
            
            If LateInTime <= 0 Then
                ITM.ListSubItems.add(, , "").ForeColor = Color
            Else
                ITM.ListSubItems.add(, , MinsToTime(Val(LateInTime))).ForeColor = Color
            End If
            
            'itm.ListSubItems.Add(, , getHrsMin(Val(!LateHours & ""))).ForeColor = Color
            'itm.ListSubItems.Add(, , getHrsMin(Val(!ExtraHours & ""))).ForeColor = Color
            
            'PrevEmp = !EmpId & ""
            
            If Val(!EntryID & "") > 0 Then
                Bold = False
            Else
                Bold = True
                'Color = vbRed
            End If
            
            ITM.Bold = Bold
            ITM.ForeColor = Color
            
            For k = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(k).Bold = Bold
                ITM.ListSubItems(k).ForeColor = Color
            Next
    
            .MoveNEXT
        Next
        .Filter = ""
        rsDept.MoveNEXT
    Next dpt
    .Close
    LV.Visible = True
    End With


End Sub

Private Sub ConfigureGrid()

    On Error GoTo err
    
    With EmpGrid
        
        .Editable = False
        .ImageList = ilsIcons
        
        .DefaultRowHeight = 15
        .RowMode = True
        
        .AddColumn , "", , , 20, , True, , False
        
        .AddColumn , "Emp ID", , , 67, , , , False
        .AddColumn , "Name", , , 227
        .AddColumn , "In Time", , , 77
        .AddColumn , "Out Time", , , 77
        .AddColumn , "Net Time", , , 77
        .AddColumn , "Late Coming", , , 77
        .AddColumn , "Short Time", , , 77
        .AddColumn , "Excess Time", , , 77
        
        .AddColumn , "Keys", , , , False
        
        .AddColumn , "Details", , , .Width, , , , , , True
        
    End With
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub RefreshLedger()
    
    
    Dim StrText As String
    
    
    Dim Rec As New ADODB.Recordset, PrevEmp As String, ITM As ListItem, Cond As String
    Dim Color As Long, Invert As Boolean, Bold As Boolean
    
    Dim rsDept As New ADODB.Recordset, DeptCond As String
    Dim LateInTime As Long, FirstInTime As String, LastOutTime As String
    Dim TotalWorkingMins As Long, TotalWorkedMins As Long
    Dim LastItmIndex As Long
    
    Cond = "WHERE (DT='" & DT & "' OR ISNULL(DT,'')='') "
    
    
    If cmbDept.ID = "0" Then
        If cmbEmp.ID <> "0" Then Cond = Cond & " AND EmpID='" & cmbEmp.ID & "'"
    Else
        If cmbEmp.ID = "0" Then
            Cond = Cond & " AND DeptID='" & cmbDept.ID & "'"
        Else
            Cond = Cond & " AND  empid='" & cmbEmp.ID & "'"
        End If
        DeptCond = "Where DeptID='" & cmbDept.ID & "'"
    End If

    Sql = "SELECT Employees.DeptID,Employees.Name,Employees.EmpID,Employees.LunchTime, EmpTimes.DT AS DT, EmpTimes.EntryID AS EntryID, EmpTimes.InTime AS InTime," & _
        " EmpTimes.OutTime AS OutTime,EmpTimes.Hrs,EmpTimes.LateHours as lateHours,EmpTimes.ExtraHrs as ExtraHours FROM (SELECT * FROM dbo.VEmptimes1 " & Cond & ") AS EmpTimes RIGHT OUTER JOIN" & _
        " dbo.VEmp as Employees ON EmpTimes.EmpID = Employees.EmpID Where Employees.Active=1"
       
    
    EmpGrid.Clear
    EmpGrid.StretchLastColumnToFit = False
    EmpGrid.Redraw = False
    EmpGrid.Visible = False
    
    Dim cmd As New ADODB.Command
    With cmd
        .CommandTimeout = 120
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_DeptSettings"
        .Parameters("@DT").Value = DT
        If cmbDept.ID <> "0" Then
            .Parameters("@DeptID").Value = cmbDept.ID
        End If
        rsDept.CursorType = adOpenForwardOnly
        rsDept.LockType = adLockReadOnly
        Set rsDept = .Execute
    End With
    '-/-/-/-/-/--/-/---/-/-/--/-
    Dim sFntBLD As New StdFont
    sFntBLD.Name = "Tahoma"
    sFntBLD.Size = 8
    sFntBLD.Bold = True
    
    Dim sFnt As New StdFont
    sFnt.Name = "Tahoma"
    sFnt.Size = 8
    
    Dim iRow As Integer
    Dim strDetail As String
    Dim lHeight As Long
    Dim iPrevRow As Integer
    '-/-/-/-/-/--/-/---/-/-/--/-
    
    With Rec
    
        .Open Sql, con, adOpenStatic, adLockReadOnly
    
        For dpt = 0 To rsDept.RecordCount - 1
        
            With rsDept
                
                sFntBLD.Name = "Verdana"
                
                Color = Val(![Color] & "")
                iRow = EmpGrid.Rows + 1
                
                EmpGrid.CellDetails iRow, 1, , , , , , , , 3
                EmpGrid.CellDetails iRow, 2, ![DeptID] & "", , , , vbBlack, sFntBLD
                EmpGrid.CellDetails iRow, 3, ![Name] & "", , , , vbBlack, sFntBLD
                EmpGrid.CellDetails iRow, 4, Format(![EnterTime] & "", "HH:NN AMPM"), , , , Color, sFntBLD
                EmpGrid.CellDetails iRow, 5, Format(![ExitTime] & "", "HH:NN AMPM"), , , , Color, sFntBLD
                EmpGrid.CellDetails iRow, 6, getHrsMin(![WorkingHrs]), , , , Color, sFntBLD
            
                If Val(![TotalEmps] & "") = 0 Then
                    strDetail = "Dept. Employees : 0        Present : 0     Absent : 0      Late Coming : 0     Short Times : 0     Excess Times : 0"
                Else
                    strDetail = "Dept. Employees : " & ![TotalEmps] & "         Present : " & Val(![TotalEmps]) - Val(![TotalAbs]) & "      Absent : " & ![TotalAbs] & "        Late Coming : " & ![LateComing] & "     Short Times : " & ![ShortTimes] & "    Excess Times : " & ![ExcessTimes]
                End If
                
                
                EmpGrid.CellDetails iRow, 10, ![DeptID]
                EmpGrid.CellDetails iRow, 11, strDetail, DT_WORDBREAK, , , vbRed, sFnt
            
                lHeight = EmpGrid.EvaluateTextHeight(iRow, 9) + EmpGrid.DefaultRowHeight + 2
                EmpGrid.RowHeight(iRow) = lHeight
            
                TotalWorkingMins = ![WorkingHrs] * 60
            
            
            End With
            
            sFntBLD.Name = "Thahoma"
            .Filter = "DeptID='" & rsDept![DeptID] & "'"
        
            For i = 1 To .RecordCount
            
                iRow = EmpGrid.Rows + 1
                
                If PrevEmp <> !EmpId & "" Then
                    Color = Val(rsDept![Color] & "")
                    If LastOutTime <> "" And FirstInTime <> "" Then

                        
                        TotalWorkedMins = DateDiff("n", CDate(FirstInTime), CDate(LastOutTime))
                    
                        ' Exclude Lunch Time From His Working Time
                        TotalWorkedMins = TotalWorkedMins - (Val(![LunchTime] & "") * 60)
                        
                        If TotalWorkedMins > TotalWorkingMins Then

                            strtxt = MinsToTime(TotalWorkedMins - TotalWorkingMins)

                            EmpGrid.CellDetails iPrevRow, 9, IIf(strtxt = "00:00", "", strtxt), , , , Color
                        Else
                            strtxt = MinsToTime(TotalWorkingMins - TotalWorkedMins)

                            EmpGrid.CellDetails iPrevRow, 8, IIf(strtxt = "00:00", "", strtxt), , , , Color
                        End If
                        
                    End If
                
                    FirstInTime = ![InTime] & ""
                    
                    EmpGrid.CellDetails iRow, 2, ![EmpId] & "", , , , , IIf(Val(![EntryID] & "") > 0, sFnt, sFntBLD)
                    EmpGrid.CellDetails iRow, 3, ![Name] & "", , , , Color, IIf(Val(![EntryID] & "") > 0, sFnt, sFntBLD)
                    
                    If Not IsNull(![InTime]) Then
                        LateInTime = DateDiff("n", CDate(rsDept![EnterTime]), CDate(![InTime]))
                    Else
                        LateInTime = 0
                    End If
                
                    PrevEmp = !EmpId & ""
                
                    LastOutTime = ![OutTime] & ""
                    iPrevRow = iRow
                    'EmpGrid.CellDetails iRow, 10, ![EmpId]
                    EmpGrid.CellDetails iRow, 10, ![DeptID]
                Else
                    Color = RGB(150, 150, 255)
                    iRow = EmpGrid.Rows + 1
                    EmpGrid.CellDetails iRow, 2, ""
                    EmpGrid.CellDetails iRow, 3, ""
                    

                    LateInTime = 0
                    LastOutTime = ![OutTime] & ""
                    'LastItmIndex = ITM.Index
                    'iPrevRow = iRow
                End If
                
                EmpGrid.CellDetails iRow, 4, Format(![InTime] & "", "HH:NN AMPM"), , , , Color
                EmpGrid.CellDetails iRow, 5, Format(![OutTime] & "", "HH:NN AMPM"), , , , Color
                EmpGrid.CellDetails iRow, 6, getHrsMin(Val(!Hrs & "")), , , , Color
                
                If LateInTime <= 0 Then
                
                    EmpGrid.CellDetails iRow, 7, "", , , , Color
                Else
                
                    EmpGrid.CellDetails iRow, 7, MinsToTime(Val(LateInTime)), , , , Color
                End If
            
                If Val(!EntryID & "") > 0 Then
                    Bold = False
                Else
                    Bold = True
                    'Color = vbRed
                End If
                
                'EmpGrid.CellDetails iRow, 10, ![EmpId]
                EmpGrid.CellDetails iRow, 10, ![DeptID]
                EmpGrid.RowVisible(iRow) = False 'Not FirstRefresh
                .MoveNEXT
            Next
            .Filter = ""
            rsDept.MoveNEXT
            
        Next dpt
        
        .Close
        
        EmpGrid.StretchLastColumnToFit = True
        EmpGrid.Redraw = True
        EmpGrid.Visible = True
    End With
    FirstRefresh = False

End Sub

Private Sub EmpGrid_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    '''
    On Error GoTo err
    
    If lRow = 0 Then Exit Sub
    
    If EmpGrid.CellText(lRow, 11) <> "" Then
    
        Call SetExpand(lRow, EmpGrid.CellText(lRow, 10))
    
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
Private Sub EmpGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single, bDoDefault As Boolean)

    Dim lCol As Long, lRow As Long
    Dim lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long
    Dim lType As Long, lClause As Long
    Dim lIconIndex As Long
    Dim i As Long

    EmpGrid.CellFromPoint x \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY, lRow, lCol
    
    If lCol = 1 Then
        If EmpGrid.CellText(lRow, 11) <> "" Then
            Call SetExpand(lRow, EmpGrid.CellText(lRow, 10))
        End If
    End If

End Sub
Private Sub SetExpand(lRow As Long, ParentID As String)

    Dim lIconIndex As Long
    Dim i As Long

    ' Set .Redraw = False to loose the animation effect
    ' when doing this (it might be too slow otherwise)
   
    lIconIndex = EmpGrid.CellExtraIcon(lRow, 1)
    If (lIconIndex = 3) Then
        ' Expand
        lIconIndex = 2
        ' Reverse order only so the "animation" looks nice!
        For i = EmpGrid.Rows To 1 Step -1
            'If (Left(EmpGrid.CellText(i, 10), 4) = ParentID) And i <> lRow Then
            If EmpGrid.CellText(i, 10) = ParentID And i <> lRow Then
                EmpGrid.RowVisible(i) = True
            End If
        Next i
    Else
        ' Collapse
        lIconIndex = 3
        For i = lRow + 1 To EmpGrid.Rows
            If (Left(EmpGrid.CellText(i, 10), 4) = ParentID) And i <> lRow Then
                EmpGrid.RowVisible(i) = False
            Else
                Exit For
            End If
        Next i
    End If
    EmpGrid.CellExtraIcon(lRow, 1) = lIconIndex
    
End Sub

'Public Sub ReadFromBiometric(ByRef iRecordRead As Integer)
'
'    On Error GoTo err
'    '/\/\/\/\/\/\/\/\/\/\/\/\/\/OPEN Comm/\/\/\/\/\/\/\/\/\/\/
'    'If OpenComm("RS232,3,9600,8,2", 1) = &HFFFFFFFF Then
'    If OpenComm("TCP,192.168.0.5,4444", 1) = &HFFFFFFFF Then
'        MsgBox str & " Failure occured"
'        MsgBox "Error Code: " & GetErrorCode()
'        MsgBox err.Description
'        Exit Sub
'    End If
'
'    Dim i As Integer
'    Dim Quit As Boolean
'
'    Quit = False
'    Do
'        Call ReadHistory(1, h)
'
'        If TimeOut() <> 1 Then
'            For i = 1 To 50
'                If Left(h.History(i).stDate, 8) = "20000000" Then
'                    Quit = True
'                    Exit For
'                End If
'                If SplitValues(h.History(i)) Then
'                    iRecordRead = iRecordRead + 1
'                End If
'                'AddMess (Left(h.History(i).stDate, 8) & " " & Left(h.History(i).stTime, 6) & " " & Format(h.History(i).stCardNo, "0000000000") & " " & Format(h.History(i).stState, "000"))
'                Debug.Print Left(h.History(i).stDate, 8) & " " & Left(h.History(i).stTime, 6) & " " & Format(h.History(i).stCardNo, "0000000000") & " " & Format(h.History(i).stState, "000")
'            Next
'
'            Call DeleteHistory(1)
'        Else
''            AddMess ("Time Out!")
'            Debug.Print "Time Out!" 'Left(h.History(i).stDate, 8) & " " & Left(h.History(i).stTime, 6) & " " & Format(h.History(i).stCardNo, "0000000000") & " " & Format(h.History(i).stState, "000")
'            Quit = True
'        End If
'    Loop Until Quit
'
'    'MsgBox "Completed", vbInformation
'    Call RefreshLedger
'    Exit Sub
'
'err:
'    'MSComm1.PortOpen = False
'    Debug.Print err.Description
'    MsgBox err.Description
'
'End Sub

'Private Function SplitValues(h As HSF1000History) As Boolean
'
'    If Format(h.stState, "000") <> "002" Then Exit Function
'    Debug.Print Left(h.stDate, 8) & " " & Left(h.stTime, 6) & " " & Format(h.stCardNo, "0000000000") & " " & Format(h.stState, "000")
'    Debug.Print
'    Dim myDT As Date
'    Dim strDT As String, strTime As String
'    Dim myTime As Date
'    Dim myID As Long
'    strDT = Left(h.stDate, 8)
'    strTime = Left(h.stTime, 2) & ":" & Mid(h.stTime, 3, 2)
'
'    myDT = DateSerial(Left(strDT, 4), Mid(strDT, 5, 2), Right(strDT, 2))
'    myTime = CDate(strTime)
'
'    SplitValues = TakeAtt(Trim(h.stCardNo), myDT, myTime)
'
'End Function

Private Function TakeAtt(EmployeeID As String, AttDate As Date, AttTime As Date, lEntryID As Long, Optional p_bByEmpID As Boolean = False, Optional bFromTxtFiles As Boolean = False) As Boolean

    Dim rs As New ADODB.Recordset
    Dim strEmpID As String, strDeptID As String, bExemptSettings As Boolean
    Dim bLunchTimeScanning As Boolean
    Dim strLunchInTime As String, strLunchOutTime As String, iEmpAttCatID As Integer
    Dim iTotalEntries As Integer, bActive As Boolean
    Dim bOverTimePaid As Boolean
    Dim strExitTime As String
    
    With rs
        If p_bByEmpID Then
            .Open "SELECT DeptID,EmpID,ExemptSettings,Active FROM Employees WHERE EmpID='" & EmployeeID & "'", con, adOpenForwardOnly, adLockReadOnly
        Else
            .Open "SELECT DeptID,EmpID,ExemptSettings,Active FROM Employees WHERE EmpID=" & Val(EmployeeID), con, adOpenForwardOnly, adLockReadOnly
        End If
        
        If .EOF = False Then
            strEmpID = !EmpId & ""
            strDeptID = !DeptID & ""
            bExemptSettings = !ExemptSettings
            bActive = !Active
            .Close
        Else
            con.Execute "Insert Into BioMetricsDataInvalid(EmpID,DT,AttTime) Values('" & EmployeeID & "','" & AttDate & "','" & Format(AttTime, "HH:nn") & "')"
            .Close
            Exit Function
        End If
        
        If bActive = False Then
            con.Execute "Insert Into BioMetricsDataInactiveEmps(EmpID,DT,AttTime) Values('" & strEmpID & "','" & AttDate & "','" & Format(AttTime, "HH:nn") & "')"
            Exit Function
        End If
        
        If bExemptSettings Then
            .Open "Select LunchInTime,LunchOutTime,EmpAttCat_RefID,ExitTime From EmpSettings Where EmpID='" & strEmpID & "'", con, adOpenForwardOnly, adLockReadOnly
        Else
            .Open "Select LunchInTime,LunchOutTime,EmpAttCat_RefID,ExitTime From GeneralSettings Where DeptID='" & strDeptID & "'", con, adOpenForwardOnly, adLockReadOnly
        End If
        
        strExitTime = !ExitTime & ""
        strLunchInTime = !LunchInTime & ""
        strLunchOutTime = !LunchOutTime & ""
        iEmpAttCatID = Val(!EmpAttCat_RefID & "")
        
        .Close
        
        .Open "SELECT LunchTimeScanning,OverTimePaid FROM EmployeeAttendanceCatagories WHERE EntryID=" & iEmpAttCatID, con, adOpenForwardOnly, adLockReadOnly
        bLunchTimeScanning = !LunchTimeScanning
        bOverTimePaid = !OverTimePaid
        .Close
        
        .Open "Select TotalEntries From VEmpDailyEntries Where EmpID='" & strEmpID & "' and DT='" & AttDate & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            iTotalEntries = 0
        Else
            iTotalEntries = .Fields(0) & ""
        End If
        .Close
        
    End With
    Set rs = Nothing
    
    If strCompany <> "Tecno" Then
        If bLunchTimeScanning = False Then
            If iTotalEntries = 1 Then
                Call TakeAttendance(strEmpID, AttDate, CDate(strLunchInTime), True)
                DoEvents
                Call TakeAttendance(strEmpID, AttDate, CDate(strLunchOutTime), True)
                DoEvents
    '            Call TakeAttendance(strEmpID, AttDate, CDate(strExitTime), True)
    '            DoEvents
            End If
        End If
    End If
    
    TakeAtt = TakeAttendance(strEmpID, AttDate, AttTime, True)
    
    If strCompany = "Tecno" Then
        If iTotalEntries = 0 Then   'Meaning If the Entry passed was the First One, Add Lunch & Check Out...
            Call TakeAttendance(strEmpID, AttDate, CDate(strLunchInTime), True)
            DoEvents
            Call TakeAttendance(strEmpID, AttDate, CDate(strLunchOutTime), True)
            DoEvents
            Call TakeAttendance(strEmpID, AttDate, CDate(strExitTime), True)
            DoEvents
        End If
    End If
    
    con.Execute "UPDATE txtFilesData SET Processed=1 WHERE EntryID=" & lEntryID
    DoEvents
    
End Function

Private Function TakeAttOld(EmployeeID As String, AttDate As Date, AttTime As Date) As Boolean

    On Error GoTo err
    TakeAttOld = False
    Dim ExtraDeds As Double, ExtraHrs As Double
    ExtraDeds = 0
    ExtraHrs = 0
    
    Dim MultipleAllowed As Boolean
    Dim bDeductLHrs As Boolean
    Dim bExempt As Boolean
    Dim EmpId As String, DeptID As String, bActive As Boolean
    Dim dLateHrs As Double
    Dim TotalEntries As Integer
    Dim AttRecs As Integer
    Dim strInTime As String, strOutTime As String, strOutTimeFormated As String
    
    Dim strEmpInTime As String, strEmpOutTime As String
    Dim strEmpLunchIn As String, strEmpLunchOut As String
    '''''''''''   This Will Is The Employee Exists Or Not    '''''''''''
    
    With con.Execute("SELECT DeptID,EmpID,EmpType,ExemptSettings,Active FROM Employees WHERE EmployeeID='" & Val(EmployeeID) & "'")
        'If !EmpId = "VSRTA-0001" Then Stop
        If .EOF Then
            con.Execute "Insert Into BioMetricsDataInvalid(EmpID,DT,AttTime) Values('" & EmployeeID & "','" & AttDate & "','" & Format(AttTime, "HH:nn") & "')"
            Exit Function
        Else
            EmpId = ![EmpId] & ""
            Emptype = ![Emptype] & ""
            bExempt = ![ExemptSettings] & ""
            DeptID = ![DeptID] & ""
            bActive = !Active
        End If
        .Close
        If bActive = False Then
            con.Execute "Insert Into BioMetricsDataInactiveEmps(EmpID,DT,AttTime) Values('" & EmpId & "','" & AttDate & "','" & Format(AttTime, "HH:nn") & "')"
            Exit Function
        End If
        If bExempt Then
            .Open "Select EnterTime,ExitTime,LunchInTime,LunchOutTime,DeductLHrs From EmpSettings Where EmpID='" & EmpId & "'", con, adOpenForwardOnly, adLockReadOnly
        Else
            .Open "Select EnterTime,ExitTime,LunchInTime,LunchOutTime,DeductLHrs From GeneralSettings Where DeptID='" & DeptID & "'", con, adOpenForwardOnly, adLockReadOnly
        End If
        If .EOF Then
            bDeductLHrs = True
        Else
            bDeductLHrs = ![DeductLHrs]
            strEmpInTime = Format(!EnterTime, "HH:NN")
            strEmpOutTime = Format(!ExitTime, "HH:NN")
            strEmpLunchIn = Format(!LunchInTime, "HH:NN")
            strEmpLunchOut = Format(!LunchOutTime, "HH:NN")
        End If
        .Close
    End With
    
    Dim Rec As New ADODB.Recordset, CurrTime As String, Hrs As Double, Sql As String, Cond As String
    Dim TimeDiff As Integer

    Cond = "WHERE EmpID='" & EmpId & "' AND DT='" & AttDate & "'"

    CurrTime = Format(AttTime, "HH:NN")

    TimeDiff = Val(con.Execute("SELECT TimeDiff FROM TimeSettings").Fields(0) & "")
        
    With Rec
        .Open "Select TotalEntries From VEmpDailyEntries Where EmpID='" & EmpId & "' and DT='" & AttDate & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            TotalEntries = 0
        Else
            TotalEntries = .Fields(0) & ""
        End If
        .Close
        
        .Open "SELECT * FROM EmpTimes WHERE EntryID=(SELECT Max(EntryID) FROM EmpTimes " & Cond & ")", con, adOpenForwardOnly, adLockReadOnly
            
        If TotalEntries > 0 Then
            
            If TotalEntries = 1 And !OutTime & "" = "" Then 'Means It's Lunch In Entry...
                '12:50 to 1:10
                
                If CurrTime < CDate(DateAdd("N", -5, strEmpLunchIn)) Then
                    ExtraDeds = 0 'DateDiff("n", CurrTime, "12:55") / 60
                ElseIf CurrTime > CDate(DateAdd("N", 5, strEmpLunchIn)) Then
                    ExtraDeds = 0 'DateDiff("n", "13:00", CurrTime) / 60
                End If
                If DateDiff("n", strEmpInTime, CDate(!InTime)) > 0 And DateDiff("n", strEmpInTime, CDate(!InTime)) < 11 Then
                    ExtraHrs = DateDiff("n", strEmpInTime, CDate(!InTime)) / 60
                ElseIf DateDiff("n", strEmpInTime, CDate(!InTime)) < 0 Then
                    ExtraDeds = ExtraDeds + DateDiff("n", CDate(!InTime), strEmpInTime) / 60
                End If
                
                If Format(AttDate, "dddd") = "Friday" Then
                    ExtraHrs = ExtraHrs + 0.25
                    If CurrTime > #12:15:00 PM# Then
                        Hrs = DateDiff("n", CDate(!InTime), #12:15:00 PM#)
                    Else
                        Hrs = DateDiff("n", CDate(!InTime), CurrTime)
                    End If
                Else
                    If CurrTime > CDate(DateAdd("N", -6, strEmpLunchIn)) Then '#12:24:00 PM#
                        Hrs = DateDiff("n", CDate(!InTime), strEmpLunchIn) '#12:30:00 PM#
                    Else
                        Hrs = DateDiff("n", CDate(!InTime), CurrTime)
                    End If
                End If
                
            ElseIf TotalEntries = 2 And !OutTime & "" = "" Then 'Means It's Lunch Out Entry
                If Format(AttDate, "dddd") = "Friday" Then
                    ExtraHrs = ExtraHrs + 0.5
                    If CDate(!InTime) > #2:05:00 PM# Then
                        ExtraDeds = 0 'DateDiff("n", "14:30", CDate(!InTime)) / 60
                    ElseIf CDate(!InTime) < #2:00:00 PM# Then
                        ExtraDeds = DateDiff("n", CDate(!InTime), "14:00") / 60
                    End If
                    If DateDiff("n", "14:00", CDate(!InTime)) > 0 And DateDiff("n", "14:00", CDate(!InTime)) < 6 Then
                        ExtraHrs = ExtraHrs + DateDiff("n", "14:00", CDate(!InTime)) / 60
                    End If
                Else
                    If DateDiff("n", strEmpLunchOut, CDate(!InTime)) > 0 And DateDiff("n", strEmpLunchOut, CDate(!InTime)) < 6 Then '"13:30"
                        ExtraHrs = DateDiff("n", strEmpLunchOut, CDate(!InTime)) / 60 '"13:30"
                    End If
                    
                    If CDate(!InTime) > CDate(DateAdd("n", 5, strEmpLunchOut)) Then  '#1:35:00 PM#
                        ExtraDeds = 0 'DateDiff("n", "14:00", CDate(!InTime)) / 60
                    ElseIf !InTime < strEmpLunchOut Then '#1:30:00 PM#
                        ExtraDeds = DateDiff("n", CDate(!InTime), strEmpLunchOut) / 60 '"13:30"
                    End If
                End If
                
                If CurrTime > strEmpOutTime Then '#5:00:00 PM#
                    ExtraDeds = ExtraDeds + 0 'DateDiff("n", "17:00", CurrTime) / 60
                End If
                
                If CurrTime > strEmpOutTime Then '#5:00:00 PM#
                    Hrs = DateDiff("n", CDate(!InTime), strEmpOutTime) '#5:00:00 PM#
                Else
                    Hrs = DateDiff("n", CDate(!InTime), CurrTime)
                End If
            End If
            'Check ... Is This Out Time Entry ?
            
            If TotalEntries < 3 Then
                If !OutTime & "" = "" Then
                    'Hrs = DateDiff("n", CDate(!InTime), CDate(CurrTime))
                
                    'Check for the Minimum Time Difference
                    If Hrs < TimeDiff Then GoTo END_OF_PROCEDURE
                
                    Hrs = Round(Hrs / 60, 2)
                
                    'Update Out Time and Do Nothing Else
                
                    Sql = "UPDATE EmpTimes SET OutTime='" & CurrTime & "',Hrs=" & Hrs & ",PayableHrs=" & (Hrs + ExtraHrs - ExtraDeds) & " WHERE EntryID=" & Val(!EntryID & "") & ""
                    con.Execute Sql
                    'con.Execute "Update AttendanceSheet Set Attendance=1 Where EmpID='" & EmpId & "' AND DT='" & AttDate & "'"
                    GoTo END_OF_PROCEDURE
                
                Else
            
                    'Check for the Minimum Time Difference
                
                    Hrs = DateDiff("n", CDate(!OutTime), CDate(CurrTime))
                    If TotalEntries <> 2 Then
                        If Hrs < TimeDiff Then GoTo END_OF_PROCEDURE
                    End If
                    
                
                End If
            End If
        Else
            If bDeductLHrs And DateDiff("n", strEmpInTime, CurrTime) > 10 Then '"08:00"
                dLateHrs = DateDiff("n", strEmpInTime, CurrTime) / 60 '"08:00"
            Else
                dLateHrs = 0
            End If
            AttRecs = Val(con.Execute("Select Count(*) From AttendanceSheet Where EmpID='" & EmpId & "' AND DT='" & AttDate & "'").Fields(0).Value & "")
            If AttRecs <> 0 Then con.Execute "Delete From AttendanceSheet Where EmpID='" & EmpId & "' AND DT='" & AttDate & "'"
            
            con.Execute "INSERT INTO AttendanceSheet(EmpID,DT,Attendance,LateHours) VALUES('" & EmpId & "','" & AttDate & "',1," & _
             dLateHrs & ")", a
            
        End If

        '.Close
        
        If Emptype = 2 Then
            
            ' If No Previous Entry Found Accept The Entry
            If TotalEntries = 0 Then GoTo ACCEPT_ENTRY
        
            'Remember The Total Entries
        
            ' If this is Second Time Entry Accept It
            If TotalEntries < 2 Then GoTo ACCEPT_ENTRY
            
        End If
        ' If Two Entries Have Been Made ... Check For Over Time Authorization
        If TotalEntries >= 2 Then
            Dim rsOT As New ADODB.Recordset
            Dim bAuthorized As Boolean
            With rsOT
                .Open "Select * From VOverTimeAuth Where EmpID='" & EmpId & "' and DT='" & AttDate & "'", con, adOpenForwardOnly, adLockReadOnly
                ' If Not Authorized for Over Time Neglect This Entry
                If .EOF Then
                    bAuthorized = False
                Else
                    bAuthorized = True
                End If
                rsOT.Close
                Set rsOT = Nothing
                If bAuthorized = False Then
                    'MsgBox "Not Authorized.", vbInformation
                    GoTo END_OF_PROCEDURE
                End If
            End With
            'If Val(![Authority] & "") <> 1 Then GoTo END_OF_PROCEDURE
            
            ' Other Wise Accept The Entry
            If !OutTime & "" <> "" Then
                con.Execute "INSERT INTO EmpTimes(EmpID,DT,InTime,ManualEntry,OverTime) VALUES('" & EmpId _
                 & "','" & AttDate & "','" & CurrTime & "',1,1)"
            Else
                Hrs = DateDiff("n", CDate(!InTime), CurrTime)
                If Hrs < TimeDiff Then GoTo END_OF_PROCEDURE
                Hrs = Round(Hrs / 60, 2)
                con.Execute "UPDATE EmpTimes SET OutTime='" & CurrTime & "',Hrs=" & Hrs & ",PayableHrs=" & (Hrs) & " WHERE EntryID=" & Val(!EntryID & "") & ""
            End If
            GoTo END_OF_PROCEDURE
            bOverTime = True
        End If
    End With


    ' Enter The New Entry With In Time

ACCEPT_ENTRY:

    Sql = "INSERT INTO EmpTimes(EmpID,DT,InTime) VALUES('" & EmpId _
     & "','" & AttDate & "','" & CurrTime & "')"
        
    con.Execute Sql
'    TotalEntries = Val(con.Execute("Select TotalEntries From VEmpDailyEntries Where EmpID='" & EmpID & "' and DT='" & AttDate & "'").Fields(0).Value & "")
'
'    If TotalEntries = 1 Then    'That Means that it's first entry of the day. So Calculate LateHours...
'        If DateDiff("n", "08:00", CurrTime) > 15 Then
'            con.Execute "Update AttendanceSheet Set LateHours=" & DateDiff("n", "08:00", CurrTime) / 60 & " Where EmpID='" & EmpID & "' and DT='" & AttDate & "'"
'        End If
'    End If
    
END_OF_PROCEDURE:

    If Rec.State = 1 Then Rec.Close
    Set Rec = Nothing
    TakeAttOld = True
    Exit Function
err:
    Debug.Print err.Description
    'So That Next Entries Should Not Be Skipped
    If CatchError(err.Description) = 0 Then
        MsgBox "Server Connection Not Found.", vbCritical
        End
    Else
        Resume Next
    End If
    

End Function

Private Function CatchError(strErr As String) As Integer

    Dim imyAff As Integer
    On Error Resume Next
    con.Execute "Insert Into Errors(ErrDescription,DT) Values('" & strErr & "','" & Now & "')", i
    CatchError = imyAff
    
End Function


Private Sub cmdImportFromTxts_Click()
    
    Dim rs As New ADODB.Recordset
    Dim strMachines() As String, iIndex As Integer
    With rs
        .Open "SELECT * FROM AttendanceMachines_txtFilePaths ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        ReDim strMachines(.RecordCount, 2)
        iIndex = 0
        Do Until .EOF
            strMachines(iIndex, 0) = !txtFilePath & ""
            strMachines(iIndex, 1) = Val(!MachineType & "")
            iIndex = iIndex + 1
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Dim j As Integer, iType As Integer
    For j = 0 To UBound(strMachines) - 1
    
        File1.Path = strMachines(j, 0)
        File1.Refresh
        iType = Val(strMachines(j, 1))
        
        Dim i As Integer
        Dim DTFile As Date
        Dim iCount As Long
        For i = 0 To File1.ListCount - 1
            If iType = 1 Then
                iCount = iCount + ReadFromTextFile(File1.Path & "\" & File1.List(i), File1.List(i), File1.Path)
            ElseIf iType = 2 Then
                iCount = iCount + ReadFromTextFileType2(File1.Path & "\" & File1.List(i), File1.List(i), File1.Path)
            End If
        Next
       
    Next
    
    Dim iTransferredToDB As Integer
    
    iTransferredToDB = TransferToDatabase(True)
    MsgBox "Read " & iCount & " Records & Transferred " & iTransferredToDB & " Records.", vbInformation
End Sub

Private Function ReadFromTextFile(strFile As String, strFileName As String, strPath As String) As Integer
    
    If strFile = "" Then Exit Function
    '---------------------------
    On Error GoTo 0
    Dim fs As Variant
    Set fs = CreateObject("Scripting.FileSystemObject")
    Dim i As Integer
    intEmpFileNbr = FreeFile
    
    'FileCopy strFile, "E:\Attendance Backup\" & Format(Now, "dd-MMM HHnn") & ".bak"
    
    Open strFile For Input As #intEmpFileNbr
    
    Dim strtemp As String
    Dim myDT As Date, myTime As String, strEmpID As String, strTemp1 As String
    Dim strDT As String, iInstrIndex As Integer
    Dim iYear As Integer, iMonth As Integer, iDay As Integer
    Dim iCount As Integer
    iCount = 0
    Dim imyPos As Integer
    Do Until EOF(intEmpFileNbr)
        
        Line Input #intEmpFileNbr, strtemp   'Acc : Fixed
        If strCompany = "Tecno" Then
            
            imyPos = InStr(strtemp, "kq,")
            If imyPos = 0 Then
                GoTo GoToNextLine
            Else
                'Data...
            End If
            'For Tecno
            '(22603754,kq,1473,2010-7-16 8:6:7,0)
            strTemp1 = Mid(strtemp, imyPos + 3)
            imyPos = InStr(strTemp1, ",")
            
            strEmpID = Left(strTemp1, imyPos - 1)
            
            strTemp1 = Mid(strTemp1, imyPos + 1)
            imyPos = InStr(strTemp1, ",")
            strTemp1 = Left(strTemp1, imyPos - 1)
            myDT = CDate(strTemp1)
            
            myTime = Format(myDT, "HH:nn")
            myDT = Format(myDT, "dd-MMM-yyyy")
            
            strEmpID = "TCN-" & Format(strEmpID, "00000")
        ElseIf strCompany = "Banzai" Then
            strEmpID = Mid$(strtemp, 4, 5)
            strEmpID = "BNZ-" & strEmpID
            
            strTemp1 = Mid$(strtemp, 10, 10)
            myDT = CDate(strTemp1)
            
            strTemp1 = Mid$(strtemp, 21, 5)
            myTime = Format(strTemp1, "HH:nn")
            
        End If
        iCount = iCount + 1
        
        On Error Resume Next
        con.Execute "INSERT INTO txtFilesData(EmpID,AttDT,AttTime) VALUES('" & strEmpID & _
         "','" & myDT & "','" & Format(myTime, "HH:nn") & "')"
GoToNextLine:
        
    Loop
    
    Close #intEmpFileNbr
    On Error Resume Next
    FileCopy strFile, strPath & "\Backup\" & Left(strFileName, Len(strFileName) - 4) & "(" & Format(Now, "dd-MMM-yyyy HHnn") & ").bak"
    On Error GoTo 0
    Kill strFile
    ReadFromTextFile = iCount
    
End Function

Private Function ReadFromTextFileType2(strFile As String, strFileName As String, strPath As String) As Integer
    
    If strFile = "" Then Exit Function
    '---------------------------
    On Error GoTo 0
    Dim fs As Variant
    Set fs = CreateObject("Scripting.FileSystemObject")
    Dim i As Integer
    intEmpFileNbr = FreeFile
     
    Open strFile For Input As #intEmpFileNbr
    
    Dim strtemp As String
    Dim myDT As Date, myTime As String, strEmpID As String, strTemp1 As String
    Dim strDT As String, iInstrIndex As Integer
    Dim iYear As Integer, iMonth As Integer, iDay As Integer
    Dim iCount As Integer
    iCount = 0
    Dim imyPos As Integer
    Do Until EOF(intEmpFileNbr)
        
        Line Input #intEmpFileNbr, strtemp
        
        strEmpID = Mid$(strtemp, 5, 6)
        strEmpID = "BNZ-" & Format(strEmpID, "00000")
        
        strTemp1 = Mid$(strtemp, 16, 10)
        myDT = CDate(strTemp1)
        
        strTemp1 = Mid$(strtemp, 28, 5)
        myTime = Format(strTemp1, "HH:nn")
        
        iCount = iCount + 1
        
        On Error Resume Next
        con.Execute "INSERT INTO txtFilesData(EmpID,AttDT,AttTime) VALUES('" & strEmpID & _
         "','" & myDT & "','" & Format(myTime, "HH:nn") & "')"
GoToNextLine:
        
    Loop
    
    Close #intEmpFileNbr
    On Error Resume Next
    FileCopy strFile, strPath & "\Backup\" & Left(strFileName, Len(strFileName) - 4) & "(" & Format(Now, "dd-MMM-yyyy HHnn") & ").bak"
    On Error GoTo 0
    Kill strFile
    ReadFromTextFileType2 = iCount
    
End Function


Private Function TransferToDatabase(Optional bFromTxtFiles As Boolean = False) As Integer

    On Error GoTo err
    
    Dim iTransfer As Integer
    Dim bTransferred As Boolean
    
    Dim strDataArray() As String
    Dim i As Long
    i = 0
     
    Dim rs As New ADODB.Recordset
    With rs
        If bFromTxtFiles Then
            .Open "SELECT * FROM txtFilesData WHERE Processed=0 AND Delay=0 ORDER By EmpID,AttDT,AttTime", con, adOpenForwardOnly, adLockReadOnly
        Else
            .Open "SELECT * FROM BioMetricsData WHERE Processed=0 AND Delay=0 ORDER BY EmployeeID,AttDT,AttTime", con, adOpenForwardOnly, adLockReadOnly
        End If
        
        ReDim strDataArray(.RecordCount, 3)
        Do Until .EOF
            If bFromTxtFiles Then
                strDataArray(i, 0) = !EmpId
            Else
                strDataArray(i, 0) = !EmployeeID
            End If
            strDataArray(i, 1) = !AttDT
            strDataArray(i, 2) = !AttTime
            strDataArray(i, 3) = !EntryID
            i = i + 1
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
     
    Dim strEmployeeID As String, AttDT As Date, AttTime As Date, lEntryID As Long
     
    For i = 0 To UBound(strDataArray) - 1
        strEmployeeID = strDataArray(i, 0)
        On Error Resume Next
        AttDT = strDataArray(i, 1)
        AttTime = strDataArray(i, 2)
        lEntryID = strDataArray(i, 3)
        
        If err.Number = 13 Then     'Type Mismatch
            con.Execute "UPDATE txtFilesData SET Delay=1 WHERE EntryID=" & lEntryID
            err.Clear
        Else
            
        
            If DateLocked(AttDT, True, True) = False Then
                bTransferred = TakeAtt(strEmployeeID, AttDT, AttTime, lEntryID, True, True)
                
                If bTransferred Then
                    iTransfer = iTransfer + 1
                Else
                    con.Execute "UPDATE txtFilesData SET Delay=1 WHERE EntryID=" & lEntryID
                End If
            Else
                con.Execute "UPDATE txtFilesData SET Delay=1 WHERE EntryID=" & lEntryID
            End If
        End If
        
        
    Next

    TransferToDatabase = iTransfer
    
    Exit Function
err:
    MsgBox err.Description
    
End Function



