VERSION 5.00
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vBalIml6.ocx"
Object = "{FEBA0457-F711-4B34-8C59-F9C1746D3AE6}#1.0#0"; "FK524PXN.ocx"
Begin VB.Form frmEmpLedger_BioMetrics 
   Caption         =   "Employee Ledger..."
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11565
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleMode       =   0  'User
   ScaleWidth      =   10358.93
   WindowState     =   2  'Maximized
   Begin FK524PXNLib.FK524PXN FK524PXN1 
      Left            =   8685
      Top             =   360
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
      CommPort        =   1
      ReadMark        =   -1  'True
      UsingModem      =   0   'False
      TelNumber       =   ""
      waitDialTime    =   0
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   90
      TabIndex        =   12
      Top             =   885
      Width           =   2355
      _ExtentX        =   4154
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
      Left            =   5325
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker DT 
      Height          =   300
      Left            =   5505
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
      Format          =   141295619
      CurrentDate     =   37384
   End
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   30
      TabIndex        =   3
      Top             =   0
      Width           =   11490
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
         Left            =   4410
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
         Left            =   4425
         TabIndex        =   5
         Top             =   135
         Width           =   2685
      End
   End
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   2430
      TabIndex        =   2
      Top             =   885
      Width           =   3075
      _ExtentX        =   5424
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
      Height          =   6045
      Left            =   60
      TabIndex        =   0
      Top             =   1200
      Width           =   11385
      Begin zkemkeeperCtl.CZKEM CZKEM1 
         Height          =   1350
         Left            =   2340
         OleObjectBlob   =   "frmEmpLedger_BioMetrics.frx":0000
         TabIndex        =   21
         Top             =   3045
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.FileListBox File1 
         Height          =   1260
         Left            =   0
         Pattern         =   "*.txt"
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   2340
      End
      Begin vbAcceleratorSGrid6.vbalGrid EmpGrid 
         Height          =   5385
         Left            =   60
         TabIndex        =   13
         Top             =   150
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   9499
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
         Left            =   345
         TabIndex        =   6
         Top             =   5505
         Width           =   10905
         Begin MSForms.CommandButton cmdImportFromTxts 
            Height          =   375
            Left            =   0
            TabIndex        =   19
            Top             =   90
            Width           =   2265
            Caption         =   "Import from Text files"
            PicturePosition =   327683
            Size            =   "3995;661"
            Accelerator     =   73
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdLateComers 
            Height          =   375
            Left            =   2010
            TabIndex        =   18
            Top             =   90
            Visible         =   0   'False
            Width           =   1425
            ForeColor       =   0
            Caption         =   "Late Comers"
            PicturePosition =   327683
            Size            =   "2514;661"
            Accelerator     =   68
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
            Left            =   6045
            TabIndex        =   15
            Top             =   90
            Visible         =   0   'False
            Width           =   1905
            ForeColor       =   0
            Caption         =   "Absent Leave Detail"
            PicturePosition =   327683
            Size            =   "3360;661"
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
            Left            =   4995
            TabIndex        =   14
            Top             =   90
            Visible         =   0   'False
            Width           =   1425
            ForeColor       =   0
            Caption         =   "A / L Summary"
            PicturePosition =   327683
            Size            =   "2514;661"
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
            Left            =   7995
            TabIndex        =   8
            Top             =   90
            Width           =   1425
            ForeColor       =   0
            Caption         =   "Print         "
            PicturePosition =   327683
            Size            =   "2514;661"
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
            Left            =   9495
            TabIndex        =   7
            Top             =   90
            Width           =   1425
            ForeColor       =   0
            Caption         =   " Close      "
            PicturePosition =   327683
            Size            =   "2514;661"
            Accelerator     =   67
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdACorrection 
            Height          =   375
            Left            =   6495
            TabIndex        =   16
            Top             =   90
            Visible         =   0   'False
            Width           =   1425
            ForeColor       =   0
            Caption         =   "A Correction"
            PicturePosition =   327683
            Size            =   "2514;661"
            Accelerator     =   65
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
      Left            =   10500
      Top             =   675
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   4592
      Images          =   "frmEmpLedger_BioMetrics.frx":0024
      Version         =   131072
      KeyCount        =   4
      Keys            =   "ÿÿÿ"
   End
   Begin MSForms.ComboBox cmbMachines 
      Height          =   345
      Left            =   10560
      TabIndex        =   17
      Top             =   840
      Width           =   1875
      DisplayStyle    =   7
      Size            =   "3307;609"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdRefreshRFID 
      Height          =   360
      Left            =   9000
      TabIndex        =   11
      Top             =   840
      Width           =   1515
      Caption         =   "Read From"
      PicturePosition =   327683
      Size            =   "2672;635"
      Accelerator     =   82
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
      Caption         =   " Department                            Employee                                                  Date"
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
      Width           =   7125
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   7305
      TabIndex        =   1
      Top             =   840
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
   Begin VB.Menu mnuLateComers 
      Caption         =   "EditLateComers"
      Visible         =   0   'False
      Begin VB.Menu mnuAddComment 
         Caption         =   "Add Comments         "
      End
   End
End
Attribute VB_Name = "frmEmpLedger_BioMetrics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalPresents As Double, TotalAbsents As Double, TotalLeaves As Double
Dim SQL As String
Dim FirstRefresh As Boolean

Private Sub cmbdept_matched()
    Call RefreshLedger
End Sub

Private Sub cmbEmp_matched()
    Call RefreshLedger
    Exit Sub
End Sub

Private Sub cmdACorrection_Click()
    Call PrintActivitySheet(True)
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
            FormulaField.Text = "'" & company & "'"
        ElseIf FormulaField.Name = "{@ForDate}" Then
            FormulaField.Text = "'" & Format(DT, "dddd dd-MMM-yyyy") & "'"
        End If
    
    Next

    rpt.SQLQueryString = "Select * from VEmp Where EmpID In (Select EmpID From AttendanceSheet Where Attendance=1 and DT='" & _
     Format(DT, "dd-MMM-yyyy") & "')"
    
    f.ShowReport " ", rpt

    Screen.MousePointer = vbDefault

End Sub

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
'            If iType = 1 Then
'                iCount = iCount + ReadFromTextFile(File1.Path & "\" & File1.List(i), File1.List(i), File1.Path)
'            ElseIf iType = 2 Then
'                iCount = iCount + ReadFromTextFileType2(File1.Path & "\" & File1.List(i), File1.List(i), File1.Path)
'            End If
            iCount = iCount + ReadFromTextFileWrangler(File1.Path & "\" & File1.List(i), File1.List(i), File1.Path)
            
        Next
       
    Next
    
    Dim iTransferredToDB As Integer
    
    iTransferredToDB = TransferToDatabase(True)
    MsgBox "Read " & iCount & " Records & Transferred " & iTransferredToDB & " Records.", vbInformation
    
End Sub

Private Sub cmdLateComers_Click()
    Call PrintDailyAbsentLeave(False, True)
End Sub

Private Sub cmdPrint_Click()
    Call cmdAttVerification_Click
    Exit Sub
    'Call PrintActivitySheet
End Sub

Private Sub PrintActivitySheet(Optional bCorrection As Boolean = False)
    
    On Error GoTo err
    Screen.MousePointer = vbHourglass
    Dim Status As Integer
    Dim strSelectionFormula As String
    
    Call StartTrans(con)
    con.Execute "DELETE FROM PrintActivitySheet"

    With EmpGrid
    
        For i = 1 To .Rows
            Status = IIf(.CellText(i, 11) <> "", 1, 0)
            
            con.Execute "INSERT INTO PrintActivitySheet(EmpID,DT,Name,InTime,OutTime,Hrs," & _
            "LateInTime,ShortTime,ExcessTime,Status,AttendanceGate,Attendance) Values('" & _
            IIf(.CellText(i, 12) = "", .CellText(i, 2), .CellText(i, 12)) & "','" & _
            Format(DT, "dd-MMM-yyyy") & "','" & .CellText(i, 3) & "','" & .CellText(i, 4) & "','" & .CellText(i, 5) & "','" & _
            .CellText(i, 6) & "','" & .CellText(i, 7) & "','" & .CellText(i, 8) & "','" & _
            .CellText(i, 9) & "'," & Status & ",'" & .CellText(i, 13) & "'," & Val(.CellText(i, 14)) & ")"
            
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
            FormulaField.Text = "'" & company & "'"
        ElseIf FormulaField.Name = "{@FromTo}" Then
            FormulaField.Text = "'" & Format(DT, "dd-MMM-yyyy") & "'"
        End If
    Next
    
    If bCorrection Then
        strSelectionFormula = "{PrintActivitySheet.InTime}<>'' AND {PrintActivitySheet.OutTime}=''"
    End If
    f.ShowReport strSelectionFormula, rpt

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
    Dim strTemp As String
    strSQL = "Insert Into PrintAttComparison Values(" & SortID & ",'" & TimeDesc & "',"
    
    For i = 1 To 31
        strTemp = strTemp & DayAtt(i) & ","
    Next i
    strSQL = strSQL & Left(strTemp, Len(strTemp) - 1) & ")"
    
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
            FormulaField.Text = "'" & company & "'"
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

Private Sub PrintDailyAbsentLeave(Optional bSummary As Boolean = False, Optional bLateComers As Boolean = False)

    On Error GoTo err
    Me.MousePointer = vbHourglass
    Dim cmd As New ADODB.Command
    With cmd
        .CommandType = adCmdStoredProc
        If bSummary Then
            .CommandText = "SP_PrintDailyAbsentLeaveSummary"
        Else
            .CommandText = "SP_PrintDailyAbsentLeaveDetail"
        End If
        Set .ActiveConnection = con
        .Parameters("@DT").Type = adDBTimeStamp
        .Parameters("@DT").Value = DT.Value
        .Execute
    End With
    Dim cmd1 As New ADODB.Command
    If bSummary = False Then
        With cmd1
            .CommandType = adCmdStoredProc
            .CommandText = "PrintEmpLateComingAndAbsentList_SP"
            Dim DTFromM As Date, DTToM As Date, DTFromY As Date, DTToY As Date
            DTFromM = DateSerial(DT.year, DT.Month, 1)
            DTToM = DateSerial(DT.year, DT.Month + 1, 0)
            If DT.Month > 6 Then
                DTFromY = DateSerial(DT.year, 7, 1)
                DTToY = DateSerial(DT.year + 1, 6, 30)
            Else
                DTFromY = DateSerial(DT.year - 1, 7, 1)
                DTToY = DateSerial(DT.year, 6, 30)
            End If
            Set .ActiveConnection = con
            .Parameters("@DTFromM").Value = DTFromM
            .Parameters("@DTToM").Value = DTToM
            .Parameters("@DTFromY").Value = DTFromY
            .Parameters("@DTToY").Value = DTToY
            .Execute
        End With
    End If
     
    Dim rpt As CRAXDDRT.Report, f As New frmPrevRpt, strSelFor As String
    If bLateComers Then
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\LateComers.rpt")
        rpt.FormulaFields(2).Text = "'" & Format(DT.Value, "dd-MMM-yyyy") & "'"
        strSelFor = "{PrintDailyAbsentLeaveDetail.LateComing}=TRUE"
    ElseIf bSummary Then
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptAbsentLeaveSummary.rpt")
        rpt.FormulaFields(2).Text = "'" & Format(DT.Value, "dd-MMM-yyyy") & "'"
    Else
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptAbsentLeaveDetail.rpt")
        strSelFor = "{PrintDailyAbsentLeaveDetail.Present}=False"
        rpt.FormulaFields(2).Text = "'" & Format(DT.Value, "dd-MMM-yyyy") & "'"
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
    
    On Error GoTo err
    
    If cmbMachines.ListIndex = -1 Then
        MsgBox "Please Select Machine.", vbInformation
        Exit Sub
    End If
    'Open Machine.....
    Dim lPortNo As Long, lPWD As Long
    lPortNo = 5005
    lPWD = 200490
    Dim lpszIPAddress As String
    lpszIPAddress = cmbMachines.Text
    
    Dim lMachineNo As Long, iMachineType As Integer
    lMachineNo = CLng(cmbMachines.List(cmbMachines.ListIndex, 1))
    iMachineType = CInt(cmbMachines.List(cmbMachines.ListIndex, 2))
    
    If iMachineType = 0 Then
        FK524PXN1.SetIPAddress lpszIPAddress, lPortNo, lPWD
        If FK524PXN1.OpenCommPort(lMachineNo, 3) = True Then
            Call ReadAllData_FK(lMachineNo)
            FK524PXN1.CloseCommPort
        End If
    ElseIf iMachineType = 1 Then
'        FP_CLOCK1.SetIPAddress lpszIPAddress, lPortNo, lPWD
'        If FP_CLOCK1.OpenCommPort(lMachineNo) = True Then
'            Call ReadAllData(lMachineNo)
'            FP_CLOCK1.CloseCommPort
'        End If
    ElseIf iMachineType = 2 Then    'Electro Machine Type..
        
        If BFLex1.SetIPAddress(lpszIPAddress) = False Then
            MsgBox "Please Check IP Address.", vbInformation
            Exit Sub
        End If
        
        If BFLex1.OpenNet() Then
            Call ReadAllData_Type2(lMachineNo)
        End If
    ElseIf iMachineType = 3 Then  'IFace
        lPortNo = 4370
        lPWD = 200490
        CZKEM1.ReadGeneralLogData 1
        If CZKEM1.Connect_Net(cmbMachines.Text, lPortNo) = 1 Then
            Call ReadAllData_IFace(lMachineNo)
            CZKEM1.Disconnect
        End If
    End If
     
    Exit Sub
err:
    MsgBox "Error Occured in Refresh Module." & vbNewLine & err.Description
    
End Sub

Private Sub ReadAllData_FK(mMachineNumber As Long)
    
    Dim vSEnrollNumber As Long
    Dim vVerifyMode As Long
    Dim vInOutMode As Long
    Dim vYear As Long
    Dim vMonth As Long
    Dim vDay As Long
    Dim vHour As Long
    Dim vMinute As Long
    Dim vErrorCode As Long
    Dim vRet As Boolean
    Dim i, n As Long
    
    Dim strValue As String
    DoEvents
    Dim myDT As Date, myTime As Date, strEmpID As String
    gstrLogItem = Array("", "EnrollNo", "VerifyMode", "InOutMode", "DateTime")
     
    MousePointer = vbHourglass
    vRet = FK524PXN1.EnableDevice(mMachineNumber, False)
    If vRet = False Then
        'lblMessage.Caption = gstrNoDevice
        MsgBox "Device Not Found.", vbInformation
        MousePointer = vbDefault
        Exit Sub
    End If
    
    vRet = FK524PXN1.ReadAllGLogData(mMachineNumber)
    If vRet = False Then
        FK524PXN1.GetLastError vErrorCode
        'MsgBox "Errors Occured.", vbInformation
        'MsgBox ErrorPrint(vErrorCode), vbInformation
    End If
    
    If vRet = True Then
        'lblMessage.Caption = "Getting..."
        MousePointer = vbHourglass
        DoEvents
        Do
            vRet = FK524PXN1.GetAllGLogData(mMachineNumber, _
                                                 vSEnrollNumber, _
                                                 vVerifyMode, _
                                                 vInOutMode, _
                                                 vYear, _
                                                 vMonth, _
                                                 vDay, _
                                                 vHour, _
                                                 vMinute)
                If vRet = False Then Exit Do
                DoEvents
               
                Debug.Print vSEnrollNumber
                
'                If vVerifyMode = 1 Then
'                    .Text = "Fp"
'                ElseIf vVerifyMode = 2 Then
'                    .Text = "Password"
'                ElseIf vVerifyMode = 3 Then
'                    .Text = "Card"
'                ElseIf vVerifyMode = 4 Then
'                    .Text = "Fp+Pass"
'                ElseIf vVerifyMode = 5 Then
'                    .Text = "Fp+Card"
'                ElseIf vVerifyMode = 6 Then
'                    .Text = "Pass+Fp"
'                ElseIf vVerifyMode = 7 Then
'                    .Text = "Card+Fp"
'                ElseIf vVerifyMode = 12 Then
'                    .Text = "Alarm"
'                Else
'                    .Text = "--"
'                End If
'                .Col = 3
'                Select Case vInOutMode
'                Case 0
'                    .Text = vInOutMode & "--" & "AL"
'                Case 1
'                    .Text = vInOutMode & "--" & "IN"
'                Case 2
'                    .Text = vInOutMode & "--" & "OUT"
'                Case 3
'                    .Text = vInOutMode & "--" & "Start"
'                Case 4
'                    .Text = vInOutMode & "--" & "End"
'                End Select
'                .Col = 4
                If vVerifyMode <> 1 Then GoTo MoveNEXT:
                
                strValue = CStr(vYear) & "/" & Format(vMonth, "0#") & "/" & Format(vDay, "0#") & _
                        " " & Format(vHour, "0#") & ":" & Format(vMinute, "0#")
                
                'LabelTotal.Caption = "Total : " & i
                DoEvents
                'i = i + 1
                myDT = DateSerial(vYear, vMonth, vDay)
                myTime = Format(vHour, "00") & ":" & Format(vMinute, "00")
                
                If strCompany = "QEL" Then
                    strEmpID = "QEL-" & Format(vSEnrollNumber, "000000")
                ElseIf (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai" Or strCompany = "Towne" Then
                    
                    strEmpID = Format(vSEnrollNumber, "000000")
                    Dim strTemp As String
                    strTemp = GetSingleStringValue("LEFT(EmpID,2)", "Employees", " WHERE RIGHT(EmpID,6)='" & strEmpID & "'")
                    
                    If strTemp = "DF" Then
                        strEmpID = "DF-" & Format(vSEnrollNumber, "000000")
                    Else
                        strEmpID = "T-" & Format(vSEnrollNumber, "000000")
                    End If
                    
                End If
                
                'Call PrepareAttendance(strEmpID, myDT, myTime)
                Call TakeAtt(strEmpID, myDT, myTime)
MoveNEXT:
            Loop
            
        MsgBox "Data Read Successfully.", vbInformation
    End If
    
    FK524PXN1.EmptyGeneralLogData (mMachineNumber)
    FK524PXN1.EnableDevice mMachineNumber, True
    MousePointer = vbDefault
    
End Sub

Private Function TakeAtt(EmployeeID As String, AttDate As Date, AttTime As Date, Optional lEntryID As Long = 0, Optional p_bByEmpID As Boolean = False, Optional bFromTxtFiles As Boolean = False) As Boolean
    
    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim strEmpID As String, strDeptID As String, bExemptSettings As Boolean
    Dim bLunchTimeScanning As Boolean
    Dim strLunchInTime As String, strLunchOutTime As String, iEmpAttCatID As Integer
    Dim iTotalEntries As Integer, bActive As Boolean
    Dim bOverTimePaid As Boolean
    strEmpID = EmployeeID
    With rs
        .Open "Select DeptID,EmpID,ExemptSettings,Active From Employees Where EmpID='" & EmployeeID & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            strEmpID = !EmpId & ""
            strDeptID = !DeptID & ""
            bExemptSettings = !ExemptSettings
            bActive = !Active
        Else
            Exit Function
        End If
        .Close
        'Save a List of Machine's Attendance Data... Done for keeping track of Wrangler's Security Guards Attendance.
        con.Execute "INSERT INTO EmployeeTimings_All(EmpID,DT,Timing) VALUES('" & strEmpID & "','" & AttDate & "','" & Format(AttTime, "HH:nn") & "')"
        
        If bExemptSettings Then
            .Open "Select LunchInTime,LunchOutTime,EmpAttCat_RefID,FridayBreakStartTime,FridayBreakEndTime From EmpSettings Where EmpID='" & strEmpID & "'", con, adOpenForwardOnly, adLockReadOnly
        Else
            .Open "Select LunchInTime,LunchOutTime,EmpAttCat_RefID,FridayBreakStartTime,FridayBreakEndTime From GeneralSettings Where DeptID='" & strDeptID & "'", con, adOpenForwardOnly, adLockReadOnly
        End If
        
        If Format(AttDate, "ddd") = "Fri" Then
            strLunchInTime = !FridayBreakStartTime & ""
            strLunchOutTime = !FridayBreakEndTime & ""
        Else
            strLunchInTime = !LunchInTime & ""
            strLunchOutTime = !LunchOutTime & ""
        End If
        iEmpAttCatID = Val(!EmpAttCat_RefID & "")
        
        .Close
        
        .Open "Select LunchTimeScanning,OverTimePaid From EmployeeAttendanceCatagories Where EntryID=" & iEmpAttCatID, con, adOpenForwardOnly, adLockReadOnly
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
    
    If bLunchTimeScanning = False And strLunchInTime <> strLunchOutTime Then
        If iTotalEntries = 1 And AttTime > #2:00:00 PM# Then
'            If Format(AttDate, "ddd") = "Fri" Then
'                Call TakeAttendance(strEmpID, AttDate, #1:00:00 PM#, True)
'                DoEvents
'                Call TakeAttendance(strEmpID, AttDate, #2:00:00 PM#, True)
'                DoEvents
''                Call TakeAttendance(strEmpID, AttDate, CDate(strLunchInTime), True)
''                DoEvents
''                Call TakeAttendance(strEmpID, AttDate, CDate(strLunchOutTime), True)
''                DoEvents
'            Else
                Call TakeAttendance(strEmpID, AttDate, CDate(strLunchInTime), True)
                DoEvents
                Call TakeAttendance(strEmpID, AttDate, CDate(strLunchOutTime), True)
                DoEvents
'            End If
        End If
    End If
    TakeAtt = TakeAttendance(strEmpID, AttDate, AttTime, True)
    DoEvents
     
    Exit Function
err:
    MsgBox "Error occured in TakeAtt Module." & vbNewLine & err.Description
    
End Function

Private Sub TakeAtt_Old(EmpId As String, AttDate As Date, AttTime As Date)

    '''''''''''   This Will Is The Employee Exists Or Not    '''''''''''
    Dim strInTime As String
    Dim strOutTime As String

    If Val(con.Execute("SELECT COUNT(*) FROM Employees WHERE EmpID='" & EmpId & "' AND Active=1").Fields(0)) = 0 Then
        MsgBox "The Following Employee " & EmpId & " Does Not Exists !", vbInformation
        Exit Sub
    End If

    With con.Execute("SELECT DeptID,EmpID FROM Employees WHERE EmpID='" & EmpId & "'")
    
        If .EOF Then
            Exit Sub
        Else
            EmpId = ![EmpId] & ""
            
            DeptID = ![DeptID] & ""
        End If
        .Close
        
        .Open "SELECT InTime,OutTime FROM Departments WHERE DeptID='" & DeptID & "'", con, adOpenForwardOnly, adLockReadOnly
        
        strInTime = !InTime & ""
        strOutTime = !OutTime & ""
        
        .Close
    End With
    
    Dim Rec As New ADODB.Recordset, CurrTime As String, Hrs As Double, SQL As String, Cond As String
    Dim TimeDiff As Integer

    Cond = "WHERE EmpID='" & EmpId & "' AND DT='" & AttDate & "'"
    CurrTime = Format(AttTime, "HH:NN")

    TimeDiff = Val(con.Execute("SELECT TimeDiff FROM TimeSettings").Fields(0) & "")

    With Rec
        .Open "SELECT * FROM EmpTimes WHERE EntryID=(SELECT MAX(EntryID) FROM EmpTimes " & Cond & ")", con, adOpenStatic, adLockReadOnly
    
        SQL = "INSERT INTO EmpTimes(EmpID,DT,InTime) VALUES('" & EmpId _
         & "','" & AttDate & "','" & CurrTime & "')"
        If .RecordCount > 0 Then
            If !OutTime & "" = "" Then
                Hrs = DateDiff("n", CDate(!InTime), CDate(CurrTime))
                If Company_DBName <> "AKI" Then
                    Hrs = Hrs - (dTeaTime * 60) - (dLunchTime * 60)
                End If
                If Hrs < TimeDiff Then Exit Sub
                Hrs = Round(Hrs / 60, 2)
                SQL = "UPDATE EmpTimes SET OutTime='" & CurrTime & "',Hrs=" & Hrs & " WHERE EntryID=" & Val(!EntryID & "") & ""
            Else
                Hrs = DateDiff("n", CDate(!OutTime), CDate(CurrTime))
                If Company_DBName <> "AKI" Then
                    Hrs = Hrs - (dTeaTime * 60) - (dLunchTime * 60)
                End If
                If Hrs < TimeDiff Then Exit Sub
            End If
        Else
        
            If DateDiff("n", strInTime, CurrTime) > 5 Then
                dLateHrs = DateDiff("n", strInTime, CurrTime) / 60
            Else
                dLateHrs = 0
            End If
            con.Execute "DELETE FROM AttendanceSheet WHERE empid='" & EmpId & "' and Dt='" & AttDate & "'"
            con.Execute "INSERT INTO AttendanceSheet(EmpID,DT,Attendance,LateHours) VALUES('" & EmpId & "','" & AttDate & "',1," & dLateHrs & ")"
        
        End If

        con.Execute SQL
    
    End With

End Sub


Private Sub DT_Change()
   ' Call RefreshLedger
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
    cmbEmp.AddVals con, "Name + ' {' + EmpID + '}'", "Employees", "EmpID", " WHERE Active=1"
    
    cmbDept.ClearVals
    cmbDept.AddItem "< All Departments >", "0"
    cmbDept.AddVals con, "Name + ' {' + DeptID + '}'", "Departments", "DeptID"
    
    cmbDept.ID = "0"
    cmbEmp.ID = "0"
    
    
    Call AddToCombo(cmbMachines, "IPAddress,MachineNo,MachineType", "AttendanceMachine")
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
    
    FLV.Move 50, 1230, Me.ScaleWidth - 100, Me.ScaleHeight - 1400
         
    'LV.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
     
    EmpGrid.Move 100, 200, Me.Width - 400, FLV.Height - FButs1.Height - 300
          
    FButs1.Move EmpGrid.Left + EmpGrid.Width - FButs1.Width, EmpGrid.Top + EmpGrid.Height
     
    'FButs2.Move EmpGrid.Left, EmpGrid.Top + EmpGrid.Height
    
End Sub

Private Sub mnuAddComment_Click()
    frmLateComers.Show vbModal
End Sub

Private Sub tcpClient_Connect()

    'MainForm.Enabled = False
    frmWait.Show
    tcpClient.SendData "RefreshList"

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


    SQL = "SELECT Employees.*, EmpTimes.DT AS DT, EmpTimes.EntryID AS EntryID, EmpTimes.InTime AS InTime," & _
        " EmpTimes.OutTime AS OutTime,EmpTimes.Hrs,EmpTimes.LateHours as lateHours,EmpTimes.ExtraHrs as ExtraHours FROM (SELECT * FROM dbo.VEmptimes1 " & Cond & ") AS EmpTimes RIGHT OUTER JOIN" & _
        " dbo.VEmp as Employees ON EmpTimes.EmpID = Employees.EmpID Where Employees.Active=1"
       
    
    
    LV.ListItems.Clear
    LV.Visible = False
    
    rsDept.Open "Select * from VDeptSettings " & DeptCond & " Order By Sno", con, adOpenForwardOnly, adLockReadOnly

    With Rec
    
    .Open SQL, con, adOpenStatic, adLockReadOnly
    
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
        
        .AddColumn , "Short Time", , , 0, False
        .AddColumn , "Excess Time", , , 0, False
        
        .AddColumn , "Keys", , , , False
        
        .AddColumn , "Details", , , .Width, , , , , , True
        .AddColumn , "EmpIDKey", , , , False
        .AddColumn , "AttGate", , , , False
        .AddColumn , "Attendance", , , , False
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
        DeptCond = " WHERE DeptID='" & cmbDept.ID & "'"
    End If

    SQL = "SELECT Employees.DeptID,Employees.Name,Employees.EmpID,Employees.LunchTime, EmpTimes.DT AS DT, EmpTimes.EntryID AS EntryID, EmpTimes.InTime AS InTime," & _
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
    
        .Open SQL, con, adOpenStatic, adLockReadOnly
    
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

Private Sub EmpGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single, bDoDefault As Boolean)

    Dim lCol As Long, lRow As Long
    Dim lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long
    Dim lType As Long, lClause As Long
    Dim lIconIndex As Long
    Dim i As Long

    EmpGrid.CellFromPoint X \ Screen.TwipsPerPixelX, Y \ Screen.TwipsPerPixelY, lRow, lCol
    
    If lCol = 1 Then
        If EmpGrid.CellText(lRow, 11) <> "" Then
            Call SetExpand(lRow, EmpGrid.CellText(lRow, 10))
        End If
    End If
    
'    If Button = 2 And EmpGrid.CellText(lRow, 7) <> "" Then
'
'        'Me.PopUpMenu mnuLateComers
'    End If
    

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
            If (EmpGrid.CellText(i, 10) = ParentID) And i <> lRow Then
                EmpGrid.RowVisible(i) = True
            End If
        Next i
    Else
        ' Collapse
        lIconIndex = 3
        For i = lRow + 1 To EmpGrid.Rows
            'If (Left(EmpGrid.CellText(i, 10), 4) = ParentID) And i <> lRow Then
            If (EmpGrid.CellText(i, 10) = ParentID) And i <> lRow Then
                EmpGrid.RowVisible(i) = False
            Else
                Exit For
            End If
        Next i
    End If
    EmpGrid.CellExtraIcon(lRow, 1) = lIconIndex
    
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

Private Sub ReadAllData(mMachineNumber As Long)

    Dim vTMachineNumber As Long
    Dim vSMachineNumber As Long
    Dim vSEnrollNumber As Long
    Dim vVerifyMode As Long
    Dim vYear As Long
    Dim vMonth As Long
    Dim vDay As Long
    Dim vHour As Long
    Dim vMinute As Long
    Dim vErrorCode As Long
    Dim vRet As Boolean
    Dim i, n As Long
    Dim vMaxLogCnt  As Long
    
    Dim strValue As String
    DoEvents
    Dim myDT As Date, myTime As Date, strEmpID As String
    gstrLogItem = Array("", "TMachineNo", "EnrollNo", "EMachineNo", "VeriMode", "DateTime")
     
    MousePointer = vbHourglass
    vRet = FP_CLOCK1.EnableDevice(mMachineNumber, False)
    If vRet = False Then
        'lblMessage.Caption = gstrNoDevice
        MsgBox "Device Not Found.", vbInformation
        MousePointer = vbDefault
        Exit Sub
    End If
    
    vRet = FP_CLOCK1.ReadAllGLogData(mMachineNumber)
    If vRet = False Then
        FP_CLOCK1.GetLastError vErrorCode
        'MsgBox "Errors Occured.", vbInformation
        'MsgBox ErrorPrint(vErrorCode), vbInformation
    End If
    
    If vRet = True Then
        'lblMessage.Caption = "Getting..."
        FP_CLOCK1.EnableDevice mMachineNumber, True
        MousePointer = vbHourglass
        DoEvents
        Do
            vRet = FP_CLOCK1.GetAllGLogData(mMachineNumber, _
                                                 vTMachineNumber, _
                                                 vSEnrollNumber, _
                                                  vSMachineNumber, _
                                                 vVerifyMode, _
                                                 vYear, _
                                                 vMonth, _
                                                 vDay, _
                                                 vHour, _
                                                 vMinute)
                If vRet = False Then Exit Do
                DoEvents
               
                Debug.Print vSEnrollNumber
                
                
                strValue = CStr(vYear) & "/" & Format(vMonth, "0#") & "/" & Format(vDay, "0#") & _
                        " " & Format(vHour, "0#") & ":" & Format(vMinute, "0#")
                
                'LabelTotal.Caption = "Total : " & i
                DoEvents
                'i = i + 1
                myDT = DateSerial(vYear, vMonth, vDay)
                myTime = Format(vHour, "00") & ":" & Format(vMinute, "00")
                
                If strCompany = "QEL" Then
                    strEmpID = "QEL-" & Format(vSEnrollNumber, "000000")
                ElseIf strCompany = "IAA" Or strCompany = "Kami" Then
                    strEmpID = "Wrangler-" & Format(vSEnrollNumber, "00000")
                ElseIf (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai" Or strCompany = "Towne" Then
                    
                    strEmpID = Format(vSEnrollNumber, "000000")
                    Dim strTemp As String
                    strTemp = GetSingleStringValue("LEFT(EmpID,2)", "Employees", " WHERE RIGHT(EmpID,6)='" & strEmpID & "'")
                    
                    If strTemp = "DF" Then
                        strEmpID = "DF-" & Format(vSEnrollNumber, "000000")
                    Else
                        strEmpID = "T-" & Format(vSEnrollNumber, "000000")
                    End If
                    
                End If
                
                'Call PrepareAttendance(strEmpID, myDT, myTime)
                Call TakeAtt(strEmpID, myDT, myTime)
            Loop
            
        MsgBox "Data Read Successfully.", vbInformation
    End If
    
    FP_CLOCK1.EmptyGeneralLogData (mMachineNumber)
    FP_CLOCK1.EnableDevice mMachineNumber, True
    MousePointer = vbDefault
    
End Sub



Private Sub ReadAllData_Type2(mMachineNumber As Long)

    Dim i As Long, LogDate As Date, bRet As Boolean
    Dim lEnrollNumber As Long, lt As ListItem
    Dim ID As Long
    Dim Glogs As Long
    
    Dim myDT As Date, myTime As Date, strEmpID As String
    ID = 1
    Glogs = BFLex1.ReadAllGLogData(mMachineNumber, ID)
    
    For i = ID To Glogs + ID - 1
    
        bRet = BFLex1.GetGeneralLogData(mMachineNumber, i, lEnrollNumber, LogDate)
        If Not bRet Then
            Exit For
        End If
        myDT = Format(LogDate, "dd-MMM-yyyy")
        myTime = Format(LogDate, "HH:nn")
                
        If strCompany = "QEL" Then
            strEmpID = "QEL-" & Format(lEnrollNumber, "000000")
        ElseIf strCompany = "IAA" Or strCompany = "Kami" Then
            strEmpID = "Wrangler-" & Format(lEnrollNumber, "00000")
        ElseIf (strCompany = "Dr-Frgz" And strSubCompany = "Intl") Or strCompany = "Banzai" Then
            strEmpID = Format(vSEnrollNumber, "000000")
            Dim strTemp As String
            strTemp = GetSingleStringValue("LEFT(EmpID,2)", "Employees", " WHERE RIGHT(EmpID,6)='" & strEmpID & "'")
            If strTemp = "DF" Then
                strEmpID = "DF-" & Format(vSEnrollNumber, "000000")
            Else
                strEmpID = "T-" & Format(vSEnrollNumber, "000000")
            End If
        ElseIf (strCompany = "Dr-Frgz" And strSubCompany = "Electro") Then
            strEmpID = "DFG-" & Format(lEnrollNumber, "000000")
        ElseIf strCompany = "Towne" Then
            strEmpID = "TWN-" & Format(vSEnrollNumber, "000000")
        End If
        Call TakeAtt(strEmpID, myDT, myTime)
    Next
    
    Call BFLex1.DeleteGLogData(mMachineNumber, 0)
    BFLex1.CloseNet
    
End Sub


Private Function ReadFromTextFileWrangler(strFile As String, strFileName As String, strPath As String) As Integer
    
    If strFile = "" Then Exit Function
    '---------------------------
    On Error GoTo 0
    Dim fs As Variant
    Set fs = CreateObject("Scripting.FileSystemObject")
    Dim i As Integer
    intEmpFileNbr = FreeFile
    
    'FileCopy strFile, "E:\Attendance Backup\" & Format(Now, "dd-MMM HHnn") & ".bak"
    
    Open strFile For Input As #intEmpFileNbr
    
    Dim strTemp As String
    Dim myDT As Date, myTime As String, strEmpID As String, strTemp1 As String
    Dim strDT As String, iInstrIndex As Integer
    Dim iYear As Integer, iMonth As Integer, iDay As Integer
    Dim iCount As Integer
    iCount = 0
    Dim imyPos As Integer, imyPos2 As Integer
    Do Until EOF(intEmpFileNbr)
        
        Line Input #intEmpFileNbr, strTemp
        
        On Error Resume Next
        
        imyPos = InStr(strTemp, vbTab)
        strEmpID = Trim(Left(strTemp, imyPos - 1))
        imyPos2 = InStr(imyPos + 1, strTemp, vbTab)
        strDT = Mid$(strTemp, imyPos + 1, 16)
        'myDT = Trim(Mid$(strTemp, imyPos - 2, 20))
        myTime = Format(strDT, "HH:nn")
        myDT = Format(strDT, "dd-MMM-yyyy")
        'Now Remove 00s.
        strEmpID = "EMR-" & Format(strEmpID, "00000")
            
        iCount = iCount + 1
        
        
        con.Execute "INSERT INTO txtFilesData(EmpID,AttDT,AttTime) VALUES('" & strEmpID & _
         "','" & myDT & "','" & Format(myTime, "HH:nn") & "')"
GoToNextLine:
        
    Loop
    
    Close #intEmpFileNbr
    On Error Resume Next
    FileCopy strFile, strPath & "\Backup\" & Left(strFileName, Len(strFileName) - 4) & "(" & Format(Now, "dd-MMM-yyyy HHnn") & ").bak"
    On Error GoTo 0
    Kill strFile
    ReadFromTextFileWrangler = iCount
    
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

Private Sub ReadAllData_IFace(mMachineNumber As Long)

    Dim vTMachineNumber As Long
    Dim vSMachineNumber As Long
    Dim vSEnrollNumber As Long
    Dim vVerifyMode As Long
    Dim vYear As Long
    Dim vMonth As Long
    Dim vDay As Long
    Dim vHour As Long
    Dim vMinute As Long
    Dim vSecond As Long
    Dim vWorkerCode As Long
    Dim vErrorCode As Long
    Dim vRet As Boolean
    Dim i, n As Long
    Dim vMaxLogCnt  As Long
    Dim vInOutMode As Long
    Dim lParam4 As Long, lParam3 As Long, lParam2 As Long, lParam1 As Long, lMainuplate As Long
    Dim strValue As String, strTime As String, strEnrollCode As String
    DoEvents
    Dim myDT As Date, myTime As Date, strEmpID As String
    gstrLogItem = Array("", "TMachineNo", "EnrollNo", "EMachineNo", "VeriMode", "DateTime")
     
    MousePointer = vbHourglass
    
'    vRet = CZKEM1.EnableDevice(mMachineNumber, False)
'    If vRet = False Then
'        MsgBox "Device Not Found.", vbInformation
'        MousePointer = vbDefault
'        Exit Sub
'    End If
    
    vRet = CZKEM1.ReadGeneralLogData(mMachineNumber)

    If vRet = False Then
        CZKEM1.GetLastError vErrorCode
    End If
    
    If vRet = 1 Then
        'lblMessage.Caption = "Getting..."
        'CZKEM1.EnableDevice mMachineNumber, True
        MousePointer = vbHourglass
        DoEvents
        Do
            
            'vRet = CZKEM1.GetAllGLogData(mMachineNumber, vTMachineNumber, vSEnrollNumber, mMachineNumber, vVerifyMode, vInOutMode, vYear, vMonth, vDay, vHour, vMinute)
            vRet = CZKEM1.SSR_GetGeneralLogData(mMachineNumber, strEnrollCode, vVerifyMode, vInOutMode, vYear, vMonth, vDay, vHour, vMinute, vSecond, vWorkerCode)
            'vRet = CZKEM1.GetGeneralLogDataStr(mMachineNumber, vSEnrollNumber, vVerifyMode, vInOutMode, strTime)
            'vRet = CZKEM1.GetGeneralLogData(mMachineNumber,  vTMachineNumber, vSEnrollNumber, mMachineNumber,vSMachineNumber, vVerifyMode, vInOutMode, vYear, vMonth, vDay, vHour, vMinute)
                If vRet = 0 Then Exit Do
                DoEvents
               
'                Debug.Print strEnrollCode
                
                strValue = CStr(vYear) & "/" & Format(vMonth, "0#") & "/" & Format(vDay, "0#") & _
                        " " & Format(vHour, "0#") & ":" & Format(vMinute, "0#")
                
                'LabelTotal.Caption = "Total : " & i
                DoEvents
                'i = i + 1
                myDT = DateSerial(vYear, vMonth, vDay)
                myTime = Format(vHour, "00") & ":" & Format(vMinute, "00")
                
                vSEnrollNumber = Val(strEnrollCode)
                
                If myTime <= #5:00:00 AM# Then
                    myDT = DateAdd("d", -1, myDT)
                    bNextDay = True
                Else
                    bNextDay = False
                End If
                
                strEmpID = "EMR-" & Format(vSEnrollNumber, "00000")
                Call TakeAtt(strEmpID, myDT, myTime)
                'Call PrepareAttendance(strEmpID, myDT, myTime)
                'Call TakeAtt(strEmpID, myDT, myTime)
            Loop
            
        MsgBox "Data Read Successfully.", vbInformation
    End If
    CZKEM1.ClearGLog (mMachineNumber)
    
    CZKEM1.EnableDevice mMachineNumber, True
    MousePointer = vbDefault
    
End Sub


