VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmPLS 
   BorderStyle     =   0  'None
   ClientHeight    =   6585
   ClientLeft      =   105
   ClientTop       =   0
   ClientWidth     =   10845
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkZero 
      Caption         =   "Show Accounts With Zero Effect"
      Height          =   195
      Left            =   4740
      TabIndex        =   5
      Top             =   6255
      Value           =   1  'Checked
      Width           =   3165
   End
   Begin VB.PictureBox PicStatus 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   45
      ScaleHeight     =   420
      ScaleWidth      =   2700
      TabIndex        =   11
      Top             =   6090
      Visible         =   0   'False
      Width           =   2700
      Begin MSComctlLib.ProgressBar PBar 
         Height          =   240
         Left            =   60
         TabIndex        =   12
         Top             =   180
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Calculating Profit && Loss ..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   13
         Top             =   -30
         Width           =   1950
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   4125
      ScaleHeight     =   405
      ScaleWidth      =   6195
      TabIndex        =   8
      Top             =   45
      Width           =   6195
      Begin VB.CheckBox chkTax 
         Caption         =   "Taxable"
         Height          =   195
         Left            =   3870
         TabIndex        =   2
         Top             =   105
         Width           =   1005
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   2415
         TabIndex        =   1
         Top             =   45
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   16646147
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   615
         TabIndex        =   0
         Top             =   45
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   16646147
         CurrentDate     =   37055
      End
      Begin MSForms.CommandButton cmdRefresh 
         Default         =   -1  'True
         Height          =   360
         Left            =   4950
         TabIndex        =   3
         Top             =   30
         Width           =   1200
         Caption         =   "Refresh"
         PicturePosition =   327683
         Size            =   "2117;635"
         Picture         =   "frmPLS.frx":0000
         Accelerator     =   82
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "To :"
         Height          =   195
         Index           =   1
         Left            =   2085
         TabIndex        =   10
         Top             =   90
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "From :"
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   9
         Top             =   90
         Width           =   525
      End
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   705
      Top             =   2220
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Trial Balance"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1695
      Top             =   3330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPLS.frx":044E
            Key             =   "ACC"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPLS.frx":08A2
            Key             =   "PARIN"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPLS.frx":09B6
            Key             =   "PAR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPLS.frx":0ACA
            Key             =   "ACCIN"
         EndProperty
      EndProperty
   End
   Begin CRVIEWERLibCtl.CRViewer CRV 
      Height          =   6015
      Left            =   0
      TabIndex        =   4
      Top             =   15
      Width           =   10425
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   0   'False
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   0   'False
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   360
      Left            =   7950
      TabIndex        =   6
      Top             =   6120
      Width           =   1425
      ForeColor       =   64
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
   Begin MSForms.CommandButton cmdClose 
      Height          =   360
      Left            =   9420
      TabIndex        =   7
      Top             =   6105
      Width           =   1335
      Caption         =   "Close       "
      PicturePosition =   327683
      Size            =   "2355;635"
      Picture         =   "frmPLS.frx":0F1E
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
Attribute VB_Name = "frmPLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim rec As New ADODB.Recordset
Dim rsAccounts As New ADODB.Recordset
Dim i As Integer
Dim Bal As Double
Dim AccList As String
Dim SubAccList As String
Dim TotalDebit As Double, TotalCredit As Double
Dim Loaded As Boolean
Public Summary  As Boolean
Public Standered As Boolean

Dim PermissionDenied As Boolean

Dim rpt As craxddrt.Report


Private Sub RefreshList()

    Dim TimeTaken As Double
    TimeTaken = Timer
    
    Screen.MousePointer = vbHourglass
    On Error GoTo err
    Dim i As Integer
    Dim Balance As Double

    Dim ACCTypes(0 To 5) As Integer

    ACCTypes(0) = 31
    ACCTypes(1) = 41
    ACCTypes(2) = 33
    ACCTypes(3) = 42
    ACCTypes(4) = 43
    ACCTypes(5) = 44


    'screen.mousepointer = vbHourglass

    'Con.BeginTrans
    con.Execute "Delete From PLS"
    PBar.Max = con.Execute("Select Count(*) From Accounts Where Type in('31','33','41','42','43','44') ").Fields(0) + 2
    PBar.value = 0
   
    PicStatus.Visible = True
    PicStatus.Refresh
   
    For l = 0 To 5
        With rsAccounts
            If .State = 1 Then .Close
            .CursorLocation = adUseClient
            .Open "Select * from Accounts Where Type='" & ACCTypes(l) & "'", con, adOpenForwardOnly
            'PBar.Max = PBar.Max + .RecordCount
            For i = 0 To .RecordCount - 1
                AccList = ""
                Balance = CalculateBalance(.Fields("AccNo"))
                PBar.value = PBar.value + 1
                If Balance = 0 And chkZero = vbUnchecked Then GoTo EndOfLoop
                Dim AccountTitle As String
                AccountTitle = GetlvlStr(![AccNo]) & ![AccTitle] & ""
                con.Execute "INSERT INTO PLS(AccTitle,Type,Amount,Parent) values ('" & AccountTitle & "'," & ![Type] & "," & Balance & "," & Abs(![Parent]) & ")"
                'Call AddToList(i, .Fields("SubaccOf") & "", .Fields("AccNo"), .Fields("AccTitle"), .Fields("OpeningBalance"), TotalDebit, TotalCredit, Balance, .Fields("AccNo"), ![Parent], True)
EndOfLoop:
                .MoveNext
            Next i
            If .RecordCount = 0 Then con.Execute "INSERT INTO PLS(AccTitle,Type,Amount) values (''," & ACCTypes(l) & ",0)"
        End With


    Next

    'Con.CommitTrans


    DtFrom.Tag = DtFrom
    DtTo.Tag = DtTo

    If Summary Then
        Me.Caption = "Profit & Loss (Summary)"
        Call ShowSummaryReport
    ElseIf Standered Then
        Me.Caption = "Profit & Loss (Standered)"
        Call ShowStanderedReport
    Else
        Me.Caption = "Profit & Loss (Detail)"
        Call ShowDetailReport
    End If
   
   
    PBar.value = PBar.Max
   
    'For i = 1 To Rpt.FormulaFields.Count
    '  Debug.Print Rpt.FormulaFields(i).Name, i
    'Next
    PicStatus.Visible = False
    CRV.ToolTipText = Me.Caption
    Screen.MousePointer = vbDefault
    'MsgBox "Time Taken:" & Timer - TimeTaken, vbInformation
    Exit Sub
err:
    If err.Number = 380 Then Resume Next Else MsgBox "ERROR :" & err.Number & vbNewLine & err.Description
    Screen.MousePointer = vbDefault
    
End Sub


Private Sub chkDetail_Click()
   Call cmdRefresh_Click
End Sub

Private Sub chkTax_Click()
    'On Error Resume Next
    If rpt Is Nothing Then Exit Sub
    rpt.FormulaFields(8).Text = chkTax.value = vbChecked
    CRV.Refresh
End Sub

Private Sub chkZero_Click()
Call cmdRefresh_Click
End Sub

Private Sub cmdClose_Click()
Set rsAccounts = Nothing
Unload Me
End Sub

Private Sub cmdPrint_Click()
On Error GoTo err
 rpt.PrintOut True
 
 Exit Sub
err:
   MsgBox "Error :" & err.Number & vbNewLine & err.Description
End Sub

Private Sub cmdRefresh_Click()
   Screen.MousePointer = vbHourglass
   Call RefreshList
   Screen.MousePointer = vbDefault
End Sub
Private Sub DtFrom_Change()
       Call setDateRange(DtFrom)
   DtTo.MinDate = DtFrom
   
End Sub


Private Sub DtFrom_LostFocus()
'If CDate(DtFrom.Tag) <> CDate(DtFrom) Then Call Form_Load
End Sub

Private Sub DtTo_Change()
    Call setDateRange(DtTo)
End Sub

Private Sub Form_Activate()
'Check For User Access
If PermissionDenied Then
   Me.Hide
   MsgBox "You Donot Have Permissions To Use This Option. Please Contact Your Administrator", vbInformation
   Unload Me
   Exit Sub
End If


If Not Loaded Then
   'Call RefreshList
   Loaded = True
End If
End Sub

Private Sub Form_Load()
'PermissionDenied = Not UserHasAccess("ProfitLoss")
Screen.MousePointer = vbHourglass

   

DtFrom = DateAdd("M", -1, Date)
DtTo = Date
Me.KeyPreview = True
Loaded = False
Screen.MousePointer = vbDefault
End Sub

Private Sub CreateLists(PRec As ADODB.Recordset, AccNo As String)
'If Len(AccList) > 0 Then AccList = AccList & "'" & AccNo & "',"
AccList = AccList & "'" & AccNo & "',"
SubAccList = ""
Dim i As Integer
With PRec
    For i = 0 To .RecordCount - 1
        AccList = AccList & "'" & .Fields("AccNo") & "',"
        SubAccList = SubAccList & "'" & .Fields("AccNo") & "',"
        .MoveNext
    Next i
    
End With
If Right$(SubAccList, 1) = "," Then SubAccList = Left$(SubAccList, Len(SubAccList) - 1)
End Sub

Private Function CalculateBalance(AccNo As String) As Double

Dim i As Integer
With rec

    If .State = 1 Then .Close
    .CursorLocation = adUseClient
    .Open "Select AccNo from Accounts Where SubAccOf='" & AccNo & "'", con, adOpenForwardOnly, adLockReadOnly
    'Do While .RecordCount > 0
        'Call CreateLists(rec, AccNo)
    '    If .State = 1 Then .Close
    '.CursorLocation = adUseClient
    
    '.Open "Select AccNo from Accounts Where trim(Accno) Like '" & AccNo & "%')", con, adOpenForwardOnly, adLockReadOnly
    'Loop
    
    'If Len(AccList) > 0 Then
    '    AccList = Left$(AccList, Len(AccList) - 1)
    'Else
    '    AccList = "'" & AccNo & "'"
    'End If
    
        
    
    If .State = 1 Then .Close
      .CursorLocation = adUseClient
      .Open "Select -Sum(Debit)+Sum(Credit) as Total from Vouchers Where (AccNo Like '" & AccNo & "%') and (Credit is not null and Debit Is Not null)" & _
      " AND (VDate Between '" & DtFrom & "' AND '" & DtTo & "')", con, adOpenForwardOnly, adLockReadOnly
    
    CalculateBalance = Val(.Fields(0) & "")
    
    .Close
End With

End Function

Private Sub Form_Resize()
On Error Resume Next
   CRV.Move 50, 50, Me.ScaleWidth - 100, Me.ScaleHeight - (cmdClose.Height + 100)
   Picture1.Move CRV.Left + CRV.Width - Picture1.Width, CRV.Top + 25
   cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, CRV.Top + CRV.Height + 50
   cmdPrint.Move cmdClose.Left - cmdPrint.Width - 50, cmdClose.Top
   PicStatus.Move CRV.Left, CRV.Top + CRV.Height
   chkZero.Move cmdPrint.Left - chkZero.Width - 50, cmdClose.Top + 50
End Sub

Private Sub Form_Unload(cancel As Integer)
   Set rsAccounts = Nothing
   
End Sub
Private Sub ShowSummaryReport()

   Screen.MousePointer = vbHourglass
   Set rpt = rptApp.OpenReport(rptProductionPath & "\PlsSummary.rpt")
   CRV.ReportSource = rpt
   
   'Rpt.Database.Tables(1).Location = DBName
   rpt.Database.Verify
   rpt.ReadRecords
 'Company Name
    rpt.FormulaFields(13).Text = "'" & strCompany & "'"
 'Date Range
    rpt.FormulaFields(12).Text = "'From " & Format(DtFrom, "dd/MM/yyyy") & " To " & Format(DtTo, "dd/MM/yyyy") & "'"
 'Taxable Or Not
    rpt.FormulaFields(8).Text = chkTax.value = vbChecked
    
Screen.MousePointer = vbDefault
On Error Resume Next
rpt.Database.Tables(1).ConnectBufferString = "DSN=DSN_Vision;;User ID=sa;;Database=Vision;;Password=visionasim"
On Error GoTo 0
CRV.ViewReport
CRV.Zoom 90

End Sub

Private Sub ShowDetailReport()

    Screen.MousePointer = vbHourglass
    Set rpt = rptApp.OpenReport(rptAccountsPath & "\PLSRpt.rpt")
    CRV.ReportSource = rpt
    Dim subRpt As craxddrt.Report

    'Rpt.Database.Tables(1).Location = DBName
    rpt.ReadRecords

   Set subRpt = rpt.OpenSubreport("Sales")
   'subRpt.Database.Tables(1).Location = DBName
   subRpt.ReadRecords
   
   Set subRpt = rpt.OpenSubreport("CostOfSales")
   'subRpt.Database.Tables(1).Location = DBName
   subRpt.ReadRecords
   
   Set subRpt = rpt.OpenSubreport("CostOfSales - 01")
   'subRpt.Database.Tables(1).Location = DBName
   subRpt.ReadRecords
   
   Set subRpt = rpt.OpenSubreport("AdminExp")
   'subRpt.Database.Tables(1).Location = DBName
   subRpt.ReadRecords
   
   Set subRpt = rpt.OpenSubreport("AdminExp - 01")
   'subRpt.Database.Tables(1).Location = DBName
   subRpt.ReadRecords
   
   Set subRpt = rpt.OpenSubreport("AdminExp - 02")
   'subRpt.Database.Tables(1).Location = DBName
   subRpt.ReadRecords
 'Company Name
    rpt.FormulaFields(13).Text = "'" & strCompany & "'"
 'Date Range
    rpt.FormulaFields(12).Text = "'From " & Format(DtFrom, "dd/MM/yyyy") & " To " & Format(DtTo, "dd/MM/yyyy") & "'"
 'Taxable Or Not
    rpt.FormulaFields(8).Text = chkTax.value = vbChecked
    
Screen.MousePointer = vbDefault
On Error Resume Next
rpt.Database.Tables(1).ConnectBufferString = "DSN=DSN_Vision;;User ID=sa;;Database=Vision;;Password=visionasim"
On Error GoTo 0
CRV.ViewReport
CRV.Zoom 90

End Sub
Private Sub ShowStanderedReport()
Screen.MousePointer = vbHourglass

Set rpt = rptApp.OpenReport(rptAccountsPath & "\PlsNotes.rpt")
CRV.ReportSource = rpt

Dim subRpt As craxddrt.Report
   
   'Rpt.Database.Tables(1).Location = DBName
   'rpt.Database.SetDataSource Con
   
   rpt.ReadRecords
   Set subRpt = rpt.OpenSubreport("Sales")
   'subRpt.Database.Tables(1).Location = DBName
   subRpt.ReadRecords
   
   Set subRpt = rpt.OpenSubreport("CostOfSales")
   'subRpt.Database.Tables(1).Location = DBName
   subRpt.ReadRecords
   
   Set subRpt = rpt.OpenSubreport("CostOfSales - 01")
   'subRpt.Database.Tables(1).Location = DBName
   subRpt.ReadRecords
   
   Set subRpt = rpt.OpenSubreport("AdminExp")
   'subRpt.Database.Tables(1).Location = DBName
   subRpt.ReadRecords
   
   Set subRpt = rpt.OpenSubreport("AdminExp - 01")
   'subRpt.Database.Tables(1).Location = DBName
   subRpt.ReadRecords
   
   Set subRpt = rpt.OpenSubreport("AdminExp - 02")
   'subRpt.Database.Tables(1).Location = DBName
   subRpt.ReadRecords
 'Company Name
    rpt.FormulaFields(13).Text = "'" & strCompany & "'"
 'Date Range
    rpt.FormulaFields(12).Text = "'From " & Format(DtFrom, "dd/MM/yyyy") & " To " & Format(DtTo, "dd/MM/yyyy") & "'"
 'Taxable Or Not
    rpt.FormulaFields(8).Text = chkTax.value = vbChecked
    
Screen.MousePointer = vbDefault
On Error Resume Next
    rpt.Database.Tables(1).ConnectBufferString = "DSN=DSN_Vision;;User ID=sa;;Database=Vision;;Password=visionasim"
    On Error GoTo 0
CRV.ViewReport
CRV.Zoom 90

End Sub

Private Function GetlvlStr(AccNo As String) As String
    GetlvlStr = Space(4 * SubStrCount(AccNo, "-"))
    If Len(GetlvlStr) > 4 Then GetlvlStr = GetlvlStr & "- "
End Function
