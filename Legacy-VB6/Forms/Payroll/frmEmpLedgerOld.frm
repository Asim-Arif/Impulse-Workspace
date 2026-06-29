VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmEmpLedgerOld 
   BorderStyle     =   0  'None
   Caption         =   "Employee Ledger..."
   ClientHeight    =   7575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11550
   LinkTopic       =   "Form19"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleMode       =   0  'User
   ScaleWidth      =   10345.49
   ShowInTaskbar   =   0   'False
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   90
      TabIndex        =   19
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
      TabIndex        =   17
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
      Format          =   57868291
      CurrentDate     =   37384
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   345
      Top             =   1635
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpLedgerOld.frx":0000
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpLedgerOld.frx":0454
            Key             =   "ORD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpLedgerOld.frx":08AC
            Key             =   "PARIN"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpLedgerOld.frx":0D00
            Key             =   "ITMIN"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpLedgerOld.frx":1154
            Key             =   "CAT"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpLedgerOld.frx":15A8
            Key             =   "OK"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpLedgerOld.frx":1A00
            Key             =   "TYPE"
         EndProperty
      EndProperty
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
      Begin VB.Frame fButs2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   450
         Left            =   60
         TabIndex        =   9
         Top             =   5520
         Visible         =   0   'False
         Width           =   4410
         Begin VB.Label lblTotalPresents 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1395
            TabIndex        =   13
            Top             =   150
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total Presents :"
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
            Index           =   0
            Left            =   45
            TabIndex        =   12
            Top             =   165
            Width           =   1320
         End
         Begin VB.Label lblTotalAbs 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3645
            TabIndex        =   11
            Top             =   150
            Width           =   660
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Total Absents :"
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
            Left            =   2340
            TabIndex        =   10
            Top             =   165
            Width           =   1260
         End
      End
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   480
         Left            =   5850
         TabIndex        =   6
         Top             =   5505
         Width           =   5400
         Begin MSForms.CommandButton cmdPresentEmps 
            Height          =   375
            Left            =   90
            TabIndex        =   21
            Top             =   90
            Width           =   2295
            ForeColor       =   0
            Caption         =   "Print Present Emp. List"
            PicturePosition =   327683
            Size            =   "4048;661"
            Accelerator     =   80
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
            Left            =   2430
            TabIndex        =   8
            Top             =   105
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
         Begin MSForms.CommandButton cmdClose 
            Height          =   360
            Left            =   3930
            TabIndex        =   7
            Top             =   105
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
         Begin MSForms.CommandButton cmdPrintExcessTimes 
            Height          =   360
            Left            =   1740
            TabIndex        =   20
            Top             =   90
            Visible         =   0   'False
            Width           =   2295
            ForeColor       =   0
            Caption         =   "Print for Excess Times"
            PicturePosition =   327683
            Size            =   "4048;635"
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
      Begin Crystal.CrystalReport cr1 
         Left            =   1950
         Top             =   1020
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
         Height          =   5325
         Left            =   45
         TabIndex        =   14
         Top             =   165
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   9393
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Emp. ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   5997
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "In Time"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Out Time"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Net Time"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Late Coming"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Short Time"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "ExcessTime"
            Object.Width           =   2028
         EndProperty
      End
      Begin MSComctlLib.TreeView TV 
         Height          =   5325
         Left            =   45
         TabIndex        =   16
         Top             =   165
         Visible         =   0   'False
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   9393
         _Version        =   393217
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         BorderStyle     =   1
         Appearance      =   0
      End
   End
   Begin MSForms.CommandButton cmdRefreshRFID 
      Height          =   360
      Left            =   7320
      TabIndex        =   18
      Top             =   840
      Width           =   2595
      Caption         =   "Refresh Data From RFIDs"
      PicturePosition =   327683
      Size            =   "4577;635"
      Picture         =   "frmEmpLedgerOld.frx":1E54
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
      TabIndex        =   15
      Top             =   630
      Width           =   7125
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   7305
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1500
      Caption         =   "Refresh        "
      PicturePosition =   327683
      Size            =   "2646;635"
      Picture         =   "frmEmpLedgerOld.frx":22A2
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
Attribute VB_Name = "frmEmpLedgerOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalPresents As Double, TotalAbsents As Double, TotalLeaves As Double
Dim Sql As String

Private Sub AddHeads()

    Dim Rec As New ADODB.Recordset, RecDept As New ADODB.Recordset
    
    TV.ImageList = ImageList1
    TV.Nodes.Clear
    
    
    RecDept.Open "SELECT DeptID,Name FROM Departments", con, adOpenStatic, adLockReadOnly
    For i = 1 To RecDept.RecordCount
        TV.Nodes.Add , , RecDept![DeptID] & "'", RecDept![Name] & "", "TYPE"
        With Rec
            .Open "Select EmpID,Name from Employees WHERE DeptID='" & RecDept!DeptID & "'" & " and Active=1 ORDER BY EmpID", con, adOpenForwardOnly, adLockReadOnly
            For j = 1 To .RecordCount
                TV.Nodes.Add RecDept![DeptID] & "'", tvwChild, ![EmpId] & "'", ![Name] & "", "CAT"
                .MoveNext
            Next
            .Close
        
        End With
        RecDept.MoveNext
    Next

End Sub

Private Sub cmbdept_matched()
    Call RefreshLedger
End Sub

Private Sub cmbEmp_matched()

    Call RefreshLedger

Exit Sub

    If cmbEmp.ID = "0" Then
        For i = 1 To TV.Nodes.Count
            TV.Nodes(i).Checked = True
        Next
    Else
        For i = 1 To TV.Nodes.Count
            If TV.Nodes(i).key = cmbEmp.ID & "'" Then
                TV.Nodes(i).Checked = True
                TV.Nodes(i).EnsureVisible
            Else
                TV.Nodes(i).Checked = False
            End If
        Next
    
    End If

    Call RefreshLedger

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPresentEmps_Click()


Screen.MousePointer = vbHourglass

Dim f As New frmPrevRpt, rpt As New rptPresentEmpList

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

Private Sub cmdPrint_Click()

Screen.MousePointer = vbHourglass
Dim Status As Integer


con.Execute "DELETE FROM PrintActivitySheet"
For i = 1 To LV.ListItems.Count
    
    With LV.ListItems(i)
        Status = IIf(.ListSubItems(2).Bold And Len(.Text) > 2, 0, 1)
        con.Execute "INSERT INTO PrintActivitySheet VALUES('" & .Tag & "','" & .SubItems(1) _
            & "','" & .SubItems(2) & "','" & .SubItems(3) & "','" & .SubItems(4) & _
            "','" & .SubItems(5) & "','" & .SubItems(6) & "','" & .SubItems(7) & "','" & Status & "')"
    End With
Next


Dim f As New frmPrevRpt, rpt As New rptEmpActivitySheet

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

Private Sub cmdPrintExcessTimes_Click()

Screen.MousePointer = vbHourglass
Dim Status As Integer


con.Execute "DELETE FROM PrintActivitySheet"
For i = 1 To LV.ListItems.Count
    
    With LV.ListItems(i)
        If .SubItems(7) <> "" Or Len(.Text) = 2 Then
            Status = IIf(.ListSubItems(2).Bold And Len(.Text) > 2, 0, 1)
            con.Execute "INSERT INTO PrintActivitySheet VALUES('" & .Tag & "','" & .SubItems(1) _
            & "','" & .SubItems(2) & "','" & .SubItems(3) & "','" & .SubItems(4) & _
            "','" & .SubItems(5) & "','" & .SubItems(6) & "','" & .SubItems(7) & "','" & Status & "')"
        End If
    End With
Next


Dim f As New frmPrevRpt, rpt As New rptEmpActivitySheet

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

Public Sub cmdRefresh_Click()

Call RefreshLedger
    
    'Call Form_Resize
End Sub

Private Sub cmdRefreshRFID_Click()

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
    'Call RefreshLedger
   
    Screen.MousePointer = vbDefault
    

End Sub

Private Sub DT_Change()
    Call RefreshLedger
End Sub

Private Sub Form_Load()
    
    Call AddHeads

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


On Error Resume Next

 fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2

 FLV.Move 50, 1230, Me.ScaleWidth - 100, Me.ScaleHeight - 1400
 
 TV.Move 100, 200, TV.Width, FLV.Height - fButs1.Height - 300
 
 ''''''''''' Commented To Hide Tree View '''''''''''''
 'LV.Move TV.Left + TV.Width - 10, 200, Me.Width - 400 - TV.Width, TV.Height
 
 LV.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
 
 fButs1.Move LV.Left + LV.Width - fButs1.Width, LV.Top + LV.Height
 
 FButs2.Move LV.Left, LV.Top + LV.Height
 
End Sub


Private Sub tcpClient_Connect()

'MainForm.Enabled = False
frmWait.Show
tcpClient.SendData "RefreshList"


End Sub

Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)
    
    Dim str As String
    tcpClient.GetData str
    If str = "Completed" Then
        Unload frmWait
        'MainForm.Enabled = True
        Call cmdRefresh_Click
    End If
    
End Sub

Private Sub TV_NodeCheck(ByVal Node As MSComctlLib.Node)
    
    If Node.Parent Is Nothing Then
        For i = Node.Index + 1 To TV.Nodes.Count
            If TV.Nodes(i).Parent Is Nothing Then Exit For
            TV.Nodes(i).Checked = Node.Checked
            
        Next i
    End If
    RefreshLedger
    
End Sub

Private Sub RefreshLedger()
    
    
    Dim StrText As String
    
    
    For i = 1 To TV.Nodes.Count
        With TV.Nodes(i)
            .Bold = False
            .ForeColor = vbbalck
        End With
    Next
    
    

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
            Set ITM = LV.ListItems.Add(, , !DeptID & "")
            ITM.ForeColor = Color
            ITM.Bold = True
            ITM.Tag = ![DeptID]
            ITM.ListSubItems.Add(, , ![Name]).Bold = True
            ITM.ListSubItems(1).ForeColor = Color
            ITM.ListSubItems.Add(, , Format(!EnterTime & "", "HH:NN AMPM")).Bold = True
            ITM.ListSubItems(2).ForeColor = Color
            ITM.ListSubItems.Add(, , Format(!ExitTime & "", "HH:NN AMPM")).Bold = True
            ITM.ListSubItems(3).ForeColor = Color
            
            ITM.ListSubItems.Add(, , getHrsMin(![WorkingHrs])).Bold = True
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
                        ITM.ListSubItems.Add
                        strtxt = MinsToTime(TotalWorkedMins - TotalWorkingMins)
                        ITM.ListSubItems.Add(, , IIf(strtxt = "00:00", "", strtxt)).ForeColor = ITM.ForeColor
                    Else
                        strtxt = MinsToTime(TotalWorkingMins - TotalWorkedMins)
                        ITM.ListSubItems.Add(, , IIf(strtxt = "00:00", "", strtxt)).ForeColor = ITM.ForeColor
                    End If
                    
                End If
                
               FirstInTime = ![InTime] & ""
                
                Color = rsDept![Color]
                
                Set ITM = LV.ListItems.Add(, , !EmpId & "")
                ITM.Tag = ![EmpId]
                ITM.ListSubItems.Add(, , !Name & "").ForeColor = Color
                
                If Not IsNull(![InTime]) Then
                    LateInTime = DateDiff("n", CDate(rsDept![EnterTime]), CDate(![InTime]))
                Else
                    LateInTime = 0
                End If
                
                PrevEmp = !EmpId & ""
                
    '            With TV.Nodes(!EmpId & "'")
    '                .Bold = True
    '                .ForeColor = RGB(100, 100, 200)
    '                .Parent.ForeColor = RGB(100, 100, 200)
    '                .Parent.Bold = True
    '            End With
                LastOutTime = ![OutTime] & ""
                LastItmIndex = ITM.Index
            Else
                Set ITM = LV.ListItems.Add(, , "")
                ITM.Tag = PrevEmp
                ITM.ListSubItems.Add , , ""
                LateInTime = 0
                LastOutTime = ![OutTime] & ""
                LastItmIndex = ITM.Index
            End If
            
            ITM.ForeColor = Color
            ITM.ListSubItems(1).ForeColor = Color
            
            ITM.ListSubItems.Add(, , Format(!InTime & "", "HH:NN AMPM")).ForeColor = Color
            ITM.ListSubItems.Add(, , Format(!OutTime & "", "HH:NN AMPM")).ForeColor = Color
            
            ITM.ListSubItems.Add(, , getHrsMin(Val(!Hrs & ""))).ForeColor = Color
            
            
            If LateInTime <= 0 Then
                ITM.ListSubItems.Add(, , "").ForeColor = Color
            Else
                ITM.ListSubItems.Add(, , MinsToTime(Val(LateInTime))).ForeColor = Color
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
            
            For k = 1 To ITM.ListSubItems.Count
                ITM.ListSubItems(k).Bold = Bold
                ITM.ListSubItems(k).ForeColor = Color
            Next
    
            .MoveNext
        Next
        .Filter = ""
        rsDept.MoveNext
    Next dpt
    .Close
    LV.Visible = True
    End With


End Sub

Private Function GetTreeValues() As String

    Dim str As String
    
    For i = 1 To TV.Nodes.Count
        If Len(TV.Nodes(i).key) <> 3 Then
            If TV.Nodes(i).Checked = True Then str = str & "'" & TV.Nodes(i).key & ","
        End If
    Next
    If str <> "" Then
        GetTreeValues = Left(str, Len(str) - 1)
    End If

End Function

