VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form SalaryHistory 
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11550
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   7395
   ScaleMode       =   0  'User
   ScaleWidth      =   10345.49
   WindowState     =   2  'Maximized
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   405
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   556
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   315
      Left            =   2340
      TabIndex        =   4
      Top             =   405
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   556
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1815
      Top             =   2280
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
            Picture         =   "SalaryHistory.frx":0000
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SalaryHistory.frx":0452
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin VB.Frame FLV 
      Height          =   6555
      Left            =   30
      TabIndex        =   0
      Top             =   735
      Width           =   11385
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
         Height          =   5835
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   10292
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
            Text            =   "EmpID"
            Object.Width           =   1614
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Employee"
            Object.Width           =   6004
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   2088
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Description"
            Object.Width           =   6112
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Salary"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   570
         Left            =   7890
         TabIndex        =   6
         Top             =   5895
         Width           =   3015
         Begin MSForms.CommandButton cmdClose 
            Height          =   360
            Left            =   1530
            TabIndex        =   8
            Top             =   195
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
            Left            =   60
            TabIndex        =   7
            Top             =   195
            Width           =   1425
            ForeColor       =   0
            Caption         =   "Print         "
            PicturePosition =   327683
            Size            =   "2514;635"
            Picture         =   "SalaryHistory.frx":0564
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
      Begin MSForms.CommandButton cmdNewST 
         Height          =   360
         Left            =   45
         TabIndex        =   5
         Top             =   6075
         Width           =   2700
         ForeColor       =   4194304
         Caption         =   "Change Salary                 "
         PicturePosition =   327683
         Size            =   "4762;635"
         Accelerator     =   78
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin MSComCtl2.DTPicker DtTo 
      Height          =   300
      Left            =   8475
      TabIndex        =   9
      Top             =   405
      Width           =   1425
      _ExtentX        =   2514
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
      Format          =   151453699
      CurrentDate     =   37081
   End
   Begin MSComCtl2.DTPicker DtFrom 
      Height          =   300
      Left            =   7095
      TabIndex        =   10
      Top             =   405
      Width           =   1380
      _ExtentX        =   2434
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
      Format          =   151453699
      CurrentDate     =   37055
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"SalaryHistory.frx":0676
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
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   165
      Width           =   9750
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   9915
      TabIndex        =   2
      Top             =   375
      Width           =   1500
      Caption         =   "Refresh        "
      PicturePosition =   327683
      Size            =   "2646;635"
      Picture         =   "SalaryHistory.frx":0716
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
Attribute VB_Name = "SalaryHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload Me
End Sub



Private Sub cmdNewST_Click()
    With IncreaseSal
        .Show 1
        If cmbEmp.ID <> "" Then Call cmdRefresh_Click
    End With
End Sub

Private Sub cmdPrint_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    With cmd
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_PrintEmpAbsentList"
        .Parameters("@DTFrom").Value = DtFrom.Value
        .Parameters("@DTTo").Value = DtTo.Value
        .Execute
    End With
    
    Cond = "{VSalaryLedger.DT}=#" & DtFrom & "# TO #" & DtTo & "#"
    
    If cmbDept.ID = "0" Then
        If cmbEmp.ID <> "0" Then Cond = Cond & " AND {VSalaryLedger.EmpID}='" & cmbEmp.ID & "'"
    Else
        If cmbEmp.ID = "0" Then
            Cond = Cond & "AND {VSalaryLedger.DeptID}='" & cmbDept.ID & "'"
        Else
        
            Cond = Cond & " AND  {VSalaryLedger.Empid}='" & cmbEmp.ID & "'"
        End If
    End If
    
    Cond = Cond & " AND {VSalaryLedger.TempDept}=False AND {VSalaryLedger.Active}=True"
    Screen.MousePointer = vbHourglass

    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpSalHistory.rpt")
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & strCompany & "'"
        ElseIf FormulaField.Name = "{@From}" Then
            FormulaField.Text = "'" & Format(DtFrom, "ddd dd-mmm-yyyy") & "'"
        ElseIf FormulaField.Name = "{@To}" Then
            FormulaField.Text = "'" & Format(DtTo, "ddd dd-mmm-yyyy") & "'"
        End If
    Next

    f.ShowReport Cond & " ", rpt

    Screen.MousePointer = vbDefault

End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLedger
    
    'Call Form_Resize
End Sub

Private Sub Form_Load()

    DtTo = GetServerDate
    DtFrom = DateAdd("M", -1, GetServerDate)
    DtTo.MaxDate = GetServerDate
    DtTo.MaxDate = GetServerDate
    
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    
    
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID"
    cmbDept.Text = "<All Departments>"
    
    'Set LV.SmallIcons = ImageList1
    
End Sub
Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    cmbEmp.AddItem "<All Employees>", "0"
    
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "Employees", "EmpID", "Where Active=1"
    Else
        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' and active=1"
    End If
    cmbEmp.Text = "<All Employees>"
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Dim TopMargin As Integer
    
    'fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
    TopMargin = cmbEmp.Top + cmbEmp.Height + 30
    
    FLV.Move 50, TopMargin, Me.ScaleWidth - 100, Me.ScaleHeight - TopMargin
    
    LV.Move 100, 200, Me.Width - 400, FLV.Height - FButs1.Height - 300
    
    FButs1.Move LV.Left + LV.Width - FButs1.Width, LV.Top + LV.Height
    
    cmdNewST.Move LV.Left, LV.Top + LV.Height

 
End Sub



Private Sub RefreshLedger()

    
    Dim rsLedger As New ADODB.Recordset
    Dim ITM As ListItem
    Dim Cond As String
    Dim SQL As String
    Dim ICO As String
    
    
    
    If Not cmbEmp.MatchFound Then
        MsgBox "Select An Employee First", vbInformation
        Exit Sub
    Else
    
        Cond = "WHERE DT BETWEEN '" & DtFrom.Value & "' AND '" & DtTo.Value & "'"
        If cmbDept.ID = "0" Then
            If cmbEmp.ID <> "0" Then Cond = Cond & " AND EmpID='" & cmbEmp.ID & "'"
        Else
            If cmbEmp.ID = "0" Then
                Cond = Cond & "AND DeptID='" & cmbDept.ID & "'"
            Else
            
                Cond = Cond & " AND  empid='" & cmbEmp.ID & "'"
            End If
        End If
    End If
    Cond = Cond & " AND TempDept=0 AND Active=1"
    SQL = "select * from VSalaryLedger " & Cond & " Order By DT Asc"

    With rsLedger
        
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            
            
            Set ITM = LV.ListItems.add(, , !EmpId & "")
            ITM.Tag = !EmpId & ""
            ITM.SubItems(1) = !Name & " " & !Rel & " " & !FName
            ITM.SubItems(2) = Format(![DT] & "", "dd-MMM-yyyy")
            ITM.SubItems(3) = ![Description] & ""
            ITM.SubItems(4) = Val(![Salary] & "")
            ITM.SubItems(5) = Val(![SalaryExt] & "")
            
            .MoveNEXT
        Next
        .Close
    End With
    Set rsLedger = Nothing
    SQL = ""
    Cond = ""
    ICO = ""
End Sub




