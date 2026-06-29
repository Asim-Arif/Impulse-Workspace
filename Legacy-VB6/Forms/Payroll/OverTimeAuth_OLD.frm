VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form OverTimeAuth 
   BorderStyle     =   0  'None
   ClientHeight    =   6360
   ClientLeft      =   15
   ClientTop       =   -90
   ClientWidth     =   10425
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   ScaleHeight     =   6360
   ScaleMode       =   0  'User
   ScaleWidth      =   9337.813
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker DTAttend 
      Height          =   285
      Left            =   5205
      TabIndex        =   10
      Top             =   345
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dddd dd-MMM-yyyy"
      Format          =   58785795
      CurrentDate     =   37055
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   930
      Top             =   2160
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
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   345
      Width           =   2430
      _ExtentX        =   4286
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
      Left            =   2670
      TabIndex        =   0
      Top             =   345
      Width           =   2520
      _ExtentX        =   4445
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
      Height          =   5550
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   9945
      Begin VB.Frame FButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   570
         Left            =   4470
         TabIndex        =   6
         Top             =   4950
         Width           =   4905
         Begin MSForms.CommandButton cmdPrint 
            CausesValidation=   0   'False
            Height          =   360
            Left            =   30
            TabIndex        =   11
            Top             =   135
            Width           =   1575
            ForeColor       =   0
            Caption         =   "Print Sheet   "
            PicturePosition =   327683
            Size            =   "2778;635"
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
            CausesValidation=   0   'False
            Height          =   360
            Left            =   3315
            TabIndex        =   8
            Top             =   135
            Width           =   1575
            ForeColor       =   0
            Caption         =   " Close"
            PicturePosition =   327683
            Size            =   "2778;635"
            Accelerator     =   67
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdSave 
            Height          =   360
            Left            =   1680
            TabIndex        =   7
            Top             =   135
            Width           =   1575
            ForeColor       =   0
            Caption         =   " Save Entry "
            PicturePosition =   327683
            Size            =   "2778;635"
            Picture         =   "OverTimeAuth.frx":0000
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
      Begin VB.Frame FButs2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   570
         Left            =   105
         TabIndex        =   3
         Top             =   4950
         Width           =   3000
         Begin MSForms.CommandButton cmdClearAll 
            Height          =   390
            Left            =   1470
            TabIndex        =   5
            Top             =   120
            Width           =   1485
            ForeColor       =   0
            Caption         =   "UnAuthorize All"
            PicturePosition =   327683
            Size            =   "2619;688"
            Accelerator     =   108
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdCheckAll 
            Height          =   360
            Left            =   0
            TabIndex        =   4
            Top             =   135
            Width           =   1425
            ForeColor       =   0
            Caption         =   "Authorize All"
            PicturePosition =   327683
            Size            =   "2514;635"
            Picture         =   "OverTimeAuth.frx":0112
            Accelerator     =   104
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   4740
         Left            =   120
         TabIndex        =   9
         Top             =   150
         Width           =   9705
         _ExtentX        =   17119
         _ExtentY        =   8361
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "EmpID"
            Object.Width           =   3969
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Employee Description"
            Object.Width           =   11483
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Status"
            Object.Width           =   4498
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7500
      Top             =   3120
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
            Picture         =   "OverTimeAuth.frx":026C
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OverTimeAuth.frx":06BE
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Department :                            Employee Name :                       Date :"
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
      Left            =   240
      TabIndex        =   12
      Top             =   90
      Width           =   7290
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete This Attendance"
      End
   End
End
Attribute VB_Name = "OverTimeAuth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Editing As Boolean
Dim forLeave As Boolean

Public Sub enterLateHours()
    Load LateHours
    LateHours.Tag = LV.SelectedItem
    LateHours.Show
End Sub

Private Sub cmbEmp_matched()

On Error Resume Next
If cmbEmp.ID = "0" Then Exit Sub
LV.ListItems(cmbEmp.ID).Selected = True
LV.ListItems(cmbEmp.ID).EnsureVisible

End Sub

Private Sub cmbShift_Change()
    Call FillList
End Sub

Private Sub cmdCheckAll_Click()
    
    
        For i = 1 To LV.ListItems.Count
            With LV.ListItems(i)
                .Checked = True
                .SubItems(2) = "OverTime"
            End With
        
        Next i
    
End Sub

Private Sub cmdClearAll_Click()

    For i = 1 To LV.ListItems.Count
        With LV.ListItems(i)
            .Checked = False
            .SubItems(2) = ""
        End With
    Next i

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub


Public Sub cmdRefresh_Click()
    'Call RefreshLedger
    
    'Call Form_Resize
End Sub


Private Sub cmdPrint_Click()

If LV.ListItems.Count = 0 Then Exit Sub

con.Execute "DELETE FROM PrintOverTime"

For i = 1 To LV.ListItems.Count
    With LV.ListItems(i)
        con.Execute "INSERT INTO PrintOverTime VALUES('" & .Text _
            & "','" & .ListSubItems(1) & "','" & .ListSubItems(2) & "')"
    End With
Next


'With cr1
'    .ReportFileName = RptPath & "\PrintOverTime.rpt"
'    .Connect = ConnectStr
'    .SelectionFormula = Cond
'
'    .Formulas(0) = "Company='" & company & "'"
'    .Formulas(1) = "ForAtt='" & Format(DTAttend, "dddd dd-mmm-yyyy") & "'"
'
'    .Action = 1
'    .PageZoomNext
'
'End With

Screen.MousePointer = vbHourglass

Dim f As New frmPrevRpt, rpt As New rptPrintOverTime

Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

Set FormulaFields = rpt.FormulaFields
For Each FormulaField In FormulaFields
    If FormulaField.Name = "{@Company}" Then
        FormulaField.Text = "'" & company & "'"
    ElseIf FormulaField.Name = "{@FromTo}" Then
        FormulaField.Text = "'" & Format(DTAttend, "dddd dd-mmm-yyyy") & "'"
    End If
Next

'SelFormula = "Year({Holidays.DT})=" & cmbYear
rpt.SQLQueryString = Sql

f.ShowReport Cond & " ", rpt

Screen.MousePointer = vbDefault


End Sub

Private Sub cmdSave_Click()

On Error GoTo ERR

Dim Sql As String

    For i = 1 To LV.ListItems.Count
        With LV.ListItems(i)
            If .Checked = True Then
                If .Tag = "" Then
                    con.Execute "INSERT INTO OverTimeAuth VALUES('" & .Text & "','" & DTAttend & "',1)"
                End If
            Else
                con.Execute "DELETE FROM OverTimeAuth WHERE EmpID='" & .Text & "' AND DT='" & DTAttend & "'"
            End If
        End With
    Next i
'    con.CommitTrans
Exit Sub

ERR:
MsgBox ERR.Description, vbInformation

End Sub



Private Sub DTAttend_Change()
    
    Call FillList
    
End Sub

Private Sub Form_Load()

    DTAttend = Format(GetServerDate, "dd-MMM-yyyy")
    
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1600
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID"
    cmbDept.ID = 0

    DTAttend.Enabled = UserHasAccess("AttDateChange")
    
End Sub
Private Sub cmbdept_matched()

    Call FillList
    
    cmbEmp.ClearVals
    cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals con, "Name + ' {' + EmpID + '}'", "Employees", "EmpID", "Where active=1"
    Else
        cmbEmp.AddVals con, "Name + ' {' + EmpID + '}'", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' and active=1 "
    End If
    cmbEmp.ID = "0"
    
End Sub





Private Sub FillList()

    
    If cmbDept.MatchFound = False Then Exit Sub
    
    LV.ListItems.Clear

    Dim rsList As New ADODB.Recordset, cmd As New ADODB.command
    Dim ITM As ListItem, Attend As String
    
    Set cmd.ActiveConnection = con
    cmd.CommandText = IIf(cmbDept.ID = 0, "SP_OverTimeDept", "SP_OverTime")
    cmd.CommandType = adCmdStoredProc
    
    cmd.Parameters("@AttDate").Type = adDBTimeStamp
    cmd.Parameters("@AttDate").value = DTAttend
    
    If cmbDept.ID <> 0 Then
        cmd.Parameters("@DeptID").Type = adVarChar
        cmd.Parameters("@DeptID").value = cmbDept.ID
    End If
    
    Set rsList = cmd.Execute
    
    With rsList
        For i = 1 To .RecordCount
            Set ITM = LV.ListItems.Add(, !EmpId, !EmpId & "")
            ITM.ListSubItems.Add , , !Name & " " & !Rel & " " & !FName
            ITM.ListSubItems.Add , , IIf(IsNull(!Authority), "", "OverTime")
            If ITM.ListSubItems(2) <> "" Then ITM.Checked = True: ITM.Tag = "Edit"
            .MoveNext
        Next
    .Close
    
    Exit Sub
    
    '''''''''' Will Check For Over Time'''''''''''''''
    .Open "SELECT * FROM VOverTimeAuth WHERE IsNull(DT,true) or DT='" & DTAttend & "' ORDER BY EmpID", con, adOpenStatic, adLockReadOnly
    For i = 1 To .RecordCount
        With LV.ListItems(i)
            If rsList!DT & "" <> "" Then
                .Checked = True
                .SubItems(2) = "OverTime"
            End If
        End With
        .MoveNext
    Next
    Set rsLedger = Nothing
    
End With

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

    Dim TopMargin As Integer
    TopMargin = cmbEmp.Top + cmbEmp.Height + 30
    
    FLV.Move 50, TopMargin, Me.ScaleWidth - 100, Me.ScaleHeight - TopMargin
    
    LV.Move 100, 200, Me.Width - 400, FLV.Height - FButs1.Height - 300
    
    FButs1.Move LV.Left + LV.Width - FButs1.Width, LV.Top + LV.Height
    
    FButs2.Move LV.Left, LV.Top + LV.Height
End Sub


Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim clr As Long
    If Item.Checked Then
        Item.Selected = True
        Item.SubItems(2) = "OverTime"
    Else
        Item.Selected = True
        Item.SubItems(2) = ""
    End If
    
End Sub


Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = vbRightButton And LV.ListItems.Count > 0 Then
    PopUpMenu mnuPopup
    Exit Sub
End If
    'Set Flag that mouse is in attendance column or not
    forLeave = CBool(x >= LV.ColumnHeaders(3).Left And x <= (LV.ColumnHeaders(3).Left + LV.ColumnHeaders(3).Width))
    

End Sub

Private Sub mnuDelete_Click()

con.Execute "DELETE FROM AttendanceSheet WHERE DT='" & DTAttend.value & "' AND EmpID='" & LV.SelectedItem.Text & "' AND Shift='" & cmbTime.Tag & "'"
Call FillList

End Sub
