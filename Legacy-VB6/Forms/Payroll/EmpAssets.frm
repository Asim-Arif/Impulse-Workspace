VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form EmpAssets 
   BorderStyle     =   0  'None
   ClientHeight    =   7395
   ClientLeft      =   15
   ClientTop       =   -90
   ClientWidth     =   11550
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   ScaleHeight     =   7395
   ScaleMode       =   0  'User
   ScaleWidth      =   10345.49
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6840
      Top             =   600
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
            Picture         =   "EmpAssets.frx":0000
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EmpAssets.frx":0452
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   330
      Left            =   60
      TabIndex        =   3
      Top             =   375
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
      Height          =   330
      Left            =   2280
      TabIndex        =   4
      Top             =   375
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
   Begin VB.Frame FLV 
      Height          =   6630
      Left            =   30
      TabIndex        =   0
      Top             =   690
      Width           =   11385
      Begin Crystal.CrystalReport cr1 
         Left            =   2580
         Top             =   435
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
         Height          =   5925
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   10451
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Assign Date"
            Object.Width           =   4322
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Asset Description"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Type"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Assigned For"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   570
         Left            =   7890
         TabIndex        =   5
         Top             =   5985
         Width           =   3015
         Begin MSForms.CommandButton cmdClose 
            Height          =   360
            Left            =   1530
            TabIndex        =   7
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
            TabIndex        =   6
            Top             =   195
            Width           =   1425
            ForeColor       =   0
            Caption         =   "Print         "
            PicturePosition =   327683
            Size            =   "2514;635"
            Picture         =   "EmpAssets.frx":0564
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
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Department                         Employee"
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
      Left            =   75
      TabIndex        =   8
      Top             =   120
      Width           =   4410
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   4620
      TabIndex        =   2
      Top             =   285
      Width           =   1500
      Caption         =   "Refresh        "
      PicturePosition =   327683
      Size            =   "2646;635"
      Picture         =   "EmpAssets.frx":0676
      Accelerator     =   82
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuNewAsset 
         Caption         =   "New Asset"
      End
      Begin VB.Menu mnuUpdateAsset 
         Caption         =   "Update This Asset"
      End
      Begin VB.Menu mnuDeleteAsset 
         Caption         =   "Delete This Asset"
      End
      Begin VB.Menu lndelete 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAssetsReport 
         Caption         =   "Print Assets Report"
      End
   End
End
Attribute VB_Name = "EmpAssets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbEmp_matched()
RefreshLedger
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub




Private Sub cmdPrint_Click()
    
    '    On Error GoTo ERR
    '    Screen.MousePointer = vbHourglass
    
    Cond = "{VSalaryLedger.EmpID}='" & cmbEmp.ID & "'"
    '
    'With cr1
    '    .ReportFileName = RptPath & "\EmpAssetsLedger.Rpt"
    '    .Connect = ConnectStr
    '    .SelectionFormula = Cond
    '    .Formulas(0) = "Company='" & strCompany & "'"
    '    .Action = 1
    '    .PageZoomNext
    'End With


    Screen.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpAssets.rpt")
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
    
    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & strCompany & "'"
        End If
    Next
    
    'SelFormula = "Year({Holidays.DT})=" & cmbYear
    Me.Hide
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
    'cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where Active=1"
    Else
        cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' and Active=1"
    End If
    
    
End Sub

Private Sub Form_Resize()

On Error Resume Next

' fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
Dim TopMargin As Integer
TopMargin = cmbEmp.Top + cmbEmp.Height + 30

 FLV.Move 50, TopMargin, Me.ScaleWidth - 100, Me.ScaleHeight - TopMargin
 
 LV.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
 
 fButs1.Move LV.Left + LV.Width - fButs1.Width, LV.Top + LV.Height
 
 'cmdNewST.Move LV.Left, LV.Top + LV.Height

 
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
        Cond = " Where  empid='" & cmbEmp.ID & "'"
    End If
    
    SQL = "SELECT * FROM EmpAssets " & Cond & " ORDER BY DT ASC"

    With rsLedger
        
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            'If ![Attendance] Then
            '    ICO = "Present"
            'Else
            '    ICO = "Absent"
            'End If
            Set ITM = LV.ListItems.add(, , Format(![DT] & "", "dddd dd MMMM ,yyyy"))
            ITM.Tag = !AssetID
            ITM.SubItems(1) = ![Description] & ""
            ITM.SubItems(2) = !Type & ""
            ITM.SubItems(3) = Val(![Amount] & "")
            ITM.SubItems(4) = Format(!AssignFor, "dddd dd-MMM-yyyy")
            If !Type & "" = "Permanent" Then ITM.SubItems(4) = ""
            .MoveNEXT
        Next
        .Close
    End With
    Set rsLedger = Nothing
    SQL = ""
    Cond = ""
    ICO = ""
End Sub




Private Sub LV_DblClick()
If LV.ListItems.count > 0 Then
    With AssignAssets
        .edit = True
        .Tag = LV.SelectedItem.Tag
        .Show 1
    End With
    cmdRefresh_Click
End If
End Sub


Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

If Button = vbRightButton Then
    PopUpMenu mnuPopup
End If

End Sub

Private Sub mnuAssetsReport_Click()
cmdPrint_Click
End Sub

Private Sub mnuDeleteAsset_Click()

If LV.ListItems.count = 0 Then Exit Sub

con.Execute "DELETE FROM EmpAssets WHERE AssetID='" & LV.SelectedItem.Tag & "'"
RefreshLedger

End Sub

Private Sub mnuNewAsset_Click()
    AssignAssets.edit = False
    AssignAssets.Show 1
    RefreshLedger
    
End Sub

Private Sub mnuUpdateAsset_Click()
LV_DblClick
End Sub
