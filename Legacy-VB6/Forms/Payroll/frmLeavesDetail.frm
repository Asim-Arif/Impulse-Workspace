VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmLeavesDetail 
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   11550
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   6900
   ScaleMode       =   0  'User
   ScaleWidth      =   10345.49
   WindowState     =   2  'Maximized
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
            Picture         =   "frmLeavesDetail.frx":0000
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLeavesDetail.frx":0452
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   75
      TabIndex        =   3
      Top             =   390
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
      Height          =   285
      Left            =   2295
      TabIndex        =   4
      Top             =   390
      Width           =   2610
      _ExtentX        =   4604
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
      Height          =   6045
      Left            =   45
      TabIndex        =   0
      Top             =   750
      Width           =   11385
      Begin VB.CommandButton cmdDelete 
         CausesValidation=   0   'False
         Height          =   255
         Left            =   10860
         Picture         =   "frmLeavesDetail.frx":0564
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Remove This Entry From Voucher"
         Top             =   180
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   570
         Left            =   7890
         TabIndex        =   5
         Top             =   5445
         Width           =   3015
         Begin MSForms.CommandButton cmdClose 
            Height          =   360
            Left            =   1530
            TabIndex        =   7
            Top             =   135
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
            Top             =   135
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
         Width           =   10785
         _ExtentX        =   19024
         _ExtentY        =   9287
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "EmpID"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   3387
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Leave Reason"
            Object.Width           =   3067
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Authority Person"
            Object.Width           =   3616
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Leave Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "From"
            Object.Width           =   1905
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "To"
            Object.Width           =   1905
         EndProperty
      End
   End
   Begin MSComCtl2.DTPicker DTFrom 
      Height          =   300
      Left            =   4905
      TabIndex        =   9
      Top             =   390
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
      Format          =   107282435
      CurrentDate     =   37384
   End
   Begin MSComCtl2.DTPicker DTTo 
      Height          =   300
      Left            =   6660
      TabIndex        =   11
      Top             =   390
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
      Format          =   107282435
      CurrentDate     =   37384
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Department                         Employee                                        From                            To"
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
      Height          =   255
      Index           =   4
      Left            =   90
      TabIndex        =   8
      Top             =   120
      Width           =   8295
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   8430
      TabIndex        =   2
      Top             =   345
      Width           =   1560
      Caption         =   "Refresh        "
      PicturePosition =   327683
      Size            =   "2752;635"
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
Attribute VB_Name = "frmLeavesDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbEmp_matched()
    Call RefreshLedger

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()

If LV.ListItems.count = 0 Then Exit Sub
Dim Cond As String

If MsgBox("Are You Sure You Want To Delete This Leave ?", vbInformation + vbYesNo) = vbNo Then Exit Sub

con.BeginTrans
With LV.SelectedItem

    con.Execute "DELETE FROM Leaves  WHERE EntryID='" & LV.SelectedItem.Tag & "'"
    If LV.SelectedItem.SubItems(4) = "Full Day" Then
        con.Execute "DELETE FROM AttendanceSheet  WHERE EmpID='" & LV.SelectedItem & "' AND DT='" & DTPicker1 & "'"
    End If
End With
con.CommitTrans
RefreshLedger

cmdDelete.Visible = False

End Sub

Private Sub cmdPrint_Click()

Screen.MousePointer = vbHourglass

Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptLeaveSlip.rpt")

Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

Set FormulaFields = rpt.FormulaFields
For Each FormulaField In FormulaFields
    If FormulaField.Name = "{@Company}" Then
        FormulaField.Text = "'" & strCompany & "'"
    End If
Next

'rpt.SQLQueryString = Sql
Cond = "{VLeaves.EntryID}=" & LV.SelectedItem.Tag & ""

f.ShowReport Cond & " ", rpt

Screen.MousePointer = vbDefault

End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLedger
    
    'Call Form_Resize
End Sub

Private Sub DTPicker1_Change()
RefreshLedger

End Sub

Private Sub Form_Load()

    DtTo = GetServerDate
    DtFrom = DateAdd("M", -1, GetServerDate)
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    
    DTPicker1 = Date
    
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
        cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where Active=1"
    Else
        cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' and Active"
    End If
    
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    Dim TopMargin As Integer
    TopMargin = cmbEmp.Top + cmbEmp.Height + 30
    
    FLV.Move 50, TopMargin, Me.ScaleWidth - 100, Me.ScaleHeight - TopMargin
    
    LV.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
    
    fButs1.Move LV.Left + LV.Width - fButs1.Width, LV.Top + LV.Height
 
End Sub



Private Sub RefreshLedger()

    
    Dim rsLedger As New ADODB.Recordset
    Dim ITM As ListItem
    Dim Cond As String
    Dim SQL As String
    Dim ICO As String
    
    
    
    If Not (cmbEmp.MatchFound Or cmbDept.MatchFound) Then
        MsgBox "Selection Criterial Is Invalid !", vbInformation
        Exit Sub
    End If
    Cond = "WHERE DT BETWEEN '" & DtFrom & "' AND '" & DtTo & "'"
    
    If cmbEmp.ID <> "0" Then
        Cond = Cond & " AND EmpID='" & cmbEmp.ID & "'"
    Else
        If cmbDept.ID <> "0" Then
            Cond = " WHERE  DeptID='" & cmbDept.ID & "'"
        End If
    End If
        
    SQL = "select * from VLeaves " & Cond & " Order By EmpID"

    With rsLedger
        
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set ITM = LV.ListItems.add(, , !EmpId & "")
            ITM.Tag = !EntryID
            ITM.SubItems(1) = !Name & ""
            ITM.SubItems(2) = Format(!DT, "dd-mm-yyyy")
            ITM.SubItems(3) = !Reason & ""
            ITM.SubItems(4) = !AuthPerson & ""
            ITM.SubItems(5) = IIf(!Full, "Full Day", "Short")
            
            If Not !Full Then
                ITM.SubItems(6) = Format(![From] & "", "hh:mm AMPM")
                ITM.SubItems(7) = Format(![To] & "" & "", "hh:mm AMPM")
            End If
            .MoveNEXT
        Next
        .Close
    End With
    Set rsLedger = Nothing
    SQL = ""
    Cond = ""
    ICO = ""
End Sub


Private Sub LV_ItemClick(ByVal Item As MSComctlLib.ListItem)

    If LV.ListItems.count = 0 Then Exit Sub
    
    cmdDelete.Move LV.ColumnHeaders(8).Left + LV.ColumnHeaders(8).Width, LV.Top + LV.SelectedItem.Top
    cmdDelete.Visible = True

End Sub
