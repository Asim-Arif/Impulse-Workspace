VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmEmpLedger 
   BorderStyle     =   0  'None
   ClientHeight    =   7005
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   11550
   LinkTopic       =   "Form19"
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleMode       =   0  'User
   ScaleWidth      =   10345.49
   ShowInTaskbar   =   0   'False
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   90
      TabIndex        =   16
      Top             =   330
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
      TabIndex        =   14
      Top             =   330
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
      Format          =   53870595
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
            Picture         =   "frmEmpLedger.frx":0000
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpLedger.frx":0454
            Key             =   "ORD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpLedger.frx":08AC
            Key             =   "PARIN"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpLedger.frx":0D00
            Key             =   "ITMIN"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpLedger.frx":1154
            Key             =   "CAT"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpLedger.frx":15A8
            Key             =   "OK"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpLedger.frx":1A00
            Key             =   "TYPE"
         EndProperty
      EndProperty
   End
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   2430
      TabIndex        =   2
      Top             =   330
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
      Height          =   6195
      Left            =   60
      TabIndex        =   0
      Top             =   645
      Width           =   11385
      Begin VB.Frame fButs2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   450
         Left            =   60
         TabIndex        =   6
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
            TabIndex        =   10
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
            TabIndex        =   9
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
            TabIndex        =   8
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
            TabIndex        =   7
            Top             =   165
            Width           =   1260
         End
      End
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   480
         Left            =   7755
         TabIndex        =   3
         Top             =   5505
         Width           =   3060
         Begin MSForms.CommandButton cmdPrint 
            Height          =   360
            Left            =   90
            TabIndex        =   5
            Top             =   90
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
            Left            =   1575
            TabIndex        =   4
            Top             =   90
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
         TabIndex        =   11
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Emp. ID"
            Object.Width           =   2962
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   8176
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "In Time"
            Object.Width           =   2990
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Out Time"
            Object.Width           =   3043
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Net Time"
            Object.Width           =   2805
         EndProperty
      End
      Begin MSComctlLib.TreeView TV 
         Height          =   5325
         Left            =   45
         TabIndex        =   13
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
      Left            =   8835
      TabIndex        =   15
      Top             =   285
      Visible         =   0   'False
      Width           =   2595
      Caption         =   "Refresh Data From RFIDs"
      PicturePosition =   327683
      Size            =   "4577;635"
      Picture         =   "frmEmpLedger.frx":1E54
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
      TabIndex        =   12
      Top             =   90
      Width           =   7125
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   7305
      TabIndex        =   1
      Top             =   285
      Width           =   1500
      Caption         =   "Refresh        "
      PicturePosition =   327683
      Size            =   "2646;635"
      Picture         =   "frmEmpLedger.frx":22A2
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



Private Sub AddHeads()

    Dim Rec As New ADODB.Recordset, RecDept As New ADODB.Recordset
    
    TV.ImageList = ImageList1
    TV.Nodes.Clear
    
    
    RecDept.Open "SELECT DeptID,Name FROM Departments", con, adOpenStatic, adLockReadOnly
    For i = 1 To RecDept.RecordCount
        TV.Nodes.Add , , RecDept![DeptID] & "'", RecDept![Name] & "", "TYPE"
        With Rec
            .Open "Select EmpID,Name from Employees WHERE DeptID='" & RecDept!DeptID & "'" & "ORDER BY EmpID", con, adOpenForwardOnly, adLockReadOnly
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

RefreshLedger

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

'With cr1
'    .ReportFileName = RptPath & "\EmpActivitySheet.rpt"
'    '.DataFiles(0) = App.Path & "\Payroll.mdb"
'    .Connect = ConnectStr
'    .SQLQuery = Sql
'    .Formulas(0) = "FromTo='For " & cmbMonth & ", " & cmbYear & ".'"
'    .Formulas(1) = "Company='" & company & "'"
'    .Action = 1
'    .PageZoomNext
'End With

Screen.MousePointer = vbHourglass

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
rpt.SQLQueryString = Sql

f.ShowReport Cond & " ", rpt

Screen.MousePointer = vbDefault

End Sub

Public Sub cmdRefresh_Click()

Call RefreshLedger
    
    'Call Form_Resize
End Sub

Private Sub cmdRefreshRFID_Click()
    Dim Rec As New ADODB.Recordset
    With Rec
        .Open "SELECT CompName FROM Computers", con, adOpenStatic, adLockReadOnly
        For i = 1 To .RecordCount
            If tcpClient.State <> sckClosed Then tcpClient.Close
            tcpClient.RemoteHost = !CompName & ""
            tcpClient.RemotePort = 1012
            tcpClient.Connect
                    
            'Sleep 1000
            'tcpClient.SendData "A"
        Next
    End With
    
    
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
    'cmbEmp.AddItem "<All Employees>", "0"
    cmbEmp.AddVals con, "Name + ' {' + EmpID + '}'", "Employees", "EmpID"
    
    cmbDept.ClearVals
    cmbDept.AddItem "< All Departments >", "0"
    'cmbEmp.AddItem "<All Employees>", "0"
    cmbDept.AddVals con, "Name + ' {' + DeptID + '}'", "Departments", "DeptID"
    
    cmbDept.ID = "0"
    cmbEmp.ID = "0"
    
End Sub
Private Sub Form_Resize()


On Error Resume Next
Dim TopMargin As Integer
' fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
TopMargin = cmbEmp.Top + cmbEmp.Height + 30
 FLV.Move 50, TopMargin, Me.ScaleWidth - 100, Me.ScaleHeight - TopMargin
 
 TV.Move 100, 200, TV.Width, FLV.Height - fButs1.Height - 300
 
 ''''''''''' Commented To Hide Tree View '''''''''''''
 'LV.Move TV.Left + TV.Width - 10, 200, Me.Width - 400 - TV.Width, TV.Height
 
 LV.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
 
 fButs1.Move LV.Left + LV.Width - fButs1.Width, LV.Top + LV.Height
 
 fButs2.Move LV.Left, LV.Top + LV.Height
 
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

For i = 1 To TV.Nodes.Count
    With TV.Nodes(i)
        .Bold = False
        .ForeColor = vbbalck
    End With
Next



Dim Rec As New ADODB.Recordset, PrevEmp As String, ITM As ListItem, Cond As String
Dim Color As Long, Invert As Boolean, Bold As Boolean

'Cond = GetTreeValues

'If Cond = "" Then LV.ListItems.Clear: Exit Sub
'Cond = "WHERE EmpID IN(" & Cond & ")"
'If Cond = "" Then
Cond = "WHERE (DT='" & DT & "' OR ISNULL(DT,'')='')"

If cmbDept.ID = 0 Then
    If cmbEmp.ID <> "0" Then Cond = Cond & " AND EmpID='" & cmbEmp.ID & "'"
Else
    If cmbEmp.ID = "0" Then
        Cond = Cond & " AND DeptID='" & cmbDept.ID & "'"
    Else
    
        Cond = Cond & " AND  empid='" & cmbEmp.ID & "'"
    End If
End If




Sql = "SELECT * FROM VEmpTimes " & Cond & " ORDER BY EmpID"

LV.ListItems.Clear
With Rec
    .Open Sql, con, adOpenStatic, adLockReadOnly
    For i = 1 To .RecordCount
        
        If PrevEmp <> !EmpId & "" Then
            Invert = Not Invert
            Color = IIf(Invert, RGB(100, 100, 200), RGB(50, 150, 250))
            
            Set ITM = LV.ListItems.Add(, , !EmpId & "")
            ITM.ListSubItems.Add(, , !Name & "").ForeColor = Color
            PrevEmp = !EmpId & ""
'            With TV.Nodes(!EmpId & "'")
'                .Bold = True
'                .ForeColor = RGB(100, 100, 200)
'                .Parent.ForeColor = RGB(100, 100, 200)
'                .Parent.Bold = True
'            End With
        Else
            Set ITM = LV.ListItems.Add(, , "")
            ITM.ListSubItems.Add , , ""
        End If
        
        ITM.ForeColor = Color
        ITM.ListSubItems(1).ForeColor = Color
        ITM.ListSubItems.Add(, , Format(!InTime & "", "HH:NN AMPM")).ForeColor = Color
        ITM.ListSubItems.Add(, , Format(!OutTime & "", "HH:NN AMPM")).ForeColor = Color
        ITM.ListSubItems.Add(, , getHrsMin(Val(!Hrs & ""))).ForeColor = Color
        'PrevEmp = !EmpId & ""
        
        If Val(!EntryID & "") > 0 Then
            Bold = True
        Else
            Bold = False
            Color = vbRed
        End If
        
        ITM.Bold = Bold
        ITM.ForeColor = Color
        For k = 1 To ITM.ListSubItems.Count
            ITM.ListSubItems(k).Bold = Bold
            ITM.ListSubItems(k).ForeColor = Color
        Next

        .MoveNext
    Next
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

