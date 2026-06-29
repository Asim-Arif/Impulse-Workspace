VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmEmpProductivityEntries 
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   12840
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleMode       =   0  'User
   ScaleWidth      =   11500.96
   WindowState     =   2  'Maximized
   Begin VB.Frame FAcc 
      ForeColor       =   &H00800000&
      Height          =   8400
      Left            =   60
      TabIndex        =   3
      Top             =   -45
      Width           =   12750
      Begin ComboList.Usercontrol1 cmbEmp 
         Height          =   285
         Left            =   75
         TabIndex        =   7
         Top             =   405
         Width           =   8415
         _ExtentX        =   14843
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
         TextBackColor   =   12640511
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         Height          =   315
         Left            =   10050
         TabIndex        =   1
         Top             =   390
         Width           =   1515
         _ExtentX        =   2672
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
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   108134403
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   8505
         TabIndex        =   0
         Top             =   390
         Width           =   1530
         _ExtentX        =   2699
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
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   108134403
         CurrentDate     =   37055
      End
      Begin MSComctlLib.ListView LV 
         Height          =   7155
         Left            =   60
         TabIndex        =   5
         Top             =   750
         Width           =   12570
         _ExtentX        =   22172
         _ExtentY        =   12621
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   15199215
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S. #"
            Object.Width           =   926
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Article"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Process"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Qty"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Type"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   375
         Left            =   11310
         TabIndex        =   6
         Top             =   7935
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "2355;661"
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   11580
         TabIndex        =   2
         Top             =   330
         Width           =   1050
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "1852;635"
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
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmEmpProductivityEntries.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   150
         Width           =   11445
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   10
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpProductivityEntries.frx":00D3
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpProductivityEntries.frx":028F
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopMenu"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add New Entry"
      End
      Begin VB.Menu dd 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit This Entry"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintReport 
         Caption         =   "Print Report"
      End
   End
End
Attribute VB_Name = "frmEmpProductivityEntries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
   Unload Me
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub



Private Sub Form_Load()
        
    mnuPop.Visible = False
    
    cmbEmp.ListHeight = 4000
    'cmbEmp.AddItem "<All Employees>", "0"
    cmbEmp.AddVals con, "'{' + EmpID + '} ' + Name", "Employees", "EmpID", " WHERE Active=1  ORDER BY EmpID"
    '-------------------------------------
    DtFrom.Value = Date
    DtTo.Value = Date
    
End Sub

Private Sub RefreshLV()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim strCondition As String
  
    Dim ITM As ListItem
    
    Screen.MousePointer = vbHourglass
    
    strCondition = " WHERE DT BETWEEN '" & DtFrom.Value & "' AND '" & DtTo.Value & "'"
    
    If cmbEmp.ID <> "0" Then
        strCondition = strCondition & " AND EmpID='" & cmbEmp.ID & "'"
    End If
    
    Dim strSQL As String
    With rs
        strSQL = "SELECT * FROM VEmpProductivityReport " & strCondition & " ORDER BY EmpID"
        .Open strSQL, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, !EntryID & "'", .AbsolutePosition)
            
            ITM.ListSubItems.add , , !ItemName & ""
            ITM.ListSubItems.add , , !Description & ""
            ITM.ListSubItems.add , , Format(!DT, "dd-MMM-yyyy")
            
             ITM.ListSubItems.add , , Val(!Qty & "")
            
            ITM.ListSubItems.add , , IIf(!iType = 0, "Normal", "Over Time")
            
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description

End Sub


Private Sub LV_DblClick()
    
    Exit Sub
    If LV.ListItems.count = 0 Then Exit Sub
    
    Load frmNewCustomInvoice
    
    Call frmNewCustomInvoice.EditProforma(Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1))
    
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)

    If LV.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo) = vbYes Then
            con.Execute "DELETE FROM EmpProductivityReport WHERE EntryID=" & Val(LV.SelectedItem.key)
            Call RefreshLV
        End If
    End If
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 2 Then Exit Sub
    Me.PopUpMenu mnuPop
      
End Sub

Private Sub mnuAdd_Click()
    frmNewEmpProductivityEntry.Show 1
End Sub

Private Sub mnuedit_Click()
    Call LV_DblClick
End Sub

Private Sub mnuSendEmail_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    
    If MsgBox("Are you sure to Send Email?", vbQuestion + vbYesNo) = vbYes Then
        con.Execute "INSERT INTO AlertsDetail(AlertType,AlertValue) VALUES('Cloth Factory Production','" & LV.SelectedItem.SubItems(3) & "')"
    End If
    
End Sub

Private Sub mnuPrintReport_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\EmployeeProductivityReport.rpt")
    Dim f As New frmPrevRpt
    Dim strSelection As String
    strSelection = "{VEmpProductivityReport.DT}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#"
    strSelection = strSelection & " AND {VEmpProductivityReport.EmpID}='" & cmbEmp.ID & "'"
    f.ShowReport strSelection, rpt
    
End Sub
