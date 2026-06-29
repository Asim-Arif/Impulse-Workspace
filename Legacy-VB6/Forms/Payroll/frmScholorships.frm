VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmScholorShips 
   ClientHeight    =   7665
   ClientLeft      =   75
   ClientTop       =   -15
   ClientWidth     =   12075
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   7665
   ScaleMode       =   0  'User
   ScaleWidth      =   10815.75
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSave 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   11520
      Picture         =   "frmScholorships.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Add This Entry To Voucher"
      Top             =   900
      UseMaskColor    =   -1  'True
      Width           =   525
   End
   Begin VB.ComboBox cmbScholorshipType 
      Height          =   315
      ItemData        =   "frmScholorships.frx":0532
      Left            =   8280
      List            =   "frmScholorships.frx":0542
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1140
      Width           =   1695
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   9960
      TabIndex        =   7
      Top             =   1140
      Width           =   525
   End
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   2580
      TabIndex        =   6
      Top             =   1170
      Width           =   3600
      _ExtentX        =   6350
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
   Begin MSComCtl2.DTPicker dt 
      Height          =   315
      Left            =   6195
      TabIndex        =   3
      Top             =   1140
      Width           =   2070
      _ExtentX        =   3651
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
      CustomFormat    =   " MMM  yyyy"
      Format          =   82903043
      CurrentDate     =   37384
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   1170
      Width           =   2490
      _ExtentX        =   4392
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
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6225
      Left            =   75
      TabIndex        =   0
      Top             =   1425
      Width           =   11970
      Begin MSComctlLib.ListView LV 
         Height          =   5535
         Left            =   75
         TabIndex        =   1
         Top             =   180
         Width           =   11790
         _ExtentX        =   20796
         _ExtentY        =   9763
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "EmpID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Employee Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Department"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Starting From"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Type"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Qty"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Amount"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   375
         Left            =   8580
         TabIndex        =   11
         Top             =   5760
         Width           =   1620
         Caption         =   "    Print          "
         PicturePosition =   327683
         Size            =   "2857;661"
         Accelerator     =   80
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdCancel 
         Height          =   375
         Left            =   10275
         TabIndex        =   4
         Top             =   5760
         Width           =   1620
         Caption         =   "Close        "
         PicturePosition =   327683
         Size            =   "2857;661"
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2070
      Top             =   11790
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
            Picture         =   "frmScholorships.frx":0567
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScholorships.frx":05C5
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin MSForms.ComboBox cmbYear 
      Height          =   285
      Left            =   60
      TabIndex        =   14
      Top             =   570
      Width           =   1140
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2011;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbMonth 
      Height          =   285
      Left            =   1200
      TabIndex        =   13
      Top             =   570
      Width           =   1350
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2381;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label1 
      Height          =   240
      Left            =   60
      TabIndex        =   12
      Top             =   330
      Width           =   2460
      BackColor       =   11517387
      VariousPropertyBits=   8388627
      Caption         =   " Year                 Month"
      Size            =   "4339;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label lblAmt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10500
      TabIndex        =   8
      Top             =   1140
      Width           =   990
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   60
      TabIndex        =   5
      Top             =   900
      Width           =   11430
      BackColor       =   11517387
      VariousPropertyBits=   8388627
      Caption         =   $"frmScholorships.frx":0623
      Size            =   "20161;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmScholorShips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim edit As Boolean

Private Sub cmbdept_matched()
    
    cmbEmp.ClearVals
    
    If cmbDept.ID = "0" Then
        cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where Active=1  and DeptID<>'VDRS' "
    Else
        cmbEmp.AddVals con, "Name", "Employees", "EmpID", "WHERE DeptID='" & cmbDept.ID & "' and Active=1  and DeptID<>'VDRS'"
        
    End If

End Sub

Private Sub cmbMonth_Change()
    Call cmbMonth_Click
End Sub

Private Sub cmbMonth_Click()
    If cmbMonth.MatchFound And cmbYear.MatchFound Then
        Call FillList
    End If
End Sub

Private Sub cmbScholorshipType_Click()

    Dim TotalAssigned As Long, ToAssing As Long
    
    TotalAssigned = con.Execute("Select isnull(Sum(isnull(Qty,0)),0) as TotalAssigned From ScholorShips " & _
     "Where year(StartDT)=" & year(DT) & " and Month(DT)=" & cmbMonth.ListIndex + 1 & _
     " and SType=" & cmbScholorshipType.ListIndex).Fields(0)
    
    Select Case cmbScholorshipType.ListIndex
        Case -1
            cmbScholorshipType.Tag = 0
            Exit Sub
        Case 0
            ToAssign = 20 - TotalAssigned '10 - TotalAssigned
            cmbScholorshipType.Tag = 100
        Case 1
            ToAssign = 20 - TotalAssigned '3 - TotalAssigned
            cmbScholorshipType.Tag = 150
        Case 2
            ToAssign = 20 - TotalAssigned '2 - TotalAssigned
            cmbScholorshipType.Tag = 200
        
    End Select
    
    txtQty.Tag = ToAssign
    
    If ToAssign = 0 Then
        MsgBox "All " & cmbScholorshipType & " Scholorships have been assigned.", vbInformation
        cmbScholorshipType.ListIndex = -1
    End If
    
    lblAmt = Format(Val(txtQty) * Val(cmbScholorshipType.Tag), "0.00")
    
End Sub

Private Sub cmbYear_Change()
    Call cmbYear_Click
End Sub

Private Sub cmbYear_Click()
    If cmbMonth.MatchFound And cmbYear.MatchFound Then
        Call FillList
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If Saved Then
        Unload Me
        NewAdvShort.Show
    End If
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub

Private Sub cmdOpt_Click()

PopUpMenu mnuPopup

End Sub

Private Sub cmdPrint_Click()
    
'    If Not (cmbYear.MatchFound And cmbMonth.MatchFound) Then
'        Exit Sub
'    End If
    
    Dim f As New frmPrevRpt
    Dim rpt As CRAXDDRT.Report
    'Set rpt = rpt.OpenSubreport(rptPayrollPath & "\rptScholorships.rpt")
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptScholorships.rpt")
    rpt.FormulaFields(6).Text = "'" & cmbMonth & " , " & cmbYear & "'"
    f.ShowReport "Month({VScholorships.DT})=" & cmbMonth.ListIndex + 1 & " AND Year({VScholorships.DT})=" & cmbYear, rpt
'    f.ShowReport " ", rpt
    Set rpt = Nothing
    Set f = Nothing
    
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
    Dim DT As Date
    DT = DateSerial(cmbYear, cmbMonth.ListIndex + 1, 1)
    
    If Not cmbEmp.MatchFound Then
        MsgBox "Please Select The Employee First", vbInformation
        cmbEmp.SetFocus
        Exit Sub
'    ElseIf year(dt) <> year(GetServerDate) Then
'        If MsgBox("You are Entering Scholorship For Another Year" & _
'        vbNewLine & "Are You Sure To Enter it in Year " & year(dt), vbQuestion + vbYesNo) = vbNo Then Exit Sub
    ElseIf cmbScholorshipType.ListIndex < 0 Then
        MsgBox "Please Select The Scholorship Type.", vbInformation
        cmbEmp.SetFocus
        Exit Sub
    ElseIf Val(txtQty) > Val(txtQty.Tag) Then
        MsgBox "Cannot Assign More Than " & txtQty.Tag & " " & cmbScholorshipType & " Scholorships", vbInformation
        txtQty.SetFocus
        Exit Sub
    End If
    
    con.Execute "Insert Into Scholorships (EmpID,StartDt,SType,Qty,DT) Values ('" & _
     cmbEmp.ID & "','" & Format(DT, "dd-MMM-yyyy") & "'," & cmbScholorshipType.ListIndex & "," & _
     Val(txtQty) & ",'" & DateSerial(cmbYear, cmbMonth.ListIndex + 2, 0) & "')"
    
    Call cmbScholorshipType_Click
    
    Call FillList
    
Exit Sub

err:
    If err.Number = -2147217900 Then
        MsgBox "This Employee Already has Scholorship(s) Assigned.", vbInformation
    Else
        MsgBox err.Description
    End If
    
End Sub

Private Sub DT_Change()
    DT = DateSerial(DT.year, DT.Month, 1)
End Sub

Private Sub Form_Load()
    
    cmbMonth.AddItem "Jan"
    cmbMonth.List(0, 1) = 1
    cmbMonth.AddItem "Feb"
    cmbMonth.List(1, 1) = 2
    cmbMonth.AddItem "Mar"
    cmbMonth.List(2, 1) = 3
    cmbMonth.AddItem "Apr"
    cmbMonth.List(3, 1) = 4
    cmbMonth.AddItem "May"
    cmbMonth.List(4, 1) = 5
    cmbMonth.AddItem "Jun"
    cmbMonth.List(5, 1) = 6
    cmbMonth.AddItem "Jul"
    cmbMonth.List(6, 1) = 7
    cmbMonth.AddItem "Aug"
    cmbMonth.List(7, 1) = 8
    cmbMonth.AddItem "Sep"
    cmbMonth.List(8, 1) = 9
    cmbMonth.AddItem "Oct"
    cmbMonth.List(9, 1) = "10"
    cmbMonth.AddItem "Nov"
    cmbMonth.List(10, 1) = 11
    cmbMonth.AddItem "Dec"
    cmbMonth.List(11, 1) = 12
    
    Dim CYear As Long
    CYear = Format(GetServerDate, "yyyy")
    For i = 2002 To CYear
        cmbYear.AddItem i
    Next i
    
    DT = GetServerDate
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1600
    
    cmbDept.ClearVals
    
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID", "Where DeptID<>'VDRS'"
    
    cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where Active=1 and EmpID Not In (Select EmpID From Scholorships Where Year(StartDT)=" & year(DT) & ")"
    cmbDept.Text = "<All Departments>"
'    Call FillList
    
End Sub


Private Sub FillList()
      
    
    
    Dim SQL As String, LItem As ListItem
    Dim rs As New ADODB.Recordset
    
    'Sql = "Select * from VScholorships Where YEAR(startDT)=" & dt.year & " AND Active=1"
    SQL = "Select * from VScholorships Where Month(DT)=" & cmbMonth.ListIndex + 1 & " AND Year(DT)=" & cmbYear
    'SQL = "Select * from VScholorships "
    LV.ListItems.Clear
    
    With rs
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            Set LItem = LV.ListItems.add(, !EntryID & "'", ![EmpId])
            LItem.SubItems(1) = ![Name] & " " & ![Rel] & " " & ![FName]
            LItem.SubItems(2) = ![DeptName] & ""
            LItem.SubItems(3) = Format(![StartDT], "MMM yyyy")
            LItem.SubItems(4) = ![ScholorshipType]
            LItem.SubItems(5) = Val(![Qty] & "")
            LItem.SubItems(6) = Format(Val(![Amount] & ""), "0.00")
            
            .MoveNEXT
        Loop
        .Close
    End With
    
    SQL = ""
    theKey = ""
    theName = ""
    
    Set rs = Nothing
    
End Sub


Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo err
    Dim iRecAff As Integer
    If KeyCode = 46 Then    'Delete Key Pressed
        If MsgBox("Are You Sure To Delete This Scholorship?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        con.Execute "Delete From Scholorships Where EntryID=" & Val(LV.SelectedItem.key), iRecAff
        LV.ListItems.Remove LV.SelectedItem.Index
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub txtQty_Change()
    lblAmt = Format(Val(txtQty) * Val(cmbScholorshipType.Tag), "0.00")
End Sub
