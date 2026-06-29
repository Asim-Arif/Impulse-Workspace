VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form EmpAbsents 
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11550
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleMode       =   0  'User
   ScaleWidth      =   10345.49
   WindowState     =   2  'Maximized
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   30
      TabIndex        =   8
      Top             =   0
      Width           =   11490
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Attendance Summary"
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
         Left            =   3300
         TabIndex        =   9
         Top             =   120
         Width           =   4935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Attendance Summary"
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
         Left            =   3315
         TabIndex        =   10
         Top             =   135
         Width           =   4935
      End
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   75
      TabIndex        =   4
      Top             =   885
      Width           =   2220
      _ExtentX        =   3916
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
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   2295
      TabIndex        =   5
      Top             =   885
      Width           =   2220
      _ExtentX        =   3916
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
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
            Picture         =   "EmpAbsents.frx":0000
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EmpAbsents.frx":0452
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin VB.Frame FLV 
      Height          =   6045
      Left            =   15
      TabIndex        =   0
      Top             =   1395
      Width           =   11385
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   4470
         TabIndex        =   14
         Top             =   1575
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   480
         Left            =   7755
         TabIndex        =   11
         Top             =   5505
         Width           =   3060
         Begin MSForms.CommandButton cmdPrint 
            Height          =   360
            Left            =   195
            TabIndex        =   13
            Top             =   60
            Visible         =   0   'False
            Width           =   1425
            ForeColor       =   0
            Caption         =   "Print         "
            PicturePosition =   327683
            Size            =   "2514;635"
            Picture         =   "EmpAbsents.frx":0564
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
            Left            =   1635
            TabIndex        =   12
            Top             =   60
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
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   9287
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   4194304
         BackColor       =   15199215
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
            Text            =   "EmpID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Employee Name"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Absents"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Presents"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Penalty Days"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   9360
      Picture         =   "EmpAbsents.frx":0676
      Top             =   840
      Width           =   465
   End
   Begin MSForms.ComboBox cmbYear 
      Height          =   300
      Left            =   4530
      TabIndex        =   7
      Top             =   885
      Width           =   1110
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "1958;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbMonth 
      Height          =   300
      Left            =   5670
      TabIndex        =   6
      Top             =   885
      Width           =   1320
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2328;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   105
      TabIndex        =   3
      Top             =   645
      Width           =   6855
      BackColor       =   11517387
      Caption         =   "  Department                         Employee                               Year                Month"
      Size            =   "12091;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   7035
      TabIndex        =   2
      Top             =   855
      Width           =   1500
      Caption         =   "Refresh        "
      PicturePosition =   327683
      Size            =   "2646;635"
      Picture         =   "EmpAbsents.frx":0C8C
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
Attribute VB_Name = "EmpAbsents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                                      
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

   ' On Error GoTo err
    Call ShowWait
    
    
    
    con.Execute "Delete  From PrintEmpLedger"
    For i = 1 To LV.ListItems.Count
      With LV.ListItems(i)
         con.Execute "Insert into  printEmpLedger (DT,Shift1,LateHours1,Shift2,LateHours2,OverTime,ShortLoan,LongLoan)" & _
         "values ('" & CDate(.Text) & "','" & .SubItems(2) & "','" & .SubItems(3) & "','" & .SubItems(4) & "'," _
         & Val(.SubItems(5)) & "," & Val(.SubItems(6)) & "," & Val(.SubItems(7)) & "," & Val(.SubItems(8)) & ")"
      End With
    Next
    
    
    
    
        
With cr1
    .ReportFileName = RptPRPath & "\EmpLedger.rpt"
    '.DataFiles(0) = App.Path & "\Payroll.mdb"
    .Connect = ConnectStr
    .Formulas(0) = "FromTo='For " & cmbMonth & ", " & cmbYear & ".'"
    .Formulas(1) = "EmpNAme='Ledger Of " & cmbEmp.Text & "'"
    .Action = 1
    .PageZoomNext
End With

Call HideWait

Exit Sub
ERR:
    MsgBox ERR.Description
    Call HideWait
End Sub

Public Sub cmdRefresh_Click()

Call RefreshLedger
    
    'Call Form_Resize
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
    CYear = Format(ServerDate, "yyyy")
    For i = 2002 To CYear
        cmbYear.AddItem i
    Next i


    cmbMonth.ListIndex = Val(Format$(ServerDate, "MM")) - 1
    cmbYear = Format$(ServerDate, "yyyy")



    DtTo = ServerDate
    DtFrom = DateAdd("M", -1, ServerDate)
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID"
    cmbDept.Text = "<All Departments>"
    'cmbEmp.Text = "<All Employees>"
    'Set LV.SmallIcons = ImageList1
    
End Sub
Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals con, "Name", "Employees", "EmpID", "WHERE Vender=0"
    Else
        cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' AND Vender=0"
    End If
    
End Sub

Private Sub Form_Resize()


On Error Resume Next

 fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2

 FLV.Move 50, 1230, Me.ScaleWidth - 100, Me.ScaleHeight - 1400
 
 LV.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
 
 fButs1.Move LV.Left + LV.Width - fButs1.Width, LV.Top + LV.Height
 

 
End Sub



Private Sub RefreshLedger()

    On Error GoTo ERR
    If cmbDept.MatchFound = False Or cmbEmp.MatchFound = False Then Exit Sub
    Dim rsLedger As New ADODB.Recordset
    Dim rsShifts As New ADODB.Recordset
    Dim ITM As ListItem
    Dim Pos As Integer
    Dim LateHours1 As Double
    Dim LateHours2 As Double
    
    Dim DT As Date
    Dim DTs As New Collection
    Dim LastDay As Integer
    Dim TheColor As Long
    Dim iDay As Integer
    Dim OverTime As String
    Dim ShortTerm As String
    Dim LongTerm As String
    
    Dim rs As New ADODB.Recordset
    
    
    If cmbEmp.ID = "0" Then
       EmpId = IIf(cmbDept.ID = "0", "", cmbDept.ID)
    Else
       EmpId = IIf(cmbEmp.ID = "0", "", cmbEmp.ID)
    End If
    
    LV.ListItems.Clear
    
    
    With rs
        'Add Employees
        .Open "Select EmpId,Name,rel,FName From Employees Where EmpID Like '" & EmpId & "%' Order By EmpID", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            Set ITM = LV.ListItems.Add(, ![EmpId] & "", ![EmpId] & "")
            ITM.SubItems(1) = ![Name] & " " & ![Rel] & " " & ![FName]
            ITM.SubItems(2) = "0"
            ITM.SubItems(3) = "0"
            .MoveNext
        Loop
        .Close
        
        'Add Absents of Employees
        .Open "Select EmpID,Count(Attendance)  as Absents From AttendanceSheet Where year(DT)=" & cmbYear & " and Month(DT)=" & cmbMonth.List(cmbMonth.ListIndex, 1) & " and EmpID like '" & EmpId & "%' and Attendance=0 Group By EmpID", con, adOpenStatic, adLockReadOnly
        
        
        Do Until .EOF
            LV.ListItems(![EmpId] & "").SubItems(2) = Val(![Absents] & "") / 2 & " "
            .MoveNext
        Loop
        .Close
        
        'Add Presnts of Employees
        .Open "Select EmpID,Count(Attendance)  as Presents From AttendanceSheet Where year(DT)=" & cmbYear & " and Month(DT)=" & cmbMonth.List(cmbMonth.ListIndex, 1) & " and EmpID like '" & EmpId & "%' and Attendance=1 Group By EmpID", con, adOpenStatic, adLockReadOnly
        
        Do Until .EOF
            LV.ListItems(![EmpId] & "").SubItems(3) = Val(![Presents] & "") / 2 & " "
            .MoveNext
        Loop
        
        .Close
        
        'Add Panalty of Employees
        .Open "SELECT EmpID, isnull(Days,0) as PenaltyDay FROM PenaltyDays Where year(DT)=" & cmbYear & " and Month(DT)=" & cmbMonth.List(cmbMonth.ListIndex, 1) & " and EmpID like '" & EmpId & "%'", con, adOpenStatic, adLockReadOnly
        
        Do Until .EOF
            LV.ListItems(![EmpId] & "").SubItems(4) = Val(![PenaltyDay] & "") & " "
            .MoveNext
        Loop
        .Close
        
    End With
    Set rs = Nothing
    
    Exit Sub
ERR:
        MsgBox ERR.Number & vbNewLine & ERR.Description
        Exit Sub
        
    
    
    
End Sub




Private Sub LV_DblClick()
If LV.ListItems.Count = 0 Then Exit Sub
    With LV.SelectedItem
        txtEdit.Move LV.Left + LV.ColumnHeaders(5).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(5).Width
        txtEdit = .SubItems(4)
    End With
    
    With txtEdit
         .Visible = True
         .SelStart = 0
         .SelLength = Len(.Text)
         .SetFocus
     End With
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        DT = Format(DateSerial(Val(cmbYear), cmbMonth.List(cmbMonth.ListIndex, 1), 1), "dd-MMM-yyyy")
        LV.SelectedItem.SubItems(4) = Val(txtEdit.Text)
        con.Execute "Delete From PenaltyDays Where year(DT)=" & cmbYear & " and Month(DT)=" & cmbMonth.List(cmbMonth.ListIndex, 1) & " and Empid='" & LV.SelectedItem.Text & "'"
        con.Execute "insert into Penaltydays (Empid,DT,Days) values ('" & LV.SelectedItem.key & "','" & DT & "'," & Val(txtEdit) & ")"
        txtEdit.Text = ""
        txtEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtEdit.Text = ""
        txtEdit.Visible = False
    Else
        KeyAscii = OnlyNumber(KeyAscii)
    End If
End Sub
