VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form Attend 
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   10095
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleMode       =   0  'User
   ScaleWidth      =   9042.228
   WindowState     =   2  'Maximized
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   75
      TabIndex        =   2
      Top             =   0
      Width           =   9915
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Attendacne Sheet"
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
         Left            =   3615
         TabIndex        =   3
         Top             =   120
         Width           =   2685
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Attendacne Sheet"
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
         Left            =   3630
         TabIndex        =   4
         Top             =   135
         Width           =   2685
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7500
      Top             =   615
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
            Picture         =   "Attend.frx":0000
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Attend.frx":0452
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   8490
      Top             =   765
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
      Left            =   60
      TabIndex        =   12
      Top             =   870
      Width           =   2220
      _ExtentX        =   3916
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
      Left            =   2280
      TabIndex        =   13
      Top             =   870
      Width           =   2220
      _ExtentX        =   3916
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   6015
      TabIndex        =   14
      Top             =   900
      Width           =   1335
      _ExtentX        =   2355
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
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   61276163
      CurrentDate     =   37055
   End
   Begin VB.Frame FLV 
      Height          =   5520
      Left            =   60
      TabIndex        =   0
      Top             =   1170
      Width           =   9945
      Begin VB.Frame FButs2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   570
         Left            =   105
         TabIndex        =   9
         Top             =   4950
         Width           =   3000
         Begin MSForms.CommandButton cmdCheckAll 
            Height          =   360
            Left            =   0
            TabIndex        =   11
            Top             =   135
            Width           =   1425
            ForeColor       =   0
            Caption         =   "Check All   "
            PicturePosition =   327683
            Size            =   "2514;635"
            Picture         =   "Attend.frx":0564
            Accelerator     =   104
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdClearAll 
            Height          =   360
            Left            =   1500
            TabIndex        =   10
            Top             =   135
            Width           =   1425
            ForeColor       =   0
            Caption         =   "Clear All    "
            PicturePosition =   327683
            Size            =   "2514;635"
            Accelerator     =   108
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Frame FButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   570
         Left            =   5550
         TabIndex        =   6
         Top             =   4950
         Width           =   3825
         Begin MSForms.CommandButton cmdSave 
            Height          =   360
            Left            =   60
            TabIndex        =   8
            Top             =   135
            Width           =   2145
            ForeColor       =   0
            Caption         =   " Save Attendence   "
            PicturePosition =   327683
            Size            =   "3784;635"
            Picture         =   "Attend.frx":06BE
            Accelerator     =   83
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
            Left            =   2295
            TabIndex        =   7
            Top             =   135
            Width           =   1425
            ForeColor       =   0
            Caption         =   " Close"
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
      Begin VB.TextBox txtLateHours 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1485
         TabIndex        =   1
         Top             =   780
         Visible         =   0   'False
         Width           =   915
      End
      Begin MSComctlLib.ListView LV 
         Height          =   4740
         Left            =   90
         TabIndex        =   5
         Top             =   165
         Width           =   9270
         _ExtentX        =   16351
         _ExtentY        =   8361
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
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
            Text            =   "Emp Name"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "F Name"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Emp. ID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Shift"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Late Hours"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   9120
      Picture         =   "Attend.frx":07D0
      Top             =   720
      Width           =   360
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   90
      TabIndex        =   16
      Top             =   615
      Width           =   7245
      BackColor       =   11517387
      Caption         =   "  Department                         Employee Name                  Shift                          Date"
      Size            =   "12779;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cmbShift 
      Height          =   285
      Left            =   4500
      TabIndex        =   15
      Top             =   900
      Width           =   1455
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "2566;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "Attend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Editing As Boolean
Public Sub enterLateHours()
    Load LateHours
    LateHours.Tag = LV.SelectedItem
    LateHours.Show
End Sub

Private Sub cmbEmp_matched()

On Error Resume Next
LV.ListItems(cmbEmp.ID).Selected = True
LV.ListItems(cmbEmp.ID).EnsureVisible

End Sub

Private Sub cmbShift_Change()
    Call FillList
End Sub

Private Sub cmdCheckAll_Click()
    With LV
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = True
        Next i
    End With
End Sub

Private Sub cmdClearAll_Click()
    With LV
        For i = 1 To .ListItems.Count
            .ListItems(i).Checked = False
        Next i
    End With
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub


Public Sub cmdRefresh_Click()
    'Call RefreshLedger
    
    'Call Form_Resize
End Sub

Private Sub cmdSave_Click()


'   If UCase(Format(DTPicker, "ddd")) = "SUN" Then
'        MsgBox "Attendance for Sunday is not allowed", vbInformation
'        Exit Sub
'   End If


    Dim Result As Boolean
    Dim LateHours As Double
    Dim Sql As String
    con.BeginTrans
    With LV
        For i = 1 To .ListItems.Count
            Result = .ListItems(i).Checked
            LateHours = Val(.ListItems(i).SubItems(4))
            If .ListItems(i).Tag = "Edit" Then
                Sql = "Update AttendanceSheet set Attendance=" & _
                Abs(Result) & ",LateHours=" & _
                IIf(LateHours = "0", "Null", LateHours) & _
                " Where EmpID='" & .ListItems(i).key & "' and Shift='" & _
                cmbShift & "' and DT='" & DTPicker1 & "'"
                con.Execute Sql
            Else
                Sql = "insert into AttendanceSheet (EmpID," & _
                 "DT,Shift,Attendance,LateHours) Values('" & _
                 .ListItems(i).key & "','" & DTPicker1 & "','" & _
                 .ListItems(i).SubItems(3) & "'," & Abs(Result) & _
                 "," & IIf(LateHours = "0", "Null", LateHours) & ")"
                 con.Execute Sql
                 
                 
                 ' Put the Second Times Attendance Same as
                 ' The First Shift By Default.
                 
                 On Error Resume Next
                 
                 If cmbShift.ListIndex = 0 Then
                     Sql = "insert into AttendanceSheet (EmpID," & _
                     "DT,Shift,Attendance,LateHours) Values('" & _
                     .ListItems(i).key & "','" & DTPicker1 & "','" & _
                     cmbShift.List(1) & "'," & Abs(Result) & _
                     ",0)"
                     con.Execute Sql
                 End If
                 
            End If
            LV.ListItems(i).Tag = "Edit"
            
        Next i
    End With
    con.CommitTrans
    
    
End Sub



Private Sub DTPicker1_Change()
    Call FillList
End Sub

Private Sub Form_Load()
Dim AMPM As String
 
 
 
 
 AMPM = Format(con.Execute("Select Getdate()").Fields(0), "AMPM")
     
     
     
    DTPicker1 = Format(ServerDate, "dd-MMM-yyyy")
    
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1600
 
    cmbShift.AddItem "First"
    cmbShift.AddItem "Second"
    
    cmbShift.Enabled = False
    
   'First Shift Attendance is Allowed Only in AM Hrs
   'Second Shift Attendance is Only in PM Hrs
 
    If AMPM = "AM" Then
      cmbShift.ListIndex = 0
    Else
      cmbShift.ListIndex = 1
    End If
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID"
    
    cmbShift.Enabled = UserHasAccess("AttShiftChange")
    DTPicker1.Enabled = UserHasAccess("AttDateChange")
    
End Sub
Private Sub cmbdept_matched()

    Call FillList
    
    cmbEmp.ClearVals
    cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where active=1 AND Vender=0"
    Else
        cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' and active=1 AND Vender=0"
    End If
End Sub





Private Sub FillList()

    If cmbDept.MatchFound = False Then Exit Sub
    Dim rsList As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim ITM As ListItem
    
    Dim CanEdit As Boolean
    
    'With rsList
    '    .Open "Select DT from AttendanceSheet Where DT>'" & Format(DTPicker1, "dd-MMM-yyyy") & "'", Con, adOpenForwardOnly, adLockReadOnly
    '    If .RecordCount > 0 Then
    '        'Exit Sub Coz Can't Edit...
    '        Exit Sub
    '    End If
    '    Set rsList = Nothing
   ' End With
    
    Dim strCond As String
    With cmd
            .CommandType = adCmdStoredProc
            
    If cmbDept.Text = "<All Departments>" Then
        .CommandText = "VTakeAtt"
    Else
        .CommandText = "VTakeDeptAtt"
    End If
        
            
             
            
            Set .ActiveConnection = con
            .Parameters.Refresh
            
            
            
            .Parameters(1).Type = adDBTimeStamp
            .Parameters(1).value = Format(DTPicker1, "dd/MMM/yyyy")
            
        
            .Parameters(2).Type = adVarChar
            
            .Parameters(2).value = cmbShift.Text


            If cmbDept.Text <> "<All Departments>" Then
               .Parameters(3).Type = adVarChar
               .Parameters(3).value = cmbDept.ID
            End If
            
                           
            rsList.CursorLocation = adUseClient
            rsList.CursorType = adOpenForwardOnly
            rsList.LockType = adLockReadOnly
            Set rsList = .Execute
        End With
        
        
        'rsList.Open ""
        
        
        
    With rsList
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            Set ITM = LV.ListItems.Add(, ![EmpId], ![Name] & "", , ICO)
            ITM.SubItems(1) = ![FName] & ""
            ITM.SubItems(2) = ![EmpId] & ""
            ITM.SubItems(3) = cmbShift & ""
            If Not IsNull(![LateHours]) Then ITM.SubItems(4) = ![LateHours] & ""
            If IsNull(![DT]) Then
                ITM.Checked = False
            Else
                If ![Attendance] Then
                    ITM.Checked = True
                Else
                    ITM.Checked = False
                End If
                ITM.Tag = "Edit"
            End If
            'End If
            .MoveNext
        Next
        .Close
    End With
    
    Set rsLedger = Nothing
    
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

 fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2

 FLV.Move 50, 1230, Me.ScaleWidth - 100, Me.ScaleHeight - 1400
 
 LV.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
 
 fButs1.Move LV.Left + LV.Width - fButs1.Width, LV.Top + LV.Height
 
 fButs2.Move LV.Left, LV.Top + LV.Height
End Sub


Private Sub LV_DblClick()
If LV.ListItems.Count = 0 Then Exit Sub
    With txtLateHours
        .Move LV.Left + LV.ColumnHeaders(5).Left + 35, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(5).Width
        .Text = Round(Val(LV.SelectedItem.SubItems(4)) * 60)
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
    LV.Enabled = False
End Sub


Private Sub txtlatehours_KeyPress(KeyAscii As Integer)

On Error GoTo ERR
    If KeyAscii = 27 Then
        txtLateHours = ""
        txtLateHours.Visible = False
    ElseIf KeyAscii = 13 Then
         If Val(txtLateHours) > 240 Then
            MsgBox "Invalid Quantity", vbInformation
            txtLateHours.SelStart = 0
            txtLateHours.SelLength = Len(txtLateHours)
            txtLateHours.SetFocus
            Exit Sub
        Else
            LV.SelectedItem.SubItems(4) = Round(Val(txtLateHours) / 60, 2)
            'LV.SelectedItem.SubItems(5) = Val(txtLateHours) * Val(LV.SelectedItem.SubItems(2))
            txtLateHours.Visible = False
            'lblTotal = GetTotal
        End If
    Else
        OnlyNumber (KeyAscii)
    End If
    Exit Sub
ERR:
    MsgBox ERR.Description
    
End Sub

Private Sub txtLateHours_LostFocus()
    LV.Enabled = True
    txtLateHours.Visible = False
End Sub

Private Sub Usercontrol11_matched()

End Sub
