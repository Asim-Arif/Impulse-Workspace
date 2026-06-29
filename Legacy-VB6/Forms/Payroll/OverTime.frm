VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form OverTime 
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   7380
   ScaleMode       =   0  'User
   ScaleWidth      =   10667.95
   WindowState     =   2  'Maximized
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11385
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Over Time"
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
         Index           =   0
         Left            =   60
         TabIndex        =   7
         Top             =   120
         Width           =   11235
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Over Time"
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
         Index           =   1
         Left            =   60
         TabIndex        =   8
         Top             =   135
         Width           =   11235
      End
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   1470
      TabIndex        =   2
      Top             =   840
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
      Left            =   3690
      TabIndex        =   3
      Top             =   840
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
            Picture         =   "OverTime.frx":0000
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OverTime.frx":0452
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker DtTo 
      Height          =   300
      Left            =   7320
      TabIndex        =   9
      Top             =   855
      Width           =   1395
      _ExtentX        =   2461
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
      Format          =   60751875
      CurrentDate     =   37081
   End
   Begin MSComCtl2.DTPicker DtFrom 
      Height          =   300
      Left            =   5940
      TabIndex        =   10
      Top             =   855
      Width           =   1335
      _ExtentX        =   2355
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
      Format          =   60751875
      CurrentDate     =   37055
   End
   Begin VB.Frame FLV 
      Height          =   5775
      Left            =   15
      TabIndex        =   0
      Top             =   1260
      Width           =   11385
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   480
         Left            =   6555
         TabIndex        =   12
         Top             =   5220
         Width           =   4635
         Begin MSForms.CommandButton cmdPrint 
            Height          =   360
            Left            =   60
            TabIndex        =   15
            Top             =   60
            Width           =   1425
            ForeColor       =   0
            Caption         =   "Print         "
            PicturePosition =   327683
            Size            =   "2514;635"
            Picture         =   "OverTime.frx":0564
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
            Left            =   3105
            TabIndex        =   14
            Top             =   60
            Width           =   1425
            ForeColor       =   0
            Caption         =   " Close      "
            PicturePosition =   327683
            Size            =   "2514;635"
            Picture         =   "OverTime.frx":0676
            Accelerator     =   67
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdNew 
            Height          =   360
            Left            =   1605
            TabIndex        =   13
            Top             =   60
            Width           =   1425
            ForeColor       =   0
            Caption         =   "New Entry   "
            PicturePosition =   327683
            Size            =   "2514;635"
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
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   5130
         TabIndex        =   11
         Top             =   2100
         Visible         =   0   'False
         Width           =   1155
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
         Height          =   5040
         Left            =   -45
         TabIndex        =   1
         Top             =   105
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   8890
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
            Text            =   "Emp ID"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emp Name"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "F Name"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date"
            Object.Width           =   4322
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Over Time"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   1500
      TabIndex        =   5
      Top             =   600
      Width           =   7185
      BackColor       =   11517387
      Caption         =   "  Department                         Employee                               From                     To"
      Size            =   "12674;423"
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
      Left            =   8805
      TabIndex        =   4
      Top             =   825
      Width           =   1500
      Caption         =   "Refresh        "
      PicturePosition =   327683
      Size            =   "2646;635"
      Picture         =   "OverTime.frx":0788
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
Attribute VB_Name = "OverTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
    Load NewOverTime
    
    If cmbEmp.Text <> "<All Employees>" Then
        NewOverTime.cmbDept.Text = cmbDept.Text
        NewOverTime.cmbEmp.Text = cmbEmp.Text
    Else
        If cmbDept.ID = 0 Then
            NewOverTime.cmbDept.Text = "<All Departments>"
        End If
    End If
    
    NewOverTime.Show 1
    Call RefreshLedger
    
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo ERR
    Call ShowWait
    Dim Cond As String
    
    
    
    If cmbEmp.ID = "0" Then
        If cmbDept.ID = "0" Then
            Cond = "({VEmpOverTime.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & "))"
        Else
            Cond = "({VEmpOverTime.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & ")) and {VEmpOverTime.DeptID}='" & cmbDept.ID & "'"
        End If
    Else
        Cond = "{VEmpOverTime.empid}='" & cmbEmp.ID & "' AND ({VEmpOverTime.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & "))"
    End If
    
With cr1
    .ReportFileName = RptPRPath & "\OverTime.rpt"
    '.DataFiles(0) = rptPRPath & "\Payroll.mdb"
    .Connect = ConnectStr
    .SelectionFormula = Cond
    
    .Formulas(0) = "FromTo='From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
    
    
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

    DtTo = ServerDate
    
    DtFrom = DateAdd("m", -1, ServerDate)
    
    cmbYear = CYear
    
    cmbMonth = Format(ServerDate, "MMM")
    
    cmbDept.ListHeight = 1600
    
    cmbEmp.ListHeight = 1250
    
    cmbDept.ClearVals
    
    cmbDept.AddItem "<All Departments>", "0"
    
    cmbDept.AddVals con, "Name", "Departments", "DeptID"
    
    cmbDept.Text = "<All Departments>"
    
    cmbEmp.Text = "<All Employees>"
    
    cmdNew.Enabled = UserHasAccess("AddOverTime")
   
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

    If cmbDept.MatchFound = False Or cmbEmp.MatchFound = False Then Exit Sub
    Dim rsLedger As New ADODB.Recordset
    
    Dim ITM As ListItem
    Dim Cond As String
    Dim Sql As String
    'Dim ICO As String
    
    
    
    If cmbEmp.ID = "0" Then
        If cmbDept.ID = "0" Then
            Cond = " where DT Between '" & Format(DtFrom, "dd-MMM-yyyy") & "' and '" & Format(DtTo, "dd-MMM-yyyy") & "'"
        Else
            Cond = " Where DeptID='" & cmbDept.ID & "' and DT Between '" & Format(DtFrom, "dd-MMM-yyyy") & "' and '" & Format(DtTo, "dd-MMM-yyyy") & "'"
        End If
    Else
        Cond = " Where empid='" & cmbEmp.ID & "' and DT Between '" & Format(DtFrom, "dd-MMM-yyyy") & "' and '" & Format(DtTo, "dd-MMM-yyyy") & "'"
    End If
    
    Sql = "select * from VEmpOverTime " & Cond & " Order By EmpID,DT Asc"
    
    With rsLedger
        .Open Sql, con, adOpenForwardOnly, adLockReadOnly
        LV.Visible = False
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            
            Set ITM = LV.ListItems.Add(, ![EmpId] & Format(![DT] & "", "dd-MMM-yyyy") & Rnd, ![EmpId] & "")
            ITM.Tag = ![DT]
            ITM.SubItems(1) = ![Name] & ""
            ITM.SubItems(2) = ![FName] & ""
            ITM.SubItems(3) = Format(![DT], "dddd dd MMMM ,yyyy")
            ITM.SubItems(4) = ![HOURS] & ""
            
            .MoveNext
            
        Next
        
        LV.Visible = True
        .Close
        
    End With
    
    Set rsLedger = Nothing
    
    Sql = ""
    Cond = ""
    
End Sub

Private Sub LV_DblClick()
If LV.ListItems.Count = 0 Or UserHasAccess("EditOverTime") = False Then Exit Sub
If CDate(LV.SelectedItem.Tag) <> ServerDate Then
    MsgBox "You Can Edit Overtime For Today Only.", vbInformation
    Exit Sub
End If
    With txtEdit
    .Move LV.Left + LV.ColumnHeaders(5).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(5).Width
    .Text = LV.SelectedItem.SubItems(4)
    .Visible = True
    .SetFocus
End With
End Sub



Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyDelete And LV.ListItems.Count > 0 Then
    If MsgBox("Delete This Entry ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    con.Execute "Delete From OverTime Where EmpID='" & LV.SelectedItem & "' and Dt='" & Format(LV.SelectedItem.Tag, "dd-MMM-yyyy") & "'"
    Call LV.ListItems.Remove(LV.SelectedItem.Index)
End If

End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    con.Execute "Update OverTime Set Hours='" & Val(txtEdit) & "' Where EmpID='" & LV.SelectedItem & "' and Dt='" & Format(LV.SelectedItem.Tag, "dd-MMM-yyyy") & "'"
    LV.SelectedItem.SubItems(4) = Val(txtEdit)
    LV.SetFocus
ElseIf KeyAscii = 27 Then
    LV.SetFocus
Else
    KeyAscii = OnlyNumber(KeyAscii, True)
End If

End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub
