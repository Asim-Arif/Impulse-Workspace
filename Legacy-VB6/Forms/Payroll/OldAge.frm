VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form OldAge 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11880
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleMode       =   0  'User
   ScaleWidth      =   10641.08
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   2625
      TabIndex        =   19
      Top             =   885
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
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11670
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Old Age"
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
         Left            =   4680
         TabIndex        =   11
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Old Age"
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
         Left            =   4695
         TabIndex        =   12
         Top             =   135
         Width           =   1275
      End
   End
   Begin VB.TextBox txtMin 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11220
      TabIndex        =   7
      Text            =   "5"
      Top             =   990
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtMax 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9870
      TabIndex        =   6
      Text            =   "28"
      Top             =   990
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5115
      Top             =   735
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
            Picture         =   "OldAge.frx":0000
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OldAge.frx":0452
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   5775
      Top             =   795
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
   Begin VB.Frame FLV 
      Height          =   5670
      Left            =   90
      TabIndex        =   0
      Top             =   1320
      Width           =   11670
      Begin VB.PictureBox PicEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1725
         Negotiate       =   -1  'True
         ScaleHeight     =   225
         ScaleWidth      =   3915
         TabIndex        =   20
         Top             =   735
         Visible         =   0   'False
         Width           =   3945
         Begin VB.TextBox txtedit 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
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
            Index           =   2
            Left            =   1890
            TabIndex        =   23
            Top             =   -45
            Width           =   945
         End
         Begin VB.TextBox txtedit 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
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
            Index           =   1
            Left            =   945
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   -30
            Width           =   945
         End
         Begin VB.TextBox txtedit 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
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
            Index           =   0
            Left            =   -15
            TabIndex        =   21
            Top             =   -30
            Width           =   945
         End
      End
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   930
         Left            =   6120
         TabIndex        =   13
         Top             =   4725
         Width           =   5250
         Begin VB.TextBox txtTotal 
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
            Left            =   3690
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   15
            Width           =   1410
         End
         Begin MSForms.CommandButton cmdSave 
            Height          =   375
            Left            =   45
            TabIndex        =   18
            Top             =   435
            Width           =   2055
            ForeColor       =   0
            VariousPropertyBits=   25
            Caption         =   "  Update Changes"
            PicturePosition =   327683
            Size            =   "3625;661"
            Picture         =   "OldAge.frx":0564
            Accelerator     =   83
            FontName        =   "Tahoma"
            FontEffects     =   1073750017
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdPrint 
            Height          =   375
            Left            =   2160
            TabIndex        =   16
            Top             =   435
            Width           =   1500
            ForeColor       =   0
            Caption         =   "Print         "
            PicturePosition =   327683
            Size            =   "2646;661"
            Picture         =   "OldAge.frx":0676
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
            Height          =   375
            Left            =   3720
            TabIndex        =   17
            Top             =   435
            Width           =   1500
            ForeColor       =   0
            Caption         =   " Close    "
            PicturePosition =   327683
            Size            =   "2646;661"
            Accelerator     =   67
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Total Amount :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   195
            Left            =   2400
            TabIndex        =   15
            Top             =   45
            Width           =   1245
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   4515
         Left            =   90
         TabIndex        =   1
         Top             =   195
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   7964
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Emp ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Emp Name"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "F Name"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Old Age #"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Designation"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Salary"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Days"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Net Salary"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Amount"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Min. Limit :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   10425
      TabIndex        =   9
      Top             =   1020
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Max Limit :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   8
      Left            =   9105
      TabIndex        =   8
      Top             =   1020
      Visible         =   0   'False
      Width           =   765
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   5370
      TabIndex        =   5
      Top             =   855
      Width           =   1500
      Caption         =   "Refresh       "
      PicturePosition =   327683
      Size            =   "2646;635"
      Picture         =   "OldAge.frx":0788
      Accelerator     =   82
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   630
      Width           =   4725
      BackColor       =   11517387
      Caption         =   "  Year                  Month                Departement"
      Size            =   "8334;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cmbYear 
      Height          =   300
      Left            =   150
      TabIndex        =   3
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
      Left            =   1290
      TabIndex        =   2
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
End
Attribute VB_Name = "OldAge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dated As Date
Dim Loading As Boolean
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

Private Sub cmbdept_matched()
Call cmdRefresh_Click
End Sub

Private Sub cmbMonth_Change()
    If Loading Or cmbMonth.MatchFound = False Then Exit Sub
    Dim iDays As Integer
    Dim DT As Date
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    iDays = DateAdd("m", 1, DT) - DT
    cmbMonth.Tag = iDays
    txtMax = iDays
    txtMin = "1"
End Sub

Private Sub cmbMonth_Click()
Call cmdRefresh_Click
End Sub

Private Sub cmbYear_Change()
    If Loading Then Exit Sub
    Dim iDays As Integer
    Dim DT As Date
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    iDays = DateAdd("m", 1, DT) - DT
    txtMax = iDays
    txtMin = "1"
End Sub

Private Sub cmdCancel_Click()
    cmdSave.Visible = False
    cmdCancel.Visible = False
    cmdRefresh.Visible = True
    cmdClose.Visible = True
    'cmdPrint.Visible = True
End Sub

Private Sub cmbYear_Click()
Call cmdRefresh_Click
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

On Error GoTo ERR
Call ShowWait

With cr1
    .ReportFileName = RptPRPath & "\OldAge.rpt"
    .Connect = ConnectStr
    .Formulas(0) = "Fromto='" & cmbMonth & "-" & cmbYear & "'"
    .SelectionFormula = "month({VOldAge.DT})=" & Month(Dated) & " and Year({VOldAGe.DT})=" & year(Dated)
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
    If Not cmbYear.MatchFound Or Not cmbMonth.MatchFound Then Exit Sub
    Call FillList
    cmdSave.Enabled = False
    'Call Form_Resize
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ERR
    
    
    con.BeginTrans
    
    
    Cond = IIf(cmbDept.ID = "0", "", " and EmpID Like '" & cmbDept.ID & "%'")
    
    con.Execute "Delete From OldAgeMonthly Where Month(dt)=" & Month(Dated) & " and Year(dt)=" & year(Dated) & Cond
    
    
    With LV
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked Then
                con.Execute "insert into OldAgeMonthly(EmpID," & _
                "DT,Salary,PerDeduct,DaysWorked) Values('" & _
                .ListItems(i).key & "','" & Dated & "'," & _
                .ListItems(i).SubItems(5) & "," & _
                .ListItems(i).SubItems(8) & "," & .ListItems(i).SubItems(6) & ")"
            End If
        Next i
    End With
    
    con.CommitTrans
    
    cmdSave.Enabled = False

    Exit Sub
ERR:
    MsgBox ERR.Description
End Sub



Private Sub DTPicker1_Change()
    Call FillList
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()


For i = 1 To LV.ColumnHeaders.Count
    LV.ColumnHeaders(i).Tag = LV.ColumnHeaders(i).Width
Next


    Loading = True
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
    cmbMonth = Format(ServerDate, "MMM")
    cmbYear = CYear
    Loading = False
    Call cmbMonth_Change
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID"
    cmbDept.Text = "<All Departments>"
    
End Sub

Private Sub FillList()
    
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim Total As Long, Cond As String
    
    Cond = IIf(cmbDept.ID = "0", "", " and DeptID='" & cmbDept.ID & "'")
    
    With rs
        .Open "Select EmpID,Name,FName,OldAgeNo,Designation,OldAgeSal From Employees Where OldAge=1 and Active=1" & Cond, con, adOpenForwardOnly, adLockReadOnly
            LV.ListItems.Clear
            For i = 0 To .RecordCount - 1
                Set ITM = LV.ListItems.Add(, ![EmpId], ![EmpId])
                ITM.SubItems(1) = ![Name] & ""
                ITM.SubItems(2) = ![FName] & ""
                ITM.SubItems(3) = ![OldAgeNo] & ""
                ITM.SubItems(4) = ![Designation] & ""
                ITM.SubItems(5) = Format(Val(![OldAgeSal] & ""), "0.00")
                .MoveNext
            Next
        .Close
    End With
    
    'Get Social Security Employees Who Are Paid
    
    Dated = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    iDays = DateAdd("m", 1, Dated) - Dated
    Dated = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), iDays)
    
    
    With cmd
        
        .CommandType = adCmdStoredProc
        .CommandText = "vOldAgeMonthly"
        Set .ActiveConnection = con
        .Parameters.Refresh
        .Parameters(1).Type = adDBTimeStamp
        .Parameters(1).value = Dated
        
        If cmbDept.ID <> "0" Then
            .Parameters(2).Type = adVarChar
            .Parameters(2).value = cmbDept.ID
        End If
        
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        
        Set rs = .Execute
    End With
    
    With rs
        If .RecordCount > 0 Then
            Dim MDays As Integer
            
            Do Until .EOF
                Set ITM = LV.ListItems(![EmpId] & "")
                ITM.Checked = True
                ITM.SubItems(5) = Val(![Salary] & "")
                ITM.SubItems(6) = Val(![DaysWorked] & "")
                MDays = DateAdd("m", 1, ![DT]) - ![DT]
                ITM.SubItems(7) = Format(Val(![Salary] & "") * Val(![DaysWorked] & "") / MDays, "0.00")
                ITM.SubItems(8) = Format(Val(![PerDeduct] & ""), "0.00")
                
                
                
                'itm.SubItems(8) = Round(Round((Val(![Salary]) / MDays)) * Val(![DaysWorked]))
                'itm.SubItems(9) = Round(Val(itm.SubItems(8)) * Val(![PerDeduct]))
                'Total = Total + Val(itm.SubItems(8))
                
                .MoveNext
            Loop
            
            'txtTotal = Total
            cmdPrint.Visible = True
        'Else
        '    Call ComputeIT
        End If
        
    End With
    rs.Close
    Set rs = Nothing
    Call GetTotal
End Sub
Private Sub ComputeIT()
    On Error GoTo ERR
    If MsgBox("Social Security For The Selected Month Has Not Been Computed." & vbNewLine & "Do You Want To Compute It...?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    Dim rsList As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim ITM As ListItem
    Dim Sql As String
    Dim Shift As String
    Dim Total As Long
    With cmd
        .CommandText = "Select * from Employees Where Social=1"
        Set .ActiveConnection = con
        .Parameters.Refresh
        '.Parameters(0).Type = adVarChar
        '.Parameters(0).value = cmbMonth
        '.Parameters(1).Type = adVarChar
        '.Parameters(1).value = cmbYear
        rsList.CursorType = adOpenForwardOnly
        rsList.LockType = adLockReadOnly
        Set rsList = .Execute
        LV.ListItems.Clear
    End With
    Dim PerSocial As Double
    Dim MonthDays As Integer
    
    Dim Sal As Double
    Dim iMax As Integer, iMin As Integer
    
    iMax = Val(txtMax)
    iMin = Val(txtMin)
    TDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    MonthDays = DateAdd("m", 1, TDate) - TDate
    
    Dim rsPer As New ADODB.Recordset
    With rsPer
        .Open "Select SocialSecurity from GeneralSettings", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            PerSocial = Val(.Fields(0)) / 100
        End If
        .Close
    End With
    Set rsPer = Nothing
    
    With rsList
        For i = 0 To .RecordCount - 1
            
            Set ITM = LV.ListItems.Add(, ![EmpId], ![EmpId])
            ITM.Tag = PerSocial
            ITM.SubItems(1) = ![Name] & ""
            ITM.SubItems(2) = ![FName] & ""
            ITM.SubItems(3) = ![SocialNo] & ""
            ITM.SubItems(4) = ![Designation] & ""
            ITM.SubItems(5) = ![Status] & ""
            Randomize
            
            ITM.SubItems(6) = Int(Rnd * (iMax - (iMin - 1)) + (iMin))
            Sal = Val(![SocialSal])
            ITM.SubItems(7) = Sal
            
            
            ITM.SubItems(8) = Round((Sal / MonthDays) * Val(ITM.SubItems(6)))
            ITM.SubItems(9) = Round(Val(ITM.SubItems(8)) * PerSocial)
            Total = Total + Val(ITM.SubItems(8))
            .MoveNext
        Next
        .Close
        txtTotal = Total
    End With
    Set rsLedger = Nothing
    Sql = ""
    
    cmdSave.Visible = True
    cmdCancel.Visible = True
    cmdClose.Visible = False
    cmdRefresh.Visible = False
    cmdPrint.Visible = False
    Exit Sub
ERR:
    MsgBox ERR.Description
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
 
 

    
End Sub

Private Sub LV_DblClick()
    If LV.ListItems.Count = 0 Then Exit Sub
    
    
    For i = 1 To LV.ColumnHeaders.Count
        LV.ColumnHeaders(i).Width = Val(LV.ColumnHeaders(i).Tag)
    Next

    
    With PicEdit
         LV.SelectedItem.EnsureVisible
         .Move LV.Left + LV.ColumnHeaders(7).Left + 35, LV.Top + LV.SelectedItem.Top + 30, LV.ColumnHeaders(9).Left + LV.ColumnHeaders(9).Width - LV.ColumnHeaders(7).Left + 30, LV.SelectedItem.Height
         txtEdit(0).Move 0, -15, LV.ColumnHeaders(7).Width
         txtEdit(0).Text = LV.SelectedItem.SubItems(6)
         
         txtEdit(1).Text = LV.SelectedItem.SubItems(7)
         txtEdit(1).Move txtEdit(0).Width, -15, LV.ColumnHeaders(8).Width
         
         txtEdit(2).Move txtEdit(1).Left + txtEdit(1).Width + 15, -15, LV.ColumnHeaders(9).Width - 15
         txtEdit(2).Text = LV.SelectedItem.SubItems(8)
         
        .Visible = True
        txtEdit(0).SetFocus
    End With
    
End Sub

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If Item.Checked Then
    Item.Checked = False
    Item.Selected = True
    Item.EnsureVisible
    Call LV_DblClick
Else
    Item.SubItems(6) = ""
    Item.SubItems(7) = ""
    Item.SubItems(8) = ""
    Call GetTotal
End If
cmdSave.Enabled = True
End Sub

Private Sub txtedit_Change(Index As Integer)
    If Index = 0 Then
        txtEdit(1) = Format(Val(LV.SelectedItem.SubItems(5)) * Val(txtEdit(0)) / Val(cmbMonth.Tag), "0.00")
    End If
    
End Sub

Private Sub txtedit_GotFocus(Index As Integer)
    txtEdit(Index).SelStart = 0
    txtEdit(Index).SelLength = Len(txtEdit(Index))
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo ERR
    If KeyAscii = 27 Then
        PicEdit.Visible = False
    ElseIf KeyAscii = 13 And Index = 2 Then
        If Val(txtEdit(0)) > Val(cmbMonth.Tag) Then
            MsgBox "Maximum Days Can Be " & cmbMonth.Tag, vbInformation
            txtEdit(0).SetFocus
            Exit Sub
        ElseIf Val(txtEdit(2)) = 0 Then
            MsgBox "Social Security Amount Cannot Be Zero.", vbInformation
            txtEdit(2).SetFocus
            Exit Sub
        Else
            With LV.SelectedItem
                If Not .Checked Then .Checked = True
                .SubItems(6) = Val(txtEdit(0))
                .SubItems(7) = Val(txtEdit(1))
                .SubItems(8) = Format(Val(txtEdit(2)), "0.00")
            End With
            cmdSave.Enabled = True
            Call GetTotal
        End If
    Else
        OnlyNumber (KeyAscii)
    End If
    Exit Sub
    
ERR:
    MsgBox ERR.Description
    
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
If LCase(Screen.ActiveControl.Name) <> "txtedit" Then PicEdit.Visible = False
End Sub

Private Sub txtMax_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumber(KeyAscii)
End Sub

Private Sub txtMax_Validate(cancel As Boolean)

    Dim iDays As Integer
    Dim DT As Date
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    iDays = DateAdd("m", 1, DT) - DT
    With txtMax
        If Val(.Text) > iDays Then
            MsgBox "Value Can't Be More Than The Days Of Selected Month.", vbInformation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            cancel = True
        ElseIf Val(.Text) < 2 Then
            MsgBox "Value Can't Be Less Than The 2.", vbInformation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            cancel = True
        ElseIf Val(.Text) < Val(txtMin) Then
            MsgBox "Value Can't Be Less Than Min Limit.", vbInformation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            cancel = True
        End If
    End With
    
End Sub

Private Sub txtMin_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumber(KeyAscii)
End Sub

Private Sub txtMin_Validate(cancel As Boolean)

    Dim iDays As Integer
    Dim DT As Date
    DT = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    iDays = DateAdd("m", 1, DT) - DT
    
    With txtMin
        If Val(.Text) < 1 Then
            MsgBox "Value Can't Be Less Than 1.", vbInformation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            cancel = True
        ElseIf Val(.Text) > (iDays - 1) Then
            MsgBox "Value Can't Be More Than The Days Of Selected Month.", vbInformation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            cancel = True
        ElseIf Val(.Text) > Val(txtMax.Text) Then
            MsgBox "Value Can't Be More Than Max Limit.", vbInformation
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            cancel = True
        End If
    End With
    
End Sub



Private Sub GetTotal()
Dim Total As Double
For i = 1 To LV.ListItems.Count
    If LV.ListItems(i).Checked Then Total = Total + Val(LV.ListItems(i).SubItems(8))
Next
txtTotal = Format(Total, "0.00")
End Sub
