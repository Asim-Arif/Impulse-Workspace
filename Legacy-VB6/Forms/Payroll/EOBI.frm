VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form EOBI 
   ClientHeight    =   7110
   ClientLeft      =   435
   ClientTop       =   5895
   ClientWidth     =   11880
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   7110
   ScaleMode       =   0  'User
   ScaleWidth      =   10641.08
   WindowState     =   2  'Maximized
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   2610
      TabIndex        =   15
      Top             =   885
      Width           =   2265
      _ExtentX        =   3995
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
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   11670
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EOBI Sheet"
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
         TabIndex        =   7
         Top             =   120
         Width           =   1830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EOBI Sheet"
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
         TabIndex        =   8
         Top             =   135
         Width           =   1830
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7680
      Top             =   705
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
            Picture         =   "EOBI.frx":0000
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "EOBI.frx":0452
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin VB.Frame FLV 
      Height          =   5670
      Left            =   75
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
         TabIndex        =   16
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
            TabIndex        =   19
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
            TabIndex        =   18
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
            TabIndex        =   17
            Top             =   -30
            Width           =   945
         End
      End
      Begin VB.Frame fButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   885
         Left            =   6885
         TabIndex        =   9
         Top             =   4725
         Width           =   4665
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
            Left            =   3240
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   75
            Width           =   1410
         End
         Begin MSForms.CommandButton cmdSave 
            Height          =   375
            Left            =   45
            TabIndex        =   14
            Top             =   435
            Visible         =   0   'False
            Width           =   1500
            ForeColor       =   0
            Caption         =   "Save"
            PicturePosition =   327683
            Size            =   "2646;661"
            Accelerator     =   83
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdPrint 
            Height          =   375
            Left            =   1590
            TabIndex        =   12
            Top             =   435
            Width           =   1500
            ForeColor       =   0
            Caption         =   "Print         "
            PicturePosition =   327683
            Size            =   "2646;661"
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
            Left            =   3165
            TabIndex        =   13
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
            Caption         =   "Total EOBI Amount :"
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
            Left            =   1470
            TabIndex        =   11
            Top             =   105
            Width           =   1680
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   4515
         Left            =   90
         TabIndex        =   1
         Top             =   195
         Width           =   11460
         _ExtentX        =   20214
         _ExtentY        =   7964
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Emp ID"
            Object.Width           =   2646
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
            SubItemIndex    =   5
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Salary"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Days"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Net Salary"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "EOBI Amt"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   5415
      TabIndex        =   5
      Top             =   885
      Width           =   1500
      Caption         =   "Refresh       "
      PicturePosition =   327683
      Size            =   "2646;635"
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
      Left            =   120
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
      Left            =   1245
      TabIndex        =   2
      Top             =   885
      Width           =   1365
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "2408;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Menu mnupop 
      Caption         =   "mnupop"
      Begin VB.Menu mnuPrintEOBISheet 
         Caption         =   "Print EOBI Sheet"
      End
      Begin VB.Menu mnuNonEOBIDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintNonEOBIList 
         Caption         =   "Print Non EOBI List"
      End
   End
End
Attribute VB_Name = "EOBI"
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
If Loading Then Exit Sub
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

    
    On Error GoTo err
    
    If cmdSave.Visible Then
        If MsgBox("EOBI Sheet Must Be Saved Before You Can Print " & vbNewLine & _
        "Do You Want To Save It ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
        Call cmdSave_Click
    End If
    
    PopUpMenu mnupop, 255, cmdPrint.Left + FLV.Left + fButs1.Left, cmdPrint.Top + FLV.Top + fButs1.Top
    
Exit Sub

err:
    MsgBox err.Description
    
End Sub

Public Sub cmdRefresh_Click()
    
    If Not cmbYear.MatchFound Or Not cmbMonth.MatchFound Then Exit Sub
    Call FillList
    
End Sub

Private Sub cmdSave_Click()
    On Error GoTo err
    
    
    con.BeginTrans
    
    
    Cond = IIf(cmbDept.ID = "0", "", " and EmpID Like '" & cmbDept.ID & "%'")
    
    con.Execute "Delete From OldAgeMonthly Where Month(dt)=" & Month(Dated) & " and Year(dt)=" & Year(Dated) & Cond
    
    
    With LV
        For i = 1 To .ListItems.Count
            If .ListItems(i).Checked Then
                con.Execute "insert into OldAgeMonthly(EmpID,DT,EmpStatus,Salary," & _
                "DaysWorked,NetSalary,EOBIAmt) Values('" & _
                .ListItems(i).key & "','" & Dated & "'," & _
                IIf(.ListItems(i).SubItems(5) = "Permanent", 1, 0) & "," & _
                Val(.ListItems(i).SubItems(6)) & "," & Val(.ListItems(i).SubItems(7)) & _
                "," & Val(.ListItems(i).SubItems(8)) & "," & Val(.ListItems(i).SubItems(9)) & ")"
            End If
        Next i
    End With
    
    con.CommitTrans
    MsgBox "Saved Successfully", vbInformation
    cmdSave.Visible = False
    
    Exit Sub
err:
    MsgBox err.Description

End Sub



Private Sub DTPicker1_Change()
    Call FillList
End Sub

Private Sub Form_Activate()
Loading = False
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
    CYear = Format(GetServerDate, "yyyy")
    
    For i = 2002 To CYear
        cmbYear.AddItem i
    Next i
    
    cmbMonth = Format(GetServerDate, "MMM")
    cmbYear = CYear
    
    
    Call cmbMonth_Change
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID"
    cmbDept.Text = "<All Departments>"
    
    mnupop.Visible = False
    cmdSave.Visible = False
End Sub

Private Sub FillList()
    
    Dim rs As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim Total As Long, Cond As String
    
    
    Cond = IIf(cmbDept.ID = "0", "", " and DeptID='" & cmbDept.ID & "'")
    
    With rs
        .Open "Select * From  VOldAge Where Month(DT)=" & (cmbMonth.ListIndex + 1) & _
        " and Year(DT)=" & Val(cmbYear) & Cond, con, adOpenForwardOnly, adLockReadOnly
        
            LV.ListItems.Clear
            
            For i = 0 To .RecordCount - 1
                Set ITM = LV.ListItems.add(, ![EmpId], ![EmpId])
                ITM.SubItems(1) = ![Name] & ""
                ITM.SubItems(2) = ![FName] & ""
                ITM.SubItems(3) = ![OldAgeNo] & ""
                ITM.SubItems(4) = ![Designation] & ""
                ITM.SubItems(5) = IIf(![EmpStatus], "Permanent", "Temporary")
                ITM.SubItems(6) = Val(![Salary] & "")
                ITM.SubItems(7) = Val(![DaysWorked] & "")
                ITM.SubItems(8) = Format(Val(![NetSalary] & ""), "0.00")
                ITM.SubItems(9) = Format(Val(![EOBIAmt] & ""), "0.00")
                
                .MoveNext
            Next
            
        .Close
    End With
    
    Set rs = Nothing
    
    
    If LV.ListItems.Count = 0 Then
        If MsgBox("EOBI Sheet for the Selected Month has not been Calculated" & _
        vbNewLine & " Do you Want To Calculate It Now?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Else
        Call GetTotal
        Exit Sub
    End If
    
    'Get Social Security Employees Who Are Paid
    
    Dated = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
    iDays = DateAdd("m", 1, Dated) - Dated
    Dated = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), iDays)
    
    With cmd
        .CommandTimeout = 120
        .CommandType = adCmdStoredProc
        .CommandText = "SP_OldAgeMonthly"
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

    LV.ListItems.Clear
    With rs
        If .RecordCount > 0 Then
        
            Dim MDays As Integer
            
            Do Until .EOF
                Set ITM = LV.ListItems.add(, ![EmpId], ![EmpId])
                ITM.SubItems(1) = ![Name] & ""
                ITM.SubItems(2) = ![FName] & ""
                ITM.SubItems(3) = ![OldAgeNo] & ""
                ITM.SubItems(4) = ![Designation] & ""
                ITM.SubItems(5) = ![EmpStatus] & ""
                ITM.SubItems(6) = Val(![EmpSalary] & "")
                ITM.SubItems(7) = Val(![SalDays] & "")
                ITM.SubItems(8) = Format(Val(![SalAmt] & ""), "0.00")
                ITM.SubItems(9) = Format(Val(![EOBIAmt] & ""), "0.00")
                .MoveNext
            Loop
            
            'txtTotal = Total
            cmdPrint.Visible = True
        'Else
        '    Call ComputeIT
        End If
        
        cmdSave.Visible = True
        
    End With
    rs.Close
    Set rs = Nothing
    Call GetTotal
    
End Sub


Private Sub Form_Resize()
  
On Error Resume Next

 fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2

 FLV.Move 50, 1230, Me.ScaleWidth - 100, Me.ScaleHeight - 1400
 
 LV.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
 
 fButs1.Move LV.Left + LV.Width - fButs1.Width, LV.Top + LV.Height
 
End Sub



Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If Item.Checked Then
    Item.Checked = False
    Item.Selected = True
    Item.EnsureVisible
    
Else
    Item.SubItems(6) = ""
    Item.SubItems(7) = ""
    Item.SubItems(8) = ""
    Call GetTotal
End If
cmdSave.Enabled = True
End Sub


Private Sub mnuPrintSSForm_Click()
    Screen.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptOldAge.rpt")
    Dim f As New frmPrevRpt
    Dim SelFormula As String
    SelFormula = "Month({VOldAge.DT})=" & cmbMonth.List(cmbMonth.ListIndex, 1) & _
    " and Year({VOldAge.DT})=" & Val(cmbYear)
    
    f.ShowReport SelFormula, rpt
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub muuPrintSSSheet_Click()
    
    Screen.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptSocialSecurity.rpt")
    Dim f As New frmPrevRpt
    Dim SelFormula As String
    SelFormula = "Month({VSocialSecurity.DT})=" & cmbMonth.List(cmbMonth.ListIndex, 1) & _
    " and Year({VSocialSecurity.DT})=" & Val(cmbYear)
    
    f.ShowReport SelFormula & " ", rpt
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub mnuPrintEOBISheet_Click()

    Screen.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptOldAge.rpt")
    Dim f As New frmPrevRpt
    Dim SelFormula As String
    
    SelFormula = "Month({VOldAge.DT})=" & cmbMonth.List(cmbMonth.ListIndex, 1) & _
     " and Year({VOldAge.DT})=" & Val(cmbYear)
    
    f.ShowReport SelFormula & " ", rpt
    
    Screen.MousePointer = vbDefault
        
End Sub

Private Sub mnuPrintNonEOBIList_Click()

    Screen.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptNonEOBI.rpt")
    Dim f As New frmPrevRpt
    
    Dim strCond As String
    strCond = " Where Month(DT)=" & cmbMonth.ListIndex + 1 & " AND Year(DT)=" & cmbYear
    
'    On Error Resume Next
'    SrcReport.Database.Tables(1).ConnectBufferString = "DSN=DSN_Vision;;User ID=sa;;Database=Vision;;Password=visionasim"
'    rpt.SQLQueryString = "Select * From VEmp Where EmpID In(Select EmpID From MonthlySalaries " & strCond & ") AND EmpID Not In(Select EmpID From OldAgeMonthly " & strCond & ")"
    rpt.FormulaFields(2).Text = "'" & cmbMonth & " - " & cmbYear & "'"
    
    f.ShowReport "", rpt, "Select * From VEmp Where EmpID In(Select EmpID From MonthlySalaries " & strCond & ") AND EmpID Not In(Select EmpID From OldAgeMonthly " & strCond & ")"
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub txtEdit_Change(Index As Integer)

    If Index = 0 Then
        txtEdit(1) = Format(Val(LV.SelectedItem.SubItems(5)) * Val(txtEdit(0)) / Val(cmbMonth.Tag), "0.00")
    End If
    
End Sub

Private Sub txtedit_GotFocus(Index As Integer)
    txtEdit(Index).SelStart = 0
    txtEdit(Index).SelLength = Len(txtEdit(Index))
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo err
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
    
err:
    MsgBox err.Description
    
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
        If LV.ListItems(i).Checked Then Total = Total + Val(LV.ListItems(i).SubItems(9))
    Next
    txtTotal = Format(Total, "0.00")
    
End Sub
