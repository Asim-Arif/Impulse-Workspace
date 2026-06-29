VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmNewShortTerm 
   BorderStyle     =   0  'None
   ClientHeight    =   6915
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10425
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   ScaleHeight     =   6915
   ScaleMode       =   0  'User
   ScaleWidth      =   9337.813
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTAttend 
      Height          =   285
      Left            =   5205
      TabIndex        =   13
      Top             =   975
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   503
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
      CustomFormat    =   "MMMM, yyyy"
      Format          =   64880643
      CurrentDate     =   37055
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   435
      Top             =   5730
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
      Left            =   240
      TabIndex        =   1
      Top             =   975
      Width           =   2430
      _ExtentX        =   4286
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
      Left            =   2670
      TabIndex        =   0
      Top             =   975
      Width           =   2520
      _ExtentX        =   4445
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
   Begin VB.Frame FLV 
      Height          =   5550
      Left            =   240
      TabIndex        =   5
      Top             =   1290
      Width           =   9945
      Begin MSComCtl2.DTPicker DTLoan 
         Height          =   285
         Left            =   4695
         TabIndex        =   17
         Top             =   1635
         Visible         =   0   'False
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   503
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
         Format          =   64880643
         CurrentDate     =   37055
      End
      Begin VB.TextBox txtAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4410
         TabIndex        =   16
         Top             =   630
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Frame FButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   810
         Left            =   4905
         TabIndex        =   9
         Top             =   4620
         Width           =   4905
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
            Left            =   3465
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   90
            Width           =   1410
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
            Left            =   2115
            TabIndex        =   19
            Top             =   120
            Width           =   1245
         End
         Begin MSForms.CommandButton cmdPrint 
            CausesValidation=   0   'False
            Height          =   360
            Left            =   30
            TabIndex        =   14
            Top             =   420
            Width           =   1575
            ForeColor       =   0
            Caption         =   "Print Sheet   "
            PicturePosition =   327683
            Size            =   "2778;635"
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
            CausesValidation=   0   'False
            Height          =   360
            Left            =   3315
            TabIndex        =   11
            Top             =   420
            Width           =   1575
            ForeColor       =   0
            Caption         =   " Close"
            PicturePosition =   327683
            Size            =   "2778;635"
            Accelerator     =   67
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdSave 
            Height          =   360
            Left            =   1680
            TabIndex        =   10
            Top             =   420
            Width           =   1575
            ForeColor       =   0
            Caption         =   " Save Entry "
            PicturePosition =   327683
            Size            =   "2778;635"
            Picture         =   "frmNewShortTerm.frx":0000
            Accelerator     =   83
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Frame FButs2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   570
         Left            =   105
         TabIndex        =   6
         Top             =   4950
         Width           =   3000
         Begin MSForms.CommandButton cmdClearAll 
            Height          =   390
            Left            =   1470
            TabIndex        =   8
            Top             =   120
            Visible         =   0   'False
            Width           =   1485
            ForeColor       =   0
            Caption         =   "UnAuthorize All"
            PicturePosition =   327683
            Size            =   "2619;688"
            Accelerator     =   108
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdCheckAll 
            Height          =   360
            Left            =   0
            TabIndex        =   7
            Top             =   120
            Visible         =   0   'False
            Width           =   1425
            ForeColor       =   0
            Caption         =   "Authorize All"
            PicturePosition =   327683
            Size            =   "2514;635"
            Accelerator     =   104
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   4425
         Left            =   120
         TabIndex        =   12
         Top             =   150
         Width           =   9705
         _ExtentX        =   17119
         _ExtentY        =   7805
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
            Text            =   "EmpID"
            Object.Width           =   3969
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Employee Description"
            Object.Width           =   10504
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Prev Bal"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   255
      TabIndex        =   2
      Top             =   120
      Width           =   9915
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Short Term Loan"
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
         Width           =   3435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Short Term Loan"
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
         Width           =   3435
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7500
      Top             =   3120
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
            Picture         =   "frmNewShortTerm.frx":0112
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewShortTerm.frx":0564
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  Department :                            Employee Name :                       Month :"
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
      Height          =   225
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   720
      Width           =   7125
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete This Attendance"
      End
   End
End
Attribute VB_Name = "frmNewShortTerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Editing As Boolean
Dim forLeave As Boolean

Public Sub enterLateHours()
    Load LateHours
    LateHours.Tag = LV.SelectedItem
    LateHours.Show
End Sub

Private Sub cmbEmp_matched()

    On Error Resume Next
    If cmbEmp.ID = "0" Then Exit Sub
    LV.ListItems(cmbEmp.ID).Selected = True
    LV.ListItems(cmbEmp.ID).EnsureVisible

End Sub

Private Sub cmbShift_Change()
    Call FillList
End Sub

Private Sub cmdCheckAll_Click()
    
    
        For i = 1 To LV.ListItems.Count
            With LV.ListItems(i)
                .Checked = True
                .SubItems(2) = "OverTime"
            End With
        
        Next i
    
End Sub

Private Sub cmdClearAll_Click()

    For i = 1 To LV.ListItems.Count
        With LV.ListItems(i)
            .Checked = False
            .SubItems(2) = ""
        End With
    Next i

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub


Public Sub cmdRefresh_Click()
    'Call RefreshLedger
    
    'Call Form_Resize
End Sub


Private Sub cmdPrint_Click()

    Dim strSelFor As String

    Screen.MousePointer = vbHourglass

    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptShortTerm.rpt")
    
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

    Set FormulaFields = rpt.FormulaFields

    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & company & "'"
        ElseIf FormulaField.Name = "{@FromTo}" Then
            FormulaField.Text = "'" & Format(DTAttend, "mmmm, yyyy") & "'"
        End If
    Next

    strSelFor = "Year({Advances.DT})=" & DTAttend.year & " AND Month({Advances.DT})=" & DTAttend.Month
    

    f.ShowReport strSelFor & " ", rpt
    Screen.MousePointer = vbDefault


End Sub

Private Sub cmdSave_Click()

On Error GoTo err

    Call StartTrans(con)
    
    'con.Execute "DELETE FROM Advances WHERE  DT='" & DTAttend & "' AND Type=0"
    
    For i = 1 To LV.ListItems.Count
    
        With LV.ListItems(i)
        
            If .Checked = True Then
            
                'If .Tag = "" Then
                
                 con.Execute "Insert Into Advances(EmpID,DT,Type,Description,Amount) Values('" & _
                  .Text & "','" & CDate(.SubItems(3)) & "',0,'Short Term Loan'," & Val(.SubItems(4)) & ")"
                    
                'End If

            End If
            
        End With
        
    Next i
    
    con.CommitTrans
    Call FillList
Exit Sub

err:
    MsgBox err.Description, vbInformation

End Sub

Private Sub DTAttend_Change()
    
    Call FillList
    
End Sub

Private Sub DTLoan_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        LV.SelectedItem.SubItems(3) = DTLoan.value
        txtAmt.SetFocus
    ElseIf KeyCode = 27 Then
        DTLoan.Visible = False
        txtAmt.Visible = False
    End If
    
End Sub

Private Sub Form_Load()

    DTAttend = Format(GetServerDate, "dd-MMM-yyyy")
    
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1600
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID", " Where IsNull(TempDept,0)=0"
    cmbDept.ID = 0

    'DTAttend.Enabled = UserHasAccess("AttDateChange")
    
End Sub
Private Sub cmbdept_matched()

    Call FillList
    
    cmbEmp.ClearVals
    cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals con, "Name + ' {' + EmpID + '}'", "VSalariedEmps", "EmpID", "Where active=1"
    Else
        cmbEmp.AddVals con, "Name + ' {' + EmpID + '}'", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' and active=1 "
    End If
    cmbEmp.ID = "0"
    
End Sub

Private Sub FillList()

    
    If cmbDept.MatchFound = False Then Exit Sub
    
    LV.ListItems.Clear

    Dim rsList As New ADODB.Recordset
    Dim SQL As String
    Dim ITM As ListItem, Attend As String
    
    
    SQL = "SELECT dbo.VSalariedEmps.empid, dbo.VSalariedEmps.name,dbo.VSalariedEmps.Rel,dbo.VSalariedEmps.FName,SLoan.TotAmount " & _
     "FROM dbo.VSalariedEmps LEFT JOIN (Select EmpID,Sum(Amount) As TotAmount From Advances Where " & _
     "Type=0 AND Month(DT) =" & Month(DTAttend) & " AND Year(DT)=" & year(DTAttend) & _
     " Group By EmpID) As SLoan On VSalariedEmps.EmpID=SLoan.EmpID " & _
     "Where Active=1"
     
    If cmbDept.ID <> "0" Then
        SQL = SQL & " AND DeptID='" & cmbDept.ID & "'"
    End If
    Dim dTotAmt As Double
    With rsList
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
              
        For i = 1 To .RecordCount
            Set ITM = LV.ListItems.add(, !EmpId, !EmpId & "")
            ITM.ListSubItems.add , , !Name & " " & !Rel & " " & !FName
            
            ITM.ListSubItems.add , , Val(![TotAmount] & "")
            dTotAmt = dTotAmt + Val(![TotAmount] & "")
            'If IsNull(![Amount]) Then ITM.Tag = "Edit" Else ITM.Checked = True
            
            .MoveNext
        Next
    .Close
    
    txtTotal = dTotAmt
    
    Exit Sub
    
    
    End With
    Set rsList = Nothing

End Sub
Private Sub AddEmps(ITM As ListItem, rs As ADODB.Recordset, Shift As String)
    With rs
        Set ITM = LV.ListItems.add(, ![EmpId] & Shift, ![EmpId] & "", , ICO)
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
 
 LV.Move 100, 200, Me.Width - 400, FLV.Height - FButs1.Height - 300
 
 FButs1.Move LV.Left + LV.Width - FButs1.Width, LV.Top + LV.Height
 
 FButs2.Move LV.Left, LV.Top + LV.Height
End Sub


Private Sub LV_DblClick()
    
    If LV.ListItems.Count = 0 Then Exit Sub
    
    With DTLoan
        
        .Move LV.Left + LV.ColumnHeaders(4).Left + 50, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(4).Width
        If LV.SelectedItem.SubItems(3) <> "" Then
            .value = CDate(LV.SelectedItem.SubItems(3))
        Else
            .value = DTAttend.value
        End If
        
        .Visible = True
        .SetFocus
    End With
    
    With txtAmt
        .Move LV.Left + LV.ColumnHeaders(5).Left + 50, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(5).Width
        .Text = Val(LV.SelectedItem.SubItems(4))
        .Visible = True
'        .SelStart = 0
'        .SelLength = Len(.Text)
'        .SetFocus
    End With
    
End Sub

Private Sub LV_GotFocus()
    DTLoan.Visible = False
    txtAmt.Visible = False
End Sub

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    
    Dim clr As Long
    Item.Selected = True
    If Item.Checked Then
        Call LV_DblClick
    Else
        Item.SubItems(2) = ""
    End If
    
   
End Sub


Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Exit Sub
If Button = vbRightButton And LV.ListItems.Count > 0 Then
    PopUpMenu mnuPopup
    Exit Sub
End If
    'Set Flag that mouse is in attendance column or not
    forLeave = CBool(x >= LV.ColumnHeaders(3).Left And x <= (LV.ColumnHeaders(3).Left + LV.ColumnHeaders(3).Width))
    

End Sub

Private Sub mnuDelete_Click()

    con.Execute "DELETE FROM AttendanceSheet WHERE DT='" & DTAttend.value & "' AND EmpID='" & LV.SelectedItem.Text & "' AND Shift='" & cmbTime.Tag & "'"
    Call FillList

End Sub

Private Sub txtAmt_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        If Val(txtAmt) <= 0 Then
            MsgBox "Invalid Amount.", vbInformation
            txtAmt.SelStart = 0
            txtAmt.SelLength = Len(txtAmt)
            txtAmt.SetFocus
            Exit Sub
        Else
            LV.SelectedItem.SubItems(4) = Val(txtAmt)
            LV.SelectedItem.Checked = True
            txtAmt.Visible = False
            DTLoan.Visible = False
        End If
        
    ElseIf KeyAscii = 27 Then
        txtAmt.Visible = False
        DTLoan.Visible = False
        If Val(LV.SelectedItem.SubItems(4)) > 0 Then LV.SelectedItem.Checked = True Else LV.SelectedItem.Checked = False
    End If
    
End Sub
