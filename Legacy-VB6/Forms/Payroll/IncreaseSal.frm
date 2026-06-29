VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form IncreaseSal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salary Change"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   9255
   StartUpPosition =   1  'CenterOwner
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   315
      Left            =   1755
      TabIndex        =   0
      Top             =   30
      Width           =   7365
      _ExtentX        =   12991
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
      Height          =   315
      Left            =   1755
      TabIndex        =   1
      Top             =   375
      Width           =   7365
      _ExtentX        =   12991
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1755
      TabIndex        =   4
      Top             =   765
      Width           =   7365
      _ExtentX        =   12991
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
      CustomFormat    =   "dddd  dd-MMM yyyy"
      Format          =   116523011
      CurrentDate     =   37384
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   1080
      Picture         =   "IncreaseSal.frx":0000
      Top             =   1920
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   360
      Picture         =   "IncreaseSal.frx":05CA
      Top             =   2040
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Description                :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   6
      Left            =   150
      TabIndex        =   14
      Top             =   1590
      Width           =   1575
   End
   Begin MSForms.TextBox txtDesc 
      Height          =   1005
      Left            =   1755
      TabIndex        =   13
      Top             =   1575
      Width           =   7365
      VariousPropertyBits=   -1467987941
      ScrollBars      =   3
      Size            =   "12991;1773"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   2025
      X2              =   9100
      Y1              =   2820
      Y2              =   2820
   End
   Begin MSForms.TextBox txtAmt 
      Height          =   315
      Left            =   8070
      TabIndex        =   2
      Top             =   1140
      Width           =   1050
      VariousPropertyBits=   746604571
      Size            =   "1852;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox txtdeduct 
      Height          =   315
      Left            =   1755
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1155
      Width           =   990
      VariousPropertyBits=   746604575
      Size            =   "1746;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Current Salary :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   5
      Left            =   585
      TabIndex        =   11
      Top             =   1185
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "New Salary :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   7110
      TabIndex        =   10
      Top             =   1185
      Width           =   915
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   405
      Left            =   4410
      TabIndex        =   5
      Top             =   3000
      Width           =   2340
      Caption         =   "Save & Close"
      PicturePosition =   327683
      Size            =   "4128;714"
      Picture         =   "IncreaseSal.frx":0BE0
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdNext 
      Height          =   405
      Left            =   2040
      TabIndex        =   3
      Top             =   3000
      Width           =   2340
      Caption         =   "Save & New    "
      PicturePosition =   327683
      Size            =   "4128;714"
      Picture         =   "IncreaseSal.frx":0D3E
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   405
      Left            =   6780
      TabIndex        =   6
      Top             =   3000
      Width           =   2340
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "4128;714"
      Picture         =   "IncreaseSal.frx":0E50
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Salary Change  Date :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   9
      Top             =   780
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Employee        :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   8
      Top             =   420
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Department    :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   150
      TabIndex        =   7
      Top             =   90
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   2025
      X2              =   9100
      Y1              =   2835
      Y2              =   2835
   End
End
Attribute VB_Name = "IncreaseSal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbdept_matched()
    cmbEmp.ClearVals
    'cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "Employees", "EmpID"
    Else
        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' AND Active=1"
    End If
    
End Sub
Private Function Saved() As Boolean

    Saved = False
    
    On Error GoTo err
    Dim SQL As String
    
    If Not cmbEmp.MatchFound Then
        MsgBox "Please Select Valid Employee From The List.", vbInformation
        Exit Function
    ElseIf Val(txtAmt) = 0 Then
        MsgBox "Salary Cannot Be Zero.", vbInformation
        Exit Function
    ElseIf Val(txtdeduct) = 0 And txtdeduct.Visible Then
        MsgBox "External Salary Cannot Be Zero.", vbInformation
        Exit Function
    'ElseIf Val(txtdeduct) > Val(txtAmt) And txtdeduct.Visible Then
    '    MsgBox "Deduction Cannot Be More Than Loan Amount.", vbInformation
    '    Exit Function
    End If
    
    
    If LongTerm Then GoTo END_OF_CHECK
    
    
    
    
END_OF_CHECK:
    
    
    con.BeginTrans
    
        SQL = "INSERT INTO SalaryLedger(EmpID,DT,Description," & _
         "Salary,SalaryExt,UserName,MachineName,PreviousSalary) VALUES('" & cmbEmp.ID & _
         "','" & DTPicker1 & "','" & txtDesc & _
         "'," & Val(txtAmt) & "," & Val(txtdeduct) & ",'" & CurrentUserName & "','" & strComputerName & "'," & Val(txtdeduct) & ")"
     
     con.Execute SQL
     
     SQL = "Update Employees Set StartingSalary=" & Val(txtAmt) & ",Salary=" & Val(txtdeduct) & " Where EmpId='" & cmbEmp.ID & "'"
     
     con.Execute SQL
    con.CommitTrans
    
    Saved = True
    
    SQL = ""
    Exit Function
err:
    con.RollbackTrans
    MsgBox err.Description
End Function

Private Sub cmbType_Change()

End Sub

Private Sub cmbEmp_matched()
    txtdeduct = GetSingleLongValue("StartingSalary", "Employees", " WHERE EmpID='" & cmbEmp.ID & "'")
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If Saved Then
         Unload Me
        IncreaseSal.Show 1
    End If
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub

Private Sub Form_Load()

    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    
    'cmbType.AddItem "Short Term"
    'cmbType.AddItem "Long Term"
    'cmbType.ListIndex = 0
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID"
    cmbDept.Text = "<All Departments>"
    Dim rs As New ADODB.Recordset
    With rs
      .Open "Select Max(Dt) From MonthlySalaries Where EmpID='" & cmbEmp.ID & "'", con, adOpenForwardOnly, adLockReadOnly
      If rs.RecordCount > 0 Then
         If IsNull(.Fields(0)) Then
               DTPicker1 = DateSerial(year(GetServerDate), Month(GetServerDate), 1)
         ElseIf Month(.Fields(0)) >= Month(GetServerDate) Then
            DTPicker1 = DateAdd("m", 1, .Fields(0))
         Else
            DTPicker1 = DateSerial(year(GetServerDate), Month(GetServerDate), 1)
         End If
      Else
         DTPicker1 = DateSerial(year(GetServerDate), Month(GetServerDate), 1)
      End If
      .Close
      
    End With
    
    Set rs = Nothing
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = OnlyNumber(KeyAscii.Value, True)
End Sub

Private Sub txtDeduct_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = OnlyNumber(KeyAscii.Value, True)
End Sub
