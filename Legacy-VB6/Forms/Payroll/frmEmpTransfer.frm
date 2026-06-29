VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmEmpTransfer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfer Employee"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9825
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
   ScaleWidth      =   9825
   StartUpPosition =   1  'CenterOwner
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   315
      Left            =   1755
      TabIndex        =   0
      Top             =   90
      Width           =   8000
      _ExtentX        =   14102
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
      Top             =   435
      Width           =   8000
      _ExtentX        =   14102
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
   Begin ComboList.Usercontrol1 cmbTransferDept 
      Height          =   315
      Left            =   1755
      TabIndex        =   2
      Top             =   765
      Width           =   8000
      _ExtentX        =   14102
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
      TabIndex        =   3
      Top             =   1305
      Width           =   8000
      _ExtentX        =   14102
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
      Format          =   117768195
      CurrentDate     =   37384
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Transfer to Dept. :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   375
      TabIndex        =   12
      Top             =   825
      Width           =   1365
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   1080
      Picture         =   "frmEmpTransfer.frx":0000
      Top             =   1980
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   360
      Picture         =   "frmEmpTransfer.frx":05CA
      Top             =   2100
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Description                :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   6
      Left            =   150
      TabIndex        =   11
      Top             =   1650
      Width           =   1575
   End
   Begin MSForms.TextBox txtDesc 
      Height          =   1005
      Left            =   1755
      TabIndex        =   4
      Top             =   1635
      Width           =   8000
      VariousPropertyBits=   -1467987941
      ScrollBars      =   3
      Size            =   "14111;1773"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   2775
      X2              =   9700
      Y1              =   2820
      Y2              =   2820
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   360
      Left            =   5175
      TabIndex        =   6
      Top             =   3075
      Width           =   2265
      Caption         =   "Save & Close"
      PicturePosition =   327683
      Size            =   "3995;635"
      Picture         =   "frmEmpTransfer.frx":0BE0
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
      Height          =   360
      Left            =   2880
      TabIndex        =   5
      Top             =   3075
      Width           =   2265
      Caption         =   "Save & New    "
      PicturePosition =   327683
      Size            =   "3995;635"
      Picture         =   "frmEmpTransfer.frx":0D3E
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
      Height          =   360
      Left            =   7490
      TabIndex        =   7
      Top             =   3075
      Width           =   2265
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "3995;635"
      Picture         =   "frmEmpTransfer.frx":0E50
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
      Caption         =   "Date of Transfer :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   435
      TabIndex        =   10
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Employee        :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   9
      Top             =   465
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Department    :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   150
      TabIndex        =   8
      Top             =   150
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   2775
      X2              =   9700
      Y1              =   2835
      Y2              =   2835
   End
End
Attribute VB_Name = "frmEmpTransfer"
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
     
    Dim rs As New ADODB.Recordset
    Dim bTempDept As Boolean
    With rs
        .Open "Select TempDept From Departments Where DeptID='" & cmbDept.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        bTempDept = !TempDept
        .Close
    End With
    Set rs = Nothing
    With cmbTransferDept
        .ClearVals
        .AddVals con, "Name", "Departments", "DeptID", " Where TempDept=" & Abs(bTempDept) & " AND DeptID<>'" & cmbDept.ID & "'"
    End With
    
End Sub
Private Function Saved() As Boolean

    Saved = False
    
    On Error GoTo err
    Dim SQL As String
    
    If Not cmbEmp.MatchFound Then
        MsgBox "Please Select Valid Employee From The List.", vbInformation
        Exit Function
    ElseIf cmbTransferDept.MatchFound = False Then
        MsgBox "Please Select Transfer Department.", vbInformation
        Exit Function
    End If
     
    Call StartTrans(con)
    
        SQL = "Insert Into EmpTransfers(EmpID,TransferDT,TransferFromDeptID,TransferToDeptID,EntryDescription," & _
         "UserName,MachineName) Values('" & cmbEmp.ID & "','" & DTPicker1 & "','" & cmbDept.ID & "','" & _
         cmbTransferDept.ID & "','" & txtDesc.Text & "','" & CurrentUserName & "','" & strComputerName & "')"
     
     con.Execute SQL
     
     SQL = "Update Employees Set DeptID='" & cmbTransferDept.ID & "' Where EmpId='" & cmbEmp.ID & "'"
     
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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If Saved Then
        Unload Me
        frmEmpTransfer.Show 1
    End If
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub




Private Sub Form_Load()

    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    cmbTransferDept.ListHeight = 1200
    'cmbType.AddItem "Short Term"
    'cmbType.AddItem "Long Term"
    'cmbType.ListIndex = 0
    DTPicker1.Value = Date
    cmbDept.ClearVals
    'cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID"
    'cmbDept.Text = "<All Departments>"
'    Dim rs As New ADODB.Recordset
'    With rs
'      .Open "Select Max(Dt) From MonthlySalaries Where EmpID='" & cmbEmp.ID & "'", con, adOpenForwardOnly, adLockReadOnly
'      If rs.RecordCount > 0 Then
'         If IsNull(.Fields(0)) Then
'               DTPicker1 = DateSerial(Year(GetServerDate), Month(GetServerDate), 1)
'         ElseIf Month(.Fields(0)) >= Month(GetServerDate) Then
'            DTPicker1 = DateAdd("m", 1, .Fields(0))
'         Else
'            DTPicker1 = DateSerial(Year(GetServerDate), Month(GetServerDate), 1)
'         End If
'      Else
'         DTPicker1 = DateSerial(Year(GetServerDate), Month(GetServerDate), 1)
'      End If
'      .Close
'
'    End With
'
'    Set rs = Nothing
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = OnlyNumber(KeyAscii.Value, True)
End Sub

Private Sub txtDeduct_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = OnlyNumber(KeyAscii.Value, True)
End Sub
