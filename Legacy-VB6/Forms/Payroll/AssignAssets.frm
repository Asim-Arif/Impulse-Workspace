VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form AssignAssets 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assign Assets"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
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
   ScaleHeight     =   4410
   ScaleWidth      =   5985
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fTitle 
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5850
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assign Assets"
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
         Left            =   1980
         TabIndex        =   10
         Top             =   60
         Width           =   2040
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assign Assets"
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
         Left            =   1950
         TabIndex        =   11
         Top             =   60
         Width           =   2040
      End
   End
   Begin VB.ComboBox cmbType 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "AssignAssets.frx":0000
      Left            =   60
      List            =   "AssignAssets.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1410
      Width           =   2925
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Top             =   840
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   2970
      TabIndex        =   1
      Top             =   840
      Width           =   2910
      _ExtentX        =   5133
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin MSComCtl2.DTPicker DTFrom 
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   2010
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
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
      Format          =   66650115
      CurrentDate     =   37384
   End
   Begin MSComCtl2.DTPicker DTFor 
      Height          =   315
      Left            =   2970
      TabIndex        =   12
      Top             =   2010
      Width           =   2925
      _ExtentX        =   5159
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
      Format          =   66650115
      CurrentDate     =   37384
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Asset Description"
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
      Index           =   5
      Left            =   60
      TabIndex        =   16
      Top             =   2370
      Width           =   5805
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Assign Date :                                         Assigned For :"
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
      Index           =   3
      Left            =   60
      TabIndex        =   15
      Top             =   1770
      Width           =   5805
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Assign Type :                                         Asset Amount :"
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
      Left            =   60
      TabIndex        =   14
      Top             =   1170
      Width           =   5805
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Department :                                       Employee :"
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
      Index           =   4
      Left            =   60
      TabIndex        =   13
      Top             =   600
      Width           =   5805
   End
   Begin MSForms.TextBox txtAmount 
      Height          =   315
      Left            =   2970
      TabIndex        =   8
      Top             =   1410
      Width           =   2880
      VariousPropertyBits=   746604571
      Size            =   "5080;556"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtDesc 
      Height          =   1005
      Left            =   60
      TabIndex        =   6
      Top             =   2610
      Width           =   5820
      VariousPropertyBits=   -1467987941
      ScrollBars      =   3
      Size            =   "10266;1773"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   75
      X2              =   5880
      Y1              =   3810
      Y2              =   3810
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   360
      Left            =   2685
      TabIndex        =   4
      Top             =   3960
      Width           =   1620
      Caption         =   "Save & Close"
      PicturePosition =   327683
      Size            =   "2857;635"
      Picture         =   "AssignAssets.frx":0024
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
      Left            =   1065
      TabIndex        =   3
      Top             =   3960
      Width           =   1590
      Caption         =   "Save & New    "
      PicturePosition =   327683
      Size            =   "2805;635"
      Picture         =   "AssignAssets.frx":0182
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
      Left            =   4320
      TabIndex        =   5
      Top             =   3960
      Width           =   1530
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2699;635"
      Picture         =   "AssignAssets.frx":0294
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   75
      X2              =   5850
      Y1              =   3825
      Y2              =   3825
   End
End
Attribute VB_Name = "AssignAssets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public edit As Boolean

Private Sub cmbdept_matched()
    cmbEmp.ClearVals
    'cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where active=1"
    Else
        cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' and Active=1"
    End If
    
End Sub
Private Function Saved() As Boolean

Saved = False






    'On Error GoTo ERR
    Dim SQL As String
    
    If Not cmbEmp.MatchFound Then
        MsgBox "Please Select Valid Employee From The List.", vbInformation
        Exit Function
    ElseIf Val(txtAmount) = 0 Then
        MsgBox "Amount Cannot Be Zero.", vbInformation
        Exit Function
    'ElseIf Val(txtdeduct) > Val(txtAmt) And txtdeduct.Visible Then
    '    MsgBox "Deduction Cannot Be More Than Loan Amount.", vbInformation
    '    Exit Function
    End If
    
    
'    If LongTerm Then GoTo END_OF_CHECK
    
    
    
    
'END_OF_CHECK:
    
    
    con.BeginTrans
    If edit = True Then
    
     SQL = "Update EmpAssets Set EmpID='" & cmbEmp.ID & "',DT='" & DtFrom _
        & "',Description='" & txtDesc & "',Type='" & cmbType & "' ,Amount=" & txtAmount _
        & ", AssignFor='" & DTFor & "' Where AssetID=" & Me.Tag & ""
    Else
    
        SQL = "Insert Into EmpAssets(EmpID,DT,Description," & _
         "Type,Amount,AssignFor) Values('" & cmbEmp.ID & _
         "','" & DtFrom & "','" & txtDesc & _
         "','" & cmbType & "'," & Val(txtAmount) & ",'" & DTFor & "')"
     End If
     
     con.Execute SQL
     
    con.CommitTrans
    edit = False
    Saved = True
    
    SQL = ""
    Exit Function
err:
    con.RollbackTrans
    MsgBox err.Description
End Function

Private Sub cmbType_Click()

If cmbType.ListIndex = 0 Then
    DTFor.Enabled = False
Else
    DTFor.Enabled = True
End If

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If Saved Then
         Unload Me
         If LongTerm Then
            MainForm.mnuLong_Click
         Else
            MainForm.mnuShort_Click
         End If
    End If
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub




Private Sub Form_Activate()

Dim Rec As New ADODB.Recordset

If edit = True Then
    With Rec
        .CursorLocation = adUseClient
        .Open "SELECT * FROM EmpAssets WHERE AssetID=" & Me.Tag & "", con, adOpenKeyset, adLockReadOnly
        If .RecordCount > 0 Then
        
            cmbEmp.ID = !EmpId
            cmbType = !Type
            txtAmount = !Amount
            txtDesc = !Description
            DtFrom = ![DT]
            DTFor = !AssignFor
        End If
        
    End With
    
End If
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
    cmbType.ListIndex = 0
    
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = OnlyNumber(KeyAscii.value, True)
End Sub

Private Sub txtDeduct_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = OnlyNumber(KeyAscii.value, True)
End Sub
