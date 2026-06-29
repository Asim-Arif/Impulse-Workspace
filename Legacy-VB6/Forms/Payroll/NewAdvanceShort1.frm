VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form NewAdvanceShort1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advances..."
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
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
   ScaleWidth      =   5100
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkPrint 
      Caption         =   "Print When Saved"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   225
      TabIndex        =   15
      Top             =   2640
      Width           =   1800
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   1755
      TabIndex        =   0
      Top             =   30
      Width           =   3120
      _ExtentX        =   5503
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
      Left            =   1755
      TabIndex        =   1
      Top             =   375
      Width           =   3120
      _ExtentX        =   5503
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
      Height          =   315
      Left            =   1785
      TabIndex        =   2
      Top             =   765
      Width           =   3075
      _ExtentX        =   5424
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
      CustomFormat    =   "dddd dd-MMM-yyyy"
      Format          =   57147395
      CurrentDate     =   37384
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
      Left            =   1770
      TabIndex        =   5
      Top             =   1575
      Width           =   3090
      VariousPropertyBits=   -1467987941
      ScrollBars      =   3
      Size            =   "5450;1773"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   195
      X2              =   4895
      Y1              =   2940
      Y2              =   2940
   End
   Begin MSForms.TextBox txtAmt 
      Height          =   315
      Left            =   1785
      TabIndex        =   3
      Top             =   1155
      Width           =   1050
      VariousPropertyBits=   746604571
      Size            =   "1852;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox txtDeduct 
      Height          =   315
      Left            =   3810
      TabIndex        =   4
      Top             =   1155
      Width           =   1050
      VariousPropertyBits=   746604571
      Size            =   "1852;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ded. Amt :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   5
      Left            =   2940
      TabIndex        =   13
      Top             =   1215
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Amount to Issue        :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   1605
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   360
      Left            =   1845
      TabIndex        =   7
      Top             =   3000
      Width           =   1710
      Caption         =   "Save & Close"
      PicturePosition =   327683
      Size            =   "3016;635"
      Picture         =   "NewAdvanceShort1.frx":0000
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
      Left            =   195
      TabIndex        =   6
      Top             =   3000
      Width           =   1590
      Caption         =   "Save & New    "
      PicturePosition =   327683
      Size            =   "2805;635"
      Picture         =   "NewAdvanceShort1.frx":015E
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
      Left            =   3600
      TabIndex        =   8
      Top             =   3000
      Width           =   1380
      Caption         =   "Cancel    "
      PicturePosition =   327683
      Size            =   "2434;635"
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
      Caption         =   "Issue Date                 :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   780
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Employee        :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   90
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   195
      X2              =   4895
      Y1              =   2955
      Y2              =   2955
   End
End
Attribute VB_Name = "NewAdvanceShort1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LongTerm As Boolean

Private Sub PrintReport()

Screen.MousePointer = vbHourglass

Dim f As New frmPrevRpt, rpt As New rptAdvancesPayslip, ID As Integer

ID = Val(con.Execute("SELECT MAX(EntryID) FROM Advances").Fields(0) & "")

Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

Set FormulaFields = rpt.FormulaFields
For Each FormulaField In FormulaFields
    If FormulaField.Name = "{@Company}" Then
        FormulaField.Text = "'" & company & "'"
    ElseIf FormulaField.Name = "{@FromTo}" Then
        FormulaField.Text = "'From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
    End If
Next

Me.Hide
f.ShowReport "{VAdvances.EntryID}=" & ID & " ", rpt


Screen.MousePointer = vbDefault


End Sub

Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    'cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where Active=1"
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
    ElseIf Val(txtAmt) = 0 Then
        MsgBox "Invalid Amount.", vbInformation
        Exit Function
    ElseIf Val(txtdeduct) = 0 And txtdeduct.Visible Then
        MsgBox "Invalid Deduction Amount.", vbInformation
        Exit Function
        
   
         
        
        
    ElseIf Val(txtdeduct) > Val(txtAmt) And txtdeduct.Visible Then
        MsgBox "Deduction Cannot Be More Than Loan Amount.", vbInformation
        Exit Function
    End If
    
 
    '''''''''''''''''''''''''''''''''
    ' Remove This Check at the moment
    
    
'    If LongTerm = True Then
'        If IsNotLongApplicable Then
'            Exit Function
'        End If
'    End If
'
    
    
    '''''''''''''''''''''''''''''''''
    ' Remove This Check at the moment
    

    
    
    
    
'    Dim AmtAllowed As Double
'    AmtAllowed = GetShortAmt(cmbEmp.ID, DTPicker1.value)
'    If Val(txtAmt) > AmtAllowed Then
'      MsgBox "Cannot Pay More Than Rs. " & Format(AmtAllowed, "0.00"), vbInformation
'      Exit Function
'    End If
'

    
    
    
    SQL = "Insert Into Advances(EmpID,DT,Type,Description," & _
     "Amount,DAmount,AmountCleared) Values('" & cmbEmp.ID & _
     "','" & DTPicker1 & "','" & Abs(LongTerm) & "','" & txtDesc & _
     "'," & Val(txtAmt) & "," & Val(txtdeduct) & ",0)"
     
    con.Execute SQL
    
    Saved = True
    If chkPrint = 1 Then Call PrintReport
    SQL = ""
    Exit Function
ERR:
    MsgBox ERR.Description


End Function

Private Sub cmbType_Change()

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
   
   Label1(5).Visible = LongTerm
   txtdeduct.Visible = LongTerm
   
If LongTerm Then
   txtDesc = "Long Term Loan"
   Me.Caption = "Long Term Loan"
Else
   txtDesc = "Short Term Loan"
   Me.Caption = "Short Term Loan"
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
      
    DTPicker1 = GetServerDate
    
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = OnlyNumber(KeyAscii.value, True)
End Sub

Private Sub txtDeduct_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = OnlyNumber(KeyAscii.value, True)
End Sub

'To Check The Long Term Settings That Are Made For Departments
Private Function IsNotLongApplicable() As Boolean

Dim Rec As New ADODB.Recordset
'dim Item as New

With Rec
    .CursorLocation = adUseClient
    .Open "SELECT JoinDate,LongTermTime,LongTermLimit FROM VOnlyEmps WHERE EmpID='" & cmbEmp.ID & "'", con, adOpenKeyset, adLockReadOnly
    
    If DateDiff("m", !JoinDate, DTPicker1) < !longtermtime Then
        MsgBox "Could Not Give Long Term Loan To Employee." & vbNewLine & "It Can Be Given After " & !longtermtime & " Months", vbInformation
        IsNotLongApplicable = True
    ElseIf Val(txtAmt) > !LongTermLimit Then
        MsgBox "Could Not Give Long Term Loan To Employee. Because Amount " & vbNewLine & " Is Greater Than Long Term Limit.", vbInformation
        IsNotLongApplicable = True
    End If
    
End With


End Function

