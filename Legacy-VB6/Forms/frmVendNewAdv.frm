VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmVendNewAdv 
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
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   330
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
      Height          =   330
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
      Format          =   51314691
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
      TabIndex        =   13
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
      Y1              =   2820
      Y2              =   2820
   End
   Begin MSForms.TextBox txtAmt 
      Height          =   315
      Left            =   1785
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   1200
      Width           =   1605
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   360
      Left            =   1755
      TabIndex        =   4
      Top             =   3000
      Width           =   1530
      Caption         =   "Save & Close"
      PicturePosition =   327683
      Size            =   "2699;635"
      Picture         =   "frmVendNewAdv.frx":0000
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
      Left            =   135
      TabIndex        =   3
      Top             =   3000
      Width           =   1530
      Caption         =   "Save & New    "
      PicturePosition =   327683
      Size            =   "2699;635"
      Picture         =   "frmVendNewAdv.frx":015E
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
      Left            =   3390
      TabIndex        =   5
      Top             =   3000
      Width           =   1530
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2699;635"
      Picture         =   "frmVendNewAdv.frx":0270
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      TabIndex        =   6
      Top             =   90
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   195
      X2              =   4895
      Y1              =   2835
      Y2              =   2835
   End
End
Attribute VB_Name = "frmVendNewAdv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LongTerm As Boolean

Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    'cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Processes>" Then
        cmbEmp.AddVals con, "VenderName", "Venders", "VendID"
    Else
        cmbEmp.AddVals con, "VenderName", "Venders", "VendID", "Where PhaseID=" & cmbDept.ID
    End If
    
End Sub
Private Function Saved() As Boolean

    
    On Error GoTo err
    Saved = False
    Dim SQL As String
    
    If Not cmbEmp.MatchFound Then
        MsgBox "Please Select Valid Vender From The List.", vbInformation
        Exit Function
    ElseIf Val(txtAmt) = 0 Then
        MsgBox "Invalid Amount.", vbInformation
        Exit Function
    'ElseIf Val(txtDeduct) = 0 And txtDeduct.Visible Then
    '    MsgBox "Invalid Deduction Amount.", vbInformation
    '    Exit Function
    ElseIf Val(txtDeduct) > Val(txtAmt) And txtDeduct.Visible Then
        MsgBox "Deduction Cannot Be More Than Loan Amount.", vbInformation
        Exit Function
    End If
    
    
    If LongTerm Then GoTo END_OF_CHECK
    'Dim AmtAllowed As Double
    'AmtAllowed = GetShortAmt(cmbEmp.ID, DTPicker1.value)
    'If Val(txtAmt) > AmtAllowed Then
    '  MsgBox "Cannot Pay More Than Rs. " & Format(AmtAllowed, "0.00")
    '  Exit Function
    'End If
    
END_OF_CHECK:
    
    
    
    SQL = "Insert Into VenderAdvances(VendID,DT,Type,Description," & _
     "Amount,DAmount,AmountCleared) Values('" & cmbEmp.ID & _
     "','" & DTPicker1 & "','" & Abs(LongTerm) & "','" & txtDesc & _
     "'," & Val(txtAmt) & "," & Val(txtDeduct) & ",0)"
     
    con.Execute SQL
    
    Saved = True
    
    SQL = ""
    Exit Function
err:
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
         If LongTerm Then
            Call MainForm.mnuVendAdvLong_Click
         Else
            Call MainForm.mnuVendAdvShort_Click
         End If
    End If
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub


Private Sub Form_Activate()
   
    Label1(5).Visible = LongTerm
    txtDeduct.Visible = LongTerm
    
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
    cmbDept.AddItem "<All Processes>", "0"
    cmbDept.AddVals con, "Description", "Processes", "ProcessID"
    cmbDept.Text = "<All Processes>"
      
    DTPicker1 = ServerDate
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    OurOwnForm = True
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = OnlyNumber(KeyAscii.value, True)
End Sub

Private Sub txtDeduct_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = OnlyNumber(KeyAscii.value, True)
End Sub
