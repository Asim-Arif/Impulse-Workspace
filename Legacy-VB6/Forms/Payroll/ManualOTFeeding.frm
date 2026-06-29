VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmManualOTFeeding 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manual Over Time Feeding"
   ClientHeight    =   1845
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
   ScaleHeight     =   1845
   ScaleWidth      =   5100
   StartUpPosition =   1  'CenterOwner
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
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   90
      X2              =   4790
      Y1              =   1230
      Y2              =   1230
   End
   Begin MSForms.TextBox txtOT 
      Height          =   315
      Left            =   3810
      TabIndex        =   8
      Top             =   765
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
      Caption         =   "Total Hours :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   5
      Left            =   2850
      TabIndex        =   7
      Top             =   810
      Width           =   930
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   360
      Left            =   1740
      TabIndex        =   3
      Top             =   1290
      Width           =   1710
      Caption         =   "Save & Close"
      PicturePosition =   327683
      Size            =   "3016;635"
      Picture         =   "ManualOTFeeding.frx":0000
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
      Left            =   90
      TabIndex        =   2
      Top             =   1290
      Width           =   1590
      Caption         =   "Save & New    "
      PicturePosition =   327683
      Size            =   "2805;635"
      Picture         =   "ManualOTFeeding.frx":015E
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
      Left            =   3495
      TabIndex        =   4
      Top             =   1290
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
      Caption         =   "Select Employee        :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   90
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   90
      X2              =   4790
      Y1              =   1245
      Y2              =   1245
   End
End
Attribute VB_Name = "frmManualOTFeeding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LongTerm As Boolean

Private Sub PrintReport()

Screen.MousePointer = vbHourglass

Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report, ID As Integer
Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptAdvancesPayslip.rpt")
ID = Val(con.Execute("Select MAX(EntryID) FROM Advances").Fields(0) & "")

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
        cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & " and Active'"
    End If
    
End Sub

Private Function Saved() As Boolean

    Saved = False
    'On Error GoTo ERR
    
    'This form is just to update manual over time for salary of july 2004
    
    
    Dim SQL As String
    
    If Not cmbEmp.MatchFound Then
        MsgBox "Please Select Valid Employee From The List.", vbInformation
        Exit Function
    End If
    
    ExtraHrs = Val(txtOT) / 27
    
    SQL = "Update AttendanceSheet Set ExtraHrs=" & ExtraHrs & " Where EmpID='" & cmbEmp.ID & "'"
    
     
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
            MainForm.mnuLong_Click
         Else
            MainForm.mnuShort_Click
         End If
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

