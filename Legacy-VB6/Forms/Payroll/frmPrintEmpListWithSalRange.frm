VERSION 5.00
Begin VB.Form frmPrintEmpListWithSalRange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Employee List"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4890
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkActive 
      Caption         =   "Show Only Active Employees"
      Height          =   270
      Left            =   255
      TabIndex        =   11
      Top             =   1125
      Value           =   1  'Checked
      Width           =   2460
   End
   Begin VB.Frame Frame1 
      Height          =   765
      Left            =   2190
      TabIndex        =   10
      Top             =   1425
      Width           =   2625
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   495
         Left            =   75
         TabIndex        =   3
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   495
         Left            =   1305
         TabIndex        =   4
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.TextBox txtAmt2 
      Alignment       =   2  'Center
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
      Left            =   4830
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox txtAmt1 
      Alignment       =   2  'Center
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
      Left            =   3690
      TabIndex        =   1
      Top             =   720
      Width           =   1125
   End
   Begin VB.ComboBox cmbCriteria 
      Height          =   315
      ItemData        =   "frmPrintEmpListWithSalRange.frx":0000
      Left            =   1425
      List            =   "frmPrintEmpListWithSalRange.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   2250
   End
   Begin VB.TextBox txt 
      Alignment       =   2  'Center
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
      Left            =   270
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "Salary Is"
      Top             =   720
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Amount"
      Height          =   195
      Index           =   3
      Left            =   4830
      TabIndex        =   9
      Top             =   495
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Amount"
      Height          =   195
      Index           =   2
      Left            =   3690
      TabIndex        =   7
      Top             =   495
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Selection"
      Height          =   195
      Index           =   1
      Left            =   1470
      TabIndex        =   6
      Top             =   495
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Criteria"
      Height          =   195
      Index           =   0
      Left            =   300
      TabIndex        =   5
      Top             =   495
      Width           =   525
   End
End
Attribute VB_Name = "frmPrintEmpListWithSalRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_bContractorsOnly As Boolean

Private Sub cmbCriteria_Click()
    If cmbCriteria.ListIndex = 5 Then
        Frame1.Left = txtAmt2.Left + txtAmt2.Width - Frame1.Width
        Label1(3).Visible = True
        txtAmt2.Visible = True
        Me.Width = 6120
    Else
        Frame1.Left = txtAmt1.Left + txtAmt1.Width - Frame1.Width
        Label1(3).Visible = False
        txtAmt2.Visible = False
        Me.Width = 4980
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Call PrintEmpList
End Sub

Private Sub Form_Load()
    cmbCriteria.ListIndex = 0
    m_bContractorsOnly = False
End Sub

Public Sub PrintEmpList()
    
    Dim SelFormula As String
    If m_bContractorsOnly Then
        SelFormula = " {VEmp.TempDept}=True"
    Else
        SelFormula = " {VEmp.TempDept}=False"
    End If
    
    If chkActive.value = vbChecked Then
        SelFormula = SelFormula & " AND {VEmp.Active}=True "
    End If
    
    Select Case cmbCriteria.ListIndex
        Case 0
            SelFormula = SelFormula & " AND {VEmp.StartingSalary}>" & Val(txtAmt1)
        Case 1
            SelFormula = SelFormula & " AND {VEmp.StartingSalary}>=" & Val(txtAmt1)
        Case 2
            SelFormula = SelFormula & " AND {VEmp.StartingSalary}<" & Val(txtAmt1)
        Case 3
            SelFormula = SelFormula & " AND {VEmp.StartingSalary}<=" & Val(txtAmt1)
        Case 4
            SelFormula = SelFormula & " AND {VEmp.StartingSalary}=" & Val(txtAmt1)
        Case 5
            SelFormula = SelFormula & " AND {VEmp.StartingSalary}=" & Val(txtAmt1) & " to " & Val(txtAmt2)
        Case Else
            Exit Sub
    End Select
    
    Screen.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt, rpt As craxddrt.Report
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpList.rpt")
    Dim FormulaFields As craxddrt.FormulaFieldDefinitions
    Dim FormulaField As craxddrt.FormulaFieldDefinition

    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & strCompany & "'"
        End If
    Next

    f.ShowReport SelFormula, rpt

    Screen.MousePointer = vbDefault

End Sub

