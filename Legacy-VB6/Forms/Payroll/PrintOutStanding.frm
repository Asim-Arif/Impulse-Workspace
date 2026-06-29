VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form PrintOutStanding 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Employees Out Stading Loans Report"
   ClientHeight    =   1695
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   165
      TabIndex        =   1
      Top             =   465
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextForeColor   =   12582912
      ListIndex       =   -1
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   2805
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Month                Year"
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
         Index           =   1
         Left            =   150
         TabIndex        =   7
         Top             =   870
         Width           =   2475
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Department"
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
         Left            =   135
         TabIndex        =   6
         Top             =   150
         Width           =   2490
      End
      Begin MSForms.ComboBox cmbMonth 
         Height          =   300
         Left            =   150
         TabIndex        =   3
         Top             =   1110
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
      Begin MSForms.ComboBox cmbYear 
         Height          =   300
         Left            =   1485
         TabIndex        =   2
         Top             =   1110
         Width           =   1155
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "2037;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   0
      Top             =   0
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
   Begin VB.Image Image2 
      Height          =   480
      Left            =   3510
      Picture         =   "PrintOutStanding.frx":0000
      Top             =   1080
      Width           =   375
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   360
      Left            =   3015
      TabIndex        =   5
      Top             =   615
      Width           =   1425
      ForeColor       =   0
      Caption         =   "Cancel       "
      PicturePosition =   327683
      Size            =   "2514;635"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   360
      Left            =   3015
      TabIndex        =   4
      Top             =   165
      Width           =   1425
      ForeColor       =   0
      Caption         =   "Print         "
      PicturePosition =   327683
      Size            =   "2514;635"
      Accelerator     =   80
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "PrintOutStanding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim Cond As String
Cond = IIf(cmbDept.ID = "0", "", "{VOutStanding.DeptID}='" & cmbDept.ID & "' and ") & _
"(month({VOutStanding.DiscontinuedOn})=" & cmbMonth.List(cmbMonth.ListIndex, 1) & _
" and Year({VOutStanding.DiscontinuedOn})=" & cmbYear & ")"



Me.Hide


Screen.MousePointer = vbHourglass

Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptOutStanding.rpt")

Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

Set FormulaFields = rpt.FormulaFields
For Each FormulaField In FormulaFields
    If FormulaField.Name = "{@Company}" Then
        FormulaField.Text = "'" & company & "'"
    ElseIf FormulaField.Name = "{@FromTo}" Then
        FormulaField.Text = "'" & cmbMonth & ", " & cmbYear & "'"
    End If
Next
Me.Hide
f.ShowReport Cond & " ", rpt


Screen.MousePointer = vbDefault


End Sub

Private Sub Form_Load()
    cmbDept.ListHeight = 900
    
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
    

    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID"
    cmbDept.Text = "<All Departments>"
    
    cmbYear = year(GetServerDate)
    cmbMonth = Format(GetServerDate, "MMM")
    
End Sub

