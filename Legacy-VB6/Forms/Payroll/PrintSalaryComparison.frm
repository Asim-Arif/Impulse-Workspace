VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form PrintSalaryComparison 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Salary Comparison Report"
   ClientHeight    =   1890
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "To :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   45
      TabIndex        =   6
      Top             =   960
      Width           =   2625
      Begin MSForms.ComboBox cmbYearTo 
         Height          =   300
         Left            =   1395
         TabIndex        =   9
         Top             =   480
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
      Begin MSForms.ComboBox cmbMonthTo 
         Height          =   300
         Left            =   60
         TabIndex        =   8
         Top             =   495
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
      Begin MSForms.Label Label2 
         Height          =   240
         Left            =   60
         TabIndex        =   7
         Top             =   240
         Width           =   2475
         BackColor       =   11517387
         VariousPropertyBits=   8388627
         Caption         =   " Month                Year"
         Size            =   "4366;423"
         BorderColor     =   16512
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "From :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   2625
      Begin MSForms.Label Label1 
         Height          =   240
         Left            =   60
         TabIndex        =   3
         Top             =   240
         Width           =   2475
         BackColor       =   11517387
         VariousPropertyBits=   8388627
         Caption         =   " Month                Year"
         Size            =   "4366;423"
         BorderColor     =   16512
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbMonth 
         Height          =   300
         Left            =   60
         TabIndex        =   2
         Top             =   495
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
         Left            =   1395
         TabIndex        =   1
         Top             =   495
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
      Height          =   435
      Left            =   3420
      Picture         =   "PrintSalaryComparison.frx":0000
      Top             =   1200
      Width           =   480
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   360
      Left            =   2805
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
      Left            =   2805
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
Attribute VB_Name = "PrintSalaryComparison"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdClose_Click()
Unload Me
End Sub




Private Sub cmdPrint_Click()

On Error GoTo err
Dim FromDate As Date, ToDate As Date, Cond As String

FromDate = DateSerial(cmbYear, cmbMonth.List(cmbMonth.ListIndex, 1), 1)
ToDate = DateSerial(cmbYearTo, cmbMonthTo.List(cmbMonthTo.ListIndex, 1), 1)

Cond = "{VSalaryComparison.Dt}=date(" & Format(FromDate, "yyyy,MM,dd") & ") to Date(" & Format(ToDate, "yyyy,MM,dd") & ")"

    
Me.Hide


    Screen.MousePointer = vbHourglass

    Dim f As New frmPrevRpt, rpt As craxddrt.Report
    Set rpt = rpt.OpenSubreport(rptPayrollPath & "\rptSalaryComparison.rpt")
    
    Dim FormulaFields As craxddrt.FormulaFieldDefinitions
    Dim FormulaField As craxddrt.FormulaFieldDefinition
    
    Set FormulaFields = rpt.FormulaFields
    
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & company & "'"
        ElseIf FormulaField.Name = "{@FromTo}" Then
            FormulaField.Text = "'From " & cmbMonth & ", " & cmbYear & " To " & cmbMonthTo & ", " & cmbYearTo & "'"
        End If
    Next
    
    f.ShowReport Cond & " ", rpt
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub

Private Sub Form_Load()

    
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
        cmbYearTo.AddItem i
    Next i
    
    
    
    
    
    cmbMonthTo.AddItem "Jan"
    cmbMonthTo.List(0, 1) = 1
    cmbMonthTo.AddItem "Feb"
    cmbMonthTo.List(1, 1) = 2
    cmbMonthTo.AddItem "Mar"
    cmbMonthTo.List(2, 1) = 3
    cmbMonthTo.AddItem "Apr"
    cmbMonthTo.List(3, 1) = 4
    cmbMonthTo.AddItem "May"
    cmbMonthTo.List(4, 1) = 5
    cmbMonthTo.AddItem "Jun"
    cmbMonthTo.List(5, 1) = 6
    cmbMonthTo.AddItem "Jul"
    cmbMonthTo.List(6, 1) = 7
    cmbMonthTo.AddItem "Aug"
    cmbMonthTo.List(7, 1) = 8
    cmbMonthTo.AddItem "Sep"
    cmbMonthTo.List(8, 1) = 9
    cmbMonthTo.AddItem "Oct"
    cmbMonthTo.List(9, 1) = "10"
    cmbMonthTo.AddItem "Nov"
    cmbMonthTo.List(10, 1) = 11
    cmbMonthTo.AddItem "Dec"
    cmbMonthTo.List(11, 1) = 12
    
    

    
   
    cmbYear = year(GetServerDate)
    cmbMonth = Format(GetServerDate, "MMM")
    
End Sub

