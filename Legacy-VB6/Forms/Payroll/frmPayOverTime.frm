VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmPayOverTime 
   BorderStyle     =   0  'None
   ClientHeight    =   6885
   ClientLeft      =   15
   ClientTop       =   195
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleMode       =   0  'User
   ScaleWidth      =   10667.95
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1095
      Top             =   -15
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   10
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPayOverTime.frx":0000
            Key             =   "Paid"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmPayOverTime.frx":015A
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPayOverTime.frx":05AC
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   330
      TabIndex        =   19
      Top             =   330
      Width           =   1890
      _ExtentX        =   3334
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
   Begin Crystal.CrystalReport Cr 
      Left            =   705
      Top             =   1650
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
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   2205
      TabIndex        =   2
      Top             =   330
      Width           =   3060
      _ExtentX        =   5398
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
      Height          =   5940
      Left            =   15
      TabIndex        =   0
      Top             =   660
      Width           =   11385
      Begin VB.Frame FButs2 
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   150
         TabIndex        =   23
         Top             =   5280
         Width           =   2475
         Begin MSForms.CommandButton cmdSave 
            Height          =   375
            Left            =   135
            TabIndex        =   24
            Top             =   45
            Visible         =   0   'False
            Width           =   2175
            ForeColor       =   0
            Caption         =   "Update Payments"
            PicturePosition =   327683
            Size            =   "3836;661"
            Accelerator     =   85
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Frame FButs1 
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   7005
         TabIndex        =   20
         Top             =   5370
         Width           =   3105
         Begin MSForms.CommandButton cmdClose 
            Height          =   360
            Left            =   1545
            TabIndex        =   22
            Top             =   45
            Width           =   1425
            ForeColor       =   0
            Caption         =   " Close      "
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
            Left            =   75
            TabIndex        =   21
            Top             =   30
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
      Begin VB.TextBox txtHours 
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
         Left            =   3225
         TabIndex        =   5
         Top             =   1380
         Visible         =   0   'False
         Width           =   915
      End
      Begin Crystal.CrystalReport cr1 
         Left            =   3840
         Top             =   195
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
      Begin MSComctlLib.ListView LV 
         Height          =   4740
         Left            =   105
         TabIndex        =   1
         Top             =   180
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   8361
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
            Text            =   "Over Time Date"
            Object.Width           =   4972
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Hours"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Rate"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Over Time Amount"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Status"
            Object.Width           =   4366
         EndProperty
      End
      Begin VB.Label lblAmtPaid 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   7815
         TabIndex        =   16
         Top             =   5010
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Amt Paid :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   7005
         TabIndex        =   15
         Top             =   5055
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Balance :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   9210
         TabIndex        =   13
         Top             =   5070
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   9975
         TabIndex        =   12
         Top             =   5010
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Prev Paid :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4560
         TabIndex        =   11
         Top             =   5048
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lblAmtPrevPaid 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   5490
         TabIndex        =   10
         Top             =   5010
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Amount :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2355
         TabIndex        =   9
         Top             =   5048
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total Hours :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   5055
         Width           =   1065
      End
      Begin VB.Label lblTotalHrs 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   1230
         TabIndex        =   7
         Top             =   5010
         Width           =   1065
      End
      Begin VB.Label lblTotalAmt 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   3165
         TabIndex        =   6
         Top             =   5010
         Width           =   1305
      End
   End
   Begin MSComCtl2.DTPicker DtTo 
      Height          =   300
      Left            =   6660
      TabIndex        =   17
      Top             =   315
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   529
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
      Format          =   70844419
      CurrentDate     =   37081
   End
   Begin MSComCtl2.DTPicker DtFrom 
      Height          =   300
      Left            =   5280
      TabIndex        =   18
      Top             =   315
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
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
      Format          =   70844419
      CurrentDate     =   37055
   End
   Begin MSForms.CheckBox chkOnlyPayable 
      Height          =   345
      Left            =   8115
      TabIndex        =   14
      Top             =   285
      Width           =   1800
      VariousPropertyBits=   1015023643
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      DisplayStyle    =   4
      Size            =   "3175;609"
      Value           =   "1"
      Caption         =   "Show Only Payables"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.Label Label2 
      Height          =   255
      Left            =   330
      TabIndex        =   4
      Top             =   90
      Width           =   7695
      BackColor       =   11517387
      VariousPropertyBits=   8388627
      Caption         =   "  Department                Employee                                                   From                    To"
      Size            =   "13573;450"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdRefresh 
      Height          =   360
      Left            =   9945
      TabIndex        =   3
      Top             =   300
      Width           =   1530
      Caption         =   "Refresh        "
      PicturePosition =   327683
      Size            =   "2699;635"
      Accelerator     =   82
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Menu mnupop 
      Caption         =   "mnupop"
      Begin VB.Menu mnuDetail 
         Caption         =   "Current Employees Detail"
      End
      Begin VB.Menu mnuOTPayslip 
         Caption         =   "Over Time Payment Slip (Current Employee)"
      End
      Begin VB.Menu dash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSummary 
         Caption         =   "Over Time Payable Summary"
      End
   End
End
Attribute VB_Name = "frmPayOverTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkOnlyPayable_Click()
   Call RefreshLedger
   cmdSave.Visible = chkOnlyPayable
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
    Load NewOverTime
    If cmbEmp.Text <> "<All Employees>" Then
        NewOverTime.cmbDept.Text = cmbDept.Text
        NewOverTime.cmbEmp.Text = cmbEmp.Text
    Else
        If cmbDept.ID = 0 Then
            NewOverTime.cmbDept.Text = "<All Departments>"
        End If
    End If
    
    NewOverTime.Show 1
    Call RefreshLedger
End Sub

Private Sub PrintDetail()

    On Error GoTo err
      
    Screen.MousePointer = vbHourglass
    Dim Cond As String
    
    
    
    If cmbEmp.ID = "0" Then
        If cmbDept.ID = "0" Then
            Cond = "({VPayOverTime.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & ")) " & IIf(chkOnlyPayable, " and len(Cstr({VPayOvertime.PaidOn}))=0", "")
        Else
            Cond = "({VPayOverTime.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & ")) and {VPayOverTime.DeptID}='" & cmbDept.ID & "' " & IIf(chkOnlyPayable, " and len(Cstr({VPayOvertime.PaidOn}))=0", "")
        End If
    Else
        Cond = "{VPayOverTime.empid}='" & cmbEmp.ID & "' AND ({VPayOverTime.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & ")) " & IIf(chkOnlyPayable, " and len(Cstr({VPayOvertime.PaidOn}))=0", "")
    End If
    
With Cr
    .ReportFileName = App.Path & "\PayOverTime.rpt"
    .DataFiles(0) = DatabasePath
    .ReplaceSelectionFormula Cond
    .Formulas(0) = "FromTo='From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
    .Action = 1
    .PageZoomNext
    .Formulas(0) = ""
End With
Screen.MousePointer = vbDefault
Exit Sub

err:
    MsgBox err.Description
    Screen.MousePointer = vbDefault
End Sub

Private Sub PrintSummary(Optional PaidOverTime As Boolean = False)

    On Error GoTo err
      
    Screen.MousePointer = vbHourglass
    Dim Cond As String
    
    
    
    
        If cmbDept.ID = "0" Then
            Cond = "({VPayOverTime.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & ")) and len(Cstr({VPayOverTime.PaidOn}))=0"
        Else
            Cond = "({VPayOverTime.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & ")) and len(Cstr({VPayOverTime.PaidOn}))=0 and {VPayOverTime.DeptID}='" & cmbDept.ID & "' and len(Cstr({VPayOvertime.PaidOn}))=0"
        End If
    
        
With cr1
    .ReportFileName = App.Path & "\OTpayableSumm.rpt"
    .DataFiles(0) = DatabasePath
    .SelectionFormula = Cond
    

    .Formulas(0) = "FromTo='From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
    .Action = 1
    .PageZoomNext
    .Formulas(0) = ""
End With
    
Screen.MousePointer = vbDefault
Exit Sub
err:
    MsgBox err.Description
    Screen.MousePointer = vbDefault
End Sub

Private Sub PrintPaySlip(Optional EmpId As String)

    
      
    Dim Cond As String
    
    
    Cond = "{VEmp.EmpID}='" & cmbEmp.ID & "'"
    'If EmpID = "" Then
    '    If cmbDept.ID = "0" Then
    '        Cond = "({VPayOverTime.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & ")) and len(Cstr({VPayOvertime.PaidOn}))<>0"
    '    Else
    '        Cond = "({VPayOverTime.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & ")) and {VPayOverTime.DeptID}='" & cmbDept.ID & "' and len(Cstr({VPayOvertime.PaidOn}))<>0"
    '    End If
    'Else
    '    Cond = "{VPayOverTime.empid}='" & EmpID & "' AND ({VPayOverTime.DT}=date(" & Format(DtFrom, "yyyy,MM,dd") & ") TO date(" & Format(DtTo, "yyyy,MM,dd") & ")) and len(Cstr({VPayOvertime.PaidOn}))<>0"
    'End If
    
'With cr1
'    .ReportFileName = RptPath & "\OverTimePayslip.rpt"
'    '.DataFiles(0) = DatabasePath
'    .SelectionFormula = Cond
'    .Formulas(0) = "Rate=" & LV.ListItems(1).SubItems(3) & ""
'    .Formulas(1) = "Amount=" & lblBalance & ""
'    .Formulas(2) = "Hrs='" & lblTotalHrs & "'"
'    .Formulas(3) = "ForMonth='From " & Format(DTFrom, "dd-MMM-yyyy") & " To " & Format(DTTo, "dd-MMM-yyyy") & "'"
'
'    .Action = 1
'    .PageZoomNext
'End With


Screen.MousePointer = vbHourglass

Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report
Set rpt = rpt.OpenSubreport(rptPayrollPath & "\rptOverTimePayslip.rpt")
Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

Set FormulaFields = rpt.FormulaFields

For Each FormulaField In FormulaFields
    If FormulaField.Name = "{@Company}" Then
        FormulaField.Text = "'" & company & "'"
    ElseIf FormulaField.Name = "{@Rate}" Then
        FormulaField.Text = LV.ListItems(1).SubItems(2)
    ElseIf FormulaField.Name = "{@Amount}" Then
        FormulaField.Text = lblTotalAmt
    ElseIf FormulaField.Name = "{@Hrs}" Then
        FormulaField.Text = "'" & lblTotalHrs & "'"
    ElseIf FormulaField.Name = "{@ForMonth}" Then
        FormulaField.Text = "'From " & Format(DtFrom, "dd-MMM-yyyy") & " To " & Format(DtTo, "dd-MMM-yyyy") & "'"
    End If
Next


f.ShowReport Cond & " ", rpt

Screen.MousePointer = vbDefault


End Sub



Private Sub cmdPrint_Click()
'PopUpMenu mnupop
    Call PrintPaySlip(cmbEmp.ID)

End Sub

Public Sub cmdRefresh_Click()
    
Dim DT As Date
'''''''''''''''''' This Will Update The LateHrs & OverTime ''''''''''
    DT = DtFrom
    For i = 1 To DateDiff("m", DtFrom, DtTo) + 1
        con.Execute "EXEC SP_CalcOvertimeEmpWise '" & DT & "','" & cmbEmp.ID & "'"
        DT = DateAdd("m", i, DtFrom)
    Next
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Call RefreshLedger
    
End Sub

Private Sub cmdSave_Click()

On Error GoTo err
con.BeginTrans
    For i = 1 To LV.ListItems.Count
        With LV.ListItems(i)
          If .Checked Then
            con.Execute "Update AttendanceSheet Set PaidOn='" & Format(Date, "dd-MMM-yyyy") _
                & "', Rate=" & Val(.SubItems(2)) & " Where EmpID='" & .Tag _
                & "' AND DT='" & .ListSubItems(1).Tag & "' "
            End If
        End With
    Next
con.CommitTrans
Exit Sub

err:
   MsgBox err.Description
   con.RollbackTrans
   
End Sub

Private Sub Form_Load()
'cmdNew.Enabled = UserHasAccess("AddOverTime")
    DtTo = Date
    DtFrom = DateAdd("m", -1, Date)
    cmbYear = CYear
    cmbMonth = Format(Date, "MMM")
    
    
    
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID"
    
    cmbDept.Text = "<All Departments>"
    
    'cmbEmp.Text = "<All Employees>"
    mnuPop.Visible = False
End Sub
Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    'cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals con, "DISTINCT('{' + Empid + '} ' + Name)", "VAuthorizedEmps", "EmpID"
    Else
        cmbEmp.AddVals con, "DISTINCT('{' + Empid + '} ' + Name)", "VAuthorizedEmps", "EmpID", "Where DeptID='" & cmbDept.ID & "'"
    End If
    
End Sub

Private Sub Form_Resize()
    

    On Error Resume Next
    Dim TopMargin As Integer
    TopMargin = cmbEmp.Top + cmbEmp.Height + 30
    
    FLV.Move 50, TopMargin, Me.ScaleWidth - 100, Me.ScaleHeight - TopMargin
    
    LV.Move 100, 200, Me.Width - 400, FLV.Height - fButs1.Height - 300
    
    fButs1.Move LV.Left + LV.Width - fButs1.Width, LV.Top + LV.Height
    
    FButs2.Move LV.Left, LV.Top + LV.Height





' FLV.Move (Me.ScaleWidth - FLV.Width) / 2
' FAcc.Move (Me.ScaleWidth - FAcc.Width) / 2
' fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
'cmdRefresh.Move FLV.Left + FLV.Width - cmdRefresh.Width
'cmdClose.Move Me.ScaleWidth - cmdClose.Width - 50, Me.ScaleHeight - cmdClose.Height - 100
'cmdPrint.Move cmdClose.Left - cmdPrint.Width - 50, cmdClose.Top

'FDrag.Move 50, Me.ScaleHeight * 3 / 5, Me.ScaleWidth - 100
 
 'If LVItems.Visible Then
 '  TheHeight = FDrag.Top
 '  FDrag.Visible = True
 'Else
 '  TheHeight = Me.ScaleHeight - cmdClose.Height - 200
 '  FDrag.Visible = False
 'End If
 
 'LV.Move 50, FAcc.Top + FAcc.Height, Me.ScaleWidth - 100, TheHeight - (FAcc.Top + FAcc.Height)
 'LVItems.Move 50, FDrag.Top + FDrag.Height, Me.ScaleWidth - 100, Me.ScaleHeight - (FAcc.Height + cmdClose.Height + LV.Height + 250)
 'FDrag.ZOrder 1
 
End Sub



Private Sub RefreshLedger()

Dim Rate As Double
Rate = 0
    If cmbDept.MatchFound = False Or cmbEmp.MatchFound = False Then Exit Sub
    Dim rsLedger As New ADODB.Recordset
    
    Dim ITM As ListItem, TotalDays As Integer
    Dim Cond As String, SQL As String
    Dim SDate As Date
    
    LV.Checkboxes = chkOnlyPayable
            
    Cond = " Where empid='" & cmbEmp.ID & "' and DT Between '" & Format(DtFrom, "dd-MMM-yyyy") & "' and '" & Format(DtTo, "dd-MMM-yyyy") & "' AND ExtraHrs>0"
    
    If chkOnlyPayable = True Then Cond = Cond & " and PaidOn Is NULL"
    
    SQL = "SELECT * FROM VOverTime " & Cond & " ORDER BY EmpID,DT"
    
    With rsLedger
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        LV.Visible = False
        Set LV.SmallIcons = ImageList1
        LV.ListItems.Clear
        For i = 0 To .RecordCount - 1
            If IsNull(!paidon) Then
                SDate = DateSerial(year(!DT), Month(!DT), 1)
                TotalDays = DateAdd("m", 1, SDate) - SDate
                'TotalDays = TotalSalDays - Val(con.Execute("Select COUNT(*) FROM Holidays Where MONTH(DT)=" & Month(!DT) & " AND YEAR(DT)=" & year(!DT) & "").Fields(0))
                'TotalDays = TotalDays - TotalSundays(Month(!DT), year(!DT), TotalSalDays)
                Rate = Round((Val(!StartingSalary & "") / TotalDays) / !OverTimeHrs, 2)
                cmdSave.Visible = True
            Else
                Rate = Val(!Rate & "")
            End If
            
            Set ITM = LV.ListItems.add(, , Format(![DT], "dddd dd-MMM-yyyy"))
               ITM.Tag = !EmpId & ""
               ITM.Checked = False
               If ITM.Tag <> "" Then ITM.SmallIcon = 1
'               Itm.SubItems(1) = !Shift & ""
               ITM.SubItems(1) = MinsToTime(Val(![ExtraHrs] & "") * 60)
               ITM.ListSubItems(1).Tag = !DT
               'Rate = Val(![Rate] & "")
               ITM.SubItems(2) = Format(Rate, "0.00")
               ITM.ListSubItems(2).Tag = Val(!ExtraHrs & "")
               ITM.SubItems(3) = Format(Rate * Val(![ExtraHrs] & ""), "0.00")
               ITM.SubItems(4) = IIf(IsNull(![paidon]), "< Payable >", "< Paid On " & Format(![paidon], "dd-MMM-yyyy") & ">")
               .MoveNext
        Next
        
        LV.Visible = True
        .Close
    End With
    Set rsLedger = Nothing
    SQL = ""
    Cond = ""
Call GetTotals
End Sub


Private Sub LV_DblClick()


Exit Sub
'''''''''''''''''''''''''''''''
    If LV.ListItems.Count = 0 Or UserHasAccess("EditOverTime") = False Then Exit Sub
    If CDate(LV.SelectedItem.Tag) <> Date Then
      MsgBox "You Can Edit Over Time Only For Today.", vbInformation
      Exit Sub
    End If

    With txtHours
        .Move LV.Left + LV.ColumnHeaders(5).Left + 35, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(5).Width
        .Text = LV.SelectedItem.SubItems(4)
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        LV.SelectedItem.SubItems(4) = "< Paid On " & Format(Date, "dd-MMM-yyyy") & " >"
    Else
        LV.SelectedItem.SubItems(4) = "< Payable >"
    End If
   Call GetTotals
End Sub

Private Sub mnuDetail_Click()
Call PrintDetail
End Sub

Private Sub mnuOTPayslip_Click()
    Call PrintPaySlip(cmbEmp.ID)
End Sub

Private Sub mnuOTPayslips_Click()
    Call PrintPaySlip
End Sub

Private Sub mnuSummary_Click()
    Call PrintSummary
End Sub

Private Sub mnuSummaryPaid_Click()
    Call PrintSummary(True)
End Sub

Private Sub txtHours_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   con.Execute "Update OverTime Set Hours=" & Val(txtHours) & " Where Empid='" & LV.SelectedItem & "' and dt=#" & CDate(LV.SelectedItem.Tag) & "#"
   LV.SelectedItem.SubItems(4) = Val(txtHours)
   LV.SetFocus
ElseIf KeyAscii = 27 Then
   LV.SetFocus
Else
   KeyAscii = OnlyNumber(KeyAscii, True)
End If
   
End Sub

Private Sub txtHours_LostFocus()
   txtHours.Visible = False
End Sub


Private Sub GetTotals()
Dim totalHrs As Double
Dim TotalAmt As Double
Dim AmtPaid As Double
Dim AmtPrevPaid As Double

For i = 1 To LV.ListItems.Count
   totalHrs = totalHrs + Val(LV.ListItems(i).ListSubItems(2).Tag)
   If LV.ListItems(i).Tag = "Paid" Then
      AmtPrevPaid = AmtPrevPaid + Val(LV.ListItems(i).SubItems(3))
   ElseIf LV.ListItems(i).Checked Then
      AmtPaid = AmtPaid + Val(LV.ListItems(i).SubItems(3))
   End If
   
   TotalAmt = TotalAmt + Val(LV.ListItems(i).SubItems(3))
   
Next

lblTotalHrs = MinsToTime(totalHrs * 60)
lblTotalAmt = Format(TotalAmt, "0.00")
lblAmtPrevPaid = Format(AmtPrevPaid, "0.00")
lblAmtPaid = Format(AmtPaid, "0.00")
lblBalance = Format(TotalAmt - AmtPaid - AmtPrevPaid, "0.00")
End Sub

