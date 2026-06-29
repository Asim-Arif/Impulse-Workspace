VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form NewAdvanceShortBPV 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Short Advance"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8475
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
   ScaleHeight     =   7035
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   315
      Left            =   1755
      TabIndex        =   0
      Top             =   30
      Width           =   6630
      _ExtentX        =   11695
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
      Width           =   6630
      _ExtentX        =   11695
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
   Begin VB.Frame FraBank 
      Caption         =   "Cheque Details :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   135
      TabIndex        =   35
      Top             =   4320
      Width           =   8220
      Begin VB.CommandButton cmdBank 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7740
         Picture         =   "NewAdvanceShortBPV.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   555
         Width           =   360
      End
      Begin VB.TextBox Tbox 
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
         Index           =   2
         Left            =   60
         TabIndex        =   37
         Top             =   1275
         Width           =   1725
      End
      Begin VB.CheckBox ChkClear 
         Caption         =   "Clear Cheque "
         Enabled         =   0   'False
         Height          =   225
         Left            =   6750
         TabIndex        =   36
         ToolTipText     =   "Clear This Cheque Automatically"
         Top             =   1335
         Value           =   1  'Checked
         Width           =   1350
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   5295
         TabIndex        =   39
         Top             =   1275
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   187564035
         CurrentDate     =   37250
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   18
         Left            =   5280
         TabIndex        =   45
         Top             =   1065
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cheque Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   17
         Left            =   1830
         TabIndex        =   44
         Top             =   1020
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cheque No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   16
         Left            =   75
         TabIndex        =   43
         Top             =   1035
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bank Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   15
         Left            =   60
         TabIndex        =   42
         Top             =   330
         Width           =   960
      End
      Begin VB.Label lblBankAcc 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   45
         TabIndex        =   41
         Top             =   570
         Width           =   7665
      End
      Begin MSForms.ComboBox cmbCheckType 
         Height          =   315
         Left            =   1815
         TabIndex        =   40
         Top             =   1275
         Width           =   3450
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "6085;556"
         ListWidth       =   2822
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Deductions :"
      Height          =   1440
      Left            =   3585
      TabIndex        =   22
      Top             =   1665
      Width           =   4800
      Begin MSForms.TextBox txtPrevST 
         Height          =   315
         Left            =   3690
         TabIndex        =   34
         Top             =   1005
         Width           =   1050
         VariousPropertyBits=   746604575
         Size            =   "1852;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Previouse Short Term :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   2025
         TabIndex        =   33
         Top             =   1035
         Width           =   1650
      End
      Begin MSForms.TextBox txtLunchAmt 
         Height          =   315
         Left            =   3690
         TabIndex        =   26
         Top             =   660
         Width           =   1050
         VariousPropertyBits=   746604575
         Size            =   "1852;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin MSForms.TextBox txtLongTerm 
         Height          =   315
         Left            =   3690
         TabIndex        =   25
         Top             =   270
         Width           =   1050
         VariousPropertyBits=   746604575
         Size            =   "1852;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lunch Amount :"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   10
         Left            =   2115
         TabIndex        =   24
         Top             =   675
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fixed Advance Inst. : "
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   9
         Left            =   2115
         TabIndex        =   23
         Top             =   315
         Width           =   1605
      End
   End
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
      Left            =   1590
      TabIndex        =   10
      Top             =   6630
      Value           =   1  'Checked
      Width           =   1800
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1755
      TabIndex        =   2
      Top             =   765
      Width           =   2655
      _ExtentX        =   4683
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
      Format          =   187564035
      CurrentDate     =   37384
   End
   Begin MSComCtl2.DTPicker DTDedStartMonth 
      Height          =   315
      Left            =   4170
      TabIndex        =   48
      Top             =   3840
      Width           =   1920
      _ExtentX        =   3387
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
      CustomFormat    =   "MMMM, yyyy"
      Format          =   187564035
      CurrentDate     =   40229
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ded. Month :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   20
      Left            =   3210
      TabIndex        =   49
      Top             =   3885
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Vchr # :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   19
      Left            =   1080
      TabIndex        =   47
      Top             =   3915
      Width           =   585
   End
   Begin MSForms.TextBox txtVchrNo 
      Height          =   315
      Left            =   1725
      TabIndex        =   46
      Top             =   3855
      Width           =   1455
      VariousPropertyBits=   746604575
      Size            =   "2566;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.TextBox txtMin 
      Height          =   315
      Left            =   7350
      TabIndex        =   32
      Top             =   3465
      Width           =   1050
      VariousPropertyBits=   746604575
      Size            =   "1852;556"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtMax 
      Height          =   315
      Left            =   1725
      TabIndex        =   31
      Top             =   3465
      Width           =   1455
      VariousPropertyBits=   746604575
      Size            =   "2566;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox txtBalance 
      Height          =   315
      Left            =   5040
      TabIndex        =   30
      Top             =   3465
      Width           =   1050
      VariousPropertyBits=   746604575
      Size            =   "1852;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Min. Allowed :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   13
      Left            =   6285
      TabIndex        =   29
      Top             =   3510
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Max. Allowed :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   12
      Left            =   645
      TabIndex        =   28
      Top             =   3510
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Balance :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   11
      Left            =   4335
      TabIndex        =   27
      Top             =   3510
      Width           =   660
   End
   Begin MSForms.TextBox txtAmount 
      Height          =   315
      Left            =   7335
      TabIndex        =   21
      Top             =   1125
      Width           =   1050
      VariousPropertyBits=   746604575
      Size            =   "1852;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Amount :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   8
      Left            =   6600
      TabIndex        =   20
      Top             =   1155
      Width           =   660
   End
   Begin MSForms.TextBox txtHrsWorked 
      Height          =   315
      Left            =   5040
      TabIndex        =   19
      Top             =   1125
      Width           =   1050
      VariousPropertyBits=   746604575
      Size            =   "1852;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hrs Worked :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   7
      Left            =   4050
      TabIndex        =   18
      Top             =   1215
      Width           =   945
   End
   Begin MSForms.TextBox txtWorkingHrs 
      Height          =   315
      Left            =   1725
      TabIndex        =   17
      Top             =   1110
      Width           =   1050
      VariousPropertyBits=   746604575
      Size            =   "1852;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Working Hrs               :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   6
      Left            =   105
      TabIndex        =   16
      Top             =   1140
      Width           =   1605
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Rate / Hr :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   5
      Left            =   6495
      TabIndex        =   15
      Top             =   840
      Width           =   765
   End
   Begin MSForms.TextBox txtRate 
      Height          =   315
      Left            =   7335
      TabIndex        =   14
      Top             =   780
      Width           =   1050
      VariousPropertyBits=   746604575
      Size            =   "1852;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Salary :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   3
      Left            =   4455
      TabIndex        =   13
      Top             =   810
      Width           =   555
   End
   Begin MSForms.TextBox txtSalary 
      Height          =   315
      Left            =   5040
      TabIndex        =   12
      Top             =   780
      Width           =   1050
      VariousPropertyBits=   746604575
      Size            =   "1852;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.TextBox txtBal 
      Height          =   315
      Left            =   7335
      TabIndex        =   11
      Top             =   3855
      Width           =   1050
      VariousPropertyBits=   746604571
      BackColor       =   16761024
      ForeColor       =   255
      Size            =   "1852;556"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
      FontWeight      =   700
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   3540
      X2              =   8240
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Amount Paid :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   6285
      TabIndex        =   9
      Top             =   3915
      Width           =   1005
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   360
      Left            =   5100
      TabIndex        =   4
      Top             =   6540
      Width           =   1560
      Caption         =   "Save & Close"
      PicturePosition =   327683
      Size            =   "2752;635"
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
      Left            =   3480
      TabIndex        =   3
      Top             =   6540
      Width           =   1560
      Caption         =   "Save & New    "
      PicturePosition =   327683
      Size            =   "2752;635"
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
      Left            =   6720
      TabIndex        =   5
      Top             =   6540
      Width           =   1560
      Caption         =   "Cancel    "
      PicturePosition =   327683
      Size            =   "2752;635"
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
      Left            =   105
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
      Left            =   120
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
      Left            =   135
      TabIndex        =   6
      Top             =   90
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   3540
      X2              =   8240
      Y1              =   6495
      Y2              =   6495
   End
End
Attribute VB_Name = "NewAdvanceShortBPV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LongTerm As Boolean
Dim m_bContractors As Boolean

Private Sub PrintReport(Optional bDirectPrint As Boolean = False)

    Screen.MousePointer = vbHourglass
    Dim lID As Long
    lID = Val(con.Execute("Select Max(EntryID) From Advances").Fields(0).Value & "")
    
    con.Execute "Delete From PrintTempAdv"
    
    con.Execute "Insert Into PrintTempAdv(EmpID,BSal,Rate,SDays,SAmt,Total,Paid,LongTerm,Balance,DT,LunchAmt,AmtPaid,AdvancesID) Values('" & _
     cmbEmp.id & "'," & Val(txtSalary) & "," & Val(txtRate) & "," & Val(txtHrsWorked) & "," & Val(txtAmount) & "," & _
     Val(txtAmt) & "," & Val(txtPrevST) & "," & Val(txtLongTerm) & "," & Val(txtBalance) & ",'" & DTPicker1 & "'," & Val(txtLunchAmt) & "," & Val(txtBal) & "," & lID & ")"
    
    Dim rpt As New CRAXDDRT.Report
    
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptTempAdv.rpt")
    
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
    
    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & strCompany & "'"
        ElseIf FormulaField.Name = "{@FromTo}" Then
            FormulaField.Text = "'" & Format(DTPicker1, "dd-MMM-yyyy") & "'"
        ElseIf FormulaField.Name = "{@MonthDays}" Then
            FormulaField.Text = Val(txtWorkingHrs) 'lWHrs * 8 'Val(txtTotalHrs)
        End If
    Next
   
    
    If bDirectPrint Then
        PrintReportDirectly " ", rpt
    Else
        Me.Hide
        Dim f As New frmPrevRpt
        f.ShowReport " ", rpt
    End If
        
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    'cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Departments>" Then
        If m_bContractors Then
            cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "VTempEmps", "EmpID", "Where Active=1 "
        Else
            cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "VSalariedEmps", "EmpID", "Where Active=1 "
        End If
    Else
        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' - ' + FName + ' {' + Designation + '}'", "Employees", "EmpID", "Where DeptID='" & cmbDept.id & "' and Active=1"
    End If
    
End Sub

Private Function Saved() As Boolean

    Saved = False
    'On Error GoTo ERR
    
'    If Not m_bContractors Then
'        If Month(DTPicker1) <> Month(GetServerDate) Then
'            MsgBox "Invalid Date.", vbInformation
'            Exit Function
'        ElseIf year(DTPicker1) <> year(GetServerDate) Then
'            MsgBox "Invalid Date.", vbInformation
'            Exit Function
'        End If
'    End If
    
    Dim Sql As String
    
    If Not cmbEmp.MatchFound Then
        MsgBox "Please Select Valid Employee From The List.", vbInformation
        Exit Function
    ElseIf Val(txtBal) = 0 Then
        MsgBox "Invalid Amount.", vbInformation
        Exit Function
    ElseIf strEmpShortTermAccNo = "" Then
        MsgBox "Can't Post, Emp. Short Term Loan Account Not Found.", vbInformation
        Exit Function
'    ElseIf Val(txtBal) > Val(txtMax) - Val(txtPrevST) Then
'        MsgBox "Sorry, Not Allowed Greater Than Rs. " & (Val(txtMax) - Val(txtPrevST)) & ".", vbInformation
'        Exit Function
    End If
    
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(*)", "Advances", " WHERE Type=0 AND EmpID='" & cmbEmp.id & "' AND DT='" & DTPicker1.Value & "'")
    If lCount > 0 Then
        If MsgBox("You have already Posted Advance in Selected date. Do you want to Add This?", vbQuestion + vbYesNo) = vbNo Then
            Exit Function
        End If
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
    Dim Issued As Boolean, ChqBookNo As Double
    
    If lblBankAcc = "" Then
        MsgBox "Invalid Bank Selected.", vbInformation
        Exit Function
    ElseIf Not ValidChqNo(GetID(lblBankAcc), Tbox(2), False, ChqBookNo) Then
        If Issued Then
            MsgBox "The Cheque With Number '" & txtChqNo & "' Has Been Issued.", vbInformation
        Else
            MsgBox "Invalid Cheque No.", vbInformation
        End If
        Tbox(2).SetFocus
        Exit Function
    End If
    Tbox(2).Tag = ChqBookNo
    
    Dim strVchrNo As String
    Call StartTrans(con)
    
    If Not m_bContractors Then
        strVchrNo = SaveTransForCheque
    End If
    
    Sql = "Insert Into Advances(EmpID,DT,Type,Description," & _
     "Amount,DAmount,AmountCleared,AccVoucherNo,DeductionStartDT,ActualAmountTaken) VALUES('" & cmbEmp.id & _
     "','" & DTPicker1 & "',0,'Short Term'," & Val(txtBal) & ",0,0,'" & _
     strVchrNo & "','" & DTDedStartMonth.Value & "'," & Val(txtBal) & ")"
    
    con.Execute Sql
    
    con.CommitTrans
        
    MsgBox "Saved Successfully.", vbInformation
    
    Saved = True
    If chkPrint = 1 Then Call PrintVoucher(MainForm.cr1, strVchrNo)
    Sql = ""
    Exit Function
err:
    MsgBox err.Description


End Function

Private Function AddVoucher() As String

    Dim i As Long, INextSno As Double, Bal As Double, Amt As Double
    Dim VchrNo As String, theDesc As String, AccNo As String
    
    Dim Dated As Date
    Dated = DTPicker1
    'Charge All Amount to Cash
    Amt = Val(txtBal)
    
    If Amt = 0 Then Exit Function
    
    'ConAcc.BeginTrans
     
    'Get The Required Values
    theDesc = "Temp. Advance Paid To " & cmbEmp.Text
    Dim bWaitForAuth As Boolean
    bWaitForAuth = VoucherRequiresAuthorization("CPV")
    If bWaitForAuth = False Then
        INextSno = getNextSno(Dated)
        VchrNo = getNextVchrNo(Dated, "CPV")
     
        ''''''''''''''''Charge To  Cash Entry'''''''''''''
    
        AccNo = cmbCashInHand.List(cmbCashInHand.ListIndex, 1) 'CashInHand
        Bal = getBalance(AccNo, Dated)
        Bal = Bal - Amt
        'Credit Entry
        con.Execute "INSERT INTO Vouchers VALUES(" & INextSno & ",'" & _
         Dated & "','" & VchrNo & "','" & AccNo & _
         "','" & theDesc & "',0," & Amt & "," & Bal & ",'',0)"
    
        Call adjustLedger(AccNo, Dated, INextSno, -Val(Amt))
     
        INextSno = INextSno + 1
'        If strCompany = "Dr-Frgz" And strSubCompany = "Intl" Or strCompany = "Banzai" Then
'            AccNo = "15-004-01001"
'        Else
'            AccNo = "13-003-19001" 'IIf(LongTerm, LongAdvAccNo, ShortAdvAccNo)
'        End If
        AccNo = strEmpShortTermAccNo
        
        Bal = getBalance(AccNo, Dated)
      
        Amt = Amt
        Bal = Bal + Amt
       
        con.Execute "insert into Vouchers Values(" & INextSno & ",'" & _
         Dated & "','" & VchrNo & "','" & AccNo & _
         "','" & theDesc & "'," & Amt & ",0," & Bal & ",'',0)"
        
        Call adjustLedger(AccNo, Dated, INextSno, Val(Amt))
    
        con.Execute "Insert Into VoucherInfo(VchrNo,UserName,MachineName,DT) Values('" & VchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
    Else
        con.Execute "INSERT INTO VouchersForAuthorization(VDate,VchrNo,Accno,Description,Debit,Credit,balance,DpstSlip,CSNo) VALUES('" & Dated & "','" & _
             VchrNo & "','" & cmbCashInHand.List(cmbCashInHand.ListIndex, 1) & "','" & theDesc & "',0," & Amt & ",0,'',0)"
        
        con.Execute "INSERT INTO VouchersForAuthorization(VDate,VchrNo,Accno,Description,Debit,Credit,balance,DpstSlip,CSNo) VALUES('" & Dated & "','" & _
             VchrNo & "','" & strEmpShortTermAccNo & "','" & theDesc & "'," & Amt & ",0,0,'',0)"
             
        con.Execute "INSERT INTO VouchersForAuthorizationInfo(VchrNo,UserName,MachineName,DT,PostedThroughJVForm,Notes,Payee) VALUES('" & VchrNo & "','" & _
         CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "',0,'','')"
    End If
    AddVoucher = VchrNo
    
End Function

Private Sub cmbEmp_matched()

    On Error GoTo err
    
    Dim TotalSalDays As Integer
    Dim Rate As Double
    Dim totalHrs As Integer
    Dim MonthDays As Long
    Dim TotalMonthDays As Integer
    Dim bPerPiece As Boolean
    
    TotalSalDays = Day(GetServerDate)
    
    'TotalHrs = Day(DateSerial(DTPicker1.year, DTPicker1.Month + 1, 0))
    totalHrs = Day(DateSerial(DTPicker1.year, DTPicker1.Month, DTPicker1.Day))
    MonthDays = totalHrs - TotalSundays(DTPicker1.Month, DTPicker1.year, totalHrs)
    TotalSalDays = Day(DateSerial(DTPicker1.year, DTPicker1.Month + 1, 0))
    Hdays = Val(con.Execute("Select Count(DT) From Holidays Where Month(DT)=" & _
    DTPicker1.Month & " and Year(DT)=" & DTPicker1.year).Fields(0))
    MonthDays = MonthDays '- Hdays
    totalHrs = MonthDays * 8
    txtWorkingHrs = totalHrs
    
    bPerPiece = con.Execute("Select IsNull(PerPieceRate,0) From Employees Where EmpID='" & cmbEmp.id & "'").Fields(0)
    
    Dim DT As Date
    'DT = DateSerial(DTPicker1.year, DTPicker1.Month, Day(GetServerDate))
    DT = DateSerial(DTPicker1.year, DTPicker1.Month, Day(DTPicker1))
    
    Dim cmd As New ADODB.Command
    
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    
    Dim Holidays As Integer
    Dim TotalSuns As Integer
    
    Dim lDedAmt As Long, lClsAmt As Long, lTotAmt As Long
    With cmd
        .CommandTimeout = 120
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        If m_bContractors Then
            .CommandText = "SP_TempEmpPaymentSheet"
        Else
            .CommandText = "VAdvanceSalary"
        End If
        .Parameters("@ParDate").Value = DT
        .Parameters("@EmpID").Value = cmbEmp.id
        Set rs = .Execute
    End With
    
    With rs
        If .EOF = True Then
            .Close
            txtHrsWorked = 0
            txtAmount = 0
            txtMax = 0
            txtMin = 0
            txtBalance = 0
            txtBal.SetFocus
            Exit Sub
        End If
        txtSalary = !StartingSalary & ""
        
        rs1.Open "Select count(DT) from Holidays Where Month(DT)=" & DTPicker1.Month & " and Year(DT)=" & DTPicker1.year & " and day(dt)<" & TotalSalDays, con, adOpenForwardOnly, adLockReadOnly
                  
        If rs1.RecordCount > 0 Then
            Holidays = rs1.Fields(0) & ""
        End If
                  
        rs1.Close
               
        TotalSuns = TotalSundays(DTPicker1.Month, DTPicker1.year, TotalSalDays)
                                              
        Holidays = TotalSuns + Holidays
        
        Rate = Val(!StartingSalary & "") / (TotalSalDays - TotalSuns)
        Rate = Round(Rate / 8, 4)
        txtRate = Rate
        
        SalHrs = Val(![HrsWorked] & "") + ((Holidays - TotalSuns) * 8)
        
        'SalHrs = SalHrs + ((Holidays - TotalSuns) * 8)  'Add Holidays
        
        SalHrs = SalHrs + (Val(![Leaves] & "") * 8)
        SalHrs = SalHrs - (Val(![WPLeaves] & "") * 8)
        SalHrs = SalHrs - (Val(![SickLeaves] & "") * 4)
        If SalHrs > Val(txtWorkingHrs) Then SalHrs = Val(txtWorkingHrs)
        txtHrsWorked = Round(SalHrs)
        
        If bPerPiece Then
            txtAmount = Val(!TotalAmt & "")
        Else
            txtAmount = Round(SalHrs * Rate)
        End If
        
        If IsNull(![DedAmt]) Then lDedAmt = 0 Else lDedAmt = ![DedAmt]
        If IsNull(![totamt]) Then lTotAmt = 0 Else lTotAmt = ![totamt]
        If IsNull(![ClsAmt]) Then lClsAmt = 0 Else lClsAmt = ![ClsAmt]
        
        If lDedAmt > (lTotAmt - lClsAmt) Then
            LongLoan = (lTotAmt - lClsAmt)
        Else
            LongLoan = lDedAmt
        End If
        
        txtLongTerm = LongLoan
        txtLunchAmt = "0" 'Round(MonthDays * 10)
        txtPrevST = !Payed & ""
        
        txtBalance = Round(Val(txtAmount) - Val(txtLongTerm) - Val(txtLunchAmt) - Val(txtPrevST))
        txtMax = Round(txtAmount)
        txtMin = Round(Val(txtBalance) * 0.9)
        
    End With
    txtBal.SetFocus
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdBank_Click()

    Dim BankID As Double
    lblBankAcc = SelectBankAcc(BankID)
    lblBankAcc.Tag = BankID
    
    'Get ChqNo
         
    Dim ChqBookNo As Double
    Tbox(2) = GetNextChqNo(GetID(lblBankAcc), ChqBookNo)
    Tbox(2).Tag = ChqBookNo
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()

    If Saved Then
        Call DTPicker1_Change
        txtBal = ""
        cmbEmp.id = ""
        cmbEmp.SetFocus
        'cmbDept.SetFocus
    End If
    
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub

Private Sub DTDedStartMonth_Change()
    DTDedStartMonth = DateSerial(DTDedStartMonth.year, DTDedStartMonth.Month + 1, 0)
End Sub

Private Sub DTPicker1_Change()
    If cmbEmp.MatchFound Then
        Call cmbEmp_matched
    End If
    DTDedStartMonth.Value = DTPicker1.Value
    NextVchrNo = getNextVchrNo(DTPicker1, "BPV")
End Sub

Private Sub Form_Load()

    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
      
    DTPicker1 = GetServerDate
    Call DTPicker1_Change
    
    m_bContractors = False
    chkPrint = Abs(bPrintWhenSaved)
    
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyNumber(KeyAscii.Value, True)
End Sub

Private Sub txtDeduct_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyNumber(KeyAscii.Value, True)
End Sub

'To Check The Long Term Settings That Are Made For Departments
Private Function IsNotLongApplicable() As Boolean

    Dim Rec As New ADODB.Recordset
    'dim Item as New
    
    With Rec
        .CursorLocation = adUseClient
        .Open "SELECT JoinDate,LongTermTime,LongTermLimit FROM VOnlyEmps WHERE EmpID='" & cmbEmp.id & "'", con, adOpenKeyset, adLockReadOnly
        
        If DateDiff("m", !JoinDate, DTPicker1) < !longtermtime Then
            MsgBox "Could Not Give Long Term Loan To Employee." & vbNewLine & "It Can Be Given After " & !longtermtime & " Months", vbInformation
            IsNotLongApplicable = True
        ElseIf Val(txtAmt) > !LongTermLimit Then
            MsgBox "Could Not Give Long Term Loan To Employee. Because Amount " & vbNewLine & " Is Greater Than Long Term Limit.", vbInformation
            IsNotLongApplicable = True
        End If
        
    End With


End Function

Public Sub ShowMe(p_bContractors)
    
    m_bContractors = p_bContractors
    Dim strCond As String
    
    If m_bContractors Then
        strCond = "Where IsNull(TempDept,0)=1"
    Else
        strCond = "Where IsNull(TempDept,0)=0"
    End If
    
    With cmbDept
        .ClearVals
        .AddItem "<All Departments>", "0"
        .AddVals con, "Name", "Departments", "DeptID", strCond
        .Text = "<All Departments>"
    End With
    
    cmbCheckType.AddItem "Open"
    cmbCheckType.AddItem "Payees A/C Only"
    cmbCheckType.ListIndex = 0
    DTPicker2 = DTPicker1.Value
    Me.Show 1
    
End Sub

Private Function SavePaymentTrans(ChqSNo As Double) As String

    Dim i As Integer
    Dim NextVchrNo As String
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim dAmount As Double
    Dim sScNo As String
    Dim sCIH As String
    Dim Bal As Double
    Dim Desc As String
    Dim AccNo As String
    Dim Dated As Date
    'Get Serial # and Balances
    Dim dLongTerm As Double
    'dLongTerm = Val(txtFixLoan.Text)
    
   
    
    AccNo = strEmpShortTermAccNo
    Dated = DTPicker2
    
    NextVchrNo = getNextVchrNo(Dated, "BPV")
    INextSno = getNextSno(Dated)
    
    SnoToUse = INextSno
    
    'Party Account's Balance
    Bal = getBalance(AccNo, Dated)
    dAmount = Val(txtBal)
    
    sScNo = ChqSNo
    
    theDesc = "Long Term Advance Paid To " & cmbEmp.Text
    
    'Bank Account No.
    sCIH = GetID(lblBankAcc)
    
    Bal = Bal + dAmount
    Dim bWaitForAuth As Boolean
    bWaitForAuth = VoucherRequiresAuthorization("BPV")
    
    'Debit Entry
    bankDesc = Right(lblBankAcc, Len(lblBankAcc) - InStr(lblBankAcc, "]")) & " Chq # :" & Tbox(2)
    theDesc = theDesc & " " & bankDesc
        
    If bWaitForAuth = False Then
        con.Execute "INSERT INTO Vouchers Values(" & INextSno & ",'" & _
         Dated & "','" & NextVchrNo & "','" & AccNo & _
         "','" & theDesc & "'," & dAmount & ",0," & Bal & ",''," & sScNo & ")"
    
        Call adjustLedger(AccNo, Dated, INextSno, Val(dAmount))
    
        Dim ChqNo As String
        INextSno = INextSno + 1
        Bal = getBalance(sCIH, Dated)
        dAmount = Val(txtBal)
        Bal = Bal - dAmount
        'Description
    
        Desc = "Cheque Paid To " & cmbEmp.Text & " Chq# : " & Tbox(2)
        ChqNo = Tbox(2)
           
        'Credit Entry
    
        con.Execute "INSERT INTO Vouchers Values(" & INextSno & ",'" & _
         Dated & "','" & NextVchrNo & "','" & sCIH & _
         "','" & theDesc & "',0," & dAmount & "," & Bal & ",''," & sScNo & ")"
    
        Call adjustLedger(sCIH, Dated, INextSno, -Val(dAmount))
    
    Else
        con.Execute "INSERT INTO VouchersForAuthorization(VDate,VchrNo,Accno,Description,Debit,Credit,balance,DpstSlip,CSNo) VALUES('" & _
         Dated & "','" & NextVchrNo & "','" & AccNo & "','" & theDesc & "'," & dAmount & ",0," & Bal & ",''," & sScNo & ")"
       
        ChqNo = Tbox(2)
        
        con.Execute "INSERT INTO VouchersForAuthorization(VDate,VchrNo,Accno,Description,Debit,Credit,balance,DpstSlip,CSNo) VALUES('" & _
         Dated & "','" & NextVchrNo & "','" & sCIH & _
         "','" & theDesc & "',0," & dAmount & "," & Bal & ",''," & sScNo & ")"
    End If
    
    Dim rsCopy As New ADODB.Recordset
    With rsCopy
        .CursorLocation = adUseClient
        .Open "Select * from ChequeDetail Where (Sno=" & sScNo & ")", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            con.Execute "Insert into VoucherDetail Values('" & ![BalancedSno] & "','" & SnoToUse & "'," & ![Amount] & ")"
            .MoveNEXT
        Loop
        .Close
    End With
    Set rsCopy = Nothing
    
    con.Execute "UPDATE Cheque SET Posted=1 WHERE Sno=" & sScNo
    If bWaitForAuth = False Then
        con.Execute "INSERT INTO VoucherInfo(VchrNo,UserName,MachineName,DT,ApprovedBy) Values('" & NextVchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "','')"
    Else
        con.Execute "INSERT INTO VouchersForAuthorizationInfo(VchrNo,UserName,MachineName,DT,ApprovedBy,Notes,Payee) VALUES('" & NextVchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "','','','')"
    End If
    
    If NextVchrNo <> txtVchrNo Then
        MsgBox "Vchr No. " & txtVchrNo & " had been used, " & NextVchrNo & " is used Instead.", vbInformation
    End If
    
    SavePaymentTrans = NextVchrNo
    Exit Function
err:
    'MsgBox "Transaction Terminated Due To Following Error;" & vbNewLine & err.Description
    SavePaymentTrans = ""
End Function

Private Function SaveTransForCheque() As String

    On Error GoTo err
    'If CheckData = False Then Exit Function

    Dim i As Integer
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim Bal As Double
    Dim RcptNo As String
    
    RcptNo = "" 'IIf(Len(Trim(Tbox(3))) = 0, "", "Rcpt # :" & Tbox(3))
    
    'Get Serial # and Balances
    INextSno = getNextSno(DTPicker1.Value, "Cheque")
    SnoToUse = INextSno
    
    AccNo = "13-003-19001" 'Employee Short Term
    
    theDesc = "Short Term Advance Paid To " & cmbEmp.Text
    'Start New Transaction
    'con.BeginTrans
    bankDesc = Right(lblBankAcc, Len(lblBankAcc) - InStr(lblBankAcc, "]")) & " Chq # :" & Tbox(2)
    
    con.Execute "INSERT INTO Cheque(SNo,CDate,AccNo,Description,Amount,BankID,chequeno,ChqBookNo,chequeType,chequeDate,Posted,Payment,Bounced,BankAccNo)  Values(" & INextSno & ",'" & Format(DTPicker1.Value, "dd-MMM-yyyy") & "','" & _
     AccNo & "','" & theDesc & " (" & bankDesc & ") '," & txtBal & "," & lblBankAcc.Tag & ",'" & _
     Tbox(2) & "'," & Tbox(2).Tag & ",'" & cmbCheckType & "','" & DTPicker2 & "',0,1,0,'" & GetID(lblBankAcc) & "')"
    'Remove the Cheque From Cheque book
    Call IssueChq(Val(Tbox(2).Tag), Tbox(2))
    
'    With LVPymt.ListItems
'        For i = 1 To .Count
'            If .Item(i).Checked Then con.Execute "Insert into ChequeDetail Values('" & SnoToUse & "','" & Val(.Item(i).key) & "'," & Val(LVPymt.ListItems(i).SubItems(4)) & ")"
'        Next
'    End With
    
    'Commit Transaction
    'con.CommitTrans
    
    SaveTransForCheque = SavePaymentTrans(SnoToUse)

    Call PlaySaveSound
    'MsgBox "Transaction Saved Successfuly.", vbInformation
   
    Exit Function
err:
    MsgBox "Transaction Terminated Due to ERROR:" & err.Number & vbNewLine & err.Description, vbCritical
    SaveTransForCheque = ""
    con.RollbackTrans
    
End Function


