VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmGratuity_Calculation 
   Caption         =   "Over Time"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12840
   ClipControls    =   0   'False
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   8685
   ScaleMode       =   0  'User
   ScaleWidth      =   11500.96
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker DTAttend 
      Height          =   285
      Left            =   165
      TabIndex        =   1
      Top             =   315
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   503
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
      CustomFormat    =   "yyyy"
      Format          =   102301699
      CurrentDate     =   37055
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   8445
      TabIndex        =   0
      Top             =   195
      Visible         =   0   'False
      Width           =   4485
      _ExtentX        =   7911
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
   Begin VB.Frame FLV 
      Height          =   8055
      Left            =   165
      TabIndex        =   5
      Top             =   570
      Width           =   12435
      Begin VB.Frame FraBank 
         Caption         =   "Posting Details :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   120
         TabIndex        =   16
         Top             =   7110
         Visible         =   0   'False
         Width           =   3570
         Begin VB.CommandButton cmdPost 
            Caption         =   "Post"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2895
            TabIndex        =   27
            Top             =   240
            Width           =   570
         End
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
            Left            =   5295
            Picture         =   "frmGratuity_Calculation.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   405
            Visible         =   0   'False
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
            Height          =   300
            Index           =   2
            Left            =   75
            TabIndex        =   18
            Top             =   510
            Width           =   1215
         End
         Begin VB.CheckBox ChkClear 
            Caption         =   "Clear Cheque "
            Enabled         =   0   'False
            Height          =   225
            Left            =   5055
            TabIndex        =   17
            ToolTipText     =   "Clear This Cheque Automatically"
            Top             =   1245
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1350
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   315
            Left            =   3585
            TabIndex        =   20
            Top             =   1185
            Visible         =   0   'False
            Width           =   1410
            _ExtentX        =   2487
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   102301699
            CurrentDate     =   37250
         End
         Begin MSComCtl2.DTPicker DTChq 
            Height          =   285
            Left            =   1305
            TabIndex        =   28
            Top             =   510
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   503
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
            Format          =   102301699
            CurrentDate     =   37055
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
            Index           =   0
            Left            =   1305
            TabIndex        =   29
            Top             =   300
            Width           =   390
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
            Index           =   14
            Left            =   3570
            TabIndex        =   26
            Top             =   975
            Visible         =   0   'False
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
            Index           =   13
            Left            =   30
            TabIndex        =   25
            Top             =   930
            Visible         =   0   'False
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
            Index           =   12
            Left            =   105
            TabIndex        =   24
            Top             =   300
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
            Index           =   11
            Left            =   3765
            TabIndex        =   23
            Top             =   165
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label lblBankAcc 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   300
            Left            =   3750
            TabIndex        =   22
            Top             =   405
            Visible         =   0   'False
            Width           =   1515
         End
         Begin MSForms.ComboBox cmbCheckType 
            Height          =   315
            Left            =   15
            TabIndex        =   21
            Top             =   1185
            Visible         =   0   'False
            Width           =   3510
            VariousPropertyBits=   679495707
            DisplayStyle    =   7
            Size            =   "6191;556"
            ListWidth       =   2822
            MatchEntry      =   1
            ShowDropButtonWhen=   1
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5310
         TabIndex        =   12
         Top             =   1830
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtValue 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7860
         TabIndex        =   2
         Top             =   7110
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.TextBox txtTotal 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10980
         TabIndex        =   9
         Top             =   7110
         Width           =   1350
      End
      Begin VB.Frame FButs1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   570
         Left            =   2820
         TabIndex        =   6
         Top             =   7455
         Width           =   9525
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   7005
            TabIndex        =   30
            Top             =   135
            Width           =   1215
         End
         Begin VB.CommandButton cmdPrintBankLetter 
            Caption         =   "Bank Letter"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   885
            TabIndex        =   15
            Top             =   135
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton cmdPrintSlips 
            Caption         =   "Print Slips"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2130
            TabIndex        =   14
            Top             =   135
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton cmdPrintFilled 
            Caption         =   "Print Filled"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3375
            TabIndex        =   13
            Top             =   135
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "Print"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   5730
            TabIndex        =   11
            Top             =   135
            Width           =   1215
         End
         Begin MSForms.CommandButton cmdClose 
            CausesValidation=   0   'False
            Height          =   360
            Left            =   8280
            TabIndex        =   4
            Top             =   135
            Width           =   1215
            ForeColor       =   0
            Caption         =   " Close"
            PicturePosition =   327683
            Size            =   "2143;635"
            Accelerator     =   67
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   6915
         Left            =   120
         TabIndex        =   3
         Top             =   150
         Width           =   12225
         _ExtentX        =   21564
         _ExtentY        =   12197
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
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
            Text            =   "Emp ID"
            Object.Width           =   1561
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Employee Description"
            Object.Width           =   12171
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Designation"
            Object.Width           =   3358
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Salary"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Gratuity"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Search Emp. ID :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   6180
         TabIndex        =   10
         Top             =   7140
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total Amount :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   9675
         TabIndex        =   8
         Top             =   7140
         Width           =   1245
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7425
      Top             =   3120
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
            Picture         =   "frmGratuity_Calculation.frx":0372
            Key             =   "Present"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGratuity_Calculation.frx":07C4
            Key             =   "Absent"
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Year"
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
      Left            =   165
      TabIndex        =   7
      Top             =   60
      Width           =   1545
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete This Attendance"
      End
   End
End
Attribute VB_Name = "frmGratuity_Calculation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Editing As Boolean
Dim forLeave As Boolean
Dim iColNo As Integer
Dim iHolidayCount As Integer, iMonthDays As Integer

Public Sub enterLateHours()
    Load LateHours
    LateHours.Tag = LV.SelectedItem
    LateHours.Show
End Sub

Private Sub cmbEmp_matched()

On Error Resume Next
If cmbEmp.id = "0" Then Exit Sub
LV.ListItems(cmbEmp.id).Selected = True
LV.ListItems(cmbEmp.id).EnsureVisible

End Sub

Private Sub cmbShift_Change()
    Call FillList
End Sub

Private Sub cmdCheckAll_Click()
    
    
        For i = 1 To LV.ListItems.count
            With LV.ListItems(i)
                .Checked = True
                .SubItems(2) = "OverTime"
            End With
        
        Next i
    
End Sub

Private Sub cmdClearAll_Click()

    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            .Checked = False
            .SubItems(2) = ""
        End With
    Next i

End Sub

Private Sub cmbType_Change()
    Call CmbType_Click
End Sub

Private Sub CmbType_Click()
    Call FillList
End Sub

Private Sub cmbUnit_Change()
    Call cmbUnit_Click
End Sub

Private Sub cmbUnit_Click()
    Call FillList
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

Private Sub cmdClose_Click()
    Unload Me
End Sub


Public Sub cmdRefresh_Click()
    'Call RefreshLedger
    
    'Call Form_Resize
End Sub


Private Sub cmdEmpWorkDone_Click()
    frmEmpWorkDone.Show 1
End Sub

Private Sub cmdPrintAll_Click()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    If chkPrintInternal.Value = vbChecked Then
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptOverTimeInternal.rpt")
    Else
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptOverTime1.rpt")
    End If
    Dim f As New frmPrevRpt
    f.ShowReport " Month({OverTimeAuth.DT})=" & Month(DTAttend) & " AND Year({OverTimeAuth.DT})=" & year(DTAttend) & " ", rpt
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdPrintDaily_Click()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report

    Set rpt = rptApp.OpenReport(rptPayrollPath & "\DailyOverTime.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "Month({EmpOTHrs.DT})=" & Month(DTAttend) & " AND Year({EmpOTHrs.DT})=" & year(DTAttend) & " AND Day({EmpOTHrs.DT})=" & Day(DTAttend), rpt
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdPrintEmpWorkDone_Click()

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptOverTimeWork.rpt")
    
    Dim f As New frmPrevRpt
    f.ShowReport " Month({OverTimeAuth.DT})=" & Month(DTAttend) & " AND Year({OverTimeAuth.DT})=" & year(DTAttend) & " ", rpt
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub cmdPost_Click()

    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(*)", "Employees_AdvanceSalary_Posting", " WHERE DT='" & DTAttend.Value & "'")
    
    If lCount > 0 Then
        MsgBox "Vouchers & Advances have been posted.", vbInformation
        Exit Sub
    End If
     
'    If strEmpShortTermAccNo = "" Then
'        MsgBox "Can't Post, Emp. Short Term Loan Account Not Found.", vbInformation
'        Exit Sub
'    End If
    
'    If lblBankAcc = "" Then
'        MsgBox "Invalid Bank Selected.", vbInformation
'        Exit Sub
'    ElseIf Not ValidChqNo(GetID(lblBankAcc), Tbox(2), False, Val(Tbox(2).Tag)) Then
'        If Issued Then
'            MsgBox "The Cheque With Number '" & txtChqNo & "' Has Been Issued.", vbInformation
'        Else
'            MsgBox "Invalid Cheque No.", vbInformation
'        End If
'        Tbox(2).SetFocus
'        Exit Sub
'    End If
    

        
'    If MsgBox("This will Post Bank & Cash Vouchers. Are you sure?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
'
'    Dim strCashVchrNo As String, strBankVchrNo As String
'    Dim lCashAmt As Long, lBankAmt As Long
'    lCashAmt = GetSingleLongValue("SUM(Approved)", "Employees_AdvanceSalary", " WHERE DT='" & DTAttend.Tag & "' AND EmpID IN(SELECT EmpID FROM Employees WHERE BankPymt=0)")
'    lBankAmt = GetSingleLongValue("SUM(Approved)", "Employees_AdvanceSalary", " WHERE DT='" & DTAttend.Tag & "' AND EmpID IN(SELECT EmpID FROM Employees WHERE BankPymt=1)")
'
    Call StartTrans(con)
'
'    Dim strDescription As String
'    Dim objVchr As New clsVouchers
'    objVchr.VoucherDate = Date
'    objVchr.VoucherType = "CPV"
'
'    strDescription = "Advance Salary of " & Format(DTAttend, "MMM, yyyy")
'    objVchr.AddVoucherHead strEmpShortTermAccNo, Val(lCashAmt), strDescription
'    objVchr.AddVoucherHead "25-005-03001", -Val(lCashAmt), strDescription
'    strCashVchrNo = objVchr.PostVoucher(False, True)
'
'    strBankVchrNo = SaveTransForCheque(lBankAmt)
    Dim strDescription As String
    strDescription = "Advance Salary"
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If Val(.SubItems(7)) > 0 Then
                con.Execute "INSERT INTO Advances(EmpID,DT,Type,Description,Amount,DAmount,AmountCleared,DeductionStartDT,Cleared,ActualAmountTaken) VALUES('" & _
                 .key & "','" & Date & "',0,'" & strDescription & "'," & Val(.SubItems(7)) & ",0,0,'" & DTAttend.Value & "',0," & Val(.SubItems(7)) & ")"
            End If
        End With
    Next
    
    con.Execute "INSERT INTO Employees_AdvanceSalary_Posting(DT,CashVchrNo,BankVchrNo,UserName,MachineName,ChqNo) VALUES('" & DTChq.Value & _
     "','" & strCashVchrNo & "','" & strBankVchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & Tbox(2) & "')"
    
    con.CommitTrans
    
End Sub

Private Sub cmdPrint_Click()
    Call PrintAdvanceReport(0)
End Sub

Private Sub PrintAdvanceReport(Optional iReportType As Integer = 0)
    
    Dim rpt As CRAXDDRT.Report, strSelection As String
    strSelection = ""
    If iReportType = 0 Then
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\AdvanceSalaryReport.rpt")
    ElseIf iReportType = 1 Then
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\AdvanceSalaryReport_Filled.rpt")
    ElseIf iReportType = 2 Then
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\AdvanceSalarySlip.rpt")
        strSelection = "{Employees_AdvanceSalary_SP;1.Approved}>0"
    ElseIf iReportType = 3 Then
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\AdvanceSalaryBankLetter.rpt")
        Dim strChqNo As String, strVchrNo As String
        strVchrNo = GetSingleStringValue("BankVchrNo", "Employees_AdvanceSalary_Posting", " WHERE DT='" & DTAttend.Value & "'")
        strChqNo = GetSingleStringValue("ChequeNo", "Cheque INNER JOIN Vouchers ON Vouchers.CSno=Cheque.SNo", " WHERE VchrNo='" & strVchrNo & "'")
        rpt.FormulaFields.GetItemByName("ChqNo").Text = "'" & strChqNo & "'"
        strSelection = "{Employees_AdvanceSalary_SP;1.Approved}>0"
    End If
    
    rpt.ParameterFields.GetItemByName("@DT").SetCurrentValue CDate(DTAttend.Value)
     
    If iReportType = 3 Then
        rpt.ParameterFields.GetItemByName("@BankPymt").SetCurrentValue True
    Else
        If cmbType.ListIndex = 0 Then
            rpt.ParameterFields.GetItemByName("@BankPymt").SetCurrentValue False
        ElseIf cmbType.ListIndex = 1 Then
            rpt.ParameterFields.GetItemByName("@BankPymt").SetCurrentValue True
        End If
    End If
    
    If cmbUnit.ListIndex = 1 Then
        rpt.ParameterFields.GetItemByName("@Unit").SetCurrentValue "U1"
    ElseIf cmbUnit.ListIndex = 2 Then
        rpt.ParameterFields.GetItemByName("@Unit").SetCurrentValue "U2"
    End If
    
    rpt.EnableParameterPrompting = False
    
    rpt.FormulaFields.GetItemByName("iMonthDays").Text = iMonthDays
    rpt.FormulaFields.GetItemByName("iHolidayCount").Text = iHolidayCount
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt, , False
    
End Sub

Private Sub cmdPrintBankLetter_Click()
    Call PrintAdvanceReport(3)
End Sub

Private Sub cmdPrintFilled_Click()
    Call PrintAdvanceReport(1)
End Sub

Private Sub cmdPrintSlips_Click()
    Call PrintAdvanceReport(2)
End Sub

Private Sub cmdSave_Click()
    
    Dim lCount As Long
    Dim lYear As Long, i As Integer
    lYear = Format(DTAttend.Value, "yyyy")
    lCount = GetSingleLongValue("COUNT(*)", "Employees_Gratuity", " WHERE GratuityYear=" & lYear & " AND Posted=1")
    If lCount > 0 Then
        MsgBox "Vouchers Posted.", vbInformation
        Exit Sub
    End If
    
    Dim lEntryID As Long
    Call StartTrans(con)
    con.Execute "DELETE FROM Employees_Gratuity WHERE GratuityYear=" & lYear
    con.Execute "INSERT INTO Employees_Gratuity(GratuityYear) VALUES(" & lYear & ")"
    lEntryID = GetSingleLongValue("MAX(EntryID)", "Employees_Gratuity")
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            con.Execute "INSERT INTO Employees_Gratuity_Detail(RefID,EmpID,GratuityAmount) VALUES(" & lEntryID & ",'" & .ListSubItems(1).Tag & "'," & Val(.SubItems(4)) & ")"
        End With
    Next
    con.CommitTrans
    
End Sub

Private Sub DTAttend_Change()
        
    'Call cmdSave_Click
    DTAttend.Tag = DateSerial(DTAttend.year, 12, 31)
    'DTPicker2 = DTAttend.Value
    Call FillList
    
End Sub

Private Sub Form_Load()

    DTAttend = Format(GetServerDate, "dd-MMM-yyyy")
    Call DTAttend_Change
    
    DTChq = Date
    
    cmbDept.ListHeight = 6000
    
    
    cmbDept.ClearVals
    'cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID", " Where TempDept=0"
    cmbDept.id = 0
    
    'DTAttend.Enabled = UserHasAccess("AttDateChange")
        
        
    'DTPicker2 = DTAttend.Value
    
    
End Sub

Private Sub cmbDept_matched()

    Call FillList
    
'    cmbEmp.ClearVals
''    cmbEmp.AddItem "<All Employees>", "0"
'    If cmbDept.Text = "<All Departments>" Then
'        cmbEmp.AddVals con, "Name + ' {' + EmpID + '}'", "Employees", "EmpID", "Where active=1"
'    Else
'        cmbEmp.AddVals con, "Name + ' {' + EmpID + '}'", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' and active=1 "
'    End If
'    cmbEmp.ID = "0"
    
End Sub

Private Sub FillList()

    'If cmbDept.MatchFound = False Then Exit Sub
    
    LV.ListItems.Clear
    
    Dim ITM As ListItem
    Dim DTJoin As Date, DTDiscontinued As Date
    DTJoin = DateSerial(DTAttend.year, 7, 1)
    DTDiscontinued = DateSerial(DTAttend.year, 1, 1)
    
    Dim rs As New adodb.Recordset, strCondition As String
    strCondition = " WHERE JoinDate<'" & DTJoin & "' AND (DiscontinuedOn IS NULL OR DiscontinuedOn>'" & DTDiscontinued & "')"
    
    Dim DTEmpJoin As Date, DTEmpDisContinued As Date, iMonths As Integer, lSalary As Long, iDays As Integer
    Dim dRate As Double, dGratuityAmt As Double, lYear As Long
    lYear = DTAttend.year
    With rs
        .Open "SELECT VEmp.EmpID,Name,Rel,FName,Designation,DeptName,BSal,JoinDate,DiscontinuedOn,StartingSalary,GratuityAmount FROM VEmp " & _
         " LEFT OUTER JOIN (SELECT EmpID,BSal FROM MonthlySalaries WHERE EntryID IN(SELECT MAX(EntryID) FROM MonthlySalaries GROUP BY EmpID)) TEmpLastSal ON VEmp.EmpID=TEmpLastSal.EmpID " & _
         " LEFT OUTER JOIN (SELECT EmpID,GratuityAmount FROM Employees_Gratuity_Detail INNER JOIN Employees_Gratuity ON Employees_Gratuity.EntryID=Employees_Gratuity_Detail.RefID WHERE GratuityYear=" & lYear & ") TG ON VEmp.EmpID=TG.EmpID " & strCondition, con, adOpenForwardOnly, adLockReadOnly
        
        LV.ListItems.Clear
        Do Until .EOF
        
            Set ITM = LV.ListItems.add(, , !EmpId & "")
            
            ITM.ListSubItems.add(, , !Name & " " & !Rel & " " & !FName).Tag = !EmpId & ""
            ITM.ListSubItems.add , , !Designation & ""
             
            If Val(!BSal & "") = 0 Then
                lSalary = Val(!StartingSalary & "")
            Else
                lSalary = Val(!BSal & "")
            End If
            dRate = lSalary / 12
            
            If !JoinDate <= DateSerial(DTAttend.year, 1, 31) Then
                DTEmpJoin = DateSerial(DTAttend.year, 1, 1)
            Else
                DTEmpJoin = !JoinDate
            End If
            
            If IsNull(!DiscontinuedOn) Then
                DTEmpDisContinued = DateSerial(DTAttend.year, 12, 31)
            Else
                If !DiscontinuedOn >= DateSerial(DTAttend.year, 12, 31) Then
                    DTEmpDisContinued = DateSerial(DTAttend.year, 12, 31)
                Else
                    DTEmpDisContinued = !DiscontinuedOn
                End If
            End If
             
            Call DateDiffInMonthsNDays(DTEmpJoin, DTEmpDisContinued, iMonths, iDays)
            
            
            ITM.ListSubItems.add , , lSalary
            If Val(!GratuityAmount & "") > 0 Then
                dGratuityAmt = Val(!GratuityAmount & "")
            Else
                dGratuityAmt = dRate * iMonths
                dGratuityAmt = Round(dGratuityAmt)
            End If
            ITM.ListSubItems.add , , dGratuityAmt
        

            .MoveNext
        Loop
        .Close
'
'        .Open "SELECT * FROM Employees_AdvanceSalary WHERE DT='" & DTAttend.Tag & "' ORDER BY EmpID", con, adOpenForwardOnly, adLockReadOnly
'
'        Do Until .EOF
'            On Error Resume Next
'            With LV.ListItems(!EmpId & "")
'                If err.Number = 35601 Then GoTo MoveNEXT
'                .Tag = rsList!EntryID
'                .SubItems(4) = Val(rs!Demand & "")
'            End With
'MoveNEXT:
'            .MoveNEXT
'        Loop
'        .Close
        
    End With
    Set rs = Nothing
    Call GetTotalAmount
    
End Sub

Private Sub AddEmps(ITM As ListItem, rs As adodb.Recordset, Shift As String)
    With rs
        Set ITM = LV.ListItems.add(, ![EmpId] & Shift, ![EmpId] & "", , ICO)
        ITM.SubItems(1) = ![Name]
        ITM.SubItems(2) = ![FName]
        ITM.SubItems(3) = Shift
        ITM.SubItems(4) = 0
    End With
    'Itm.SubItems(4) = Format(![DT], "dddd dd MMMM ,yyyy")
        
End Sub




Private Sub Form_Resize()
    
    Exit Sub

    On Error Resume Next

    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2

    FLV.Move 50, 1230, Me.ScaleWidth - 100, Me.ScaleHeight - 1400

    LV.Move 100, 200, Me.Width - 400, FLV.Height - FButs1.Height - 300

    FButs1.Move LV.Left + LV.Width - FButs1.Width, LV.Top + LV.Height

    fButs2.Move LV.Left, LV.Top + LV.Height
    
End Sub




Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    Dim lCount As Long
    'lCount = GetSingleLongValue("COUNT(*)", "Employees_AdvanceSalary_Posting", " WHERE DT='" & DTAttend.Value & "'")
    
    If lCount > 0 Then
        MsgBox "Vouchers & Advances have been posted. Editing is Locked Now.", vbInformation
        Exit Sub
    End If
    
    iColNo = 5
    
    With txtEdit
        .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(iColNo).Width
        .Text = Val(LV.SelectedItem.SubItems(iColNo - 1))
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
    
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If x > Val(LV.ColumnHeaders(i).Left) And x < (Val(LV.ColumnHeaders(i).Left) + Val(LV.ColumnHeaders(i).Width)) Then
            iColNo = i
            Exit For
        End If
    Next
End Sub

Private Sub mnuDelete_Click()

con.Execute "DELETE FROM AttendanceSheet WHERE DT='" & DTAttend.Value & "' AND EmpID='" & LV.SelectedItem.Text & "' AND Shift='" & cmbTime.Tag & "'"
Call FillList

End Sub

Private Sub TFrom_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then
        If TTo.Visible Then TTo.SetFocus Else txtHours.SetFocus
    ElseIf KeyCode = 27 Then
        TFrom.Visible = False
        TTo.Visible = False
        txtHours.Visible = False
    End If
    
End Sub

Private Sub TTo_KeyDown(KeyCode As Integer, Shift As Integer)


Dim LHours As Integer, LMinute As Integer
Dim OHours As Integer, OMinute As Integer, Fine As Double, ExtraHrs As Double

Dim DefinedWorkingMins As Long, EmpWorkingMins As Long
Dim MinDiff As Long

If KeyCode = 13 Then
   
    
    With LV.SelectedItem

        .Checked = True
        
        .SubItems(2) = Format(TFrom.Value, "HH:mm")
        .SubItems(3) = Format(TTo.Value, "HH:mm")
        
        TTo.Visible = False
        TFrom.Visible = False
    
        LV.SetFocus
    End With
End If

End Sub


Private Sub txtEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        LV.SelectedItem.SubItems(iColNo - 1) = Val(txtEdit)
        txtEdit.Visible = False
        Call GetTotalAmount
        
        Dim iRowAffect As Integer, strFieldName As String
         
        'con.Execute "UPDATE Employees_AdvanceSalary SET " & strFieldName & "=" & Val(txtEdit) & " WHERE EmpID='" & LV.SelectedItem.key & "' AND DT='" & DTAttend.Value & "'", iRowAffect
       
'        If iRowAffect = 0 Then
'            con.Execute "INSERT INTO Employees_AdvanceSalary(EmpID,DT," & strFieldName & ") VALUES('" & LV.SelectedItem.key & "','" & DTAttend.Value & "'," & Val(txtEdit) & ")"
'        End If
        
'        If LV.SelectedItem.Index < LV.ListItems.count Then
'            LV.ListItems(LV.SelectedItem.Index + 1).Selected = True
'            LV.SelectedItem.EnsureVisible
'            Call LV_DblClick
'        Else
'            txtEdit.Visible = False
'        End If
        
    ElseIf KeyAscii = 27 Then
        txtEdit.Visible = False
    End If
    
End Sub

Private Sub txtEdit_LostFocus()
    txtEdit.Visible = False
End Sub

Private Sub txtvalue_Change()
    

    Dim StartIdx As Integer
    
    
    StartIdx = 1
    
    Dim fldValue As String

    For i = StartIdx To LV.ListItems.count
       
       
     
        fldValue = LV.ListItems(i)
    
        If LCase(fldValue) Like "*" & LCase(Trim(txtValue)) & "*" Then
            LV.ListItems(i).Selected = True
            LV.ListItems(i).EnsureVisible
            If i <> StartIdx Then Exit For
        End If
       
    Next

    If i = LV.ListItems.count + 1 Then
        MsgBox "Item Not Found.", vbInformation
        Set LV.SelectedItem = Nothing
    End If
    
    Exit Sub

End Sub

Private Sub txtEmpID_KeyPress(KeyAscii As Integer)
    
End Sub

Private Sub txtHours_KeyPress(KeyAscii As Integer)

    Dim LHours As Integer, LMinute As Integer
    Dim OHours As Integer, OMinute As Integer, Fine As Double, ExtraHrs As Double

    Dim DefinedWorkingMins As Long, EmpWorkingMins As Long
    Dim MinDiff As Long
    
    Dim sTto As String
    Dim iAdd As Integer
    Dim iTemp As Integer
    
    If KeyAscii = 13 Then
'        If Val(txtHours) <= 0 Or Val(txtHours) + Val(LV.SelectedItem.SubItems(5)) > 12 Or Val(txtHours) > 6 Then
'            MsgBox "Invalid Entry.", vbInformation
'            Exit Sub
'        End If
        If Val(txtHours) <= 0 Then
            MsgBox "Invalid Entry.", vbInformation
            Exit Sub
'        ElseIf Val(txtHours) > 3 Then
'            MsgBox "Invalid Entry.", vbInformation
'            Exit Sub
        End If
        
        With LV.SelectedItem
            
            .Checked = True
        
            '.SubItems(2) = Format(TFrom.value, "HH:mm")
            'sTto = Format(DateAdd("n", Val(txtHours) * 60, Format(TFrom.value, "HH:mm")), "HH:mm")
            'sTto = Val(Left(Format(TFrom.value, "HH:mm"), 2)) + Val(txtHours) & Right(Format(TFrom.value, "HH:mm"), 3)
            'iAdd = Int(Rnd * (Val(5) - (Val(1) - 1)) + (Val(1)))
            
            'iTemp = Val(Right(sTto, 2)) + iAdd
            
            'If iTemp > 59 Then
           '     sTto = (Val(Left(sSecOut, 2)) + 1) & ":" & Format((iAdd - 60), "00")
          '  Else
         '       sTto = Left(sTto, 3) & Format(iTemp, "00")
        '    End If
            
            .SubItems(3) = Val(txtHours) 'Format(TTo.value, "HH:mm")
            .SubItems(4) = Val(txtHours) * Val(.SubItems(2))
            'TTo.Visible = False
            'TFrom.Visible = False
            txtHours.Visible = False
            
            LV.SetFocus
            GetTotalOTAmt
        End With
        
    ElseIf KeyAscii = 27 Then
        'TTo.Visible = False
        'TFrom.Visible = False
        txtHours.Visible = False
    End If

End Sub

Private Sub GetTotalOTAmt()

    Dim i As Integer
    Dim dTotal As Double
    For i = 1 To LV.ListItems.count
        dTotal = dTotal + Val(LV.ListItems(i).SubItems(4))
    Next
    txtTotal = dTotal
End Sub

Private Sub txtValue_GotFocus()
    With txtValue
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Function SaveTransForCheque(lAmt As Long) As String

    On Error GoTo err
    'If CheckData = False Then Exit Function

    Dim i As Integer
    Dim INextSno As Double
    Dim SnoToUse As Double
    Dim Bal As Double
    Dim RcptNo As String
    
    RcptNo = "" 'IIf(Len(Trim(Tbox(3))) = 0, "", "Rcpt # :" & Tbox(3))
    
    'Get Serial # and Balances
    INextSno = getNextSno(Date, "Cheque")
    SnoToUse = INextSno
    
    AccNo = strEmpShortTermAccNo
    
    theDesc = "Advance Salary of " & Format(DTAttend, "MMM, yyyy") & " Chq # :" & Tbox(2)
    'Start New Transaction
    'con.BeginTrans
    bankDesc = theDesc 'Right(lblBankAcc, Len(lblBankAcc) - InStr(lblBankAcc, "]")) & " Chq # :" & Tbox(2)
    
    con.Execute "INSERT INTO Cheque (SNo,CDate,AccNo,Description,Amount,BankID,chequeno,ChqBookNo,chequeType,chequeDate,Posted,Payment,Bounced,BankAccNo)  Values(" & INextSno & ",'" & Format(Date, "dd-MMM-yyyy") & "','" & _
    AccNo & "','" & theDesc & " (" & bankDesc & ") '," & lAmt & "," & lblBankAcc.Tag & ",'" & _
    Tbox(2) & "'," & Tbox(2).Tag & ",'" & cmbCheckType & "','" & DTPicker2 & "',0,1,0,'" & GetID(lblBankAcc) & "')"
 
    Call IssueChq(Val(Tbox(2).Tag), Tbox(2))
    
    SaveTransForCheque = SavePaymentTrans(SnoToUse, lAmt)

    Call PlaySaveSound
  
   
    Exit Function
err:
    MsgBox "Transaction Terminated Due to ERROR:" & err.Number & vbNewLine & err.Description, vbCritical
    SaveTransForCheque = ""
    
    
End Function

Private Function SavePaymentTrans(ChqSNo As Double, lAmt As Long) As String

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
   
    Dim dLongTerm As Double
    
    AccNo = strEmpShortTermAccNo
    
    Dated = DTPicker2
    
    NextVchrNo = getNextVchrNo(Dated, "BPV")
    INextSno = getNextSno(Dated)
    
    SnoToUse = INextSno
    
    'Party Account's Balance
    Bal = getBalance(AccNo, Dated)
    dAmount = lAmt
    
    sScNo = ChqSNo
    
    theDesc = "Advance Salary of " & Format(DTAttend, "MMM, yyyy") & " Chq # :" & Tbox(2)
    
    'Bank Account No.
    sCIH = GetID(lblBankAcc)
    
    Bal = Bal + dAmount
    'Debit Entry
    bankDesc = theDesc 'Right(lblBankAcc, Len(lblBankAcc) - InStr(lblBankAcc, "]")) & " Chq # :" & Tbox(2)
    'theDesc = theDesc & " " & bankDesc
     
    
    con.Execute "INSERT INTO Vouchers Values(" & INextSno & ",'" & _
     Dated & "','" & NextVchrNo & "','" & AccNo & _
     "','" & theDesc & "'," & dAmount & ",0," & Bal & ",''," & sScNo & ")"
    
        Call adjustLedger(AccNo, Dated, INextSno, Val(lAmt))
    
    Dim ChqNo As String
    INextSno = INextSno + 1
    Bal = getBalance(sCIH, Dated)
    dAmount = lAmt
    Bal = Bal - dAmount
        'Description
    
    Desc = theDesc '"Cheque Paid To " & cmbEmp.Text & " Chq# : " & Tbox(2)
    ChqNo = Tbox(2)
           
        'Credit Entry
    
    con.Execute "INSERT INTO Vouchers Values(" & INextSno & ",'" & _
     Dated & "','" & NextVchrNo & "','" & sCIH & _
     "','" & theDesc & "',0," & dAmount & "," & Bal & ",''," & sScNo & ")"
    
    Call adjustLedger(sCIH, Dated, INextSno, -Val(dAmount))
   
    con.Execute "UPDATE Cheque Set Posted=1 WHERE Sno=" & sScNo
    
    Dim myDT As Date, dpAmount As Double, DTLastReconciliation As Date
    
    con.Execute "INSERT INTO VoucherInfo(VchrNo,UserName,MachineName,DT) VALUES('" & NextVchrNo & "','" & CurrentUserName & "','" & strComputerName & "','" & GetServerDate(True) & "')"
   
    SavePaymentTrans = NextVchrNo
    Exit Function
err:
    'MsgBox "Transaction Terminated Due To Following Error;" & vbNewLine & err.Description
    SavePaymentTrans = ""
End Function


Private Sub GetTotalAmount()

    Dim i As Integer, lTotalAmt As Long
    For i = 1 To LV.ListItems.count
        lTotalAmt = lTotalAmt + Val(LV.ListItems(i).SubItems(4))
    Next
    txtTotal = lTotalAmt
    
End Sub
