VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmLeavesOld 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Leaves / Absent"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
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
   ScaleHeight     =   3150
   ScaleWidth      =   6870
   StartUpPosition =   1  'CenterOwner
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   135
      TabIndex        =   0
      Top             =   330
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Height          =   285
      Left            =   2265
      TabIndex        =   1
      Top             =   330
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin VB.OptionButton OptType 
      Caption         =   "Absent"
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
      Height          =   285
      Index           =   0
      Left            =   1965
      TabIndex        =   4
      Top             =   1365
      Width           =   900
   End
   Begin VB.OptionButton OptType 
      Caption         =   "Leave"
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
      Height          =   285
      Index           =   1
      Left            =   135
      TabIndex        =   3
      Top             =   1350
      Value           =   -1  'True
      Width           =   840
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1680
      Left            =   105
      TabIndex        =   14
      Top             =   1695
      Width           =   4905
      Begin VB.OptionButton opts 
         Caption         =   "Without Deduction"
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
         Height          =   285
         Index           =   0
         Left            =   1695
         TabIndex        =   21
         Top             =   1410
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.OptionButton OptFullDay 
         Caption         =   "Full Day Leave"
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
         Height          =   255
         Left            =   30
         TabIndex        =   27
         Top             =   15
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton OptHalfDay 
         Caption         =   "Half Day Leave"
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
         Height          =   255
         Left            =   1845
         TabIndex        =   26
         Top             =   15
         Visible         =   0   'False
         Width           =   1725
      End
      Begin ComboList.Usercontrol1 cmbAuth 
         Height          =   285
         Left            =   2490
         TabIndex        =   19
         Top             =   525
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   503
         BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   -1
      End
      Begin VB.OptionButton opts 
         Caption         =   "With Deduction"
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
         Height          =   285
         Index           =   1
         Left            =   15
         TabIndex        =   20
         Top             =   1395
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.CheckBox chkFull 
         Caption         =   "Full Day Leave"
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
         Height          =   240
         Left            =   3690
         TabIndex        =   5
         Top             =   45
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Frame Fra2 
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   45
         TabIndex        =   15
         Top             =   285
         Width           =   2415
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Leave Type"
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
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   2415
         End
         Begin MSForms.ComboBox cmbLeaveType 
            Height          =   285
            Left            =   0
            TabIndex        =   6
            Top             =   240
            Width           =   2415
            VariousPropertyBits=   746604571
            BorderStyle     =   1
            DisplayStyle    =   7
            Size            =   "4260;503"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin VB.Frame Fra1 
         BorderStyle     =   0  'None
         Height          =   570
         Left            =   -60
         TabIndex        =   17
         Top             =   285
         Visible         =   0   'False
         Width           =   2565
         Begin MSComCtl2.DTPicker DTFrom 
            Height          =   315
            Left            =   90
            TabIndex        =   8
            Top             =   480
            Width           =   1200
            _ExtentX        =   2117
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
            CustomFormat    =   "hh:mm  tt"
            Format          =   64356355
            UpDown          =   -1  'True
            CurrentDate     =   37384
         End
         Begin MSComCtl2.DTPicker DTTo 
            Height          =   315
            Left            =   1320
            TabIndex        =   9
            Top             =   240
            Width           =   1230
            _ExtentX        =   2170
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
            CustomFormat    =   "hh:mm  tt"
            Format          =   64356355
            UpDown          =   -1  'True
            CurrentDate     =   37384
         End
         Begin VB.Label lblFrom 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   " From :                 To :"
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
            Left            =   120
            TabIndex        =   18
            Top             =   0
            Width           =   2400
         End
      End
      Begin MSForms.TextBox txtDesc 
         Height          =   315
         Left            =   30
         TabIndex        =   24
         Top             =   1080
         Width           =   4860
         VariousPropertyBits=   -1467987941
         ScrollBars      =   3
         Size            =   "8572;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Leave Reason"
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
         Index           =   6
         Left            =   30
         TabIndex        =   23
         Top             =   855
         Width           =   4845
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Authority Person :"
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
         Left            =   2505
         MouseIcon       =   "frmLeavesOld.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   22
         Top             =   285
         Width           =   2385
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
      Height          =   255
      Left            =   5040
      TabIndex        =   13
      Top             =   2670
      Visible         =   0   'False
      Width           =   1800
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   135
      TabIndex        =   2
      Top             =   900
      Width           =   2130
      _ExtentX        =   3757
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
      CustomFormat    =   "ddd dd-MMM yyyy"
      Format          =   64356355
      CurrentDate     =   37384
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   2265
      TabIndex        =   28
      Top             =   900
      Width           =   2745
      _ExtentX        =   4842
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
      CustomFormat    =   "ddd dd-MMM yyyy"
      Format          =   64356355
      CurrentDate     =   37384
   End
   Begin MSForms.CommandButton cmdPrintLeaves 
      Height          =   360
      Left            =   5085
      TabIndex        =   29
      Top             =   1380
      Width           =   1710
      ForeColor       =   0
      Caption         =   "Print Balances"
      PicturePosition =   327683
      Size            =   "3016;635"
      Accelerator     =   76
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " From                                     To"
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
      Height          =   240
      Index           =   0
      Left            =   150
      TabIndex        =   25
      Top             =   645
      Width           =   4845
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Select Department            Select Employee"
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
      Left            =   135
      TabIndex        =   12
      Top             =   90
      Width           =   4845
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   360
      Left            =   5100
      TabIndex        =   10
      Top             =   525
      Width           =   1695
      Caption         =   "Save & Close"
      PicturePosition =   327683
      Size            =   "2990;635"
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
      Left            =   5085
      TabIndex        =   7
      Top             =   105
      Width           =   1725
      Caption         =   "Save & New    "
      PicturePosition =   327683
      Size            =   "3043;635"
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
      Left            =   5085
      TabIndex        =   11
      Top             =   945
      Width           =   1710
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "3016;635"
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
Attribute VB_Name = "frmLeavesOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public edit As Boolean

Private Sub PrintReport()

    Screen.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report, ID As Integer
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptLeaveSlip.rpt")
    
    ID = Val(con.Execute("Select MAX(EntryID) FROM Leaves").Fields(0) & "")
    
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
    
    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & company & "'"
        End If
    Next
    
    'rpt.SQLQueryString = Sql
    Cond = "{VLeaves.EntryID}=" & ID & ""
    
    Me.Hide
    f.ShowReport Cond & " ", rpt
    
    Screen.MousePointer = vbDefault


End Sub


Private Sub cmbdept_matched()
    
    cmbEmp.ClearVals
    
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where Active=1"
    Else
        cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' and ACtive=1"
    End If
    
End Sub

Private Function Saved() As Boolean

    Dim Affect As Long, First As Integer, Second As Integer, IsFirst As Boolean, IsSecond As Boolean
    Dim Full As Integer, Hrs As Double, TFrom As String, TTo As String
    Dim PrevCasualLeaves  As Integer
    Saved = False
    
    'On Error GoTo err
    Dim SQL As String, Cond As String, Deduct As Integer
            
    Dim iLoop As Integer
    Dim iTotal As Integer
    Dim iTotalLimit As Integer
    Dim iTotalTaken As Integer
    Dim DT As Date
    Dim StartDT As Date
    Dim FinalDT As Date
    
    
    If Not cmbEmp.MatchFound Then
        MsgBox "Please Select Valid Employee From The List.", vbInformation
        Exit Function
    End If
    If DTPicker2.value < DTPicker1.value Then
        MsgBox "Invalid Date Range Selected.", vbInformation
        Exit Function
    End If
    
    If IsProbation(cmbEmp.ID, DTPicker1.value) Then
        MsgBox "Employee is still in his Probation. Can't Take Leave.", vbInformation
        Exit Function
    End If
    
    iLoop = DateDiff("d", DTPicker1, DTPicker2)
    iTotal = iLoop + 1
    
    Cond = " Where DT Between '" & DTPicker1 & "' AND '" & DTPicker2 & "' AND EmpID='" & cmbEmp.ID & "'"
      
    If OptType(1) Then
        If Not cmbEmp.MatchFound Then
                MsgBox "Please Select Valid Employee From The List.", vbInformation
                Exit Function
        ElseIf Not cmbAuth.MatchFound Then
            MsgBox "Authoriy Person That You Entered Does Not Exits !", vbInformation
            'cmbAuth.SetFocus
            Exit Function
        'ElseIf IsPrevTime And chkFull = vbUnchecked Then
        '    MsgBox "You Have Already Taken Leave For This Time !", vbInformation
        '    DTFrom.SetFocus
        '    Exit Function
        ElseIf cmbAuth.ID = cmbEmp.ID Then
            MsgBox "Employee And Authority Person Could Not Be Same !", vbInformation
            cmbAuth.SetFocus
            Exit Function
        'ElseIf DTFrom.value >= DTTo.value And chkFull = vbUnchecked Then
        '    MsgBox "Out Value Could Not Be Greater/Equal In Value  !", vbInformation
        '    DTFrom.SetFocus
        '    Exit Function
        ElseIf cmbLeaveType.ListIndex < 0 Then
            MsgBox "Please Select Leave Type", vbInformation
            cmbLeaveType.SetFocus
            Exit Function
        End If
    
        If Month(DTPicker1) < 7 Then
            StartDT = DateSerial(Year(DTPicker1) - 1, 7, 1) ' '7/1/' + Cast(Year(@DT)-1 As Varchar)
        Else
            StartDT = DateSerial(Year(DTPicker1), 7, 1)
        End If
        FinalDT = DateSerial(Year(StartDT) + 1, 6, 30)

        
        Dim FieldName As String
              
        If OptFullDay Then
            FieldName = "," & cmbLeaveType.List(cmbLeaveType.ListIndex, 1)
        Else
            FieldName = ""
        End If
        
        iTotalTaken = con.Execute("Select IsNull(Sum(" & cmbLeaveType.List(cmbLeaveType.ListIndex, 1) & "),0) From Leaves Where DT Between '" & StartDT & "' AND '" & FinalDT & "' AND EmpID='" & cmbEmp.ID & "'").Fields(0).value
        iTotalLimit = Val(cmbLeaveType.List(cmbLeaveType.ListIndex, 2))
        If iTotalLimit <> 0 And OptType(1) Then
            If iTotalTaken + iTotal > iTotalLimit Then
                MsgBox "The Date Range Exceeds The Limit Of " & cmbLeaveType.Text & " Leaves." & vbNewLine & "Can't Add Leaves.", vbInformation
                Exit Function
            End If
        End If
    End If
    
    First = Val(con.Execute("Select COUNT(*) FROM AttendanceSheet " & Cond).Fields(0) & "")
    If First > 0 Then
        'Put Code Here If You Want To Confirm The Deletion Of Attendance
        If MsgBox("The Attendance of this employee is already taken for the selected Dates, Do you want to Delete That?", vbQuestion + vbYesNo) = vbNo Then
            Exit Function
        End If
    End If
    
    Call StartTrans(con)
    con.Execute "DELETE FROM AttendanceSheet " & Cond
    con.Execute "DELETE FROM EmpTimes " & Cond
    con.Execute "DELETE FROM Leaves " & Cond
    
    For i = 0 To iLoop
        DT = DateAdd("d", i, DTPicker1)
        If OptType(0).value Then
            con.Execute "INSERT INTO AttendanceSheet(EmpID,DT,Attendance) VALUES('" & cmbEmp.ID & "','" & DT & "',0)"
        Else
            If Opts(1).value Then
                con.Execute "INSERT INTO AttendanceSheet(EmpID,DT,Attendance) VALUES('" & cmbEmp.ID & "','" & DT & "',3)"
            Else
                con.Execute "INSERT INTO AttendanceSheet(EmpID,DT,Attendance) VALUES('" & cmbEmp.ID & "','" & DT & "',2)"
            End If
            TFrom = Format(DtFrom, "HH:mm")
            TTo = Format(DtTo, "HH:mm")
                  
            Hrs = Round(DateDiff("n", TFrom, TTo) / 60, 2)
              
            '''''''''' WILL INSERT INTO LEAVES '''''''''''
                  
            con.Execute "INSERT INTO Leaves(EmpID,DT,AuthPerson,Reason," & _
             "[From],[To],[Full],Hrs,Deduct" & FieldName & ") VALUES('" & cmbEmp.ID _
             & "','" & DT & "','" & cmbAuth.Text & "','" & txtDesc _
             & "','','','" & Full & "'," & Hrs & "," & Deduct & _
             IIf(FieldName = "", ")", IIf(OptHalfDay, ",0.5", ",1") & ")")
        End If
    Next
    
    con.CommitTrans
        
    Saved = True
    Unload Me
    SQL = ""
    
    Exit Function
    
err:
    con.RollbackTrans
    MsgBox err.Description
    
End Function

Private Sub TakeAtt(DT As Date)
    
    
     
End Sub

''''''''    Will Select The Attendance From AttendanceSheet '''''''''''

Private Function IsPrevLeave(Cond As String) As Boolean
    
    Dim First As Integer
    First = Val(con.Execute("Select COUNT(*) FROM AttendanceSheet " & Cond & "").Fields(0) & "")
    If First <> 0 Then
        If MsgBox("The Attendance For This Employee Is Already Taken,Do You Want To Delete That", vbInformation + vbYesNo) = vbYes Then
            con.Execute "DELETE FROM AttendanceSheet " & Cond
            con.Execute "DELETE FROM EmpTimes " & Cond
            con.Execute "DELETE FROM Leaves " & Cond & " AND Full=1"
            IsPrevLeave = True
        Else
            IsPrevLeave = False
        End If
    End If
    
End Function

'Private Function IsPrevEntry(Cond As String) As Boolean
'On Error Resume Next
'Dim Shift As String
'Shift = con.Execute("Select Shift FROM Leaves " & Cond & "").Fields(0)
'If Shift = "" Then
'    IsPrevEntry = False: Exit Function
'Else
'    If Shift = "Complete Day" Then
'        IsPrevEntry = True: Exit Function
'    End If
'    If Shift = cmbShift Then
'        IsPrevEntry = True: Exit Function
'    End If
'End If
'End Function

Private Function IsPrevTime() As Boolean


Dim Temp As New ADODB.Recordset
Dim TimeFrom  As Date, TimeTo As Date, CTimeFrom As Date, CTimeTo As Date

With Temp
    .CursorLocation = adUseClient
    .Open "SELECT * FROM Leaves WHERE DT='" & DTPicker1 & "' AND EmpID='" & cmbEmp.ID & "'", con, adOpenStatic, adLockReadOnly
    
    CTimeFrom = Format(DtFrom, "HH:mm")
    CTimeTo = Format(DtTo, "HH:mm")
            
    If .RecordCount > 0 Then
        For i = 1 To .RecordCount
            TimeFrom = !From & ""
            TimeTo = !To & ""
            
            If CTimeFrom > TimeFrom And CTimeTo < TimeTo Then
                IsPrevTime = True
                Exit For
            ElseIf CTimeFrom = TimeFrom Then
                IsPrevTime = True
                Exit For
            End If
            .MoveNext
        Next
    
    End If
    
End With

End Function

Private Sub cmbEmp_matched()
    Dim Emptype As Long
    Emptype = con.Execute("Select EmpType From Employees Where EmpID='" & cmbEmp.ID & "'").Fields(0)
    
    cmbAuth.ClearVals
    cmbAuth.ClearVals
    cmbAuth.AddVals con, "PersonName", "LeaveAuthPersons", "EntryID", " Order By PersonName"
'    If Emptype = 2 Then
'        cmbAuth.AddItem "Ahsan Naeem", "0"
'        cmbAuth.AddItem "Mohsin Naeem", "1"
'        cmbAuth.AddItem "Aslam Pervaiz", "2"
''        cmbAuth.AddItem "Asifa Lone", "3"
''        cmbAuth.AddItem "Muhammad Wasiq", "4"
''        cmbAuth.AddItem "Abid Butt", "5"
'    Else
'        cmbAuth.AddItem "Ahsan Naeem", "0"
'        cmbAuth.AddItem "Mohsin Naeem", "1"
'        cmbAuth.AddItem "Aslam Pervaiz", "2"
''        cmbAuth.AddItem "Asifa Lone", "3"
''        cmbAuth.AddItem "Muhammad Wasiq", "4"
''        cmbAuth.AddItem "Abid Butt", "5"
'    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If Saved Then
        Unload Me
        frmLeaves.Show 1
    End If
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub

Private Sub cmdPrintLeaves_Click()

    On Error GoTo err
    Dim cmd As New ADODB.command
    With cmd
        .CommandTimeout = 120
        .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_EmpLeaveBalances"
        .Parameters("@DT").value = DateSerial(DTPicker1.Year, DTPicker1.Month, 1)
        .Execute
    End With
    
    Dim rpt As CRAXDDRT.Report
    If cmbEmp.MatchFound Then
        'Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpLeaveBalanceMonthlyDetail.rpt")
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpLeaveBalanceMonthly.rpt")
    Else
        Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptEmpLeaveBalanceMonthly.rpt")
    End If
    rpt.FormulaFields(16).Text = "'" & cmbMonth & " - " & cmbYear & "'"
    Dim f As New frmPrevRpt
    Dim forSel As String
    
    If cmbEmp.MatchFound = False Then
        forSel = " "
    Else
        forSel = "{VEmpLeaveBalance.EmpID}='" & cmbEmp.ID & "' "
    End If
    
    Dim StartDT As Date
    Dim EndDT As Date
    
    If DTPicker1.Month < 7 Then
        StartDT = DateSerial(DTPicker1.Year - 1, 7, 1)
    Else
        StartDT = DateSerial(DTPicker1.Year, 7, 1)
    End If
    EndDT = DTPicker1
    
    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition
    
    Set FormulaFields = rpt.FormulaFields
    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & company & "'"
        End If
        If cmbEmp.MatchFound Then
            If FormulaField.Name = "{@ForStartDT}" Then
                FormulaField.Text = "#" & StartDT & "#"
            ElseIf FormulaField.Name = "{@ForEndDT}" Then
                FormulaField.Text = "#" & EndDT & "#"
            End If
        End If
    Next
    
    
    f.ShowReport forSel, rpt
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub DTPicker1_Change()
    DTPicker2.value = DTPicker1.value
End Sub

Private Sub Form_Load()

    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    cmbAuth.ListHeight = 600
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals con, "Name", "Departments", "DeptID"
    
    'cmbDept.Text = "<All Departments>"
    'cmbAuth.AddVals con, "Name", "Employees", "EmpID", "Where Active=1"


    DTPicker1 = Date
    DTPicker2 = Date
    DtFrom = Now
    DtTo = Now

    With cmbLeaveType
        .AddItem "Casual Leave"
        .List(0, 1) = "CL"
        .List(0, 2) = "10"
        .AddItem "Sick Leave"
        .List(1, 1) = "SL"
        '.List(1, 2) = "8"
        .List(1, 2) = "16"
        .AddItem "Annual Leave"
        .List(2, 1) = "AL"
        .List(2, 2) = "14"
        .AddItem "Compensatory Leave"
        .List(3, 1) = "CPL"
        .List(3, 2) = "10"
        .AddItem "Unapproved Leave"
        .List(4, 1) = "LWP"
        .List(4, 2) = "0"
        .AddItem "Maternity Leave"
        .List(5, 1) = "ML"
        .List(5, 2) = "90"
        .AddItem "Special Approved Leave"
        .List(6, 1) = "SAL"
        .List(6, 2) = "0"
    End With
        
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyNumber(KeyAscii.value, True)
End Sub

Private Sub txtDeduct_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyNumber(KeyAscii.value, True)
End Sub

Private Sub Label1_Click(Index As Integer)
    If Index = 1 Then
        Dim f As New frmMiscS
        f.ShowForm "LeaveAuthPersons", "PersonName", "Auth Person", "EntryID", "Leave Auth Persons"
        Set f = Nothing
        cmbAuth.ClearVals
        cmbAuth.AddVals con, "PersonName", "LeaveAuthPersons", "EntryID", " Order By PersonName"
    End If
End Sub

Private Sub OptType_Click(Index As Integer)
    Frame1.Enabled = OptType(1)
End Sub
