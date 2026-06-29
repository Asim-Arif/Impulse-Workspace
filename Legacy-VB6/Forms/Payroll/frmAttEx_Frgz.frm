VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmAttEx_Frgz 
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   -45
   ClientWidth     =   6960
   ControlBox      =   0   'False
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8685
   ScaleWidth      =   6960
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkUseServer 
      Caption         =   "Use System Time"
      Height          =   225
      Left            =   45
      TabIndex        =   17
      Top             =   8070
      Width           =   1665
   End
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   315
      Left            =   45
      TabIndex        =   0
      Top             =   390
      Width           =   4860
      _ExtentX        =   8573
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
   Begin MSComCtl2.DTPicker DTEdit 
      Height          =   315
      Left            =   5370
      TabIndex        =   12
      Top             =   4530
      Visible         =   0   'False
      Width           =   1260
      _ExtentX        =   2223
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
      Format          =   152109059
      UpDown          =   -1  'True
      CurrentDate     =   37384
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4905
      Top             =   5070
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAttEx_Frgz.frx":0000
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAttEx_Frgz.frx":0454
            Key             =   "ORD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAttEx_Frgz.frx":08AC
            Key             =   "PARIN"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAttEx_Frgz.frx":0D00
            Key             =   "ITMIN"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAttEx_Frgz.frx":1154
            Key             =   "CAT"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAttEx_Frgz.frx":15A8
            Key             =   "OK"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAttEx_Frgz.frx":1A00
            Key             =   "TYPE"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   4905
      TabIndex        =   1
      Top             =   375
      Width           =   1995
      _ExtentX        =   3519
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
      CustomFormat    =   "ddd, dd-MMM-yyyy"
      Format          =   152109059
      CurrentDate     =   37384
   End
   Begin MSComCtl2.DTPicker InTime 
      Height          =   330
      Left            =   60
      TabIndex        =   2
      Top             =   960
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
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
      CheckBox        =   -1  'True
      CustomFormat    =   "hh:mm  tt"
      Format          =   152109059
      UpDown          =   -1  'True
      CurrentDate     =   37384
   End
   Begin MSComCtl2.DTPicker BrkOut 
      Height          =   330
      Left            =   1575
      TabIndex        =   3
      Top             =   960
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   582
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
      CheckBox        =   -1  'True
      CustomFormat    =   "hh:mm  tt"
      Format          =   152109059
      UpDown          =   -1  'True
      CurrentDate     =   37384
   End
   Begin MSComctlLib.ListView LV 
      Height          =   5550
      Left            =   45
      TabIndex        =   11
      Top             =   2475
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   9790
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ITM"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "In Time"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Out Time"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Net Time"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComCtl2.DTPicker BrkIn 
      Height          =   330
      Left            =   60
      TabIndex        =   4
      Top             =   1545
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
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
      CheckBox        =   -1  'True
      CustomFormat    =   "hh:mm  tt"
      Format          =   152109059
      UpDown          =   -1  'True
      CurrentDate     =   37384
   End
   Begin MSComCtl2.DTPicker ChkOut 
      Height          =   330
      Left            =   1575
      TabIndex        =   5
      Top             =   1545
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   582
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
      CheckBox        =   -1  'True
      CustomFormat    =   "dd-MMM-yyy hh:mm  tt"
      Format          =   152109059
      UpDown          =   -1  'True
      CurrentDate     =   37384
   End
   Begin MSComCtl2.DTPicker OTIn 
      Height          =   330
      Left            =   60
      TabIndex        =   6
      Top             =   2130
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   582
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
      CheckBox        =   -1  'True
      CustomFormat    =   "hh:mm  tt"
      Format          =   152109059
      UpDown          =   -1  'True
      CurrentDate     =   37384
   End
   Begin MSComCtl2.DTPicker OTOut 
      Height          =   330
      Left            =   1575
      TabIndex        =   7
      Top             =   2130
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   582
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
      CheckBox        =   -1  'True
      CustomFormat    =   "hh:mm  tt"
      Format          =   152109059
      UpDown          =   -1  'True
      CurrentDate     =   37384
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1740
      Left            =   4905
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1995
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Over Time In                        Over Time Out"
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
      Index           =   2
      Left            =   45
      TabIndex        =   16
      Top             =   1905
      Width           =   4155
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Break - In                              Check Out"
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
      Left            =   45
      TabIndex        =   15
      Top             =   1305
      Width           =   4155
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Check In Time                      Break - Out"
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
      Left            =   45
      TabIndex        =   14
      Top             =   720
      Width           =   4155
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   435
      Left            =   2925
      TabIndex        =   13
      Top             =   8145
      Visible         =   0   'False
      Width           =   1965
      Caption         =   "Feed All Attendance"
      PicturePosition =   327683
      Size            =   "3466;767"
      Accelerator     =   78
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
      Caption         =   " Employee                                                                                        Date"
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
      Left            =   60
      TabIndex        =   10
      Top             =   150
      Width           =   6840
   End
   Begin MSForms.CommandButton cmdNext 
      Default         =   -1  'True
      Height          =   1740
      Left            =   4215
      TabIndex        =   8
      Top             =   705
      Width           =   675
      Caption         =   "Add"
      PicturePosition =   327683
      Size            =   "1191;3069"
      Accelerator     =   65
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
      Left            =   5280
      TabIndex        =   9
      Top             =   8070
      Width           =   1605
      Caption         =   "Close            "
      PicturePosition =   327683
      Size            =   "2831;635"
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
Attribute VB_Name = "frmAttEx_Frgz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public edit As Boolean
Dim EditOutTime As Boolean
Dim bEmployees As Boolean
Dim bOverTime As Boolean

Private Sub AddHeads()

End Sub

Private Sub TakeRegularAtt()
        'Call MakeAllPresentForJuly
    On Error GoTo err
    
    Dim rs As New ADODB.Recordset, rs1 As New ADODB.Recordset
    Dim i As Integer
    rs1.Open "Select DeptID,LunchInTime,LunchOutTime From GeneralSettings", con, adOpenKeyset, adLockReadOnly
    Dim iRet As Integer
    Dim AttTime As String
    Call StartTrans(con)
    With rs
        '.Open "Select EmpID From EmpTimes Where DT='" & DTPicker1 & "' AND OutTime Is Null", con, adOpenForwardOnly, adLockReadOnly
        '.Open "Select EmpID From Employees Where Active=1 AND DeptID<>'ASTT' AND EmpID Not In(Select EmpID From AttendanceSheet Where DT='" & DTPicker1.value & "') Order By EmpID", con, adOpenForwardOnly, adLockReadOnly
        .Open "Select DeptID,EmpID From Employees Where EmpID In('VBPE-0007','VHRM-0005','VHRM-0006','VPDT-0063','VPDT-0064','VPDT-0065','VPDT-0066','VPDT-0067','VPDT-0068','VQAS-0011','VQTC-0057','VQTC-0058','VQTC-0059','VQTC-0060','VQTC-0061','VQTC-0062','VQTC-0063','VQTC-0064','VQTC-0065','VQTC-0066','VRND-0015','VRND-0016','VRND-0017','VRND-0018','VSIP-0004','VSIP-0005','VSRT-0033','VSRT-0034','VSRT-0035','VSRT-0036','VSRT-0037','VSRT-0038','VSRT-0039','VSRT-0040','VSRT-0041','VSRTA-0001','VSRTA-0002','VSVS-0086','VSVS-0087','VWSH-0060','VWSH-0061','VWSH-0062','VWSH-0063','VWSH-0064','VWSH-0065','VWSH-0066','VWSH-0067','VWSH-0068','VWSH-0069','VWSH-0070','VWSH-0071','VWSH-0072','VWSH-0073','VWSH-0074','VWSH-0075') Order By EmpID", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            rs1.Find "DeptID='" & !DeptID & "'"
            For i = 1 To 29
                DTPicker1.Value = DateSerial(2008, 2, i)
                iRet = Int(Rnd * (Val(9) - (Val(1) - 1)) + (Val(1)))
                AttTime = "08:" & Format(iRet, "00")
                Call TakeAttendance(!EmpId, DTPicker1.Value, Format(AttTime, "HH:NN"))
                
                iRet = Int(Rnd * (Val(9) - (Val(1) - 1)) + (Val(1)))
                'AttTime = "12:" & Format(30 + iRet, "00")
                AttTime = DateAdd("n", iRet, rs1!LunchInTime)
                Call TakeAttendance(!EmpId, DTPicker1.Value, Format(AttTime, "HH:NN"))
                
                iRet = Int(Rnd * (Val(9) - (Val(1) - 1)) + (Val(1)))
                'AttTime = "13:" & Format(25 + iRet, "00")
                AttTime = DateAdd("n", -iRet, rs1!LunchOutTime)
                Call TakeAttendance(!EmpId, DTPicker1.Value, Format(AttTime, "HH:NN"))
                
                iRet = Int(Rnd * (Val(9) - (Val(1) - 1)) + (Val(1)))
                AttTime = "17:" & Format(iRet, "00")
                Call TakeAttendance(!EmpId, DTPicker1.Value, Format(AttTime, "HH:NN"))
            Next
            Debug.Print !EmpId
            .MoveNEXT
        Loop
        .Close
    End With
    con.CommitTrans
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub cmbEmp_GotFocus()
    cmdNext.Default = False
End Sub

Private Sub cmbEmp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmbEmp.MatchFound Then
            Call cmdNext_Click
        End If
    End If
End Sub

Private Sub cmbEmp_LostFocus()
    cmdNext.Default = True
End Sub

Private Sub CommandButton1_Click()
    Call TakeRegularAtt
    MsgBox "OK"
    'Call TakeRamzanBreakAtt
End Sub

Private Sub TakeRamzanBreakAtt()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim iRet As Integer
    Dim AttTime As String
    Dim AttDT As Date
    Dim strTime As String
    Dim lEntryID As Long
    Dim dHrs As Double
    
    Call StartTrans(con)
    With rs
        'It was taken from 14-Sep-07 to 20-Sep-07
        .Open "Select Min(EntryID) as EntryID,EmpID,DT,Min(InTime) as InTime,Min(OutTime) as OutTime From EmpTimes " & _
        "Where DT Between '10/11/2007' AND '10/12/2007' GROUP BY EmpID,DT Having Count(EmpId) = 1", con, adOpenForwardOnly, adLockReadOnly
        '.Open "Select EntryID,EmpID,DT,InTime,OutTime From EmpTimes Where DT Between '9/14/2007' AND '9/20/2007' AND PayableHrs<5 ", con, adOpenForwardOnly, adLockReadOnly
        '.Open "Select EntryID,EmpID,DT,InTime,OutTime From EmpTimes Where month(DT)=10 and year(DT)=2006 And OutTime Is Null AND EmpID In(Select EmpID From VEmp Where TempDept=0) AND InTime<'10:00' ORDER By EmpID,DT", con, adOpenForwardOnly, adLockReadOnly
        '.Open "Select EntryID,EmpID,DT,InTime,OutTime From EmpTimes Where DT In('10/13/2006','10/14/2006','10/18/2006') AND EmpID='VSVS-0071'", con, adOpenForwardOnly, adLockReadOnly
        '.Open "Select EntryID,EmpID,DT,InTime,OutTime From EmpTimes Where DT Between '10/01/2006' AND '10/21/2006' AND EmpTimes.EmpID In(Select EmpID From VEmp Where TempDept=0 AND Designation<>'Security Guard') Group By EmpTimes.EmpID,DT Having Count(EmpTimes.EmpID)=1", con, adOpenForwardOnly, adLockReadOnly
        '.Open "Select EntryID,EmpID,DT,InTime,OutTime From EmpTimes Where DT ='10/20/2006' AND PayableHrs<5 AND EmpID='VMIS-0001'", con, adOpenForwardOnly, adLockReadOnly
        '.Open "Select EmpID,DT From EmpTimes Where DT Between '10/1/2006' AND '10/21/2006' AND EmpID Not In('VSVS-0065') AND EmpID In(Select EmpID From VEmp Where TempDept=0 AND Designation<>'Security Guard') Group By EmpID,DT Having Count(EmpID)=1 Order By EmpID,DT", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            
            lEntryID = Val(!EntryID & "")
            
            iRet = Int(Rnd * (Val(9) - (Val(1) - 1)) + (Val(1)))
            AttTime = "13:" & Format(iRet, "00")
            If !InTime < "08:11" Then
                dHrs = DateDiff("n", "08:00", AttTime) / 60
            Else
                dHrs = DateDiff("n", !InTime, AttTime) / 60
            End If
            
            con.Execute "Update EmpTimes Set OutTime='" & AttTime & "',Hrs=" & dHrs & ",PayableHrs=" & dHrs & " Where EntryID=" & lEntryID
            'con.Execute "Update EmpTimes Set OutTime='" & AttTime & "',Hrs=" & dHrs & ",PayableHrs=" & dHrs & " Where EmpID='" & !EmpId & "' AND DT='" & !DT & "'"
            'Call TakeAttendance(!EmpId, AttDT, Format(AttTime, "HH:NN"))
            
            If Format(!DT, "dddd") = "Friday" Then
                iRet = Int(Rnd * (Val(9) - (Val(1) - 1)) + (Val(1)))
                AttTime = "14:" & Format(iRet, "00")
            Else
                iRet = Int(Rnd * (Val(9) - (Val(1) - 1)) + (Val(1)))
                AttTime = "13:" & Format(25 + iRet, "00")
            End If
            strTime = AttTime
            
            iRet = Int(Rnd * (Val(9) - (Val(1) - 1)) + (Val(1)))
            AttTime = "15:" & Format(30 + iRet, "00")
            dHrs = DateDiff("n", strTime, AttTime) / 60
            con.Execute "Insert Into EmpTimes(EmpID,DT,InTime,OutTime,Hrs,PayableHrs) Values('" & _
             !EmpId & "','" & !DT & "','" & strTime & "','" & AttTime & "'," & dHrs & ",3)"
            
            Debug.Print !EmpId
            .MoveNEXT
        Loop
        .Close
    End With
    con.CommitTrans
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub DTEdit_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim InTime As String, OutTime As String, Hrs As Double
    Dim dLateHours As Double
    Dim EmpId As String, DT As Date
    
    If KeyCode = 13 Then
        If EditOutTime Then
            OutTime = Format(DTEdit, "HH:NN")
            InTime = con.Execute("Select InTime FROM EmpTimes WHERE EntryID=" & LV.SelectedItem.Tag & "").Fields(0)
            Hrs = DateDiff("n", CDate(InTime), CDate(OutTime)) - 60              ' Lunch Time
            con.Execute "UPDATE EmpTimes SET OutTime='" & OutTime & "',Hrs=" & Round(Hrs / 60, 2) & " WHERE EntryID=" & LV.SelectedItem.Tag & ""
        Else
            InTime = Format(DTEdit, "HH:NN")
            OutTime = con.Execute("Select OutTime FROM EmpTimes WHERE EntryID=" & LV.SelectedItem.Tag & "").Fields(0) & ""
            If OutTime <> "" Then Hrs = DateDiff("n", CDate(InTime), CDate(OutTime)) - 60        ' Lunch Time
            con.Execute "UPDATE EmpTimes Set InTime='" & InTime & "',Hrs=" & Round(Hrs / 60, 2) & " WHERE EntryID=" & LV.SelectedItem.Tag & ""
            
            If LV.SelectedItem.Index = 1 Then 'Update Late Hours If It's Ist Entry Of The Day.\
                dLateHours = DateDiff("n", "08:00", InTime) / 60
                If dLateHours > 10 Then
                    con.Execute "Update AttendanceSheet Set LateHours=" & dLateHours & " Where EmpID='" & cmbEmp.ID & "' And DT='" & DTPicker1 & "'"
                Else
                    con.Execute "Update AttendanceSheet Set LateHours=Null Where EmpID='" & cmbEmp.ID & "' And DT='" & DTPicker1 & "'"
                End If
            End If
        End If
        DTEdit.Visible = False
        Call RefreshList
    End If
    
End Sub

Private Sub LV_Click()
    DTEdit.Visible = False
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo err
    
    If LV.ListItems.count = 0 Then Exit Sub
    
    If KeyCode = 46 Then
        If MsgBox("Do you really want to delete ?", vbInformation + vbYesNo) = vbYes Then
            
'            con.Execute "Delete From EmpTimes Where EntryID=" & Val(LV.SelectedItem.Tag)
            
'            If LV.ListItems.Count = 1 Then
'                con.Execute "Delete From AttendanceSheet Where EmpID='" & cmbEmp.ID & "' AND DT='" & DTPicker1 & "'"
'            End If
            con.Execute "Delete From AttendanceSheet Where EmpID='" & cmbEmp.ID & "' AND DT='" & DTPicker1.Value & "'"
            con.Execute "Delete From EmpTimes Where EmpID='" & cmbEmp.ID & "' AND DT='" & DTPicker1.Value & "'"
            
            Call RefreshList
        End If
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub LV_DblClick()
    
    Exit Sub    'Stop The Editing For Now.... Can Still Delete The Att. And Take New.
    If Not EditOutTime Then
        DTEdit.Move LV.Left + LV.ColumnHeaders(2).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(2).Width
        DTEdit = CDate(LV.SelectedItem.SubItems(1))
        DTEdit.Visible = True
        DTEdit.SetFocus
    ElseIf EditOutTime And LV.SelectedItem.SubItems(2) <> "" Then
        DTEdit.Move LV.Left + LV.ColumnHeaders(3).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(3).Width
        DTEdit = CDate(LV.SelectedItem.SubItems(2))
        DTEdit.Visible = True
        DTEdit.SetFocus
    End If
    
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If x > LV.ColumnHeaders(3).Left Then
        EditOutTime = True
    Else
        EditOutTime = False
    End If
    
End Sub

Private Sub TV_Click()
    If Len(TV.SelectedItem.key) <> 3 Then
        If Left(TV.SelectedItem.key, 3) = "PSS" Then
            cmbEmp.ID = Left(TV.SelectedItem.key, Len(TV.SelectedItem.key) - 1)
        End If
    End If
End Sub

Private Sub TV_Collapse(ByVal Node As MSComctlLib.Node)
    'Node.Image = "ITMIN"
End Sub

Private Sub TV_Expand(ByVal Node As MSComctlLib.Node)
    'Node.Image = "ITM"
End Sub

Private Sub SaveNowOld(EmpId As String, AttDate As Date, AttTime As Date)

    On Error GoTo err

    Dim ExtraDeds As Double, ExtraHrs As Double
    ExtraDeds = 0
    ExtraHrs = 0
    Dim MultipleAllowed As Boolean
    Dim bDeductLHrs As Boolean
    Dim bExempt As Boolean
    Dim DeptID As String
    Dim dLateHrs As Double
    Dim TotalEntries As Integer
    Dim AttRecs As Integer
    Dim bOverTime As Boolean
    Dim strEmpInTime As String, strEmpOutTime As String
    Dim strEmpLunchIn As String, strEmpLunchOut As String
    bOverTime = False
    '''''''''''   This Will Is The Employee Exists Or Not    '''''''''''
    
    With con.Execute("Select DeptID,EmpID,EmpType,ExemptSettings FROM Employees WHERE EmpID='" & EmpId & "'")
        If .EOF Then
            Exit Sub
        Else
            EmpId = ![EmpId] & ""
            Emptype = ![Emptype] & ""
            bExempt = ![ExemptSettings] & ""
            DeptID = ![DeptID] & ""
        End If
        .Close
        
        If bExempt Then
            .Open "Select DeductLHrs From EmpSettings Where EmpID='" & EmpId & "'", con, adOpenForwardOnly, adLockReadOnly
        Else
            .Open "Select DeductLHrs From GeneralSettings Where DeptID='" & DeptID & "'", con, adOpenForwardOnly, adLockReadOnly
        End If
        
        If .EOF Then
            bDeductLHrs = True
        Else
            bDeductLHrs = ![DeductLHrs]
'            strEmpInTime = Format(!EnterTime, "HH:NN")
'            strEmpOutTime = Format(!ExitTime, "HH:NN")
'            strEmpLunchIn = Format(!LunchInTime, "HH:NN")
'            strEmpLunchOut = Format(!LunchOutTime, "HH:NN")
        End If
        .Close
    End With
    
    Dim Rec As New ADODB.Recordset, CurrTime As String, Hrs As Double, SQL As String, Cond As String
    Dim TimeDiff As Integer

    Cond = "WHERE EmpID='" & EmpId & "' AND DT='" & AttDate & "'"

    CurrTime = Format(AttTime, "HH:NN")

    TimeDiff = Val(con.Execute("Select TimeDiff FROM TimeSettings").Fields(0) & "")
        
    With Rec
        .Open "Select TotalEntries From VEmpDailyEntries Where EmpID='" & EmpId & "' and DT='" & AttDate & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            TotalEntries = 0
        Else
            TotalEntries = .Fields(0) & ""
        End If
        .Close
        
        .Open "SELECT * FROM EmpTimes WHERE EntryID=(SELECT Max(EntryID) FROM EmpTimes " & Cond & ")", con, adOpenForwardOnly, adLockReadOnly
            
        If TotalEntries > 0 Then
            
            If TotalEntries = 1 And !OutTime & "" = "" Then 'Means It's Lunch In Entry...
                '12:50 to 1:10
                If CurrTime < #12:25:00 PM# Then
                    ExtraDeds = 0 'DateDiff("n", CurrTime, "12:55") / 60
                ElseIf CurrTime > #12:35:00 PM# Then
                    ExtraDeds = 0 'DateDiff("n", "13:00", CurrTime) / 60
                End If
                
                If DateDiff("n", "8:00", CDate(!InTime)) > 0 And DateDiff("n", "8:00", CDate(!InTime)) < 26 Then
                    ExtraHrs = DateDiff("n", "8:00", CDate(!InTime)) / 60
                ElseIf DateDiff("n", "8:00", CDate(!InTime)) < 0 Then
                    ExtraDeds = ExtraDeds + DateDiff("n", CDate(!InTime), "8:00") / 60
                End If
                
                If Format(AttDate, "dddd") = "Friday" Then
                    ExtraHrs = ExtraHrs + 0.25
                    If CurrTime > #12:15:00 PM# Then
                        Hrs = DateDiff("n", CDate(!InTime), #12:15:00 PM#)
                    Else
                        Hrs = DateDiff("n", CDate(!InTime), CurrTime)
                    End If
                Else
                    If CurrTime > #12:24:00 PM# Then
                        Hrs = DateDiff("n", CDate(!InTime), #12:30:00 PM#)
                    Else
                        Hrs = DateDiff("n", CDate(!InTime), CurrTime)
                    End If
                End If
                
            ElseIf TotalEntries = 2 And !OutTime & "" = "" Then 'Means It's Lunch Out Entry
                If Format(AttDate, "dddd") = "Friday" Then
                    ExtraHrs = ExtraHrs + 0.5
                    If CDate(!InTime) > #2:05:00 PM# Then
                        ExtraDeds = 0 'DateDiff("n", "14:30", CDate(!InTime)) / 60
                    ElseIf CDate(!InTime) < #2:00:00 PM# Then
                        ExtraDeds = DateDiff("n", CDate(!InTime), "14:00") / 60
                    End If
                    If DateDiff("n", "14:00", CDate(!InTime)) > 0 And DateDiff("n", "14:00", CDate(!InTime)) < 6 Then
                        ExtraHrs = ExtraHrs + DateDiff("n", "14:00", CDate(!InTime)) / 60
                    End If
                Else
                    If DateDiff("n", "13:30", CDate(!InTime)) > 0 And DateDiff("n", "13:30", CDate(!InTime)) < 6 Then
                        ExtraHrs = DateDiff("n", "13:30", CDate(!InTime)) / 60
                    End If
                    
                    If CDate(!InTime) > #1:35:00 PM# Then
                        ExtraDeds = 0 'DateDiff("n", "14:00", CDate(!InTime)) / 60
                    ElseIf !InTime < #1:30:00 PM# Then
                        ExtraDeds = DateDiff("n", CDate(!InTime), "13:30") / 60
                    End If
                End If
                
                If CurrTime > #5:00:00 PM# Then
                    ExtraDeds = ExtraDeds + 0 'DateDiff("n", "17:00", CurrTime) / 60
                End If
                
                If CurrTime > #5:00:00 PM# Then
                    Hrs = DateDiff("n", CDate(!InTime), #5:00:00 PM#)
                Else
                    Hrs = DateDiff("n", CDate(!InTime), CurrTime)
                End If
            End If
            'Check ... Is This Out Time Entry ?
            
            If TotalEntries < 3 Then
                If !OutTime & "" = "" Then
                    'Hrs = DateDiff("n", CDate(!InTime), CDate(CurrTime))
                
                    'Check for the Minimum Time Difference
                    If Hrs < TimeDiff Then GoTo END_OF_PROCEDURE
                
                    Hrs = Round(Hrs / 60, 2)
                            
                    'Update Out Time and Do Nothing Else
                
                    SQL = "UPDATE EmpTimes SET OutTime='" & CurrTime & "',Hrs=" & Hrs & ",PayableHrs=" & (Hrs + ExtraHrs - ExtraDeds) & " WHERE EntryID=" & Val(!EntryID & "") & ""
                    con.Execute SQL
                    'con.Execute "Update AttendanceSheet Set Attendance=1 Where EmpID='" & EmpId & "' AND DT='" & AttDate & "'"
                    GoTo END_OF_PROCEDURE
                Else
                    'Check for the Minimum Time Difference
                
                    Hrs = DateDiff("n", CDate(!OutTime), CDate(CurrTime))
                    If TotalEntries <> 2 Then
                        If Hrs < TimeDiff Then GoTo END_OF_PROCEDURE
                    End If
                End If
            End If
        Else
            If bDeductLHrs And DateDiff("n", "08:00", CurrTime) > 10 Then
                dLateHrs = DateDiff("n", "08:00", CurrTime) / 60
            Else
                dLateHrs = 0
            End If
            AttRecs = Val(con.Execute("Select Count(*) From AttendanceSheet Where EmpID='" & EmpId & "' AND DT='" & AttDate & "'").Fields(0).Value & "")
            If AttRecs <> 0 Then con.Execute "Delete From AttendanceSheet Where EmpID='" & EmpId & "' AND DT='" & AttDate & "'"
            
            con.Execute "INSERT INTO AttendanceSheet(EmpID,DT,Attendance,LateHours) VALUES('" & EmpId & "','" & AttDate & "',1," & _
             dLateHrs & ")", a
            
        End If

        '.Close
        
        If Emptype = 2 Then
            
            ' If No Previous Entry Found Accept The Entry
            If TotalEntries = 0 Then GoTo ACCEPT_ENTRY
        
            'Remember The Total Entries
        
            ' If this is Second Time Entry Accept It
            If TotalEntries < 2 Then GoTo ACCEPT_ENTRY

        End If
            
        ' If Two Entries Have Been Made ... Check For Over Time Authorization
        If TotalEntries >= 2 Then
            Dim rsOT As New ADODB.Recordset
            Dim bAuthorized As Boolean
            With rsOT
                .Open "Select * From VOverTimeAuth Where EmpID='" & EmpId & "' and DT='" & AttDate & "'", con, adOpenForwardOnly, adLockReadOnly
                ' If Not Authorized for Over Time Neglect This Entry
                If .EOF Then
                    bAuthorized = False
                Else
                    bAuthorized = True
                End If
                rsOT.Close
                Set rsOT = Nothing
                If bAuthorized = False Then
                    MsgBox "Not Authorized.", vbInformation
                    GoTo END_OF_PROCEDURE
                End If
            End With
            'If Val(![Authority] & "") <> 1 Then GoTo END_OF_PROCEDURE
            
            ' Other Wise Accept The Entry
            If !OutTime & "" <> "" Then
                con.Execute "INSERT INTO EmpTimes(EmpID,DT,InTime,ManualEntry,OverTime,UserName,ComputerName) VALUES('" & EmpId _
                 & "','" & AttDate & "','" & CurrTime & "',1,1,'" & CurrentUserName & "','" & strComputerName & "')"
            Else
                Hrs = DateDiff("n", CDate(!InTime), CurrTime)
                If Hrs < TimeDiff Then GoTo END_OF_PROCEDURE
                Hrs = Round(Hrs / 60, 2)
                con.Execute "UPDATE EmpTimes SET OutTime='" & CurrTime & "',Hrs=" & Hrs & ",PayableHrs=" & (Hrs) & " WHERE EntryID=" & Val(!EntryID & "") & ""
            End If
            GoTo END_OF_PROCEDURE
            bOverTime = True
        End If
        
    End With

    ' Enter The New Entry With In Time

ACCEPT_ENTRY:
    
    SQL = "INSERT INTO EmpTimes(EmpID,DT,InTime,ManualEntry,UserName,ComputerName) VALUES('" & EmpId _
     & "','" & AttDate & "','" & CurrTime & "',1,'" & CurrentUserName & "','" & strComputerName & "')"
    
    con.Execute SQL
'    TotalEntries = Val(con.Execute("Select TotalEntries From VEmpDailyEntries Where EmpID='" & EmpID & "' and DT='" & AttDate & "'").Fields(0).Value & "")
'
'    If TotalEntries = 1 Then    'That Means that it's first entry of the day. So Calculate LateHours...
'        If DateDiff("n", "08:00", CurrTime) > 15 Then
'            con.Execute "Update AttendanceSheet Set LateHours=" & DateDiff("n", "08:00", CurrTime) / 60 & " Where EmpID='" & EmpID & "' and DT='" & AttDate & "'"
'        End If
'    End If
    
END_OF_PROCEDURE:
    
    If Rec.State = 1 Then Rec.Close
    Set Rec = Nothing
    
    Exit Sub
err:
    Debug.Print err.Description
    'So That Next Entries Should Not Be Skipped
    Resume Next
    
End Sub


Private Function Saved() As Boolean
    
    If InTime.Value >= BrkOut.Value Or BrkOut.Value >= BrkIn.Value Or BrkIn.Value >= ChkOut.Value Then
        MsgBox "Out Value Could Not Be Greater/Equal In Value  !", vbInformation
        
        Exit Function
    End If
    
'    If IsPrevTime Then
'        MsgBox "You Have Already Taken Attendance For This Time !", vbInformation
'        DTFrom.SetFocus
'        Exit Function
'    End If
    Dim bUseServer As Boolean
    If chkUseServer.Value = vbChecked Then
        bUseServer = True
    Else
        bUseServer = False
    End If
    
    Call StartTrans(con)
    
    If Not (bUseServer And LV.ListItems.count > 0) Then
        con.Execute "DELETE FROM AttendanceSheet WHERE EmpID='" & cmbEmp.ID & "' AND DT='" & DTPicker1.Value & "'"
        con.Execute "DELETE FROM EmpTimes WHERE EmpID='" & cmbEmp.ID & "' AND DT='" & DTPicker1.Value & "'"
    
        If Not IsNull(InTime) Then
            Call TakeAttendance(cmbEmp.ID, DTPicker1, Format(IIf(chkUseServer.Value = vbChecked, Time, InTime), "HH:NN"))
        End If
        con.CommitTrans
    
        Call StartTrans(con)
        If Not IsNull(BrkOut) And BrkOut.Enabled = True Then
            Call TakeAttendance(cmbEmp.ID, DTPicker1, Format(BrkOut, "HH:NN"))
        End If
        con.CommitTrans
    End If
    
    Dim bSecondTimeOnly As Boolean
    bSecondTimeOnly = False
    If IsNull(InTime) And IsNull(BrkOut) Then
        bSecondTimeOnly = True
    End If
    
    Call StartTrans(con)
    If Not IsNull(BrkIn) Then
        Call TakeAttendance(cmbEmp.ID, DTPicker1, Format(BrkIn, "HH:NN"), , bSecondTimeOnly)
    End If
    con.CommitTrans
    
    Call StartTrans(con)
    If Not IsNull(ChkOut) Then
        Call TakeAttendance(cmbEmp.ID, Format(ChkOut, "dd-MMM-yyyy"), Format(IIf(chkUseServer.Value = vbChecked, Time, ChkOut), "HH:NN"), , bSecondTimeOnly)
    End If
    con.CommitTrans
     
    Call StartTrans(con)
    If Not IsNull(OTIn) Then
        Call TakeAttendance(cmbEmp.ID, DTPicker1, Format(OTIn, "HH:NN"))
    End If
    con.CommitTrans
    
     Call StartTrans(con)
    If Not IsNull(OTOut) Then
        Call TakeAttendance(cmbEmp.ID, DTPicker1, Format(OTOut, "HH:NN"))
    End If
    con.CommitTrans
    Call RefreshList
    '------------------------------------------------------------------------------------------------------------
    cmbEmp.ID = ""
    cmbEmp.SetFocus
    
End Function

Private Function SavedOld() As Boolean

    Dim Affect As Long, First As Integer, Second As Integer, IsFirst As Boolean, IsSecond As Boolean
    Dim Deduction As Integer, Hrs As Double, TFrom As String, TTo As String
    
    Dim rs As New ADODB.Recordset
    Dim bDeductLHrs As Boolean
    Dim bExempt As Boolean
    
    With rs
        .Open "Select ExemptSettings From Employees Where EmpID='" & cmbEmp.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        bExempt = !ExemptSettings & ""
        .Close
        If bExempt Then
            .Open "Select DeductLHrs From EmpSettings Where EmpID='" & cmbEmp.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        Else
            .Open "Select DeductLHrs From GeneralSettings Where DeptID='" & DeptID & "'", con, adOpenForwardOnly, adLockReadOnly
        End If
        If .EOF Then
            bDeductLHrs = True
        Else
            bDeductLHrs = ![DeductLHrs]
        End If
        .Close
    End With
    
    dLateHours = 0
    
    SavedOld = False

    'On Error GoTo err
    Dim SQL As String, Cond As String
    
    If Not cmbEmp.MatchFound Then
        MsgBox "Please Select Valid Employee From The List.", vbInformation
        Exit Function
    End If
    If Not IsNull(DtFrom) And Not IsNull(DtTo) Then
        If IsPrevTime Then
            MsgBox "You Have Already Taken Attendance For This Time !", vbInformation
            DtFrom.SetFocus
            Exit Function
        ElseIf DtFrom.Value >= DtTo.Value Then
            MsgBox "Out Value Could Not Be Greater/Equal In Value  !", vbInformation
            DtFrom.SetFocus
            Exit Function
        
        End If
    
        Cond = "WHERE EmpID='" & cmbEmp.ID & "' AND DT='" & DTPicker1 & "'"
    
        con.BeginTrans
        
        TFrom = Format(DtFrom, "HH:mm")
        TTo = Format(DtTo, "HH:mm")
        
        Hrs = Round(DateDiff("n", TFrom, TTo) / 60, 2)
        
        ''''''''''Will INSERT INTO Leaves'''''''''''
        
        con.Execute "INSERT INTO EmpTimes(EmpID,DT,InTime,OutTime,Hrs,UserName,ComputerName) VALUES('" & cmbEmp.ID _
            & "','" & DTPicker1 & "','" & Format(DtFrom, "HH:mm") & "','" & Format(DtTo, "HH:mm") & "'," & Hrs & _
            ",'" & CurrentUserName & "','" & strComputerName & "')"
            
        If LV.ListItems.count = 0 Then
            On Error Resume Next
            If bDeductLHrs And DateDiff("n", "08:00", Format(DtFrom, "HH:mm")) > 15 Then
                dLateHrs = DateDiff("n", "08:00", Format(DtFrom, "HH:mm")) / 60
            Else
                dLateHrs = 0
            End If
            con.Execute "INSERT INTO AttendanceSheet(EmpID,DT,Attendance,LateHours) VALUES('" & cmbEmp.ID & "','" & DTPicker1 & "',1," & dLateHrs & ")"
            If err.Number = -2147217900 Then
                If MsgBox("This Employee is Absent on This Date. Update This Entry ?", vbQuestion + vbYesNo) = vbYes Then
                    con.Execute "Delete From AttendanceSheet Where EmpID='" & cmbEmp.ID & "' and DT='" & DTPicker1 & "'"
                    con.Execute "INSERT INTO AttendanceSheet(EmpID,DT,Attendance) VALUES('" & cmbEmp.ID & "','" & DTPicker1 & "',1)"
                End If

            End If
        End If
        On Error GoTo err
        
        con.CommitTrans
    Else
    
        If Not IsNull(DtFrom) Then
            Call TakeAtt(cmbEmp.ID, Format(DtFrom, "HH:NN"))
        ElseIf Not IsNull(DtTo) Then
            Call TakeAtt(cmbEmp.ID, Format(DtTo, "HH:NN"))
        End If
    
    End If
    
    
    Call RefreshList
    SavedOld = True
    SQL = ""
    Exit Function
err:
    con.RollbackTrans
    MsgBox err.Description
End Function

''''''''Will Select The Attendance From AttendanceSheet'''''''''''
Private Sub CheckLeave(First As Integer, Second As Integer, Cond As String)

    First = Val(con.Execute("Select COUNT(*) FROM AttendanceSheet " & Cond & " AND Shift='First'").Fields(0) & "")
    Second = Val(con.Execute("Select COUNT(*) FROM AttendanceSheet " & Cond & " AND Shift='Second'").Fields(0) & "")

End Sub

Private Function IsPrevEntry(Cond As String) As Boolean
On Error Resume Next
Dim Shift As String
Shift = con.Execute("Select Shift FROM Leaves " & Cond & "").Fields(0)
If Shift = "" Then
    IsPrevEntry = False: Exit Function
Else
    If Shift = "Complete Day" Then
        IsPrevEntry = True: Exit Function
    End If
    If Shift = cmbShift Then
        IsPrevEntry = True: Exit Function
    End If
End If
End Function

Private Sub RefreshList()
    
    On Error GoTo err
    If cmbEmp.MatchFound = False Then Exit Sub
    Dim Rec As New ADODB.Recordset, ITM As ListItem
    Dim iRecCount As Integer
    
    InTime.Value = Null
    BrkOut.Value = Null
    BrkIn.Value = Null
    ChkOut.Value = Null
    OTIn.Value = Null
    OTOut.Value = Null
    LV.ListItems.Clear
    With Rec
        .Open "SELECT * FROM EmpTimes WHERE DT='" & DTPicker1 & "' AND EmpID='" & cmbEmp.ID & "' ORDER BY EntryID", con, adOpenStatic, adLockReadOnly
        iRecCount = .RecordCount
        For i = 1 To .RecordCount
            Set ITM = LV.ListItems.add(, , "")
            ITM.Tag = !EntryID & ""
            If i = 1 Then
                InTime.Value = Format(!InTime & "", "HH:NN AMPM")
            ElseIf i = 2 Then
                BrkIn.Value = Format(!InTime & "", "HH:NN AMPM")
            ElseIf i = 3 Then
                OTIn.Value = Format(!InTime & "", "HH:NN AMPM")
            End If
            ITM.ListSubItems.add , , Format(!InTime & "", "HH:NN AMPM")
            If i = 1 Then
                BrkOut.Value = Format(!OutTime & "", "HH:NN AMPM")
            ElseIf i = 2 Then
                ChkOut.Value = DTPicker1.Value & " " & Format(!OutTime & "", "HH:NN AMPM")
            ElseIf i = 3 Then
                OTOut.Value = Format(!OutTime & "", "HH:NN AMPM")
            End If
            ITM.ListSubItems.add , , Format(!OutTime & "", "HH:NN AMPM")
            ITM.ListSubItems.add , , getHrsMin(Val(!Hrs & ""))
            .MoveNEXT
        Next
        .Close
    End With
    Dim strBrkOutTime As String, strBrkInTime As String
    Dim strInTime As String, strOutTime As String
    
    Call getLunchTimingsByEmpID(cmbEmp.ID, strBrkOutTime, strBrkInTime)
    Call getInOutTimingsByEmpID(cmbEmp.ID, strInTime, strOutTime)
    
    If Format(DTPicker1.Value, "ddd") = "Fri" Then
        Dim rs As New ADODB.Recordset
        With rs
            .Open "SELECT FridayBreakStartTime,FridayBreakEndTime,FridayCheckOut FROM VEmpSettings WHERE EmpID='" & cmbEmp.ID & "'", con, adOpenForwardOnly, adLockReadOnly
            If .EOF = False Then
                strBrkOutTime = Format(!FridayBreakStartTime & "", "HH:NN AMPM")
                strBrkInTime = Format(!FridayBreakEndTime & "", "HH:NN AMPM")
                strOutTime = Format(!FridayCheckOut & "", "HH:NN AMPM")
            End If
            .Close
        End With
    End If
    
    If InTime.Value & "" = "" Then InTime.Value = strInTime
'    If Format(DTPicker1.Value, "ddd") = "Fri" Then
'        If BrkOut.Value & "" = "" Then BrkOut.Value = "13:00"
'        If BrkIn.Value & "" = "" Then BrkIn.Value = "14:00"
'    Else
'        If BrkOut.Value & "" = "" Then BrkOut.Value = Format(strBrkOutTime, "HH:NN AMPM")
'        If BrkIn.Value & "" = "" Then BrkIn.Value = Format(strBrkInTime, "HH:NN AMPM")
'    End If
    If strBrkInTime = strBrkOutTime Then
        BrkOut.Enabled = False
        BrkIn.Enabled = False
    Else
        BrkIn.Enabled = True
        BrkOut.Enabled = True
        If BrkOut.Value & "" = "" Then BrkOut.Value = Format(strBrkOutTime, "HH:NN AMPM")
        If BrkIn.Value & "" = "" Then BrkIn.Value = Format(strBrkInTime, "HH:NN AMPM")
    End If
        
        
    If ChkOut.Value & "" = "" Then
        ChkOut.Value = DTPicker1.Value & " " & strOutTime
    End If
    
    If OTIn.Value & "" = "" Then
        OTIn.Value = "16:31"
        OTIn.Value = Null
    End If
    
    If OTOut.Value & "" = "" Then
        OTOut.Value = "22:01"
        OTOut.Value = Null
    End If
    Set rs = Nothing
    
    If iRecCount = 0 Then
        'BrkIn.Value = Null
        'If strBrkInTime <> strBrkOutTime Then
        'ChkOut.Value = Null
        'End If
'    ElseIf iRecCount = 1 Then
'        InTime.Value = Null
'        BrkOut.Value = Null
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Function IsPrevTime() As Boolean


Dim Temp As New ADODB.Recordset
Dim TimeFrom  As Date, TimeTo As Date, CTimeFrom As Date, CTimeTo As Date

With Temp
    .CursorLocation = adUseClient
    .Open "SELECT * FROM EmpTimes WHERE DT='" & DTPicker1 & "' AND EmpID='" & cmbEmp.ID & "' AND ISNULL(OutTime,'')<>''", con, adOpenStatic, adLockReadOnly
    
    If Not IsNull(DtFrom) Then CTimeFrom = Format(DtFrom, "HH:mm")
    If Not IsNull(DtTo) Then CTimeTo = Format(DtTo, "HH:mm")
            
    If .RecordCount > 0 Then
        For i = 1 To .RecordCount
            TimeFrom = !InTime & ""
            TimeTo = !OutTime & ""
            
            If CTimeFrom >= TimeFrom And CTimeFrom <= TimeTo Then
                IsPrevTime = True
                Exit For
            ElseIf CTimeTo >= TimeFrom And CTimeTo <= TimeTo Then
                IsPrevTime = True
                Exit For
            ElseIf CTimeFrom = TimeFrom Then
                IsPrevTime = True
                Exit For
            End If
            .MoveNEXT
        Next
    
    End If
    
End With

End Function

Private Sub cmbEmp_matched()

    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT EmpPic FROM Employees WHERE EmpID='" & cmbEmp.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        If Not IsNull(![EmpPic]) Then
            Set Image1.DataSource = rs
            Image1.DataField = "EmpPic"
        Else
            Image1.Picture = LoadPicture("")
        End If
        .Close
    End With
    Set rs = Nothing
    
    Call RefreshList
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
        
'    If DTPicker1.Value > Date Then
'        MsgBox "Attendance in Future dates not Allowed.", vbInformation
'        Exit Sub
'    End If
    
    If bOverTime Then
        If LV.ListItems.count <> 2 Then
            MsgBox "You can Only Add OverTime.", vbInformation
            Exit Sub
        End If
    End If
    
    If Format(DTPicker1, "ddd") = "Sun" And Not (strCompany = "Weldon Instruments" Or strCompany = "QEL" Or strCompany = "Dr-Frgz" Or strCompany = "IAA" Or strCompany = "Kami" Or strCompany = "Towne") Then
        MsgBox "Can't Add Attendance For Sunday.", vbInformation
        Exit Sub
    End If
    
    If strCompany <> "Dr-Frgz" Then
        If con.Execute("Select DT From Holidays Where DT='" & DTPicker1 & "'").RecordCount > 0 Then
            MsgBox "Can't Add Attendance For Holidays.", vbInformation
            Exit Sub
        End If
    End If
    
    If con.Execute("SELECT DT FROM Leaves WHERE EmpID='" & cmbEmp.ID & "' AND DT='" & DTPicker1 & "'").RecordCount > 0 Then
        MsgBox "Can't Add Attendance, Leave Feeded.", vbInformation
        Exit Sub
    End If
    
    If chkUseServer.Value = vbChecked Then
        Dim strInTime As String
        strInTime = GetSingleStringValue("TOP 1 InTime", "EmpTimes", " WHERE EmpID='" & cmbEmp.ID & "' AND DT='" & DTPicker1.Value & "'")
        If strInTime <> "" Then
            If DateDiff("N", strInTime, Time) <= 45 Then
                MsgBox "Employee Already In.", vbInformation
                Exit Sub
            End If
        End If
    End If
    
    If Saved Then
         'Unload Me
        'IncreaseSal.Show 1
    End If
    
    
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub

Public Sub ShowMe(Optional m_bEmployees As Boolean = True, Optional m_bOverTime As Boolean)
    
    bOverTime = m_bOverTime
    bEmployees = m_bEmployees
    If bEmployees Then
        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' ' + Rel + ' ' + FName", "Employees", "EmpID", "WHERE Active=1 AND DeptID In(Select DeptID From Departments Where TempDept=0)"
    Else
        cmbEmp.AddVals con, "'{' + EmpID + '} ' +  Name + ' ' + Rel + ' ' + FName", "Employees", "EmpID", "WHERE Active=1 AND DeptID In(Select DeptID From Departments Where TempDept=1)"
    End If
    
    Call AddHeads
    Me.Show
    
End Sub


Private Sub DTPicker1_Change()
    Call RefreshList
    If IsNull(ChkOut) = False Then
        ChkOut = DTPicker1.Value & " " & Format(ChkOut.Value, "hh:mm")
    End If
End Sub

Private Sub Form_Load()
    
    bOverTime = False
    bEmployees = False
    'cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 7485
    
    cmbEmp.ClearVals
    'cmbEmp.AddItem "<All Employees>", "0"
     
    DTPicker1 = Date
    DtFrom = Now
    DtTo = Now
    ChkOut = Date
    
End Sub

Private Sub txtAmt_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyNumber(KeyAscii.Value, True)
End Sub

Private Sub txtDeduct_KeyPress(KeyAscii As MSForms.ReturnInteger)
    KeyAscii = OnlyNumber(KeyAscii.Value, True)
End Sub

Private Sub TakeAtt(EmpId As String, AttTime As String)

    '''''''''''   This Will Is The Employee Exists Or Not    '''''''''''

    Dim Rec As New ADODB.Recordset, CurrTime As String, Hrs As Double, SQL As String, Cond As String
    Dim TimeDiff As Integer

    Cond = "WHERE EmpID='" & EmpId & "' AND DT='" & DTPicker1 & "'"
    CurrTime = AttTime 'Format(Now, "HH:NN")

    TimeDiff = Val(con.Execute("Select TimeDiff FROM TimeSettings").Fields(0) & "")

    With Rec
        .Open "SELECT * FROM EmpTimes WHERE EntryID=(SELECT MAX(EntryID) FROM EmpTimes " & Cond & ")", con, adOpenStatic, adLockReadOnly
    
        SQL = "INSERT INTO EmpTimes(EmpID,DT,InTime) VALUES('" & EmpId _
         & "','" & DTPicker1 & "','" & CurrTime & "')"
        If .RecordCount > 0 Then
            If !OutTime & "" = "" Then
                Hrs = DateDiff("n", CDate(!InTime), CDate(CurrTime))
                If Hrs < TimeDiff Then Exit Sub
                Hrs = Round(Hrs / 60, 2)
                SQL = "UPDATE EmpTimes SET OutTime='" & CurrTime & "',Hrs=" & Hrs & " WHERE EntryID=" & Val(!EntryID & "") & ""
            Else
                Hrs = DateDiff("n", CDate(!OutTime), CDate(CurrTime))
                If Hrs < TimeDiff Then Exit Sub
            End If
        Else
            On Error Resume Next
            con.Execute "INSERT INTO AttendanceSheet(EmpID,DT,Attendance) VALUES('" & EmpId & "','" & DTPicker1 & "',1)"
        
            If err.Number = -2147217900 Then
                If MsgBox("This Employee is Absent on This Date. Update This Entry ?", vbQuestion + vbYesNo) = vbYes Then
                    con.Execute "Delete From AttendanceSheet Where EmpID='" & cmbEmp.ID & "' and DT='" & DTPicker1 & "'"
                    con.Execute "INSERT INTO AttendanceSheet(EmpID,DT,Attendance) VALUES('" & cmbEmp.ID & "','" & DTPicker1 & "',1)"
                End If

            End If
        
        End If

        con.Execute SQL
    End With

End Sub





''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' This procedure is made to take attendance of all employees of vision
' technologies for the month of july 2004 automatically at default time
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



Private Sub MakeAllPresentForJuly()
    
    Dim DT As Date
    DT = #7/1/2004#
    
    Dim rs As New ADODB.Recordset
    
    rs.Open "Select EmpID From Employees Order By EmpID", con, adOpenForwardOnly, adLockReadOnly
    
    For i = 1 To 31
                
        If Format(DT, "ddd") = "Sun" Then GoTo Next_day
        
        
        
        With rs
            .MoveFirst
            For j = 0 To rs.RecordCount - 1
                con.Execute "INSERT INTO EmpTimes(EmpID,DT,InTime,OutTime,Hrs) VALUES('" & ![EmpId] _
                & "','" & DT & "','08:00 AM','05:00 PM',8)"
                
                con.Execute "INSERT INTO AttendanceSheet(EmpID,DT,Attendance) VALUES('" & ![EmpId] & "','" & DT & "',1)"
                
                .MoveNEXT
            Next
    
        End With
        
Next_day:
    DT = DT + 1
    
    Next
    
    
    
    
End Sub

