VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form Attendance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attendance"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
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
   ScaleHeight     =   2535
   ScaleWidth      =   4365
   StartUpPosition =   3  'Windows Default
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   330
      Left            =   1920
      TabIndex        =   8
      Top             =   30
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
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
      Left            =   1920
      TabIndex        =   9
      Top             =   375
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   582
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
      Height          =   285
      Left            =   1950
      TabIndex        =   5
      Top             =   765
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   52166659
      CurrentDate     =   37384
   End
   Begin MSForms.ComboBox cmbShift 
      Height          =   285
      Left            =   1950
      TabIndex        =   7
      Top             =   1110
      Width           =   1275
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "2249;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Shift :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   1440
      TabIndex        =   6
      Top             =   1140
      Width           =   435
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   195
      X2              =   4295
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date Of Attendance :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   315
      TabIndex        =   4
      Top             =   795
      Width           =   1560
   End
   Begin MSForms.CommandButton cmdAbsent 
      Height          =   360
      Left            =   3150
      TabIndex        =   3
      Top             =   2130
      Width           =   1125
      Caption         =   "Make Absent"
      Size            =   "1984;635"
      Accelerator     =   65
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdPresent 
      Height          =   360
      Left            =   1965
      TabIndex        =   2
      Top             =   2130
      Width           =   1125
      Caption         =   "Make Present"
      Size            =   "1984;635"
      Accelerator     =   80
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Employee        :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   420
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Department    :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   315
      TabIndex        =   0
      Top             =   90
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   195
      X2              =   4295
      Y1              =   2055
      Y2              =   2055
   End
End
Attribute VB_Name = "Attendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals Con, "Name", "Employees", "EmpID"
    Else
        cmbEmp.AddVals Con, "Name", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "'"
    End If
End Sub



Private Sub cmdAbsent_Click()
    If cmbEmp.MatchFound Then
        Call TakeAtt(False)
    Else
        MsgBox "Please Select Employee From The List.", vbInformation
    End If
End Sub

Private Sub cmdPresent_Click()
    If cmbEmp.MatchFound Then
        Call TakeAtt(True)
    Else
        MsgBox "Please Select Employee From The List.", vbInformation
    End If
End Sub
Private Sub TakeAtt(Result As Boolean)
    On Error GoTo err
    Dim Sql As String
    Dim MsgText As String
    
    If cmbEmp.ID = "0" Then
        If cmbDept.MatchFound Then
            Dim SubSql As String
            Dim rs As New ADODB.Recordset
            
            If cmbDept.ID = "0" Then
                SubSql = "Select EmpID from Employees"
            Else
                SubSql = "Select EmpID from Employees Where deptID='" & cmbDept.ID & "'"
            End If
            
            With rs
                .Open SubSql, Con, adOpenForwardOnly, adLockReadOnly
                Con.BeginTrans
                Do Until .EOF
                    Con.Execute "Insert into AttendanceSheet (EmpID,DT,Shift,Attendance) Values('" & ![EmpID] & "',#" & DTPicker1 & "#,'" & cmbShift & "'," & Result & ")"
                    .MoveNext
                Loop
                Con.CommitTrans
                .Close
                
            End With
        Else
            MsgBox "Invalid Department Name.", vbInformation
            Exit Sub
        End If
    Else
        Sql = "Insert into AttendanceSheet (EmpID,DT,Shift,Attendance) Values('" & cmbEmp.ID & "',#" & DTPicker1 & "#,'" & cmbShift & "'," & Result & ")"
        Con.Execute Sql
    End If
    
    Exit Sub
    
err:
    MsgBox err.Description
End Sub
Private Sub Form_Load()

    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    cmbShift.AddItem "First"
    cmbShift.AddItem "Second"
    cmbShift.ListIndex = 0
    
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals Con, "Name", "Departments", "DeptID"
End Sub

