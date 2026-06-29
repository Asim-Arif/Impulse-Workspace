VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form NewOverTime 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Over Time..."
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
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
   ScaleWidth      =   4440
   StartUpPosition =   1  'CenterOwner
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   1755
      TabIndex        =   0
      Top             =   30
      Width           =   2625
      _ExtentX        =   4630
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
      Width           =   2625
      _ExtentX        =   4630
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   1785
      TabIndex        =   2
      Top             =   765
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   48168963
      CurrentDate     =   37384
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   960
      Picture         =   "NewOverTime.frx":0000
      Top             =   1560
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   240
      Picture         =   "NewOverTime.frx":0548
      Top             =   1560
      Width           =   465
   End
   Begin VB.Image Image3 
      Height          =   750
      Left            =   3120
      Picture         =   "NewOverTime.frx":0B5E
      Top             =   1200
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Over Time Hours       :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   150
      TabIndex        =   10
      Top             =   1155
      Width           =   1575
   End
   Begin MSForms.TextBox txtHours 
      Height          =   285
      Left            =   1785
      TabIndex        =   3
      Top             =   1125
      Width           =   855
      VariousPropertyBits=   746604571
      Size            =   "1508;503"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   360
      Left            =   1545
      TabIndex        =   5
      Top             =   2130
      Width           =   1395
      Caption         =   "Save & Close"
      PicturePosition =   327683
      Size            =   "2461;635"
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
      Left            =   105
      TabIndex        =   4
      Top             =   2130
      Width           =   1395
      Caption         =   "Save & New"
      PicturePosition =   327683
      Size            =   "2461;635"
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
      Left            =   2970
      TabIndex        =   6
      Top             =   2130
      Width           =   1395
      Caption         =   "Cancel        "
      PicturePosition =   327683
      Size            =   "2461;635"
      Picture         =   "NewOverTime.frx":18DD
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   100
      X2              =   4400
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Over Time Date         :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   9
      Top             =   795
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Employee        :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   90
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   100
      X2              =   4400
      Y1              =   2055
      Y2              =   2055
   End
End
Attribute VB_Name = "NewOverTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    'cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals Con, "Name", "Employees", "EmpID", "Where active=1 AND Vender=0"
    Else
        cmbEmp.AddVals Con, "Name", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' and Active=1 AND Vender=0"
    End If
    
    
End Sub
Private Function Saved() As Boolean

    On Error GoTo ERR
    Dim Sql As String
    If Not cmbEmp.MatchFound Then
        MsgBox "Please Select Valid Employee From The List.", vbInformation
        Exit Function
    ElseIf Val(txtHours) > Val(txtHours.Tag) Then
        MsgBox "Invalid Over Time, You Can Enter Only " & Val(txtHours.Tag) & ".", vbInformation
        Exit Function
    ElseIf Val(txtHours) <= 0 Then
        MsgBox "Invalid Over Time, You Can Enter Only " & Val(txtHours.Tag) & ".", vbInformation
        Exit Function
    End If
    
    Sql = "Insert Into OverTime (EmpID,DT,Hours) Values('" & cmbEmp.ID & "','" & DTPicker1 & "'," & txtHours & ")"
    Con.Execute Sql
    Saved = True
    
    Sql = ""
    Exit Function
ERR:
If ERR.Number = -2147217900 Then
    MsgBox "Over Time For " & cmbEmp.Text & " for " & Format(DTPicker1, "dddd dd-MMM-yyyy") & " is already Entered.", vbInformation
Else
    MsgBox ERR.Description
End If
End Function

Private Sub cmbEmp_matched()
With Con.Execute("Select OverTime From VEmp where Empid='" & cmbEmp.ID & "'")
    If Not .EOF Then
     txtHours.Tag = .Fields(0)
     
    End If
    
End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If Saved Then
        'Unload Me
        cmbEmp.Text = ""
        txtHours = ""
        cmbEmp.SetFocus
    End If
End Sub


Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub

Private Sub Form_Load()
    DTPicker1 = ServerDate
    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals Con, "Name", "Departments", "DeptID"
End Sub

Private Sub txtHours_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = OnlyNumber(KeyAscii.value, True)
End Sub
