VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form NewPayEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Payment Entry..."
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
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
   ScaleHeight     =   1950
   ScaleWidth      =   4665
   StartUpPosition =   1  'CenterOwner
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   330
      Left            =   1605
      TabIndex        =   0
      Top             =   30
      Width           =   3015
      _ExtentX        =   5318
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
      ListForeColor   =   12582912
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextForeColor   =   6244673
      ListIndex       =   -1
   End
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   330
      Left            =   1605
      TabIndex        =   1
      Top             =   375
      Width           =   3015
      _ExtentX        =   5318
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
      ListForeColor   =   12582912
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextForeColor   =   6244673
      ListIndex       =   -1
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1635
      TabIndex        =   2
      Top             =   765
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   47579139
      CurrentDate     =   37384
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Amount :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   3030
      TabIndex        =   10
      Top             =   810
      Width           =   660
   End
   Begin MSForms.TextBox txtHours 
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Top             =   765
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
      Left            =   1575
      TabIndex        =   5
      Top             =   1530
      Width           =   1575
      Caption         =   "Save & Close"
      PicturePosition =   327683
      Size            =   "2778;635"
      Picture         =   "NewPayEntry.frx":0000
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
      Left            =   45
      TabIndex        =   4
      Top             =   1530
      Width           =   1485
      Caption         =   "Save & New"
      PicturePosition =   327683
      Size            =   "2619;635"
      Picture         =   "NewPayEntry.frx":044E
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
      Left            =   3180
      TabIndex        =   6
      Top             =   1530
      Width           =   1395
      Caption         =   "Cancel        "
      PicturePosition =   327683
      Size            =   "2461;635"
      Picture         =   "NewPayEntry.frx":0560
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
      X1              =   75
      X2              =   4525
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Payment Date :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   465
      TabIndex        =   9
      Top             =   795
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Employee :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   315
      TabIndex        =   8
      Top             =   420
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Department :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   150
      TabIndex        =   7
      Top             =   90
      Width           =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   75
      X2              =   4525
      Y1              =   1455
      Y2              =   1455
   End
End
Attribute VB_Name = "NewPayEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    'cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.Text = "<All Departments>" Then
        cmbEmp.AddVals Con, "Name", "Contractors", "EmpID", "WHERE Vender=0"
    Else
        cmbEmp.AddVals Con, "Name", "Contractors", "EmpID", "Where DeptID='" & cmbDept.ID & "' AND Vender=0"
    End If
    
End Sub
Private Function Saved() As Boolean

    On Error GoTo ERR
    Dim Sql As String
    If Not cmbEmp.MatchFound Then
        MsgBox "Please Select Valid Employee From The List.", vbInformation
        Exit Function
    End If
    Sql = "Insert Into ContPayment (EmpID,DT,Amt) Values('" & cmbEmp.ID & "','" & DTPicker1 & "'," & txtHours & ")"
    Con.Execute Sql
    Saved = True
    
    Sql = ""
    Exit Function
ERR:
    MsgBox ERR.Description
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If Saved Then
        Unload Me
        NewOverTime.Show 1
    End If
    
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub

Private Sub Form_Load()
    DTPicker1 = ServerDate
    cmbDept.ListHeight = 1550
    cmbEmp.ListHeight = 1250
    
    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals Con, "Name", "ContDepartments", "DeptID"
    cmbDept.Text = "<All Departments>"
End Sub

