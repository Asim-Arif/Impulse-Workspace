VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form PrintExtSal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "External Salary Sheet"
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
      Top             =   480
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
      Begin MSForms.Label Label1 
         Height          =   240
         Left            =   150
         TabIndex        =   5
         Top             =   855
         Width           =   2475
         BackColor       =   11517387
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
      Begin MSForms.Label Label2 
         Height          =   240
         Left            =   150
         TabIndex        =   4
         Top             =   180
         Width           =   2475
         BackColor       =   11517387
         Caption         =   "  Department"
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
      Height          =   375
      Left            =   3000
      Picture         =   "PrintExtSal.frx":0000
      Top             =   1080
      Width           =   465
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3600
      Picture         =   "PrintExtSal.frx":0616
      Top             =   1080
      Width           =   765
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   360
      Left            =   3015
      TabIndex        =   7
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
      TabIndex        =   6
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
Attribute VB_Name = "PrintExtSal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
Call ShowWait(Me)
Dim f As New SalarySheet
With f
    Call .ShowExternalSalSheet(cmbMonth.Text, cmbYear.Text, cmbDept.Text, cr1)
    Me.Hide
End With

Call HideWait
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
    CYear = Format(ServerDate, "yyyy")
    
    For i = 2002 To CYear
        cmbYear.AddItem i
    Next i
    

    
    cmbDept.ClearVals
    cmbDept.AddItem "<All Departments>", "0"
    cmbDept.AddVals Con, "Name", "Departments", "DeptID"
    cmbDept.Text = "<All Departments>"
    
    cmbYear = year(ServerDate)
    cmbMonth = Format(ServerDate, "MMM")
    
End Sub

