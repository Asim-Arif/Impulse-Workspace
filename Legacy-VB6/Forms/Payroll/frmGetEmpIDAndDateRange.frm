VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmGetEmpIDAndDateRange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Date..."
   ClientHeight    =   2790
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGetEmpIDAndDateRange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   90
      TabIndex        =   9
      Top             =   285
      Width           =   3945
      _ExtentX        =   6959
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
      Height          =   285
      Left            =   90
      TabIndex        =   10
      Top             =   795
      Width           =   3945
      _ExtentX        =   6959
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
   Begin VB.CheckBox chkInvert 
      Caption         =   "Invert"
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   45
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   2235
      Width           =   1275
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1785
      Width           =   1275
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   105
      TabIndex        =   0
      Top             =   1665
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dddd, dd-MMM-yyyy"
      Format          =   116064259
      CurrentDate     =   39372
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   90
      TabIndex        =   1
      Top             =   2280
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dddd, dd-MMM-yyyy"
      Format          =   116064259
      CurrentDate     =   39372
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&From Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   1425
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&To Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   6
      Top             =   2055
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Employee :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   5
      Top             =   570
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Department :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   75
      Width           =   1125
   End
End
Attribute VB_Name = "frmGetEmpIDAndDateRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bInvert As Boolean
Dim bCancelled As Boolean
Dim strEmpID As String
Dim strDeptID As String
Dim DtFrom As Date
Dim DtTo As Date

Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub

Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    cmbEmp.AddItem "<All Employees>", "0"
    If cmbDept.ID = "0" Then
        cmbEmp.AddVals con, "'{' + EmpID + '} ' + Name", "VEmp", "EmpID", " WHERE TempDept=0"
    Else
        cmbEmp.AddVals con, "'{' + EmpID + '} ' + Name", "VEmp", "EmpID", " WHERE DeptID='" & cmbDept.ID & "'"
    End If
    cmbEmp.ID = "0"
    
End Sub

Private Sub Form_Load()
    
    DTPicker1.Value = DateAdd("d", -7, Date)
    DTPicker2.Value = Date
    cmbDept.ListHeight = 2200
    cmbEmp.ListHeight = 1800
    cmbDept.AddItem "< All Departments >", "0"
    cmbDept.AddVals con, "Name + ' {' + DeptID + '}'", "Departments", "DeptID", " WHERE TempDept=0"
    cmbDept.ID = "0"
    bCancelled = True
    bInvert = False
    
End Sub

Private Sub OKButton_Click()
    
    If cmbDept.MatchFound = False Then
        MsgBox "Select Department.", vbInformation
        cmbDept.SetFocus
        Exit Sub
    ElseIf cmbEmp.MatchFound = False Then
        MsgBox "Select Employee.", vbInformation
        cmbEmp.SetFocus
        Exit Sub
    End If
     
    bCancelled = False
    'bInvert = (chkInvert.value = vbChecked)
    strDeptID = cmbDept.ID
    strEmpID = cmbEmp.ID
    DtFrom = DTPicker1.Value
    DtTo = DTPicker2.Value
    
    Unload Me
    
End Sub

Public Function getEmpIDAndDateRange(ByRef myDTFrom As Date, ByRef myDTTo As Date, ByRef m_strDeptID As String, ByRef m_strEmpID As String) As Boolean
     
    Me.Show 1
'    bPrintEarlier = bEarlier
    'm_bInvert = bInvert
    m_strDeptID = strDeptID
    m_strEmpID = strEmpID
    myDTFrom = DtFrom
    myDTTo = DtTo
    getEmpIDAndDateRange = Not bCancelled
    
End Function
