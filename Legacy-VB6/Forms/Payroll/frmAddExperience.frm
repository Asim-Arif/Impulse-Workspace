VERSION 5.00
Begin VB.Form frmAddExperience 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Experience"
   ClientHeight    =   3285
   ClientLeft      =   7170
   ClientTop       =   5820
   ClientWidth     =   4965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   45
      TabIndex        =   8
      Top             =   0
      Width           =   4845
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   4
         Left            =   75
         TabIndex        =   3
         Top             =   1575
         Width           =   4695
      End
      Begin VB.TextBox txts 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Index           =   3
         Left            =   75
         TabIndex        =   4
         Top             =   2160
         Width           =   2325
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   75
         TabIndex        =   1
         Top             =   1005
         Width           =   2325
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   2430
         TabIndex        =   2
         Top             =   1005
         Width           =   2325
      End
      Begin VB.CommandButton cmdSavenNew 
         Caption         =   "Save && &New"
         Height          =   405
         Left            =   1350
         TabIndex        =   5
         Top             =   2715
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   75
         TabIndex        =   0
         Top             =   375
         Width           =   4695
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Save && Close"
         Height          =   405
         Left            =   2490
         TabIndex        =   6
         Top             =   2715
         Width           =   1110
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   405
         Left            =   3645
         TabIndex        =   7
         Top             =   2715
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Why Left"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   105
         TabIndex        =   13
         Top             =   1365
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salary :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   75
         TabIndex        =   12
         Top             =   1965
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Designation :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   2460
         TabIndex        =   11
         Top             =   810
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Experience (Years)"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   10
         Top             =   810
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Institute"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   165
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmAddExperience"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bCancelled As Boolean
Dim strInstitute As String, strDesignation As String, strPeriod As String, iSalary As Integer, strWhyLeft As String

Public Function ShowMe(ByRef p_strInstitute As String, ByRef p_strDesignation As String, ByRef p_strPeriod As String, ByRef p_iSalary As Integer, ByRef p_strWhyLeft As String) As Integer

    Me.Show 1
    
    p_strInstitute = strInstitute
    
    p_strDesignation = strDesignation
    p_strPeriod = strPeriod
    p_iSalary = iSalary
    p_strWhyLeft = strWhyLeft
    
    ShowMe = Not bCancelled
    
End Function


Private Sub cmdCancel_Click()
    bCancelled = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    If Not ValidateData Then
        Exit Sub
    End If
    bCancelled = False
    strInstitute = txts(0)
    strDesignation = txts(1)
    strPeriod = Val(txts(2))
    iSalary = Val(txts(3))
    strWhyLeft = txts(4)
    Unload Me
    
End Sub

Private Function ValidateData() As Boolean

    If txts(0) = "" Then
        MsgBox "Invalid Institute.", vbInformation
        Exit Function
    ElseIf txts(1) = "" Then
        MsgBox "Invalid Designation.", vbInformation
        Exit Function
    ElseIf Val(txts(2)) <= 0 Then
        MsgBox "Invalid Experience.", vbInformation
        Exit Function
    ElseIf Val(txts(3)) <= 0 Then
        MsgBox "Invalid Salary.", vbInformation
        Exit Function
    ElseIf txts(4) = "" Then
        MsgBox "Please Add Why Left.", vbInformation
        Exit Function
    End If
    ValidateData = True
    
End Function


