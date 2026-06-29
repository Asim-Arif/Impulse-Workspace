VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAddMedicalCheckup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Medical Check-Up"
   ClientHeight    =   3375
   ClientLeft      =   7170
   ClientTop       =   5865
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
   ScaleHeight     =   3375
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3285
      Left            =   45
      TabIndex        =   9
      Top             =   0
      Width           =   4845
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   4
         Left            =   75
         TabIndex        =   5
         Top             =   2325
         Width           =   4680
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   75
         TabIndex        =   4
         Top             =   1650
         Width           =   4680
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   1470
         TabIndex        =   3
         Top             =   990
         Width           =   3255
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   2430
         TabIndex        =   1
         Top             =   360
         Width           =   2325
      End
      Begin VB.CommandButton cmdSavenNew 
         Caption         =   "Save && &New"
         Height          =   405
         Left            =   1350
         TabIndex        =   8
         Top             =   2790
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
         Width           =   2325
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Save && Close"
         Height          =   405
         Left            =   2490
         TabIndex        =   6
         Top             =   2790
         Width           =   1110
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   405
         Left            =   3645
         TabIndex        =   7
         Top             =   2790
         Width           =   1110
      End
      Begin MSComCtl2.DTPicker DT 
         Height          =   330
         Left            =   75
         TabIndex        =   2
         Top             =   990
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   90963971
         CurrentDate     =   36891
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   75
         TabIndex        =   15
         Top             =   2115
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Treatment"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   75
         TabIndex        =   14
         Top             =   1440
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   75
         TabIndex        =   13
         Top             =   795
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lab"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   2460
         TabIndex        =   12
         Top             =   165
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   1500
         TabIndex        =   11
         Top             =   810
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Medical Test"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   10
         Top             =   165
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmAddMedicalCheckup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bCancelled As Boolean
Dim strMedicalTest As String, strLab As String, myDT As Date, strStatus As String, strTreatment As String, strRemarks As String

Public Function ShowMe(ByRef p_strMedicalTest As String, ByRef p_strLab As String, ByRef p_DT As Date, ByRef p_strStatus As String, ByRef p_strTreatment As String, ByRef p_strRemarks As String) As Integer

    Me.Show 1
    
    p_strMedicalTest = strMedicalTest
    p_strLab = strLab
    p_DT = myDT
    p_strStatus = strStatus
    p_strTreatment = strTreatment
    p_strRemarks = strRemarks
    
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
    
    strMedicalTest = txts(0)
    strLab = txts(1)
    myDT = DT.Value
    strStatus = txts(2)
    strTreatment = txts(3)
    strRemarks = txts(4)
    
    Unload Me
    
End Sub

Private Function ValidateData() As Boolean

    If txts(0) = "" Then
        MsgBox "Invalid Test.", vbInformation
        Exit Function
    ElseIf txts(1) = "" Then
        MsgBox "Invalid Lab.", vbInformation
        Exit Function
    End If
    ValidateData = True
    
End Function


