VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAddNotice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Notice"
   ClientHeight    =   4155
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   45
      TabIndex        =   3
      Top             =   0
      Width           =   4845
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   2325
         Index           =   1
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   1185
         Width           =   4635
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   330
         Left            =   105
         TabIndex        =   4
         Top             =   450
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   41615363
         CurrentDate     =   39930
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   1635
         TabIndex        =   0
         Top             =   450
         Width           =   3075
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Save && Close"
         Height          =   405
         Left            =   2505
         TabIndex        =   1
         Top             =   3585
         Width           =   1110
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   405
         Left            =   3630
         TabIndex        =   2
         Top             =   3585
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notice"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   8
         Top             =   990
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notice Issued By"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   1680
         TabIndex        =   6
         Top             =   255
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   5
         Top             =   225
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmAddNotice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bCancelled As Boolean
Dim DT As Date, strWarnedBy As String, strWarning As String, bEdit As Boolean

Public Function ShowMe(ByRef p_DT As Date, ByRef p_strWarnedBy As String, ByRef p_strWarning As String, Optional ByRef p_bEdit As Boolean = False) As Boolean
    
'    bEdit = p_bEdit
'
'    DT = p_DT
'    strReviewedBy = p_ReviewedBy
'    strReview = p_Review
    
    
    Me.Show 1
    
    p_DT = DT
    p_strWarnedBy = strWarnedBy
    p_strWarning = strWarning
    
    ShowMe = Not bCancelled

End Function

Private Sub cmdCancel_Click()

    bCancelled = True
        
    Unload Me
    
End Sub

Private Sub cmdOK_Click()
    
    If ValidateData = False Then Exit Sub
    
    bCancelled = False
    DT = DTPicker1.Value
    strWarnedBy = txts(0)
    strWarning = txts(1)
    
    Unload Me
    
End Sub

Private Function ValidateData() As Boolean

    If txts(0) = "" Then
        MsgBox "Invalid Warned By.", vbInformation
        Exit Function
    ElseIf txts(1) = "" Then
        MsgBox "Invalid Warning.", vbInformation
        Exit Function
    End If
    ValidateData = True
    
End Function



Private Sub Form_Load()
    
    DT = GetServerDate
    
    
End Sub

