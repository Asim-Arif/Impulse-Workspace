VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAddReview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Review"
   ClientHeight    =   4755
   ClientLeft      =   7170
   ClientTop       =   5865
   ClientWidth     =   5070
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
   ScaleHeight     =   4755
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   45
      TabIndex        =   3
      Top             =   0
      Width           =   4950
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   1635
         TabIndex        =   11
         Top             =   1155
         Width           =   3075
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   105
         TabIndex        =   9
         Top             =   1155
         Width           =   1530
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   2325
         Index           =   3
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   1755
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
         Format          =   87752707
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
         Left            =   2550
         TabIndex        =   1
         Top             =   4155
         Width           =   1110
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   405
         Left            =   3675
         TabIndex        =   2
         Top             =   4155
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Review By"
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
         Index           =   4
         Left            =   1665
         TabIndex        =   12
         Top             =   945
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rating"
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
         Index           =   3
         Left            =   165
         TabIndex        =   10
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Review"
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
         Left            =   165
         TabIndex        =   8
         Top             =   1545
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Width           =   420
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
         Left            =   165
         TabIndex        =   5
         Top             =   225
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmAddReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bCancelled As Boolean
Dim DT As Date, strReviewedBy As String, strReview As String, bEdit As Boolean
Dim strRating As String, strType As String

Public Function ShowMe(ByRef p_DT As Date, ByRef p_ReviewedBy As String, ByRef p_Review As String, ByRef p_strType As String, ByRef p_strRating As String, Optional ByRef p_bEdit As Boolean = False) As Boolean
    
'    bEdit = p_bEdit
'
'    DT = p_DT
'    strReviewedBy = p_ReviewedBy
'    strReview = p_Review
    
    
    Me.Show 1
    
    p_DT = DT
    p_ReviewedBy = strReviewedBy
    p_Review = strReview
    p_strRating = strRating
    p_strType = strType
    
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
    
    strType = txts(0)
    strRating = txts(1)
    strReviewedBy = txts(2)
    strReview = txts(3)
    
    Unload Me
    
End Sub

Private Function ValidateData() As Boolean

    If txts(0) = "" Then
        MsgBox "Invalid Type.", vbInformation
        Exit Function
    ElseIf txts(1) = "" Then
        MsgBox "Invalid Rating.", vbInformation
        Exit Function
    ElseIf txts(2) = "" Then
        MsgBox "Invalid Review By.", vbInformation
        Exit Function
    ElseIf txts(3) = "" Then
        MsgBox "Invalid Review.", vbInformation
        Exit Function
    End If
    ValidateData = True
    
End Function



Private Sub Form_Load()
    
    DT = GetServerDate
    
End Sub

