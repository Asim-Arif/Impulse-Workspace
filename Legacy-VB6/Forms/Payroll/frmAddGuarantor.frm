VERSION 5.00
Begin VB.Form frmAddGuarantor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Guarantor"
   ClientHeight    =   3270
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
   ScaleHeight     =   3270
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3240
      Left            =   45
      TabIndex        =   7
      Top             =   0
      Width           =   4845
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   90
         TabIndex        =   2
         Top             =   2220
         Width           =   2340
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   405
         Left            =   3645
         TabIndex        =   6
         Top             =   2715
         Width           =   1110
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Save && Close"
         Height          =   405
         Left            =   2490
         TabIndex        =   5
         Top             =   2715
         Width           =   1110
      End
      Begin VB.CommandButton cmdSavenNew 
         Caption         =   "Save && &New"
         Height          =   405
         Left            =   1350
         TabIndex        =   4
         Top             =   2715
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   960
         Index           =   2
         Left            =   90
         TabIndex        =   1
         Top             =   1020
         Width           =   4695
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   2475
         TabIndex        =   3
         Top             =   2220
         Width           =   2295
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   0
         Top             =   375
         Width           =   4695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   11
         Top             =   2025
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   10
         Top             =   795
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Known By :"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   2475
         TabIndex        =   9
         Top             =   2025
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   8
         Top             =   165
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmAddGuarantor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bCancelled As Boolean
Dim strName As String, strAddress As String, strKnownBy As String, strPhoneNo As String

Public Function ShowMe(ByRef p_strName As String, p_strAddress As String, p_strKnownBy As String, p_strPhoneNo As String) As Integer

    Me.Show 1
    
    p_strName = strName
    
    p_strAddress = strAddress
    p_strKnownBy = strKnownBy
    p_strPhoneNo = strPhoneNo
    
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
    strName = txts(0)
    
    strPhoneNo = txts(3)
    strAddress = txts(2)
    strKnownBy = txts(1)
    Unload Me
    
End Sub

Private Function ValidateData() As Boolean

    If txts(0) = "" Then
        MsgBox "Invalid Name.", vbInformation
        Exit Function
    ElseIf txts(1) = "" Then
        MsgBox "Invalid Known By.", vbInformation
        Exit Function
    ElseIf txts(2) = "" Then
        MsgBox "Invalid Address.", vbInformation
        Exit Function
    ElseIf txts(3) = "" Then
        MsgBox "Invalid Phone.", vbInformation
        Exit Function
    End If
    ValidateData = True
    
End Function
