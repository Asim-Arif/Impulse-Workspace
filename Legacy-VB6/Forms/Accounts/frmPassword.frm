VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Security Password"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   ControlBox      =   0   'False
   LinkTopic       =   "Form23"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   15
      TabIndex        =   0
      Top             =   -60
      Width           =   4020
      Begin VB.TextBox Pass 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1185
         PasswordChar    =   "l"
         TabIndex        =   1
         Top             =   750
         Width           =   2685
      End
      Begin VB.Label lblLabels 
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   165
         TabIndex        =   3
         Top             =   795
         Width           =   1020
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   150
         Picture         =   "frmPassword.frx":0000
         Top             =   180
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "This option is password protected, It requires password."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   390
         Left            =   810
         TabIndex        =   2
         Top             =   210
         Width           =   3000
      End
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Canceled As Boolean

Private Sub Form_Load()
Canceled = False
End Sub

Private Sub Pass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Me.Hide
 Canceled = False
ElseIf KeyAscii = 27 Then
 Canceled = True
 Me.Hide
End If
End Sub
