VERSION 5.00
Begin VB.Form frmRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CTW Acconting System Registration"
   ClientHeight    =   4035
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5925
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   5925
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame f3 
      Height          =   3195
      Left            =   135
      TabIndex        =   7
      Top             =   105
      Width           =   5625
      Begin VB.Frame Frame2 
         Height          =   1575
         Left            =   315
         TabIndex        =   8
         Top             =   1485
         Width           =   4950
         Begin VB.TextBox txtCompany 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   135
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   420
            Width           =   4650
         End
         Begin VB.TextBox txtReg 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   135
            MaxLength       =   5
            TabIndex        =   1
            Top             =   1170
            Width           =   1095
         End
         Begin VB.TextBox txtReg 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   1320
            MaxLength       =   5
            TabIndex        =   2
            Top             =   1170
            Width           =   1095
         End
         Begin VB.TextBox txtReg 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   2505
            MaxLength       =   5
            TabIndex        =   3
            Top             =   1170
            Width           =   1095
         End
         Begin VB.TextBox txtReg 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   3690
            MaxLength       =   5
            TabIndex        =   4
            Top             =   1170
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Company Name "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   270
            Index           =   2
            Left            =   135
            TabIndex        =   11
            Top             =   135
            Width           =   4380
         End
         Begin VB.Label Label3 
            Caption         =   "Registration Key"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   270
            Index           =   0
            Left            =   150
            TabIndex        =   9
            Top             =   870
            Width           =   4380
         End
      End
      Begin VB.Label Label3 
         Caption         =   $"frmRegister.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1170
         Index           =   1
         Left            =   870
         TabIndex        =   10
         Top             =   240
         Width           =   4395
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   195
         Picture         =   "frmRegister.frx":00EA
         Top             =   315
         Width           =   480
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4365
      TabIndex        =   6
      Top             =   3525
      Width           =   1425
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   3525
      Width           =   1425
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CancelButton_Click()
    If MsgBox("Are You Sure To Quit ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Unload Me
    Call CleanUp
End Sub

Private Sub Form_Load()
    txtCompany = company
End Sub

Private Sub OKButton_Click()
Dim RegNo As String
   For i = 0 To txtReg.Count - 1
       RegNo = RegNo & txtReg(i)
   Next

   If Not Registered(txtCompany, RegNo) Then
      MsgBox "Invalid Registration Key, You Must Enter The Valid Registration Key", vbInformation
      Exit Sub
   End If
   
    Con.Execute "Update RegInfo Set CompanyName='" & txtCompany & "',RegKey='" & RegNo & "'"
    
   If IsPackageRegistered Then
      MsgBox "Registered Successfully." & vbNewLine & "Thank You For Registeration.", vbInformation
      Unload Me
   Else
      MsgBox "Cannot Register The Package...", vbCritical
   End If
End Sub


Private Sub txtreg_Change(Index As Integer)
If Len(txtReg(Index)) = 5 Then
 If Index < 3 Then
  txtReg(Index + 1).SetFocus
 Else
  OKButton.SetFocus
 End If
End If
End Sub

