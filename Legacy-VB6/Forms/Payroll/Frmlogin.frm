VERSION 5.00
Begin VB.Form FrmLogin 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   Icon            =   "Frmlogin.frx":0000
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUser 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1455
      TabIndex        =   0
      Text            =   "Administrator"
      Top             =   945
      Width           =   2940
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1965
      TabIndex        =   2
      Top             =   2010
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   3315
      TabIndex        =   3
      Top             =   2010
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Height          =   1830
      Left            =   90
      TabIndex        =   4
      Top             =   45
      Width           =   4515
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
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1365
         PasswordChar    =   "l"
         TabIndex        =   1
         Top             =   1320
         Width           =   2940
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1365
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   900
         Visible         =   0   'False
         Width           =   2940
      End
      Begin VB.Image Image2 
         Height          =   705
         Left            =   360
         Picture         =   "Frmlogin.frx":0442
         Top             =   120
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "Please Enter User Name and Password"
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
         Height          =   435
         Left            =   2400
         TabIndex        =   8
         Top             =   240
         Width           =   1875
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   1800
         Picture         =   "Frmlogin.frx":0E75
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "&Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   390
         TabIndex        =   7
         Top             =   1425
         Width           =   825
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "&User Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   6
         Top             =   930
         Width           =   945
      End
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public LoginSucceeded As Boolean
Dim rec As New ADODB.Recordset

Private Sub cmdCancel_Click()
 If MsgBox("Are You Sure To Quit ?", vbYesNo + vbQuestion) = vbYes Then
   Con.Close
   Set Con = Nothing
   Unload MainForm
   Unload Me
   End
 End If
End Sub

Private Sub cmdOK_Click()
Combo1.ListIndex = -1
Call FindInCombo(Combo1, txtUser)
If LCase(Combo1) <> LCase(txtUser) Then
    MsgBox "Invalid User Name", vbInformation
    txtUser.SelStart = 0
    txtUser.SelLength = Len(Pass)
    txtUser.SetFocus
    Exit Sub
End If




Dim check As String
    'check for correct password
    
   rec.Open "Select password from Users where UserName='" & Combo1.Text & "'", Con, adOpenForwardOnly, adLockReadOnly
    If Pass & "" = rec![password] & "" Then
    
     

    Call ShowWait
      CurrentUserName = Combo1
      MainForm.mnuLogOff.Caption = "Log Off " & CurrentUserName & " ..."
      Me.Hide
      'Apply User Access Rights
      Call MainForm.ApplySecuritySettings
      
    Call HideWait
      
      Unload Me
      If UserHasAccess("Reminders") Then frmReminders.Show 1
    Else
      MsgBox "Cannot Login Due To Invalid Password, Note That Password is Case Sensitive." & vbNewLine & _
      "Make Sure That Caps Lock Is Not Accidently ON.", vbCritical, "Login"
      Pass.SelStart = 0
      Pass.SelLength = Len(Pass)
      Pass.SetFocus
    End If
     
   rec.Close
   Set rec = Nothing
End Sub


Private Sub Combo1_Click()
   UserName = Combo1.Text
End Sub

Private Sub Form_Activate()
'   Pass.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Form_Load()
Me.KeyPreview = True
   Call AddToCombo(Combo1, "UserName", "Users", "Order By UserName")
   If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub

Private Sub txtUser_GotFocus()
cmdOK.Default = False
End Sub
Private Sub txtUser_LostFocus()
cmdOK.Default = True
End Sub
