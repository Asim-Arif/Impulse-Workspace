VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login to CTW Accounts"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   ControlBox      =   0   'False
   Icon            =   "Frmlogin.frx":0000
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1965
      TabIndex        =   5
      Top             =   2010
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   3315
      TabIndex        =   4
      Top             =   2010
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Height          =   1830
      Left            =   90
      TabIndex        =   0
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
         TabIndex        =   7
         Top             =   1335
         Width           =   2940
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1365
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   900
         Width           =   2940
      End
      Begin VB.Label Label1 
         Caption         =   "Select User Name And Enter A Valid Password For That User To Login ..."
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
         Left            =   870
         TabIndex        =   6
         Top             =   210
         Width           =   3555
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   165
         Picture         =   "Frmlogin.frx":0442
         Top             =   210
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
         TabIndex        =   3
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
         TabIndex        =   2
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
        con.Close
        Set con = Nothing
        Unload frmMain
        End
    End If
End Sub

Private Sub cmdOK_Click()
Dim check As String
    'check for correct password
    
   rec.Open "Select password from Users where UserName='" & Combo1.Text & "'", con, adOpenForwardOnly, adLockReadOnly
    If Pass & "" = rec![password] & "" Then
    
     

    Screen.MousePointer = vbHourglass
      CurrentUserName = Combo1
      frmMain.mnuLogOff.Caption = "Log Off " & CurrentUserName & " ..."
      Me.Hide
      Call DoStartUpActivities
    Screen.MousePointer = vbDefault
      
      Unload Me
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
   Pass.SetFocus
End Sub

Private Sub Form_Load()
   
   ''''''''''''''''' Check FOR SMC ''''''''''''''''''''
   
If GetServerDate >= #7/15/2003# Then
    MsgBox "Cannot Continue, CTW Accounting System Has Expired..." & vbNewLine & _
    "Please Contact Computer Technology World, 0432-255960", vbCritical
    Call CleanUp
End If


   Call AddToCombo(Combo1, "UserName", "Users", "Order By UserName")
   If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
   
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Unload Me
End Sub

Private Sub DoStartUpActivities()
Screen.MousePointer = vbHourglass
   If Not OutOfFinancialYear Then Call CheckForAutoVouchers
   
   Call LoadUserSettings
   Call frmMain.ApplySecuritySettings
   
   If OutOfFinancialYear Then Exit Sub
   
   If AutoDepriciation Then Call CheckDepriciation
   
   
Screen.MousePointer = vbDefault
End Sub
