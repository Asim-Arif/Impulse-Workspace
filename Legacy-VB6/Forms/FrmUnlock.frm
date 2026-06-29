VERSION 5.00
Begin VB.Form FrmUnlock 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unlock..."
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmUnlock.frx":0000
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   2820
      TabIndex        =   3
      Top             =   2010
      Width           =   1785
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1830
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   4515
      Begin VB.TextBox txtUser 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1365
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   840
         Width           =   2940
      End
      Begin VB.TextBox Pass 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1365
         PasswordChar    =   "l"
         TabIndex        =   5
         Top             =   1335
         Width           =   2940
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password To Unlock..."
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   870
         TabIndex        =   4
         Top             =   210
         Width           =   3555
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   165
         Picture         =   "FrmUnlock.frx":0442
         Top             =   210
         Width           =   480
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "&Password :"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   2
         Top             =   1395
         Width           =   900
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "&User Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
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
         TabIndex        =   1
         Top             =   870
         Width           =   1005
      End
   End
End
Attribute VB_Name = "FrmUnlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LoginSucceeded As Boolean

Private Sub cmdOK_Click()
    
    On Error GoTo err
        
    Dim rs As New ADODB.Recordset
    
    'check for correct password
    With rs
        .Open "Select UserID,Password from Users where UserName='" & txtUser.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        If Pass & "" = ![Password] & "" Then
            Screen.MousePointer = vbHourglass
            
            Me.Hide
            UserID = ![UserID]

            Screen.MousePointer = vbDefault
          
            Unload Me
        Else
            MsgBox "Invalid Password, Note That Password is Case Sensitive." & vbNewLine & _
             "Make Sure That Caps Lock Is Not Accidently ON.", vbCritical, "Login"
         
            Pass.SelStart = 0
            Pass.SelLength = Len(Pass)
            Pass.SetFocus
        End If
         
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
    
err:
    MsgBox err.Description
End Sub

Private Sub Form_Activate()
    Pass.SetFocus
End Sub

Private Sub Form_Load()
    txtUser = CurrentUserName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    OurOwnForm = True
End Sub
