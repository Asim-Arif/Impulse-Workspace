VERSION 5.00
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmCallLogin 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login To Call Settings"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   ControlBox      =   0   'False
   Icon            =   "frmCallLogin.frx":0000
   LinkTopic       =   "Form14"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   1455
      TabIndex        =   0
      Top             =   960
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1980
      TabIndex        =   2
      Top             =   2010
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   3330
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
         Top             =   1260
         Width           =   2940
      End
      Begin VB.Label lblCaption 
         Caption         =   "Select Employee Name And Enter A Valid Password For That Employee To Check The Calls Detail ..."
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
         Height          =   570
         Left            =   870
         TabIndex        =   7
         Top             =   210
         Width           =   3555
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   165
         Picture         =   "frmCallLogin.frx":0442
         Top             =   210
         Width           =   480
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "&Password :"
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
         TabIndex        =   6
         Top             =   1335
         Width           =   945
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Employee :"
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
         Left            =   390
         TabIndex        =   5
         Top             =   930
         Width           =   945
      End
   End
End
Attribute VB_Name = "FrmCallLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public LoginSucceeded As Boolean
Dim rec As New ADODB.Recordset

Private Sub cmbEmp_GotFocus()

cmdOK.Default = False

End Sub

Private Sub cmbEmp_LostFocus()

cmdOK.Default = True

End Sub

Private Sub cmbEmp_matched()

Pass.SetFocus

End Sub

Private Sub cmdCancel_Click()
    Unload Me
    
End Sub

Private Sub cmdOK_Click()

If cmbEmp.MatchFound = False Then
    MsgBox "The Emplyee Name That You Write Does Not Exist !", vbInformation
    cmbEmp.SetFocus
ElseIf cmbEmp.ID <> Pass Then
    MsgBox "Cannot Login Due To Invalid Password, Note That Password is Case Sensitive." & vbNewLine & _
    "Make Sure That Caps Lock Is Not Accidently ON.", vbCritical, "Login"
    Pass.SelStart = 0
    Pass.SelLength = Len(Pass)
    Pass.SetFocus
Else

    With frmNumbersDetail
        .cmbEmp.Text = cmbEmp.Text
        .cmbEmp.Enabled = False
        Unload Me
        .Show
    End With
    
End If
     
End Sub


Private Sub Combo1_Click()
   UserName = Combo1.Text
End Sub

Private Sub Form_Activate()
   'Pass.SetFocus
End Sub

Private Sub Form_Load()

cmbEmp.ListHeight = 1000
cmbEmp.AddVals con, "Name", "Employees", "Password"

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
