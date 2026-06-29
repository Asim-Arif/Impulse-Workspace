VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmUserInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Information"
   ClientHeight    =   2940
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2355
      Left            =   60
      TabIndex        =   5
      Top             =   45
      Width           =   5115
      Begin VB.TextBox txts 
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
         Index           =   2
         Left            =   1875
         PasswordChar    =   "l"
         TabIndex        =   2
         Top             =   1890
         Width           =   2940
      End
      Begin VB.TextBox txts 
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
         Index           =   0
         Left            =   1875
         PasswordChar    =   "l"
         TabIndex        =   0
         Top             =   1020
         Width           =   2940
      End
      Begin VB.TextBox txts 
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
         Index           =   1
         Left            =   1875
         PasswordChar    =   "l"
         TabIndex        =   1
         Top             =   1455
         Width           =   2940
      End
      Begin VB.Image Image1 
         Height          =   435
         Left            =   120
         Picture         =   "frmUserInfo.frx":0000
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lbls 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Confirm Password :"
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
         Index           =   2
         Left            =   225
         TabIndex        =   9
         Top             =   1950
         Width           =   1605
      End
      Begin VB.Label lbls 
         Alignment       =   1  'Right Justify
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
         Left            =   825
         TabIndex        =   8
         Top             =   1065
         Width           =   1005
      End
      Begin VB.Label lbls 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Enter &Password :"
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
         Left            =   435
         TabIndex        =   7
         Top             =   1500
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   $"frmUserInfo.frx":025E
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   630
         Left            =   765
         TabIndex        =   6
         Top             =   210
         Width           =   4260
      End
   End
   Begin MSForms.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3945
      TabIndex        =   4
      Top             =   2505
      Width           =   1245
      Caption         =   " Cancel  "
      PicturePosition =   327683
      Size            =   "2196;661"
      Picture         =   "frmUserInfo.frx":02F2
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdSave 
      Default         =   -1  'True
      Height          =   375
      Left            =   2625
      TabIndex        =   3
      Top             =   2505
      Width           =   1245
      Caption         =   "OK"
      PicturePosition =   327683
      Size            =   "2196;661"
      Accelerator     =   79
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Show_Option As UserInfoOptions
Dim Loaded As Boolean

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo ERR

Select Case Show_Option
   Case Opt_ChangeUserPassword
         If txts(0) & "" <> Con.Execute("Select [Password] From Users Where [UserName]='" & Me.Tag & "'").Fields(0) & "" Then
            MsgBox "Incorrect Old Password", vbInformation
            txts(0).SetFocus
            Exit Sub
         ElseIf txts(1) <> txts(2) Then
            MsgBox "New Password Doesnot Match Confirmed Password.", vbInformation
            txts(2).SetFocus
            Exit Sub
         End If
         Con.Execute "Update Users Set [Password]='" & txts(1) & "' Where [UserName]='" & Me.Tag & "'"
         'PlaySaveSound
         MsgBox "Password Has Been Changed Successfully.", vbInformation
   Case Opt_ResetPassword
         If txts(1) <> txts(2) Then
            MsgBox "New Password Doesnot Match Confirmed Password.", vbInformation
            txts(2).SetFocus
            Exit Sub
         End If
         Con.Execute "Update Users Set [Password]='" & txts(1) & "' Where [UserName]='" & Me.Tag & "'"
'         PlaySaveSound
         MsgBox "Password Has Been Changed Successfully.", vbInformation
   Case Opt_NewUser
         If txts(0) = "" Then
            MsgBox "Invalid UserName", vbInformation
            txts(0).SetFocus
            Exit Sub
         ElseIf txts(1) <> txts(2) Then
            MsgBox "Password Doesnot Match Confirmed Password.", vbInformation
            txts(2).SetFocus
            Exit Sub
         End If
         Con.Execute "insert into Users([UserName],[Password]) Values ('" & txts(0) & "','" & txts(1) & "')"
   End Select
   
Unload Me
   
Exit Sub

ERR:
   MsgBox "Error :" & ERR.Number & vbNewLine & ERR.Description, vbInformation
   
End Sub

Private Sub Form_Activate()
If Loaded Then Exit Sub

If Show_Option = Opt_ChangeUserPassword Then
   Label1.Caption = "Enter Your Old Password, Then Enter New Password, Confirm It And Press Save Button To Change You Password"
   Set txts(0).Font = txts(1).Font
   txts(0).PasswordChar = txts(1).PasswordChar
   lbls(0).Caption = "Enter &Old Password :"
ElseIf Show_Option = Opt_ResetPassword Then
   Label1.Caption = "Enter New Password For User, Then Confirm It, You Can Change User Access By Changing Its Access Settings..."
   txts(0).Font.Name = "Tahoma"
   txts(0) = Me.Tag
   txts(0).PasswordChar = ""
   txts(0).Enabled = False
ElseIf Show_Option = Opt_NewUser Then
   Label1.Caption = "Enter User Name and Password For New User, By Default It Has Operater Level Access You Can Change User Access By Changing Its Access Settings..."
   txts(0).PasswordChar = ""
   txts(0).Font.Name = "Tahoma"
End If
Loaded = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frmUSerInfo = Nothing
End Sub

Private Sub txts_GotFocus(Index As Integer)
   txts(Index).SelStart = 0
   txts(Index).SelLength = Len(txts(Index))
End Sub
