VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Manager"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   Icon            =   "frmUsers.frx":0000
   LinkTopic       =   "Form24"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5625
      Left            =   105
      TabIndex        =   7
      Top             =   90
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   9922
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      BackColor       =   -2147483644
      TabCaption(0)   =   "User List"
      TabPicture(0)   =   "frmUsers.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdNew"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdDelete"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdCopy"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "List1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.Frame Frame1 
         Caption         =   "Settings for default"
         Height          =   1305
         Left            =   210
         TabIndex        =   10
         Top             =   4110
         Width           =   4815
         Begin MSForms.CommandButton cmdSettings 
            Height          =   375
            Left            =   2370
            TabIndex        =   5
            Top             =   810
            Width           =   1515
            Caption         =   "Change Settings..."
            Size            =   "2672;661"
            MousePointer    =   99
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdPass 
            Height          =   375
            Left            =   750
            TabIndex        =   4
            Top             =   810
            Width           =   1515
            Caption         =   "Set Password..."
            Size            =   "2672;661"
            MousePointer    =   99
            Accelerator     =   83
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Label Label3 
            Caption         =   "Use these buttons to specify a password or to setup the user rights for this project."
            Height          =   525
            Left            =   750
            TabIndex        =   11
            Top             =   330
            Width           =   3885
         End
         Begin VB.Image Image2 
            Height          =   675
            Left            =   150
            Picture         =   "frmUsers.frx":0028
            Top             =   300
            Width           =   510
         End
      End
      Begin VB.ListBox List1 
         Height          =   2400
         Left            =   210
         TabIndex        =   0
         Top             =   1560
         Width           =   3285
      End
      Begin MSForms.CommandButton cmdCopy 
         Height          =   375
         Left            =   3630
         TabIndex        =   3
         Top             =   2550
         Visible         =   0   'False
         Width           =   1425
         VariousPropertyBits=   25
         Caption         =   "Make a Copy"
         Size            =   "2514;661"
         Accelerator     =   77
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdDelete 
         Height          =   375
         Left            =   3630
         TabIndex        =   2
         Top             =   2070
         Width           =   1425
         Caption         =   "Delete"
         Size            =   "2514;661"
         Accelerator     =   68
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdNew 
         Height          =   375
         Left            =   3630
         TabIndex        =   1
         Top             =   1590
         Width           =   1425
         ForeColor       =   0
         Caption         =   "New User..."
         PicturePosition =   327683
         Size            =   "2514;661"
         Accelerator     =   78
         MouseIcon       =   "frmUsers.frx":12B2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Label2 
         Caption         =   "&Users"
         Height          =   225
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   $"frmUsers.frx":15CC
         Height          =   735
         Left            =   870
         TabIndex        =   8
         Top             =   510
         Width           =   4095
      End
      Begin VB.Image Image1 
         Height          =   675
         Left            =   240
         Picture         =   "frmUsers.frx":167F
         Top             =   450
         Width           =   540
      End
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   405
      Left            =   3870
      TabIndex        =   6
      Top             =   5835
      Width           =   1470
      ForeColor       =   0
      Caption         =   "Close"
      Size            =   "2593;714"
      Accelerator     =   67
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rec As New ADODB.Recordset
Private Sub cmdClose_Click()
 Unload Me
End Sub


Private Sub cmdCopy_Click()

 Dim Rec1 As New ADODB.Recordset
 Dim rec2 As New ADODB.Recordset
 
 On Error GoTo err
 Dim UserName As String
 Dim UPass As String
 UserName = InputBox("Enter the name of new User", "New User")
 UPass = InputBox("Enter the Password for " & UserName, "New User")
 
 If UserName = "" Then Exit Sub
 Dim i As Integer
 Rec1.CursorLocation = adUseClient
 Rec1.Open "Select * from Users Where loginName='" & List1.Text & "'", con, adOpenStatic, adLockReadOnly
 
 rec2.CursorLocation = adUseClient
 rec2.Open "Select * from users", con, adOpenStatic, adLockOptimistic
 
 rec2.AddNew
 
  rec2.Fields(0) = UserName
 For i = 1 To rec2.Fields.Count - 1
  rec2.Fields(i) = Rec1.Fields(i)
  
 Next i
  rec2.Update
  
 
   Call FillList
 
 
 
 Exit Sub
err:
 MsgBox err.Description
End Sub



Private Sub CmdDelete_Click()
On Error GoTo err
' Is it Attempt to Delete The Administrator ? Fail IT

If UCase(List1.Text) = UCase("Administrator") Then
   MsgBox "User Administrator Cannot Be Deleted.", vbInformation
   Exit Sub
End If


 If MsgBox("Are You Sure To Delete " & List1.Text & " ?", vbQuestion + vbYesNo, "User Settings") = vbYes Then
 
  con.Execute "Delete From Users Where UserName='" & List1.Text & "'"
  Call FillList
 End If
Exit Sub
err:
 MsgBox err.Description
End Sub

Private Sub cmdNew_Click()

Me.Hide
 With frmUserInfo
   .Show_Option = Opt_NewUser
   .Show 1
   Call FillList
 End With

Me.Show 1

End Sub

Private Sub cmdPass_Click()

If UCase(List1.Text) = UCase("Administrator") And CurrentUserName <> "Administrator" Then
   MsgBox "Only Administrator Can Change His Password.", vbInformation
   Exit Sub
End If

Me.Hide
 With frmUserInfo
   .Show_Option = Opt_ResetPassword
   .Tag = List1
   .Show 1
 End With
Me.Show 1
End Sub


Private Sub cmdSettings_Click()

If UCase(List1.Text) = UCase("Administrator") Then
   MsgBox "User Administrator Has Full Control Its Access Rights Cannot Be Changed.", vbInformation
   Exit Sub
End If


Me.Hide
 With frmUserSettings
   .Tag = List1
   .Show 1
 End With
Me.Show 1
End Sub




Private Sub Form_Load()
   Call FillList
End Sub


Private Sub Disable()
 cmdDelete.Enabled = False
 cmdCopy.Enabled = False
 cmdPass.Enabled = False
 cmdSettings.Enabled = False
End Sub

Private Sub Enable()
 cmdDelete.Enabled = True
 'cmdCopy.Enabled = True
 cmdPass.Enabled = True
 cmdSettings.Enabled = True
End Sub

Private Sub List1_Click()
 Call Enable
End Sub


Public Sub FillList()
Dim i As Integer
If rec.State = 1 Then rec.Close
 rec.CursorLocation = adUseClient
 rec.Open "Select UserName from Users", con, adOpenForwardOnly, adLockReadOnly
 
 List1.Clear
 For i = 0 To rec.RecordCount - 1
  List1.AddItem rec.Fields(0)
  rec.MoveNext
 Next i
 rec.Close
 Set rec = Nothing
End Sub
