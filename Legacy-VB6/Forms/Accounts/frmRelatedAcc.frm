VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmRelatedAcc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Related Accounts Info"
   ClientHeight    =   3420
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fAcrAcc 
      Caption         =   "Related Expense Account Info :"
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
      Height          =   1845
      Left            =   255
      TabIndex        =   21
      Top             =   960
      Visible         =   0   'False
      Width           =   5310
      Begin VB.TextBox txtAcrAccTitle 
         Height          =   315
         Left            =   1200
         MaxLength       =   255
         TabIndex        =   2
         Top             =   825
         Width           =   3960
      End
      Begin VB.TextBox txtAmt 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4020
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   1275
         Width           =   1125
      End
      Begin VB.TextBox txtExpBal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   1275
         Width           =   1095
      End
      Begin VB.Frame fExp 
         Height          =   510
         Left            =   120
         TabIndex        =   22
         Top             =   210
         Width           =   5055
         Begin VB.OptionButton OptExp 
            Caption         =   "Prepaid  Expense Account"
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
            Height          =   225
            Index           =   1
            Left            =   2580
            TabIndex        =   1
            Top             =   195
            Width           =   2280
         End
         Begin VB.OptionButton OptExp 
            Caption         =   "Accured Expense Account"
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
            Height          =   225
            Index           =   0
            Left            =   180
            TabIndex        =   0
            Top             =   195
            Width           =   2280
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Account Title :"
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
         Height          =   195
         Index           =   25
         Left            =   135
         TabIndex        =   26
         Top             =   885
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Amount / Month :"
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
         Height          =   195
         Index           =   18
         Left            =   2745
         TabIndex        =   25
         Top             =   1335
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Opening Bal. :"
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
         Height          =   195
         Index           =   19
         Left            =   150
         TabIndex        =   24
         Top             =   1335
         Width           =   1020
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " CR"
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
         Height          =   315
         Index           =   20
         Left            =   2280
         TabIndex        =   23
         Top             =   1275
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      Height          =   825
      Left            =   255
      TabIndex        =   27
      Top             =   90
      Width           =   5325
      Begin VB.Label lblAccTitle 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   1725
         TabIndex        =   31
         Top             =   435
         Width           =   3525
      End
      Begin VB.Label lblAccNo 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   285
         Left            =   90
         TabIndex        =   30
         Top             =   435
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Account # :"
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
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Account Title :"
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
         Height          =   195
         Index           =   0
         Left            =   1755
         TabIndex        =   28
         Top             =   210
         Width           =   1035
      End
   End
   Begin VB.Frame fAsset 
      Caption         =   "Depriciation Account Information :"
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
      Height          =   1875
      Left            =   255
      TabIndex        =   7
      Top             =   930
      Visible         =   0   'False
      Width           =   5310
      Begin VB.TextBox txtAccDep 
         Height          =   315
         Left            =   135
         MaxLength       =   255
         TabIndex        =   12
         Top             =   480
         Width           =   3570
      End
      Begin VB.TextBox txtDep 
         Height          =   315
         Left            =   135
         MaxLength       =   255
         TabIndex        =   11
         Top             =   1050
         Width           =   3570
      End
      Begin VB.TextBox txtAccDepBal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3750
         MaxLength       =   25
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   480
         Width           =   1185
      End
      Begin VB.TextBox txtdepBal 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3750
         MaxLength       =   25
         TabIndex        =   9
         Text            =   "0.00"
         Top             =   1050
         Width           =   1185
      End
      Begin VB.TextBox txtRate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4185
         MaxLength       =   5
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   1455
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Acc. Dep. Account Title"
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
         Height          =   195
         Index           =   12
         Left            =   150
         TabIndex        =   20
         Top             =   255
         Width           =   1680
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Depriciation Account"
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
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   19
         Top             =   825
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Opening Bal."
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
         Height          =   195
         Index           =   11
         Left            =   3765
         TabIndex        =   18
         Top             =   240
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Opening Bal."
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
         Height          =   195
         Index           =   13
         Left            =   3765
         TabIndex        =   17
         Top             =   825
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DR"
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
         Index           =   14
         Left            =   4995
         TabIndex        =   16
         Top             =   1125
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CR"
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
         Index           =   15
         Left            =   4995
         TabIndex        =   15
         Top             =   525
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Yearly Depriciation Rate :"
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
         Height          =   195
         Index           =   16
         Left            =   2325
         TabIndex        =   14
         Top             =   1500
         Width           =   1830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "%"
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
         Index           =   17
         Left            =   4995
         TabIndex        =   13
         Top             =   1515
         Width           =   195
      End
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   375
      Left            =   4290
      TabIndex        =   6
      Top             =   2940
      Width           =   1290
      Caption         =   "Cancel    "
      PicturePosition =   327683
      Size            =   "2275;661"
      Picture         =   "frmRelatedAcc.frx":0000
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdOK 
      Height          =   375
      Left            =   2925
      TabIndex        =   5
      Top             =   2940
      Width           =   1290
      Caption         =   "Save      "
      PicturePosition =   327683
      Size            =   "2275;661"
      Picture         =   "frmRelatedAcc.frx":0112
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmRelatedAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Loaded As Boolean
Public RelatedExpenseAccount As Boolean, RelatedAssetAccount As Boolean, AddNew As Boolean, Edit As Boolean
Dim PermissionDenied As Boolean
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
   If SaveAccInfo Then Unload Me
End Sub

Private Sub Form_Activate()
'Check For User Access
If PermissionDenied Then
   Me.Hide
   MsgBox "You Donot Have Permissions To Use This Option. Please Contact Your Administrator", vbInformation
   Unload Me
   Exit Sub
End If


If Not Loaded Then
   fAsset.Visible = RelatedAssetAccount
   fAcrAcc.Visible = RelatedExpenseAccount
   Call GetAccInfo
   If RelatedExpenseAccount Then
      If Edit Then Call GetExpAccInfo
      fExp.Enabled = AddNew
      Me.Caption = "Related Expense Accounts"
   ElseIf RelatedAssetAccount Then
      Call GetAssetAccInfo
      Me.Caption = "Related Asset Accounts"
   End If
Loaded = False
End If
End Sub

Private Sub Form_Load()
PermissionDenied = Not UserHasAccess("EditAccount")

   
   RelatedExpenseAccount = False
   RelatedAssetAccount = False
   AddNew = False
   Edit = False
End Sub


Private Sub GetExpAccInfo()
   Dim rs As New ADODB.Recordset
   With rs
      .Open "Select * from VExpenseAccounts Where ExpAccno='" & Me.Tag & "'", con, adOpenForwardOnly, adLockReadOnly
      If .RecordCount = 0 Then Exit Sub
      LblAccNo = Me.Tag
      lblAccTitle = ![ExpAccTitle] & ""
      txtAcrAccTitle.Tag = ![AcrExpAccNo] & ""
      txtAcrAccTitle = ![AcrAccTitle] & ""
      OptExp(1) = ![PrePaid]
      txtExpBal = Abs(Val(![AcrBalance] & ""))
      txtAmt = Val(![Amount] & "")
      .Close
   End With
   Set rs = Nothing
   
End Sub

Private Sub GetAssetAccInfo()
   
   Dim rs As New ADODB.Recordset
   
   With rs
      .Open "Select * from VAssets Where Accno='" & Me.Tag & "'", con, adOpenForwardOnly, adLockReadOnly
      If .RecordCount = 0 Then
         .Close
         .Open "Select * from VAssets Where AccDepAccNo='" & Me.Tag & "'", con, adOpenForwardOnly, adLockReadOnly
         If .RecordCount = 0 Then Exit Sub
      End If
      'Asset Account
      LblAccNo = Me.Tag
      lblAccTitle = ![AccTitle] & ""
      
      'Accumulated Depriciation Account
      txtAccDep = ![AccDepTitle] & ""
      txtAccDep.Tag = ![AccDepAccno] & ""
      
      txtAccDepBal = Abs(Val(![AccDepBal] & ""))
      txtAccDepBal.Tag = Val(![AccDepBal] & "")
      'Depriciation Account
      txtDep = ![DepTitle] & ""
      txtDep.Tag = ![DepAccNo] & ""
      
      txtdepBal = Abs(Val(![Depbal] & ""))
      txtdepBal.Tag = Val(![Depbal] & "")
      
      txtRate = Val(![PercentRate] & "")
      
      .Close
   
   End With
   
   Set rs = Nothing
   
End Sub

Private Sub OptExp_Click(Index As Integer)
   Label1(20) = IIf(OptExp(0), " CR", " DR")
If AddNew Then
      If OptExp(0) Then txtAcrAccTitle = "Accrued " & lblAccTitle Else txtAcrAccTitle = "Prepaid " & lblAccTitle
'      txtAcrAccTitle = "Accured " & lblAccTitle
End If
End Sub



Private Function AddAcrExpAcc(AccTitle As String) As Boolean
   Dim NextAccNo As String
   
   Dim MainAcc As String
   
   ' "24-005" is Accrued Expense Main Account and  "14-001" is Prepaid Main Account
   con.BeginTrans
   If OptExp(0) Then MainAcc = "24-005" Else MainAcc = "14-001"
   'Opening Balance of Accured
   Bal = IIf(OptExp(0), -Val(txtExpBal), Val(txtExpBal))
   
   NextAccNo = MainAcc & "-" & Format(Val(GetMax("Right(Accno,3)", "Accounts", "Where Subaccof='" & MainAcc & "'")) + 1, "000")
   
   con.Execute "Insert into Accounts Values('" & NextAccNo & "','" & AccTitle & "','" & MainAcc & "','" & Left(MainAcc, 2) & "'," & Bal & ",'" & Date & "'," & Bal & " ,False,True)"
   
   con.Execute "Insert into ExpenseAccounts values ('" & LblAccNo & "','" & NextAccNo & "'," & Val(txtAmt) & "," & OptExp(1) & ")"



   AddAcrExpAcc = True
   
con.CommitTrans

Exit Function
ERR:
   MsgBox "Cannot Update Changes To Expense Accounts Due to Error :" & ERR.Number & _
   vbNewLine & ERR.Description, vbInformation
   AddAcrExpAcc = False
   con.RollbackTrans
End Function


Private Function SaveAccInfo() As Boolean
  
   If RelatedExpenseAccount Then
      If AddNew Then
         If Not AddAcrExpAcc(txtAcrAccTitle) Then Exit Function
      Else
         If Not UpdateExpAccount Then Exit Function
      End If
      
   ElseIf RelatedAssetAccount Then
      If Not UpdateAssetAccount Then Exit Function
   End If
   SaveAccInfo = True
End Function



Private Function UpdateExpAccount() As Boolean
UpdateExpAccount = False
Dim Diff As Double

con.BeginTrans
   Diff = Val(txtExpBal.Tag) - Val(txtExpBal)
   con.Execute "Update Accounts Set Acctitle='" & txtAcrAccTitle & "',OpenBal=" & Val(txtExpBal) & _
               ",Balance=Balance+" & Diff & " Where AccNo='" & txtAcrAccTitle.Tag & "'"
                         
   'Adjust The Ledger
      
   Call adjustLedger(txtAcrAccTitle.Tag, #1/1/1900#, 0, Diff, True)
   con.Execute "Update ExpenseAccounts Set Amount=" & Val(txtAmt) & " Where ExpAccno='" & LblAccNo & "'"
   
   UpdateExpAccount = True
   
con.CommitTrans

Exit Function
ERR:
   MsgBox "Cannot Update Changes To Expense Accounts Due to Error :" & ERR.Number & _
   vbNewLine & ERR.Description, vbInformation
   UpdateExpAccount = False
   con.RollbackTrans
End Function

Private Function UpdateAssetAccount() As Boolean
UpdateAssetAccount = False
Dim Diff As Double
'''''''''''''''''''''' Accumulated Depriciation Account '''''''''''''''''''
con.BeginTrans
Diff = Val(txtAccDepBal.Tag) - Val(txtAccDepBal)
   con.Execute "Update Accounts Set Acctitle='" & txtAccDep & "',OpenBal=" & -Val(txtAccDepBal) & _
               ",Balance=Balance+" & Diff & " Where AccNo='" & txtAccDep.Tag & "'"
               
   'Adjust The Ledger
   Call adjustLedger(txtAccDep.Tag, #1/1/1900#, 0, Diff, True)
   
''''''''''''''''''''''  Depriciation Account '''''''''''''''''''

   Diff = Val(txtdepBal.Tag) - Val(txtdepBal)
   con.Execute "Update Accounts Set Acctitle='" & txtDep & "',OpenBal=" & Val(txtdepBal) & _
               ", Balance=Balance+" & Diff & " Where AccNo='" & txtDep.Tag & "'"
               
   'Adjust The Ledger
   Call adjustLedger(txtDep.Tag, #1/1/1900#, 0, Diff, True)
   
''''''''''''''''''''''  Asset Accounts Table  '''''''''''''''''''

con.Execute "Update Assets Set PercentRate=" & Val(txtRate) & " Where Accno='" & LblAccNo & "'"
   
   
UpdateAssetAccount = True
con.CommitTrans

Exit Function
ERR:
   MsgBox "Cannot Update Changes To Expense Accounts Due to Error :" & ERR.Number & _
   vbNewLine & ERR.Description, vbInformation
UpdateAssetAccount = False
con.RollbackTrans
   
End Function


Private Sub GetAccInfo()
On Error GoTo ERR
   LblAccNo = Me.Tag
   lblAccTitle = con.Execute("Select AccTitle From Accounts Where AccNo='" & Me.Tag & "'").Fields(0) & ""
   Exit Sub
ERR:
   MsgBox "Cannot Get Account Info This Account Is Not Found", vbInformation
End Sub
