VERSION 5.00
Begin VB.Form frmBankOD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Over Draft"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3390
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   3390
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAllowOD 
      Caption         =   "Allow Overdraft"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   90
      TabIndex        =   6
      Top             =   180
      Width           =   1905
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   2025
      TabIndex        =   5
      Top             =   1710
      Width           =   1275
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   720
      TabIndex        =   4
      Top             =   1710
      Width           =   1275
   End
   Begin VB.TextBox txtODRate 
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtODAmt 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Overdraft Rate (%)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1710
      TabIndex        =   2
      Top             =   750
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Overdraft Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   750
      Width           =   1320
   End
End
Attribute VB_Name = "frmBankOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strAccNo As String

Private Sub chkAllowOD_Click()
    If chkAllowOD.value = vbUnchecked Then
        txtODAmt = ""
        txtODRate = ""
        txtODAmt.Locked = True
        txtODRate.Locked = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    
    con.Execute "Update BankAccounts Set AllowOD=" & Abs(chkAllowOD.value) & ",ODAmt=" & Val(txtODAmt) & _
     ",ODRate=" & Val(txtODRate) & " Where Accno='" & strAccNo & "'"
     
    Unload Me
    
End Sub


Public Sub ShowMe(p_strAccNo As String)

    strAccNo = p_strAccNo
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select AllowOD,ODAmt,ODRate FROM BankAccounts Where AccNo='" & strAccNo & "'", con, adOpenForwardOnly, adLockReadOnly
        If !AllowOD Then
            chkAllowOD.value = vbChecked
        Else
            chkAllowOD.value = vbUnchecked
        End If
        txtODAmt = Val(!ODAmt & "")
        txtODRate = Val(!ODRate & "")
        .Close
    End With
    Set rs = Nothing
    
    Me.Show 1
    
End Sub
