VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmProductionGenSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Settings"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4680
   Begin VB.Frame Frame2 
      Caption         =   "Other General Options :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1275
      Left            =   75
      TabIndex        =   10
      Top             =   1935
      Width           =   4500
      Begin VB.TextBox txtRepairRate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1860
         TabIndex        =   11
         Top             =   225
         Width           =   750
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Repair Deduction Rate :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   12
         Top             =   240
         Width           =   1725
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Advance Salary Settings :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1275
      Left            =   75
      TabIndex        =   5
      Top             =   585
      Width           =   4500
      Begin VB.TextBox txtAdvSalPer 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2220
         TabIndex        =   6
         Top             =   885
         Width           =   750
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00AFBDCB&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Process"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   9
         Top             =   285
         Width           =   4320
      End
      Begin MSForms.ComboBox cmbProcess 
         Height          =   285
         Left            =   60
         TabIndex        =   8
         Top             =   555
         Width           =   4320
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "7620;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Advance Salary Percentage :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   3
         Left            =   75
         TabIndex        =   7
         Top             =   930
         Width           =   2100
      End
   End
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4650
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "General Settings"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   390
         Index           =   1
         Left            =   630
         TabIndex        =   1
         Top             =   135
         Width           =   3435
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "General Settings"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Index           =   2
         Left            =   645
         TabIndex        =   2
         Top             =   150
         Width           =   3435
      End
   End
   Begin MSForms.CommandButton cmdOK 
      Height          =   360
      Left            =   1695
      TabIndex        =   4
      Top             =   3285
      Width           =   1425
      ForeColor       =   0
      Caption         =   "  OK"
      PicturePosition =   327683
      Size            =   "2514;635"
      Picture         =   "frmProductionGenSettings.frx":0000
      Accelerator     =   79
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   360
      Left            =   3165
      TabIndex        =   3
      Top             =   3285
      Width           =   1425
      ForeColor       =   0
      Caption         =   "  Close"
      PicturePosition =   327683
      Size            =   "2514;635"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmProductionGenSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbProcess_Change()
    Call cmbProcess_Click
End Sub

Private Sub cmbProcess_Click()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select AdvSalPer From GeneralSettings Where ProcessID=" & cmbProcess.List(cmbProcess.ListIndex, 1), con, adOpenForwardOnly, adLockReadOnly
        If .EOF Then
            txtAdvSalPer.Tag = "Add"
            txtAdvSalPer = 0
        Else
            txtAdvSalPer.Tag = "Edit"
            txtAdvSalPer = Val(.Fields(0) & "")
        End If
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    On Error GoTo err
    If Not cmbProcess.MatchFound Then Exit Sub
    
    Call StartTrans(con)
    
    If txtAdvSalPer.Tag = "Add" Then
        con.Execute "Insert Into GeneralSettings(ProcessID,AdvSalPer) Values(" & cmbProcess.List(cmbProcess.ListIndex, 1) & _
        "," & Val(txtAdvSalPer) & ")"
    Else
        con.Execute "Update GeneralSettings Set AdvSalPer=" & Val(txtAdvSalPer) & _
        " Where ProcessID=" & cmbProcess.List(cmbProcess.ListIndex, 1)
    End If
    
    
    con.Execute "Update GeneralSettings1 Set RepairRate=" & Val(txtRepairRate)
    
    con.CommitTrans
    'Unload Me
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select ProcessID,Description From Processes", con, adOpenForwardOnly, adLockReadOnly
        cmbProcess.Clear
        Do Until .EOF
            cmbProcess.AddItem ![Description] & ""
            cmbProcess.List(.AbsolutePosition - 1, 1) = ![ProcessID] & ""
            .MoveNext
        Loop
        .Close
        
        .Open "Select * From GeneralSettings1", con, adOpenForwardOnly, adLockReadOnly
        txtRepairRate = Val(!RepairRate & "")
        .Close
    End With
    Set rs = Nothing
    cmbProcess.AddItem "Sub Contractors"
    cmbProcess.List(cmbProcess.ListCount - 1, 1) = -1
    cmbProcess.AddItem "Repair Venders"
    cmbProcess.List(cmbProcess.ListCount - 1, 1) = 0
    'Call AddToCombo(cmbProcess, "Description", "Processes", , , "ProcessID")
       
    Exit Sub
err:
    MsgBox err.Description
End Sub



Private Sub txtDays_Validate(Cancel As Boolean)
    If Val(txtDays) = 0 Then Cancel = True
End Sub
