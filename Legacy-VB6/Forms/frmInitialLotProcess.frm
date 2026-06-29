VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmInitialLotProcess 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Initial Lot Process..."
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4680
   Begin VB.TextBox txtDays 
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
      Left            =   3105
      TabIndex        =   8
      Top             =   1230
      Width           =   750
   End
   Begin VB.Frame fTitle 
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4650
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set Initial Lot Process"
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
         TabIndex        =   2
         Top             =   135
         Width           =   3405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set Initial Lot Process"
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
         TabIndex        =   3
         Top             =   150
         Width           =   3405
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00AFBDCB&
      BackStyle       =   0  'Transparent
      Caption         =   "Delete Entry Without Authentication After                  Days."
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
      Height          =   255
      Index           =   3
      Left            =   60
      TabIndex        =   7
      Top             =   1260
      Width           =   4500
   End
   Begin MSForms.ComboBox cmbProcess 
      Height          =   285
      Left            =   60
      TabIndex        =   6
      Top             =   870
      Width           =   4515
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "7964;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdOK 
      Height          =   360
      Left            =   1680
      TabIndex        =   5
      Top             =   1845
      Width           =   1425
      ForeColor       =   0
      Caption         =   "  OK"
      PicturePosition =   327683
      Size            =   "2514;635"
      Picture         =   "frmInitialLotProcess.frx":0000
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
      Left            =   3150
      TabIndex        =   4
      Top             =   1845
      Width           =   1425
      ForeColor       =   0
      Caption         =   "  Cancel"
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
      TabIndex        =   0
      Top             =   600
      Width           =   4515
   End
End
Attribute VB_Name = "frmInitialLotProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo err
    If Not cmbProcess.MatchFound Then Exit Sub
    
    Call StartTrans(con)
    
    con.Execute "Delete From InitialLotProcess"
    con.Execute "Insert Into InitialLotProcess " & _
    "(ProcessID,NoOfDays) Values(" & cmbProcess.List(cmbProcess.ListIndex, 1) & "," & Val(txtDays) & ")"
    
    con.CommitTrans
    Unload Me
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    On Error GoTo err
    
    Call AddToCombo(cmbProcess, "Description", "Processes", , , "ProcessID")
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select * From InitialLotProcess", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            txtDays = Val(![NoOfDays] & "")
            For i = 0 To cmbProcess.ListCount - 1
                If cmbProcess.List(i, 1) = ![ProcessID] & "" Then
                    cmbProcess.ListIndex = i
                    Exit For
                End If
            Next i
        End If
        .Close
    End With
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
End Sub



Private Sub txtDays_Validate(Cancel As Boolean)
    If Val(txtDays) = 0 Then Cancel = True
End Sub
