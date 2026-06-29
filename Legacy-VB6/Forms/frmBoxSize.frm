VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmBoxSize 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Define Box Size"
   ClientHeight    =   1965
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5475
   Icon            =   "frmBoxSize.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   105
      TabIndex        =   5
      Top             =   15
      Width           =   3435
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Tag             =   "ItemUsage"
         Top             =   1485
         Width           =   2475
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "4366;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "3rd Dimension"
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
         Index           =   1
         Left            =   135
         TabIndex        =   8
         Top             =   1275
         Width           =   1005
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Tag             =   "ItemUsage"
         Top             =   945
         Width           =   2475
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "4366;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "2nd Dimension"
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
         Index           =   0
         Left            =   135
         TabIndex        =   7
         Top             =   735
         Width           =   1035
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Tag             =   "ItemUsage"
         Top             =   375
         Width           =   2475
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "4366;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ist Dimension"
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
         Left            =   135
         TabIndex        =   6
         Top             =   165
         Width           =   960
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4185
      Picture         =   "frmBoxSize.frx":0442
      Top             =   1260
      Width           =   765
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   360
      Left            =   3630
      TabIndex        =   4
      Top             =   555
      Width           =   1755
      ForeColor       =   0
      Caption         =   "Cancel       "
      PicturePosition =   327683
      Size            =   "3096;635"
      Picture         =   "frmBoxSize.frx":0C93
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
      Height          =   360
      Left            =   3630
      TabIndex        =   3
      Top             =   105
      Width           =   1755
      ForeColor       =   0
      Caption         =   "Save           "
      PicturePosition =   327683
      Size            =   "3096;635"
      Picture         =   "frmBoxSize.frx":0DA5
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
Attribute VB_Name = "frmBoxSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    
    On Error GoTo err
    
    If Val(TBox(0)) = 0 Then
        MsgBox "Dimensions Can't Be 0.", vbInformation
        TBox(0).SetFocus
        Exit Sub
    ElseIf Val(TBox(1)) = 0 Then
        MsgBox "Dimensions Can't Be 0.", vbInformation
        TBox(1).SetFocus
        Exit Sub
    ElseIf Val(TBox(2)) = 0 Then
        MsgBox "Dimensions Can't Be 0.", vbInformation
        TBox(2).SetFocus
        Exit Sub
    End If
    
    con.Execute "Update BoxSize Set Dimension1=" & Val(TBox(0)) & _
     ",Dimension2=" & Val(TBox(1)) & ",Dimension3=" & Val(TBox(2))
     
     
    Unload Me
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select * From BoxSize", con, adOpenForwardOnly, adLockReadOnly
        TBox(0) = ![Dimension1] & ""
        TBox(1) = ![Dimension2] & ""
        TBox(2) = ![Dimension3] & ""
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub
