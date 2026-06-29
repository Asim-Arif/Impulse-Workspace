VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmWait 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4620
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.Animation Animation1 
      Height          =   705
      Left            =   3870
      TabIndex        =   1
      Top             =   90
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   1244
      _Version        =   393216
      Center          =   -1  'True
      BackStyle       =   1
      BackColor       =   0
      FullWidth       =   43
      FullHeight      =   47
   End
   Begin MSComCtl2.Animation Animation2 
      Height          =   195
      Left            =   525
      TabIndex        =   2
      Top             =   600
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   344
      _Version        =   393216
      Center          =   -1  'True
      BackStyle       =   1
      BackColor       =   16777215
      FullWidth       =   195
      FullHeight      =   13
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      Height          =   840
      Left            =   30
      Top             =   30
      Width           =   4575
   End
   Begin MSForms.Label Label1 
      Height          =   390
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   3555
      ForeColor       =   8388608
      BackColor       =   16777215
      VariousPropertyBits=   276824091
      Caption         =   "Please Wait While System Processes Data Reading Data From BioMetric"
      Size            =   "6271;688"
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Animation1_Click()
    Unload Me
End Sub

Private Sub Form_Deactivate()
'    Me.SetFocus
    Me.Refresh
End Sub

Private Sub Form_GotFocus()
    Me.Refresh
End Sub

Private Sub Form_Load()
    Animation1.Open App.Path & "\avtar.avi"
    Animation1.Play
    '
    Animation2.Open App.Path & "\progress.avi"
    Animation2.Play
End Sub

Private Sub Label1_Click()
    Unload Me
End Sub
