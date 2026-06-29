VERSION 5.00
Begin VB.Form frmAway 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   705
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   6030
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5625
      Top             =   0
   End
   Begin VB.Label lblAway 
      Caption         =   "Welcome Back To Export Documentation Package..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   780
      TabIndex        =   0
      Top             =   270
      Width           =   5115
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   150
      Picture         =   "frmAway.frx":0000
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmAway"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tCounter As Integer
Private Sub Form_Click()
    Call AtUnload
End Sub

Private Sub Image1_Click()
    Call AtUnload
End Sub

Private Sub Label1_Click()
    Call AtUnload
End Sub

Public Sub ShowForm()

    Dim HOURS As Long
    Dim MINS As Long
    Dim SECS As Long
       
    MINS = DateDiff("n", AwayTime, ServerTime)
    SECS = DateDiff("s", AwayTime, ServerTime)
    
    SECS = SECS - (MINS * 60)
    
    HOURS = MINS / 60
    
    MINS = MINS - (HOURS * 60)
    
    lblAway = "You Have Been Away For " & HOURS & " Hours " & MINS & " Minutes and " & SECS & " Seconds."
    
    tCounter = 0
    Timer1.Enabled = True
    
    Me.Show 1
    
    
End Sub

Private Sub Timer1_Timer()
    Call AtUnload
End Sub

Private Sub AtUnload()

'    For i = 1 To MainForm.Toolbar.count
'        MainForm.Toolbar(i).Refresh
'    Next i
    Unload Me
    
End Sub
