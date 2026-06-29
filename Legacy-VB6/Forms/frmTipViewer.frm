VERSION 5.00
Begin VB.Form frmTipViewer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tip"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9345
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrev 
      Caption         =   "&Prev"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3375
      TabIndex        =   7
      Top             =   525
      Width           =   1350
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4800
      TabIndex        =   6
      Top             =   525
      Width           =   1350
   End
   Begin VB.CheckBox chkShowAtStartup 
      Caption         =   "Show at Startup"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   255
      TabIndex        =   5
      Top             =   5640
      Width           =   3660
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   7890
      TabIndex        =   4
      Top             =   5565
      Width           =   1350
   End
   Begin VB.TextBox txtTipDetail 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3060
      Left            =   255
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2460
      Width           =   5175
   End
   Begin VB.TextBox txtTip 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   585
      Left            =   255
      TabIndex        =   1
      Top             =   1530
      Width           =   5175
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3960
      Left            =   5505
      Stretch         =   -1  'True
      Top             =   1530
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "[Showing]"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2250
      TabIndex        =   8
      Top             =   135
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Detail"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   255
      TabIndex        =   2
      Top             =   2145
      Width           =   5175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Tip"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   255
      TabIndex        =   0
      Top             =   1155
      Width           =   5175
   End
End
Attribute VB_Name = "frmTipViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lTipIndex As Long

Private Sub chkShowAtStartup_Click()
    con.Execute "UPDATE Users SET ShowTips=" & Abs(chkShowAtStartup.Value) & " WHERE UserName='" & CurrentUserName & "'"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    Call ShowTip(1)
End Sub

Private Sub cmdPrev_Click()
    Call ShowTip(0)
End Sub

Private Sub Form_Load()

    lTipIndex = GetSingleLongValue("Tip_RefID", "Users", " WHERE UserName='" & CurrentUserName & "'")
    
    Dim bShowTips As Boolean
    bShowTips = GetSingleBooleanValue("ShowTips", "Users", " WHERE UserName='" & CurrentUserName & "'")
    
    chkShowAtStartup.Value = Abs(bShowTips)
    
    Call ShowTip(1)
    
End Sub

Private Sub ShowTip(iMovement As Integer)
    
    Dim rs As New ADODB.Recordset, lIndex As Long
    If iMovement = 1 Then 'Move Next
        lIndex = GetSingleLongValue("MIN(EntryID)", "Tips", " WHERE EntryID>" & lTipIndex)
    ElseIf iMovement = 0 Then 'Move Back.
        lIndex = GetSingleLongValue("MAX(EntryID)", "Tips", " WHERE EntryID<" & lTipIndex)
    End If
     
    If lIndex = 0 Then
        lIndex = lTipIndex
    End If
    
    With rs
        .Open "SELECT * FROM VTips WHERE EntryID=" & lIndex, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            txtTip = !TipName & ""
            txtTipDetail = !TipDetail & ""
            
            If Not IsNull(![TipPic]) Then
                Set Image1.DataSource = rs
                Image1.DataField = "TipPic"
            Else
                Image1.Picture = LoadPicture("")
            End If
        End If
        .Close
    End With
    Set rs = Nothing
    
    con.Execute "UPDATE Users SET Tip_RefID=" & lIndex & " WHERE UserName='" & CurrentUserName & "'"
    
End Sub
