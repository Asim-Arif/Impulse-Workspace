VERSION 5.00
Begin VB.Form frmProductionLinks 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   10005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6570
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProductionLinks.frx":0000
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   10005
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00EFE0D6&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   10000
      Index           =   0
      Left            =   0
      ScaleHeight     =   10005
      ScaleWidth      =   6570
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "PRODUCTION LINKS"
      Top             =   0
      Width           =   6570
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receiving"
         BeginProperty Font 
            Name            =   "Segoe Condensed"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3465
         TabIndex        =   4
         Top             =   615
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issuance"
         BeginProperty Font 
            Name            =   "Segoe Condensed"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   300
         TabIndex        =   3
         Top             =   615
         Width           =   780
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   3180
         X2              =   3180
         Y1              =   885
         Y2              =   9800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   3165
         X2              =   3165
         Y1              =   885
         Y2              =   9800
      End
      Begin VB.Label lblRcv 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rcv Label"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   0
         Left            =   3465
         MouseIcon       =   "frmProductionLinks.frx":014A
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   1020
         Width           =   885
      End
      Begin VB.Label lblIss 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Iss Label"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   0
         Left            =   285
         MouseIcon       =   "frmProductionLinks.frx":029C
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   1020
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmProductionLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bMouseDown As Boolean
Const lCaptionHeight = 600
Dim lDownX As Long, lDownY As Long
Dim iColNo As Integer
Dim iColNoTemp As Integer
Dim bSaved As Boolean
Dim lEntryID As Long

Private Sub DrawBorder(Pic As PictureBox)

    Dim lBorderColor As Long
    lBorderColor = RGB(0, 200, 0)    'RGB(69, 140, 201) 'RGB(69, 140, 201)
    Pic.Line (0, 0)-(Pic.Width, 0), lBorderColor 'Top
    Pic.Line (0, Pic.Height - 85)-(Pic.Width, Pic.Height - 85), lBorderColor 'Bottom
    Pic.Line (0, 0)-(0, Pic.Height), lBorderColor    'Left
    Pic.Line (Pic.Width - 85, 0)-(Pic.Width - 85, Pic.Height), lBorderColor  'Right
    
End Sub

Private Sub DrawPicCaptions()

    Dim i As Integer
    For i = 0 To PicWC.count - 1
        Call DrawCaptions(PicWC(i))
    Next
    
End Sub

Private Sub DrawCaptions(Pic As PictureBox)

    Dim lBorderColor As Long, lFillColor
    lBorderColor = RGB(70, 70, 181)
    lFillColor = RGB(70, 70, 181)
    Pic.Line (0, 20)-(Pic.Width, 20), lBorderColor
    
    If Pic.Tag <> "" Then
        Pic.Line (0, 350)-(Pic.Width, 350), lBorderColor
        
        Pic.Line (0, 50)-(Pic.Width, 600), lFillColor, BF
        Pic.ForeColor = vbWhite
        Pic.CurrentX = (Pic.Width - Pic.TextWidth(Pic.Tag)) / 2
        Pic.CurrentY = 180
        Pic.Print Pic.Tag
    End If
        
    'Pic.Line (0, Pic.Height - 30)-(Pic.Width, Pic.Height - 30), lBorderColor 'Bottom
    Pic.Line (0, Pic.Height - 100)-(Pic.Width, Pic.Height), lFillColor, BF
    
    Pic.Line (20, 0)-(20, Pic.Height), lBorderColor                          'Left
    Pic.Line (Pic.Width - 30, 0)-(Pic.Width - 30, Pic.Height), lBorderColor  'Right
    
    'Pic.Line (0, Pic.Height - 500)-(50, Pic.Height), lBorderColor
'    Me.BackColor = lBorderColor
End Sub




Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub Form_Load()

    Call DrawPicCaptions
    bMouseDown = False
    bSaved = False
    
End Sub

Private Sub lblIss_Click(Index As Integer)
    ShowIssForm (lblIss(Index).Tag)
End Sub

Private Sub lblRcv_Click(Index As Integer)
    ShowRcvVForm (lblRcv(Index).Tag)
End Sub

Private Sub PicWC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y < 600 Then
        If Button = 1 Then
            If bMouseDown = False Then
                bMouseDown = True
                lDownX = X
                lDownY = Y
            End If
        End If
    End If
End Sub

Private Sub PicWC_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Exit Sub    'Restring Moving
    If bMouseDown Then
        Me.Move Me.Left + ((X - lDownX) / 15), Me.Top + ((Y - lDownY) / 15)
    End If
End Sub

Private Sub PicWC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    bMouseDown = False
End Sub

Public Function ShowMe() As Boolean
    
    'Will hold the handle to the new region.
    Dim hRegion As Long
    '
    'This API can create a Round or Oval Shaped Region.
    hRegion = CreateRoundRectRgn(0, 0, Me.Width / 15, Me.Height / 15, 20, 20)
    '
    'Change the shape of the Form to the newly created Region.
    SetWindowRgn Me.hwnd, hRegion, True
    '
    'Show the new form with the specified shape.
    'frmRegion.Show

   
    'Me.Move 5200 + 2000, 4600
    Me.Move 3200 + 2000, 200
    
    Call AddLinks
    Me.Show
    ShowMe = True
    
End Function

Private Sub AddLinks()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim i As Integer, iNewIndex As Integer
    i = 0
    iNewIndex = 0
    With rs
        .Open "SELECT * FROM Processes", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            If i > 0 Then
                iNewIndex = lblIss.UBound + 1 'i + 1
                Load lblIss(iNewIndex)
                lblIss(iNewIndex).Top = lblIss(iNewIndex - 1).Top + lblIss(iNewIndex - 1).Height + 100
                lblIss(iNewIndex).Visible = True
                '----------------------------------
                Load lblRcv(iNewIndex)
                lblRcv(iNewIndex).Top = lblRcv(iNewIndex - 1).Top + lblRcv(iNewIndex - 1).Height + 100
                lblRcv(iNewIndex).Visible = True
            End If
            lblIss(iNewIndex).Caption = !Description & ""
            lblIss(iNewIndex).Tag = Val(!ProcessID & "")
            '-------------------------------------------
            lblRcv(iNewIndex).Caption = !Description & ""
            lblRcv(iNewIndex).Tag = Val(!ProcessID & "")
            i = i + 1
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
End Sub
