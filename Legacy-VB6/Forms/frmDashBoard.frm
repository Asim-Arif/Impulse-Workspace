VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDashBoard 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Delete Voucher"
   ClientHeight    =   10395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDashBoard.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10395
   ScaleWidth      =   11340
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   4545
      Index           =   1
      Left            =   180
      ScaleHeight     =   4545
      ScaleWidth      =   10920
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "RECEIVING"
      Top             =   5190
      Width           =   10920
      Begin MSComctlLib.ListView LVRcving 
         Height          =   4200
         Left            =   30
         TabIndex        =   6
         Top             =   315
         Width           =   10860
         _ExtentX        =   19156
         _ExtentY        =   7408
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   65280
         BackColor       =   4210752
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Lot No."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Process"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Customer Order"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Qty"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Maker"
            Object.Width           =   4586
         EndProperty
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   165
      TabIndex        =   3
      Top             =   225
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   108658691
      CurrentDate     =   40509
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9885
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9870
      Width           =   1215
   End
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   4545
      Index           =   0
      Left            =   180
      ScaleHeight     =   4545
      ScaleWidth      =   10920
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "ISSUANCE"
      Top             =   570
      Width           =   10920
      Begin MSComctlLib.ListView LVIssuance 
         Height          =   4200
         Left            =   30
         TabIndex        =   4
         Top             =   315
         Width           =   10860
         _ExtentX        =   19156
         _ExtentY        =   7408
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   65280
         BackColor       =   4210752
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Lot No."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Process"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Customer Order"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Qty"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Maker"
            Object.Width           =   4586
         EndProperty
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DASH BOARD"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   10920
   End
End
Attribute VB_Name = "frmDashBoard"
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
    Dim lBorderColor As Long, lForeColor As Long
    For i = 0 To PicWC.count - 1
        If i = 1 Then
            lBorderColor = RGB(153, 204, 255)
            lForeColor = 0
        ElseIf i = 0 Then
            lBorderColor = RGB(255, 204, 153)
            lForeColor = 0
        ElseIf i = 3 Then
            lBorderColor = RGB(255, 153, 204)
            lForeColor = 0
        ElseIf i = 2 Then
            lBorderColor = RGB(0, 255, 0)
            lForeColor = 0
        ElseIf i = 4 Then
            lBorderColor = RGB(200, 200, 192)
            lForeColor = 0
        End If
        Call DrawCaptions(PicWC(i), lBorderColor, lForeColor)
    Next
    
End Sub

Private Sub DrawCaptions(Pic As PictureBox, lBorderColor As Long, lForeColor As Long)

    Dim lFillColor
    'lBorderColor = 0 'RGB(226, 103, 161) 'RGB(70, 70, 181)
    lFillColor = lBorderColor 'RGB(226, 103, 161) 'RGB(70, 70, 181)
    Pic.Line (0, 20)-(Pic.Width, 20), lBorderColor
    
    If Pic.Tag <> "" Then
        Pic.Line (0, 150)-(Pic.Width, 150), lBorderColor
        
        Pic.Line (0, 50)-(Pic.Width, 300), lFillColor, BF
        Pic.ForeColor = lForeColor 'vbGreen
        Pic.CurrentX = (Pic.Width - Pic.TextWidth(Pic.Tag)) / 2
        Pic.CurrentY = 0
        Pic.Print Pic.Tag
    End If
        
    'Pic.Line (0, Pic.Height - 30)-(Pic.Width, Pic.Height - 30), lBorderColor 'Bottom
    Pic.Line (0, Pic.Height - 50)-(Pic.Width, Pic.Height), lFillColor, BF
    
    Pic.Line (20, 0)-(20, Pic.Height), lBorderColor                          'Left
    Pic.Line (Pic.Width - 20, 0)-(Pic.Width - 30, Pic.Height), lBorderColor  'Right
    
    'Pic.Line (0, Pic.Height - 500)-(50, Pic.Height), lBorderColor
'    Me.BackColor = lBorderColor
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub DTPicker1_Change()
    Call UpdateData
End Sub

Private Sub Form_Load()

    Call DrawPicCaptions
    bMouseDown = False
    bSaved = False
    lEntryID = 0
    
    Call UpdateData
    
End Sub

Private Sub PicWC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If y < 600 Then
        If Button = 1 Then
            If bMouseDown = False Then
                bMouseDown = True
                lDownX = x
                lDownY = y
            End If
        End If
    End If
End Sub

Private Sub PicWC_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If bMouseDown Then
        Me.Move Me.Left + ((x - lDownX) / 15), Me.Top + ((y - lDownY) / 15)
    End If
End Sub

Private Sub PicWC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
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

    Me.Show
    ShowMe = bSaved
    
End Function

Private Sub Timer1_Timer()
    UpdateData
End Sub

Public Sub UpdateData()

    If LVIssuance.Tag <> DTPicker1.Value Then
        LVIssuance.ListItems.Clear
        LVRcving.ListItems.Clear
    End If
    Dim strIDs As String
    strIDs = "0"
    Dim i As Integer
    For i = 1 To LVIssuance.ListItems.count
        With LVIssuance.ListItems(i)
            strIDs = strIDs & "," & Val(.key)
        End With
    Next
    
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT VendIssdDetail.EntryID,OrderNo,ItemCode,LotNo,Description,IssQty,Makers.VendID1,VenderName FROM VendIssdDetail " & _
         "INNER JOIN Processes ON VendIssdDetail.RcvProcessID=Processes.ProcessID " & _
         "INNER JOIN VendIssued ON VendIssdDetail.RefID=VendIssued.EntryID " & _
         "INNER JOIN Makers ON VendIssued.VendID=Makers.VendID " & _
         "WHERE DT='" & DTPicker1.Value & "' AND VendIssdDetail.EntryID NOT IN(" & strIDs & ") ORDER BY VendIssdDetail.EntryID", con, adOpenForwardOnly, adLockReadOnly
         
        If .EOF = False Then
            Call PlayIssuanceSound
        End If
        
        Do Until .EOF
            Set ITM = LVIssuance.ListItems.add(1, !EntryID & "'", !LotNo & "")
            ITM.ListSubItems.add , , !Description & ""
            ITM.ListSubItems.add , , !ItemCode & ""
            ITM.ListSubItems.add , , !OrderNo & ""
            ITM.ListSubItems.add , , Val(!IssQty & "")
            ITM.ListSubItems.add , , "(" & !VendID1 & ") " & !VenderName & ""
            .MoveNEXT
        Loop
        .Close
        If LVIssuance.ListItems.count > 0 Then
            LVIssuance.ListItems(1).EnsureVisible
        End If
        LVIssuance.Tag = DTPicker1.Value
        
        
        strIDs = "0"
        For i = 1 To LVRcving.ListItems.count
            With LVRcving.ListItems(i)
                strIDs = strIDs & "," & Val(.key)
            End With
        Next
        
        
        .Open "SELECT VendRcvdDetail.EntryID,OrderNo,ItemCode,LotNo,Description,RcvdQty,Makers.VendID1,VenderName FROM VendRcvdDetail " & _
         "INNER JOIN Processes ON VendRcvdDetail.ProcessID=Processes.ProcessID " & _
         "INNER JOIN VendReceived ON VendRcvdDetail.RefID=VendReceived.EntryID " & _
         "INNER JOIN Makers ON VendReceived.VendID=Makers.VendID " & _
         "WHERE CAST(CONVERT(Varchar(50),DT,6) AS DateTime)='" & DTPicker1.Value & "' AND VendRcvdDetail.EntryID NOT IN(" & strIDs & ") ORDER BY VendRcvdDetail.EntryID", con, adOpenForwardOnly, adLockReadOnly
         
        If .EOF = False Then
            Call PlayReceivingSound
        End If
        Do Until .EOF
            Set ITM = LVRcving.ListItems.add(1, !EntryID & "'", !LotNo & "")
            ITM.ListSubItems.add , , !Description & ""
            ITM.ListSubItems.add , , !ItemCode & ""
            ITM.ListSubItems.add , , !OrderNo & ""
            ITM.ListSubItems.add , , Val(!RcvdQty & "")
            ITM.ListSubItems.add , , "(" & !VendID1 & ") " & !VenderName & ""
            .MoveNEXT
        Loop
        .Close
        If LVRcving.ListItems.count > 0 Then
            LVRcving.ListItems(1).EnsureVisible
        End If
    End With
    Set rs = Nothing
    
End Sub
