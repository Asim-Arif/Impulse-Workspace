VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeleteVoucher 
   BorderStyle     =   0  'None
   Caption         =   "Delete Voucher"
   ClientHeight    =   7905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9225
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDeleteVoucher.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FAF5FE&
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
      Height          =   7905
      Index           =   0
      Left            =   0
      ScaleHeight     =   7905
      ScaleWidth      =   9225
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "Delete Voucher"
      Top             =   0
      Width           =   9225
      Begin VB.TextBox txtVchrNo 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   6735
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1290
         Width           =   2070
      End
      Begin VB.TextBox txtAccTitle 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1290
         Width           =   3765
      End
      Begin VB.TextBox txtAccNo 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   345
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1290
         Width           =   2550
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FAF5FE&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6885
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   7365
         Width           =   1905
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FAF5FE&
         Caption         =   "Delete It !"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4905
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   7380
         Width           =   1935
      End
      Begin VB.TextBox txtComputer 
         Height          =   315
         Left            =   6885
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   6540
         Width           =   1905
      End
      Begin VB.TextBox txtUserName 
         Height          =   315
         Left            =   4905
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   6555
         Width           =   1935
      End
      Begin VB.TextBox txtReason 
         Height          =   1215
         Left            =   345
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   6570
         Width           =   4500
      End
      Begin VB.TextBox txtDescription 
         Height          =   315
         Left            =   2070
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2055
         Width           =   6735
      End
      Begin VB.TextBox txtDate 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   345
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   2055
         Width           =   1680
      End
      Begin MSComctlLib.ListView LV 
         Height          =   3690
         Left            =   345
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2430
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   6509
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         PictureAlignment=   5
         _Version        =   393217
         ForeColor       =   7621
         BackColor       =   15199215
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Acc #"
            Object.Width           =   2910
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Acc Title"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Debit"
            Object.Width           =   1993
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Credit"
            Object.Width           =   2028
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   15
         Picture         =   "frmDeleteVoucher.frx":014A
         Top             =   60
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Computer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00001DC5&
         Height          =   195
         Index           =   7
         Left            =   6870
         TabIndex        =   13
         Top             =   6300
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00001DC5&
         Height          =   195
         Index           =   6
         Left            =   4920
         TabIndex        =   11
         Top             =   6315
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reason of Deletion"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00001DC5&
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   9
         Top             =   6330
         Width           =   1590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vchr #"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00001DC5&
         Height          =   195
         Index           =   4
         Left            =   6780
         TabIndex        =   8
         Top             =   1065
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acc Title"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00001DC5&
         Height          =   195
         Index           =   3
         Left            =   2955
         TabIndex        =   7
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acc No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00001DC5&
         Height          =   195
         Index           =   2
         Left            =   345
         TabIndex        =   6
         Top             =   1065
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00001DC5&
         Height          =   195
         Index           =   1
         Left            =   2085
         TabIndex        =   3
         Top             =   1830
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00001DC5&
         Height          =   195
         Index           =   0
         Left            =   375
         TabIndex        =   1
         Top             =   1815
         Width           =   405
      End
   End
End
Attribute VB_Name = "frmDeleteVoucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const lCaptionHeight = 600
Dim bMouseDown As Boolean
Dim lDownX As Long, lDownY As Long

Private Sub DrawBorder(Pic As PictureBox)

    Dim lBorderColor As Long
    lBorderColor = &H8080FF      'RGB(69, 140, 201) 'RGB(69, 140, 201)
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
    lBorderColor = &H8080FF      'RGB(51, 102, 153)
    lFillColor = &H8080FF      'RGB(51, 102, 153)
    Pic.Line (0, 20)-(Pic.Width, 20), lBorderColor
    
    If Pic.Tag <> "" Then
        Pic.Line (0, 350)-(Pic.Width, 350), lBorderColor
        
        Pic.Line (0, 50)-(Pic.Width, 600), lFillColor, BF
        Pic.ForeColor = vbWhite
        Pic.CurrentX = (Pic.Width - Pic.TextWidth(Pic.Tag)) / 2
        Pic.CurrentY = 180
        Pic.Print Pic.Tag
    End If
        
    Pic.Line (0, Pic.Height - 20)-(Pic.Width, Pic.Height - 20), lBorderColor 'Bottom

    Pic.Line (20, 0)-(20, Pic.Height), lBorderColor    'Left
    Pic.Line (Pic.Width - 20, 0)-(Pic.Width - 20, Pic.Height), lBorderColor  'Right

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    
    If txtReason.Text = "" Then
        MsgBox "Enter Reason of Deletion.", vbInformation
        txtReason.SetFocus
        Exit Sub
    End If
    
    Dim getPassword As Boolean
    getPassword = getDBPassword("DeleteVoucher")
    
    If getPassword Then
        If MsgBox("Are You Sure To Delete Voucher # : " & txtVchrNo.Text, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        'If Not getPassword Then Exit Sub
        
        If DeleteVoucher(txtVchrNo.Text, , txtReason.Text) Then MsgBox "Voucher Deleted Successfully.", vbInformation
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    
    Call DrawPicCaptions
    bMouseDown = False
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
    If bMouseDown Then
        Me.Move Me.Left + ((X - lDownX) / 15), Me.Top + ((Y - lDownY) / 15)
    End If
End Sub

Private Sub PicWC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    bMouseDown = False
End Sub

Public Sub ShowMe(VchrNo As String)
    
    Dim rs As New ADODB.Recordset
    Dim AccNo As String, AccTitle As String
    With rs
        .Open "Select AccNo,AccTitle From VVouchers Where VchrNo='" & VchrNo & "'", con, adOpenForwardOnly, adLockReadOnly
        AccNo = !AccNo & ""
        AccTitle = !AccTitle & ""
        .Close
    End With
    Set rs = Nothing
    txtAccNo = AccNo
    txtAccTitle = AccTitle
    txtVchrNo = VchrNo
    Me.Show 1
    
End Sub

Private Sub txtVchrNo_Change()
    
    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "Select * From VVouchers Where VchrNo='" & txtVchrNo.Text & "'", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            txtDate.Text = Format(!Vdate, "dd-MMM-yyyy")
            txtDescription.Text = !Description & ""
            Do Until .EOF
                Set ITM = LV.ListItems.add(, , !AccNo & "")
                ITM.SubItems(1) = !AccTitle & ""
                ITM.SubItems(2) = Val(!Debit & "")
                ITM.SubItems(3) = Val(!Credit & "")
                .MoveNEXT
            Loop
        End If
        .Close
        .Open "Select * From VoucherInfo Where VchrNo='" & txtVchrNo & "'", con, adOpenForwardOnly, adLockReadOnly
        txtUserName = !UserName & ""
        txtComputer = !MachineName & ""
        .Close
    End With
    Set rs = Nothing
    Exit Sub
    
err:
    MsgBox err.Description
End Sub
