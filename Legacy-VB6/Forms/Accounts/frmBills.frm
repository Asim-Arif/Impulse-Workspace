VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmDueRcptPymt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Due Receipts And Payments"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7140
   ScaleWidth      =   10590
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   9
      ImageHeight     =   5
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBills.frx":0000
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBills.frx":00E0
            Key             =   "Up"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10635
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Due Receipts And Payments"
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
         Index           =   0
         Left            =   60
         TabIndex        =   6
         Top             =   165
         Width           =   10605
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00AFBDCB&
         BackStyle       =   0  'Transparent
         Caption         =   "Due Receipts And Payments"
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
         Height          =   465
         Index           =   2
         Left            =   75
         TabIndex        =   7
         Top             =   180
         Width           =   10590
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   10335
      TabIndex        =   2
      Top             =   3795
      Width           =   10335
      Begin VB.Label lblRcpt 
         Alignment       =   2  'Center
         BackColor       =   &H00BECCD8&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Due Receipts From Customers"
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
         Left            =   135
         MousePointer    =   7  'Size N S
         TabIndex        =   3
         Top             =   0
         Width           =   10260
      End
   End
   Begin MSComctlLib.ListView LVPymt 
      Height          =   2805
      Left            =   0
      TabIndex        =   0
      Top             =   1020
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   4948
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   15199215
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Voucher #"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Vender"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   "Amt"
         Text            =   "Total Amout"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Amount Due"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Pending From"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSComctlLib.ListView LvRcpt 
      Height          =   2280
      Left            =   90
      TabIndex        =   8
      Top             =   4050
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   4022
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   15199215
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Voucher #"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Customer"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   "Amt"
         Text            =   "Total Amout"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Amount Due"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Pending From"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   390
      Left            =   8775
      TabIndex        =   4
      Top             =   6585
      Width           =   1560
      ForeColor       =   64
      Caption         =   "Close            "
      PicturePosition =   327683
      Size            =   "2752;688"
      Picture         =   "frmBills.frx":01C0
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label lblPymt 
      Alignment       =   2  'Center
      BackColor       =   &H00BECCD8&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Due Payments To Venders"
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
      Height          =   270
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   1
      Top             =   690
      Width           =   10335
   End
End
Attribute VB_Name = "frmDueRcptPymt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReadyToDrag As Boolean

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Call GetPymts
   Call GetRcpts
End Sub

Private Sub Form_Resize()

On Error Resume Next

lblPymt.Move (Me.ScaleWidth - lblPymt.Width) / 2
LVPymt.Move 50, lblPymt.Top + lblPymt.Height, Me.ScaleWidth - 100, Me.ScaleHeight * 3 / 7

cmdClose.Move Me.ScaleWidth - cmdClose.Width - 100, Me.ScaleHeight - cmdClose.Height - 100

pic.Move 50, LVPymt.Top + LVPymt.Height, LVPymt.Width

LvRcpt.Move 50, pic.Top + pic.Height, Me.ScaleWidth - 100, cmdClose.Top - (pic.Top + pic.Height) - 50

lblPymt.Move LVPymt.Left, lblPymt.Top, LVPymt.Width

Frame1.Move (Me.ScaleWidth - Frame1.Width) / 2

   
   
End Sub



Private Sub LVPymt_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)



If LVPymt.SortKey = ColumnHeader.Index - 1 Then
   If LVPymt.SortOrder = lvwAscending Then
      LVPymt.SortOrder = lvwDescending
      ColumnHeader.Icon = "Up"
   Else
      LVPymt.SortOrder = lvwAscending
      ColumnHeader.Icon = "Down"
   End If
Else
 LVPymt.ColumnHeaderIcons = Nothing
 Set LVPymt.ColumnHeaderIcons = ImageList1
 LVPymt.SortKey = ColumnHeader.Index - 1
 LVPymt.SortOrder = lvwAscending
 ColumnHeader.Icon = "Down"
End If


LVPymt.Sorted = True



End Sub
Private Sub LvRcpt_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)



If LvRcpt.SortKey = ColumnHeader.Index - 1 Then
   If LvRcpt.SortOrder = lvwAscending Then
      LvRcpt.SortOrder = lvwDescending
      ColumnHeader.Icon = "Up"
   Else
      LvRcpt.SortOrder = lvwAscending
      ColumnHeader.Icon = "Down"
   End If
Else
 LvRcpt.ColumnHeaderIcons = Nothing
 Set LvRcpt.ColumnHeaderIcons = ImageList1
 LvRcpt.SortKey = ColumnHeader.Index - 1
 LvRcpt.SortOrder = lvwAscending
 ColumnHeader.Icon = "Down"
End If


LvRcpt.Sorted = True



End Sub

Private Sub Pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReadyToDrag = True
    
    pic.ZOrder
End Sub

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim TheTop As Long
TheTop = pic.Top + Y

    If ReadyToDrag Then
        If (TheTop > (lblPymt.Top + lblPymt.Height + 500)) And (TheTop < cmdClose.Top - 800) Then
            pic.Top = TheTop
        End If
    End If

    
End Sub

Private Sub Pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If ReadyToDrag Then
    
       LVPymt.Height = pic.Top - LVPymt.Top
       LvRcpt.Move 50, pic.Top + pic.Height, LvRcpt.Width, Me.ScaleHeight - (Frame1.Height + lblPymt.Height + cmdClose.Height + LVPymt.Height + pic.Height + 250)
     
       ReadyToDrag = False
        
        pic.ZOrder 1
    End If
    
End Sub

Private Sub pic_Resize()
   lblRcpt.Move 0, 0, pic.Width, pic.Height
End Sub

Private Sub lblRcpt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Pic_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblRcpt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Pic_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblRcpt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call Pic_MouseUp(Button, Shift, X, Y)
End Sub



Private Sub GetPymts()
Dim rsList As New ADODB.Recordset
rsList.CursorLocation = adUseClient
rsList.Open "Select Sno,AccTitle,Vdate,VchrNo,Desc,TotalAmount,AmountLeft from AccPay Where  AccNo in (Select Accno From VVenders) AND  (Vdate<=#" & Date & "#) and ((AmountLeft>0) or (AmountLeft is null))", Con, adOpenForwardOnly, adLockReadOnly
 rsList.Sort = "Sno"
With LVPymt.ListItems
    .Clear
    'cmd.Caption = "Auto Apply         "
    Dim itm As ListItem
    Do Until rsList.EOF
        Set itm = .Add(, rsList![Sno] & "?", rsList![VchrNo])
        itm.ListSubItems.Add(, , rsList![AccTitle] & "").ToolTipText = rsList![AccTitle] & ""
        itm.ListSubItems.Add(, , rsList![Desc] & "").ToolTipText = rsList![Desc] & ""
        itm.SubItems(3) = rsList![TotalAmount] & ""
        If IsNull(rsList![AmountLeft]) Then
            itm.SubItems(4) = rsList![TotalAmount]
        Else
            itm.SubItems(4) = rsList![AmountLeft] & ""
        End If
        
        itm.SubItems(5) = IIf((Date - rsList![Vdate]) = 0, "Today.", Date - rsList![Vdate] & " Days.")
        
        
        rsList.MoveNext
    Loop
    'If .Count > 0 Then cmd.Enabled = True Else cmd.Enabled = False
End With

End Sub


Private Sub GetRcpts()
Dim rsList As New ADODB.Recordset
rsList.CursorLocation = adUseClient
rsList.Open "Select Sno,AccTitle,Vdate,VchrNo,Desc,TotalAmount,AmountLeft from AccRec Where Trim(AccNo) in (Select Accno From VCustomers) AND  (Vdate<=#" & Date & "#) and ((AmountLeft>0) or (AmountLeft is null))", Con, adOpenForwardOnly, adLockReadOnly
 rsList.Sort = "Sno"
With LvRcpt.ListItems
    .Clear
    'cmd.Caption = "Auto Apply         "
    Dim itm As ListItem
    Do Until rsList.EOF
        Set itm = .Add(, rsList![Sno] & "?", rsList![VchrNo])
        itm.ListSubItems.Add(, , rsList![AccTitle] & "").ToolTipText = rsList![AccTitle] & ""
        itm.ListSubItems.Add(, , rsList![Desc] & "").ToolTipText = rsList![Desc] & ""
        itm.SubItems(3) = rsList![TotalAmount] & ""
        If IsNull(rsList![AmountLeft]) Then
            itm.SubItems(4) = rsList![TotalAmount]
        Else
            itm.SubItems(4) = rsList![AmountLeft] & ""
        End If
        
        itm.SubItems(5) = IIf((Date - rsList![Vdate]) = 0, "Today.", Date - rsList![Vdate] & " Days.")
        
        
        rsList.MoveNext
    Loop
    'If .Count > 0 Then cmd.Enabled = True Else cmd.Enabled = False
End With

End Sub

