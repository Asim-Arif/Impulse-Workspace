VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmGenerateIPOLots 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generate IPO Lots"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtLotQty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3525
      TabIndex        =   5
      Top             =   1485
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   420
      Left            =   4140
      TabIndex        =   4
      Top             =   5250
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   420
      Left            =   5295
      TabIndex        =   3
      Top             =   5250
      Width           =   1095
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4635
      Left            =   90
      TabIndex        =   2
      Top             =   570
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   8176
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item Code"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Item Name"
         Object.Width           =   4586
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Lot Qty"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.TextBox txtIPONo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1155
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   165
      Width           =   1770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "IPO No. :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   375
      TabIndex        =   0
      Top             =   210
      Width           =   690
   End
End
Attribute VB_Name = "frmGenerateIPOLots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bReturnValue As Boolean
Dim lIPONo As Long

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Function ShowMe(p_lIPONo As Long) As Boolean

    lIPONo = p_lIPONo
    txtIPONo = p_lIPONo
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT T1.EntryID,CompItemCode,ProductionQty,ItemName,MaxLotSize FROM InternalProductionOrderDetail T1 INNER JOIN InternalProductionOrder " & _
         "ON T1.RefID=InternalProductionOrder.EntryID INNER JOIN Items ON T1.CompItemCode=Items.ItemID INNER JOIN " & _
         "ItemCatagories ON Items.CatID=ItemCatagories.CatID WHERE IPONo=" & lIPONo, con, adOpenForwardOnly, adLockReadOnly
         
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, !EntryID & "'", !CompItemCode & "")
            ITM.ListSubItems.add , , !ItemName & ""
            ITM.ListSubItems.add , , Val(!ProductionQty & "")
            ITM.ListSubItems.add , , Val(!MaxLotSize & "")
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    Me.Show 1
    
    ShowMe = bReturnValue
    
End Function
    

Private Sub SaveLots(lRefID As Long, lMaxLotQty As Long, lQty As Long)

    Dim lLots As Long, lExtraLotQty As Long
    lExtraLotQty = 0
    lLots = Floor(lQty / lMaxLotQty)
    If lQty > (lLots * lMaxLotQty) Then
        lExtraLotQty = lQty - (lLots * lMaxLotQty)
    End If
    
    Dim i As Integer
    For i = 1 To lLots
        con.Execute "INSERT INTO InternalProductionOrderDetailLots(IPOD_RefID,LotQty) VALUES(" & _
         lRefID & "," & lMaxLotQty & ")"
    Next
    
    If lExtraLotQty > 0 Then
        con.Execute "INSERT INTO InternalProductionOrderDetailLots(IPOD_RefID,LotQty) VALUES(" & _
         lRefID & "," & lExtraLotQty & ")"
    End If
    
End Sub


Private Sub cmdGenerate_Click()
    Dim i As Integer
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If Val(.SubItems(3)) = 0 Then
                MsgBox "No Lot Qty. for " & .Text, vbCritical
                Exit Sub
            End If
        End With
    Next
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            Call SaveLots(Val(.key), Val(.SubItems(3)), Val(.SubItems(2)))
        End With
    Next
    bReturnValue = True
    Unload Me
    
End Sub

Private Sub Form_Load()
    bReturnValue = False
    lIPONo = 0
End Sub

Private Sub LV_DblClick()
    If LV.ListItems.count = 0 Then Exit Sub
    With txtLotQty
        .Move LV.Left + LV.ColumnHeaders(4).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(4).Width
        .Text = Val(LV.SelectedItem.SubItems(3))
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub txtLotQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LV.SelectedItem.SubItems(3) = Val(txtLotQty)
        txtLotQty.Visible = False
    ElseIf KeyAscii = 27 Then
        txtLotQty.Visible = False
    End If
End Sub

Private Sub txtLotQty_LostFocus()
    txtLotQty.Visible = False
End Sub
