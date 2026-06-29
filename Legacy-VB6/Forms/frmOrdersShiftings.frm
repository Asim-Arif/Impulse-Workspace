VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOrdersShiftings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Order to Order Shifting..."
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
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
   ScaleHeight     =   6525
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTotal 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   7485
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   5610
      Width           =   1200
   End
   Begin VB.TextBox txtEdit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4215
      TabIndex        =   11
      Top             =   2370
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   360
      Left            =   6090
      TabIndex        =   10
      Top             =   6105
      Width           =   1260
   End
   Begin VB.TextBox txtOrderQty 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   7485
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   810
      Width           =   1200
   End
   Begin VB.TextBox txtProcess 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1215
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   435
      Width           =   7470
   End
   Begin VB.TextBox txtItemCode 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1215
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   60
      Width           =   7470
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   360
      Left            =   7425
      TabIndex        =   1
      Top             =   6105
      Width           =   1260
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4350
      Left            =   60
      TabIndex        =   0
      Top             =   1230
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   7673
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "S #"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PTC #"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Order #"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Shift"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total :"
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
      Index           =   4
      Left            =   6900
      TabIndex        =   12
      Top             =   5685
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Order Qty :"
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
      Index           =   3
      Left            =   6510
      TabIndex        =   8
      Top             =   885
      Width           =   915
   End
   Begin MSForms.ComboBox cmbOrderNo 
      Height          =   360
      Left            =   1215
      TabIndex        =   7
      Top             =   810
      Width           =   5250
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "9260;635"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Order No. :"
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
      Index           =   2
      Left            =   300
      TabIndex        =   6
      Top             =   870
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Process :"
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
      Index           =   1
      Left            =   420
      TabIndex        =   4
      Top             =   480
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Item Code :"
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
      Index           =   0
      Left            =   195
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmOrdersShiftings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strItemCode As String, iProcessNo As Integer
Dim bReturnValue As Boolean

Public Function ShowMe(p_strItemCode As String, p_iProcessNo As Integer) As Boolean
    strItemCode = p_strItemCode
    iProcessNo = p_iProcessNo
    
    txtProcess = GetSingleStringValue("Description", "Processes", " WHERE ProcessID=" & iProcessNo)
    txtItemCode = strItemCode
    
    Call AddToCombo(cmbOrderNo, "OrderNo", "FOrderItems", " WHERE CompItemCode='" & strItemCode & "'", , "Qty")
    
    Me.Show 1
    ShowMe = bReturnValue
    
End Function

Private Sub cmbOrderNo_Change()
    Call cmbOrderNo_Click
End Sub

Private Sub cmbOrderNo_Click()

    If cmbOrderNo.MatchFound Then
        txtOrderQty = Val(cmbOrderNo.List(cmbOrderNo.ListIndex, 1))
    Else
        txtOrderQty = ""
        LV.ListItems.Clear
        Exit Sub
    End If
    
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT * FROM VendRcvdDetail WHERE ItemCode='" & strItemCode & "' AND ReqAuth=0 AND RcvdQty>IssQty AND OrderNo<>'" & cmbOrderNo.Text & "' ORDER BY LotNo", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, !EntryID & "'", LV.ListItems.count + 1)
            ITM.ListSubItems.add , , Val(!LotNo & "")
            ITM.ListSubItems.add , , !OrderNo & ""
            ITM.ListSubItems.add , , Val(!RcvdQty & "") - Val(!IssQty & "")
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    
    Dim i As Integer, lEntryID As Long
    Call StartTrans(con)
    
    con.Execute "INSERT INTO OrdersShiftings(OrderNo,ItemCode,ProcessID,UserName,MachineName) VALUES('" & cmbOrderNo.Text & _
    "','" & strItemCode & "'," & iProcessNo & ",'" & CurrentUserName & "','" & strComputerName & "')"
    
    lEntryID = GetSingleLongValue("MAX(EntryID)", "OrdersShiftings")
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                con.Execute "INSERT INTO OrdersShiftingsDetails(OS_RefID,VRD_RefID,OriginalOrderNo) VALUES(" & lEntryID & "," & Val(.key) & _
                 ",'" & .SubItems(2) & "')"
                
                con.Execute "UPDATE VendRcvdDetail SET OrderNo='" & cmbOrderNo.Text & "' WHERE EntryID=" & Val(.key)
            End If
        End With
    Next
    
    con.CommitTrans
    
    bReturnValue = True
    Unload Me
    
End Sub

Private Sub Form_Load()
    bReturnValue = False
End Sub

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        Item.SubItems(4) = Val(Item.SubItems(3))
    Else
        Item.SubItems(4) = ""
    End If
    
    Dim i As Integer, lTotal As Long
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            lTotal = lTotal + Val(LV.ListItems(i).SubItems(4))
        End If
    Next
    txtTotal = lTotal
End Sub
