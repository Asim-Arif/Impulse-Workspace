VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmReserveForggingOpenPO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Forging Reservation"
   ClientHeight    =   8685
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReserveForggingOpenPO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtArticle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   165
      TabIndex        =   9
      Top             =   270
      Width           =   6750
   End
   Begin ComboList.Usercontrol1 cmbItemID 
      Height          =   285
      Left            =   165
      TabIndex        =   6
      Top             =   1620
      Visible         =   0   'False
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextBackColor   =   12640511
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
      Appearance      =   0
   End
   Begin ComboList.Usercontrol1 Usercontrol12 
      Height          =   285
      Left            =   6945
      TabIndex        =   8
      Top             =   270
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextBackColor   =   12640511
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdUnSelectAll 
      Caption         =   "Uncheck All"
      Height          =   375
      Left            =   1455
      TabIndex        =   4
      Top             =   8220
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Check All"
      Height          =   375
      Left            =   135
      TabIndex        =   3
      Top             =   8220
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ListView LV 
      Height          =   7200
      Left            =   150
      TabIndex        =   2
      Top             =   615
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   12700
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Raw Material"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Qty In Hand"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Open PO Qty"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Total"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Reserve"
         Object.Width           =   2117
      EndProperty
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   450
      Left            =   8355
      TabIndex        =   1
      Top             =   8145
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Reserve"
      Height          =   450
      Left            =   7110
      TabIndex        =   0
      Top             =   8145
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Order No."
      Height          =   195
      Index           =   2
      Left            =   6960
      TabIndex        =   7
      Top             =   75
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Article"
      Height          =   195
      Index           =   1
      Left            =   165
      TabIndex        =   5
      Top             =   60
      Width           =   450
   End
End
Attribute VB_Name = "frmReserveForggingOpenPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean
Dim strItemID As String

Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub

Private Sub cmdSelectAll_Click()
    Dim i As Integer
    LV.ListItems(1).Checked = False
    For i = 2 To LV.ListItems.count
        LV.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdUnSelectAll_Click()
    Dim i As Integer
    
    For i = 1 To LV.ListItems.count
        LV.ListItems(1).Checked = False
    Next
End Sub

Private Sub Form_Load()

      
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "SELECT DeptID,Name FROM Departments ORDER BY Name", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Set ITM = LV.ListItems.add(, , "< All >")
        ITM.Tag = "0"
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , !DeptID & "")
            ITM.SubItems(1) = !Name & ""
            ITM.Tag = !DeptID & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    LV.ListItems(1).Checked = True
    
End Sub


Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    Dim i As Integer
    If Item.Tag = "0" Then
        For i = 2 To LV.ListItems.count
            LV.ListItems(i).Checked = False
        Next
    Else
        LV.ListItems(1).Checked = False
    End If
    
End Sub


Public Function ReserveForgging(p_strItemID As String, p_strItemDescription As String) As Boolean
     
    Me.Show 1
    
    strItemID = p_strItemID
    txtArticle = p_strItemDescription
    
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT VRM.RMName,QtyInStock,QtyToRcv,VItemsRMComp.RMID FROM VItemsRMComp INNER JOIN VRM ON VRM.RMID=VItemsRMComp.RMID LEFT OUTER JOIN " & _
         "(SELECT RMID1,SUM(QtyToRcv) AS QtyToRcv FROM VVendOrdersToRcv GROUP BY RMID1) T1 ON VRM.RMID1=T1.RMID1 " & _
         "LEFT OUTER JOIN VForggingStockReservations ON VRM.ID=ForggingStockReservations.RMID WHERE ItemID='" & _
         strItemID & "' AND GroupID IN(SELECT Group_ID FROM RMGroupIDsForForging)", con, adOpenForwardOnly, adLockReadOnly
        
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , !RMName & "")
            ITM.Tag = Val(!RMID & "")
            ITM.ListSubItems.add , , Val(!QtyInStock & "")
            ITM.ListSubItems.add , , Val(!QtyToRcv & "")
            ITM.ListSubItems.add , , Val(!QtyToRcv & "") + Val(!QtyInStock & "")
            .MoveNEXT
        Loop
        .Close
    End With
    
    ReserveForgging = Not bCancelled
    
End Function
