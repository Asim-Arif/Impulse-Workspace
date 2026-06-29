VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmArticlewiseWIP 
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleMode       =   0  'User
   ScaleWidth      =   10721.69
   WindowState     =   2  'Maximized
   Begin VB.Frame fraPODetail 
      Caption         =   "Qty Ordered to Maker"
      Height          =   6480
      Left            =   165
      TabIndex        =   6
      Top             =   1140
      Visible         =   0   'False
      Width           =   11730
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   270
         Left            =   105
         TabIndex        =   8
         Top             =   6075
         Width           =   915
      End
      Begin MSComctlLib.ListView LVOrd 
         Height          =   5775
         Left            =   105
         TabIndex        =   7
         Top             =   240
         Width           =   11550
         _ExtentX        =   20373
         _ExtentY        =   10186
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   15199215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sr. No."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "PO No."
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Internal #"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Maker"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "PO #"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Qty Ordered"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Received"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Receivable"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin VB.Frame FAcc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   8400
      Left            =   60
      TabIndex        =   1
      Top             =   -45
      Width           =   11865
      Begin ComboList.Usercontrol1 cmbArticle 
         Height          =   285
         Left            =   195
         TabIndex        =   2
         Top             =   345
         Width           =   6660
         _ExtentX        =   11748
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
      Begin MSComctlLib.ListView LV 
         Height          =   7110
         Left            =   45
         TabIndex        =   3
         Top             =   765
         Width           =   11700
         _ExtentX        =   20638
         _ExtentY        =   12541
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   15199215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sr. No."
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Processes"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Qty Ordered"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Received"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Issued"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "In-Hand"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Receivable"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComCtl2.DTPicker DtTo 
         Height          =   315
         Left            =   8625
         TabIndex        =   10
         Top             =   330
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   162332675
         CurrentDate     =   37081
      End
      Begin MSComCtl2.DTPicker DtFrom 
         Height          =   315
         Left            =   6900
         TabIndex        =   11
         Top             =   330
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   162332675
         CurrentDate     =   37055
      End
      Begin VB.Label lbl 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   8610
         TabIndex        =   12
         Top             =   105
         Width           =   1155
      End
      Begin VB.Label lbl 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   6900
         TabIndex        =   9
         Top             =   135
         Width           =   1155
      End
      Begin VB.Label lbl 
         Caption         =   "Article"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   5
         Top             =   150
         Width           =   1155
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   375
         Left            =   10440
         TabIndex        =   4
         Top             =   7935
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "2355;661"
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdRefresh 
         Height          =   360
         Left            =   10260
         TabIndex        =   0
         Top             =   315
         Width           =   1500
         Caption         =   "Refresh        "
         PicturePosition =   327683
         Size            =   "2646;635"
         Accelerator     =   82
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "PopMenu"
   End
End
Attribute VB_Name = "frmArticlewiseWIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iColNo As Integer

Private Sub cmbArticle_matched()

    On Error GoTo err
    
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "ArticlewiseWIP_SP"
        .Parameters("@ItemID").Value = cmbArticle.ID
        .Parameters("@DTFrom").Value = DtFrom.Value
        .Parameters("@DTTo").Value = DtTo.Value
        Set rs = .Execute
    End With
    
    LV.ListItems.Clear
    With rs
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            ITM.Tag = Val(!ProcessID & "")
            ITM.SubItems(1) = !Description & ""
            ITM.SubItems(2) = Val(!IssQty & "")
            ITM.SubItems(3) = Val(!RcvdQty & "")
            ITM.SubItems(4) = Val(!ReIssQty & "")   'Issd to Next Process
            ITM.SubItems(5) = Val(!RcvdQty & "") - Val(!ReIssQty & "")
            ITM.SubItems(6) = Val(!IssQty & "") - Val(!RcvdQty & "")
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
    fraPODetail.Visible = False
    FAcc.Enabled = True
End Sub

Private Sub cmdRefresh_Click()
    Call cmbArticle_matched
End Sub

Private Sub Form_Load()
    iColNo = 0
    mnuPop.Visible = False
    cmbArticle.ListHeight = 5800
    cmbArticle.DropDownOffLine = True
    Call cmbArticle.AddVals(con, "'{ ' + ItemID + ' } ' + ItemName ", "Items", "ItemID", " WHERE InActive=0")
    
    DtFrom.Value = DateSerial(year(Date), Month(Date), 1)
    DtTo.Value = Date
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    FAcc.Top = (Me.ScaleHeight - FAcc.Height) / 2
    
    
    fraPODetail.Left = (Me.ScaleWidth - fraPODetail.Width) / 2
    fraPODetail.Top = (Me.ScaleHeight - fraPODetail.Height) / 2
End Sub

Private Sub LV_DblClick()
    If iColNo = 3 Then
        Call ShowIssuanceDetail
    End If
End Sub

Private Sub LV_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i As Integer
    For i = 1 To LV.ColumnHeaders.count
        If X > LV.ColumnHeaders(1).Left And X < LV.ColumnHeaders(i).Left + LV.ColumnHeaders(i).Width Then
            iColNo = i
            Exit For
        End If
    Next
    
End Sub

Private Sub ShowIssuanceDetail()

    On Error GoTo err
    
    Dim cmd As New ADODB.Command
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "ArticlewiseWIP_SP"
        .Parameters("@ItemID").Value = cmbArticle.ID
        .Parameters("@Options").Value = 2
        .Parameters("@ProcessID").Value = LV.SelectedItem.Tag
        .Parameters("@DTFrom").Value = DtFrom.Value
        .Parameters("@DTTo").Value = DtTo.Value
        Set rs = .Execute
    End With
    
    LVOrd.ListItems.Clear
    Dim strLastOrderNo As String
    strLastOrderNo = ""
    Dim lIssdQty As Long, lIssdQtyTotal As Long
    Dim lRcvdQty As Long, lRcvdQtyTotal As Long
    Dim lRcvableQty As Long, lRcvableQtyTotal As Long
    lIssdQty = 0
    lRcvdQty = 0
    lRcvableQty = 0
    Dim iIndex As Integer
    iIndex = 0
    Dim i As Integer
    With rs
        Do Until .EOF
            If !OrderNo <> strLastOrderNo Then
                Set ITM = LVOrd.ListItems.add(, !OrderNo & "'", iIndex + 1)
                iIndex = iIndex + 1
'                ITM.ListSubItems.add(, , "").Bold = True
                ITM.ListSubItems.add(, , !OrderNo & "").Bold = True
                ITM.ListSubItems.add(, , !InternalRefNo & "").Bold = True
                ITM.ListSubItems.add(, , "-----").Bold = True
                ITM.ListSubItems.add(, , "-----").Bold = True
                ITM.ListSubItems.add(, , "").Bold = True
                ITM.ListSubItems.add(, , "").Bold = True
                ITM.ForeColor = vbRed
                For i = 1 To ITM.ListSubItems.count
                    ITM.ListSubItems(i).ForeColor = vbRed
                Next
            End If
            
            Set ITM = LVOrd.ListItems.add()
            ITM.ListSubItems.add , , "-----"
            ITM.ListSubItems.add , , "-----"
            ITM.ListSubItems.add , , !VendID1 & ""
            
            ITM.Tag = !OrderNo & ""
            ITM.ListSubItems.add , , !RecieptID & ""
            ITM.ListSubItems.add , , Val(!IssQty & "")
            ITM.ListSubItems.add , , Val(!RcvdQty & "")
            ITM.ListSubItems.add , , Val(!IssQty & "") - Val(!RcvdQty & "")
            
            If strLastOrderNo <> !OrderNo & "" And strLastOrderNo <> "" Then
                Set ITM = LVOrd.ListItems(strLastOrderNo & "'")
                ITM.SubItems(5) = lIssdQty
                ITM.SubItems(6) = lRcvdQty
                ITM.SubItems(7) = lRcvableQty
                lIssdQty = 0
                lRcvdQty = 0
                lRcvableQty = 0
            End If
            lIssdQty = lIssdQty + Val(!IssQty & "")
            lRcvdQty = lRcvdQty + Val(!RcvdQty & "")
            lRcvableQty = lRcvableQty + (Val(!IssQty & "") - Val(!RcvdQty & ""))
            
            lIssdQtyTotal = lIssdQtyTotal + Val(!IssQty & "")
            lRcvdQtyTotal = lRcvdQtyTotal + Val(!RcvdQty & "")
            lRcvableQtyTotal = lRcvableQtyTotal + (Val(!IssQty & "") - Val(!RcvdQty & ""))
            
            strLastOrderNo = !OrderNo & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    If strLastOrderNo <> "" Then
        Set ITM = LVOrd.ListItems(strLastOrderNo & "'")
        ITM.SubItems(5) = lIssdQty
        ITM.SubItems(6) = lRcvdQty
        ITM.SubItems(7) = lRcvableQty
    End If
    Set ITM = LVOrd.ListItems.add()
    ITM.ListSubItems.add(, , "Total").Bold = True
    ITM.ListSubItems.add(, , "-----").Bold = True
    ITM.ListSubItems.add(, , "-----").Bold = True
    ITM.ListSubItems.add(, , "-----").Bold = True
    ITM.ListSubItems.add(, , lIssdQtyTotal).Bold = True
    ITM.ListSubItems.add(, , lRcvdQtyTotal).Bold = True
    ITM.ListSubItems.add(, , lRcvableQtyTotal).Bold = True
    FAcc.Enabled = False
    
    fraPODetail.Visible = True
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button <> 2 Then Exit Sub
    
    If LV.ListItems.count = 0 Then Exit Sub
    If Val(LV.SelectedItem.SubItems(5)) <= 0 Then Exit Sub
    'Now If It's Not Zero...
    Dim strRet As String
    strRet = ShowPopUpMenu1("Order to Order Shifting")
    
    If strRet = "Order to Order Shifting" Then
        Dim f As New frmOrdersShiftings
        Load f
        If f.ShowMe(cmbArticle.ID, Val(LV.SelectedItem.Tag)) Then
            cmdRefresh_Click
        End If
    End If
    
End Sub

