VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmAlertsDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alerts Details"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   450
      Left            =   8655
      TabIndex        =   7
      Top             =   7410
      Visible         =   0   'False
      Width           =   1260
   End
   Begin MSComctlLib.ListView LVPlanningDetail 
      Height          =   7155
      Left            =   0
      TabIndex        =   5
      Top             =   150
      Visible         =   0   'False
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   12621
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User Name"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Task Completed"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Task Completed Time"
         Object.Width           =   3175
      EndProperty
   End
   Begin VB.CommandButton cmdHideDetail 
      Caption         =   "&Hide Detail"
      Height          =   450
      Left            =   0
      TabIndex        =   6
      Top             =   7410
      Visible         =   0   'False
      Width           =   1260
   End
   Begin MSComctlLib.ListView LVPlanning 
      Height          =   7155
      Left            =   0
      TabIndex        =   4
      Top             =   150
      Visible         =   0   'False
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   12621
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date Plan"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Item Code"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Remarks"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Task Completed"
         Object.Width           =   2469
      EndProperty
   End
   Begin MSComctlLib.ListView LVEvents 
      Height          =   7155
      Left            =   0
      TabIndex        =   3
      Top             =   150
      Visible         =   0   'False
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   12621
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "S #"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Customer"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Event"
         Object.Width           =   15522
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Hidden"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Hidden"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   450
      Left            =   9990
      TabIndex        =   1
      Top             =   7410
      Width           =   1260
   End
   Begin MSComctlLib.ListView LVPurchaseAlerts 
      Height          =   7155
      Left            =   0
      TabIndex        =   2
      Top             =   150
      Visible         =   0   'False
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   12621
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "S #"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Alert"
         Object.Width           =   17639
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Hidden"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Hidden"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView LVLateOrders 
      Height          =   7155
      Left            =   0
      TabIndex        =   0
      Top             =   150
      Visible         =   0   'False
      Width           =   11250
      _ExtentX        =   19844
      _ExtentY        =   12621
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "S #"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Customer"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Order No."
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Delivery"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Late Days"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "# of Items"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
   End
End
Attribute VB_Name = "frmAlertsDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iAlertType As Integer
Dim strLateOrderSQL As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub ShowAlerts(m_iAlertType As Integer)
    iAlertType = m_iAlertType
    '0=Purchase Alerts;1=Late Orders;2=Customer Events;3=Initial Planning.
    cmdPrint.Visible = False
    Dim rs As New ADODB.Recordset, ITM As ListItem
    If iAlertType = 1 Then
        With rs
            strLateOrderSQL = "SELECT * FROM VFOrderList WHERE OrderNo IN(SELECT OrderNo FROM VUnShippedOrderList) AND DeliveryDT<'" & Date & "' AND CustCode NOT IN('STOCK','PARTS') AND LateOrderAlerts=1"
            .Open strLateOrderSQL, con, adOpenForwardOnly, adLockReadOnly
            LVLateOrders.ListItems.Clear
            Do Until .EOF
                Set ITM = LVLateOrders.ListItems.add(, , LVLateOrders.ListItems.count + 1)
                ITM.ListSubItems.add , , !CustCode & ""
                ITM.ListSubItems.add(, , IIf(strCompany = "IAA" Or strCompany = "Kami", !InternalRefNo, !OrderNo & "")).Tag = !OrderNo & ""
                ITM.ListSubItems.add , , Format(!DeliveryDT & "", "dd-MMM-yyyy")
                ITM.ListSubItems.add , , DateDiff("d", !DeliveryDT, Date)
                ITM.ListSubItems.add , , Val(!TotalArticles & "") - Val(!ShippedItemCount & "")
                ITM.ListSubItems.add , , Val(!TotalBalanceQty & "")
                .MoveNEXT
            Loop
            .Close
        End With
        LVLateOrders.Visible = True
        cmdPrint.Visible = True
    ElseIf iAlertType = 0 Then
   
        Dim strAlert As String, dBalance As Double, lForeColor As Long, i As Integer
        With rs
            .Open "SELECT EntryIDReturnDTs,ItemName,Makers.VenderName,ItemCode,IssQty,RcvQty,Makers.VendID1,Prefix,FollowUpData.EntryID AS FollowUpEntryID,MasterPONO FROM " & _
            "VPurchaseCalendar INNER JOIN Items ON VPurchaseCalendar.ItemCode=Items.ItemID INNER JOIN Makers ON " & _
            "VPurchaseCalendar.VendID=Makers.VendID LEFT OUTER JOIN FollowUpData ON EntryIDReturnDTs=VIDRDTS_RefID " & _
            "AND Prefix=FollowUpData.FollowUp WHERE ReturnDT='" & Date & "' AND IssQty>RcvQty " & _
            "AND CustCode<>'Stock' AND MasterPONo IS NOT NULL AND Makers.VendID<>207 AND CustCode IN(SELECT CustCode FROM Users_Customers " & _
            "WHERE UserID=" & UserID & ")", con, adOpenForwardOnly, adLockReadOnly
        
            LVPurchaseAlerts.ListItems.Clear
        
            Do Until .EOF
        
                Set ITM = LVPurchaseAlerts.ListItems.add(, , LVPurchaseAlerts.ListItems.count + 1)
                ITM.Tag = Val(!EntryIDReturnDTs & "")
                
                strAlert = ""
                If !Prefix <> "" Then
                    strAlert = !Prefix & ""
                End If
                dBalance = Val(!IssQty & "") - Val(!RcvQty & "")
                
                strAlert = strAlert & " " & dBalance & " Pcs of " & !ItemCode & " (" & !ItemName & ") from " & !VendID1 & " (" & !VenderName & ")"
                ITM.ListSubItems.add(, , strAlert).Tag = !Prefix & ""
                ITM.ListSubItems.add(, , "").Tag = Val(!FollowUpEntryID & "")
                ITM.ListSubItems.add(, , "").Tag = !MasterPONo & ""
                
                If Val(!FollowUpEntryID & "") = 0 Then
                    If !Prefix & "" = "(D)" Then
                        lForeColor = RGB(50, 50, 255)
                    Else
                        lForeColor = RGB(255, 50, 50)
                    End If
                Else
                    lForeColor = RGB(100, 200, 100)
                End If
                ITM.ForeColor = lForeColor
                For i = 1 To ITM.ListSubItems.count
                    ITM.ListSubItems(i).ForeColor = lForeColor
                Next
                .MoveNEXT
            Loop
            .Close
        End With
        LVPurchaseAlerts.Visible = True
    ElseIf iAlertType = 2 Then
    
        With rs
            .Open "SELECT CustCode,Description FROM ForeignCustomers_Events WHERE EventDT='" & Date & "' AND RepeatedType=0" & _
                "UNION ALL SELECT CustCode,Description FROM ForeignCustomers_Events WHERE CONVERT(Varchar(50),EventDT,11)='" & Format(Date, "yy/mm") & "' AND RepeatedType=1 " & _
                "UNION ALL SELECT Description AS CustCode,CONVERT(VARCHAR(50),DT,6) AS Description FROM VCompanyCertificationExpiries WHERE DT<='" & Date & "'", con, adOpenForwardOnly, adLockReadOnly
            LVEvents.ListItems.Clear
            Do Until .EOF
                Set ITM = LVEvents.ListItems.add(, , LVEvents.ListItems.count + 1)
                ITM.ListSubItems.add , , !CustCode & ""
                ITM.ListSubItems.add , , !Description & ""
                .MoveNEXT
            Loop
            .Close
        End With
        LVEvents.Visible = True
        
    ElseIf iAlertType = 3 Then  'Initial Planning....
        Dim strAcknowledge As String
        With rs
            .Open "SELECT DT,GeneralRemarks,CompItemCode,FOrderItems_InitialPlanningMemo_Users.* FROM FOrderItems_InitialPlanningMemo_Users INNER JOIN FOrderItems_InitialPlanningMemo ON FOrderItems_InitialPlanningMemo_Users.FOI_IPM_RefID=FOrderItems_InitialPlanningMemo.EntryID INNER JOIN FOrderItems ON FOI_RefID=FOrderItems.ID WHERE FOrderItems_InitialPlanningMemo_Users.UserName='" & CurrentUserName & "' AND Acknowledged=0", con, adOpenForwardOnly, adLockReadOnly
            LVPlanning.ListItems.Clear
            Do Until .EOF
                Set ITM = LVPlanning.ListItems.add(, , Format(!DT, "dd-MMM-yyyy"))
                ITM.Tag = Val(!FOI_IPM_RefID & "")
                ITM.ListSubItems.add , , !CompItemCode & ""
                ITM.ListSubItems.add , , !GeneralRemarks & ""
                If !Acknowledged Then
                    strAcknowledge = "Yes"
                Else
                    strAcknowledge = "No"
                End If
                ITM.ListSubItems.add , , strAcknowledge
                .MoveNEXT
            Loop
            .Close
        End With
        LVPlanning.Visible = True
    End If
    
    Set rs = Nothing
    If iAlertType = 3 Then
        Me.Show , MainForm
    Else
        Me.Show 1
    End If
    
End Sub
    
Private Sub c_Click()
    
End Sub

Private Sub cmdHideDetail_Click()
    LVPlanningDetail.Visible = False
    cmdHideDetail.Visible = False
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    Me.MousePointer = vbHourglass
    
    Dim f As New frmPrevRpt
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptFOrderList.rpt")
        
    
    rpt.FormulaFields.GetItemByName("ForCustomer").Text = "'<Late Orders>'"
    rpt.FormulaFields.GetItemByName("ForFromTo").Text = "''"
    rpt.FormulaFields.GetItemByName("OrderType").Text = "''"
    
    Dim frm As New frmPrevRpt
    Me.Hide
    frm.ShowReport " ", rpt, strLateOrderSQL, , , Me
    
    Me.MousePointer = vbNormal
    
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub LVLateOrders_DblClick()

    If strCompany <> "IAA" Then Exit Sub
    
    If LVLateOrders.ListItems.count = 0 Then Exit Sub
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptProductionOrder_Status.rpt")
    rpt.ParameterFields.GetItemByName("@OrderNo").SetCurrentValue LVLateOrders.SelectedItem.ListSubItems(2).Tag
    rpt.FormulaFields.GetItemByName("UserName").Text = "'" & CurrentUserName & "'"
    
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    Me.Hide
    f.ShowReport "", rpt, , False, , Me
    
End Sub

Private Sub LVPlanning_DblClick()

    If LVPlanning.ListItems.count = 0 Then Exit Sub
    If LVPlanning.SelectedItem.SubItems(3) = "No" Then
        con.Execute "UPDATE FOrderItems_InitialPlanningMemo_Users SET Acknowledged=1,AcknowledgedEntryDT=getDate() WHERE FOI_IPM_RefID=" & Val(LVPlanning.SelectedItem.Tag) & " AND UserName='" & CurrentUserName & "'"
        LVPlanning.SelectedItem.SubItems(3) = "Yes"
    End If
    
End Sub

Private Sub LVPlanning_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If LVPlanning.ListItems.count = 0 Then Exit Sub
    If Button <> 2 Then Exit Sub
    Dim strRet As String
    strRet = ShowPopUpMenu1("Show Detail", "Show Dashboard", "Initial Planning")
    
    If strRet = "Show Detail" Then
        Dim rs As New ADODB.Recordset, ITM As ListItem, strAcknowledge As String
        With rs
            .Open "SELECT * FROM FOrderItems_InitialPlanningMemo_Users WHERE FOI_IPM_RefID=" & Val(LVPlanning.SelectedItem.Tag), con, adOpenForwardOnly, adLockReadOnly
            LVPlanningDetail.ListItems.Clear
            Do Until .EOF
                Set ITM = LVPlanningDetail.ListItems.add(, , !UserName & "")
                If !Acknowledged Then
                    strAcknowledge = "Yes"
                Else
                    strAcknowledge = "No"
                End If
                ITM.ListSubItems.add , , strAcknowledge
                If IsNull(!AcknowledgedEntryDT) Then
                    ITM.ListSubItems.add , , "-"
                Else
                    ITM.ListSubItems.add , , Format(!AcknowledgedEntryDT, "dd-MMM-yyyy HH:nn")
                End If
                .MoveNEXT
            Loop
            .Close
        End With
        Set rs = Nothing
        LVPlanningDetail.Visible = True
        cmdHideDetail.Visible = True
    ElseIf strRet = "Show Dashboard" Then
        Dim f As New frmProductionPlanningDashBoard
        Load f
        f.ShowMe (LVPlanning.SelectedItem.SubItems(1))
    ElseIf strRet = "Initial Planning" Then
        Dim lFOI_RefID As Long
        lFOI_RefID = GetSingleLongValue("FOI_RefID", "FOrderItems_InitialPlanningMemo", " WHERE EntryID=" & Val(LVPlanning.SelectedItem.Tag))
        Call ShowInitialPlanningMemo(lFOI_RefID, 0, 0, 0, 0, 0)
    End If
    
End Sub

