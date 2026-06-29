VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmProdTimeLineDashBoardDetails 
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   45
   ClientWidth     =   13860
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
   ScaleHeight     =   8835
   ScaleMode       =   0  'User
   ScaleWidth      =   12414.59
   WindowState     =   2  'Maximized
   Begin VB.Frame FAcc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   8925
      Left            =   60
      TabIndex        =   0
      Top             =   -90
      Width           =   13755
      Begin VB.Frame fTitle 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   0
         TabIndex        =   8
         Top             =   -15
         Width           =   13740
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Production Timeline Dashboard"
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
            Index           =   2
            Left            =   0
            TabIndex        =   9
            Top             =   135
            Width           =   13710
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Production Timeline Dashboard"
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
            Height          =   390
            Index           =   3
            Left            =   0
            TabIndex        =   10
            Top             =   150
            Width           =   13710
         End
      End
      Begin VB.Frame Frame1 
         Height          =   480
         Left            =   60
         TabIndex        =   4
         Top             =   8400
         Width           =   5475
         Begin VB.ComboBox cmbfield 
            Height          =   315
            ItemData        =   "frmProdTimeLineDashBoardDetails.frx":0000
            Left            =   60
            List            =   "frmProdTimeLineDashBoardDetails.frx":0019
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   135
            Width           =   1875
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4395
            TabIndex        =   6
            Top             =   135
            Width           =   1020
         End
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   1965
            TabIndex        =   5
            Top             =   135
            Width           =   2430
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   7830
         Left            =   60
         TabIndex        =   1
         Top             =   570
         Width           =   13605
         _ExtentX        =   23998
         _ExtentY        =   13811
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S #"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Customer Code"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Order #"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Item Code"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Description"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Qty"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Lot #"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSForms.CommandButton cmdPurchaseCalendar 
         Height          =   375
         Left            =   10305
         TabIndex        =   11
         Top             =   8460
         Visible         =   0   'False
         Width           =   1290
         ForeColor       =   0
         Caption         =   "Calendar"
         PicturePosition =   327683
         Size            =   "2275;661"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdAuth 
         Height          =   375
         Left            =   11655
         TabIndex        =   3
         Top             =   8460
         Visible         =   0   'False
         Width           =   1035
         ForeColor       =   0
         Caption         =   "Authorize"
         PicturePosition =   327683
         Size            =   "1826;661"
         Accelerator     =   65
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   375
         Left            =   12750
         TabIndex        =   2
         Top             =   8460
         Width           =   915
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "1614;661"
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   11
      ImageHeight     =   10
      MaskColor       =   255
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProdTimeLineDashBoardDetails.frx":0060
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProdTimeLineDashBoardDetails.frx":021C
            Key             =   "Up"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmProdTimeLineDashBoardDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bCourierInvoice As Boolean
Dim bShipping As Boolean
Dim bCustomInvoice As Boolean
Dim strReportSQL As String

Public Sub ShowMe(Optional p_bCustomInvoice As Boolean = True)
    bCustomInvoice = p_bCustomInvoice
    Me.Show
End Sub

Public Sub ShowForShipping()
    bShipping = True
    Me.Show
End Sub

Private Sub chkMasterPOOpen_Click()
    If chkMasterPOOpen.Value = vbChecked Then
        chkShowMasterPOOnly.Value = vbUnchecked
        chkRepairLots.Value = vbUnchecked
    End If
End Sub

Private Sub chkRegularLotsOnly_Click()
    If chkRegularLotsOnly.Value = vbChecked Then
        chkRepairLots.Value = vbUnchecked
    End If
End Sub

Private Sub chkRepairLots_Click()

    If chkRepairLots.Value = vbChecked Then
        chkShowMasterPOOnly.Value = vbUnchecked
        chkMasterPOOpen.Value = vbUnchecked
        chkRegularLotsOnly.Value = vbUnchecked
    End If
    
End Sub

Private Sub chkShowMasterPOOnly_Click()

    If chkShowMasterPOOnly.Value = vbChecked Then
        chkMasterPOOpen.Value = vbUnchecked
        chkRepairLots.Value = vbUnchecked
    End If
    
End Sub

Private Sub cmbDueDays_Change()
    Call cmbDueDays_Click
End Sub

Private Sub cmbDueDays_Click()

    txtDueDays.Locked = True
    If cmbDueDays.ListIndex = 0 Then
        txtDueDays.Text = ""
    ElseIf cmbDueDays.ListIndex = 1 Then
        txtDueDays.Text = "3"
    ElseIf cmbDueDays.ListIndex = 2 Then
        txtDueDays.Text = "7"
    Else
        txtDueDays.Text = ""
        txtDueDays.Locked = False
    End If
    
End Sub

Private Sub cmbRange_Click()

    DtTo = GetServerDate
    DtFrom.Enabled = False
    DtTo.Enabled = False
    Select Case cmbRange.ListIndex
        Case 0
            DtFrom = Date
        Case 1
            DtFrom = DateAdd("d", -15, Date)
        Case 2
            DtFrom = DateAdd("d", -30, Date)
        Case 3
            DtFrom = DateAdd("d", -60, Date)
        Case 4
            DtFrom = DateAdd("d", -90, Date)
        Case 5
            DtFrom.Enabled = True
            DtTo.Enabled = True
            If DtFrom.Visible Then
                DtFrom.SetFocus
            End If
            Exit Sub
    End Select
    
End Sub


Private Sub cmdCheckAll_Click(Index As Integer)
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Checked = (Index = 0)
    Next
End Sub

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdFind_Click()

    Dim StartIdx As Integer

    If LV.SelectedItem Is Nothing Then
       StartIdx = 1
    Else
       If LV.SelectedItem.Index = LV.ListItems.count Then
          StartIdx = 1
       Else
          StartIdx = LV.SelectedItem.Index
       End If
    End If

    Dim fldValue As String




    For i = StartIdx To LV.ListItems.count
       
       
       If cmbfield.ListIndex = 0 Then
          fldValue = Trim(Replace(LV.ListItems(i), "»", ""))
       Else
          fldValue = Trim(Replace(LV.ListItems(i).SubItems(cmbfield.ListIndex), "»", ""))
       End If
    
       If LCase(fldValue) Like "*" & LCase(Trim(txtFind)) & "*" Then
          LV.ListItems(i).Selected = True
          LV.ListItems(i).EnsureVisible
          If i <> StartIdx Then Exit For
       End If
    Next
    
    If i = LV.ListItems.count + 1 Then
        MsgBox "No Mathes Found.", vbInformation
        Set LV.SelectedItem = Nothing
    End If

End Sub

Private Sub cmdPurchaseCalendar_Click()

    Dim rpt As CRAXDDRT.Report, strSelectionMain As String
    Set rpt = rptApp.OpenReport(RptPath & "\PurchaseCalendar.rpt")
    Dim f As New frmPrevRpt
    strSelectionMain = "{DateLookup.DateFull}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#"
    
    rpt.FormulaFields.GetItemByName("Maker").Text = cmbMaker.ID
    rpt.FormulaFields.GetItemByName("ItemGroup").Text = cmbItemGroup.ID
    rpt.FormulaFields.GetItemByName("Customer").Text = "'" & cmbCustomer.ID & "'"
    
    f.ShowReport strSelectionMain, rpt
    
End Sub


Private Sub DtFrom_Change()
    'Call setDateRange(DtFrom)
End Sub

Private Sub DtTo_Change()
    'Call setDateRange(DtTo)
End Sub


Private Sub Form_Load()
        
    With cmbfield
        .Clear
        For i = 1 To LV.ColumnHeaders.count
            cmbfield.AddItem LV.ColumnHeaders(i).Text
        Next
    End With
     
End Sub

Private Sub RefreshLV(lIssProcessID As Long, lRcvProcessID As Long, myDT As Date)
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim Cond As String
    Dim ITM As ListItem
    
    Screen.MousePointer = vbHourglass
   
    Dim lForeColor As Long
    
    With rs
        
        .Open "SELECT * FROM VVendRcvdDetail_Simple WHERE ProcessID=" & lIssProcessID & " AND DT='" & myDT & "' AND LotNo NOT IN(SELECT LotNo FROM VVendRcvdDetail_Simple WHERE ProcessID=" & lRcvProcessID & " AND DT='" & myDT & "')", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        'LV.Sorted = False
        Do Until .EOF
            If IsNull(!ReWorklot) Then
                lForeColor = LV.ForeColor
            ElseIf !ReWorklot Then
                lForeColor = lForeColor_RepairLot
'            ElseIf Val(!BookmarkEntryID & "") > 0 Then
'                lForeColor = RGB(255, 0, 255)
'            ElseIf !Closed Then
'                lForeColor = RGB(59, 59, 59)
            Else
                lForeColor = LV.ForeColor
            End If
            
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            ITM.Tag = Val(![EntryID] & "")
            
            ITM.ListSubItems.add , , !CustCode & ""
            ITM.ListSubItems.add , , !InternalRefNo & ""
            ITM.ListSubItems.add , , ![ItemCode] & ""
            ITM.ListSubItems.add , , !Description & ""
            ITM.ListSubItems.add , , Format(!DT & "", "dd-MMM-yyyy")
            ITM.ListSubItems.add , , Val(!RcvdQty & "")
            ITM.ListSubItems.add(, , ![LotNo] & "").Tag = ""
           
            ITM.ForeColor = lForeColor
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ForeColor = lForeColor
            Next
            
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    FAcc.Left = (Me.ScaleWidth - FAcc.Width) / 2
    FAcc.Top = (Me.ScaleHeight - FAcc.Height) / 2
    
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    If LV.SortKey + 1 = ColumnHeader.Index Then
        If LV.SortOrder = lvwAscending Then
            LV.SortOrder = lvwDescending
            ColumnHeader.Icon = "Down"
        Else
            LV.SortOrder = lvwAscending
            ColumnHeader.Icon = "Up"
        End If
    Else
        LV.ColumnHeaderIcons = Nothing
        Set LV.ColumnHeaderIcons = ImageList2
        LV.SortKey = ColumnHeader.Index - 1
        LV.SortOrder = lvwAscending
        ColumnHeader.Icon = "Down"
    End If
    Call RefreshSnos
    'If LV.Sorted = False Then LV.Sorted = True
End Sub

Private Sub ShowMenusForShipping()

    On Error Resume Next
    Dim c As Control
    For Each c In Controls
        If TypeOf c Is Menu Then
            If c.Name <> "mnuPrintPList" Then
                c.Visible = False
            End If
        End If
    Next
    
    On Error GoTo 0
    
    If LV.ListItems.count > 0 Then
        mnuPrintPList.Visible = True
        mnuCustomPList.Visible = True
        mnuCustomPListOther.Visible = True
        mnuCommercialPList.Visible = True
        mnuCommercialPListOther.Visible = True
        mnuDashPackingLabels.Visible = True
        mnuPrintPackingLabels.Visible = True
    End If
'    mnuPrintDecDocs.Visible = False
    Me.PopUpMenu mnuPop
    
End Sub

Private Sub ShowMenus(vis As Boolean)
            
    If LV.ListItems.count = 0 Then
        mnuEditDatesDash.Visible = False
        mnuEditPromises.Visible = False
    Else
        If LV.SelectedItem.ListSubItems(9).Tag <> "" Then
            mnuEditDatesDash.Visible = True
            mnuEditPromises.Visible = True
        Else
            mnuEditDatesDash.Visible = False
            mnuEditPromises.Visible = False
        End If
    End If
    
    If strCompany = "QEL" Then
        mnuPrintGatePass.Visible = True
    Else
        mnuPrintGatePass.Visible = False
    End If
    
    If strCompany <> "IAA" Then
        mnuAddFollowupEntries.Visible = False
    Else
        If LV.ListItems.count > 0 Then
            If LV.SelectedItem.SubItems(10) <> "" Then
                mnuAddFollowupEntries.Visible = True
            Else
                mnuAddFollowupEntries.Visible = False
            End If
        End If
    End If
    Me.PopUpMenu mnuPop
    
End Sub

Private Sub ShowMenusforCourier()
         
    On Error Resume Next
    Dim c As Control
    For Each c In Controls
        If TypeOf c Is Menu Then
            c.Visible = False
        End If
    Next
    
    On Error GoTo 0
    mnuAdd.Visible = True
    If LV.ListItems.count > 0 Then
        dd.Visible = True
        mnuEdit.Visible = True
        mnuComDocs.Visible = True
        mnuPrintComInv.Visible = True
    End If
    mnuPrintDecDocs.Visible = False
    Me.PopUpMenu mnuPop
    
End Sub


Private Sub mnuAdd_Click()
    'Load frmNewCustomInvoice
    'frmNewCustomInvoice.Show
    'ShowForm frmNewCustomInvoice, MainForm, False, "New Custom Invoice"
    'frmNewCustomInvoice.AddNewProforma
    Load frmNewCustomInvoice
    frmNewCustomInvoice.AddNewInvoice True, bCourierInvoice
End Sub

Private Sub mnuAddShippInfo_Click()

    Load frmEditShippInfo
    'ShowForm f, MainForm, , "Shipp. Info"
    Call frmEditShippInfo.EditEntry(LV.SelectedItem.SubItems(3))
    
End Sub

Private Sub mnuBillOfExchange_Click()
    Call PrintBillOfExchange
End Sub

Private Sub PrintBillOfExchange(Optional Firstrpt As Boolean = True)

    On Error GoTo err
    Me.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptBillOfExchange.rpt")
    rpt.FormulaFields(3).Text = Firstrpt
    Dim frm As New frmPrevRpt
    frm.ShowReport "{CustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    Me.MousePointer = vbNormal
    
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description

End Sub

    
Private Sub mnuBillOfExchange2_Click()
    Call PrintBillOfExchange(False)
End Sub

Private Sub mnuCommercialPList_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3), False, True, LV.SelectedItem.SubItems(1) = "DECATHLON", LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2))
End Sub

Private Sub mnuCommercialPListOther_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3), True, True, LV.SelectedItem.SubItems(1) = "DECATHLON", LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2))
End Sub

Private Sub mnuCoveringLetter_Click()
    
    On Error GoTo err
    Me.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptComCoveringLetter.rpt")
    If LCase(CurrentUserName) = "mnaeem" Then
        rpt.FormulaFields(3).Text = "'Naeem Ahmed'"
    End If
    
    Dim frm As New frmPrevRpt
    frm.ShowReport "{CustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub mnuCustomPList_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3))
End Sub

Private Sub mnuCustomPListOther_Click()
    'Call PrintPackingList(LV.SelectedItem.SubItems(3), True)
End Sub

Private Sub mnuDelPackingInfo_Click()
    
    If LV.ListItems.count = 0 Then Exit Sub
    Dim iAffect As Integer
    If MsgBox("Are you sure to Delete Packing List of Invoice # " & LV.SelectedItem.SubItems(3) & " ?", vbQuestion + vbYesNo) = vbYes Then
        Call StartTrans(con)
        con.Execute "Delete From CustomPList Where CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'", iAffect
        con.CommitTrans
        If iAffect > 0 Then
            LV.SelectedItem.ListSubItems(2).Tag = 0
        End If
    End If
    
End Sub


Private Sub mnuEForm_Click()

    Load frmPrintFormE
    If LV.ListItems.count > 0 Then
        frmPrintFormE.ShowForm LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
    Else
        frmPrintFormE.Show , MainForm
    End If
    
End Sub

Private Sub mnuGSP_Click()

'    Load frmGSP
'    If LV.ListItems.Count > 0 Then
'        frmGSP.ShowForm LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
'    Else
'        frmGSP.Show , MainForm
'    End If
    
End Sub

Private Sub mnuPacking_Click()

'    If Left(mnuPacking.Caption, 3) = "Add" Then
'
'        Load frmNewPList
'        frmNewPList.ShowMe LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
'    Else
'        Load frmNewPList
'        Call frmNewPList.EditPList(LV.SelectedItem.SubItems(3))
'        ShowForm frmNewPList, MainForm, , "Packing List"
'    End If
    
End Sub

Private Sub mnuPrintSummary_Click()

    On Error GoTo err
    Me.MousePointer = vbHourglass
    
    Call StartTrans(con)
    con.Execute "Delete From PrintFOrderList"
    
    For i = 1 To LV.ListItems.count
        con.Execute "Insert Into PrintFOrderList(Sno,Customer,Country," & _
         "OrderNo,OrderDT,OrderAmt,Status,Curr) Values(" & LV.ListItems(i) & _
         ",'" & LV.ListItems(i).SubItems(1) & "','" & LV.ListItems(i).SubItems(2) & _
         "','" & LV.ListItems(i).SubItems(3) & "','" & _
         LV.ListItems(i).SubItems(4) & "'," & Val(LV.ListItems(i).ListSubItems(5).Tag & "") & _
         ",'" & LV.ListItems(i).SubItems(6) & "','" & LV.ListItems(i).Tag & "')"
    Next i
    
    con.CommitTrans
    
    Load frmPrevRpt
    
    Dim strFromTo As String
    Dim iNumber As Integer
    If cmbRange.ListIndex = 0 Then
        iNumber = -15
    ElseIf cmbRange.ListIndex = 1 Then
        iNumber = -30
    ElseIf cmbRange.ListIndex = 2 Then
        iNumber = -60
    ElseIf cmbRange.ListIndex = 3 Then
        iNumber = -90
    End If
    
    If cmbRange.ListIndex = 4 Then
        strFromTo = Format(DtFrom, "dd-MMM-yyyy") & " To  " & Format(DtTo, "dd-MMM-yyyy")
    Else
        strFromTo = Format(DateAdd("d", iNumber, ServerDate), "dd-MMM-yyyy") & " To  " & Format(ServerDate, "dd-MMM-yyyy")
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptFProformaList.rpt")
    rpt.FormulaFields(3).Text = "'" & cmbMaker.Text & "'"
    rpt.FormulaFields(4).Text = "'" & strFromTo & "'"
    
    frmPrevRpt.ShowReport "", rpt
    
    Me.MousePointer = vbNormal
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub mnuPrintCoverLetter_Click()

    On Error GoTo err

    Screen.MousePointer = vbHourglass
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptDecathlonCoveringLetter.rpt")
    Load frmPrevRpt
    frmPrevRpt.ShowReport "{VCustomInvoice.CustomInvoice}='" & Left(LV.SelectedItem.key, Len(LV.SelectedItem.key) - 1) & "' ", rpt

    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub mnuPrintDecDocs_Click()
    Load frmPrintDecathlonDocs
    frmPrintDecathlonDocs.ShowMe (LV.SelectedItem.SubItems(3))
End Sub

Private Sub mnuPrintPackingLabels_Click()

    On Error GoTo err
    
    If LV.SelectedItem.SubItems(1) <> "DECATHLON" Then
        MsgBox "Packing Labels are available only for DECATHLON at the moment.", vbInformation
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    
    Dim rpt As CRAXDDRT.Report, f As New frmPrevRpt
    Set rpt = rptApp.OpenReport(rptexportpath & "\rptDecathlonPackingLabels.rpt")
    f.ShowReport "{VCustomInvoice.CustomInvoice}='" & LV.SelectedItem.SubItems(3) & "'", rpt
    
    Me.MousePointer = vbNormal
    
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub mnuShippingInstructions_Click()
'    Load frmShippingInstructions
'    If LV.ListItems.Count = 0 Then
'        frmShippingInstructions.Show , MainForm
'    Else
'        frmShippingInstructions.ShowForm LV.SelectedItem.SubItems(1), LV.SelectedItem.SubItems(2), LV.SelectedItem.SubItems(3)
'    End If
End Sub

Private Sub mnuUpdateCustomPrice_Click()
    If LV.ListItems.count = 0 Then Exit Sub
    If MsgBox("Are you Sure, You want to Update Custom Prices?", vbInformation + vbYesNo) = vbNo Then Exit Sub
    
    Dim dCustomPrice As Double
    dCustomPrice = Val(con.Execute("Select SUM(Qty*Price)/SUM(Qty) From CustomInvoiceItems Where CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'").Fields(0).Value & "")
    con.Execute "Update CustomInvoiceItems SET CustomPrice=" & dCustomPrice & " WHERE CustomInvoice='" & LV.SelectedItem.SubItems(3) & "'"
    
End Sub

Private Sub txtDesc_GotFocus()
    cmdRefresh.Default = False
    txtDesc.SelStart = 0
    txtDesc.SelLength = Len(txtDesc)
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
  
    If KeyAscii = 13 Then
    
        If Trim(txtDesc = "") Then
            MsgBox "Invalid Description", vbInformation
            Exit Sub
        End If
    
        con.Execute "Update Vouchers Set [Description]='" & txtDesc & "' Where Sno=" & Val(LV.SelectedItem.key)
        LV.SelectedItem.ListSubItems(3) = txtDesc & txtDesc.Tag
        txtDesc.Visible = False
        cmdRefresh.Default = False
    ElseIf KeyAscii = 27 Then
        txtDesc.Visible = False
        cmdRefresh.Default = False
    End If

End Sub

Private Sub txtDesc_LostFocus()
    txtDesc.Visible = False
End Sub
Private Sub RefreshSnos()
    For i = 1 To LV.ListItems.count
        LV.ListItems(i).Text = i
    Next i
End Sub

Private Sub mnuAddBookmarks_Click()
     
    Dim f As New frmAddBookmarkforReceiving
    Load f
    f.AddBookmark (Val(LV.SelectedItem.ListSubItems(2).Tag))
    
End Sub

Private Sub mnuAddFollowupEntries_Click()

    Dim lReturnDTsEntryID As Long
    lReturnDTsEntryID = GetSingleLongValue("MAX(EntryID)", "VendIssdDetail_ReturnDTs", " WHERE VIS_RefID=" & Val(LV.SelectedItem.ListSubItems(2).Tag))
    Dim f As New frmAddFollowupData
    Load f
    f.AddFollowup lReturnDTsEntryID, "(F1)"
    
End Sub

Private Sub mnuCloseMakerPO_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    If Val(LV.SelectedItem.SubItems(9)) > 0 Then Exit Sub
    
    If getDBPassword("CloseMakerPO") Then
        con.Execute "UPDATE VendIssued SET Closed=1 WHERE EntryID=" & Val(LV.SelectedItem.Tag)
    End If
    
End Sub

Private Sub mnuDelete_Click()

    On Error GoTo err
    If LV.ListItems.count = 0 Then Exit Sub
    If getDBPassword("DeleteProdIss") = False Then Exit Sub
    
    'Check for Receivings...
    If LV.SelectedItem.SubItems(6) = "Un-Auth." Then
        If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo) = vbYes Then
            Call DeleteIssuance
            con.Execute "DELETE FROM VendIssued WHERE RecieptID='" & LV.SelectedItem.SubItems(2) & "'"
            'Call RefreshLV
        End If
    Else
        If MsgBox("Are you sure to Delete this Aurhorized Order ?", vbQuestion + vbYesNo) = vbYes Then
            'Check Rcving Against this Order.
            Dim lCount As Long
            lCount = GetSingleLongValue("COUNT(Issuance_RefID)", "VendReceived", "WHERE Issuance_RefID=" & Val(LV.SelectedItem.Tag))
            If lCount > 0 Then
                MsgBox "Can't Delete, It has been Received.", vbInformation
                Exit Sub
            Else
                Call DeleteIssuance
                'Call RefreshLV
            End If
        End If
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub DeleteIssuance()

    Dim lRcvd_RefID As Long, lEntryID As Long, lDetailEntryID As Long
    lEntryID = Val(LV.SelectedItem.Tag)
    lDetailEntryID = GetSingleLongValue("EntryID", "VendIssdDetail", " WHERE RefID=" & lEntryID)
    lRcvd_RefID = GetSingleLongValue("Rcvd_RefID", "VendIssdDetail", " WHERE RefID=" & Val(LV.SelectedItem.Tag))
        
    Dim bRepairLot As Boolean
    bRepairLot = GetSingleBooleanValue("ReWorkLot", "VendIssdDetail", " WHERE RefID=" & lEntryID)
    
    Call StartTrans(con)
    
    If bRepairLot And lRcvd_RefID = 0 Then
        'It's Repair Lot & It's not Issued after Receiving, It was issued from Re-Work Issuance....
        con.Execute "UPDATE VendRcvdDetailReWorkDetail SET VendRcvdDetailReWorkDetail.IssQty=VendRcvdDetailReWorkDetail.IssQty-VendReWorkIssuanceDetail.Qty FROM VendRcvdDetailReWorkDetail INNER JOIN " & _
         "VendReWorkIssuanceDetail ON VendRcvdDetailReWorkDetail.EntryID=VRDRWD_RefID INNER JOIN VendIssdDetail ON " & _
         "VendReWorkIssuanceDetail.VID_RefID=VendIssdDetail.EntryID WHERE VendIssdDetail.EntryID=" & lDetailEntryID
        
        con.Execute "DELETE FROM VendReWorkIssuanceDetail WHERE VID_RefID=" & lDetailEntryID
        
    End If
    
    'Update Entries that were SF Issuance...
    
    con.Execute "UPDATE VendRcvdDetail SET VendRcvdDetail.IssQty=VendRcvdDetail.IssQty-ISNULL(T1.QtyIssd,0) FROM VendRcvdDetail " & _
     "INNER JOIN (SELECT SOO_RefID,SUM(QtyIssd) AS QtyIssd FROM StockOrderOpening_Issuance WHERE VID_RefID=" & lDetailEntryID & " GROUP BY SOO_RefID)  T1 " & _
     "ON VendRcvdDetail.Opening_RefID=T1.SOO_RefID"
    
    con.Execute "UPDATE VendRcvdDetail SET VendRcvdDetail.IssQty=VendRcvdDetail.IssQty-VendIssdDetail.IssQty FROM " & _
     "VendRcvdDetail INNER JOIN VendIssdDetail ON VendRcvdDetail.EntryID=VendIssdDetail.Rcvd_RefID " & _
     "WHERE VendIssdDetail.RefID=" & lEntryID
    'con.Execute "UPDATE VendRcvdDetail SET IssQty=IssQty-" & val( WHERE"
    con.Execute "DELETE FROM VendIssued WHERE RecieptID='" & LV.SelectedItem.SubItems(2) & "'"
                
    con.CommitTrans
    
End Sub

Private Sub mnuEditPromises_Click()
    Dim f As New frmEditMakerPOReturnDates
    Load f
    f.ShowMe Val(LV.SelectedItem.Tag)
End Sub

Private Sub mnuMakerBalanceReport_Click()

    Dim rpt As CRAXDDRT.Report, strSelection As String
    Set rpt = rptApp.OpenReport(RptPath & "\MakerBalanceReport.rpt")
    
    strSelection = "{Makers.VendID}<>" & iFactoryMakerID
    
    strSelection = strSelection & " AND {VendIssdDetail.RcvdQty}<{VendIssdDetail.IssQty}"
    
    If cmbProcess.ID <> "0" Then
        strSelection = strSelection & " AND {VendIssued.ProcessID} IN[" & Replace(cmbProcess.ID, "'", "") & "]"
    End If
    
    If cmbMaker.ID <> "0" Then
        strSelection = strSelection & " AND {VendIssued.VendID}=" & cmbMaker.ID
    End If
    
    Dim f As New frmPrevRpt
    f.ShowReport strSelection, rpt
    
End Sub

Private Sub mnuPrintGatePass_Click()
    Call PrintIssSlip(True)
End Sub

Private Sub mnuPrintItemPicture_Click()
    Call PrintMasterPOPictures(False)
End Sub

Private Sub PrintMasterPOPictures(Optional bPrintArtwork As Boolean = False)

    On Error GoTo err
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\IssListArticlePic.rpt")
    
    Dim f As New frmPrevRpt
    
    Dim strSelection As String

    strSelection = "{VendIssued.MasterPONo}='" & LV.SelectedItem.ListSubItems(9).Tag & "'"
    
    f.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub mnuPrintList_Click()
    Call PrintMasterPO
End Sub

Private Sub PrintMasterPO(Optional bHideRate As Boolean)
    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\IssList.rpt")
    
    If strCompany = "IAA" Or strCompany = "Kami" Then
        rpt.FormulaFields.GetItemByName("HideRate").Text = bHideRate
    End If
'    Dim strList As String
'    Dim i As Integer
'    For i = 1 To LV.ListItems.count
'        With LV.ListItems(i)
'            If .Checked Then
'                strList = strList & .Tag & ","
'            End If
'        End With
'    Next
'    strList = Left(strList, Len(strList) - 1)
    Dim f As New frmPrevRpt
    
    Dim strSelection As String
'    strSelection = "{VendIssued.EntryID} IN[" & strList & "]"
    strSelection = "{VendIssued.MasterPONo}='" & LV.SelectedItem.ListSubItems(9).Tag & "'"
'    Dim rs As New ADODB.Recordset
'    Dim strOrders As String
'    strOrders = ""
'    With rs
'        .Open "SELECT DISTINCT OrderNo FROM VendIssdDetail WHERE RefID IN(" & strList & ")", con, adOpenForwardOnly, adLockReadOnly
'        Do Until .EOF
'            strOrders = strOrders & !OrderNo & ","
'            .MoveNEXT
'        Loop
'        .Close
'    End With
'    Set rs = Nothing
'    strOrders = Left(strOrders, Len(strOrders) - 1)
'
'    rpt.FormulaFields.GetItemByName("Orders").Text = "'" & strOrders & "'"
    
    f.ShowReport strSelection, rpt
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
    
Private Sub mnuPrintMasterPOHideRate_Click()
    Call PrintMasterPO(True)
End Sub

Private Sub mnuPrintMasterPOStatus_Click()
    Call PrintMasterPOStatus(LV.SelectedItem.SubItems(10))
End Sub

Private Sub mnuPrintPTC_Click()
    Call PrintPTC(LV.SelectedItem.ListSubItems(9), False)
End Sub

Private Sub mnuPrintPurchaseOrder_Click()
    Call PrintIssSlip
End Sub

Private Sub PrintIssSlip(Optional bPrintGP As Boolean = False, Optional bHideCustomer As Boolean = False)

    On Error GoTo err
    Dim rpt As CRAXDDRT.Report
    If bPrintGP Then
        Set rpt = rptApp.OpenReport(RptPath & "\IssSlipGP.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\IssSlip.rpt")
    End If
    
    
    Dim strList As String
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                strList = strList & .Tag & ","
            End If
        End With
    Next
    
    Dim f As New frmPrevRpt
    If strList = "" Then
        f.ShowReport "{VendIssued.EntryID} IN[" & LV.SelectedItem.Tag & "]", rpt
    Else
        f.ShowReport "{VendIssued.EntryID} IN[" & Left(strList, Len(strList) - 1) & "]", rpt
    End If
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub mnuPrintPurchaseOrderWithoutRate_Click()
    Call PrintIssSlip(False, True)
End Sub

Private Sub mnuPrintSFStock_Click()

    Dim strCond As String
    strCond = "{@NetQty}>0"
    If cmbItems.ID <> "0" Then
        strCond = strCond & " AND {StockOrderOpening.ItemID}='" & cmbItems.ID & "'"
    End If
    If cmbProcess.ID <> "0" Then
        strCond = strCond & " AND {StockOrderOpening.ProcessID}=" & cmbProcess.ID
    End If
    
    Dim strStores As String
    Dim frm As New frmGetStores
    Load frm
    If frm.getData(strStores) = False Then Exit Sub
    
    If strStores <> "All" Then
        strCond = strCond & " AND {StockOrderOpening.Store_RefID} IN[" & Replace(strStores, "'", "") & "]"
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\SFStockReportStorewise.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport strCond, rpt
    
End Sub

Private Sub mnuPrintThisList_Click()

    Dim strFilters As String
    strFilters = "Maker:" & cmbMaker.Text & " | Process:" & Replace(cmbProcess.Text, "'", "") & " | Catagory:" & Replace(cmbItemCat.Text, "'", "") & " | Article:" & cmbItems.Text & " | Customer:" & Replace(cmbCustomer.ID, "'", "")
    
    If Trim(txtLotNo) <> "" Then
        strFilters = strFilters & " | Lot No.:" & Trim(txtLotNo)
    End If
    
    If Trim(txtOrderNo) <> "" Then
        strFilters = strFilters & " | Order No.:" & Trim(txtOrderNo)
    End If
    
    If Trim(txtMasterPONo) <> "" Then
        strFilters = strFilters & " | Master PO:" & Trim(txtMasterPONo)
    End If
    
    If Trim(txtInActiveDays) <> "" Then
        strFilters = strFilters & " | In-Active Days:" & Trim(txtInActiveDays)
    End If
        
    If chkShowMasterPOOnly.Value = vbChecked Then
        strFilters = strFilters & " | Master POs Only"
    End If
    
    If chkMasterPOOpen.Value = vbChecked Then
        strFilters = strFilters & " | Open Master POs Only"
    End If
    
    If chkRepairLots.Value = vbChecked Then
        strFilters = strFilters & " | Repair Lots Only"
    End If
    
    Dim strDateRange As String
    strDateRange = Format(DtFrom.Value, "dd-MMM-yyyy") & " to " & Format(DtTo.Value, "dd-MMM-yyyy")
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\IssuanceList.rpt")
    rpt.FormulaFields.GetItemByName("Filters").Text = "'" & strFilters & "'"
    rpt.FormulaFields.GetItemByName("DateRange").Text = "'" & strDateRange & "'"
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, strReportSQL
    
End Sub

Private Sub mnuPurchaseRegister_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\PurchaseCalendar.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{DateLookup.DateFull}=#" & DtFrom.Value & "# TO #" & DtTo.Value & "#", rpt
     
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 2 Then Exit Sub
    If LV.ListItems.count = 0 Then Exit Sub
    '---------------------------------------------
    Dim strRet As String
    strRet = ShowPopUpMenu1("Print PTC")
    If strRet = "Print PTC" Then
        Dim strLotNo As String
        strLotNo = LV.SelectedItem.SubItems(7)
        Call PrintPTC(strLotNo, False)
    End If
    
    
End Sub

Private Sub txtFind_Change()
    cmdFind.Enabled = Len(Trim(txtFind)) > 0
End Sub

Public Sub ShowIssuanceList(p_lIssProcessID As Long, p_lRcvProcessID As Long, myDT As Date)
    
    
    Call RefreshLV(p_lIssProcessID, p_lRcvProcessID, myDT)
    Me.Show
    
End Sub
